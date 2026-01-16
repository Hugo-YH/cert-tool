#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
实验动物质量合格证：二维码解析 -> 打开URL -> 抓取页面字段 -> 导出Excel
一体化脚本（自动创建虚拟环境 + 安装依赖 + 安装 Playwright 浏览器）

用法：
1) 双击/运行：
   python cert_tool_allinone.py
   -> 弹窗选择一个或多个 PDF/图片，导出 Excel

2) “拖拽文件到脚本上”（Windows/macOS 常见）或命令行传参：
   python cert_tool_allinone.py 文件1.pdf 文件2.jpg
   -> 自动批处理并导出 Excel（默认输出到当前目录）
"""

import os
import re
import sys
import time
import subprocess
import traceback
import warnings
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

# 安装/启动过程的详细日志（静默模式下写入此文件）
SETUP_LOG_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".cert_tool_setup.log")

# 尽量减少启动时的噪声提示（例如 SyntaxWarning 等）
warnings.filterwarnings("ignore", category=SyntaxWarning)

# 开启调试输出：export CERT_TOOL_DEBUG=1
DEBUG_MODE = os.environ.get("CERT_TOOL_DEBUG", "").strip() == "1"


# =========================
# 0) 虚拟环境与依赖自举
# =========================

VENV_DIRNAME = ".venv_cert"
PLAYWRIGHT_MARK = ".playwright_browsers_installed"

REQUIRED_PACKAGES = [
    "pandas",
    "openpyxl",
    "pillow",
    "opencv-python",
    "pymupdf",
    "playwright",
    "beautifulsoup4",
    "lxml",
]

# 你如果希望固定版本，可改为如 "pandas==2.2.2" 这种形式


def is_in_venv() -> bool:
    # 在 venv 内：sys.prefix != sys.base_prefix
    return getattr(sys, "base_prefix", sys.prefix) != sys.prefix


def venv_python_path(venv_dir: str) -> str:
    if os.name == "nt":
        return os.path.join(venv_dir, "Scripts", "python.exe")
    return os.path.join(venv_dir, "bin", "python")


def run_cmd(
    cmd: List[str],
    cwd: Optional[str] = None,
    quiet: bool = False,
    progress_label: Optional[str] = None,
    log_path: Optional[str] = None,
) -> None:
    """运行子命令。

    quiet=True 时：不输出命令与子进程输出（写入 log_path），终端仅显示一行进度动画。
    """
    if not quiet:
        subprocess.check_call(cmd, cwd=cwd)
        return

    if log_path is None:
        log_path = SETUP_LOG_PATH

    spinner = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
    label = progress_label or "⏳ 正在加载"

    with open(log_path, "a", encoding="utf-8") as lf:
        lf.write("\n" + "=" * 80 + "\n")
        lf.write("$ " + " ".join(cmd) + "\n")
        lf.flush()

        p = subprocess.Popen(cmd, cwd=cwd, stdout=lf, stderr=lf)
        i = 0
        last_msg_time = time.time()
        
        while True:
            ret = p.poll()
            if ret is not None:
                break
            
            # 每隔3秒输出一次"还在加载"提示
            current_time = time.time()
            if current_time - last_msg_time >= 3.0:
                sys.stdout.write("\r" + " " * (len(label) + 4) + "\r")
                sys.stdout.flush()
                print(f"  ⏳ {label} (还在初始化中，请稍候...)")
                last_msg_time = current_time
            
            sys.stdout.write("\r" + f"{spinner[i % len(spinner)]} {label}")
            sys.stdout.flush()
            time.sleep(0.05)  # 更频繁的刷新，让 spinner 动画更流畅
            i += 1

        # 清理进度行
        sys.stdout.write("\r" + " " * (len(label) + 4) + "\r")
        sys.stdout.flush()

        if ret != 0:
            raise subprocess.CalledProcessError(ret, cmd)


def ensure_venv_and_rerun() -> None:
    """
    若当前不在 venv 中，则创建 venv 并使用 venv 的 python 重新执行本脚本。
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    venv_dir = os.path.join(script_dir, VENV_DIRNAME)
    py_in_venv = venv_python_path(venv_dir)

    if is_in_venv():
        return

    if not os.path.exists(py_in_venv):
        import venv

        print("  → 创建虚拟环境...")
        builder = venv.EnvBuilder(with_pip=True, clear=False, upgrade=False)
        builder.create(venv_dir)
        print("  ✓ 虚拟环境已创建\n")

    # 使用 venv 的 python 重新执行本脚本（并把参数原样传递）
    print("  → 启动虚拟环境并加载依赖/浏览器...\n")
    cmd = [py_in_venv, os.path.abspath(__file__)] + sys.argv[1:]
    subprocess.check_call(cmd, cwd=script_dir)
    sys.exit(0)


def pip_install(pkgs: List[str], progress_label: str = "[INFO] 正在安装依赖…") -> None:
    # 使用 -q 降低pip输出噪声；详细日志写入 .cert_tool_setup.log
    run_cmd(
        [sys.executable, "-m", "pip", "install", "-U", "-q"] + pkgs,
        quiet=True,
        progress_label=progress_label,
    )


def pip_install_with_progress(pkgs: List[str]) -> None:
    """逐个安装包并显示进度，格式: 📦 安装中 [1/8] pandas"""
    for idx, pkg in enumerate(pkgs, 1):
        label = f"📦 安装中 [{idx}/{len(pkgs)}] {pkg}"
        print(f"  {label}", end="", flush=True)
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "-U", "-q", pkg],
            capture_output=True,
            check=True
        )
        print(" ✓")



def ensure_packages_installed() -> None:
    """
    尝试导入核心库；缺失则 pip install。
    """
    missing = []

    # 用“导入探针”避免误判
    probes = {
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "pillow": "PIL",
        "opencv-python": "cv2",
        "pymupdf": "fitz",
        "playwright": "playwright",
        "beautifulsoup4": "bs4",
        "lxml": "lxml",
    }

    for pkg, mod in probes.items():
        try:
            __import__(mod)
        except Exception:
            missing.append(pkg)

    if missing:
        print(f"  → 检测到缺失库：{', '.join(missing)}")
        pip_install_with_progress(missing)
        print("  ✓ 依赖库已安装\n")

    # 确保 pip 自身更新（可选，静默）
    try:
        print("  → 更新安装工具...")
        pip_install(["pip", "setuptools", "wheel"], progress_label="📦 更新工具中")
        print("  ✓ 工具已更新\n")
    except Exception:
        pass


def ensure_playwright_browsers() -> None:
    """
    Playwright 需要额外下载浏览器内核；用标记文件避免每次都执行。
    """
    script_dir = os.path.dirname(os.path.abspath(__file__))
    mark_path = os.path.join(script_dir, PLAYWRIGHT_MARK)
    if os.path.exists(mark_path):
        return

    print("  → 下载/安装 Chromium、Firefox、WebKit 浏览器引擎...")
    print("     (这可能需要 1-5 分钟，取决于网络速度)\n")
    
    browsers = ["chromium", "firefox", "webkit"]
    for idx, browser in enumerate(browsers, 1):
        label = f"🌐 安装中 [{idx}/{len(browsers)}] {browser}"
        print(f"  {label}", end="", flush=True)
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", browser],
            cwd=script_dir,
            capture_output=True,
            check=True
        )
        print(" ✓")
    
    print("  ✓ 浏览器引擎已安装\n")

    with open(mark_path, "w", encoding="utf-8") as f:
        f.write(str(time.time()))


# =========================
# 1) 主功能：文件 -> 图片
# =========================

def _pdf_first_page_to_image(pdf_path: str, dpi: int = 500):
    import fitz  # PyMuPDF
    from PIL import Image

    doc = fitz.open(pdf_path)
    page = doc.load_page(0)

    # 提高渲染分辨率，提升小二维码识别成功率（必要时可调到 600）
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=mat, alpha=False)

    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


def _image_path_to_image(img_path: str):
    from PIL import Image
    return Image.open(img_path).convert("RGB")


def file_to_image(file_path: str):

    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        return _pdf_first_page_to_image(file_path)
    elif ext in [".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"]:
        return _image_path_to_image(file_path)
    else:
        raise ValueError(f"不支持的文件类型: {ext}")


# =========================
# PDF角落裁剪渲染 + QR识别辅助（提升小二维码识别成功率）
# =========================

def _pdf_render_clip_to_image(pdf_path: str, clip_rect, dpi: int = 900):
    """将PDF第一页指定区域以高DPI渲染为图片（用于小二维码识别）。"""
    import fitz  # PyMuPDF
    from PIL import Image

    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pix = page.get_pixmap(matrix=mat, alpha=False, clip=clip_rect)
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    doc.close()
    return img


def decode_qr_from_pdf(pdf_path: str) -> List[str]:
    """优先渲染PDF角落区域（高DPI）来识别二维码，命中率通常高于整页识别。"""
    import fitz  # PyMuPDF

    doc = fitz.open(pdf_path)
    page = doc.load_page(0)
    rect = page.rect
    w, h = rect.width, rect.height

    # 角落裁剪：先大后小，左下为主，右下为备
    clips = [
        fitz.Rect(0, h * 0.55, w * 0.45, h),
        fitz.Rect(0, h * 0.65, w * 0.35, h),
        fitz.Rect(0, h * 0.70, w * 0.30, h),
        fitz.Rect(w * 0.55, h * 0.55, w, h),
        fitz.Rect(w * 0.65, h * 0.65, w, h),
    ]
    doc.close()

    for clip in clips:
        try:
            clip_img = _pdf_render_clip_to_image(pdf_path, clip_rect=clip, dpi=900)
            qr_list = decode_qr_from_image(clip_img)
            if qr_list:
                return qr_list
        except Exception:
            pass

    return []


# =========================
# 2) 图像 -> QR 内容（OpenCV）
# =========================

def decode_qr_from_image(pil_img) -> List[str]:
    import cv2
    import numpy as np

    def try_decode(bgr) -> Optional[str]:
        detector = cv2.QRCodeDetector()
        data, _, _ = detector.detectAndDecode(bgr)
        if data and data.strip():
            return data.strip()
        return None

    bgr_full = cv2.cvtColor(np.array(pil_img), cv2.COLOR_RGB2BGR)
    h, w = bgr_full.shape[:2]

    # 证件二维码通常在角落：优先裁剪角落区域（从大到小逐步收敛）
    crops = []
    # 左下角（常见）
    crops.append(bgr_full[int(h * 0.55):h, 0:int(w * 0.45)])
    crops.append(bgr_full[int(h * 0.65):h, 0:int(w * 0.35)])
    crops.append(bgr_full[int(h * 0.70):h, 0:int(w * 0.30)])
    # 右下角（防模板变化）
    crops.append(bgr_full[int(h * 0.55):h, int(w * 0.55):w])

    candidates = [bgr_full] + crops
    scales = [1.0, 2.0, 3.0]
    rotations = [None, cv2.ROTATE_90_CLOCKWISE, cv2.ROTATE_180, cv2.ROTATE_90_COUNTERCLOCKWISE]

    for img in candidates:
        if img is None or img.size == 0:
            continue

        for sc in scales:
            if sc != 1.0:
                img2 = cv2.resize(img, None, fx=sc, fy=sc, interpolation=cv2.INTER_CUBIC)
            else:
                img2 = img

            # 直接识别
            got = try_decode(img2)
            if got:
                return [got]

            # 阈值增强后再识别
            gray = cv2.cvtColor(img2, cv2.COLOR_BGR2GRAY)
            gray = cv2.GaussianBlur(gray, (3, 3), 0)
            thr = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY, 31, 5
            )
            thr_bgr = cv2.cvtColor(thr, cv2.COLOR_GRAY2BGR)
            got = try_decode(thr_bgr)
            if got:
                return [got]

            # 旋转后识别（部分二维码方向/页旋转导致失败）
            for rot in rotations[1:]:
                rot_img = cv2.rotate(img2, rot)
                got = try_decode(rot_img)
                if got:
                    return [got]

    return []


def extract_pcid(url: str) -> Optional[str]:
    m = re.search(r"[?&]pcId=([0-9a-fA-F]+)", url)
    return m.group(1) if m else None


# =========================
# 3) URL -> 网页抽取（Playwright）
# =========================

def _normalize_text(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())


def extract_fields_from_html(html: str) -> Dict[str, str]:
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    fields: Dict[str, str] = {}

    # 1) 表格：两列 key/value
    for table in soup.find_all("table"):
        for tr in table.find_all("tr"):
            tds = tr.find_all(["td", "th"])
            if len(tds) >= 2:
                key = _normalize_text(tds[0].get_text(" ", strip=True))
                val = _normalize_text(tds[1].get_text(" ", strip=True))
                if key and val and key not in fields:
                    fields[key] = val

    # 2) dl/dt/dd
    for dl in soup.find_all("dl"):
        dts = dl.find_all("dt")
        dds = dl.find_all("dd")
        for dt, dd in zip(dts, dds):
            key = _normalize_text(dt.get_text(" ", strip=True))
            val = _normalize_text(dd.get_text(" ", strip=True))
            if key and val and key not in fields:
                fields[key] = val

    # 3) 兜底：匹配 “键：值”
    text = soup.get_text("\n", strip=True)
    for line in text.split("\n"):
        line = _normalize_text(line)
        if "：" in line:
            k, v = [p.strip() for p in line.split("：", 1)]
            if k and v and len(k) <= 30 and k not in fields:
                fields[k] = v

    # 过滤少量噪声（可按实际再加）
    noise_keys = {"首页", "返回", "打印", "下载", "关闭"}
    for nk in list(fields.keys()):
        if nk in noise_keys:
            fields.pop(nk, None)

    return fields


def _flatten_json(obj, parent_key: str = "", sep: str = ".") -> Dict[str, str]:
    """把JSON递归拍平为 {key: value}（value统一转为字符串），用于导出Excel。"""
    out: Dict[str, str] = {}

    def _add(k: str, v):
        if v is None:
            return
        s = str(v).strip()
        if s == "":
            return
        # 避免覆盖：如重复key则追加序号
        if k in out:
            i = 2
            nk = f"{k}{sep}{i}"
            while nk in out:
                i += 1
                nk = f"{k}{sep}{i}"
            out[nk] = s
        else:
            out[k] = s

    if isinstance(obj, dict):
        for k, v in obj.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else str(k)
            if isinstance(v, (dict, list)):
                out.update(_flatten_json(v, new_key, sep=sep))
            else:
                _add(new_key, v)
    elif isinstance(obj, list):
        for idx, v in enumerate(obj):
            new_key = f"{parent_key}{sep}{idx}" if parent_key else str(idx)
            if isinstance(v, (dict, list)):
                out.update(_flatten_json(v, new_key, sep=sep))
            else:
                _add(new_key, v)
    else:
        _add(parent_key or "value", obj)

    return out


def _pick_best_json(captured: List[Tuple[str, object]], url_hint: str = "") -> Optional[object]:
    """从捕获到的多个JSON响应中挑选最可能是“证照详情”的那个。"""
    if not captured:
        return None

    # 1) 优先包含 pcId 的响应
    pcid = extract_pcid(url_hint) if url_hint else None
    if pcid:
        for u, j in captured:
            try:
                s = str(j)
                if pcid in s:
                    return j
            except Exception:
                pass

    # 2) 其次：URL里带 detail/record/cert/sales/qr 等关键词
    keywords = ["detail", "record", "cert", "certificate", "sales", "qr", "code", "pcid"]
    for u, j in captured:
        lu = (u or "").lower()
        if any(k in lu for k in keywords):
            return j

    # 3) 兜底：选择“拍平后字段最多”的JSON
    best = None
    best_n = -1
    for u, j in captured:
        try:
            n = len(_flatten_json(j))
            if n > best_n:
                best_n = n
                best = j
        except Exception:
            pass
    return best


def scrape_cert_page(url: str, timeout_ms: int = 20000, wait_sec: float = 2.0) -> Tuple[str, Dict[str, str]]:
    """抓取证照网页。

    优先策略：监听网络响应抓后端JSON（字段更全/更稳/更快）
    兜底策略：抓HTML再解析（兼容没有JSON接口或接口加密的情况）
    """
    from playwright.sync_api import sync_playwright

    captured_json: List[Tuple[str, object]] = []

    def on_response(resp):
        try:
            ct = (resp.headers.get("content-type") or "").lower()
            if "application/json" in ct or "text/json" in ct:
                j = resp.json()
                captured_json.append((resp.url, j))
        except Exception:
            pass

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()
        page.on("response", on_response)

        # 用 networkidle 更容易等到接口返回
        try:
            page.goto(url, wait_until="networkidle", timeout=timeout_ms)
        except Exception:
            page.goto(url, wait_until="domcontentloaded", timeout=timeout_ms)

        # 额外等待一点点，给晚到的接口响应时间
        try:
            page.wait_for_timeout(int(wait_sec * 1000))
        except Exception:
            time.sleep(wait_sec)

        title = _normalize_text(page.title())
        html = page.content()
        context.close()
        browser.close()

    # 1) JSON优先
    best_json = _pick_best_json(captured_json, url_hint=url)
    if best_json is not None:
        fields = _flatten_json(best_json)
        # 标记来源，便于你核对
        fields["_source"] = "json"
        return title, fields

    # 2) 兜底：HTML解析
    fields = extract_fields_from_html(html)
    fields["_source"] = "html"
    return title, fields


# =========================
# 4) 汇总导出 Excel
# =========================


@dataclass
class CertResult:
    source_file: str
    qr_url: str
    pcid: Optional[str]
    page_title: str
    fields: Dict[str, str]
    error: Optional[str] = None


# =========================
# 辅助：推断合格证编号、输出文件名
# =========================

def _derive_cert_no(fields: Dict[str, str]) -> Optional[str]:
    r"""从抓取到的字段中推断“合格证编号”（常见形如 Bxxxx...）。

    优先：字段名包含“合格证编号/证书编号/编号/certNo/certificateNo”等
    兜底：在所有 value 里扫描类似 \bB\d{3,}\b 的编号
    """
    if not fields:
        return None

    # 1) 优先按字段名命中
    key_hints = [
        "合格证编号", "证书编号", "证书编号", "编号", "合格证号", "证书号",
        "certno", "cert_no", "certificateno", "certificate_no", "certificateid", "certid",
    ]

    for k, v in fields.items():
        lk = (k or "").lower()
        if any(h in k for h in key_hints[:6]) or any(h in lk for h in key_hints[6:]):
            if v:
                m = re.search(r"\bB\d{3,}\b", str(v))
                if m:
                    return m.group(0)
                # 若不是B开头，也先返回原值（做最小清洗）
                vv = str(v).strip()
                if vv:
                    return vv

    # 2) 兜底：扫所有 value
    for v in fields.values():
        if not v:
            continue
        m = re.search(r"\bB\d{3,}\b", str(v))
        if m:
            return m.group(0)

    return None


def _safe_excel_path(out_dir: str, base_name: str) -> str:
    """生成不会覆盖的输出xlsx路径。"""
    base = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff._-]+", "_", base_name).strip("_")
    if not base:
        base = "合格证解析结果"
    path = os.path.join(out_dir, f"{base}.xlsx")
    if not os.path.exists(path):
        return path
    # 若同名已存在，追加时间戳
    ts = time.strftime("%Y%m%d_%H%M%S")
    return os.path.join(out_dir, f"{base}_{ts}.xlsx")


def process_files(file_paths: List[str]) -> List[CertResult]:
    results: List[CertResult] = []
    for fp in file_paths:
        fp_abs = os.path.abspath(fp)
        try:
            ext = os.path.splitext(fp_abs)[1].lower()

            # 1) 优先对PDF做角落高DPI识别；图片文件直接走图像识别
            if ext == ".pdf":
                qr_list = decode_qr_from_pdf(fp_abs)
                # 同时渲染整页用于调试查看清晰度/位置
                img = file_to_image(fp_abs)
            else:
                img = file_to_image(fp_abs)
                qr_list = []

            # 4) 兜底：如果角落识别没读到，再对整页/整图跑一次多策略识别
            if not qr_list:
                qr_list = decode_qr_from_image(img)

            # 若仍未识别到二维码，才落盘调试图，便于定位问题
            if not qr_list:
                try:
                    debug_png = os.path.splitext(fp_abs)[0] + "_debug_page.png"
                    img.save(debug_png)
                    print("[DEBUG] 已保存渲染页图：", debug_png)
                except Exception:
                    pass

                if ext == ".pdf":
                    try:
                        import fitz
                        doc = fitz.open(fp_abs)
                        page = doc.load_page(0)
                        rect = page.rect
                        w, h = rect.width, rect.height
                        clip = fitz.Rect(0, h * 0.65, w * 0.35, h)
                        doc.close()

                        clip_img = _pdf_render_clip_to_image(fp_abs, clip_rect=clip, dpi=900)
                        debug_clip_png = os.path.splitext(fp_abs)[0] + "_debug_clip_bl.png"
                        clip_img.save(debug_clip_png)
                        print("[DEBUG] 已保存PDF左下角裁剪图：", debug_clip_png)
                    except Exception:
                        pass

            if not qr_list:
                results.append(CertResult(fp_abs, "", None, "", {}, error="未识别到二维码"))
                continue

            qr_url = qr_list[0]
            pcid = extract_pcid(qr_url) if qr_url else None

            title, fields = scrape_cert_page(qr_url)
            # 调试：可选落盘字段（export CERT_TOOL_DEBUG=1 开启）
            if DEBUG_MODE:
                try:
                    import json
                    dbg_fields = os.path.splitext(fp_abs)[0] + "_debug_fields.json"
                    with open(dbg_fields, "w", encoding="utf-8") as f:
                        json.dump(fields, f, ensure_ascii=False, indent=2)
                    print("[DEBUG] 已保存抓取字段：", dbg_fields)
                except Exception:
                    pass
            results.append(CertResult(fp_abs, qr_url, pcid, title, fields, error=None))

        except Exception as e:
            results.append(CertResult(fp_abs, "", None, "", {}, error=str(e)))

    return results


def export_to_excel(results: List[CertResult], out_xlsx: str) -> None:
    import pandas as pd

    # Sheet1：一证一行（宽表）
    all_keys = set()
    for r in results:
        all_keys.update(r.fields.keys())
    all_keys = sorted(all_keys)

    wide_rows = []
    for r in results:
        row = {
            "source_file": r.source_file,
            "qr_url": r.qr_url,
            "pcId": r.pcid,
            "page_title": r.page_title,
            "error": r.error or "",
        }
        for k in all_keys:
            row[k] = r.fields.get(k, "")
        wide_rows.append(row)

    df_wide = pd.DataFrame(wide_rows)

    # Sheet2：长表（更稳）
    long_rows = []
    for r in results:
        if r.fields:
            for k, v in r.fields.items():
                long_rows.append({
                    "source_file": r.source_file,
                    "qr_url": r.qr_url,
                    "pcId": r.pcid,
                    "page_title": r.page_title,
                    "field": k,
                    "value": v,
                    "error": r.error or "",
                })
        else:
            long_rows.append({
                "source_file": r.source_file,
                "qr_url": r.qr_url,
                "pcId": r.pcid,
                "page_title": r.page_title,
                "field": "",
                "value": "",
                "error": r.error or "",
            })

    df_long = pd.DataFrame(long_rows)

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df_wide.to_excel(writer, index=False, sheet_name="wide")
        df_long.to_excel(writer, index=False, sheet_name="long")


# =========================
# 5) 入口：命令行交互
# =========================


def parse_paths_from_input_line(line: str) -> List[str]:
    """解析终端一行输入中的一个或多个路径（支持 Finder 拖拽的转义/引号）。"""
    import shlex

    line = (line or "").strip()
    if not line:
        return []

    # 允许逗号/分号分隔
    chunks = [c.strip() for c in re.split(r"[;,]", line) if c.strip()]

    paths: List[str] = []
    for chunk in chunks:
        try:
            items = shlex.split(chunk)
        except Exception:
            items = [chunk]

        for it in items:
            p = os.path.expanduser(it)
            if p:
                paths.append(p)

    # 去重 + 规范化 + 存在性检查
    norm: List[str] = []
    seen = set()
    for p in paths:
        ap = os.path.abspath(p)
        if ap in seen:
            continue
        seen.add(ap)
        if not os.path.exists(ap):
            print(f"[WARN] 路径不存在，已跳过：{ap}")
            continue
        ext = os.path.splitext(ap)[1].lower()
        if ext not in {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}:
            print(f"[WARN] 不支持的文件类型，已跳过：{ap}")
            continue
        norm.append(ap)

    return norm


def interactive_drag_drop_loop():
    """交互模式：每次拖入一个（或多个）合格证文件路径后立即处理。

    - 直接把文件从 Finder 拖到终端窗口，回车就开始处理
    - 输入 esc 退出（也支持 quit/exit）
    """
    out_dir = os.path.dirname(os.path.abspath(__file__))

    print("\n✅ 已就绪。请拖入合格证文件（PDF/图片），或输入 esc 退出\n")

    while True:
        try:
            line = input("拖入> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\n[INFO] 退出。")
            return

        if not line:
            continue

        if line.lower() in {"esc", "quit", "exit"}:
            print("[INFO] 退出。")
            return

        file_paths = parse_paths_from_input_line(line)
        if not file_paths:
            continue

        # 逐个处理并按“合格证编号（Bxxxx）”命名导出
        for fp in file_paths:
            print(f"\n📄 正在处理：{os.path.basename(fp)}")
            results = process_files([fp])
            r0 = results[0] if results else None

            if not r0 or r0.error:
                err = r0.error if r0 else "未知错误"
                print(f"❌ 失败：{err}")
                # 失败也导出一份（便于留痕），用文件名+时间
                ts = time.strftime("%Y%m%d_%H%M%S")
                out_xlsx = _safe_excel_path(out_dir, f"失败_{os.path.splitext(os.path.basename(fp))[0]}_{ts}")
                export_to_excel(results, out_xlsx)
                print(f"📊 已导出：{os.path.basename(out_xlsx)}")
                continue

            cert_no = _derive_cert_no(r0.fields) or os.path.splitext(os.path.basename(fp))[0]
            out_xlsx = _safe_excel_path(out_dir, cert_no)
            export_to_excel(results, out_xlsx)
            print(f"✅ 成功！已导出 → {os.path.basename(out_xlsx)}")


def default_output_path(script_dir: str) -> str:
    ts = time.strftime("%Y%m%d_%H%M%S")
    return os.path.join(script_dir, f"合格证解析结果_{ts}.xlsx")


def main():
    args = [a for a in sys.argv[1:] if a and not a.startswith("--")]

    # 有参数：批处理，输出仍用时间戳文件名
    if args:
        file_paths = [os.path.abspath(os.path.expanduser(a)) for a in args]
        print(f"\n📋 批处理 {len(file_paths)} 个文件...\n")
        results = process_files(file_paths)

        ok_n = sum(1 for r in results if not r.error)
        bad_n = len(results) - ok_n
        print(f"\n📊 处理完成：✅ {ok_n} 成功，❌ {bad_n} 失败")

        script_dir = os.path.dirname(os.path.abspath(__file__))
        out_xlsx = default_output_path(script_dir)
        export_to_excel(results, out_xlsx)
        print(f"\n📁 已导出：{os.path.basename(out_xlsx)}")

        if bad_n:
            print("\n⚠️  失败文件：")
            for r in results:
                if r.error:
                    print(f"  • {os.path.basename(r.source_file)} → {r.error}")
        return

    # 无参数：交互拖拽模式（拖入即处理，按Bxxxx命名，等待下一个）
    interactive_drag_drop_loop()


if __name__ == "__main__":
    try:
        # 若不在 venv，则创建 venv 并用 venv python 重新执行
        ensure_venv_and_rerun()

        # 以下代码只会在 venv 内执行
        ensure_packages_installed()
        ensure_playwright_browsers()

        main()

    except KeyboardInterrupt:
        print("\n[INFO] 用户中断。")
    except Exception:
        print("[ERROR] 发生异常：")
        traceback.print_exc()
        print(f"[INFO] 详细安装/启动日志：{SETUP_LOG_PATH}")
        sys.exit(1)