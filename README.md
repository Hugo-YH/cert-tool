# cert-tool

实验动物质量合格证解析工具

## 功能概述

- 从 PDF / 图片 中识别二维码  
- 打开二维码中的 URL，抓取页面或接口中的证件字段  
- 将结果导出为 JSON 格式（方便后期接入其他功能）

## 快速开始

### 交互模式（推荐）

```bash
python cert_tool_allinone.py
# 然后拖拽合格证文件（PDF/图片）到终端，回车开始处理
# 输入 esc 退出
```

### 批处理模式

```bash
python cert_tool_allinone.py 文件1.pdf 文件2.jpg
# 自动处理所有文件
```

## 输出说明

- **输出格式**：JSON（每个源文件一个 JSON 文件）
- **输出位置**：与源文件相同目录
- **文件名**：根据推断的合格证编号命名（如 `B202601150412.json`），冲突时追加时间戳

### JSON 结构

```json
{
  "source_file": "原文件绝对路径",
  "qr_url": "二维码包含的 URL",
  "pcId": "证件 ID",
  "page_title": "网页标题",
  "fields": {
    "证书编号": "B202601150412",
    "单位": "示例养殖场",
    ...
  },
  "error": "错误信息（成功时为空）"
}
```

## 依赖

脚本会自动创建虚拟环境并安装依赖：

- pillow, opencv-python, pymupdf（图像处理与二维码识别）
- playwright, beautifulsoup4, lxml（网页抓取与解析）

## 调试

### 启用调试输出

在环境变量中设置 `CERT_TOOL_DEBUG=1`，脚本会在源文件目录下生成 `*_debug_fields.json`

```bash
export CERT_TOOL_DEBUG=1
python cert_tool_allinone.py 文件.pdf
```

### 调试文件

- `*_debug_page.png`：渲染的整页图片（用于检查二维码位置）
- `*_debug_clip_bl.png`：PDF 左下角裁剪图（高 DPI 识别用）
- `*_debug_fields.json`：抓取的原始字段（DEBUG 模式下）

## 仓库

https://github.com/Hugo-YH/cert-tool
