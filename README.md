# cert-tool

实验动物质量合格证解析工具

功能概述：
- 从 PDF / 图片 中识别二维码
- 打开二维码中的 URL，抓取页面或接口中的证件字段
- 将结果导出为 Excel（wide / long 两个 sheet）

快速开始：

```bash
python cert_tool_allinone.py  # 交互拖拽或传入文件路径进行批处理
```

依赖（脚本会自动安装）：
- pandas, openpyxl, pillow, opencv-python, pymupdf, playwright, beautifulsoup4, lxml

调试：
- 导出调试字段：在环境变量中设置 `CERT_TOOL_DEBUG=1`

仓库： https://github.com/Hugo-YH/cert-tool
