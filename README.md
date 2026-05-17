# Invoicely

Invoicely 是一个本地 macOS 发票整理工具。它使用系统自带 Vision 做 OCR，PDF 会优先读取文本层，打包后可以作为独立 `.app` 使用。

## 项目结构

```text
Invoicely/
  src/invoicely/          Python 源码
  assets/                 App 图标等资源
  scripts/                构建脚本
  samples/invoice_templates/ 示例发票模板
  dist/                   打包后的 App，默认不进入 git
  requirements.txt        Python 依赖
```

## 本地运行源码

```bash
cd /Users/neoshi/Developer/Invoicely
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
PYTHONPATH=src python src/invoicely/invoice_sorter_native.py
```

## 重新打包 App

```bash
cd /Users/neoshi/Developer/Invoicely
./scripts/build_app.sh
```

打包结果会生成到：

```text
dist/Invoicely发票整理.app
```

构建临时文件会放在 `/private/tmp/invoicely_packaging`，脚本结束后会自动清理。
