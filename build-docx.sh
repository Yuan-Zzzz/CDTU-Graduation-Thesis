#!/bin/bash

# ==========================================
# 论文编译配置区 (根据你的实际文件名修改)
# ==========================================
INPUT="thesis.md"            # 你的 Markdown 源文件
OUTPUT="thesis.docx"         # 输出的 Word 文件名
TEMPLATE="template.docx"     # Word 模板文件
BIB="refs.bib"         # 参考文献库文件
CSL="china-national-standard-gb-t-7714-2015-numeric.csl" # 引用格式文件
TOC_DEPTH=3

# ==========================================
# 执行区 (通常不需要修改)
# ==========================================
echo "⏳ 开始生成论文 Word 文档..."

# 检查输入文件是否存在
if [ ! -f "$INPUT" ]; then
    echo "❌ 错误: 找不到 Markdown 文件 ($INPUT)"
    exit 1
fi

# 执行 Pandoc 命令
pandoc "$INPUT" \
  --reference-doc "$TEMPLATE" \
  --bibliography "$BIB" \
  --csl "$CSL" \
  --citeproc \
  --toc \
  --toc-depth="$TOC_DEPTH" \
  -o "$OUTPUT"

# 检查命令是否执行成功
if [ $? -eq 0 ]; then
    echo "✅ 成功! 论文已生成: $OUTPUT"
else
    echo "❌ 生成失败! 请检查上方的报错信息。"
fi
