#!/bin/bash

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

ABSTRACT="${REPO_ROOT}/manuscript/abstract.md"
INPUT="${REPO_ROOT}/manuscript/thesis.md"
ACK="${REPO_ROOT}/manuscript/acknowledgments.md"

ABSTRACT_OUT="${REPO_ROOT}/output/abstract.docx"
THESIS_OUT="${REPO_ROOT}/output/thesis.docx"
ACK_OUT="${REPO_ROOT}/output/acknowledgments.docx"
FINAL_OUT="${REPO_ROOT}/output/Final_thesis.docx"

TEMPLATE="${REPO_ROOT}/assets/template.docx"
BIB="${REPO_ROOT}/references/refs.bib"
CSL="${REPO_ROOT}/assets/csl/china-national-standard-gb-t-7714-2015-numeric.csl"
TOC_DEPTH=3

echo "开始生成论文 Word 文档..."

for file in "$ABSTRACT" "$INPUT" "$ACK" "$TEMPLATE" "$BIB" "$CSL"; do
  if [ ! -f "$file" ]; then
    echo "错误: 找不到文件 $file"
    exit 1
  fi
done

mkdir -p "${REPO_ROOT}/output"

echo "生成摘要..."
pandoc "$ABSTRACT" \
  --reference-doc "$TEMPLATE" \
  -o "$ABSTRACT_OUT"

echo "生成正文..."
pandoc "$INPUT" \
  --bibliography "$BIB" \
  --csl "$CSL" \
  --citeproc \
  --reference-doc "$TEMPLATE" \
  --toc \
  --toc-depth="$TOC_DEPTH" \
  --metadata toc-title="目录" \
  --metadata link-citations=true \
  -o "$THESIS_OUT"

echo "生成致谢..."
pandoc "$ACK" \
  --reference-doc "$TEMPLATE" \
  -o "$ACK_OUT"

echo "合并文档..."
if [ ! -d "${REPO_ROOT}/venv" ]; then
  python3 -m venv "${REPO_ROOT}/venv"
fi

source "${REPO_ROOT}/venv/bin/activate"
pip install python-docx docxcompose > /dev/null 2>&1

python "${REPO_ROOT}/scripts/fix_styles.py" "$ABSTRACT_OUT"
python "${REPO_ROOT}/scripts/fix_styles.py" "$ACK_OUT"
python "${REPO_ROOT}/scripts/fix_styles.py" "$THESIS_OUT"
python "${REPO_ROOT}/scripts/fix_tables.py" "$THESIS_OUT"
python "${REPO_ROOT}/scripts/merge_docx.py"
python "${REPO_ROOT}/scripts/fix_pages.py"
python "${REPO_ROOT}/scripts/fix_citations.py" "$FINAL_OUT"

echo "成功: 最终论文已生成 ${FINAL_OUT}"
