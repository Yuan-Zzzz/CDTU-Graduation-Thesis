@echo off
chcp 65001 >nul
echo ⏳ 开始生成带目录的论文 Word 文档...

set INPUT=thesis.md
set OUTPUT=thesis.docx
set TEMPLATE=template.docx
set BIB=references.bib
set CSL=china-national-standard-gb-t-7714-2015-numeric.csl
set TOC_DEPTH=3

pandoc "%INPUT%" --reference-doc "%TEMPLATE%" --bibliography "%BIB%" --csl "%CSL%" --citeproc --toc --toc-depth="%TOC_DEPTH%" -o "%OUTPUT%"

if %ERRORLEVEL% EQU 0 (
    echo ✅ 成功! 论文已生成: %OUTPUT%
) else (
    echo ❌ 生成失败! 请检查报错信息。
)
pause
