@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

set SCRIPT_DIR=%~dp0
set REPO_ROOT=%SCRIPT_DIR%..

set ABSTRACT=%REPO_ROOT%\manuscript\abstract.md
set INPUT=%REPO_ROOT%\manuscript\thesis.md
set ACK=%REPO_ROOT%\manuscript\acknowledgments.md

set ABSTRACT_OUT=%REPO_ROOT%\output\abstract.docx
set THESIS_OUT=%REPO_ROOT%\output\thesis.docx
set ACK_OUT=%REPO_ROOT%\output\acknowledgments.docx
set FINAL_OUT=%REPO_ROOT%\output\Final_thesis.docx

set TEMPLATE=%REPO_ROOT%\assets\template.docx
set BIB=%REPO_ROOT%\references\refs.bib
set CSL=%REPO_ROOT%\assets\csl\china-national-standard-gb-t-7714-2015-numeric.csl
set TOC_DEPTH=3

if not exist "%REPO_ROOT%\output" mkdir "%REPO_ROOT%\output"

echo 开始生成论文 Word 文档...
echo 生成摘要...
pandoc "%ABSTRACT%" --reference-doc "%TEMPLATE%" -o "%ABSTRACT_OUT%"
if not %ERRORLEVEL% EQU 0 goto :fail

echo 生成正文...
pandoc "%INPUT%" --reference-doc "%TEMPLATE%" --bibliography "%BIB%" --csl "%CSL%" --citeproc --toc --toc-depth="%TOC_DEPTH%" --metadata toc-title="目录" -o "%THESIS_OUT%"
if not %ERRORLEVEL% EQU 0 goto :fail

echo 生成致谢...
pandoc "%ACK%" --reference-doc "%TEMPLATE%" -o "%ACK_OUT%"
if not %ERRORLEVEL% EQU 0 goto :fail

if not exist "%REPO_ROOT%\venv" python -m venv "%REPO_ROOT%\venv"
call "%REPO_ROOT%\venv\Scripts\activate.bat"
pip install python-docx docxcompose >nul 2>nul

python "%REPO_ROOT%\scripts\fix_headings.py" "%ABSTRACT_OUT%"
if not %ERRORLEVEL% EQU 0 goto :fail
python "%REPO_ROOT%\scripts\fix_headings.py" "%ACK_OUT%"
if not %ERRORLEVEL% EQU 0 goto :fail
python "%REPO_ROOT%\scripts\merge_docx.py"
if not %ERRORLEVEL% EQU 0 goto :fail
python "%REPO_ROOT%\scripts\fix_pages.py"
if not %ERRORLEVEL% EQU 0 goto :fail

echo 成功: 最终论文已生成 %FINAL_OUT%
goto :end

:fail
echo 失败: 请检查上方报错信息

:end
pause
