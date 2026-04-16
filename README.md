## 前言

这个仓库是一个**论文工具链**（脚本、模板、引用样式），支持用 Pandoc 将 Markdown 转为 Word。适用于CDTU的论文格式，让各位同学们可以专注论文内容的写作，不用在Word中被格式困扰

如遇到问题 欢迎提ISSUE

## 当前目录结构

```text
GraduationThesis/
├─ README.md
├─ .gitignore
├─ build-docx.sh               # 根目录快捷入口（转发到 scripts/）
├─ build-docx.bat              # 根目录快捷入口（转发到 scripts/）
│
├─ scripts/                    # 构建与文档处理脚本
│  ├─ build-docx.sh
│  ├─ build-docx.bat
│  ├─ fix_headings.py
│  ├─ fix_pages.py
│  └─ merge_docx.py
│
├─ assets/                     # 模板与引用相关静态资源
│  ├─ template.docx
│  └─ csl/
│     └─ china-national-standard-gb-t-7714-2015-numeric.csl
│
├─ references/
│  └─ refs.bib
│
├─ manuscript/                 # 论文正文
│  ├─ thesis.md
│  ├─ abstract.md
│  └─ acknowledgments.md
│
└─ output/                     # 导出产物
   ├─ thesis.docx
   ├─ abstract.docx
   └─ Final_thesis.docx
```

## 使用前提

- `pandoc >= 2.14`
- Microsoft Word（用于查看和微调）

检查：

```bash
pandoc --version
```

## 常用命令

Linux / macOS：

```bash
chmod +x build-docx.sh
./build-docx.sh
```

Windows：

```bat
call build-docx.bat
```
