# CDTU 毕业论文 Word 文档生成工具

用 Pandoc 把 Markdown 转成 Word，专门给成都工业学院的毕业论文格式定制。

写论文不用再跟 Word 的排版打架了，用 Markdown 写完一键导出。

## 功能

- Markdown 写作，纯文本方便 git 版本控制
- 自动生成目录，支持三级标题
- 参考文献用 GB/T 7714-2015 格式
- 基于 Word 模板，格式统一
- Windows / macOS / Linux 都能用
- 一键编译，不用手动操作

## 依赖

| 软件 | 版本 | 用途 | 下载 |
|------|------|------|------|
| Pandoc | ≥ 2.14 | Markdown 转 Word | [官网](https://pandoc.org/installing.html) |
| Microsoft Word | 2016+ | 编辑模板、查看文档 | [Microsoft 365](https://www.microsoft.com/zh-cn/microsoft-365) |

验证 Pandoc 安装：

```bash
pandoc --version
```

## 文件结构

```
CDTU-Graduation-Thesis/
├── thesis.md              # 论文正文（Markdown）
├── refs.bib               # 参考文献库
├── template.docx          # Word 格式模板
├── china-national-standard-gb-t-7714-2015-numeric.csl  # 引用格式
├── build-docx.sh          # Linux/macOS 编译脚本
├── build-docx.bat         # Windows 编译脚本
├── thesis.docx            # 生成的 Word 文件
├── README.md
└── .gitignore
```

## 快速开始

### 安装 Pandoc

macOS:
```bash
brew install pandoc
```

Ubuntu/Debian:
```bash
sudo apt-get update && sudo apt-get install pandoc
```

Windows: 从 [GitHub Releases](https://github.com/jgm/pandoc/releases) 下载安装包

### 克隆仓库

```bash
git clone https://github.com/Yuan-Zzzz/CDTU-Graduation-Thesis.git
cd CDTU-Graduation-Thesis
```

### 写论文

编辑 `thesis.md` 文件：

```markdown
---
title: "基于Unity的2D游戏设计与实现"
author: "张三"
date: "2026-04-14"
lang: "zh-CN"
toc: true
toc-depth: 3
numbersections: true
bibliography: refs.bib
---

# 摘要

本文研究了...[@unityManual2024]

# 第1章 绪论

## 1.1 选题背景
...
```

### 添加参考文献

在 `refs.bib` 里添加：

```bibtex
@book{sommerville2015,
  author    = {Sommerville, Ian},
  title     = {Software Engineering},
  year      = {2015},
  publisher = {Pearson}
}
```

正文里用 `[@sommerville2015]` 引用。

### 生成 Word

Linux/macOS:
```bash
chmod +x build-docx.sh && ./build-docx.sh
```

Windows:
```batch
call build-docx.bat
```

生成后的文件是 `thesis.docx`。

## 使用说明

### Markdown 语法

| 功能 | 语法 | 效果 |
|------|------|------|
| 一级标题 | `# 标题` | 章标题 |
| 二级标题 | `## 标题` | 1.1 小节 |
| 三级标题 | `### 标题` | 1.1.1 小节 |
| 粗体 | `**文字**` | 加粗 |
| 斜体 | `*文字*` | 斜体 |
| 引用 | `[@key]` | 参考文献 |
| 列表 | `- 项目` / `1. 项目` | 无序/有序 |
| 表格 | `| 表头 |` | 表格 |

### 元数据配置

`thesis.md` 文件开头 `---` 之间的配置：

```yaml
---
title: "论文标题"
author: "作者姓名"
date: "2026-04-14"
lang: "zh-CN"
toc: true
toc-depth: 3
numbersections: true
bibliography: refs.bib
---
```

### 参考文献

支持的条目类型：`@book`, `@article`, `@inproceedings`, `@manual`, `@online`

引用格式：
- 单个：`[@key]`
- 多个：`[@key1;@key2]`

### 修改模板

1. 用 Word 打开 `template.docx`
2. 修改样式（标题、正文、引用等）
3. 保存后重新运行编译脚本

## 编译配置

编辑 `build-docx.sh` 修改这些变量：

```bash
INPUT="thesis.md"            # 输入文件
OUTPUT="thesis.docx"         # 输出文件
TEMPLATE="template.docx"     # 模板
BIB="refs.bib"               # 参考文献
CSL="china-national-standard-gb-t-7714-2015-numeric.csl"
TOC_DEPTH=3                  # 目录深度
```

## 常见问题

**编译时提示 "pandoc: command not found"**

没装 Pandoc，先装一下。

**参考文献引用显示不对**

- 检查 `refs.bib` 语法
- 确认引用的 key 和条目匹配
- 检查 YAML 里的 bibliography 路径

**生成的 Word 格式不对**

- 打开 `template.docx` 改样式
- 调标题、段落、字体
- 保存后重新编译

**中文字体显示异常**

- 系统要有中文字体（思源黑体、微软雅黑等）
- 在 `template.docx` 里指定字体

**目录页码不对**

- Word 里打开 `thesis.docx`
- 右键目录 → 更新域 → 更新整个目录

## 相关技术

- [Pandoc](https://pandoc.org/) - 文档转换
- [CSL](https://citationstyles.org/) - 引用格式标准
- GB/T 7714-2015 - 国标参考文献格式

## 许可证

模板遵循学术规范，代码用 MIT 许可证。

CSL 样式来自 [Zotero Styles](https://www.zotero.org/styles)，CC BY-SA 3.0。

## 致谢

用到了这些工具：
- [Pandoc](https://pandoc.org/)
- [Zotero](https://www.zotero.org/)
- [zotero-chinese/styles](https://github.com/zotero-chinese/styles)

## 反馈

有问题或建议可以提 Issue。
