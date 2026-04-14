# CDTU 毕业论文 Word 文档生成工具

基于 **Pandoc** 的 Markdown 转 Word 工具，用于生成符合 **CDTU（成都工业学院）** 毕业论文格式要求的 Word 文档。

> 采用「Markdown 写作 + Word 模板 + 自动引用」的工作流，告别 Word 排版困扰。

---

## 功能特性

- ✅ 使用 Markdown 编写论文内容，纯文本易于版本控制
- ✅ 支持自动目录生成（三级标题深度可配置）
- ✅ 支持 GB/T 7714-2015 格式参考文献引用
- ✅ 基于 Word 模板渲染，确保格式符合学校要求
- ✅ 跨平台支持（Windows / macOS / Linux）
- ✅ 一键编译脚本，无需手动操作

---

## 依赖项

### 必需软件

| 软件 | 版本要求 | 用途 | 下载链接 |
|------|---------|------|----------|
| **Pandoc** | ≥ 2.14 | Markdown 转 Word 核心工具 | [官网下载](https://pandoc.org/installing.html) |
| **Microsoft Word** | 2016+ | 模板编辑与文档查看 | [Microsoft 365](https://www.microsoft.com/zh-cn/microsoft-365) |

### 验证安装

```bash
# 检查 Pandoc 是否安装成功
pandoc --version
```

---

## 文件结构

```
CDTU-Graduation-Thesis/
├── thesis.md                              # 论文正文（Markdown 源文件）
├── refs.bib                               # 参考文献库（BibTeX 格式）
├── template.docx                          # Word 格式模板
├── china-national-standard-gb-t-7714-2015-numeric.csl  # 引用格式
├── build-docx.sh                          # Linux/macOS 编译脚本
├── build-docx.bat                         # Windows 编译脚本
├── thesis.docx                            # 生成的论文 Word 文件（自动生成）
├── README.md                              # 本说明文档
└── .gitignore                             # Git 忽略配置
```

---

## 快速开始

### 1. 安装 Pandoc

**macOS:**
```bash
brew install pandoc
```

**Ubuntu/Debian:**
```bash
sudo apt-get update
sudo apt-get install pandoc
```

**Windows:**
下载安装包并安装：[Pandoc Releases](https://github.com/jgm/pandoc/releases)

### 2. 克隆仓库

```bash
git clone https://github.com/Yuan-Zzzz/CDTU-Graduation-Thesis.git
cd CDTU-Graduation-Thesis
```

### 3. 编辑论文内容

打开 `thesis.md` 文件，使用 Markdown 语法编写论文内容。

**示例：**
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

### 4. 添加参考文献

在 `refs.bib` 中添加参考文献条目：

```bibtex
@book{sommerville2015,
  author    = {Sommerville, Ian},
  title     = {Software Engineering},
  year      = {2015},
  publisher = {Pearson}
}
```

然后在正文中使用 `[@sommerville2015]` 引用。

### 5. 生成 Word 文档

**Linux / macOS:**
```bash
chmod +x build-docx.sh
./build-docx.sh
```

**Windows:**
```batch
call build-docx.bat
```

编译完成后，将生成 `thesis.docx` 文件。

---

## 使用说明

### Markdown 语法支持

Pandoc 支持的 Markdown 扩展语法：

| 功能 | Markdown 语法 | 说明 |
|------|--------------|------|
| 一级标题 | `# 标题` | 对应论文章标题 |
| 二级标题 | `## 标题` | 对应 1.1 小节 |
| 三级标题 | `### 标题` | 对应 1.1.1 小节 |
| 粗体 | `**文字**` | 加粗显示 |
| 斜体 | `*文字*` | 斜体显示 |
| 引用 | `[@key]` | 引用参考文献 |
| 列表 | `- 项目` 或 `1. 项目` | 无序/有序列表 |
| 表格 | `| 表头 |` | 普通表格 |

### YAML 元数据（Front Matter）

在 `thesis.md` 开头的 `---` 区域内配置：

```yaml
---
title: "论文标题"           # 论文题目
author: "作者姓名"          # 你的姓名
date: "2026-04-14"        # 日期
lang: "zh-CN"             # 语言（中文）
toc: true                 # 是否生成目录
toc-depth: 3              # 目录深度（1-6）
numbersections: true      # 章节自动编号
bibliography: refs.bib    # 参考文献文件
---
```

### 参考文献管理

1. 在 `refs.bib` 中添加条目
2. 支持类型：`@book`, `@article`, `@inproceedings`, `@manual`, `@online`
3. 在正文中使用 `[@key]` 格式引用，如 `[@unityManual2024]`
4. 引用多个文献：`[@key1;@key2]`

### 自定义 Word 模板

如需修改论文格式：

1. 打开 `template.docx`
2. 在 Word 中修改样式（标题、正文、引用等）
3. 保存后再次运行编译脚本

---

## 编译配置

编辑脚本文件可自定义配置：

```bash
# build-docx.sh 中的配置变量
INPUT="thesis.md"            # 输入 Markdown 文件
OUTPUT="thesis.docx"         # 输出 Word 文件
TEMPLATE="template.docx"     # 模板文件
BIB="refs.bib"               # 参考文献文件
CSL="china-national-standard-gb-t-7714-2015-numeric.csl"
TOC_DEPTH=3                  # 目录深度
```

---

## 常见问题

### Q: 编译时报 "pandoc: command not found"
**A:** 未安装 Pandoc，请先按照「依赖项」章节安装。

### Q: 参考文献引用显示不正确
**A:** 
1. 检查 `refs.bib` 中的条目语法是否正确
2. 确认引用的 `key` 与文献条目匹配
3. 检查 `thesis.md` 中 YAML 的 `bibliography` 路径是否正确

### Q: 生成的 Word 格式不符合要求
**A:** 
1. 打开 `template.docx` 修改样式
2. 确保标题样式、段落格式、字体等符合学校要求
3. 重新运行编译脚本

### Q: 中文字体显示异常
**A:** 
1. 确保系统安装了中文字体（如「思源黑体」「微软雅黑」等）
2. 在 `template.docx` 中指定中文字体

### Q: 目录页码不对
**A:** 
1. 在 Word 中打开生成的 `thesis.docx`
2. 右键目录 → 「更新域」→ 「更新整个目录」

---

## 技术栈

- [Pandoc](https://pandoc.org/) - 文档转换引擎
- [CSL (Citation Style Language)](https://citationstyles.org/) - 引用格式标准
- GB/T 7714-2015 - 中国国家标准参考文献格式

---

## 许可证

本项目模板文件遵循相关学术规范，代码部分采用 MIT 许可证。

CSL 样式文件来自 [Zotero Styles Repository](https://www.zotero.org/styles)，遵循 CC BY-SA 3.0 协议。

---

## 致谢

感谢以下项目提供的工具和资源：
- [Pandoc](https://pandoc.org/) - 文档转换神器
- [Zotero](https://www.zotero.org/) - 参考文献管理
- [GB/T 7714-2015 CSL](https://github.com/zotero-chinese/styles) - 中文引用格式

---

## 联系与反馈

如有问题或建议，欢迎提交 Issue 或 Pull Request。
