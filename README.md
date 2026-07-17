# docx-formatter

将已有的 Word 文档（`.docx`）格式化为统一的中文学术论文或中文公文样式。

本项目是一个可被 AI Agent / Codex 触发的 Skill，也可以直接作为命令行工具使用。它通过 `python-docx` 修改段落字体、字号、对齐、行距、缩进以及页面边距，并尽量保留原文档中的表格、图片和其他复杂内容。

## 支持的样式

| 样式 | 参数 | 适用场景 |
| --- | --- | --- |
| 中文学术论文 | `academic` | 毕业论文、学位论文、研究报告、期刊论文 |
| 中文公文 | `gongwen` | 通知、报告、请示、批复、意见、函、纪要等行政公文 |

详细的字号、字体、页面边距和行距规范见：

- [`references/academic_style.md`](references/academic_style.md)
- [`references/gongwen_style.md`](references/gongwen_style.md)

## 环境要求

- Python 3.8 或更高版本
- `python-docx`
- 输入文件必须是 `.docx`；`.doc` 文件请先用 Word 或 WPS 另存为 `.docx`

首次使用时运行：

```bash
bash scripts/setup.sh
```

也可以手动安装依赖：

```bash
python3 -m pip install python-docx
```

## 快速开始

### 直接格式化

适合标题已经使用 Word 标题样式，或标题符合项目内置编号规则的文档：

```bash
python3 scripts/format_docx.py \
  --input draft.docx \
  --style academic \
  --output draft_formatted.docx
```

公文格式只需将样式参数改为 `gongwen`：

```bash
python3 scripts/format_docx.py \
  --input notice.docx \
  --style gongwen \
  --output notice_formatted.docx
```

脚本会在终端输出识别到的标题、正文段落数量，以及使用自动检测或语义映射的统计信息。

### 推荐：先提取段落，再提供语义映射

当文档标题没有使用规范的 Word 样式，或存在容易被误判的编号列表时，先提取段落：

```bash
python3 scripts/extract_paragraphs.py --input draft.docx
```

根据输出的段落索引判断层级，然后将 JSON 映射传给格式化脚本：

```bash
python3 scripts/format_docx.py \
  --input draft.docx \
  --style academic \
  --heading-map '{"0":"title","5":"h1","11":"h2","18":"h3"}' \
  --output draft_formatted.docx
```

支持的段落类型如下：

| 类型 | 含义 |
| --- | --- |
| `title` | 文档总标题或公文主标题 |
| `h1` | 章标题 / 一级标题 |
| `h2` | 节标题 / 二级标题 |
| `h3` | 小节标题 / 三级标题 |
| `body` | 正文；不写入映射时默认按正文处理 |

## AI Skill 使用流程

当用户要求对 Word 文档进行“排版”“套格式”“改成论文格式”或“改成公文格式”时，可按以下流程调用本 Skill：

1. 判断目标样式：学术论文使用 `academic`，中文公文使用 `gongwen`。
2. 确认输入文件是 `.docx`。
3. 运行 `extract_paragraphs.py`，读取段落索引和文本。
4. 根据语义生成 `heading_map`，不要仅依据数字编号判断标题。
5. 运行 `format_docx.py` 生成带 `_formatted` 后缀的输出文件。
6. 向用户报告使用的样式、识别到的标题层级和可能需要确认的判断。

Skill 的触发说明和完整操作约定见 [`SKILL.md`](SKILL.md)。

## 处理范围

### 会处理

- 正文及标题段落的字体、字号、加粗、对齐方式
- 行距、首行缩进、段前段后间距
- 所有分节的纸张尺寸和页面边距
- `Title`、`Heading 1`、`Heading 2`、`Heading 3` 等常见标题样式
- 常见中文标题编号的自动识别

### 默认保留原样

- 表格及表格内容
- 图片和图形
- 页眉、页脚
- 脚注、尾注
- 批注和修订记录

这是有意的保守策略，避免统一段落样式时破坏用户已经设计好的复杂版式。

## 自动识别说明

脚本会按以下顺序判断段落类型：

1. 使用 `--heading-map` 中明确指定的类型；
2. 识别 Word 样式名；
3. 识别 OpenXML 的大纲级别；
4. 根据中文章节编号进行启发式判断。

自动识别适合结构规范的文档。若标题和正文都使用普通段落样式，建议使用 `extract_paragraphs.py` 配合 `--heading-map`，以减少误判。

## 测试文档

项目提供了两个用于本地验证的“混乱格式”测试文档生成器：

```bash
python3 scripts/create_test_doc.py \
  --type academic \
  --output /tmp/test_academic.docx

python3 scripts/create_test_doc.py \
  --type gongwen \
  --output /tmp/test_gongwen.docx
```

然后分别运行格式化命令，打开输出文件检查字体、行距、缩进和页边距：

```bash
python3 scripts/format_docx.py \
  --input /tmp/test_academic.docx \
  --style academic \
  --output /tmp/test_academic_formatted.docx
```

## 项目结构

```text
.
├── SKILL.md                       # AI Skill 定义与触发/执行规范
├── README.md                      # 项目说明与使用指南
├── references/
│   ├── academic_style.md           # 中文学术论文排版参数
│   └── gongwen_style.md            # 中文公文排版参数
└── scripts/
    ├── setup.sh                   # 安装运行依赖
    ├── extract_paragraphs.py      # 提取段落供语义分析
    ├── format_docx.py             # 核心格式化脚本
    └── create_test_doc.py         # 生成测试文档
```

## 常见问题

### 提示 `No module named 'docx'`

运行 `bash scripts/setup.sh`，或手动执行 `python3 -m pip install python-docx`。

### 没有识别出标题

这通常表示文档没有使用标题样式，且文本不符合自动检测规则。先运行 `extract_paragraphs.py`，再使用 `--heading-map` 明确指定标题层级。

### 公文标题字体显示异常

公文标题使用 `FZXiaoBiaoSong-B05S`。如果本机没有安装方正小标宋，Word/WPS 可能显示替代字体；文件中的字体名称仍会被正确写入，在安装对应字体的环境中打开即可。

### 公式、页眉页脚或表格没有被统一格式化

这些内容不在当前方案的处理范围内，脚本会尽量保留其原有格式。需要处理时请单独在 Word/WPS 中调整。

## 许可

当前仓库未声明开源许可证。如需公开分发或允许他人再使用，建议补充 `LICENSE` 文件并明确许可条款。
