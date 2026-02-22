# DeepExtract

中文 | [English](#english)

DeepExtract 是一个文档结构化转换工具：支持 PDF / 图片 / Word / PPT / HTML / Markdown 输入，输出 Markdown 或 Word（`.docx`）。

## 两种使用方式

### A) 普通用户（Web 界面启动项目）
```bash
pip install -r backend/requirements.txt
python start.py
```
打开：http://localhost:5000

这是默认方式，直接本地启动服务后在浏览器使用。

### B) AI 工作流用户（安装 Skill）

支持两种安装方式：

1. npx 一键安装（推荐）

```bash
npx @coreyxiang/deepextract-skill
```

2. 手动安装（拖拽/复制 Skill 目录）

将仓库内 `skill/deepextract-doc-converter` 复制到：

`~/.config/opencode/skills/deepextract-doc-converter`

然后按需配置环境变量：

```bash
export DEEPEXTRACT_ROOT="/你的/deepextract_github/路径"
export MINERU_API_KEY="你的MinerUKey"
```

你也可以在自然语言里直接给 Word 排版要求，例如：

- "转 Word，行间距 1.75，宋体，正文12号"
- "转成 docx，一级标题16号，二级标题14号，段间距8pt"

Skill 会自动映射为对应参数（如 `line_spacing`、`font_zh`、`font_size_body` 等）。

## 功能亮点

- 多格式输入，统一转换流程
- 支持公式与表格解析（MinerU）
- 支持 Markdown 转可编辑 Word（含排版选项）
- 前后端同端口（Flask）
- 任务结果默认 5 分钟自动清理

## 项目结构

```text
deepextract_github/
├── front.html
├── qrcode.jpg
├── start.py
├── start.bat
├── md2word_final.py
├── mineru_extract.py
├── apikey.md
├── backend/
│   ├── app.py
│   └── requirements.txt
├── skill/
│   └── deepextract-doc-converter/
│       ├── SKILL.md
│       └── scripts/convert_with_deepextract.py
├── CONTRIBUTING.md
└── LICENSE
```

## 快速开始

1) 安装依赖

```bash
pip install -r backend/requirements.txt
```

2) 配置 `MINERU_API_KEY`（二选一）

方式 A（推荐，环境变量）

```bash
export MINERU_API_KEY="your_mineru_key"
```

方式 B（本地文件）

编辑 `apikey.md`：

```text
MINERU_API_KEY=your_mineru_key
```

3) 启动服务

Mac / Linux:

```bash
python start.py
```

Windows:

```bat
start.bat
```

4) 打开浏览器

`http://localhost:5000`

## 公众号二维码

欢迎关注阿玮的AI实战与商业思考

![微信公众号二维码](./qrcode.jpg)

## 安全说明

- 仓库不包含任何有效 API Key
- 请不要提交你自己的 `apikey.md`
- 推荐始终使用环境变量配置敏感信息

## 贡献

请先阅读 `CONTRIBUTING.md`。

## 许可

MIT，见 `LICENSE`。

---

## English

DeepExtract is a document structuring and conversion tool.
It accepts PDF / images / Word / PPT / HTML / Markdown and exports Markdown or Word (`.docx`).

### Two usage modes

#### A) Web app users

Run the local service and use the browser UI.

#### B) AI workflow users (Skill)

1. npx install (recommended):

```bash
npx @coreyxiang/deepextract-skill
```

2. Manual install:

Copy `skill/deepextract-doc-converter` to:

`~/.config/opencode/skills/deepextract-doc-converter`

Then set env vars if needed:

```bash
export DEEPEXTRACT_ROOT="/path/to/deepextract_github"
export MINERU_API_KEY="your_mineru_key"
```

You can also include Word style requirements in natural language, for example:

- "Convert to Word with line spacing 1.75, SimSun, body size 12"
- "Export as docx, H1 16pt, H2 14pt, paragraph spacing 8pt"

The skill maps these to style options automatically.

### Features

- Multi-format input with a unified conversion flow
- Formula and table extraction via MinerU
- Markdown to editable Word with style options
- Single-port Flask app for frontend + backend
- Result cleanup after 5 minutes by default

### Quick Start

1. Install dependencies:

```bash
pip install -r backend/requirements.txt
```

2. Configure `MINERU_API_KEY` (choose one):

Option A (recommended):

```bash
export MINERU_API_KEY="your_mineru_key"
```

Option B: edit `apikey.md`

```text
MINERU_API_KEY=your_mineru_key
```

3. Run:

```bash
python start.py
```

4. Open:

`http://localhost:5000`

### WeChat QR Code

The frontend loads `qrcode.jpg` from the project root and shows it in the success modal.

### Security

- No real API key is included in this repo
- Never commit your own secrets
- Prefer environment variables for credentials

### Contributing & License

- Contributing guide: `CONTRIBUTING.md`
- License: MIT (`LICENSE`)
