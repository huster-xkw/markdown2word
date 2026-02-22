# Skill 说明

这个目录用于 AI 工作流，不是给普通用户直接点开使用的界面。

## 目录内容

- `deepextract-doc-converter/SKILL.md`: Skill 规则
- `deepextract-doc-converter/scripts/convert_with_deepextract.py`: 实际转换脚本

## 手动安装（不走 npx）

把 `deepextract-doc-converter` 整个目录复制到：

`~/.config/opencode/skills/deepextract-doc-converter`

然后建议配置：

```bash
export DEEPEXTRACT_ROOT="/你的/deepextract_github/路径"
export MINERU_API_KEY="你的MinerUKey"
```

## 普通用户使用方式

普通用户请直接启动 Web 界面：

```bash
python start.py
```

打开 `http://localhost:5000`
