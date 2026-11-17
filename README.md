# AIdailybrief — AI 每日资讯汇总

自动生成每日 AI 资讯汇总，写入桌面 Excel，并支持从 B 站筛选 AI / 硬件相关视频。

## 功能

- 聚合近 24 小时 AI/ML 资讯：官方实验室博客、开源项目、X、Reddit、Hacker News、KOL 观点、GitHub Trending。
- 专门的 B 站检索：按「最新发布」搜索 AI + 硬件关键词，逐个核实 UP 主粉丝数，只保留官方/认证号或粉丝 ≥10 万。
- 输出桌面 Excel，表头固定为：
  `时间戳 | 今日要点 | 热门开源项目 | 值得关注的讨论/观点 | 来源链接 | 简要说明`
- 每天追加一行；同一天重复运行则更新该行，不产生重复；历史行保留。

## 目录结构

- `build_daily.mjs` — 把当天内容写入 / 追加到每日 Excel
- `skills_checklist.mjs` — 扫描本机技能目录，生成「技能与功能清单」Excel
- `automation.md` — 每日自动化任务的提示词模板
- `input.example.json` — `build_daily.mjs` 的输入示例

## 依赖

- Node.js（推荐使用 Codex 运行时自带的 node）
- `@oai/artifact-tool`：把脚本所在目录的 `node_modules` 做 junction 指向 Codex 运行时的 `node_modules` 即可，例如：

```powershell
New-Item -ItemType Junction -Path .\node_modules -Target "C:\Program Files\WindowsApps\<runtime>\dependencies\node\node_modules"
```

## 快速开始

1. 复制 `input.example.json` 为 `input.json` 并填写当天内容
2. 运行 `node build_daily.mjs input.json`
3. 结果输出到 `output/`，复制到桌面即可

## 安全提醒

B 站 `cookies.txt`、`config.json`（含本机路径）等凭据严禁提交，已加入 `.gitignore`。
