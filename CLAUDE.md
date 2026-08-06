# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

**AI大家学** — 中文AI教育视频内容仓库，面向B站/抖音发布的科普频道。三个系列：

| 系列 | 目录 | 定位 |
|------|------|------|
| 🔧 VibeCoding实战篇 | `03-VibeCoding实战篇/` | 用Claude Code从零搭建AgentFlow平台（8期连续剧） |
| 🧬 大模型技术篇 | `04-大模型技术篇/` | LLM底层机制，5单元23期（API→部署→Agent→RAG→高阶） |
| 🚀 前沿AI认知篇 | `02-前沿认知篇/` | 追踪AI最新发展（Kimi K3等热点解读） |

频道总体规划见 `AI大家学/00-频道/频道规划书.md`，进度跟踪见 `AI大家学/00-频道/PROGRESS.md`。

## 内容生产流程

```
脚本.md → 分镜 (storyboard) → SVG素材 → PPT生成 → 录制 → 发布
```

每集目录标准结构：`脚本.md` + `素材/` + `代码/`（可选）+ `CLAUDE.md`（可选）

## PPT生成

两种方式并存：

- **Python**：`python-pptx` + `lxml` 注入动画XML。运行：`python generate_ppt.py`（从集数目录执行）
- **Node.js**：`pptxgenjs`。运行：`node generate-ppt.js`（从集数目录执行，需先 `npm install`）
- **ppt-master技能**：SVG手写生成后通过 `svg_to_pptx.py` 导出PPTX

修改PPT时应从SVG源文件重新导出，不要直接编辑PPTX。

## 设计规范

全局配色（米色系温馨风格，禁止深色/灰色/黑色背景）：

- 背景：`#F5F0E8` / 卡片：`#FAF7F2` / 主色：`#8B6F47` / 强调色：`#E8A87C`
- 正文：`#5D4E37` / 辅助文字：`#8B7D6B` / 边框：`#EDE5D5`

## 约定

- 所有对话和内容默认使用**中文**
- 脚本版本命名：`*-v{N}.md`（保留历史版本，不覆盖）
- PPTX版本命名：`*-v{N}.pptx`
- 分镜按镜头编号（镜1、镜2…）；脚本结构变更时需同步更新分镜
- 脚本中引用的数据需核实来源和时效性

## 依赖

根目录无 `package.json`（已gitignore）。各集数目录独立管理依赖：

- `AI大家学/02-前沿认知篇/E01-KimiK3/` — Node依赖：`pptxgenjs`
- Python脚本依赖：`python-pptx`、`lxml`、`beautifulsoup4`、`python-docx`（按需）
