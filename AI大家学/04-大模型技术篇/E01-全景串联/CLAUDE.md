# CLAUDE.md

本文件为 Claude Code（claude.ai/code）在此仓库中工作时提供指引。

## 项目概述

自媒体教学视频项目："从GPT、Deepseek到OpenClaw、Hermes — 一期视频打通AI底层逻辑"

时长约 12-15 分钟的中文视频，按递进链讲解 13 个 AI 核心概念：LLM → Token → Embedding → Context/Window → Prompt → System Prompt → RAG+向量数据库 → Function Calling → MCP → Agent → Agent Skill → Harness Engineering → 2026 前沿。

## 仓库结构

- `脚本-v3.md` — 当前脚本（13 个板块，约 12-15 分钟）
- `分镜-v3.md` — 42 镜分镜脚本，含画面/音频/转场细节
- `全景串联-v3.pptx` — 最新 PPT 导出（米色系温馨风格）
- `CSDN_LLM_AgentSkill_原文.docx` — CSDN 原文参考
- `csdn_raw.html` / `csdn_to_docx.py` — 原文抓取与转换工具
- `素材/` — 视频素材（待制作）

## 设计规范

**配色方案**（米色系温馨风，禁止深色/灰色/黑色）：
- 背景：`#F5F0E8` / 卡片：`#FAF7F2` / 主色：`#8B6F47` / 强调色：`#E8A87C`
- 正文：`#5D4E37` / 辅助文字：`#8B7D6B` / 边框：`#EDE5D5`

**PPT 工作流**：使用 `ppt-master` 技能流水线。SVG 手写生成（不用脚本批量生成），通过 `svg_to_pptx.py` 导出为 PPTX。

## 约定

- 脚本版本命名：`脚本-v{N}.md`（保留历史版本）
- PPTX 版本命名：`全景串联-v{N}.pptx`
- PPT 源 SVG 存放在 `projects/*/svg_output/`；修改 PPT 应从 SVG 重新导出，不要直接编辑 PPTX
- 分镜按镜头编号（镜1–镜42）；脚本板块变更时需同步更新分镜
- 所有对话默认使用中文
