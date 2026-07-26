# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Self-media educational video project: "从GPT、Deepseek到OpenClaw、Hermes — 一期视频打通AI底层逻辑"

A ~12-15 minute Chinese-language video covering 13 AI concepts in a progressive chain: LLM → Token → Embedding → Context/Window → Prompt → System Prompt → RAG+VectorDB → Function Calling → MCP → Agent → Agent Skill → Harness Engineering → 2026 Frontier.

## Repository Structure

- `ai-video-script-v3.md` — Current script (13 sections, ~12-15 min)
- `ai-video-script-v2.md` — Previous version (preserved for reference)
- `storyboard-v3.md` — 42-shot storyboard with visual/audio/transition details
- `ai-concepts-video-v3.pptx` — Latest PPT export (20 pages, warm beige style)
- `projects/ai-concepts-video_ppt169_20260726/` — PPT project files (SVGs, design spec, spec_lock, speaker notes)

## Key Design Specs

**Color palette** (warm beige, NO dark/gray/black):
- Background: `#F5F0E8` / Card: `#FAF7F2` / Primary: `#8B6F47` / Accent: `#E8A87C`
- Text: `#5D4E37` / Muted: `#8B7D6B` / Border: `#EDE5D5`

**PPT workflow**: Uses `ppt-master` skill pipeline. SVGs are hand-written (not script-generated), exported to PPTX via `svg_to_pptx.py`.

## Conventions

- Script versioning: `ai-video-script-v{N}.md` (keep previous versions)
- PPTX versioning: `ai-concepts-video-v{N}.pptx` (previous version may be locked by PowerPoint)
- PPT source SVGs live in `projects/*/svg_output/`; regenerate PPTX from there rather than editing PPTX directly
- Storyboard references shots by number (镜1–镜42); update storyboard when script sections change
