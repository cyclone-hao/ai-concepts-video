# Kimi K3 宣发稿 - Design Spec

> AI-driven promotional deck for Kimi K3 — China's breakthrough AI model.
> Machine-readable execution contract: `spec_lock.md`

## I. Project Information

| Item | Value |
| ---- | ----- |
| **Project Name** | kimi-k3-promo |
| **Canvas Format** | PPT 16:9 (1280×720) |
| **Page Count** | 14 |
| **Design Style** | A) General Versatile + 温暖科技感 |
| **Target Audience** | AI爱好者、科技关注者、B站/抖音用户 |
| **Use Case** | 频道宣传物料、社交媒体分享 |
| **Created Date** | 2026-08-08 |

---

## II. Canvas Specification

| Property | Value |
| -------- | ----- |
| **Format** | PPT 16:9 |
| **Dimensions** | 1280×720 |
| **viewBox** | `0 0 1280 720` |
| **Margins** | left/right 60px, top/bottom 50px |
| **Content Area** | 1160×620 (x=60, y=50, w=1160, h=620) |

---

## III. Visual Theme

### Theme Style

- **Style**: A) General Versatile + 温暖科技感
- **Theme**: Light theme — warm beige tones
- **Tone**: Warm, approachable, tech-forward, data-driven

### Color Scheme

| Role | HEX | Purpose |
| ---- | --- | ------- |
| **Background** | `#F5F0E8` | Page background |
| **Secondary bg** | `#FAF7F2` | Card/section background |
| **Primary** | `#8B6F47` | Title decorations, key sections |
| **Accent** | `#E8A87C` | Data highlights, key info, emphasis |
| **Secondary accent** | `#C9A96E` | Secondary emphasis, gradients |
| **Body text** | `#5D4E37` | Main body text |
| **Secondary text** | `#8B7D6B` | Captions, annotations |
| **Tertiary text** | `#A89B8A` | Supplementary info, footers |
| **Border/divider** | `#EDE5D5` | Card borders, divider lines |
| **Success** | `#6B8E5A` | Positive indicators (muted green) |
| **Warning** | `#C75C3A` | Negative/warning indicators (muted red-orange) |

### Gradient Scheme

```xml
<linearGradient id="titleGradient" x1="0%" y1="0%" x2="100%" y2="100%">
  <stop offset="0%" stop-color="#8B6F47"/>
  <stop offset="100%" stop-color="#E8A87C"/>
</linearGradient>

<linearGradient id="accentBar" x1="0%" y1="0%" x2="100%" y2="0%">
  <stop offset="0%" stop-color="#8B6F47"/>
  <stop offset="100%" stop-color="#C9A96E"/>
</linearGradient>

<radialGradient id="bgDecor" cx="80%" cy="20%" r="50%">
  <stop offset="0%" stop-color="#E8A87C" stop-opacity="0.12"/>
  <stop offset="100%" stop-color="#E8A87C" stop-opacity="0"/>
</radialGradient>
```

---

## IV. Typography System

### Font Plan

**Typography direction**: 楷体标题 × 黑体正文 — 人文温度感与现代科技感的对比

| Role | Chinese | English | Fallback tail |
| ---- | ------- | ------- | ------------- |
| **Title** | KaiTi | Georgia | serif |
| **Body** | "Microsoft YaHei" | Arial | sans-serif |
| **Emphasis** | — | — | same as Body |
| **Code** | — | Consolas | monospace |

**Per-role font stacks**:
- Title: `KaiTi, Georgia, serif`
- Body: `"Microsoft YaHei", Arial, sans-serif`
- Emphasis: same as Body
- Code: `Consolas, "Courier New", monospace`

### Font Size Hierarchy

**Baseline**: Body = 20px (medium density)

| Purpose | Ratio to body | @ body=20 | Weight |
| ------- | ------------- | --------- | ------ |
| Cover title | 3x | 60px | Bold |
| Chapter / section opener | 2.2x | 44px | Bold |
| Page title | 1.7x | 34px | Bold |
| Hero number | 2x | 40px | Bold |
| Subtitle | 1.3x | 26px | SemiBold |
| **Body content** | **1x** | **20px** | Regular |
| Annotation / caption | 0.75x | 15px | Regular |
| Page number / footnote | 0.6x | 12px | Regular |

---

## V. Layout Principles

### Page Structure

- **Header area**: y=50, h=80 — title + subtitle + accent bar
- **Content area**: y=130, h=520 — main content
- **Footer area**: y=670, h=30 — page number + branding

### Layout Pattern Library

| Pattern | Suitable Scenarios |
| ------- | ----------------- |
| **Single column centered** | Covers, conclusions, hero quotes |
| **Asymmetric split (3:7)** | Chart + takeaway |
| **Three-column cards** | Parallel points, feature lists |
| **Top-bottom split** | Timelines, process flows |
| **Z-pattern / waterfall** | Storytelling, case studies |
| **Center-radiating** | Ecosystem diagrams |
| **Full-bleed + floating text** | Breathing pages, emphasis |

### Spacing Specification

**Universal**:

| Element | Value |
| ------- | ----- |
| Safe margin from canvas edge | 60px |
| Content block gap | 30px |
| Icon-text gap | 12px |

**Card-based layouts**:

| Element | Value |
| ------- | ----- |
| Card gap | 24px |
| Card padding | 24px |
| Card border radius | 12px |

---

## VI. Icon Usage Specification

### Source

- **Built-in icon library**: `tabler-filled`
- **Brand library**: `simple-icons` (for brand logos only)

### Icon Inventory

| Purpose | Icon Path | Pages |
| ------- | --------- | ----- |
| Speed/Efficiency | `tabler-filled/bolt` | P03, P06 |
| Global | `tabler-filled/world` | P11 |
| Time | `tabler-filled/clock` | P05 |
| Coding | `tabler-filled/code-circle-2` | P04 |
| Success/Check | `tabler-filled/circle-check` | P04, P08 |
| Failure/Warning | `tabler-filled/circle-x` | P13 |
| Alert | `tabler-filled/alert-triangle` | P07 |
| Finance | `tabler-filled/coin-yuan` | P09 |
| Star/Excellence | `tabler-filled/star` | P04 |
| Shield | `tabler-filled/shield` | P07 |
| Target | `tabler-filled/target` | P06 |
| Growth | `tabler-filled/arrow-up-circle` | P09 |
| Fire/Trending | `tabler-filled/flame` | P02 |
| Cash | `tabler-filled/cash-banknote` | P10 |
| Right arrow | `tabler-filled/arrow-right-circle` | P11 |

---

## VII. Visualization Reference List

Catalog read: 71 templates

| Page | Template | Path | Summary-quote (verbatim) | Usage |
| ---- | -------- | ---- | ------------------------- | ----- |
| P04 | horizontal_bar_chart | `templates/charts/horizontal_bar_chart.svg` | "Pick for ranking 5-12 items, especially with long labels. Skip if <=8 short-label items (use bar_chart)." | Benchmark排名对比 |
| P09 | timeline | `templates/charts/timeline.svg` | "Pick for 3-8 milestone events on a horizontal time axis (no duration). Skip for tasks with start/end ranges (use gantt_chart) or vertical layout (use roadmap_vertical)." | 估值飙升时间线 |

**Runners-up considered**:
- `bar_chart` | rejected for P04: labels较长（benchmark名称），horizontal更清晰
- `line_chart` | rejected for P09: 需要突出里程碑事件而非连续趋势
- `waterfall_chart` | rejected for P09: 估值不是增减分解，而是时间点跳跃

---

## VIII. Image Resource List

无图片需求。数据图表和信息图为主，使用占位符标记需要截图的位置。

---

## IX. Content Outline

### Part 1: 开篇

#### Slide 01 - 封面

- **Layout**: Single column centered, decorative gradient
- **Title**: Kimi K3
- **Subtitle**: 一个中国90后造的AI，正在改写全球科技格局
- **Info**: AI大家学 · 前沿认知篇 | 2026.08

#### Slide 02 - Hook: 72小时，4700亿

- **Layout**: Full-bleed breathing page, large hero number
- **Title**: 72小时
- **Content**:
  - 4700亿美元蒸发（hero number）
  - 2026.7.17 上海WAIC → 全球科技股跳水 → 白宫指控
  - 杨植麟，92年，月之暗面
  - "这不是科幻小说，这是真实发生的事"

### Part 2: 实力篇

#### Slide 03 - 硬参数：全球开源之王

- **Layout**: 4-card grid (2×2)
- **Title**: Kimi K3 到底强在哪？
- **Content**:
  - Card 1: 2.8万亿参数 — 全球最大开源模型
  - Card 2: 896专家MoE — 每次激活16个
  - Card 3: 100万Token上下文
  - Card 4: ~$15/百万Token — 便宜2-3倍

#### Slide 04 - 成绩单：全球第三，开源第一

- **Layout**: horizontal_bar_chart adapted
- **Title**: 全球第三，开源第一
- **Visualization**: horizontal_bar_chart
- **Content**:
  - 综合排名全球第三（仅次Claude Fable 5、GPT-5.6）
  - 软件工程排名第一
  - 前端开发全球登顶
  - 智能体工作排名第二
  - Terminal Bench 88.3分

#### Slide 05 - 2小时 vs 1-2周

- **Layout**: Breathing page, asymmetric split
- **Title**: 真实案例：天体物理研究
- **Content**:
  - Left: K3 — 2小时（大数字 + 20篇论文/300种方程/3000行代码）
  - Right: 人类 — 1-2周
  - Bottom: "连马斯克都点了赞"

### Part 3: 技术篇

#### Slide 06 - KDA技术秘密

- **Layout**: Z-pattern, visual diagram
- **Title**: 9年来首次重写AI底层机制
- **Content**:
  - KDA（Kimi Delta Attention）注意力机制
  - 75% KDA线性注意力 + 25% 完整注意力（3:1混合）
  - 训练成本仅460万美元（美国同级上千万至上亿）
  - "同样的效果，成本砍了两到三倍"

### Part 4: 争议篇

#### Slide 07 - 蒸馏风波

- **Layout**: Breathing page, narrative storytelling
- **Title**: 白宫指控"窃取技术"
- **Content**:
  - Michael Kratsios公开指控蒸馏Claude
  - 美财长声称检出"水印"
  - "中国AI偷技术"叙事铺天盖地

#### Slide 08 - 三大反驳

- **Layout**: 3-column cards with check icons
- **Title**: 事情很快反转了
- **Content**:
  - Card 1: 时间线对不上
  - Card 2: 架构是原创的（896专家MoE + KDA）
  - Card 3: 蒸馏做不出原创突破
  - Bottom: "每次中国AI突破，第一反应都是'你作弊了'"

### Part 5: 资本篇

#### Slide 09 - 资本狂飙

- **Layout**: timeline chart adapted
- **Title**: 从30亿到500亿，8个月翻10倍+
- **Visualization**: timeline
- **Content**:
  - 2024.8 ~30亿 → 2026.1 48亿 → 2026上半年 200→315→350亿
  - F轮35亿美元，超募3倍
  - Pre-IPO估值500亿

#### Slide 10 - IPO与关键数据

- **Layout**: KPI cards (2×2)
- **Title**: 月之暗面：资本市场的宠儿
- **Content**:
  - 累计融资376亿元
  - ARR突破3亿美元
  - 拟募资30亿美元IPO
  - 投资方：红杉/真格/IDG/阿里/腾讯

### Part 6: 格局篇

#### Slide 11 - 全球AI格局改写

- **Layout**: 3-column with icons
- **Title**: K3正在改写全球AI格局
- **Content**:
  - 各国可本地部署顶级AI
  - 韩国成"主权AI试验场"
  - 从"比谁聪明"到"比谁用得多"
  - 开源排行榜前六全是中国模型

#### Slide 12 - 中国AI闭环

- **Layout**: Breathing page, center-radiating
- **Title**: 中国AI基础设施闭环加速形成
- **Content**:
  - 国产芯片（华为昇腾）
  - 开源模型（Kimi K3）
  - 部署框架（vLLM）
  - "自己的芯片 + 自己的模型 + 自己的框架 = 完整闭环"

### Part 7: 收尾

#### Slide 13 - 不完美的真相

- **Layout**: Dense, pros/cons style
- **Title**: 不完美的真相
- **Content**:
  - ⚠️ 幻觉率51%（但Opus 5也是50%）
  - ⚠️ 复杂编程与Claude有差距
  - ⚠️ 模型1.4TB，部署门槛高
  - 核心价值："中国团队在最前沿做出了原创性贡献"

#### Slide 14 - 结语

- **Layout**: Single column centered, anchor page
- **Title**: 中国人能做出全球最好的大模型之一
- **Content**:
  - "从GPT到DeepSeek，从DeepSeek到Kimi K3"
  - "竞争越激烈，模型越好用，价格越便宜"
  - 频道引导：关注AI大家学
  - B站/抖音搜索"AI大家学"

---

## X. Speaker Notes Requirements

- **Filename**: match SVG names (e.g., `01_封面.md`)
- **Style**: 对话式叙述，像和朋友聊天
- **Duration**: ~15-20分钟完整演讲
- **Language**: 中文

---

## XI. Technical Constraints Reminder

### SVG Generation Must Follow:

1. viewBox: `0 0 1280 720`
2. Background uses `<rect>` elements
3. Text wrapping uses `<tspan>` (`<foreignObject>` FORBIDDEN)
4. Transparency uses `fill-opacity` / `stroke-opacity`; `rgba()` FORBIDDEN
5. FORBIDDEN: `mask`, `<style>`, `class`, `foreignObject`
6. FORBIDDEN: `textPath`, `animate*`, `script`
7. Text characters: raw Unicode; XML reserved chars escaped
8. `marker-start` / `marker-end` conditionally allowed
9. `clipPath` only on `<image>` elements

### PPT Compatibility Rules:

- `<g opacity="...">` FORBIDDEN; set on each child individually
- Image transparency uses overlay mask layer
- Inline styles only; external CSS and `@font-face` FORBIDDEN
