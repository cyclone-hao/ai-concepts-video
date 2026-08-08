# Design Specification — E01 全景串联 v4

## I. Project Info

| Field | Value |
|-------|-------|
| Project Name | AI概念打通 — 一期视频打通AI底层逻辑 |
| Source | 脚本-v3.md (722 lines, ~15-18 min video) |
| Canvas | PPT 16:9 (1280×720) |
| Pages | 38 |
| Audience | B站/抖音观众：AI小白到开发者 |
| Style | A) General Versatile — 米色系温馨教育风格 |
| Language | Chinese (Simplified) |

## II. Canvas Format

PPT 16:9 — 1280×720px

## III. Visual Theme

Warm beige educational style, consistent with channel's established visual identity (Kimi K3 deck). No dark/gray/black backgrounds (except 镜1 typewriter which uses black bg intentionally).

### Color Palette

| Role | HEX | Name |
|------|-----|------|
| Background | `#F5F0E8` | 米白 |
| Card Background | `#FAF7F2` | 奶白 |
| Code Block BG | `#F0EBE0` | 浅米 |
| Primary | `#8B6F47` | 暖棕 |
| Accent | `#E8A87C` | 柔和橙 |
| Dark Accent | `#C4956A` | 深暖棕 |
| Body Text | `#5D4E37` | 深棕 |
| Secondary Text | `#8B7D6B` | 辅助棕 |
| Border | `#EDE5D5` | 浅米边 |
| Success/Positive | `#8FBC8F` | 柔绿 |
| Error/Warning | `#D4837F` | 柔红 |
| Section BG Gradient Start | `#8B6F47` | 章节渐变起 |
| Section BG Gradient End | `#6B5535` | 章节渐变终 |

## IV. Typography

| Role | Font Stack | Size | Weight |
|------|-----------|------|--------|
| Title | `"Microsoft YaHei", "PingFang SC", Georgia, serif` | 34px | bold |
| Subtitle | `"Microsoft YaHei", "PingFang SC", Arial, sans-serif` | 20-24px | normal/bold |
| Body | `"Microsoft YaHei", "PingFang SC", Arial, sans-serif` | 18px | normal |
| Small/Annotation | `"Microsoft YaHei", "PingFang SC", Arial, sans-serif` | 12-14px | normal |
| Code | `Consolas, "Courier New", monospace` | 14-16px | normal |
| Hero Number | `Georgia, "Microsoft YaHei", serif` | 48-72px | bold |
| Section Title (章节页) | `"Microsoft YaHei", "PingFang SC", Georgia, serif` | 48-52px | bold |

### Size Ramp (anchored on body=18px)

| Slot | Ratio | Size |
|------|-------|------|
| Hero | 2.5-4x | 48-72px |
| Title | 1.6-2x | 28-36px |
| Subtitle | 1.2-1.4x | 20-24px |
| Body | 1x | 18px |
| Annotation | 0.7-0.8x | 12-14px |
| Code | 0.8-0.9x | 14-16px |

## V. Layout Templates

No layout template selected — free design for all pages.

## VI. Icon Strategy

| Setting | Value |
|---------|-------|
| Library | `tabler-filled` |
| Stroke Width | N/A (filled library) |
| Style | Smooth, rounded, organic contours |

### Icon Inventory (approved)

- `tabler-filled/message` — 对话/消息
- `tabler-filled/brain` — 推理/思考
- `tabler-filled/bolt` — 速度/效率
- `tabler-filled/database` — 数据库/存储
- `tabler-filled/search` — 搜索/检索
- `tabler-filled/settings` — 配置/工具
- `tabler-filled/shield` — 安全/治理
- `tabler-filled/eye` — 观察
- `tabler-filled/refresh` — 循环/刷新
- `tabler-filled/bulb` — 洞察/提示
- `tabler-filled/world` — 全球/网络
- `tabler-filled/code` — 代码/编程
- `tabler-filled/file-text` — 文档
- `tabler-filled/device-desktop` — 设备/部署
- `tabler-filled/cloud` — 云端
- `tabler-filled/cash-banknote` — 费用/成本
- `tabler-filled/clock` — 时间
- `tabler-filled/star` — 星标
- `tabler-filled/rocket` — 发射/前沿
- `tabler-filled/brand-github` — GitHub
- `tabler-filled/layers` — 层级/堆叠
- `tabler-filled/box` — 模块/组件
- `tabler-filled/link` — 连接/链接
- `tabler-filled/book` — 知识/学习
- `tabler-filled/check` — 正确/成功
- `tabler-filled/x` — 错误/失败
- `tabler-filled/alert-triangle` — 警告
- `tabler-filled/arrow-right` — 递进/下一步
- `tabler-filled/plug` — 接口/协议
- `tabler-filled/puzzle` — 技能/模块
- `tabler-filled/chart-bar` — 数据/图表
- `tabler-filled/users` — 用户/协作

## VII. Visualization Reference List

| Page | Visualization | Chart Template | Notes |
|------|--------------|----------------|-------|
| P05 | 预测下一个词概率 | Custom bubble/bar | Three candidate words with probability % |
| P07 | 五层堆叠图 | module_composition | Five stacked layers, bottom→top |
| P16 | KV Cache 柱状图 | column_chart | X: rounds 1-10, Y: memory GB |
| P19 | Few-shot + CoT 对比 | Custom | Side-by-side accuracy comparison |
| P20 | Lost in the Middle | Custom horizontal bar | Green→Red→Green attention gradient |
| P23 | Embedding 2D 向量 | Custom scatter | Word clusters in 2D space |
| P25 | RAG 三步流程 | pipeline_with_stages | 建库→检索→生成 |
| P28 | Function Calling 流程 | process_flow | 4-step horizontal flow |
| P31 | MCP N×M vs N+M | Custom topology | Messy vs star topology |
| P33 | ReAct 循环 | circular_stages | 观察→思考→行动→观察 loop |
| P37 | Harness 六部件 | hub_spoke | LLM center, 6 components around |
| P38 | 进化时间轴 | timeline_horizontal | 2 nodes: Prompt Eng → Harness Eng |
| P41 | 全局串联图 | process_flow | 12-node progressive chain |

## VIII. Image Resource List

| ID | Description | Acquire Via | Status | Notes |
|----|-------------|-------------|--------|-------|
| img-01 | OpenClaw 龙虾 Logo | placeholder | Placeholder | Dashed border placeholder, user replaces later |
| img-02 | 大厂"免费安装OpenClaw"新闻截图 | placeholder | Placeholder | 3 news card placeholders |

## IX. Content Outline & Page Roster

### P01 — 打字机开场 (anchor)
Black background, white typewriter text. Four lines appearing sequentially:
- 2023: ChatGPT → 科技达人
- 2024: Prompt → 行业专家
- 2025: Agent → 技术大牛
- 2026: Harness/OpenClaw/Hermes → ???(highlighted in #E8A87C)

### P02 — 开场主视觉 (anchor)
Warm beige background. Title: "从 GPT、DeepSeek 到 OpenClaw、Hermes". Subtitle: "一期视频打通 AI 底层逻辑". Keyword tag cloud below.

### P03 — OpenClaw龙虾梗 (breathing)
OpenClaw logo placeholder on right. 3 news card placeholders on left. Brief text explaining the "养虾" phenomenon.

### P04 — 01 LLM标题 (anchor)
Section number "01" in large warm brown. Title: "LLM — 大语言模型". Core definition: "给一段文字，预测下一个最可能出现的词".

### P05 — 预测下一个词 (dense)
Sentence: "今天天气真好，我们去公园___". Three candidate word bubbles with probabilities: 散步(45%), 玩(30%), 跑步(15%).

### P06 — Attention连线图 (dense)
Sentence: "小明在北京参加了一场关于___的学术研讨会". Two curved arc lines from "北京" and "学术研讨会" converging to the blank. "Attention" label on arcs.

### P07 — 五层堆叠图 (dense)
Five stacked rounded rectangles from bottom to top:
1. 自回归架构 — 让它能学 (#8B6F47, darkest)
2. 海量预训练 — 让它学得广 (4500亿Token)
3. 代码训练 — 让它学得深 (逻辑推理)
4. SFT指令微调 — 让它会对话
5. RLHF — 让它说得好 (lightest)

### P08 — 02 Token标题 (anchor)
Section number "02". Title: "Token — 模型的最小阅读单位". Core: "模型不认识'字'，它认识的是Token".

### P09 — BPE切分与合并 (dense)
Top: "我喜欢吃苹果" split into 6 colored blocks. Bottom: BPE merge animation — "人"+"工"→"人工"→"人工智能".

### P10 — Tokenizer对比 (dense)
Three-row comparison: Qwen(6 tokens), GPT-4(5 tokens, "喜欢" merged), Claude(6 tokens). Colored blocks for each token.

### P11 — Token计费 (dense)
Calculator-style card. DeepSeek-V3 pricing. Input: 50 Token × ¥2/1M. Output: 500 Token × ¥8/1M. Total: ¥0.004. Monthly: ~¥12,000.

### P12 — 特殊Token (breathing)
Token sequence with special tokens highlighted: `<bos>` | 我 | 喜欢 | 吃 | 苹果 | `<eos>` | `<pad>`. Annotation bubbles below each special token.

### P13 — 03 Context标题 (anchor)
Section "03". Title: "Context — 大模型没有'记忆'". Core: "每次对话都是失忆的".

### P14 — 多轮对话messages增长 (dense)
Code editor card showing messages array growing from 1 message (round 1) to 3 messages (round 2). Dashed outline showing "previous content resent".

### P15 — Context Window信息 (dense)
Four model cards horizontally: GPT-5.6(105万), Claude Opus 5(100万), Gemini(100万+), DeepSeek-V3/Kimi K3(128K).

### P16 — KV Cache柱状图 (dense)
Column chart. X: conversation rounds 1-10. Y: KV Cache memory. Bars grow from ~1GB to ~20GB+. Gradient from light to dark warm brown.

### P17 — 04 Prompt标题 (anchor)
Section "04". Title: "Prompt — 你怎么问，决定它怎么答".

### P18 — Prompt对比 (dense)
Left-right comparison. Left: vague prompt "帮我写个排序算法" (light border, ❓). Right: precise prompt (accent border, ✅).

### P19 — Few-shot + CoT (dense)
Top: Few-shot translation examples (你好→Hello, 谢谢→Thank you, 再见→Goodbye). Bottom: CoT accuracy comparison 17.7% → 78.7%.

### P20 — Lost in the Middle (breathing)
Horizontal bar divided into 3 zones: start(green ✅) → middle(red ⚠️) → end(green ✅). "128K ≠ 完美处理128K" label.

### P21 — System Prompt结构 (dense)
Code card showing messages array. System message highlighted in #E8A87C. User message in light bg. "新员工手册" metaphor icon on right.

### P22 — 05 Embedding标题 (anchor)
Section "05". Title: "Embedding — 让文字变成'坐标'". Bridge: "Token只是编号，Embedding赋予语义".

### P23 — 二维词向量可视化 (dense)
2D coordinate system with word dots. Fruit cluster (苹果, 香蕉, 橙子, 水果) upper-left. Vehicle cluster (汽车, 卡车, 公交) lower-right. Distance labels: 苹果→水果="近"(green), 苹果→汽车="远"(red).

### P24 — 06 RAG标题 (anchor)
Section "06". Title: "RAG — 让AI查资料再回答". "Retrieval-Augmented Generation".

### P25 — RAG三步流程图 (dense)
Three-step horizontal pipeline:
1. 建库: 文档→切片→向量→向量数据库
2. 检索: 问题→向量→语义搜索→Top-K文档
3. 生成: 文档+问题→LLM→回答
Bottom: "RAG = 搜索引擎 + 大模型"

### P26 — RAG vs Fine-tuning (dense)
Left-right comparison. Left: RAG = 开卷考试 (成本低, 可更新, 有引用). Right: Fine-tuning = 上课学习 (更快更自然). Bottom: "实际项目经常结合使用".

### P27 — 07 Tool标题 (anchor)
Section "07". Title: "Tool — 从'会说'到'能做'".

### P28 — Function Calling四步 (dense)
Four-step flow with code:
1. 用户提问 → 2. 模型调工具 → 3. 程序执行 → 4. 生成回答
Code snippets for each step below.

### P29 — SFT训练数据 (dense)
Code card showing SFT training data format. Three annotation arrows: ① 什么时候调 ② 调哪个 ③ 参数怎么填.

### P30 — 08 MCP标题 (anchor)
Section "08". Title: "MCP — AI世界的USB-C". "Model Context Protocol".

### P31 — N×M vs N+M (dense)
Left: messy connection lines (N×M = 适配地狱). Right: star topology with MCP Hub (N+M = 一次开发处处运行). Bottom: 高德地图 example.

### P32 — 09 Agent标题 (anchor)
Section "09". Title: "Agent — 让AI自己决定下一步".

### P33 — ReAct循环 (dense)
Left: circular flow diagram (观察→思考→行动→观察). Right: Agent execution log (竞品调研 scenario).

### P34 — 10 Skill标题 (anchor)
Section "10". Title: "Agent Skill — 给AI一本'操作手册'".

### P35 — Skill比喻+结构 (dense)
Top: Left (scattered notes = 即兴发言) vs Right (structured doc = 排练过的演讲稿). Bottom: 4 elements (触发条件, 执行流程, 约束规则, 输出模板).

### P36 — 11 Harness标题 (anchor)
Section "11". Title: "Harness Engineering — 让Agent靠谱地工作". "Mitchell Hashimoto, 2026".

### P37 — Harness六部件 (dense)
Center: "LLM" circle. 6 surrounding component cards:
1-4 (warm, 能力): System Prompt, Skills, Tools, Context Window
5-6 (green, 保障): Governance, Observability
Bottom: "Agent = Model + Harness"

### P38 — 进化时间轴 (dense)
Horizontal timeline, 2 nodes:
1. 2023-2024: Prompt Engineering "怎么写好一条指令"
2. 2025-2026: Harness Engineering "怎么搭一套系统让AI可靠运行"

### P39 — 12 前沿标题 (anchor)
Section "12". Title: "2026前沿 — OpenClaw · Hermes · Vibe Coding".

### P40 — 三大前沿项目 (dense)
Three vertical cards:
1. OpenClaw: 🦞 400K+ Stars, 4.4万技能
2. Hermes: 🧠 220K+ Stars, 自我学习循环
3. Vibe Coding: ⚡ 用嘴写代码

### P41 — 全局串联图 (dense)
12-node progressive chain in two rows:
LLM → Token → Context → Prompt → Embedding → RAG → Tool → MCP → Agent → Skill → Harness → 2026前沿
Arrow labels explain why each layer needs the next.

### P42 — 结尾 (anchor)
Ending text: "如果这期内容对你有帮助，别忘了三连！". Channel info. Next episode preview: "大模型技术篇 · Agent底层原理".

## X. Speaker Notes

Notes style: conversational narration matching video voiceover. Natural Chinese, matching the script's tone. Each page's notes = 2-5 sentences of pure spoken narration, with natural transitions between pages.

## XI. Technical Constraints

- All SVGs: 1280×720, viewBox="0 0 1280 720"
- No banned features (no mask, no style, no class, no animate, no foreignObject)
- All fonts must end with pre-installed families
- Icons: tabler-filled only, placeholder syntax
- Colors: HEX only, transparency via fill-opacity/stroke-opacity
- Group elements semantically with `<g id="...">` (3-8 content groups per slide)
