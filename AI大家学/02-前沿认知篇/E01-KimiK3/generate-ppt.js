const pptxgen = require('pptxgenjs');
const path = require('path');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE'; // 16:9

// ========== 胡桃色温馨配色 ==========
const C = {
  bg:       'F5F0E8',
  card:     'FAF7F2',
  primary:  '8B6F47',
  accent:   'E8A87C',
  teal:     '7BA7A7',
  text:     '5D4E37',
  subtext:  '8B7D6B',
  border:   'EDE5D5',
  white:    'FFFFFF',
  red:      'C25B56',
  green:    '7BA77B',
  purple:   '8B7BA7',
};

// ========== 辅助函数 ==========
function addBg(slide) {
  slide.background = { color: C.bg };
  // 每页统一过渡效果
  slide.transition = { type: 'fade', speed: 0.8 };
}

function addFooter(slide, text) {
  slide.addText(text, { x:0, y:6.8, w:13.33, h:0.3, fontSize:8, color:C.subtext, align:'center', fontFace:'Microsoft YaHei' });
}
function addSectionNum(slide, num) {
  slide.addText(num, { x:0.6, y:0.8, w:1.5, h:1.2, fontSize:60, bold:true, color:C.border, fontFace:'Georgia',
    animation: { type:'wipe', delay:0.0, duration:0.6, dir:'fromRight' }
  });
}

// ========== Slide 1: 封面 ==========
let s1 = pptx.addSlide();
addBg(s1);
s1.addShape(pptx.ShapeType.rect, { x:0, y:0, w:13.33, h:7.5, fill:{ color:C.bg } });
s1.addShape(pptx.ShapeType.rect, { x:1.5, y:2.2, w:4, h:0.04, fill:{ color:C.accent },
  animation: { type:'wipe', delay:0.0, duration:0.5, dir:'fromLeft' }
});
s1.addText('Kimi K3', { x:1.5, y:2.5, w:10, h:1.5, fontSize:54, bold:true, color:C.primary, fontFace:'Georgia',
  animation: { type:'fade', delay:0.3, duration:0.8 }
});
s1.addText('当开源模型第一次逼到闭源喉咙口', { x:1.5, y:3.8, w:10, h:0.8, fontSize:22, color:C.text, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.8, duration:0.6 }
});
s1.addText('AI大家学 · 前沿AI认知篇', { x:1.5, y:5.0, w:10, h:0.5, fontSize:14, color:C.subtext, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:1.2, duration:0.5 }
});
s1.addShape(pptx.ShapeType.ellipse, { x:9.5, y:1.0, w:3.5, h:3.5, fill:{ color:C.card }, line:{ color:C.border, width:1.5 },
  animation: { type:'zoom', delay:0.5, duration:0.8 }
});
s1.addText('2.8T', { x:9.5, y:2.0, w:3.5, h:0.8, fontSize:36, bold:true, color:C.accent, align:'center', fontFace:'Georgia',
  animation: { type:'fade', delay:1.0, duration:0.6 }
});
s1.addText('参数', { x:9.5, y:2.8, w:3.5, h:0.5, fontSize:14, color:C.subtext, align:'center', fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:1.2, duration:0.5 }
});
addFooter(s1, '数据来源：Artificial Analysis · Frontend Code Arena · 月之暗面官方技术报告');

// ========== Slide 2: Hook - 4700亿 ==========
let s2 = pptx.addSlide();
addBg(s2);
s2.transition = { type:'fade', speed:0.5 };
s2.addText('72小时', { x:1.5, y:1.5, w:5, h:1.0, fontSize:44, bold:true, color:C.primary, fontFace:'Georgia',
  animation: { type:'wipe', delay:0.0, duration:0.6, dir:'fromLeft' }
});
s2.addShape(pptx.ShapeType.rect, { x:1.5, y:2.5, w:8, h:0.03, fill:{ color:C.accent },
  animation: { type:'wipe', delay:0.4, duration:0.5, dir:'fromLeft' }
});
s2.addText('-$4,700 亿', { x:1.5, y:2.8, w:10, h:1.5, fontSize:64, bold:true, color:C.red, fontFace:'Georgia',
  animation: { type:'zoom', delay:0.8, duration:0.8 }
});
s2.addText('全球科技股市值蒸发', { x:1.5, y:4.2, w:10, h:0.6, fontSize:20, color:C.text, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:1.5, duration:0.5 }
});
s2.addText([
  { text:'7月16日  ', options:{ color:C.subtext, fontSize:14 } },
  { text:'月之暗面发布 Kimi K3', options:{ color:C.primary, fontSize:14, bold:true } },
], { x:1.5, y:5.2, w:10, h:0.4, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:2.0, duration:0.5 }
});
s2.addText([
  { text:'7月19日  ', options:{ color:C.subtext, fontSize:14 } },
  { text:'华尔街称之为"Kimi时刻"', options:{ color:C.primary, fontSize:14, bold:true } },
], { x:1.5, y:5.7, w:10, h:0.4, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:2.5, duration:0.5 }
});
addFooter(s2, '据多家财经媒体统计');

// ========== Slide 3: 章节 - 它有多强 ==========
let s3 = pptx.addSlide();
addBg(s3);
addSectionNum(s3, '01');
s3.addText('它有多强', { x:2.2, y:1.0, w:8, h:1.0, fontSize:36, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.3, duration:0.6 }
});
s3.addText('2.8万亿参数 · MoE 896专家 · 100万Token上下文', { x:2.2, y:2.0, w:8, h:0.5, fontSize:16, color:C.subtext, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.7, duration:0.5 }
});
s3.addShape(pptx.ShapeType.rect, { x:2.2, y:2.6, w:3, h:0.03, fill:{ color:C.accent },
  animation: { type:'wipe', delay:1.0, duration:0.4, dir:'fromLeft' }
});

// ========== Slide 4: MoE 架构 ==========
let s4 = pptx.addSlide();
addBg(s4);
s4.addText('MoE 混合专家架构', { x:0.8, y:0.4, w:8, h:0.7, fontSize:28, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
s4.addText('知识储备极大，单次推理只调用 1.8% 的专家', { x:0.8, y:1.0, w:10, h:0.4, fontSize:13, color:C.subtext, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.3, duration:0.5 }
});
const moeData = [
  { label:'总参数', value:'2.8万亿', color:C.primary },
  { label:'专家总数', value:'896', color:C.teal },
  { label:'每次激活', value:'16个', color:C.accent },
  { label:'激活比例', value:'1.8%', color:C.accent },
];
moeData.forEach((d, i) => {
  const x = 0.8 + i * 3.1;
  s4.addShape(pptx.ShapeType.roundRect, { x:x, y:1.8, w:2.8, h:1.6, rectRadius:0.15, fill:{ color:C.card }, line:{ color:C.border, width:1 },
    animation: { type:'wipe', delay: 0.5 + i*0.4, duration:0.5, dir:'fromBottom' }
  });
  s4.addText(d.label, { x:x, y:1.95, w:2.8, h:0.5, fontSize:12, color:C.subtext, align:'center', fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 0.7 + i*0.4, duration:0.4 }
  });
  s4.addText(d.value, { x:x, y:2.4, w:2.8, h:0.8, fontSize:28, bold:true, color:d.color, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.8 + i*0.4, duration:0.4 }
  });
});
// 说明区
s4.addShape(pptx.ShapeType.roundRect, { x:0.8, y:3.8, w:11.73, h:2.8, rectRadius:0.15, fill:{ color:C.card }, line:{ color:C.border, width:1 },
  animation: { type:'wipe', delay:2.2, duration:0.6, dir:'fromBottom' }
});
s4.addText('翻译成人话', { x:1.2, y:3.95, w:3, h:0.5, fontSize:16, bold:true, color:C.accent, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:2.5, duration:0.4 }
});
s4.addText('896个专家随时待命，但每次推理只叫16个出来干活。\n\n就像一家有896个专科医生的医院，你每次看病只需要挂16个相关科室。\n\n→ 知识储备极大，但单次成本很低', {
  x:1.2, y:4.5, w:11, h:2.0, fontSize:14, color:C.text, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.4,
  animation: { type:'fade', delay:2.8, duration:0.8 }
});
addFooter(s4, '来源：月之暗面官方技术报告');

// ========== Slide 5: 榜单成绩 ==========
let s5 = pptx.addSlide();
addBg(s5);
s5.addText('成绩单 — 硬指标', { x:0.8, y:0.4, w:8, h:0.7, fontSize:28, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
const benchmarks = [
  { name:'Frontend Code Arena', sub:'前端代码能力', score:'1679', rank:'#1', note:'全球第一 · 开源首次登顶', color:C.teal },
  { name:'AA Intelligence Index', sub:'智能指数', score:'57', rank:'#3', note:'仅次 Fable 5 (60) 和 GPT-5.6 Sol (59)', color:C.accent },
  { name:'BrowseComp', sub:'长周期检索', score:'91.2', rank:'#1', note:'全球第一', color:C.purple },
  { name:'GDPval-AA v2', sub:'真实经济任务', score:'1668', rank:'#4', note:'开源模型最高', color:C.primary },
];
benchmarks.forEach((b, i) => {
  const y = 1.4 + i * 1.3;
  s5.addShape(pptx.ShapeType.roundRect, { x:0.8, y:y, w:11.73, h:1.1, rectRadius:0.12, fill:{ color:C.card }, line:{ color:C.border, width:1 },
    animation: { type:'wipe', delay: 0.3 + i*0.6, duration:0.5, dir:'fromLeft' }
  });
  s5.addShape(pptx.ShapeType.rect, { x:0.8, y:y, w:0.08, h:1.1, fill:{ color:b.color },
    animation: { type:'wipe', delay: 0.5 + i*0.6, duration:0.3, dir:'fromTop' }
  });
  s5.addText(b.name, { x:1.2, y:y+0.1, w:4, h:0.4, fontSize:14, bold:true, color:C.text, fontFace:'Georgia',
    animation: { type:'fade', delay: 0.6 + i*0.6, duration:0.4 }
  });
  s5.addText(b.sub, { x:1.2, y:y+0.5, w:4, h:0.4, fontSize:11, color:C.subtext, fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 0.7 + i*0.6, duration:0.4 }
  });
  s5.addText(b.score, { x:7, y:y+0.15, w:3, h:0.7, fontSize:28, bold:true, color:b.color, align:'right', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.8 + i*0.6, duration:0.4 }
  });
  s5.addShape(pptx.ShapeType.roundRect, { x:10.2, y:y+0.25, w:0.9, h:0.5, rectRadius:0.25, fill:{ color:b.color },
    animation: { type:'zoom', delay: 0.9 + i*0.6, duration:0.3 }
  });
  s5.addText(b.rank, { x:10.2, y:y+0.25, w:0.9, h:0.5, fontSize:14, bold:true, color:C.white, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 1.0 + i*0.6, duration:0.3 }
  });
  s5.addText(b.note, { x:11.3, y:y+0.3, w:2, h:0.4, fontSize:10, color:C.subtext, fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 1.0 + i*0.6, duration:0.3 }
  });
});
addFooter(s5, 'Artificial Analysis · Frontend Code Arena · 月之暗面技术报告');

// ========== Slide 6: 三模型对比 ==========
let s6 = pptx.addSlide();
addBg(s6);
s6.addText('三模型对比', { x:0.8, y:0.4, w:8, h:0.7, fontSize:28, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
const colX = [0.8, 4.8, 8.0, 11.0];
const colW = [3.8, 3.0, 2.8, 1.73];
const models = [
  { name:'对比维度', codeArena:'Code Arena', aaIndex:'AA Index', openClosed:'开源/闭源', color:C.subtext },
  { name:'Kimi K3', codeArena:'1679', aaIndex:'57', openClosed:'开源', color:C.teal },
  { name:'Claude Fable 5', codeArena:'1631', aaIndex:'60', openClosed:'闭源', color:C.accent },
  { name:'GPT-5.6 Sol', codeArena:'1618', aaIndex:'59', openClosed:'闭源', color:C.red },
];
models.forEach((m, i) => {
  const y = 1.4 + i * 1.2;
  const bgColor = i === 0 ? C.primary : C.card;
  const textColor = i === 0 ? C.white : C.text;
  s6.addShape(pptx.ShapeType.roundRect, { x:0.8, y:y, w:11.73, h:1.0, rectRadius: i===0 ? 0.12 : 0.08, fill:{ color:bgColor }, line:{ color:C.border, width: i===0?0:1 },
    animation: { type:'wipe', delay: 0.2 + i*0.5, duration:0.5, dir:'fromLeft' }
  });
  s6.addText(m.name, { x:colX[0], y:y+0.15, w:colW[0], h:0.7, fontSize: i===0?13:16, bold:true, color: i===0?C.white:m.color, fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 0.4 + i*0.5, duration:0.4 }
  });
  s6.addText(m.codeArena, { x:colX[1], y:y+0.15, w:colW[1], h:0.7, fontSize: i===0?12:18, bold:i>0, color:textColor, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.5 + i*0.5, duration:0.4 }
  });
  s6.addText(m.aaIndex, { x:colX[2], y:y+0.15, w:colW[2], h:0.7, fontSize: i===0?12:18, bold:i>0, color:textColor, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.6 + i*0.5, duration:0.4 }
  });
  s6.addText(m.openClosed, { x:colX[3], y:y+0.15, w:colW[3], h:0.7, fontSize: i===0?12:13, color: i===0?C.white:m.color, align:'center', fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 0.7 + i*0.5, duration:0.4 }
  });
});
// 结论
s6.addShape(pptx.ShapeType.roundRect, { x:0.8, y:6.0, w:11.73, h:0.7, rectRadius:0.1, fill:{ color:C.card }, line:{ color:C.accent, width:1.5 },
  animation: { type:'wipe', delay:2.5, duration:0.5, dir:'fromBottom' }
});
s6.addText('开源第一，综合第三，性价比维度几乎没有对手', { x:0.8, y:6.05, w:11.73, h:0.6, fontSize:16, bold:true, color:C.accent, align:'center', fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:2.8, duration:0.5 }
});

// ========== Slide 7: I-Love-Q 时钟对比 ==========
let s7 = pptx.addSlide();
addBg(s7);
s7.addText('真实生产力 — I-Love-Q 中子星研究', { x:0.8, y:0.4, w:10, h:0.7, fontSize:24, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
// K3 卡片 - 从左飞入
s7.addShape(pptx.ShapeType.roundRect, { x:0.8, y:1.4, w:5.5, h:5.0, rectRadius:0.15, fill:{ color:C.card }, line:{ color:C.teal, width:2 },
  animation: { type:'fly', delay:0.3, duration:0.7, dir:'fromLeft' }
});
s7.addShape(pptx.ShapeType.rect, { x:0.8, y:1.4, w:5.5, h:0.6, fill:{ color:C.teal }, rectRadius:0.15,
  animation: { type:'wipe', delay:0.8, duration:0.3, dir:'fromLeft' }
});
s7.addText('Kimi K3', { x:0.8, y:1.42, w:5.5, h:0.55, fontSize:16, bold:true, color:C.white, align:'center', fontFace:'Georgia',
  animation: { type:'fade', delay:1.0, duration:0.4 }
});
s7.addText('~2 小时', { x:0.8, y:2.5, w:5.5, h:1.0, fontSize:44, bold:true, color:C.teal, align:'center', fontFace:'Georgia',
  animation: { type:'zoom', delay:1.2, duration:0.6 }
});
s7.addText('自主阅读 20+ 篇论文\n评估 300+ 种状态方程\n生成 3000+ 行 Python\n马斯克评价："Impressive"', {
  x:1.2, y:3.6, w:4.7, h:2.5, fontSize:13, color:C.text, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.6,
  animation: { type:'fade', delay:1.6, duration:0.8 }
});
// VS - 缩放出现
s7.addText('VS', { x:5.8, y:3.5, w:1.5, h:0.8, fontSize:20, bold:true, color:C.subtext, align:'center', fontFace:'Georgia',
  animation: { type:'zoom', delay:2.0, duration:0.5 }
});
// 人类卡片 - 从右飞入
s7.addShape(pptx.ShapeType.roundRect, { x:7, y:1.4, w:5.5, h:5.0, rectRadius:0.15, fill:{ color:C.card }, line:{ color:C.accent, width:2 },
  animation: { type:'fly', delay:2.3, duration:0.7, dir:'fromRight' }
});
s7.addShape(pptx.ShapeType.rect, { x:7, y:1.4, w:5.5, h:0.6, fill:{ color:C.accent }, rectRadius:0.15,
  animation: { type:'wipe', delay:2.8, duration:0.3, dir:'fromRight' }
});
s7.addText('资深研究人员', { x:7, y:1.42, w:5.5, h:0.55, fontSize:16, bold:true, color:C.white, align:'center', fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:3.0, duration:0.4 }
});
s7.addText('1-2 周', { x:7, y:2.5, w:5.5, h:1.0, fontSize:44, bold:true, color:C.accent, align:'center', fontFace:'Georgia',
  animation: { type:'zoom', delay:3.2, duration:0.6 }
});
s7.addText('需要领域专业知识积累\n手动检索和筛选文献\n编写和调试计算代码\n反复验证状态方程', {
  x:7.4, y:3.6, w:4.7, h:2.5, fontSize:13, color:C.text, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.6,
  animation: { type:'fade', delay:3.6, duration:0.8 }
});
addFooter(s7, '来源：月之暗面官方演示');

// ========== Slide 8: 章节 - 怎么做到的 ==========
let s8 = pptx.addSlide();
addBg(s8);
addSectionNum(s8, '02');
s8.addText('怎么做到的', { x:2.2, y:1.0, w:8, h:1.0, fontSize:36, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.3, duration:0.6 }
});
s8.addText('不是堆参数，是改地基', { x:2.2, y:2.0, w:8, h:0.5, fontSize:16, color:C.subtext, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.7, duration:0.5 }
});
s8.addShape(pptx.ShapeType.rect, { x:2.2, y:2.6, w:3, h:0.03, fill:{ color:C.accent },
  animation: { type:'wipe', delay:1.0, duration:0.4, dir:'fromLeft' }
});

// ========== Slide 9: 三大技术 ==========
let s9 = pptx.addSlide();
addBg(s9);
s9.addText('三项自研底层技术', { x:0.8, y:0.4, w:10, h:0.7, fontSize:28, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
s9.addText('对沿用近十年的 AI 基础机制做系统性重写', { x:0.8, y:1.0, w:10, h:0.4, fontSize:13, color:C.subtext, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.3, duration:0.5 }
});
const techs = [
  { name:'KDA 混合线性注意力', desc:'不同通道 × 不同衰减速率\n管理长序列记忆\n+ Attention Residuals', result:'推理效率 ↑ ~25%', color:C.teal },
  { name:'Kimi Linear Tension', desc:'自研线性注意力机制\n解决超长任务性能衰减\n上下文窗口 ×10', result:'长上下文的工程底气', color:C.accent },
  { name:'Moon Clip 二阶优化器', desc:'首次用于大模型训练\n20T 数据跑出 40T 效果\n训练成本和算力功耗减半', result:'训练效率翻倍', color:C.purple },
];
techs.forEach((t, i) => {
  const x = 0.8 + i * 4.1;
  s9.addShape(pptx.ShapeType.roundRect, { x:x, y:1.7, w:3.8, h:4.5, rectRadius:0.15, fill:{ color:C.card }, line:{ color:C.border, width:1 },
    animation: { type:'wipe', delay: 0.5 + i*0.7, duration:0.6, dir:'fromBottom' }
  });
  s9.addShape(pptx.ShapeType.rect, { x:x, y:1.7, w:3.8, h:0.55, fill:{ color:t.color }, rectRadius:0.15,
    animation: { type:'wipe', delay: 0.8 + i*0.7, duration:0.3, dir:'fromLeft' }
  });
  s9.addText(t.name, { x:x, y:1.72, w:3.8, h:0.5, fontSize:14, bold:true, color:C.white, align:'center', fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 0.9 + i*0.7, duration:0.4 }
  });
  s9.addText(t.desc, { x:x+0.3, y:2.6, w:3.2, h:2.5, fontSize:12, color:C.text, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.6,
    animation: { type:'fade', delay: 1.1 + i*0.7, duration:0.5 }
  });
  s9.addShape(pptx.ShapeType.rect, { x:x+0.3, y:4.8, w:3.2, h:0.02, fill:{ color:C.border },
    animation: { type:'wipe', delay: 1.3 + i*0.7, duration:0.3, dir:'fromLeft' }
  });
  s9.addText(t.result, { x:x, y:5.0, w:3.8, h:0.6, fontSize:14, bold:true, color:t.color, align:'center', fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 1.4 + i*0.7, duration:0.4 }
  });
});
// 结果条
s9.addShape(pptx.ShapeType.roundRect, { x:2, y:6.4, w:9.33, h:0.6, rectRadius:0.1, fill:{ color:C.card }, line:{ color:C.accent, width:1.5 },
  animation: { type:'wipe', delay:3.8, duration:0.5, dir:'fromBottom' }
});
s9.addText('三项叠加 → 整体扩展效率提升 ~2.5 倍', { x:2, y:6.42, w:9.33, h:0.55, fontSize:16, bold:true, color:C.accent, align:'center', fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:4.1, duration:0.5 }
});

// ========== Slide 10: 章节 - 为什么全球紧张 ==========
let s10 = pptx.addSlide();
addBg(s10);
addSectionNum(s10, '03');
s10.addText('为什么全球都紧张', { x:2.2, y:1.0, w:8, h:1.0, fontSize:36, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.3, duration:0.6 }
});
s10.addText('算力挤兑 · 蒸馏争议 · 从卡芯片到卡算法', { x:2.2, y:2.0, w:8, h:0.5, fontSize:16, color:C.subtext, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.7, duration:0.5 }
});
s10.addShape(pptx.ShapeType.rect, { x:2.2, y:2.6, w:3, h:0.03, fill:{ color:C.accent },
  animation: { type:'wipe', delay:1.0, duration:0.4, dir:'fromLeft' }
});

// ========== Slide 11: 蒸馏争议 ==========
let s11 = pptx.addSlide();
addBg(s11);
s11.addText('蒸馏争议 — 三条反驳', { x:0.8, y:0.4, w:10, h:0.7, fontSize:28, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
// 指控 - 先出现
s11.addShape(pptx.ShapeType.roundRect, { x:0.8, y:1.3, w:11.73, h:0.9, rectRadius:0.1, fill:{ color:'FDF0F0' }, line:{ color:C.red, width:1 },
  animation: { type:'wipe', delay:0.3, duration:0.5, dir:'fromLeft' }
});
s11.addText('白宫指控：月之暗面蒸馏了 Anthropic 的 Fable 模型，财长威胁制裁', { x:1.2, y:1.35, w:11, h:0.8, fontSize:14, color:C.red, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.6, duration:0.5 }
});
const refutations = [
  { num:'1', title:'性能无法靠蒸馏获得', desc:'K3 排全球第二/第三/前端第一，仅靠蒸馏不可能', color:C.teal },
  { num:'2', title:'开源技术报告可复核', desc:'开放权重 + 全公开架构，蒸馏出来的是别人的架构', color:C.accent },
  { num:'3', title:'时间线对不上', desc:'Fable 发布到 K3 训练完成之间不够走完蒸馏流程', color:C.purple },
];
refutations.forEach((r, i) => {
  const y = 2.6 + i * 1.4;
  s11.addShape(pptx.ShapeType.roundRect, { x:0.8, y:y, w:11.73, h:1.2, rectRadius:0.12, fill:{ color:C.card }, line:{ color:C.border, width:1 },
    animation: { type:'wipe', delay: 1.2 + i*0.8, duration:0.5, dir:'fromLeft' }
  });
  s11.addShape(pptx.ShapeType.ellipse, { x:1.1, y:y+0.2, w:0.7, h:0.7, fill:{ color:r.color },
    animation: { type:'zoom', delay: 1.5 + i*0.8, duration:0.3 }
  });
  s11.addText(r.num, { x:1.1, y:y+0.2, w:0.7, h:0.7, fontSize:20, bold:true, color:C.white, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 1.6 + i*0.8, duration:0.3 }
  });
  s11.addText(r.title, { x:2.2, y:y+0.15, w:6, h:0.45, fontSize:16, bold:true, color:C.text, fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 1.7 + i*0.8, duration:0.4 }
  });
  s11.addText(r.desc, { x:2.2, y:y+0.6, w:9, h:0.45, fontSize:12, color:C.subtext, fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 1.8 + i*0.8, duration:0.4 }
  });
  s11.addText('✓', { x:11.5, y:y+0.25, w:0.8, h:0.7, fontSize:24, bold:true, color:r.color, align:'center', fontFace:'Georgia',
    animation: { type:'zoom', delay: 2.0 + i*0.8, duration:0.3 }
  });
});
addFooter(s11, '来源：Nathan Lambert · TechCrunch · BBC中文');

// ========== Slide 12: 卡芯片到卡算法 ==========
let s12 = pptx.addSlide();
addBg(s12);
s12.addText('从"卡芯片"到"卡算法"', { x:0.8, y:0.4, w:10, h:0.7, fontSize:28, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
// 轴线
s12.addShape(pptx.ShapeType.rect, { x:1.5, y:3.0, w:10.33, h:0.03, fill:{ color:C.border },
  animation: { type:'wipe', delay:0.3, duration:0.6, dir:'fromLeft' }
});
const phases = [
  { name:'卡芯片', year:'2022-2024', desc:'禁运高端GPU', color:C.red },
  { name:'指控蒸馏', year:'2025-2026', desc:'DeepSeek + K3', color:C.accent },
  { name:'实体清单', year:'2026', desc:'威胁制裁', color:C.red },
  { name:'下一步？', year:'', desc:'全面封锁？', color:C.purple },
];
phases.forEach((p, i) => {
  const x = 1.5 + i * 3.3;
  s12.addShape(pptx.ShapeType.roundRect, { x:x, y:1.5, w:2.8, h:1.2, rectRadius:0.1, fill:{ color:C.card }, line:{ color:C.border, width:1 },
    animation: { type:'wipe', delay: 0.6 + i*0.5, duration:0.5, dir:'fromBottom' }
  });
  s12.addText(p.name, { x:x, y:1.55, w:2.8, h:0.5, fontSize:16, bold:true, color:p.color, align:'center', fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 0.8 + i*0.5, duration:0.4 }
  });
  s12.addText(p.year, { x:x, y:2.0, w:2.8, h:0.35, fontSize:10, color:C.subtext, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.9 + i*0.5, duration:0.3 }
  });
  s12.addShape(pptx.ShapeType.ellipse, { x:x+0.6, y:2.85, w:0.3, h:0.3, fill:{ color:p.color },
    animation: { type:'zoom', delay: 1.0 + i*0.5, duration:0.3 }
  });
  s12.addText(p.desc, { x:x, y:3.4, w:2.8, h:0.4, fontSize:11, color:C.subtext, align:'center', fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 1.1 + i*0.5, duration:0.3 }
  });
});
// 底部分析
s12.addShape(pptx.ShapeType.roundRect, { x:0.8, y:4.5, w:11.73, h:2.2, rectRadius:0.12, fill:{ color:C.card }, line:{ color:C.border, width:1 },
  animation: { type:'wipe', delay:3.0, duration:0.6, dir:'fromBottom' }
});
s12.addText('Pattern：你强了，就说你偷', { x:1.2, y:4.6, w:11, h:0.5, fontSize:16, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:3.3, duration:0.5 }
});
s12.addText('• 6月：Anthropic 指控阿里巴巴"蒸馏攻击"\n• 7月：白宫指控月之暗面蒸馏 Fable\n• 专家反驳：性能无法靠蒸馏获得，KDA 架构完全原创\n\n中国商务部：美方"缺乏事实依据、双重标准，是典型的 AI 霸权主义"', {
  x:1.2, y:5.15, w:11, h:1.4, fontSize:12, color:C.text, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.5,
  animation: { type:'fade', delay:3.6, duration:0.8 }
});

// ========== Slide 13: 章节 - 为什么重要 ==========
let s13 = pptx.addSlide();
addBg(s13);
addSectionNum(s13, '04');
s13.addText('为什么重要', { x:2.2, y:1.0, w:8, h:1.0, fontSize:36, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.3, duration:0.6 }
});
s13.addText('趋势信号 · 开源突围 · 资本定价', { x:2.2, y:2.0, w:8, h:0.5, fontSize:16, color:C.subtext, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.7, duration:0.5 }
});
s13.addShape(pptx.ShapeType.rect, { x:2.2, y:2.6, w:3, h:0.03, fill:{ color:C.accent },
  animation: { type:'wipe', delay:1.0, duration:0.4, dir:'fromLeft' }
});

// ========== Slide 14: 趋势时间轴 ==========
let s14 = pptx.addSlide();
addBg(s14);
s14.addText('中国AI突破密度 — 不是偶然，是趋势', { x:0.8, y:0.4, w:10, h:0.7, fontSize:24, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
// 轴线
s14.addShape(pptx.ShapeType.rect, { x:0.8, y:3.3, w:11.73, h:0.03, fill:{ color:C.border },
  animation: { type:'wipe', delay:0.3, duration:0.8, dir:'fromLeft' }
});
const events = [
  { name:'DeepSeek R1', date:'2025.01', desc:'推理模型\n震惊华尔街', color:C.accent },
  { name:'DeepSeek V3', date:'2025年中', desc:'开源基座\n逼近GPT-4', color:C.accent },
  { name:'Kimi K2', date:'2025年底', desc:'MoE架构\n1万亿参数', color:C.teal },
  { name:'Kimi K3', date:'2026.07', desc:'2.8万亿\n全球第三', color:C.teal, highlight:true },
];
events.forEach((e, i) => {
  const x = 1.2 + i * 3.1;
  const yTop = e.highlight ? 1.2 : 1.5;
  const hCard = e.highlight ? 1.8 : 1.5;
  s14.addShape(pptx.ShapeType.roundRect, { x:x, y:yTop, w:2.8, h:hCard, rectRadius:0.12, fill:{ color:C.card }, line:{ color: e.highlight ? e.color : C.border, width: e.highlight ? 2 : 1 },
    animation: { type:'wipe', delay: 0.6 + i*0.6, duration:0.5, dir:'fromBottom' }
  });
  s14.addText(e.name, { x:x, y:yTop+0.05, w:2.8, h:0.45, fontSize: e.highlight?16:14, bold:true, color:e.color, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.8 + i*0.6, duration:0.4 }
  });
  s14.addText(e.date, { x:x, y:yTop+0.45, w:2.8, h:0.3, fontSize:10, color:C.subtext, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.9 + i*0.6, duration:0.3 }
  });
  s14.addText(e.desc, { x:x, y:yTop+0.8, w:2.8, h:0.7, fontSize:11, color:C.text, align:'center', fontFace:'Microsoft YaHei', lineSpacingMultiple:1.3,
    animation: { type:'fade', delay: 1.0 + i*0.6, duration:0.4 }
  });
  s14.addShape(pptx.ShapeType.ellipse, { x:x+0.8, y:3.15, w:0.3, h:0.3, fill:{ color:e.color },
    animation: { type:'zoom', delay: 1.1 + i*0.6, duration:0.3 }
  });
});
// 结论
s14.addShape(pptx.ShapeType.roundRect, { x:1.5, y:4.0, w:10.33, h:2.5, rectRadius:0.12, fill:{ color:C.card }, line:{ color:C.border, width:1 },
  animation: { type:'wipe', delay:3.2, duration:0.6, dir:'fromBottom' }
});
s14.addText('关键判断', { x:6.66, y:4.1, w:4, h:0.5, fontSize:16, bold:true, color:C.primary, align:'center', fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:3.5, duration:0.4 }
});
s14.addText('DeepSeek 出来时，有人说是灵光一现。K3 出来，你不能再这么说了。\n\n中国AI公司正在系统性逼近、甚至局部超越美国顶级实验室。\n不是一次，是一种趋势。', {
  x:2, y:4.65, w:9.33, h:1.7, fontSize:13, color:C.text, fontFace:'Microsoft YaHei', lineSpacingMultiple:1.5, align:'center',
  animation: { type:'fade', delay:3.8, duration:0.8 }
});

// ========== Slide 15: 生态闭环 ==========
let s15 = pptx.addSlide();
addBg(s15);
s15.addText('中国AI基础设施闭环', { x:0.8, y:0.4, w:10, h:0.7, fontSize:28, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
const nodes = [
  { name:'国产芯片', sub:'华为昇腾 · 沐曦 · 燧原', x:4.67, y:1.5, color:C.red },
  { name:'开源模型', sub:'Kimi K3 · DeepSeek', x:8.5, y:3.5, color:C.teal },
  { name:'部署框架', sub:'vLLM · SGLang', x:4.67, y:5.5, color:C.accent },
  { name:'应用生态', sub:'企业应用 · 开发者', x:0.8, y:3.5, color:C.purple },
];
nodes.forEach((n, i) => {
  s15.addShape(pptx.ShapeType.roundRect, { x:n.x, y:n.y, w:4, h:1.3, rectRadius:0.12, fill:{ color:C.card }, line:{ color:n.color, width:1.5 },
    animation: { type:'wipe', delay: 0.3 + i*0.5, duration:0.5, dir: i===0?'fromTop':i===1?'fromRight':i===2?'fromBottom':'fromLeft' }
  });
  s15.addText(n.name, { x:n.x, y:n.y+0.15, w:4, h:0.5, fontSize:16, bold:true, color:n.color, align:'center', fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 0.6 + i*0.5, duration:0.4 }
  });
  s15.addText(n.sub, { x:n.x, y:n.y+0.65, w:4, h:0.4, fontSize:11, color:C.subtext, align:'center', fontFace:'Microsoft YaHei',
    animation: { type:'fade', delay: 0.7 + i*0.5, duration:0.4 }
  });
});
// 中心文字
s15.addText('自主可控\n完整闭环', { x:5.17, y:3.3, w:3, h:1.2, fontSize:18, bold:true, color:C.primary, align:'center', fontFace:'Microsoft YaHei', lineSpacingMultiple:1.4,
  animation: { type:'zoom', delay:2.5, duration:0.6 }
});
addFooter(s15, '华为昇腾、沐曦、燧原已适配 K3 · vLLM 提供生产级支持');

// ========== Slide 16: 估值 ==========
let s16 = pptx.addSlide();
addBg(s16);
s16.addText('估值飙升 — 真金白银的市场定价', { x:0.8, y:0.4, w:10, h:0.7, fontSize:24, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.0, duration:0.5 }
});
// 轴线
s16.addShape(pptx.ShapeType.rect, { x:1.2, y:3.5, w:10.5, h:0.03, fill:{ color:C.border },
  animation: { type:'wipe', delay:0.3, duration:0.6, dir:'fromLeft' }
});
const vals = [
  { date:'2024.08', val:'$33亿', x:1.2 },
  { date:'2025.12', val:'$43亿', x:3.4 },
  { date:'2026.01', val:'$100亿+', x:5.6 },
  { date:'2026.05', val:'$200亿', x:7.8 },
  { date:'2026.07', val:'$315亿', x:10.0, highlight:true },
];
vals.forEach((v, i) => {
  s16.addShape(pptx.ShapeType.ellipse, { x:v.x+0.5, y:3.35, w:0.3, h:0.3, fill:{ color: v.highlight ? C.accent : C.primary },
    animation: { type:'zoom', delay: 0.6 + i*0.4, duration:0.3 }
  });
  s16.addText(v.val, { x:v.x, y:2.5, w:1.6, h:0.6, fontSize: v.highlight?20:16, bold:true, color: v.highlight?C.accent:C.primary, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.7 + i*0.4, duration:0.4 }
  });
  s16.addText(v.date, { x:v.x, y:3.8, w:1.6, h:0.4, fontSize:10, color:C.subtext, align:'center', fontFace:'Georgia',
    animation: { type:'fade', delay: 0.8 + i*0.4, duration:0.3 }
  });
});
// 目标
s16.addShape(pptx.ShapeType.roundRect, { x:10.0, y:1.5, w:2.5, h:1.5, rectRadius:0.1, fill:{ color:C.card }, line:{ color:C.purple, width:1.5, dashType:'dash' },
  animation: { type:'wipe', delay:2.8, duration:0.5, dir:'fromRight' }
});
s16.addText('$500亿?', { x:10.0, y:1.6, w:2.5, h:0.6, fontSize:20, bold:true, color:C.purple, align:'center', fontFace:'Georgia',
  animation: { type:'fade', delay:3.1, duration:0.4 }
});
s16.addText('Pre-IPO\n中金+高盛 · 港股', { x:10.0, y:2.2, w:2.5, h:0.7, fontSize:10, color:C.subtext, align:'center', fontFace:'Microsoft YaHei', lineSpacingMultiple:1.3,
  animation: { type:'fade', delay:3.3, duration:0.4 }
});
// 结论
s16.addShape(pptx.ShapeType.roundRect, { x:2, y:4.8, w:9.33, h:1.5, rectRadius:0.12, fill:{ color:C.card }, line:{ color:C.accent, width:1.5 },
  animation: { type:'wipe', delay:3.6, duration:0.5, dir:'fromBottom' }
});
s16.addText('从 $33亿 到 $315亿\n冲刺港股 IPO', { x:2, y:4.9, w:9.33, h:1.3, fontSize:18, bold:true, color:C.accent, align:'center', fontFace:'Microsoft YaHei', lineSpacingMultiple:1.5,
  animation: { type:'fade', delay:3.9, duration:0.5 }
});
addFooter(s16, 'Bloomberg · TechCrunch · The Decoder · 电子工程专辑');

// ========== Slide 17: 章节 - 我的判断 ==========
let s17 = pptx.addSlide();
addBg(s17);
addSectionNum(s17, '05');
s17.addText('我的判断', { x:2.2, y:1.0, w:8, h:1.0, fontSize:36, bold:true, color:C.primary, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.3, duration:0.6 }
});
s17.addText('中国团队可以在最前沿的底层机制上做原创', { x:2.2, y:2.0, w:8, h:0.5, fontSize:16, color:C.subtext, fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:0.7, duration:0.5 }
});
s17.addShape(pptx.ShapeType.rect, { x:2.2, y:2.6, w:3, h:0.03, fill:{ color:C.accent },
  animation: { type:'wipe', delay:1.0, duration:0.4, dir:'fromLeft' }
});

// ========== Slide 18: 结尾 ==========
let s18 = pptx.addSlide();
addBg(s18);
s18.addShape(pptx.ShapeType.rect, { x:4.17, y:2.5, w:5, h:0.04, fill:{ color:C.accent },
  animation: { type:'wipe', delay:0.0, duration:0.5, dir:'fromLeft' }
});
s18.addText('关注 AI大家学', { x:1.5, y:2.8, w:10.33, h:1.2, fontSize:40, bold:true, color:C.primary, align:'center', fontFace:'Microsoft YaHei',
  animation: { type:'zoom', delay:0.3, duration:0.8 }
});
s18.addText('AI 世界快人一步', { x:1.5, y:4.0, w:10.33, h:0.6, fontSize:18, color:C.subtext, align:'center', fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:1.0, duration:0.6 }
});
s18.addText('开源 vs 闭源，你站哪边？评论区聊聊 👇', { x:1.5, y:5.2, w:10.33, h:0.5, fontSize:13, color:C.accent, align:'center', fontFace:'Microsoft YaHei',
  animation: { type:'fade', delay:1.5, duration:0.5 }
});

// ========== 导出 ==========
const outPath = path.join(__dirname, 'KimiK3-演示文稿-v2.pptx');
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log('PPT 已生成: ' + outPath);
}).catch(err => {
  console.error('生成失败:', err);
});
