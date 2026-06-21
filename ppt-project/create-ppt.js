const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const imgDir = path.join(__dirname, "images");
if (!fs.existsSync(imgDir)) fs.mkdirSync(imgDir);

// Color scheme
const C = {
  bg: `FAF7F2`, darkBg: `1C1C1E`, warmBrown: `8B6914`,
  deepBrown: `5C3D0E`, earthGreen: `5B8C5A`, systemBlue: `3A7BD5`,
  alertRed: `D94A4A`, alertOrange: `E8913A`, successGreen: `4CAF50`,
  textDark: `2C2C2C`, textGray: `6B6B6B`, cardBg: `F5F0E8`,
  white: `FFFFFF`, black: `000000`, lightLine: `E8E0D5`, lightOverlay: `FFF8F0`,
};

// Font presets
const FONT = `Microsoft YaHei`;

// Shortcuts
function bg(slide, color) { slide.background = { fill: color }; }

function placeholder(slide, x, y, w, h, label, note) {
  slide.addShape(`rect`, { x, y, w, h, fill: { color: `F0EBE0` }, line: { color: C.warmBrown, width: 2, dashType: `dash` }, rectRadius: 0.05 });
  slide.addText(label, { x, y, w, h, align: `center`, valign: `middle`, fontSize: 10, fontFace: FONT, color: C.textGray });
  if (note) slide.addNotes(note);
}

function pn(slide, num) {
  slide.addText(`${num} / 30`, { x: 8.8, y: 5.2, w: 1, h: 0.4, fontSize: 9, fontFace: FONT, color: C.textGray, align: `right` });
}

function rect(slide, x, y, w, h, opts) { return slide.addShape(`rect`, { x, y, w, h, ...opts }); }

const pptx = new PptxGenJS();
pptx.layout = `LAYOUT_WIDE`;
pptx.author = `张家硕`;
pptx.title = `如果她（他）活到今天——AI驱动下新时代乡村共创课堂`;

// ===== P1 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  placeholder(s, 0, 0, 10, 5.63, `【AI生图-豆包】西南山区清晨薄雾村庄，画面底部渐暗留标题区`, `豆包提示词：中国西南山区乡村清晨，薄雾笼罩村庄，远处层叠山峦，近处青瓦白墙农舍，纪录片电影质感，温暖自然光，真实摄影风格，横向构图，画面底部渐暗，不要人物，不要文字`);
  rect(s, 0, 3.2, 10, 2.43, { fill: { color: `000000`, transparency: 40 } });
  s.addText(`如果她（他）活到今天`, { x: 0.8, y: 3.5, w: 8.4, h: 1.0, fontSize: 48, fontFace: FONT, bold: true, color: C.white, align: `left`, valign: `bottom` });
  s.addText(`AI驱动下新时代乡村共创课堂`, { x: 0.8, y: 4.5, w: 8.4, h: 0.5, fontSize: 20, fontFace: FONT, color: `D4C8B0`, align: `left` });
  s.addText(`文化素养+AI  |  思政×地理  |  高中段`, { x: 0.8, y: 5.0, w: 8.4, h: 0.35, fontSize: 12, fontFace: FONT, color: `A09080` });
}

// ===== P2 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  rect(s, 0, 0, 10, 0.5, { fill: { color: `0D0D0F` } });
  s.addText(`AI乡村共创系统  v2.4  |  任务状态：待启动  |  实践团编号：YLC-2024-003`, { x: 0.3, y: 0.05, w: 9.4, h: 0.4, fontSize: 11, fontFace: `Courier New`, color: `4A9F4A` });
  rect(s, 1.2, 1.2, 7.6, 3.2, { fill: { color: `1A1A2E` }, rectRadius: 0.1 });
  rect(s, 1.2, 1.2, 7.6, 0.06, { fill: { color: C.systemBlue } });
  s.addText(`【系统提示】`, { x: 1.6, y: 1.5, w: 6.8, h: 0.5, fontSize: 18, fontFace: FONT, bold: true, color: C.systemBlue });
  s.addText(`欢迎进入「新时代乡村共创行动」\n\n你们将接手一座正在推进乡村振兴的村庄。\n请完成以下任务：`, { x: 1.6, y: 2.0, w: 6.8, h: 1.2, fontSize: 16, fontFace: FONT, color: `C8C8D0` });
  const tasks = [{ i: `\u{1F4CB}`, t: `任务一`, d: `群众调研与沟通` },{ i: `\u{1F50D}`, t: `任务二`, d: `村庄问题诊断分析` },{ i: `\u{1F3D8}`, t: `任务三`, d: `乡村共创方案设计` }];
  tasks.forEach((tk, i) => {
    const bx = 1.6 + i * 2.5;
    rect(s, bx, 3.4, 2.2, 0.7, { fill: { color: `252540` }, rectRadius: 0.05 });
    s.addText(`${tk.i}  ${tk.t}`, { x: bx + 0.15, y: 3.42, w: 1.9, h: 0.32, fontSize: 13, fontFace: FONT, bold: true, color: C.systemBlue });
    s.addText(tk.d, { x: bx + 0.15, y: 3.74, w: 1.9, h: 0.3, fontSize: 10, fontFace: FONT, color: `9090A0` });
  });
  rect(s, 1.6, 4.4, 6.8, 0.08, { fill: { color: `333355` }, rectRadius: 0.04 });
  rect(s, 1.6, 4.4, 0.5, 0.08, { fill: { color: C.systemBlue }, rectRadius: 0.04 });
  s.addText(`系统初始化中... 5%`, { x: 1.6, y: 4.5, w: 3, h: 0.35, fontSize: 10, fontFace: `Courier New`, color: `7070A0` });
  s.addText(`教师导入语：今天这节课，你们不只是课堂中的学生。\n你们将作为「新时代青年乡村共创实践团成员」，进入一个真实乡村建设任务。`, { x: 0.5, y: 4.9, w: 9, h: 0.55, fontSize: 13, fontFace: FONT, color: `A0A0B0`, align: `center` });
  pn(s, 2);
}

// ===== P3 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  placeholder(s, 0.8, 0.6, 8.4, 4.0, `【AI视频-可灵AI】20秒短视频\n西南山区航拍→盘山公路→村庄→村民劳作\n（生成后替换为实际视频嵌入）`, `可灵AI分段提示词见文档P3说明。4段各5秒，必剪拼接。访问 kling.kuaishou.com`);
  s.addText(`云岭村 · 西南喀斯特山区`, { x: 0.8, y: 4.7, w: 8.4, h: 0.4, fontSize: 16, fontFace: FONT, color: `B0B0C0`, align: `center` });
  pn(s, 3);
}

// ===== P4 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`云岭村基础资料`, { x: 0.5, y: 0.2, w: 5, h: 0.6, fontSize: 28, fontFace: FONT, bold: true, color: C.textDark });
  rect(s, 0.5, 0.75, 2.5, 0.04, { fill: { color: C.warmBrown } });
  placeholder(s, 0.5, 1.1, 4.3, 3.3, `【AI生图-豆包】云岭村俯视图/位置示意图\n山区聚落沿山坡分散分布`, `豆包提示词：中国西南山区村庄俯视图，村落沿山坡分散分布，喀斯特地貌特征，可见盘山小路连接各户，自然纪录片航拍风格`);
  const data = [
    { l: `户数`, v: `286户`, sub: `分散于6个山坡片区` },
    { l: `主要作物`, v: `柑橘`, sub: `种植面积约800亩` },
    { l: `交通`, v: `距县城约47公里`, sub: `仅一条盘山公路连通` },
    { l: `雨季`, v: `5—8月`, sub: `年均塌方阻断3—5次` },
    { l: `青壮年外流`, v: `约62%`, sub: `常住人口以老人妇女为主` },
  ];
  data.forEach((item, i) => {
    const by = 1.1 + i * 0.82;
    rect(s, 5.2, by, 4.3, 0.7, { fill: { color: C.cardBg }, rectRadius: 0.05 });
    s.addText(item.l, { x: 5.4, y: by + 0.05, w: 1.8, h: 0.28, fontSize: 11, fontFace: FONT, color: C.textGray });
    s.addText(item.v, { x: 5.4, y: by + 0.28, w: 1.8, h: 0.32, fontSize: 18, fontFace: FONT, bold: true, color: C.textDark });
    s.addText(item.sub, { x: 7.3, y: by + 0.15, w: 2.0, h: 0.45, fontSize: 10, fontFace: FONT, color: C.textGray, valign: `middle` });
  });
  rect(s, 0.5, 4.6, 9, 0.6, { fill: { color: C.lightOverlay }, rectRadius: 0.05 });
  s.addText(`\u{1F914} 思考：为什么一个拥有自然资源和产业基础的村庄，发展仍然困难重重？`, { x: 0.7, y: 4.62, w: 8.6, h: 0.55, fontSize: 16, fontFace: FONT, color: C.deepBrown, valign: `middle` });
  pn(s, 4);
}

// ===== P5 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  rect(s, 0, 0, 10, 0.08, { fill: { color: C.alertRed } });
  rect(s, 1.5, 0.5, 7, 0.7, { fill: { color: `2A1515` }, rectRadius: 0.05 });
  s.addText(`⚠  【当前任务状态】乡村共创行动推进失败`, { x: 1.7, y: 0.55, w: 6.6, h: 0.6, fontSize: 20, fontFace: FONT, bold: true, color: C.alertRed, valign: `middle` });
  const reasons = [`群众不愿配合调研`, `村民对实践团缺乏信任`, `产业方案推进停滞`];
  reasons.forEach((r, i) => {
    rect(s, 1.8, 1.5 + i * 0.55, 6.4, 0.45, { fill: { color: `1E1414` }, rectRadius: 0.03 });
    s.addText(`✘  ${r}`, { x: 2.0, y: 1.5 + i * 0.55, w: 6.0, h: 0.45, fontSize: 15, fontFace: FONT, color: `E08080`, valign: `middle` });
  });
  s.addText(`—— 村民反馈信息流 ——`, { x: 1.5, y: 3.2, w: 7, h: 0.4, fontSize: 12, fontFace: FONT, color: `808090`, align: `center` });
  const msgs = [
    { t: `以前也有人来调研，最后没下文。`, a: `张大爷`, c: `4A3A3A` },
    { t: `你们年轻人懂农村吗？`, a: `李婶`, c: `3A3A4A` },
    { t: `产业项目以前赔过，干部说得好听。`, a: `王叔`, c: `3A4A3A` },
    { t: `最后还不是我们自己承担风险？`, a: `陈伯`, c: `4A4A3A` },
  ];
  msgs.forEach((msg, i) => {
    const mx = i % 2 === 0 ? 0.8 : 5.2;
    const my = 3.7 + Math.floor(i / 2) * 0.65;
    s.addShape(`ellipse`, { x: mx, y: my, w: 0.4, h: 0.4, fill: { color: msg.c } });
    s.addText(msg.a[0], { x: mx, y: my, w: 0.4, h: 0.4, fontSize: 10, fontFace: FONT, color: C.white, align: `center`, valign: `middle` });
    rect(s, mx + 0.5, my, 3.5, 0.42, { fill: { color: msg.c }, rectRadius: 0.05 });
    s.addText(`${msg.t}`, { x: mx + 0.6, y: my, w: 3.3, h: 0.42, fontSize: 11, fontFace: FONT, color: `D0C8C0`, valign: `middle` });
    s.addText(msg.a, { x: mx, y: my + 0.42, w: 0.4, h: 0.2, fontSize: 8, fontFace: FONT, color: `808090`, align: `center` });
  });
  s.addText(`教师提问：「为什么乡村建设方案还没开始，任务就已经停滞了？」`, { x: 0.5, y: 5.1, w: 9, h: 0.4, fontSize: 13, fontFace: FONT, color: `A0A0B0`, align: `center` });
  pn(s, 5);
}

// ===== P6 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  placeholder(s, 0, 0, 10, 5.63, `【AI生图-豆包】夜晚村庄模糊背景\n雨后泥路反射灯光，氛围感暗色调`, `豆包提示词：中国西南山村夜晚，雨后泥泞路面反射远处窗户灯光，氛围感夜景，电影质感，画面略微模糊作为背景，色调偏暗暖`);
  rect(s, 0, 0, 10, 5.63, { fill: { color: `000000`, transparency: 45 } });
  rect(s, 1.5, 1.2, 7, 3.2, { fill: { color: `1A1A2E` }, rectRadius: 0.08 });
  rect(s, 1.5, 1.2, 7, 0.05, { fill: { color: C.systemBlue } });
  s.addText(`【系统检测】`, { x: 2.0, y: 1.5, w: 6, h: 0.5, fontSize: 18, fontFace: FONT, bold: true, color: C.systemBlue });
  s.addText(`群众工作推进受阻。\n检测到可解锁资源：基层工作档案。\n\n是否解锁？`, { x: 2.0, y: 2.1, w: 6, h: 1.2, fontSize: 18, fontFace: FONT, color: `C8C8D0`, align: `center` });
  rect(s, 2.8, 3.5, 1.8, 0.7, { fill: { color: C.systemBlue }, rectRadius: 0.08 });
  s.addText(`是`, { x: 2.8, y: 3.5, w: 1.8, h: 0.7, fontSize: 26, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `middle` });
  rect(s, 5.4, 3.5, 1.8, 0.7, { fill: { color: `3A3A3A` }, rectRadius: 0.08 });
  s.addText(`否`, { x: 5.4, y: 3.5, w: 1.8, h: 0.7, fontSize: 26, fontFace: FONT, bold: true, color: `808080`, align: `center`, valign: `middle` });
  s.addNotes(`课堂操作：教师请学生齐声选择。点击「是」进入P7。这页是课堂节奏的转折点——由学生来决定推进故事。`);
  pn(s, 6);
}

// ===== P7 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  placeholder(s, 0, 0, 10, 5.63, `【AI生图-即梦AI】夜晚暴雨后泥路\n年轻女性驻村干部背影，车停在村口\n（本页是全课视觉锚点，用即梦生成更高画质）`, `即梦AI提示词：夜晚的中国西南山村，村口停着一辆沾满泥浆的白色轿车，雨后泥泞的土路反射着车灯光，远处一两户人家窗户透出暖黄色灯光，年轻女性驻村干部的背影（短发、深色外套、背着工作包），她正从车里出来准备入户，纪录片真实摄影风格，电影级构图，画面传递出坚定又疲惫的情绪，不要正脸特写`);
  rect(s, 0, 4.0, 10, 1.63, { fill: { color: `000000`, transparency: 35 } });
  s.addText(`2018年  ·  驻村第67天  ·  第一次大规模入户调研`, { x: 1.0, y: 4.2, w: 8, h: 0.5, fontSize: 18, fontFace: FONT, color: `B0A090`, align: `center` });
  s.addText(`你将作为实践团成员，跟随黄文秀开展第一次入户工作。`, { x: 1.0, y: 4.7, w: 8, h: 0.5, fontSize: 16, fontFace: FONT, bold: true, color: C.white, align: `center` });
  pn(s, 7);
}

// ===== P8 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  rect(s, 0, 0, 10, 0.6, { fill: { color: C.deepBrown } });
  s.addText(`【情境1】第一次入户 —— 村民拒绝开门`, { x: 0.5, y: 0.05, w: 9, h: 0.5, fontSize: 17, fontFace: FONT, bold: true, color: C.white, valign: `middle` });
  placeholder(s, 0.5, 0.9, 4.3, 3.2, `【AI生图-豆包】紧闭的农村老旧木门\n门缝透出暖黄灯光`, `豆包提示词：中国西南农村老旧木门特写，门紧闭，门缝透出暖黄色灯光，门上贴着褪色的对联，傍晚氛围，真实农村摄影`);
  rect(s, 5.2, 0.9, 4.3, 1.3, { fill: { color: C.cardBg }, rectRadius: 0.08 });
  s.addShape(`triangle`, { x: 5.0, y: 1.3, w: 0.25, h: 0.25, fill: { color: C.cardBg }, rotate: 90 });
  s.addText(`\u{1F5E3} 村民反馈：`, { x: 5.4, y: 0.95, w: 3.9, h: 0.3, fontSize: 12, fontFace: FONT, bold: true, color: C.textGray });
  s.addText(`「不用调研了，以前也有人来，\n最后什么都没解决。」`, { x: 5.4, y: 1.25, w: 3.9, h: 0.85, fontSize: 16, fontFace: FONT, color: C.textDark, valign: `middle` });
  rect(s, 5.2, 2.5, 4.3, 0.5, { fill: { color: `E8EDF5` }, rectRadius: 0.05 });
  s.addText(`【任务选择】如果你是实践团成员，你会：`, { x: 5.4, y: 2.52, w: 3.9, h: 0.45, fontSize: 12, fontFace: FONT, bold: true, color: C.systemBlue, valign: `middle` });
  const opts = [
    { l: `A`, t: `强调这是政策要求，希望村民配合`, c: C.alertRed, f: `村民情绪更加抵触。` },
    { l: `B`, t: `留下调查表，改天再来`, c: C.alertOrange, f: `信息收集效率较低，群众关系没有改善。` },
    { l: `C`, t: `主动帮助村民做事，再寻找交流机会`, c: C.successGreen, f: `黄文秀选择：先帮村民扫院子，再慢慢聊天。` },
  ];
  opts.forEach((o, i) => {
    const oy = 3.2 + i * 0.75;
    rect(s, 5.2, oy, 4.3, 0.6, { fill: { color: C.white }, line: { color: o.c, width: 2 }, rectRadius: 0.05 });
    rect(s, 5.3, oy + 0.08, 0.44, 0.44, { fill: { color: o.c }, rectRadius: 0.03 });
    s.addText(o.l, { x: 5.3, y: oy + 0.08, w: 0.44, h: 0.44, fontSize: 18, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `middle` });
    s.addText(o.t, { x: 5.85, y: oy, w: 3.5, h: 0.6, fontSize: 12, fontFace: FONT, color: C.textDark, valign: `middle` });
  });
  s.addNotes(`课堂操作：学生小组讨论后举手选择。教师点击对应选项展示反馈。选C进入P9。`);
  pn(s, 8);
}

// ===== P9 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  rect(s, 0, 0, 10, 0.6, { fill: { color: C.earthGreen } });
  s.addText(`✓ 正确选择 —— 黄文秀的工作方式`, { x: 0.5, y: 0.05, w: 9, h: 0.5, fontSize: 17, fontFace: FONT, bold: true, color: C.white, valign: `middle` });
  const scenes = [
    { l: `帮村民扫院子`, p: `中国西南农村小院，年轻女性基层干部拿着扫帚在打扫院子，旁边有老人在看，清晨光线，温暖自然光，真实纪录风格` },
    { l: `帮村民收玉米`, p: `中国农村玉米地，年轻女性在帮老人收玉米，穿着朴素深色外套，阳光透过玉米叶，劳动场景真实感` },
    { l: `一边干活一边聊家常`, p: `傍晚农家门口，年轻女干部和老农妇坐在小板凳上聊天，自然放松的氛围，暖色调晚霞光，纪录片偷拍视角` },
  ];
  scenes.forEach((sc, i) => {
    const sx = 0.5 + i * 3.1;
    placeholder(s, sx, 0.85, 2.9, 2.2, `【豆包/即梦】\n${sc.l}`, `提示词：${sc.p}`);
    s.addText(sc.l, { x: sx, y: 3.1, w: 2.9, h: 0.35, fontSize: 12, fontFace: FONT, color: C.textGray, align: `center` });
  });
  rect(s, 1.0, 3.7, 8, 0.7, { fill: { color: C.lightOverlay }, rectRadius: 0.05 });
  s.addText(`「黄文秀选择：先帮村民干活，再慢慢聊天。」`, { x: 1.2, y: 3.72, w: 7.6, h: 0.65, fontSize: 20, fontFace: FONT, bold: true, color: C.deepBrown, align: `center`, valign: `middle` });
  s.addText(`——源自黄文秀在百坭村的真实工作记录`, { x: 1.0, y: 4.5, w: 8, h: 0.3, fontSize: 11, fontFace: FONT, color: C.textGray, align: `center` });
  s.addText(`教师追问（口头）：黄文秀明明是来推进工作的，为什么却先去帮村民干活？`, { x: 0.5, y: 5.05, w: 9, h: 0.35, fontSize: 12, fontFace: FONT, color: C.systemBlue, align: `center` });
  pn(s, 9);
}

// ===== P10 =====
{
  const s = pptx.addSlide(); bg(s, `1C1410`);
  placeholder(s, 0, 0, 10, 5.63, `【复用P9图】半透明遮罩叠加\n黄文秀帮村民劳动画面`, `将P9生成的图片置入后设透明度60%`);
  rect(s, 0, 0, 10, 5.63, { fill: { color: `000000`, transparency: 50 } });
  const qs = [
    { n: `第一问`, t: `黄文秀明明是来推进工作的，\n为什么却先去帮村民干活？` },
    { n: `第二问`, t: `如果连群众真实想法都不了解，\n工作还能推进吗？` },
    { n: `第三问`, t: `她是在「管理群众」，\n还是在「走进群众」？` },
  ];
  qs.forEach((q, i) => {
    const qy = 0.5 + i * 1.6;
    rect(s, 1.0, qy, 1.3, 0.45, { fill: { color: C.warmBrown }, rectRadius: 0.05 });
    s.addText(q.n, { x: 1.0, y: qy, w: 1.3, h: 0.45, fontSize: 16, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `middle` });
    s.addText(q.t, { x: 2.5, y: qy, w: 6.5, h: 0.8, fontSize: 22, fontFace: FONT, bold: true, color: C.white });
    if (i < 2) rect(s, 2.5, qy + 1.2, 6, 0.01, { fill: { color: `5A4A3A` } });
  });
  s.addNotes(`课堂操作：教师逐一点击展示问题。前两问铺垫，第三问逼近「群众路线」核心。`);
  pn(s, 10);
}

// ===== P11 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  s.addText(`【黄文秀基层工作档案 · AI对话系统】`, { x: 0.5, y: 0.15, w: 9, h: 0.5, fontSize: 18, fontFace: FONT, bold: true, color: C.systemBlue });
  rect(s, 0.3, 0.8, 3.2, 4.3, { fill: { color: `1A1A2E` }, rectRadius: 0.05 });
  placeholder(s, 0.7, 1.0, 2.4, 2.2, `【AI生图】黄文秀工作照风格\n背影/侧面即可，暖色调`, `用即梦AI生成，不要正脸`);
  s.addText(`AI智能体：黄文秀基层工作档案`, { x: 0.5, y: 3.4, w: 2.8, h: 0.4, fontSize: 11, fontFace: FONT, bold: true, color: C.systemBlue });
  s.addText(`基于黄文秀在百坭村的真实工作理念构建。\n可自由提问关于基层工作、群众沟通的任何问题。`, { x: 0.5, y: 3.75, w: 2.8, h: 0.6, fontSize: 9, fontFace: FONT, color: `9090A0` });
  rect(s, 0.5, 4.5, 2.8, 0.5, { fill: { color: `1D2E1D` }, rectRadius: 0.03 });
  s.addText(`创建方式：扣子Coze (coze.cn)\n嵌入方式：PPT Web Viewer插件\n或分屏浏览器展示`, { x: 0.6, y: 4.5, w: 2.6, h: 0.5, fontSize: 8, fontFace: FONT, color: `80B080`, valign: `middle` });
  rect(s, 3.8, 0.8, 5.9, 4.3, { fill: { color: `111122` }, rectRadius: 0.05 });
  const chats = [
    { sd: `学生`, t: `为什么群众不愿意相信干部？`, c: `2A4A2A`, a: `right` },
    { sd: `AI`, t: `如果连群众真正担心什么都不知道，再好的方案也推进不了。`, c: `2A2A4A`, a: `left` },
    { sd: `学生`, t: `为什么你总要反复入户？`, c: `2A4A2A`, a: `right` },
    { sd: `AI`, t: `群众不愿意配合，不是因为他们不讲道理，而是因为他们有真实顾虑。只有真正走进群众，群众才愿意相信你。`, c: `2A2A4A`, a: `left` },
  ];
  chats.forEach((msg, i) => {
    const cy = 1.1 + i * 1.0;
    const cx = msg.a === `right` ? 5.8 : 4.0;
    const cw = msg.a === `right` ? 3.5 : 3.8;
    rect(s, cx, cy, cw, 0.75, { fill: { color: msg.c }, rectRadius: 0.05 });
    s.addText(msg.sd, { x: cx + 0.15, y: cy + 0.03, w: cw - 0.3, h: 0.2, fontSize: 8, fontFace: FONT, bold: true, color: `A0B0A0` });
    s.addText(msg.t, { x: cx + 0.15, y: cy + 0.22, w: cw - 0.3, h: 0.5, fontSize: 10, fontFace: FONT, color: `D0D0D0`, valign: `top` });
  });
  rect(s, 4.0, 4.55, 4.0, 0.38, { fill: { color: `1A1A30` }, rectRadius: 0.03 });
  s.addText(`输入你的问题...`, { x: 4.1, y: 4.55, w: 3.0, h: 0.38, fontSize: 10, fontFace: FONT, color: `606080`, valign: `middle` });
  rect(s, 8.1, 4.55, 0.8, 0.38, { fill: { color: C.systemBlue }, rectRadius: 0.03 });
  s.addText(`发送`, { x: 8.1, y: 4.55, w: 0.8, h: 0.38, fontSize: 10, fontFace: FONT, color: C.white, align: `center`, valign: `middle` });
  s.addNotes(`创新亮点1：AI智能体嵌入课件。课前用扣子Coze创建好「黄文秀基层工作档案」Bot。人设提示词见文档P11。课堂上教师打开嵌入网页，学生口头发问或上台输入，AI实时回答。`);
  pn(s, 11);
}

// ===== P12 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  rect(s, 0, 0, 10, 0.6, { fill: { color: C.alertRed } });
  s.addText(`⚠  【情境2】产业推进警报 —— 大量村民反对`, { x: 0.5, y: 0.05, w: 9, h: 0.5, fontSize: 17, fontFace: FONT, bold: true, color: C.white, valign: `middle` });
  placeholder(s, 0.5, 0.9, 4.5, 2.6, `【AI生图-豆包】滞销柑橘堆积\n果园里烂掉的果子，阴天`, `豆包提示词：中国农村果园，成堆柑橘堆在地上，部分果实开始腐烂，阴天，画面传达出焦急和无奈的情绪，新闻纪实摄影风格，真实感`);
  rect(s, 5.3, 0.9, 4.2, 2.6, { fill: { color: C.cardBg }, rectRadius: 0.05 });
  s.addText(`贡柑滞销事件档案`, { x: 5.5, y: 1.0, w: 3.8, h: 0.4, fontSize: 15, fontFace: FONT, bold: true, color: C.alertRed });
  [`\u{1F342} 果子烂在地里，无人收购`, `\u{1F4B0} 被收购商持续压价`, `\u{1F4B8} 农户血本无归`, `\u{1F3C3} 收购商跑路，定金不退`].forEach((e, i) => {
    s.addText(e, { x: 5.5, y: 1.5 + i * 0.45, w: 3.8, h: 0.4, fontSize: 13, fontFace: FONT, color: C.textDark, valign: `middle` });
  });
  rect(s, 0.8, 3.8, 8.4, 0.7, { fill: { color: `FFF0E8` }, rectRadius: 0.08 });
  s.addText(`\u{1F5E3} 村民真实语音：「你们说发展产业，万一又赔了怎么办？」`, { x: 1.0, y: 3.8, w: 8.0, h: 0.7, fontSize: 18, fontFace: FONT, bold: true, color: C.deepBrown, align: `center`, valign: `middle` });
  const opts2 = [
    { l: `A`, t: `强调国家乡村振兴政策` },
    { l: `B`, t: `要求群众必须服从安排` },
    { l: `C`, t: `先了解群众顾虑，再共同寻找解决办法` },
  ];
  opts2.forEach((o, i) => {
    const ox = 0.8 + i * 3.0;
    const isC = i === 2;
    rect(s, ox, 4.7, 2.8, 0.65, { fill: { color: C.white }, line: { color: isC ? C.successGreen : C.textGray, width: 2 }, rectRadius: 0.05 });
    rect(s, ox + 0.08, 4.78, 0.4, 0.4, { fill: { color: isC ? C.successGreen : C.textGray }, rectRadius: 0.03 });
    s.addText(o.l, { x: ox + 0.08, y: 4.78, w: 0.4, h: 0.4, fontSize: 16, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `middle` });
    s.addText(o.t, { x: ox + 0.55, y: 4.7, w: 2.15, h: 0.65, fontSize: 10, fontFace: FONT, color: C.textDark, valign: `middle` });
  });
  pn(s, 12);
}

// ===== P13 =====
{
  const s = pptx.addSlide(); bg(s, `1C1410`);
  placeholder(s, 0, 0, 10, 5.63, `【复用P9劳动场景图】半透明叠加`, ``);
  rect(s, 0, 0, 10, 5.63, { fill: { color: `000000`, transparency: 55 } });
  const qs2 = [
    { n: `第一问`, t: `为什么黄文秀总是强调——\n先听群众怎么想？` },
    { n: `第二问`, t: `她为什么没有替群众\n直接做决定？` },
    { n: `第三问`, t: `她的这些做法，\n体现了一种怎样的工作思想？` },
  ];
  qs2.forEach((q, i) => {
    const qy = 0.6 + i * 1.6;
    rect(s, 1.0, qy, 1.3, 0.45, { fill: { color: C.warmBrown }, rectRadius: 0.05 });
    s.addText(q.n, { x: 1.0, y: qy, w: 1.3, h: 0.45, fontSize: 16, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `middle` });
    s.addText(q.t, { x: 2.5, y: qy, w: 6.5, h: 0.8, fontSize: 22, fontFace: FONT, bold: true, color: C.white });
    if (i < 2) rect(s, 2.5, qy + 1.2, 6, 0.01, { fill: { color: `5A4A3A` } });
  });
  s.addText(`关键词提示：相信人民群众...  一切为了群众...  从群众中来...  到群众中去...`, { x: 2.5, y: 5.0, w: 6.5, h: 0.35, fontSize: 12, fontFace: FONT, color: `A09080` });
  pn(s, 13);
}

// ===== P14 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`群众观点`, { x: 0.5, y: 0.2, w: 4, h: 0.6, fontSize: 32, fontFace: FONT, bold: true, color: C.deepBrown });
  rect(s, 0.5, 0.75, 3, 0.04, { fill: { color: C.warmBrown } });
  const pts = [
    { t: `相信人民群众自己解放自己`, d: `黄文秀没有替群众包办一切，\n而是发动群众共同参与建设` },
    { t: `全心全意为人民服务`, d: `先帮村民干活，再了解需求，\n把群众放在心中最高位置` },
    { t: `一切向人民群众负责`, d: `反复入户，确保了解每一户\n真实情况，不走过场` },
    { t: `虚心向人民群众学习`, d: `先听群众怎么想，\n不替群众做决定` },
  ];
  pts.forEach((p, i) => {
    const py = 1.1 + i * 1.05;
    rect(s, 0.5, py, 5.2, 0.9, { fill: { color: C.white }, line: { color: C.lightLine, width: 1 }, rectRadius: 0.05 });
    rect(s, 0.5, py, 0.06, 0.9, { fill: { color: C.warmBrown } });
    s.addText(p.t, { x: 0.75, y: py + 0.05, w: 4.8, h: 0.35, fontSize: 15, fontFace: FONT, bold: true, color: C.deepBrown });
    s.addText(p.d, { x: 0.75, y: py + 0.38, w: 4.8, h: 0.45, fontSize: 11, fontFace: FONT, color: C.textGray });
  });
  placeholder(s, 6.0, 1.1, 3.5, 3.8, `【复用P9三格图中的一幅】\n黄文秀帮村民劳动`, ``);
  rect(s, 0.5, 4.95, 9, 0.45, { fill: { color: C.lightOverlay }, rectRadius: 0.03 });
  s.addText(`「黄文秀没有把群众当成「被管理对象」，而是始终尊重群众、理解群众、向群众学习。」`, { x: 0.7, y: 4.95, w: 8.6, h: 0.45, fontSize: 13, fontFace: FONT, italic: true, color: C.warmBrown, align: `center`, valign: `middle` });
  pn(s, 14);
}

// ===== P15 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`群众路线`, { x: 0.5, y: 0.2, w: 4, h: 0.6, fontSize: 32, fontFace: FONT, bold: true, color: C.deepBrown });
  rect(s, 0.5, 0.75, 3, 0.04, { fill: { color: C.warmBrown } });
  const rts = [
    { t: `一切为了群众`, d: `她推动的每一项工作，\n都是为了让群众过上好日子` },
    { t: `一切依靠群众`, d: `她没有替群众包办，\n而是发动群众共同参与建设` },
    { t: `从群众中来`, d: `反复入户调研，了解群众真实需求，\n不是坐在办公室制定方案` },
    { t: `到群众中去`, d: `把方案带回村里，和群众一起商量，\n一起实施、一起改进` },
  ];
  rts.forEach((r, i) => {
    const ry = 1.1 + i * 0.95;
    rect(s, 0.5, ry, 5.2, 0.8, { fill: { color: C.white }, line: { color: C.lightLine, width: 1 }, rectRadius: 0.05 });
    rect(s, 0.5, ry, 0.06, 0.8, { fill: { color: C.earthGreen } });
    s.addText(r.t, { x: 0.75, y: ry + 0.05, w: 4.8, h: 0.3, fontSize: 15, fontFace: FONT, bold: true, color: C.earthGreen });
    s.addText(r.d, { x: 0.75, y: ry + 0.33, w: 4.8, h: 0.42, fontSize: 11, fontFace: FONT, color: C.textGray });
  });
  placeholder(s, 6.0, 1.1, 3.5, 3.8, `【AI生图-豆包】\n驻村干部与村民围坐讨论\n温暖，参与感`, `豆包提示词：中国农村傍晚，年轻驻村干部和几位村民围坐在院子里讨论，气氛融洽，温暖晚霞光线，真实纪录片风格`);
  rect(s, 0.5, 4.95, 9, 0.45, { fill: { color: C.lightOverlay }, rectRadius: 0.03 });
  s.addText(`「她没有坐在办公室制定方案，而是反复入户、了解群众真实需求、发动群众共同参与建设。」`, { x: 0.7, y: 4.95, w: 8.6, h: 0.45, fontSize: 13, fontFace: FONT, italic: true, color: C.earthGreen, align: `center`, valign: `middle` });
  pn(s, 15);
}

// ===== P16 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`课后任务`, { x: 0.5, y: 0.2, w: 3, h: 0.5, fontSize: 14, fontFace: FONT, bold: true, color: C.systemBlue });
  s.addText(`《云岭村群众工作微方案》`, { x: 0.5, y: 0.6, w: 9, h: 0.7, fontSize: 28, fontFace: FONT, bold: true, color: C.textDark });
  const tts = [`当前群众问题分析`, `群众不信任的原因探究`, `黄文秀哪些做法体现了群众观点？`, `黄文秀哪些做法体现了群众路线？`, `作为实践团成员，你会怎样真正发动群众参与建设？`, `人民群众为什么能够推动乡村建设？`];
  tts.forEach((t, i) => {
    const ty = 1.5 + i * 0.55;
    rect(s, 0.5, ty, 6.5, 0.45, { fill: { color: C.white }, line: { color: C.lightLine, width: 1 }, rectRadius: 0.03 });
    s.addText(`${i + 1}.`, { x: 0.6, y: ty, w: 0.4, h: 0.45, fontSize: 14, fontFace: FONT, bold: true, color: C.warmBrown, valign: `middle` });
    s.addText(t, { x: 1.0, y: ty, w: 5.8, h: 0.45, fontSize: 13, fontFace: FONT, color: C.textDark, valign: `middle` });
  });
  rect(s, 7.3, 1.5, 2.3, 3.3, { fill: { color: C.lightOverlay }, rectRadius: 0.08 });
  s.addText(`\u{1F4F1} 提交方式`, { x: 7.5, y: 1.6, w: 1.9, h: 0.35, fontSize: 14, fontFace: FONT, bold: true, color: C.deepBrown });
  s.addText(`扫码进入\n「乡村共创\n任务平台」\n在线提交方案`, { x: 7.5, y: 2.1, w: 1.9, h: 1.0, fontSize: 12, fontFace: FONT, color: C.textGray, align: `center` });
  placeholder(s, 7.7, 3.2, 1.5, 1.4, `【简道云二维码】\n生成后替换`, `在简道云(jiandaoyun.com)创建表单后生成二维码`);
  s.addNotes(`创新亮点2：任务提交平台。用简道云搭建「乡村共创任务平台」，学生在线提交方案，教师后台实时查看数据。详见文档P16。`);
  pn(s, 16);
}

// ===== P17 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  placeholder(s, 0, 0, 10, 5.63, `【AI视频-可灵AI】25秒混剪\n黄文秀入户→调研→帮村民劳动→笑脸\n（图生视频模式，将P7/P9/P12图动起来）`, `用可灵AI图生视频功能，将前面生成的静态图转为5秒动态片段，必剪拼接。详见文档P17。`);
  rect(s, 0, 0, 10, 5.63, { fill: { color: `000000`, transparency: 40 } });
  s.addText(`只有扎根泥土，才能懂得人民。`, { x: 0.5, y: 2.0, w: 9, h: 1.5, fontSize: 44, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `middle` });
  s.addText(`—— 课时一结束 ——`, { x: 4.0, y: 4.8, w: 2, h: 0.4, fontSize: 12, fontFace: FONT, color: `908070`, align: `center` });
  pn(s, 17);
}

// ===== P18 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  rect(s, 2.5, 0.8, 5, 1.0, { fill: { color: `1A2E1A` }, rectRadius: 0.08 });
  s.addText(`✓  任务一已完成`, { x: 2.5, y: 0.85, w: 5, h: 0.5, fontSize: 22, fontFace: FONT, bold: true, color: C.successGreen, align: `center` });
  s.addText(`群众参与度提升  |  群众调研已完成`, { x: 2.5, y: 1.3, w: 5, h: 0.3, fontSize: 12, fontFace: FONT, color: `80A080`, align: `center` });
  rect(s, 2.0, 2.3, 6, 1.5, { fill: { color: `2A1515` }, rectRadius: 0.08 });
  rect(s, 2.0, 2.3, 6, 0.05, { fill: { color: C.alertRed } });
  s.addText(`⚠  新问题出现`, { x: 2.0, y: 2.5, w: 6, h: 0.5, fontSize: 22, fontFace: FONT, bold: true, color: C.alertRed, align: `center` });
  s.addText(`即使群众愿意参与，云岭村的发展仍然推进缓慢。`, { x: 2.2, y: 3.1, w: 5.6, h: 0.5, fontSize: 15, fontFace: FONT, color: `D0A0A0`, align: `center` });
  s.addText(`任务二：为什么云岭村的发展总是「卡住」？`, { x: 0.5, y: 4.3, w: 9, h: 0.8, fontSize: 26, fontFace: FONT, bold: true, color: C.white, align: `center` });
  s.addText(`——从地理视角重新认识乡村问题`, { x: 0.5, y: 4.95, w: 9, h: 0.4, fontSize: 14, fontFace: FONT, color: `A0B0C0`, align: `center` });
  pn(s, 18);
}

// ===== P19 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  placeholder(s, 0.5, 0.3, 9, 3.5, `【AI视频-可灵AI】15秒无人机航拍\n山地村落分散→盘山公路→塌方路段→梯田→货车运输`, `可灵AI文生视频。分段提示词见文档P19。访问 kling.kuaishou.com`);
  s.addText(`云岭村发展数据问题面板`, { x: 0.5, y: 4.0, w: 9, h: 0.4, fontSize: 15, fontFace: FONT, bold: true, color: C.systemBlue });
  const probs = [`雨季道路频繁中断`, `农产品运输成本高`, `青壮年持续外流`, `山地耕地破碎`, `产业收益不稳定`];
  probs.forEach((p, i) => {
    const px = 0.5 + i * 1.85;
    rect(s, px, 4.5, 1.7, 0.7, { fill: { color: `1A1A2E` }, rectRadius: 0.05 });
    s.addText(p, { x: px, y: 4.5, w: 1.7, h: 0.7, fontSize: 11, fontFace: FONT, color: `D0A0A0`, align: `center`, valign: `middle` });
  });
  s.addText(`教师提问：「为什么群众已经愿意参与建设，村庄发展仍然困难重重？」`, { x: 0.5, y: 5.25, w: 9, h: 0.3, fontSize: 11, fontFace: FONT, color: `9090A0`, align: `center` });
  pn(s, 19);
}

// ===== P20 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  rect(s, 1.5, 1.0, 7, 3.5, { fill: { color: `1A1A2E` }, rectRadius: 0.08 });
  rect(s, 1.5, 1.0, 7, 0.05, { fill: { color: C.earthGreen } });
  s.addText(`【系统检测】`, { x: 2.0, y: 1.3, w: 6, h: 0.5, fontSize: 18, fontFace: FONT, bold: true, color: C.earthGreen });
  s.addText(`当前问题涉及：自然环境与乡村发展的关系。\n\n是否接入「地理智库系统」？`, { x: 2.0, y: 1.9, w: 6, h: 1.2, fontSize: 18, fontFace: FONT, color: `C8C8D0`, align: `center` });
  rect(s, 3.5, 3.3, 3, 0.7, { fill: { color: C.earthGreen }, rectRadius: 0.08 });
  s.addText(`接  入`, { x: 3.5, y: 3.3, w: 3, h: 0.7, fontSize: 22, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `middle` });
  s.addText(`竺可桢地理研究档案 · 人地关系与区域发展`, { x: 0.5, y: 4.8, w: 9, h: 0.4, fontSize: 12, fontFace: FONT, color: `80A080`, align: `center` });
  pn(s, 20);
}

// ===== P21 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`云岭村地形诊断（一）：为什么路总修不好？`, { x: 0.5, y: 0.15, w: 9, h: 0.55, fontSize: 22, fontFace: FONT, bold: true, color: C.textDark });
  placeholder(s, 0.3, 0.8, 5.5, 3.6, `【AI生图-豆包+手动标注】云岭村地形图\n（山地海拔/聚落/道路/塌方区）\n生成底图后在PPT中添加标注线`, `豆包提示词：中国西南喀斯特山区地形鸟瞰示意图，山地起伏、河谷、分散小块聚落、蜿蜒山路，类似地理教科书插图但更真实，俯视角，自然色调`);
  s.addText(`教师引导问题`, { x: 6.1, y: 0.8, w: 3.5, h: 0.4, fontSize: 15, fontFace: FONT, bold: true, color: C.textDark });
  [`为什么这里的道路建设成本特别高？`, `为什么雨季容易中断交通？`, `为什么村民会分散居住在不同山坡？`].forEach((q, i) => {
    rect(s, 6.1, 1.3 + i * 0.65, 3.5, 0.55, { fill: { color: C.cardBg }, rectRadius: 0.05 });
    s.addText(q, { x: 6.2, y: 1.3 + i * 0.65, w: 3.3, h: 0.55, fontSize: 11, fontFace: FONT, color: C.textDark, valign: `middle` });
  });
  rect(s, 0.5, 4.6, 9, 0.75, { fill: { color: C.lightOverlay }, rectRadius: 0.05 });
  s.addText(`【知识点】地形会影响：聚落分布  ·  交通建设  ·  基础设施成本\n「自然地理环境，会深刻影响乡村发展方式。」`, { x: 0.7, y: 4.6, w: 8.6, h: 0.75, fontSize: 14, fontFace: FONT, color: C.deepBrown, align: `center`, valign: `middle` });
  pn(s, 21);
}

// ===== P22 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`【知识点总结】地形如何影响乡村发展`, { x: 0.5, y: 0.2, w: 9, h: 0.55, fontSize: 24, fontFace: FONT, bold: true, color: C.textDark });
  const cards = [
    { i: `\u{1F3D8}`, t: `聚落分布`, d: `山区聚落沿坡地、谷地\n分散分布，难以形成\n规模化基础设施` },
    { i: `\u{1F6E3}`, t: `交通建设`, d: `喀斯特山区施工难度大\n道路建设与养护成本\n远高于平原地区` },
    { i: `⚡`, t: `基础设施成本`, d: `分散聚落导致供水、\n供电、网络等基础设施\n人均成本大幅上升` },
  ];
  cards.forEach((c, i) => {
    const cx = 0.7 + i * 3.1;
    rect(s, cx, 1.1, 2.8, 3.2, { fill: { color: C.white }, line: { color: C.lightLine, width: 1 }, rectRadius: 0.08 });
    rect(s, cx, 1.1, 2.8, 0.06, { fill: { color: C.earthGreen } });
    s.addText(c.i, { x: cx, y: 1.4, w: 2.8, h: 0.6, fontSize: 30, align: `center` });
    s.addText(c.t, { x: cx + 0.2, y: 2.1, w: 2.4, h: 0.4, fontSize: 17, fontFace: FONT, bold: true, color: C.earthGreen, align: `center` });
    s.addText(c.d, { x: cx + 0.2, y: 2.6, w: 2.4, h: 1.2, fontSize: 12, fontFace: FONT, color: C.textGray, align: `center` });
  });
  rect(s, 0.5, 4.6, 9, 0.6, { fill: { color: C.lightOverlay }, rectRadius: 0.03 });
  s.addText(`核心结论：有些问题，并不仅仅是「努力」就能解决的。需要科学分析自然条件。`, { x: 0.7, y: 4.6, w: 8.6, h: 0.6, fontSize: 14, fontFace: FONT, color: C.deepBrown, align: `center`, valign: `middle` });
  pn(s, 22);
}

// ===== P23 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`云岭村地形诊断（二）：为什么种了果树还是赚不到钱？`, { x: 0.5, y: 0.15, w: 9, h: 0.55, fontSize: 20, fontFace: FONT, bold: true, color: C.textDark });
  placeholder(s, 0.3, 0.8, 4.5, 2.6, `【Canva制作】柑橘产业运输路线图\n标注：山区路线→收购站→县城市场\n+豆包配图：山区货车运输`, `豆包提示词：中国西南山区盘山公路，小型货车满载柑橘缓慢行驶，路面狭窄，真实摄影风格`);
  rect(s, 5.1, 0.8, 4.6, 2.6, { fill: { color: `111122` }, rectRadius: 0.05 });
  s.addText(`竺可桢地理智库 · AI对话`, { x: 5.3, y: 0.85, w: 4.2, h: 0.3, fontSize: 11, fontFace: FONT, bold: true, color: C.earthGreen });
  s.addText(`为什么自然环境会影响\n经济发展？`, { x: 5.3, y: 1.25, w: 3.0, h: 0.65, fontSize: 9, fontFace: FONT, color: `80C080` });
  s.addText(`自然环境会影响交通、农业与\n产业发展条件。不同地区的发\n展方式，需要结合当地地理条件。`, { x: 5.3, y: 2.0, w: 4.2, h: 0.8, fontSize: 9, fontFace: FONT, color: `C0C0D0` });
  s.addText(`输入问题...`, { x: 5.4, y: 3.0, w: 3.0, h: 0.3, fontSize: 9, fontFace: FONT, color: `606080` });
  rect(s, 8.5, 3.0, 0.8, 0.3, { fill: { color: C.earthGreen }, rectRadius: 0.02 });
  s.addText(`发送`, { x: 8.5, y: 3.0, w: 0.8, h: 0.3, fontSize: 8, fontFace: FONT, color: C.white, align: `center`, valign: `middle` });
  s.addText(`教师连续追问：`, { x: 0.5, y: 3.6, w: 9, h: 0.3, fontSize: 14, fontFace: FONT, bold: true, color: C.textDark });
  [`问题只是「不会种植」吗？`, `为什么交通会影响产业收益？`, `为什么山区产业更容易面临风险？`].forEach((q, i) => {
    rect(s, 0.5 + i * 3.1, 4.0, 2.9, 0.7, { fill: { color: C.cardBg }, rectRadius: 0.05 });
    s.addText(q, { x: 0.6 + i * 3.1, y: 4.0, w: 2.7, h: 0.7, fontSize: 12, fontFace: FONT, color: C.textDark, align: `center`, valign: `middle` });
  });
  s.addNotes(`第三处AI智能体使用：用扣子Coze创建「竺可桢地理智库」Bot，人设提示词见文档P23。与P11的「黄文秀AI智能体」形成「双智囊」结构。`);
  pn(s, 23);
}

// ===== P24 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`《云岭村地理诊断卡》`, { x: 0.5, y: 0.2, w: 9, h: 0.55, fontSize: 26, fontFace: FONT, bold: true, color: C.textDark });
  const dcards = [
    { n: `01`, t: `自然条件限制`, d: `云岭村面临哪些\n自然条件限制？\n请结合地形、气候、\n土壤等方面分析` },
    { n: `02`, t: `交通与产业影响`, d: `地形如何影响\n交通与产业发展？\n分析道路建设、运输\n成本与市场接入` },
    { n: `03`, t: `发展风险评估`, d: `为什么山区乡村\n发展风险更高？\n从自然灾害、市场\n波动等角度分析` },
    { n: `04`, t: `人地关系总结`, d: `自然环境与乡村\n发展之间的关系？\n用具体案例说明\n人地互动规律` },
  ];
  dcards.forEach((c, i) => {
    const cx = 0.4 + i * 2.4;
    rect(s, cx, 1.0, 2.2, 3.5, { fill: { color: C.white }, line: { color: C.lightLine, width: 1 }, rectRadius: 0.08 });
    rect(s, cx, 1.0, 2.2, 0.06, { fill: { color: C.earthGreen } });
    s.addText(c.n, { x: cx + 0.15, y: 1.2, w: 0.6, h: 0.45, fontSize: 28, fontFace: FONT, bold: true, color: C.earthGreen });
    s.addText(c.t, { x: cx + 0.15, y: 1.7, w: 1.9, h: 0.35, fontSize: 15, fontFace: FONT, bold: true, color: C.textDark });
    s.addText(c.d, { x: cx + 0.15, y: 2.2, w: 1.9, h: 1.5, fontSize: 11, fontFace: FONT, color: C.textGray });
    rect(s, cx + 0.3, 3.9, 1.6, 0.35, { fill: { color: C.earthGreen }, rectRadius: 0.03 });
    s.addText(`在线提交 →`, { x: cx + 0.3, y: 3.9, w: 1.6, h: 0.35, fontSize: 10, fontFace: FONT, color: C.white, align: `center`, valign: `middle` });
  });
  rect(s, 0.5, 4.8, 9, 0.55, { fill: { color: C.lightOverlay }, rectRadius: 0.03 });
  s.addText(`「真正的乡村建设，既要理解群众，也要理解土地。」`, { x: 0.7, y: 4.8, w: 8.6, h: 0.55, fontSize: 18, fontFace: FONT, bold: true, color: C.deepBrown, align: `center`, valign: `middle` });
  pn(s, 24);
}

// ===== P25 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  s.addText(`任务三：什么才是真正适合云岭村的发展道路？`, { x: 0.5, y: 0.2, w: 9, h: 0.7, fontSize: 24, fontFace: FONT, bold: true, color: C.white });
  s.addText(`——新时代乡村共创行动`, { x: 0.5, y: 0.8, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: `A0B0C0` });
  const panels = [
    { t: `群众问题`, c: C.alertOrange, items: [`村民担忧产业失败`, `青壮年外流严重`, `群众参与度需提升`, `信任仍在重建中`] },
    { t: `地理条件`, c: C.earthGreen, items: [`喀斯特山区地形`, `山地聚落分散`, `雨季道路塌方`, `山地农业基础`] },
    { t: `已有资源`, c: C.systemBlue, items: [`柑橘种植基础`, `山地生态资源`, `部分返乡青年`, `基本道路条件`] },
  ];
  panels.forEach((p, i) => {
    const px = 0.5 + i * 3.15;
    rect(s, px, 1.4, 3.0, 3.2, { fill: { color: `1A1A2E` }, rectRadius: 0.08 });
    rect(s, px, 1.4, 3.0, 0.06, { fill: { color: p.c } });
    s.addText(p.t, { x: px + 0.2, y: 1.6, w: 2.6, h: 0.4, fontSize: 17, fontFace: FONT, bold: true, color: p.c });
    p.items.forEach((item, j) => {
      s.addText(`•  ${item}`, { x: px + 0.2, y: 2.2 + j * 0.5, w: 2.6, h: 0.4, fontSize: 12, fontFace: FONT, color: `C0C0D0` });
    });
  });
  rect(s, 1.5, 4.8, 7, 0.55, { fill: { color: `1D2E1D` }, rectRadius: 0.05 });
  s.addText(`【系统提示】请结合前两次任务成果，完成云岭村乡村共创行动方案设计。`, { x: 1.7, y: 4.8, w: 6.6, h: 0.55, fontSize: 14, fontFace: FONT, color: `A0D0A0`, align: `center`, valign: `middle` });
  pn(s, 25);
}

// ===== P26 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`三种发展方案对比分析`, { x: 0.5, y: 0.1, w: 9, h: 0.5, fontSize: 24, fontFace: FONT, bold: true, color: C.textDark });
  const plans = [
    { t: `方案A\n大规模工业开发`, img: `【豆包】大型工厂与\n山区对比图`, analysis: [`短期收益较高`, `生态风险较大`, `部分群众反对`, `山区施工难度高`], c: C.alertRed, v: `❌ 不适合` },
    { t: `方案B\n网红旅游村建设`, img: `【豆包】网红民宿\n与山村对比图`, analysis: [`前期投入较大`, `交通条件不足`, `同质化风险明显`, `季节性强`], c: C.alertOrange, v: `⚠ 风险大` },
    { t: `方案C\n山地特色农业\n+合作共创`, img: `【即梦AI】合作社\n村民+电商场景`, analysis: [`依托现有产业基础`, `群众参与度较高`, `风险相对较低`, `更符合当地条件`], c: C.successGreen, v: `✓ 推荐` },
  ];
  plans.forEach((p, i) => {
    const px = 0.3 + i * 3.2;
    rect(s, px, 0.8, 3.0, 4.3, { fill: { color: C.white }, line: { color: p.c, width: 2 }, rectRadius: 0.08 });
    rect(s, px, 0.8, 3.0, 0.95, { fill: { color: p.c }, rectRadius: 0.05 });
    s.addText(p.t, { x: px + 0.1, y: 0.85, w: 2.8, h: 0.85, fontSize: 15, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `middle` });
    placeholder(s, px + 0.2, 1.95, 2.6, 1.5, p.img, ``);
    p.analysis.forEach((a, j) => {
      s.addText(`${j + 1}. ${a}`, { x: px + 0.2, y: 3.6 + j * 0.3, w: 2.6, h: 0.25, fontSize: 9, fontFace: FONT, color: C.textDark });
    });
    rect(s, px + 0.3, 4.7, 2.4, 0.3, { fill: { color: p.c, transparency: 85 }, rectRadius: 0.03 });
    s.addText(p.v, { x: px + 0.3, y: 4.7, w: 2.4, h: 0.3, fontSize: 14, fontFace: FONT, bold: true, color: p.c, align: `center`, valign: `middle` });
  });
  s.addText(`教师提问：「什么样的发展道路，才是真正适合云岭村的？」`, { x: 0.5, y: 5.2, w: 9, h: 0.3, fontSize: 12, fontFace: FONT, color: C.systemBlue, align: `center` });
  pn(s, 26);
}

// ===== P27 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  rect(s, 0, 0, 5, 5.63, { fill: { color: `1C1410` } });
  s.addText(`黄文秀精神档案`, { x: 0.3, y: 0.5, w: 4.4, h: 0.5, fontSize: 18, fontFace: FONT, bold: true, color: C.warmBrown });
  placeholder(s, 0.5, 1.2, 4.0, 2.2, `【复用P9劳动图】\n暖色调处理`, ``);
  s.addText(`「真正能留下来的产业，\n一定是群众愿意参与、\n能够参与的产业。」`, { x: 0.5, y: 3.6, w: 4.0, h: 1.0, fontSize: 18, fontFace: FONT, italic: true, color: `D4A870`, align: `center` });
  rect(s, 4.95, 0.5, 0.1, 4.5, { fill: { color: `4A4A4A` } });
  rect(s, 5, 0, 5, 5.63, { fill: { color: `101C14` } });
  s.addText(`竺可桢地理档案`, { x: 5.3, y: 0.5, w: 4.4, h: 0.5, fontSize: 18, fontFace: FONT, bold: true, color: C.earthGreen });
  placeholder(s, 5.5, 1.2, 4.0, 2.2, `【复用P21地形图】\n冷色调处理`, ``);
  s.addText(`「适合土地的发展，\n才能真正持续下去。」`, { x: 5.5, y: 3.6, w: 4.0, h: 1.0, fontSize: 18, fontFace: FONT, italic: true, color: `80C0A0`, align: `center` });
  rect(s, 4.3, 4.7, 1.4, 0.5, { fill: { color: `3A3A3A` }, rectRadius: 0.05 });
  s.addText(`共同指向 →`, { x: 4.3, y: 4.7, w: 1.4, h: 0.5, fontSize: 10, fontFace: FONT, color: `C0C0C0`, align: `center`, valign: `middle` });
  pn(s, 27);
}

// ===== P28 =====
{
  const s = pptx.addSlide(); bg(s, C.bg);
  s.addText(`《云岭村乡村共创行动方案》`, { x: 0.5, y: 0.15, w: 9, h: 0.55, fontSize: 26, fontFace: FONT, bold: true, color: C.textDark });
  s.addText(`请结合《群众工作方案》+《地理诊断卡》，完成以下五个模块：`, { x: 0.5, y: 0.7, w: 9, h: 0.35, fontSize: 13, fontFace: FONT, color: C.textGray });
  const mods = [
    { n: `01`, t: `群众需求分析`, d: `村民最关心的\n问题是什么？\n如何保障群众\n真正受益？` },
    { n: `02`, t: `地理条件分析`, d: `地形、交通、区位\n对方案的制约与\n机遇是什么？` },
    { n: `03`, t: `产业发展路径`, d: `选择什么产业？\n为什么选它？\n如何落地实施？` },
    { n: `04`, t: `群众参与方式`, d: `如何发动群众\n共同参与？\n利益如何分配？` },
    { n: `05`, t: `可持续发展`, d: `生态保护、\n长期效益、\n风险应对方案` },
  ];
  mods.forEach((m, i) => {
    const mx = 0.3 + i * 1.92;
    rect(s, mx, 1.2, 1.8, 3.2, { fill: { color: C.white }, line: { color: C.lightLine, width: 1 }, rectRadius: 0.05 });
    rect(s, mx, 1.2, 1.8, 0.06, { fill: { color: C.warmBrown } });
    s.addText(m.n, { x: mx + 0.1, y: 1.35, w: 0.5, h: 0.4, fontSize: 22, fontFace: FONT, bold: true, color: C.warmBrown });
    s.addText(m.t, { x: mx + 0.1, y: 1.75, w: 1.6, h: 0.4, fontSize: 14, fontFace: FONT, bold: true, color: C.textDark });
    s.addText(m.d, { x: mx + 0.1, y: 2.2, w: 1.6, h: 1.8, fontSize: 10, fontFace: FONT, color: C.textGray });
  });
  rect(s, 0.5, 4.6, 9, 0.8, { fill: { color: C.lightOverlay }, rectRadius: 0.05 });
  s.addText(`\u{1F4F1} 提交方式：扫码进入「乡村共创任务平台」→ 选择「乡村共创行动方案」表单 → 在线填写提交`, { x: 0.7, y: 4.6, w: 6.5, h: 0.4, fontSize: 12, fontFace: FONT, color: C.textDark });
  s.addText(`系统将自动汇总全班方案并生成分析报告，第二课时进行成果展示与互评。`, { x: 0.7, y: 5.0, w: 6.5, h: 0.3, fontSize: 11, fontFace: FONT, color: C.textGray });
  placeholder(s, 7.8, 4.55, 1.4, 0.9, `【简道云二维码】`, `在简道云创建表单后生成二维码`);
  s.addNotes(`创新亮点：完整的「提交→分析→反馈」闭环。学生提交后，教师可用Claude Code脚本对接简道云API自动分析方案。详见文档P28。`);
  pn(s, 28);
}

// ===== P29 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  placeholder(s, 0, 0, 10, 5.63, `【AI视频-可灵AI】延时摄影\n乡村从清晨到夜晚，山村灯光亮起\n（用作背景，加半透明遮罩）`, `可灵AI提示词：中国乡村从清晨到夜晚的延时摄影，山村灯光亮起，年轻人走在村路上，温暖夜色，微电影质感`);
  rect(s, 0, 0, 10, 5.63, { fill: { color: `000000`, transparency: 45 } });
  s.addText(`【新时代身份确认】`, { x: 1.0, y: 0.8, w: 8, h: 0.6, fontSize: 20, fontFace: FONT, bold: true, color: C.systemBlue, align: `center` });
  s.addText(`乡村振兴不是历史中的任务。`, { x: 1.0, y: 2.0, w: 8, h: 0.6, fontSize: 28, fontFace: FONT, bold: true, color: C.white, align: `center` });
  s.addText(`而是属于这一代青年的新时代长征。`, { x: 1.0, y: 2.6, w: 8, h: 0.6, fontSize: 28, fontFace: FONT, bold: true, color: C.white, align: `center` });
  rect(s, 1.5, 3.6, 7, 0.8, { fill: { color: `1D2E1D` }, rectRadius: 0.08 });
  s.addText(`【最终问题】如果未来有一天，你的家乡也需要建设，你愿意回来吗？`, { x: 1.7, y: 3.6, w: 6.6, h: 0.8, fontSize: 18, fontFace: FONT, bold: true, color: `C0E0C0`, align: `center`, valign: `middle` });
  s.addText(`（留给学生思考与回答的安静时刻）`, { x: 1.0, y: 4.8, w: 8, h: 0.4, fontSize: 12, fontFace: FONT, color: `808090`, align: `center` });
  s.addNotes(`课堂操作：这一页留白10-15秒。让学生真正思考「我愿不愿意」。教师邀请2-3名学生自由表达。不需要标准答案。`);
  pn(s, 29);
}

// ===== P30 =====
{
  const s = pptx.addSlide(); bg(s, C.darkBg);
  placeholder(s, 0, 0, 10, 5.63, `【AI视频-可灵AI+必剪】35-40秒最终蒙太奇\n黄文秀背影→帮村民劳动→航拍村庄→村民议事→青年返乡→山村夜景灯光→渐暗`, `用可灵AI「图生视频」将前面各页生成的静态图转为5秒动态片段，必剪拼接。详见文档P30。这是本作品最重要的一段视频。`);
  rect(s, 0, 0, 10, 5.63, { fill: { color: `000000`, transparency: 55 } });
  s.addText(`如果她（他）活到今天`, { x: 0.5, y: 2.0, w: 9, h: 1.0, fontSize: 42, fontFace: FONT, bold: true, color: C.white, align: `center`, valign: `bottom` });
  s.addText(`我们无法替他们重写历史。`, { x: 0.5, y: 3.8, w: 9, h: 0.5, fontSize: 22, fontFace: FONT, color: `C0B8B0`, align: `center` });
  s.addText(`但我们可以接过他们未走完的长征。`, { x: 0.5, y: 4.3, w: 9, h: 0.6, fontSize: 24, fontFace: FONT, bold: true, color: C.white, align: `center` });
  pn(s, 30);
}

// ===== SAVE =====
const outputPath = path.join(__dirname, `AI乡村共创课堂_30页完整版.pptx`);
pptx.writeFile({ fileName: outputPath }).then(() => {
  console.log(`✅ PPT生成成功：${outputPath}`);
  console.log(`\u{1F4CA} 共30页`);
  console.log(``);
  console.log(`\u{1F4CC} 下一步操作：`);
  console.log(`1. 用豆包/即梦AI生成各页标注的图片，替换PPT中的占位图`);
  console.log(`2. 用可灵AI (kling.kuaishou.com) 生成视频片段，用必剪拼接`);
  console.log(`3. 在扣子Coze (coze.cn) 创建两个AI智能体，嵌入P11和P23`);
  console.log(`4. 在简道云 (jiandaoyun.com) 搭建任务提交平台，二维码替换P16/P28`);
  console.log(`5. 根据需要微调字体、动画效果和图片位置`);
}).catch(err => {
  console.error(`生成失败：`, err);
});
