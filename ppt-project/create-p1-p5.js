const PptxGenJS = require("pptxgenjs");
const path = require("path");

// ===== 配色体系 =====
const C = {
  // 大地色系
  darkEarth:  "1C1410",
  warmBrown:  "8B6914",
  deepBrown:  "5C3D0E",
  riceGold:   "C8A96E",
  cream:      "F5F0E8",
  // 山脉灰绿
  earthGreen: "5B8C5A",
  mossGreen:  "4A6B3A",
  sageGreen:  "8AAA80",
  // 科技蓝灰（系统提示色）
  darkNavy:   "0D1117",
  panelBg:    "151A24",
  systemBlue: "3A7BD5",
  techCyan:   "4ECDC4",
  // 警示色
  alertRed:   "D94A4A",
  alertOrange:"E8913A",
  successGreen:"4CAF50",
  // 基础色
  white:      "FFFFFF",
  textDark:   "2C2C2C",
  textGray:   "787880",
  lightBg:    "FAF7F2",
  cardBg:     "F5EDE0",
  // 功能色
  softGold:   "D4C8B0",
  dimGold:    "A09080",
  darkOverlay:"000000",
};

const FONT = "Microsoft YaHei";
const MONO = "Consolas";

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "张家硕";
pptx.title = "如果她（他）活到今天——前5页预览";

const SW = 10;  // slide width in inches
const SH = 5.63; // slide height in inches

// ===== 工具函数 =====
function bg(s, color) { s.background = { fill: color }; }

function shape(s, type, x, y, w, h, opts = {}) {
  return s.addShape(type, { x, y, w, h, ...opts });
}

function rect(s, x, y, w, h, opts = {}) {
  return shape(s, "rect", x, y, w, h, opts);
}

function roundRect(s, x, y, w, h, r = 0.05, opts = {}) {
  return rect(s, x, y, w, h, { rectRadius: r, ...opts });
}

function txt(s, text, x, y, w, h, opts = {}) {
  return s.addText(text, {
    x, y, w, h,
    fontFace: opts.fontFace || FONT,
    fontSize: opts.fontSize || 14,
    color: opts.color || C.textDark,
    bold: opts.bold || false,
    italic: opts.italic || false,
    align: opts.align || "left",
    valign: opts.valign || "top",
    ...opts,
  });
}

function placeholder(s, x, y, w, h, label, toolNote) {
  roundRect(s, x, y, w, h, 0.04, {
    fill: { color: "EDE8DD" },
    line: { color: C.warmBrown, width: 1.5, dashType: "dash" },
  });
  txt(s, `📷 ${label}`, x + 0.15, y + 0.15, w - 0.3, h - 0.3, {
    fontSize: 10, color: C.textGray, align: "center", valign: "middle",
  });
  if (toolNote) s.addNotes(toolNote);
}

function pageNum(s, n) {
  txt(s, `${n} / 30`, SW - 1.2, SH - 0.45, 1, 0.35, {
    fontSize: 8, color: "999999", align: "right",
  });
}

// ============================================================
//  P1 — 封面：如果她（他）活到今天
// ============================================================
{
  const s = pptx.addSlide();
  bg(s, C.darkEarth);

  // 模拟山脉轮廓 —— 用多层不规则形状营造山区氛围
  // 远山（淡）
  s.addShape("triangle", {
    x: -0.5, y: 2.8, w: 3.0, h: 2.4,
    fill: { color: "2A2520", transparency: 30 },
  });
  s.addShape("triangle", {
    x: 1.5, y: 2.5, w: 3.5, h: 2.8,
    fill: { color: "252018", transparency: 20 },
  });
  s.addShape("triangle", {
    x: 4.0, y: 2.6, w: 4.0, h: 2.6,
    fill: { color: "2A2220", transparency: 25 },
  });
  s.addShape("triangle", {
    x: 6.5, y: 2.3, w: 4.5, h: 3.0,
    fill: { color: "221D18", transparency: 15 },
  });

  // 中景山
  rect(s, 0, 3.8, SW, 1.2, {
    fill: { color: "1A1510", transparency: 10 },
  });

  // 近景地面
  rect(s, 0, 4.6, SW, 1.03, { fill: { color: "0D0B08" } });

  // 地面纹理线
  rect(s, 0, 4.65, SW, 0.02, { fill: { color: "3A3020", transparency: 60 } });

  // 薄雾效果 —— 半透明白色矩形
  rect(s, 0, 3.0, SW, 1.5, {
    fill: { color: "E8E0D0", transparency: 92 },
  });

  // 标题区
  txt(s, "《如果她（他）活到今天》", 0.6, 3.4, 8.8, 0.9, {
    fontSize: 46, color: C.white, bold: true, align: "left", valign: "bottom",
  });

  // 金色装饰线
  rect(s, 0.6, 4.35, 2.5, 0.035, { fill: { color: C.riceGold } });

  // 副标题
  txt(s, "AI驱动下新时代乡村共创课堂", 0.6, 4.45, 8.8, 0.45, {
    fontSize: 20, color: C.softGold, align: "left",
  });

  // 底部标签
  txt(s, "文化素养+AI  |  思政×地理  |  高中段", 0.6, 5.05, 8.8, 0.3, {
    fontSize: 11, color: C.dimGold,
  });

  pageNum(s, 1);
  s.addNotes("豆包生图提示词：中国西南山区乡村清晨，薄雾笼罩村庄，远处层叠山峦，近处青瓦白墙农舍，纪录片电影质感，温暖自然光，真实摄影风格，横向构图，留白天空部分作为标题区，乡村肌理清晰可见，不要人物，不要文字。生成后替换背景。");
}

// ============================================================
//  P2 — 课堂启动：新时代乡村共创行动
// ============================================================
{
  const s = pptx.addSlide();
  bg(s, C.darkNavy);

  // 顶部系统状态栏
  rect(s, 0, 0, SW, 0.4, { fill: { color: "05080C" } });
  txt(s, "AI乡村共创系统 v2.4  │  任务状态：待启动  │  实践团编号：YLC-2024-003", 0.3, 0.02, 9.4, 0.36, {
    fontSize: 10, fontFace: MONO, color: "4A9F4A", valign: "middle",
  });

  // 主面板
  roundRect(s, 0.8, 0.75, 8.4, 3.8, 0.08, {
    fill: { color: C.panelBg },
    line: { color: "2A3040", width: 1 },
  });
  // 面板顶部强调线
  rect(s, 0.8, 0.75, 8.4, 0.045, { fill: { color: C.systemBlue } });

  // 系统提示标题
  txt(s, "【 系统提示 】", 1.3, 1.05, 7.4, 0.45, {
    fontSize: 20, color: C.systemBlue, bold: true, valign: "middle",
  });

  // 欢迎语
  txt(s, "欢迎进入「新时代乡村共创行动」", 1.3, 1.55, 7.4, 0.5, {
    fontSize: 24, color: C.white, bold: true,
  });

  txt(s, "你们将接手一座正在推进乡村振兴的村庄。请完成以下三项核心任务：", 1.3, 2.1, 7.4, 0.4, {
    fontSize: 14, color: "B0B8C8",
  });

  // 三张任务卡片
  const tasks = [
    { icon: "📋", title: "任务一", desc: "群众调研与沟通", sub: "了解村民真实想法", color: C.systemBlue },
    { icon: "🔍", title: "任务二", desc: "村庄问题诊断分析", sub: "从地理视角看发展瓶颈", color: C.techCyan },
    { icon: "🏗", title: "任务三", desc: "乡村共创方案设计", sub: "提出可行的发展路径", color: C.successGreen },
  ];
  tasks.forEach((t, i) => {
    const cx = 1.3 + i * 2.8;
    roundRect(s, cx, 2.75, 2.5, 1.45, 0.06, {
      fill: { color: "1A1F2E" },
      line: { color: "2A3050", width: 0.5 },
    });
    // 卡片顶部色条
    rect(s, cx, 2.75, 2.5, 0.04, { fill: { color: t.color } });
    // 图标和标题
    txt(s, `${t.icon}  ${t.title}`, cx + 0.2, 2.9, 2.1, 0.4, {
      fontSize: 15, color: t.color, bold: true, valign: "middle",
    });
    txt(s, t.desc, cx + 0.2, 3.3, 2.1, 0.35, {
      fontSize: 13, color: "D0D4E0", valign: "middle",
    });
    txt(s, t.sub, cx + 0.2, 3.65, 2.1, 0.35, {
      fontSize: 10, color: "7880A0", valign: "middle",
    });
  });

  // 进度条
  roundRect(s, 1.3, 4.35, 7.4, 0.06, 0.03, { fill: { color: "252A40" } });
  rect(s, 1.3, 4.35, 0.4, 0.06, { fill: { color: C.systemBlue } });
  txt(s, "系统初始化中... 5%", 1.3, 4.43, 3, 0.28, {
    fontSize: 9, fontFace: MONO, color: "6070A0",
  });

  // 教师导入语
  txt(s, "教师导入语：今天这节课，你们不只是课堂中的学生。你们将作为「新时代青年乡村共创实践团成员」，进入一个真实的乡村建设任务。", 0.5, 4.85, 9, 0.5, {
    fontSize: 13, color: "90A0B8", align: "center", valign: "middle",
  });

  pageNum(s, 2);
  s.addNotes("系统UI风格用Canva辅助设计（canva.cn搜「科技UI」模板）。系统提示音用PPT自带音频录制简短提示音。");
}

// ============================================================
//  P3 — 云岭村初印象（视频页）
// ============================================================
{
  const s = pptx.addSlide();
  bg(s, "0A0D10");

  // 视频外框 —— 模拟播放器界面
  roundRect(s, 0.6, 0.4, 8.8, 4.2, 0.06, {
    fill: { color: "11161E" },
    line: { color: "2A3040", width: 1.5 },
  });

  // 视频画面占位区
  placeholder(s, 1.0, 0.7, 8.0, 3.3,
    "【AI视频 · 可灵AI】20秒短视频\n\n西南山区航拍 → 云雾村庄 → 盘山公路 → 村民劳作\n（4段×5秒，用必剪拼接）\n\n工具：可灵AI (kling.kuaishou.com)",
    "可灵AI分段提示词见文档P3。4段各5秒：①航拍喀斯特山区云雾 ②盘山公路 ③村庄近景村民劳作 ④镜头推向村口。必剪(bcut.bilibili.cn)免费剪辑。"
  );

  // 播放器控件装饰
  rect(s, 0.6, 4.2, 8.8, 0.4, { fill: { color: "0D1018" } });
  txt(s, "▶  00:00 / 00:20", 0.8, 4.2, 2, 0.4, {
    fontSize: 9, fontFace: MONO, color: "6080A0", valign: "middle",
  });
  txt(s, "云岭村 · 西南喀斯特山区", 4.0, 4.2, 2, 0.4, {
    fontSize: 9, fontFace: MONO, color: "506080", align: "center", valign: "middle",
  });

  // 底部标题
  txt(s, "云岭村  ·  初印象", 0.6, 4.85, 8.8, 0.55, {
    fontSize: 22, color: "C0C8D8", align: "center", bold: true,
  });

  pageNum(s, 3);
  s.addNotes("视频是关键的情绪入口。用可灵AI生成4个片段，必剪拼接。加简单转场，整体色调偏暖。");
}

// ============================================================
//  P4 — 云岭村基础资料（数据面板）
// ============================================================
{
  const s = pptx.addSlide();
  bg(s, C.lightBg);

  // 页面标题
  txt(s, "云岭村基础资料", 0.5, 0.15, 5, 0.55, {
    fontSize: 26, color: C.textDark, bold: true,
  });
  rect(s, 0.5, 0.68, 2.2, 0.035, { fill: { color: C.warmBrown } });

  // 左侧：村庄位置示意图占位
  placeholder(s, 0.5, 0.95, 4.5, 3.2,
    "【AI生图 · 豆包】\n云岭村俯视图 / 位置示意图\n\n山区聚落沿山坡分散分布\n喀斯特地貌特征",
    "豆包提示词：中国西南山区村庄俯视图，村落沿山坡分散分布，喀斯特地貌特征，可见盘山小路连接各户，自然纪录片航拍风格"
  );

  // 右侧：数据卡片面板
  txt(s, "村庄关键数据", 5.4, 0.95, 4.2, 0.35, {
    fontSize: 13, color: C.textGray, bold: true, valign: "middle",
  });

  const dataCards = [
    { label: "户数", value: "286户", note: "分散于6个山坡片区", icon: "🏘" },
    { label: "主要作物", value: "柑橘", note: "种植面积约800亩", icon: "🍊" },
    { label: "交通", value: "距县城47km", note: "仅一条盘山公路连通", icon: "🛣" },
    { label: "雨季", value: "5—8月", note: "年均塌方阻断3—5次", icon: "🌧" },
    { label: "青壮年外流", value: "约62%", note: "常住人口以老人妇女为主", icon: "👥" },
  ];

  dataCards.forEach((d, i) => {
    const dy = 1.4 + i * 0.76;
    roundRect(s, 5.2, dy, 4.3, 0.64, 0.05, {
      fill: { color: C.white },
      line: { color: "E8E0D5", width: 0.5 },
    });
    // 左侧色条
    rect(s, 5.2, dy + 0.08, 0.04, 0.48, { fill: { color: C.warmBrown } });
    // 图标 + 标签
    txt(s, `${d.icon} ${d.label}`, 5.45, dy + 0.05, 1.5, 0.25, {
      fontSize: 10, color: C.textGray, valign: "middle",
    });
    // 数值
    txt(s, d.value, 5.45, dy + 0.28, 1.5, 0.3, {
      fontSize: 18, color: C.textDark, bold: true, valign: "middle",
    });
    // 备注
    txt(s, d.note, 7.0, dy + 0.12, 2.3, 0.4, {
      fontSize: 10, color: C.textGray, valign: "middle",
    });
  });

  // 底部引导问题
  roundRect(s, 0.5, 4.35, 9, 0.6, 0.05, { fill: { color: "F8F2E8" } });
  txt(s, "🤔  思考：为什么一个拥有自然资源和产业基础的村庄，发展仍然困难重重？", 0.7, 4.35, 8.6, 0.6, {
    fontSize: 15, color: C.deepBrown, valign: "middle",
  });

  // 系统状态栏
  roundRect(s, 0.5, 5.1, 9, 0.35, 0.03, { fill: { color: "E8E2D8" } });
  txt(s, "【系统状态】资料加载完成 | 等待任务指令...", 0.7, 5.1, 8.6, 0.35, {
    fontSize: 9, fontFace: MONO, color: C.textGray, valign: "middle",
  });

  pageNum(s, 4);
  s.addNotes("数据面板用Canva辅助（canva.cn搜「信息图」模板）。数据卡片内容可根据实际情况调整。右侧面板和P2保持系统风格一致。");
}

// ============================================================
//  P5 — 任务推送：为什么群众不愿意参与建设？
// ============================================================
{
  const s = pptx.addSlide();
  bg(s, "0D0808");

  // 顶部红色警报条
  rect(s, 0, 0, SW, 0.06, { fill: { color: C.alertRed } });

  // 警告标题面板
  roundRect(s, 1.2, 0.4, 7.6, 0.65, 0.05, { fill: { color: "1E1010" } });
  // 闪烁红点
  s.addShape("ellipse", {
    x: 1.45, y: 0.56, w: 0.22, h: 0.22,
    fill: { color: C.alertRed },
  });
  txt(s, "【当前任务状态】乡村共创行动推进失败", 1.8, 0.4, 6.8, 0.65, {
    fontSize: 20, color: C.alertRed, bold: true, valign: "middle",
  });

  // 三条失败原因
  const reasons = [
    "群众不愿配合调研",
    "村民对实践团缺乏信任",
    "产业方案推进停滞",
  ];
  reasons.forEach((r, i) => {
    const ry = 1.3 + i * 0.55;
    roundRect(s, 1.5, ry, 7.0, 0.46, 0.03, { fill: { color: "150C0C" } });
    txt(s, `✘   ${r}`, 1.7, ry, 6.6, 0.46, {
      fontSize: 14, color: "D08080", valign: "middle",
    });
  });

  // 分隔线
  rect(s, 2.0, 3.05, 6.0, 0.01, { fill: { color: "3A2020" } });

  // 村民反馈信息流标题
  txt(s, "—— 村民反馈信息流 ——", 1.5, 3.2, 7.0, 0.35, {
    fontSize: 11, color: "786868", align: "center",
  });

  // 四条消息气泡（交替左右排列，模拟聊天界面）
  const messages = [
    { text: "以前也有人来调研，最后没下文。", name: "张大爷", avatar: "张", bg: "2A1A1A" },
    { text: "你们年轻人懂农村吗？", name: "李婶", avatar: "李", bg: "1A1A2A" },
    { text: "产业项目以前赔过，干部说得好听。", name: "王叔", avatar: "王", bg: "1A2A1A" },
    { text: "最后还不是我们自己承担风险？", name: "陈伯", avatar: "陈", bg: "2A2A1A" },
  ];

  messages.forEach((msg, i) => {
    const isLeft = i % 2 === 0;
    const mx = isLeft ? 0.6 : 5.3;
    const my = 3.65 + Math.floor(i / 2) * 0.62;

    // 头像圆圈
    s.addShape("ellipse", {
      x: mx, y: my, w: 0.38, h: 0.38,
      fill: { color: msg.bg },
    });
    txt(s, msg.avatar, mx, my, 0.38, 0.38, {
      fontSize: 12, color: "E0D8D0", align: "center", valign: "middle",
    });

    // 消息气泡
    roundRect(s, mx + 0.48, my, 3.6, 0.4, 0.06, {
      fill: { color: msg.bg },
    });
    txt(s, `"${msg.text}"`, mx + 0.6, my, 3.35, 0.4, {
      fontSize: 10.5, color: "D0C8C0", valign: "middle",
    });

    // 发送者名字
    txt(s, msg.name, mx, my + 0.4, 0.38, 0.18, {
      fontSize: 7, color: "706060", align: "center",
    });
  });

  // 底部教师引导
  txt(s, "教师提问：「为什么乡村建设方案还没开始，任务就已经停滞了？」", 0.5, 5.1, 9, 0.4, {
    fontSize: 12, color: "908088", align: "center", valign: "middle",
  });

  pageNum(s, 5);
  s.addNotes("村民头像用豆包生成（提示词：中国农村老年男性/女性正面肖像，60岁左右，黝黑皮肤，生活照质感）。红色警告图标可在Canva搜索「警告图标」修改颜色，或使用PPT自带图标库。消息气泡每条依次弹出（PPT「飞入」动画，间隔0.5秒）。");
}

// ===== 保存 =====
const outputPath = path.join(__dirname, "AI乡村共创课堂_前5页预览.pptx");
pptx.writeFile({ fileName: outputPath }).then(() => {
  console.log("✅ PPT生成成功！");
  console.log(`📁 文件位置：${outputPath}`);
  console.log("");
  console.log("📊 包含页面：");
  console.log("  P1 — 封面：如果她（他）活到今天");
  console.log("  P2 — 课堂启动：新时代乡村共创行动");
  console.log("  P3 — 云岭村初印象（视频页）");
  console.log("  P4 — 云岭村基础资料（数据面板）");
  console.log("  P5 — 任务推送：村民反馈信息流");
  console.log("");
  console.log("📌 后续步骤：");
  console.log("  1. 用豆包/即梦AI生成各页标注的图片，替换占位区");
  console.log("  2. 用可灵AI生成P3视频片段");
  console.log("  3. 为P4/P5中的系统面板添加PPT动画效果");
  console.log("  4. 确认配色和字体满意后，继续制作P6-P30");
}).catch(err => {
  console.error("生成失败：", err);
});
