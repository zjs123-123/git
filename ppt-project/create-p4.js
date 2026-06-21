const PptxGenJS = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const pptx = new PptxGenJS();
pptx.layout = "LAYOUT_WIDE";
pptx.author = "张家硕";

const IMG = "C:/Users/ASUS/Desktop/Claude/比赛作品";
const FONT = "Microsoft YaHei";

// 读取图片为 base64
function readImg(filename) {
  const filePath = path.join(IMG, filename);
  const buf = fs.readFileSync(filePath);
  return buf.toString("base64");
}

const imgVillage = readImg("村庄俯视图.png");
const imgInfo    = readImg("信息图模板.png");
const imgSys     = readImg("系统界面展示图模板.png");

const s = pptx.addSlide();
s.background = { fill: "FAF7F2" };

// ===== 顶部标题栏 =====
s.addText("云岭村基础资料", {
  x: 0.5, y: 0.15, w: 5, h: 0.55,
  fontSize: 26, fontFace: FONT, bold: true, color: "2C2C2C",
});
s.addShape("rect", {
  x: 0.5, y: 0.68, w: 2.2, h: 0.035,
  fill: { color: "8B6914" },
});

// 右上角标签
s.addShape("rect", {
  x: 7.3, y: 0.2, w: 2.4, h: 0.4,
  fill: { color: "5C3D0E" }, rectRadius: 0.04,
});
s.addText("【系统自动加载】", {
  x: 7.3, y: 0.2, w: 2.4, h: 0.4,
  fontSize: 10, fontFace: FONT, color: "FFFFFF", align: "center", valign: "middle",
});

// ===== 左侧：村庄俯视图 =====
s.addImage({
  data: "data:image/png;base64," + imgVillage,
  x: 0.3, y: 0.95, w: 4.7, h: 3.15,
  sizing: { type: "cover", w: 4.7, h: 3.15 },
});
// 图片底部标签
s.addText("▲ 云岭村俯视图 · 喀斯特山区聚落分布", {
  x: 0.3, y: 4.15, w: 4.7, h: 0.25,
  fontSize: 9, fontFace: FONT, color: "787880", align: "center",
});

// ===== 右侧：数据面板 =====
s.addShape("rect", {
  x: 5.2, y: 0.95, w: 4.5, h: 3.45,
  fill: { color: "FFFFFF" },
  line: { color: "E8E0D5", width: 0.5 },
  rectRadius: 0.06,
});

s.addText("村庄关键数据", {
  x: 5.5, y: 1.05, w: 3.9, h: 0.35,
  fontSize: 14, fontFace: FONT, bold: true, color: "5C3D0E",
});

// 信息图模板小图
s.addImage({
  data: "data:image/png;base64," + imgInfo,
  x: 8.1, y: 1.05, w: 1.3, h: 1.8,
  sizing: { type: "contain", w: 1.3, h: 1.8 },
});

// 五条数据
const dataCards = [
  { icon: "🏘", label: "户数",      value: "286 户",           note: "分散于 6 个山坡片区" },
  { icon: "🍊", label: "主要作物",  value: "柑橘",             note: "种植面积约 800 亩" },
  { icon: "🛣", label: "交通",      value: "距县城约 47 公里",  note: "仅一条盘山公路连通" },
  { icon: "🌧", label: "雨季",      value: "5 — 8 月",        note: "年均塌方阻断 3—5 次" },
  { icon: "👥", label: "青壮年外流", value: "约 62%",          note: "常住人口以老人妇女为主" },
];

dataCards.forEach((d, i) => {
  const dy = 1.5 + i * 0.56;
  s.addShape("rect", {
    x: 5.4, y: dy, w: 4.1, h: 0.48,
    fill: { color: "FBF8F2" },
    rectRadius: 0.04,
  });
  s.addShape("rect", {
    x: 5.4, y: dy + 0.06, w: 0.04, h: 0.36,
    fill: { color: "8B6914" },
  });
  s.addText(`${d.icon}  ${d.label}`, {
    x: 5.6, y: dy + 0.02, w: 1.6, h: 0.22,
    fontSize: 10, fontFace: FONT, color: "787880", valign: "middle",
  });
  s.addText(d.value, {
    x: 5.6, y: dy + 0.22, w: 1.6, h: 0.24,
    fontSize: 15, fontFace: FONT, bold: true, color: "2C2C2C", valign: "middle",
  });
  s.addText(d.note, {
    x: 7.1, y: dy + 0.08, w: 2.3, h: 0.32,
    fontSize: 9, fontFace: FONT, color: "787880", valign: "middle",
  });
});

// ===== 底部 =====
s.addImage({
  data: "data:image/png;base64," + imgSys,
  x: 0.3, y: 4.55, w: 9.4, h: 0.45,
  sizing: { type: "cover", w: 9.4, h: 0.45 },
});

// 思考问题
s.addShape("rect", {
  x: 0.5, y: 5.0, w: 9.0, h: 0.45,
  fill: { color: "F5EDE0" },
  rectRadius: 0.04,
});
s.addText("🤔  思考：为什么一个拥有自然资源和产业基础的村庄，发展仍然困难重重？", {
  x: 0.7, y: 5.0, w: 8.6, h: 0.45,
  fontSize: 14, fontFace: FONT, color: "5C3D0E", valign: "middle",
});

// 保存
const outputPath = path.join(__dirname, "P4_云岭村基础资料.pptx");
pptx.writeFile({ fileName: outputPath }).then(() => {
  console.log("✅ P4 生成成功！");
  console.log("📁 " + outputPath);
}).catch(err => {
  console.error("生成失败：", err);
});
