const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Claude";
pres.title = "AIS 31 随机数生成器功能等级标准 - 教学讲解";

// ============================================================
// Color Palette: Professional Dark Navy + Teal
// ============================================================
const C = {
  darkBg: "0F1B2D",       // very dark navy (title/conclusion bg)
  primary: "1A3A5C",       // dark navy blue
  secondary: "2176A1",     // medium blue
  accent: "00A8CC",        // teal/cyan accent
  lightAccent: "E8F4F8",   // very light blue (card bg)
  white: "FFFFFF",
  offWhite: "F4F7FA",      // slide bg
  textDark: "1A1A2E",      // main text
  textGray: "4A5568",      // muted text
  tableHeader: "1A3A5C",
  tableStripe: "EDF2F7",
  border: "D0D9E4",
  goodGreen: "38A169",
  warnOrange: "DD6B20",
};

// ============================================================
// Reusable factory functions (avoids shared-object mutation bug)
// ============================================================
const mkShadow = () => ({ type: "outer", color: "000000", blur: 4, offset: 1.5, angle: 135, opacity: 0.1 });
const mkCardShadow = () => ({ type: "outer", color: "000000", blur: 3, offset: 1, angle: 135, opacity: 0.08 });
const mkAccentBar = (color) => ({ x: 0.6, y: 0.55, w: 0.06, h: 0.35, fill: { color: color || C.accent } });

// ============================================================
// SLIDE 1: Title
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };

  // Top accent line
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: C.accent } });

  // Decorative shapes
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 0.7, w: 0.5, h: 0.5, fill: { color: C.accent, transparency: 20 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 8.7, y: 4.4, w: 0.5, h: 0.5, fill: { color: C.secondary, transparency: 40 } });

  s.addText("AIS 31", {
    x: 1.2, y: 0.8, w: 7.6, h: 0.7,
    fontSize: 18, fontFace: "Arial", color: C.accent, bold: true,
    charSpacing: 6,
  });

  s.addText("随机数生成器\n功能等级标准", {
    x: 1.2, y: 1.5, w: 7.6, h: 1.8,
    fontSize: 40, fontFace: "Arial Black", color: C.white, bold: true,
    lineSpacingMultiple: 1.3,
  });

  s.addText([
    { text: "教学讲解", options: { fontSize: 22, color: C.white } },
  ], { x: 1.2, y: 3.4, w: 7.6, h: 0.5, fontSize: 22, fontFace: "Arial", color: C.white });

  s.addShape(pres.shapes.RECTANGLE, { x: 1.2, y: 4.05, w: 1.5, h: 0.03, fill: { color: C.accent } });

  s.addText("A Proposal for Functionality Classes for Random Number Generators\nVersion 2.35 DRAFT — BSI (德国联邦信息安全办公室)", {
    x: 1.2, y: 4.3, w: 7.6, h: 0.7,
    fontSize: 12, fontFace: "Arial", color: "8899AA",
  });

  // Bottom bar
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 5.55, w: 10, h: 0.075, fill: { color: C.accent } });
}

// ============================================================
// SLIDE 2: Document Overview
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("文档概述", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  const items = [
    ["文档全称", "A Proposal for Functionality Classes for Random Number Generators"],
    ["作者", "Matthias Peter & Werner Schindler (BSI)"],
    ["版本", "Version 2.35 — DRAFT (2022年9月)"],
    ["页数", "239 页"],
    ["定位", "德国 Common Criteria 认证体系中 AIS 20/31 的数学技术参考文档"],
    ["背景", "合并更新了 AIS 20 (确定性RNG) 和 AIS 31 (物理真随机数生成器)"],
  ];

  const rowH = 0.53;
  const tableY = 1.2;
  items.forEach((item, i) => {
    // Label card
    s.addShape(pres.shapes.RECTANGLE, {
      x: 0.6, y: tableY + i * (rowH + 0.08), w: 1.8, h: rowH,
      fill: { color: C.primary }, shadow: mkCardShadow(),
    });
    s.addText(item[0], {
      x: 0.6, y: tableY + i * (rowH + 0.08), w: 1.8, h: rowH,
      fontSize: 12, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0,
    });
    // Value
    s.addShape(pres.shapes.RECTANGLE, {
      x: 2.55, y: tableY + i * (rowH + 0.08), w: 6.8, h: rowH,
      fill: { color: C.white }, shadow: mkCardShadow(),
    });
    s.addText(item[1], {
      x: 2.7, y: tableY + i * (rowH + 0.08), w: 6.5, h: rowH,
      fontSize: 11, fontFace: "Arial", color: C.textDark, valign: "middle", margin: 0,
    });
  });

  // Footer
  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 3: Why RNG Security Evaluation?
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("为什么需要 RNG 安全评估？", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  // Left column: 4 icon cards
  const cards = [
    { icon: "01", title: "密码应用的基础", desc: "会话密钥、签名参数、Nonces、挑战值、盲化与掩码值 — 几乎所有密码操作都依赖随机数" },
    { icon: "02", title: "弱RNG的后果", desc: "弱随机数生成器可决定性地削弱整个密码应用的安全性，导致密钥可预测、签名可伪造" },
    { icon: "03", title: "20年评估传统", desc: "德国 Common Criteria 体系约 20 年来一直使用 AIS 20/31 作为 RNG 安全评估标准" },
    { icon: "04", title: "三方受众", desc: "本文档面向开发者（设计指导）、评估者（评估依据）和认证机构（认证标准）" },
  ];

  cards.forEach((card, i) => {
    const y = 1.2 + i * 1.05;
    // Number circle
    s.addShape(pres.shapes.OVAL, { x: 0.6, y: y + 0.05, w: 0.5, h: 0.5, fill: { color: C.accent } });
    s.addText(card.icon, { x: 0.6, y: y + 0.05, w: 0.5, h: 0.5, fontSize: 16, fontFace: "Arial", color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    // Card bg
    s.addShape(pres.shapes.RECTANGLE, { x: 1.3, y: y, w: 8.1, h: 0.85, fill: { color: C.white }, shadow: mkCardShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 1.3, y: y, w: 0.06, h: 0.85, fill: { color: C.accent } });
    s.addText(card.title, { x: 1.55, y: y + 0.05, w: 7.7, h: 0.3, fontSize: 14, fontFace: "Arial", color: C.primary, bold: true, margin: 0 });
    s.addText(card.desc, { x: 1.55, y: y + 0.38, w: 7.7, h: 0.42, fontSize: 10, fontFace: "Arial", color: C.textGray, margin: 0 });
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 4: RNG Three Categories
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("RNG 三大分类", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  const rngCards = [
    {
      color: "2B6CB0", title: "DRNG", subtitle: "确定性随机数生成器",
      items: ["将短随机种子\"扩展\"为长序列", "总熵 ≤ 种子熵", "计算安全性", "例: NIST SP800-90A DRBG"],
    },
    {
      color: "2C7A7B", title: "PTRNG", subtitle: "物理真随机数生成器",
      items: ["基于物理噪声源产生随机比特", "可精确估计熵值", "信息论安全性", "例: 环形振荡器、齐纳二极管"],
    },
    {
      color: "9B6B2C", title: "NPTRNG", subtitle: "非物理真随机数生成器",
      items: ["从非物理噪声源收集熵", "难以精确估计熵值", "跨平台运行", "例: Linux /dev/random"],
    },
  ];

  rngCards.forEach((card, i) => {
    const x = 0.4 + i * 3.15;
    // Card
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.2, w: 2.95, h: 3.8, fill: { color: C.white }, shadow: mkCardShadow() });
    // Top strip
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.2, w: 2.95, h: 0.06, fill: { color: card.color } });
    // Title
    s.addText(card.title, { x: x + 0.25, y: 1.45, w: 2.45, h: 0.45, fontSize: 22, fontFace: "Arial Black", color: card.color, bold: true, margin: 0 });
    s.addText(card.subtitle, { x: x + 0.25, y: 1.9, w: 2.45, h: 0.3, fontSize: 11, fontFace: "Arial", color: C.textGray, margin: 0 });
    // Divider
    s.addShape(pres.shapes.RECTANGLE, { x: x + 0.25, y: 2.3, w: 0.8, h: 0.025, fill: { color: card.color } });
    // Bullets
    const bulletText = card.items.map((item, idx) => ({
      text: item, options: { bullet: true, breakLine: idx < card.items.length - 1, fontSize: 11, color: C.textDark },
    }));
    s.addText(bulletText, { x: x + 0.25, y: 2.55, w: 2.5, h: 2.3, fontFace: "Arial", paraSpaceAfter: 8 });
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 5: Functionality Classes Overview (with Table)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("六大功能等级总览", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  const headerOpts = { fill: { color: C.tableHeader }, color: C.white, bold: true, fontSize: 12, fontFace: "Arial", align: "center", valign: "middle" };
  const cellOpts = (color, isCenter) => ({ fill: { color: color || C.white }, color: C.textDark, fontSize: 11, fontFace: "Arial", align: isCenter ? "center" : "left", valign: "middle" });

  const tableData = [
    [
      { text: "类别", options: headerOpts },
      { text: "等级", options: headerOpts },
      { text: "类型", options: headerOpts },
      { text: "安全强度", options: headerOpts },
      { text: "说明", options: headerOpts },
    ],
    [
      { text: "DRNG", options: cellOpts(C.white, true) },
      { text: "DRG.2", options: cellOpts(C.white, true) },
      { text: "确定性", options: cellOpts(C.white, true) },
      { text: "基础级", options: cellOpts(C.white, true) },
      { text: "适用于不需要新鲜熵保证的应用", options: cellOpts(C.white, false) },
    ],
    [
      { text: "DRNG", options: cellOpts(C.tableStripe, true) },
      { text: "DRG.3", options: cellOpts(C.tableStripe, true) },
      { text: "确定性", options: cellOpts(C.tableStripe, true) },
      { text: "增强级", options: cellOpts(C.tableStripe, true) },
      { text: "增加增强后向保密 + 种子树结构", options: cellOpts(C.tableStripe, false) },
    ],
    [
      { text: "DRNG", options: cellOpts(C.white, true) },
      { text: "DRG.4", options: cellOpts(C.white, true) },
      { text: "混合型", options: cellOpts(C.white, true) },
      { text: "最高级", options: cellOpts(C.white, true) },
      { text: "混合DRNG: 密码学后处理带记忆 + 持续熵刷新", options: cellOpts(C.white, false) },
    ],
    [
      { text: "PTRNG", options: cellOpts(C.tableStripe, true) },
      { text: "PTG.2", options: cellOpts(C.tableStripe, true) },
      { text: "物理真随机", options: cellOpts(C.tableStripe, true) },
      { text: "基础级", options: cellOpts(C.tableStripe, true) },
      { text: "随机模型 + 在线/总失效测试, 可选非密码学后处理", options: cellOpts(C.tableStripe, false) },
    ],
    [
      { text: "PTRNG", options: cellOpts(C.white, true) },
      { text: "PTG.3", options: cellOpts(C.white, true) },
      { text: "物理真随机(混合)", options: cellOpts(C.white, true) },
      { text: "最强级", options: cellOpts(C.white, true) },
      { text: "密码学后处理(带记忆), 整体安全性最强的等级", options: cellOpts(C.white, false) },
    ],
    [
      { text: "NPTRNG", options: cellOpts(C.tableStripe, true) },
      { text: "NTG.1", options: cellOpts(C.tableStripe, true) },
      { text: "非物理真随机", options: cellOpts(C.tableStripe, true) },
      { text: "—", options: cellOpts(C.tableStripe, true) },
      { text: "跨平台熵收集, 以 min-entropy 量化安全性", options: cellOpts(C.tableStripe, false) },
    ],
  ];

  s.addTable(tableData, {
    x: 0.4, y: 1.15, w: 9.2,
    colW: [1.1, 1.0, 1.5, 1.1, 4.5],
    border: { pt: 0.5, color: C.border },
    rowH: [0.5, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45],
    margin: [4, 8, 4, 8],
  });

  // Note
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 4.65, w: 9.2, h: 0.35, fill: { color: "FEF3E2" } });
  s.addText("注: DRG.1 和 PTG.1 已在 v2.35 中取消。DRG.4、PTG.3、NTG.1 均为混合型 RNG (Hybrid RNG)。", {
    x: 0.6, y: 4.65, w: 8.8, h: 0.35, fontSize: 10, fontFace: "Arial", color: C.warnOrange, valign: "middle", margin: 0,
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 6: DRG.2 Details
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("DRG.2 — 确定性RNG基础等级", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 26, fontFace: "Arial Black", color: C.white, bold: true });

  // Left: requirements
  s.addText("核心要求", { x: 0.6, y: 1.1, w: 4, h: 0.35, fontSize: 16, fontFace: "Arial", color: C.primary, bold: true });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.45, w: 0.8, h: 0.025, fill: { color: C.accent } });

  const reqs = [
    "种子由 TRNG 生成 (PTG.2/PTG.3/NTG.1)",
    "两次种子间最多 2⁴⁸ 次请求，单次 ≤ 2¹⁹ bits",
    "有效内部状态 ≥ 252 bits",
    "初始内部状态: min-entropy ≥ 240 bits\n  或 Shannon 熵 ≥ 250 bits",
    "前向保密 + 后向保密 (内部随机数粒度)",
    "状态转移函数 ϕ 或输出函数 ψ 是密码学的",
    "统计测试无法区分内部随机数与理想RNG输出",
  ];

  const reqBullets = reqs.map((r, i) => ({
    text: r, options: { bullet: true, breakLine: i < reqs.length - 1, fontSize: 11, color: C.textDark },
  }));
  s.addText(reqBullets, {
    x: 0.6, y: 1.6, w: 4.8, h: 3.3, fontFace: "Arial", paraSpaceAfter: 10,
  });

  // Right: visual card
  s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: 1.1, w: 3.8, h: 3.8, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y: 1.1, w: 3.8, h: 0.5, fill: { color: C.primary } });
  s.addText("DRG.2 适用场景", { x: 6, y: 1.15, w: 3.4, h: 0.4, fontSize: 13, fontFace: "Arial", color: C.white, bold: true, margin: 0 });

  s.addText([
    { text: "不需要保证新鲜熵的密码应用", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "一般安全级别的密钥生成", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "Nonces 与挑战值生成", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "不需要原子性请求保证", options: { bullet: true, fontSize: 11, color: C.textDark } },
  ], { x: 6.05, y: 1.8, w: 3.3, h: 2.8, fontFace: "Arial", paraSpaceAfter: 10 });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 7: DRG.3 Details
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("DRG.3 — 确定性RNG增强等级", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 26, fontFace: "Arial Black", color: C.white, bold: true });

  // Comparison: DRG.2 vs DRG.3
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 4.2, h: 0.4, fill: { color: "E2E8F0" } });
  s.addText("DRG.2 已有要求 (继承)", { x: 0.8, y: 1.15, w: 3.8, h: 0.4, fontSize: 11, fontFace: "Arial", color: C.textGray, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "种子安全要求 (DRG.2.1)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textGray } },
    { text: "请求限制 (DRG.2.2)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textGray } },
    { text: "内部状态大小 (DRG.2.3)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textGray } },
    { text: "初始熵要求 (DRG.2.4)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textGray } },
    { text: "前向/后向保密 (DRG.2.5/6)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textGray } },
    { text: "统计不可区分性 (DRG.2.9)", options: { bullet: true, fontSize: 10, color: C.textGray } },
  ], { x: 0.8, y: 1.55, w: 3.8, h: 2.7, fontFace: "Arial", paraSpaceAfter: 4 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.15, w: 4.4, h: 0.4, fill: { color: C.accent } });
  s.addText("DRG.3 新增/增强要求", { x: 5.4, y: 1.15, w: 4, h: 0.4, fontSize: 11, fontFace: "Arial", color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "增强后向保密 (请求粒度)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "ϕ 和 ψ 都必须是密码学的", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "ϕ 必须是单向函数 (One-Way)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "支持种子树结构 (高度限制)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "种子可由另一 DRNG 提供", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "根DRNG必须用TRNG种子", options: { bullet: true, fontSize: 10, color: C.textDark } },
  ], { x: 5.4, y: 1.55, w: 4, h: 2.7, fontFace: "Arial", paraSpaceAfter: 4 });

  // Bottom: seed tree explanation
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.3, w: 9, h: 0.7, fill: { color: C.lightAccent } });
  s.addText("种子树结构: 根 DRNG (TRNG种子) → 子 DRNG → 孙 DRNG ... 支持逐级种子传递，减轻认证负担", {
    x: 0.8, y: 4.3, w: 8.6, h: 0.7, fontSize: 11, fontFace: "Arial", color: C.primary, valign: "middle", margin: 0,
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 8: DRG.4 Details
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("DRG.4 — 混合型确定性RNG (最高级)", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 26, fontFace: "Arial Black", color: C.white, bold: true });

  // Hybrid concept
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 8.8, h: 0.8, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 0.06, h: 0.8, fill: { color: C.accent } });
  s.addText("DRG.4 是混合型 RNG (Hybrid RNG)：结合了 DRNG 的计算安全性和 TRNG 的信息论安全性，通过持续熵输入刷新内部状态。", {
    x: 0.9, y: 1.15, w: 8.3, h: 0.8, fontSize: 12, fontFace: "Arial", color: C.primary, valign: "middle", margin: 0,
  });

  // Three key features
  const features = [
    { title: "密码学后处理", desc: "必须使用带记忆的密码学后处理算法，确保即使部分内部状态泄露也难以回溯", icon: "🔐" },
    { title: "持续熵刷新", desc: "输出由熵输入持续刷新，防止内部状态因长时间运行而熵耗尽", icon: "🔄" },
    { title: "全DRG.3基础", desc: "继承DRG.3全部要求（增强后向保密、单向函数状态转移、种子树等）", icon: "📋" },
  ];

  features.forEach((f, i) => {
    const y = 2.2;
    const x = 0.6 + i * 3.05;
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 2.85, h: 2.1, fill: { color: C.white }, shadow: mkCardShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 2.85, h: 0.06, fill: { color: C.accent } });
    s.addText(f.icon, { x: x + 0.2, y: y + 0.25, w: 0.5, h: 0.5, fontSize: 24, align: "center", valign: "middle", margin: 0 });
    s.addText(f.title, { x: x + 0.2, y: y + 0.8, w: 2.45, h: 0.3, fontSize: 13, fontFace: "Arial", color: C.primary, bold: true, margin: 0 });
    s.addText(f.desc, { x: x + 0.2, y: y + 1.15, w: 2.45, h: 0.8, fontSize: 10, fontFace: "Arial", color: C.textGray, margin: 0 });
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 9: PTG.2 Details
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("PTG.2 — 物理真随机数生成器基础等级", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 26, fontFace: "Arial Black", color: C.white, bold: true });

  // Key requirements as visual blocks
  const reqBlocks = [
    { label: "随机模型", desc: "必须有正式的随机模型描述噪声源的随机特性，是PTRNG评估的核心", color: "2B6CB0" },
    { label: "平稳性", desc: "原始随机数满足时间局部平稳性 (time-local stationarity)", color: "2C7A7B" },
    { label: "在线测试", desc: "有效的在线测试和总失效测试，实时监控噪声源质量", color: "9B6C2C" },
    { label: "熵要求", desc: "Shannon熵每比特 > 0.9998\n或 min-entropy 每比特 > 0.98", color: "7B3A6E" },
    { label: "后处理", desc: "可选用非密码学后处理 (如 Von Neumann 去偏、固定压缩率)", color: "3A6B7B" },
  ];

  reqBlocks.forEach((b, i) => {
    const y = 1.15 + i * 0.8;
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y, w: 0.06, h: 0.65, fill: { color: b.color } });
    s.addText(b.label, { x: 0.9, y: y, w: 1.8, h: 0.3, fontSize: 11, fontFace: "Arial", color: b.color, bold: true, margin: 0 });
    s.addText(b.desc, { x: 0.9, y: y + 0.28, w: 4.5, h: 0.4, fontSize: 10, fontFace: "Arial", color: C.textGray, margin: 0 });
  });

  // Right side: visual diagram concept
  s.addShape(pres.shapes.RECTANGLE, { x: 5.5, y: 1.15, w: 4.1, h: 3.9, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addText("PTG.2 数据流", { x: 5.7, y: 1.25, w: 3.7, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.primary, bold: true, margin: 0 });
  // Simple boxes showing flow
  const flow = ["物理噪声源", "数字化", "在线测试", "后处理(可选)", "输出"];
  flow.forEach((step, i) => {
    const fy = 1.75 + i * 0.65;
    s.addShape(pres.shapes.RECTANGLE, { x: 6.3, y: fy, w: 2.2, h: 0.45, fill: { color: i === 0 ? "2B6CB0" : i === 4 ? C.accent : C.lightAccent } });
    s.addText(step, { x: 6.3, y: fy, w: 2.2, h: 0.45, fontSize: 10, fontFace: "Arial", color: i < 2 ? C.white : C.textDark, align: "center", valign: "middle", margin: 0 });
    if (i < 4) {
      s.addText("↓", { x: 7.2, y: fy + 0.45, w: 0.4, h: 0.2, fontSize: 10, align: "center", color: C.textGray });
    }
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 10: PTG.3 Details
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("PTG.3 — 最高安全等级 (混合型PTRNG)", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 26, fontFace: "Arial Black", color: C.white, bold: true });

  // Central highlight
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.1, w: 8.8, h: 0.6, fill: { color: "EBF5FB" } });
  s.addText("PTG.3 是整体安全性最强的功能等级，适用于最高安全需求的密码应用（如长期密钥生成、高安全签名等）", {
    x: 0.8, y: 1.1, w: 8.4, h: 0.6, fontSize: 12, fontFace: "Arial", color: C.primary, bold: true, valign: "middle", margin: 0,
  });

  // PTG.2 baseline + PTG.3 additions
  s.addText("PTG.2 基础要求", { x: 0.6, y: 1.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", color: "2C7A7B", bold: true });
  s.addText([
    { text: "随机模型 (强制)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "时间局部平稳性", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "在线测试 + 总失效测试", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "启动测试", options: { bullet: true, fontSize: 10, color: C.textDark } },
  ], { x: 0.8, y: 2.3, w: 3.8, h: 2, fontFace: "Arial", paraSpaceAfter: 6 });

  s.addText("PTG.3 新增要求", { x: 5.2, y: 1.95, w: 4, h: 0.35, fontSize: 14, fontFace: "Arial", color: C.accent, bold: true });
  s.addText([
    { text: "密码学后处理 (强制, 带记忆)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "支持多种熵声明方式", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "可声明特定熵界 (更灵活)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "内部随机数增强要求", options: { bullet: true, fontSize: 10, color: C.textDark } },
  ], { x: 5.4, y: 2.3, w: 4, h: 2, fontFace: "Arial", paraSpaceAfter: 6 });

  // Bottom entropy options
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.35, w: 8.8, h: 0.7, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addText("PTG.3 熵声明选项: ① Shannon 熵每比特 > 0.9998  ② min-entropy 每比特 > 0.98  ③ PTRNG 特定熵界 (如基于随机模型计算的精确熵界)", {
    x: 0.8, y: 4.35, w: 8.4, h: 0.7, fontSize: 10, fontFace: "Arial", color: C.textDark, valign: "middle", margin: 0,
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 11: NTG.1 Details
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("NTG.1 — 非物理真随机数生成器", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 26, fontFace: "Arial Black", color: C.white, bold: true });

  // Two columns
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 4.2, h: 0.4, fill: { color: "9B6B2C" } });
  s.addText("NTG.1 特点", { x: 0.8, y: 1.15, w: 3.8, h: 0.4, fontSize: 13, fontFace: "Arial", color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "熵源不受设计者控制 (跨平台)", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "难以精确估计熵值", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "使用 min-entropy 量化每比特熵下界", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "同样需要在线测试 + 总失效测试", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "典型案例: Linux /dev/random", options: { bullet: true, fontSize: 11, color: C.textDark } },
  ], { x: 0.8, y: 1.65, w: 3.8, h: 2.5, fontFace: "Arial", paraSpaceAfter: 8 });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.15, w: 4.4, h: 0.4, fill: { color: C.primary } });
  s.addText("与其他等级的对比", { x: 5.4, y: 1.15, w: 4, h: 0.4, fontSize: 13, fontFace: "Arial", color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "vs PTG.2/3: 不需要随机模型 (因为无法建模)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "vs PTG.2/3: 熵估计更保守 (只使用 min-entropy)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "vs DRG.2/3: 提供真实熵而非计算安全性", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "可作为 DRG 的熵种子源", options: { bullet: true, fontSize: 10, color: C.textDark } },
  ], { x: 5.4, y: 1.65, w: 4, h: 2.5, fontFace: "Arial", paraSpaceAfter: 8 });

  // Linux box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.25, w: 9, h: 0.8, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addText("Linux RNG 架构 (/dev/random & /dev/urandom): BSI委托的长期研究分析了Linux内核RNG，验证了其与NTG.1和DRG.3的一致性。内核版本5.6+采用了新的NPTRNG功能设计。", {
    x: 0.8, y: 4.25, w: 8.6, h: 0.8, fontSize: 10, fontFace: "Arial", color: C.textDark, valign: "middle", margin: 0,
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 12: Stochastic Model
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("随机模型 (Stochastic Model)", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  s.addText("TRNG 评估的核心概念 — 用数学方式描述物理噪声源的随机特性", {
    x: 0.6, y: 1.05, w: 8.8, h: 0.35, fontSize: 13, fontFace: "Arial", color: C.accent, bold: true, margin: 0,
  });

  // Four components in a 2x2 grid
  const comps = [
    { title: "随机模型定义", desc: "描述噪声源输出的概率分布模型。例: 抛硬币模型 B(1, p) — 偏向概率为p的独立同分布比特序列", color: "2B6CB0" },
    { title: "在线测试 (Online Test)", desc: "实时监控噪声源是否偏离模型预期。目标: 好区域内低误报率，坏区域内高检测率", color: "2C7A7B" },
    { title: "总失效测试 (Total Failure)", desc: "检测噪声源的完全失效 (如硬件故障)。比在线测试更严格，覆盖更广泛的失效模式", color: "9B6C2C" },
    { title: "启动测试 (Start-up Test)", desc: "设备上电或复位后执行，在输出随机数之前验证噪声源工作正常", color: "7B3A6E" },
  ];

  comps.forEach((c, i) => {
    const x = 0.6 + (i % 2) * 4.6;
    const y = 1.6 + Math.floor(i / 2) * 1.85;
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 4.3, h: 1.6, fill: { color: C.white }, shadow: mkCardShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.06, h: 1.6, fill: { color: c.color } });
    s.addText(c.title, { x: x + 0.3, y: y + 0.1, w: 3.7, h: 0.3, fontSize: 13, fontFace: "Arial", color: c.color, bold: true, margin: 0 });
    s.addText(c.desc, { x: x + 0.3, y: y + 0.45, w: 3.7, h: 1.0, fontSize: 10, fontFace: "Arial", color: C.textGray, margin: 0 });
  });

  s.addText("对于 PTG.2 和 PTG.3，随机模型是强制要求 (mandatory)。NTG.1 不需要随机模型。", {
    x: 0.6, y: 5.2, w: 8.8, h: 0.3, fontSize: 9, fontFace: "Arial", color: C.textGray, italic: true,
  });
}

// ============================================================
// SLIDE 13: Online Test & Total Failure Test
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("在线测试与总失效测试原理", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  // Online test detailed
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.15, w: 4.2, h: 0.4, fill: { color: "2C7A7B" } });
  s.addText("在线测试 (Online Test)", { x: 0.8, y: 1.15, w: 3.8, h: 0.4, fontSize: 13, fontFace: "Arial", color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "持续检测噪声源输出质量", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "检测统计特性是否偏离随机模型", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "异常时触发噪声报警 (noise alarm)", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "不允许输出未通过测试的随机数", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "关键指标: 低误报率, 高检测率", options: { bullet: true, fontSize: 11, color: C.textDark } },
  ], { x: 0.8, y: 1.65, w: 3.8, h: 2.3, fontFace: "Arial", paraSpaceAfter: 8 });

  // Total failure test
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.15, w: 4.4, h: 0.4, fill: { color: "9B6C2C" } });
  s.addText("总失效测试 & 启动测试", { x: 5.4, y: 1.15, w: 4, h: 0.4, fontSize: 13, fontFace: "Arial", color: C.white, bold: true, valign: "middle", margin: 0 });

  s.addText([
    { text: "总失效测试: 检测噪声源完全失效", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "比在线测试更严格", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "启动测试: 上电/复位后立即执行", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "确保初始状态下噪声源正常工作", options: { bullet: true, breakLine: true, fontSize: 11, color: C.textDark } },
    { text: "三者协同: 全生命周期保障", options: { bullet: true, fontSize: 11, color: C.textDark } },
  ], { x: 5.4, y: 1.65, w: 4, h: 2.3, fontFace: "Arial", paraSpaceAfter: 8 });

  // Test regions explanation
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.1, w: 9, h: 0.9, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addText([
    { text: "测试空间划分:  ", options: { fontSize: 10, color: C.textGray } },
    { text: "A_good (好区域)", options: { fontSize: 10, color: C.goodGreen, bold: true } },
    { text: " — 噪声源参数在此区域内应通过测试    ", options: { fontSize: 10, color: C.textGray } },
    { text: "A_bad (坏区域)", options: { fontSize: 10, color: "E53E3E", bold: true } },
    { text: " — 噪声源参数在此区域内应触发报警    ", options: { fontSize: 10, color: C.textGray } },
    { text: "A_real (实际区域)", options: { fontSize: 10, color: C.accent, bold: true } },
    { text: " — 噪声源参数的真实分布区域，应完全包含在 A_good 内", options: { fontSize: 10, color: C.textGray } },
  ], { x: 0.8, y: 4.1, w: 8.6, h: 0.9, fontFace: "Arial", valign: "middle", margin: 0 });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 14: Mathematical Foundations - Entropy
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("数学基础 — 熵与工作量因子", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  // Three entropy types as cards
  const entropyCards = [
    {
      title: "Min-Entropy (最小熵)",
      formula: "H_min = -log₂(max(pᵢ))",
      desc: "最保守的熵度量，只考虑最可能出现的值。评估中最常用的指标。",
      color: "2B6CB0",
    },
    {
      title: "Shannon 熵",
      formula: "H = -Σ pᵢ · log₂(pᵢ)",
      desc: "信息论标准度量，反映平均不确定性。需要平稳分布假设。",
      color: "2C7A7B",
    },
    {
      title: "Collision 熵",
      formula: "H_coll = -log₂(Σ pᵢ²)",
      desc: "基于碰撞概率的熵度量。用于某些特定类型的统计测试。",
      color: "9B6C2C",
    },
  ];

  entropyCards.forEach((ec, i) => {
    const x = 0.4 + i * 3.15;
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.15, w: 2.95, h: 2.4, fill: { color: C.white }, shadow: mkCardShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: 1.15, w: 2.95, h: 0.06, fill: { color: ec.color } });
    s.addText(ec.title, { x: x + 0.2, y: 1.35, w: 2.55, h: 0.35, fontSize: 14, fontFace: "Arial", color: ec.color, bold: true, margin: 0 });
    s.addText(ec.formula, { x: x + 0.2, y: 1.75, w: 2.55, h: 0.35, fontSize: 10, fontFace: "Consolas", color: C.textDark, margin: 0 });
    s.addText(ec.desc, { x: x + 0.2, y: 2.2, w: 2.55, h: 1.1, fontSize: 10, fontFace: "Arial", color: C.textGray, margin: 0 });
  });

  // Guess work + random mappings
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 3.8, w: 4.4, h: 1.2, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addText("Guess Work (猜测工作量)", { x: 0.6, y: 3.85, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.primary, bold: true, margin: 0 });
  s.addText("攻击者以最优策略猜测随机值所需的期望工作量。对给定的成功概率 λ，工作因子 w_λ 量化了猜测难度。min-entropy 与 w_1/2 直接相关。", {
    x: 0.6, y: 4.2, w: 4, h: 0.7, fontSize: 10, fontFace: "Arial", color: C.textGray, margin: 0,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 3.8, w: 4.4, h: 1.2, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addText("关键结论", { x: 5.4, y: 3.85, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.accent, bold: true, margin: 0 });
  s.addText("min-entropy 是最保守但也最可靠的熵度量。在安全认证中，当无法精确建模随机源时，推荐使用 min-entropy 作为熵评估依据。", {
    x: 5.4, y: 4.2, w: 4, h: 0.7, fontSize: 10, fontFace: "Arial", color: C.textGray, margin: 0,
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 15: NIST SP800-90A Conformity
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("NIST SP800-90A 一致性分析", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  s.addText("文档第5.3节分析了 NIST 批准的 DRBG 设计与 DRG.3 / DRG.4 的一致性关系", {
    x: 0.6, y: 1.05, w: 8.8, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.accent, bold: true, margin: 0,
  });

  // NIST DRBG functional model
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.6, w: 4.4, h: 2.5, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addText("NIST SP800-90A DRBG 功能模型", { x: 0.8, y: 1.7, w: 4, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "熵源 (Entropy Source)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "种子材料 → DRBG 内部状态", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "状态转移函数 (单向函数)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "输出函数 → 随机比特", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "支持重新播种 (Reseeding)", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "可选附加输入 (Additional Input)", options: { bullet: true, fontSize: 10, color: C.textDark } },
  ], { x: 0.8, y: 2.1, w: 4, h: 1.8, fontFace: "Arial", paraSpaceAfter: 6 });

  // Hash DRBG analysis
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.6, w: 4.2, h: 2.5, fill: { color: C.white }, shadow: mkCardShadow() });
  s.addText("Hash DRBG 安全评估", { x: 5.6, y: 1.7, w: 3.8, h: 0.3, fontSize: 12, fontFace: "Arial", color: C.primary, bold: true, margin: 0 });
  s.addText([
    { text: "详细分析了 Hash DRBG 的安全性", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "验证了状态转移函数的单向性", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "确认满足前后向保密要求", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "统计测试不可区分性验证", options: { bullet: true, breakLine: true, fontSize: 10, color: C.textDark } },
    { text: "结论: 符合条件的 Hash DRBG\n  可达到 DRG.3 或 DRG.4", options: { bullet: true, fontSize: 10, color: C.accent, bold: true } },
  ], { x: 5.6, y: 2.1, w: 3.8, h: 1.8, fontFace: "Arial", paraSpaceAfter: 6 });

  // Benefit box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.35, w: 9, h: 0.65, fill: { color: "EBF5FB" } });
  s.addText("对开发者的价值: 可直接引用第5.3节的分析结论作为安全证据，无需自行证明 NIST DRBG 设计的安全性，大幅减轻认证负担。", {
    x: 0.8, y: 4.35, w: 8.6, h: 0.65, fontSize: 11, fontFace: "Arial", color: C.primary, valign: "middle", margin: 0,
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 16: Linux /dev/random Analysis
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("Linux /dev/random 与 /dev/urandom 分析", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 26, fontFace: "Arial Black", color: C.white, bold: true });

  s.addText("BSI 委托的长期研究 (第5.6节)，验证 Linux 内核 RNG 与 AIS 31 功能等级的一致性", {
    x: 0.6, y: 1.05, w: 8.8, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.accent, bold: true, margin: 0,
  });

  // Architecture
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.6, w: 4.4, h: 0.35, fill: { color: C.primary } });
  s.addText("Linux NPTRNG 架构 (内核 5.6+)", { x: 0.8, y: 1.6, w: 4, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, valign: "middle", margin: 0 });

  const archItems = [
    "多种熵源: 键盘/鼠标中断、磁盘I/O、网络中断、RDRAND等",
    "熵池 (Entropy Pool) 收集与混合",
    "CRNG (ChaCha20-based RNG) 作为DRNG",
    "/dev/random: 阻塞式，保证新鲜熵",
    "/dev/urandom: 非阻塞，不保证新鲜熵",
  ];
  s.addText(archItems.map((a, i) => ({
    text: a, options: { bullet: true, breakLine: i < archItems.length - 1, fontSize: 10, color: C.textDark },
  })), { x: 0.8, y: 2.05, w: 4, h: 2.2, fontFace: "Arial", paraSpaceAfter: 6 });

  // Conformity table
  s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.6, w: 4.2, h: 0.35, fill: { color: C.primary } });
  s.addText("一致性对照", { x: 5.6, y: 1.6, w: 3.8, h: 0.35, fontSize: 12, fontFace: "Arial", color: C.white, bold: true, valign: "middle", margin: 0 });

  const confHeader = { fill: { color: C.primary }, color: C.white, bold: true, fontSize: 10, fontFace: "Arial", align: "center", valign: "middle" };
  const confCell = (color) => ({ fill: { color: color || C.white }, color: C.textDark, fontSize: 10, fontFace: "Arial", align: "center", valign: "middle" });

  s.addTable([
    [{ text: "组件", options: confHeader }, { text: "等级", options: confHeader }, { text: "状态", options: confHeader }],
    [{ text: "Entropy Source", options: confCell() }, { text: "NTG.1", options: confCell() }, { text: "合规", options: confCell() }],
    [{ text: "CRNG (DRNG)", options: confCell(C.tableStripe) }, { text: "DRG.3", options: confCell(C.tableStripe) }, { text: "待确认", options: confCell(C.tableStripe) }],
    [{ text: "/dev/random", options: confCell() }, { text: "NTG.1", options: confCell() }, { text: "合规", options: confCell() }],
    [{ text: "/dev/urandom", options: confCell(C.tableStripe) }, { text: "DRG.3", options: confCell(C.tableStripe) }, { text: "合规", options: confCell(C.tableStripe) }],
  ], {
    x: 5.4, y: 2.05, w: 4.2, colW: [1.6, 1.3, 1.3],
    border: { pt: 0.5, color: C.border },
    rowH: [0.35, 0.3, 0.3, 0.3, 0.3],
    margin: [4, 6, 4, 6],
  });

  // Note
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.5, w: 9, h: 0.45, fill: { color: C.lightAccent } });
  s.addText("开发者可引用 BSI 的 Linux RNG 研究报告作为其产品中 Linux RNG 组件的安全证据。", {
    x: 0.8, y: 4.5, w: 8.6, h: 0.45, fontSize: 11, fontFace: "Arial", color: C.primary, valign: "middle", margin: 0,
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 17: Practical Examples
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.offWhite };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.primary } });
  s.addText("实际应用示例", { x: 0.6, y: 0.15, w: 8.8, h: 0.6, fontSize: 28, fontFace: "Arial Black", color: C.white, bold: true });

  // 2x3 grid of examples
  const examples = [
    { title: "双噪声二极管 PTRNG", desc: "基于两个齐纳二极管的噪声电压差，Gamma分布建模时间间隔，经典的物理真随机设计方案" },
    { title: "放射性衰变 PTRNG", desc: "利用放射性同位素衰变事件的随机时间间隔，物理随机性极好但实现复杂" },
    { title: "PLL 锁相环噪声源", desc: "基于锁相环的抖动噪声，数字化后输出随机比特，适合集成电路实现" },
    { title: "Von Neumann 去偏", desc: "经典去偏算法: 将比特对 (0,1)→0, (1,0)→1, (0,0)和(1,1)→丢弃。消除偏向但降低吞吐率" },
    { title: "AES-OFB 模式 DRNG", desc: "使用AES在输出反馈模式下生成伪随机数，需验证状态转换函数的单向性" },
    { title: "在线测试程序示例", desc: "卡方拟合优度检验、Coron熵测试、多层级测试流程 (ER1→ER2→ER3) 构成完整在线测试体系" },
  ];

  examples.forEach((ex, i) => {
    const x = 0.4 + (i % 3) * 3.15;
    const y = 1.15 + Math.floor(i / 3) * 2.0;
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 2.95, h: 1.8, fill: { color: C.white }, shadow: mkCardShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: x, y: y, w: 0.06, h: 1.8, fill: { color: C.accent } });
    s.addText(ex.title, { x: x + 0.2, y: y + 0.1, w: 2.55, h: 0.45, fontSize: 12, fontFace: "Arial", color: C.primary, bold: true, margin: 0 });
    s.addText(ex.desc, { x: x + 0.2, y: y + 0.6, w: 2.55, h: 1.1, fontSize: 9, fontFace: "Arial", color: C.textGray, margin: 0 });
  });

  s.addText("AIS 31 随机数生成器功能等级标准", { x: 0.6, y: 5.25, w: 8.8, h: 0.3, fontSize: 8, color: C.textGray });
}

// ============================================================
// SLIDE 18: Summary (Dark background like title)
// ============================================================
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };

  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.04, fill: { color: C.accent } });

  s.addText("总结", {
    x: 0.6, y: 0.3, w: 8.8, h: 0.6,
    fontSize: 32, fontFace: "Arial Black", color: C.white, bold: true,
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 0.95, w: 1.5, h: 0.03, fill: { color: C.accent } });

  const summaryItems = [
    "AIS 31 是德国 BSI 发布的 RNG 安全评估权威标准",
    "定义了三类 RNG：DRNG (确定性)、PTRNG (物理真随机)、NPTRNG (非物理真随机)",
    "六个功能等级：DRG.2 → DRG.3 → DRG.4 → PTG.2 → PTG.3 (最强) → NTG.1",
    "核心技术支柱：随机模型 (Stochastic Model)、熵评估、在线测试 & 总失效测试",
    "与 NIST SP800-90A 等国际标准互补，提供标准间一致性分析",
    "为 Common Criteria 认证中 RNG 评估提供完整技术依据",
  ];

  const sumBullets = summaryItems.map((item, i) => ({
    text: item, options: { bullet: true, breakLine: i < summaryItems.length - 1, fontSize: 13, color: C.white },
  }));
  s.addText(sumBullets, {
    x: 0.6, y: 1.25, w: 8.8, h: 3.0, fontFace: "Arial", paraSpaceAfter: 14,
  });

  // Bottom decorative
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 4.35, w: 8.8, h: 0.025, fill: { color: C.accent, transparency: 50 } });
  s.addText("PTG.3 是整体安全性最强的功能等级", {
    x: 0.6, y: 4.5, w: 8.8, h: 0.4, fontSize: 18, fontFace: "Arial", color: C.accent, bold: true, align: "center",
  });
  s.addText("Matthias Peter & Werner Schindler — BSI — Version 2.35 DRAFT (September 2022)", {
    x: 0.6, y: 5.1, w: 8.8, h: 0.3, fontSize: 9, fontFace: "Arial", color: "8899AA", align: "center",
  });
}

// ============================================================
// Save
// ============================================================
pres.writeFile({ fileName: "D:\\ClaudeCodeTest\\test5\\AIS31_教学讲解.pptx" })
  .then(() => console.log("PPT saved successfully!"))
  .catch(err => console.error(err));
