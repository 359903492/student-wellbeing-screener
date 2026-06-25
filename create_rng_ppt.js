const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Claude Code";
pres.title = "随机数生成器（RNG）全面解析";

// ============================================================
// Color Palette
// ============================================================
const C = {
  darkBg:    "0B1120",
  darkBg2:   "111D32",
  accent:    "06B6D4",
  accent2:   "0891B2",
  accent3:   "22D3EE",
  white:     "FFFFFF",
  lightBg:   "F0F9FF",
  lightBg2:  "E0F2FE",
  textDark:  "1E293B",
  textMuted: "64748B",
  cardBg:    "FFFFFF",
  green:     "10B981",
  orange:    "F59E0B",
  purple:    "8B5CF6",
  red:       "EF4444",
  border:    "E2E8F0",
};

const FONT = "Microsoft YaHei";
const MONO = "Consolas";

const makeShadow = () => ({ type: "outer", color: "000000", blur: 8, offset: 3, angle: 135, opacity: 0.12 });

// ============================================================
// SLIDE 1 — Title
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.accent } });
  slide.addShape(pres.shapes.OVAL, { x: 8.2, y: -0.8, w: 3, h: 3, fill: { color: C.accent2, transparency: 88 } });
  slide.addShape(pres.shapes.OVAL, { x: 7.5, y: -1.2, w: 2, h: 2, fill: { color: C.accent, transparency: 85 } });
  slide.addShape(pres.shapes.RECTANGLE, { x: 1.5, y: 5.4, w: 7, h: 0.04, fill: { color: C.accent, transparency: 40 } });

  slide.addText("随机数生成器", {
    x: 1.5, y: 1.0, w: 7, h: 1.0,
    fontSize: 48, fontFace: FONT, color: C.white, bold: true, margin: 0,
  });
  slide.addText("RNG 全面解析 — 硬件设计与算法", {
    x: 1.5, y: 1.9, w: 7, h: 0.8,
    fontSize: 30, fontFace: FONT, color: C.accent, bold: false, margin: 0,
  });
  slide.addText("TRNG 电路架构 · PRNG 算法对比 · CSPRNG 标准 · 后处理技术 · 2025 前沿趋势", {
    x: 1.5, y: 3.4, w: 7, h: 0.6,
    fontSize: 14, fontFace: FONT, color: C.textMuted, margin: 0,
  });
  slide.addText("2025年6月", {
    x: 1.5, y: 4.8, w: 3, h: 0.4,
    fontSize: 12, fontFace: FONT, color: C.textMuted, margin: 0,
  });
}

// ============================================================
// SLIDE 2 — 三大类别
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent } });

  slide.addText("RNG 三大类别", {
    x: 0.6, y: 0.35, w: 8, h: 0.65,
    fontSize: 30, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText("随机数生成器的分类与核心差异", {
    x: 0.6, y: 0.95, w: 8, h: 0.4,
    fontSize: 13, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  const cards = [
    {
      label: "TRNG", title: "真随机数生成器",
      desc: "基于物理过程的\n本质不确定性",
      features: ["环形振荡器抖动 / 亚稳态", "放射性衰变 / 量子光学", "不可预测、不可复现", "需专用硬件电路"],
      accent: C.green, x: 0.4,
    },
    {
      label: "PRNG", title: "伪随机数生成器",
      desc: "确定性算法 +\n种子值生成序列",
      features: ["LCG / MT19937 / Xorshift", "PCG / LFSR / 混沌映射", "可复现、速度极快", "非密码学用途"],
      accent: C.accent, x: 3.5,
    },
    {
      label: "CSPRNG", title: "密码学安全 PRNG",
      desc: "PRNG + 密码学\n安全属性",
      features: ["ChaCha20 / AES-CTR DRBG", "BLAKE3 / Fortuna / SHA-2", "前向+后向安全性", "NIST SP 800-90A 标准"],
      accent: C.purple, x: 6.6,
    },
  ];

  cards.forEach((c) => {
    slide.addShape(pres.shapes.RECTANGLE, { x: c.x, y: 1.7, w: 2.95, h: 3.5, fill: { color: C.cardBg }, shadow: makeShadow() });
    slide.addShape(pres.shapes.RECTANGLE, { x: c.x, y: 1.7, w: 2.95, h: 0.06, fill: { color: c.accent } });
    slide.addText(c.label, { x: c.x + 0.25, y: 2.0, w: 2.45, h: 0.45, fontSize: 20, fontFace: MONO, color: c.accent, bold: true, margin: 0 });
    slide.addText(c.title, { x: c.x + 0.25, y: 2.5, w: 2.45, h: 0.4, fontSize: 15, fontFace: FONT, color: C.textDark, bold: true, margin: 0 });
    slide.addText(c.desc, { x: c.x + 0.25, y: 2.9, w: 2.45, h: 0.55, fontSize: 11, fontFace: FONT, color: C.textMuted, margin: 0 });
    slide.addText(
      c.features.map((f, i) => ({
        text: f, options: { bullet: true, breakLine: i < c.features.length - 1, fontSize: 10, fontFace: FONT, color: C.textDark },
      })),
      { x: c.x + 0.25, y: 3.5, w: 2.5, h: 1.6, paraSpaceAfter: 3, margin: 0 }
    );
  });
}

// ============================================================
// SLIDE 3 — TRNG 硬件架构（一）：环形振荡器与抖动采样
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.green } });

  slide.addText("TRNG 硬件设计（一）：环形振荡器与抖动采样", {
    x: 0.6, y: 0.35, w: 9, h: 0.6,
    fontSize: 26, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText("利用半导体热噪声导致时钟边沿随机抖动 — 最成熟的 TRNG 架构", {
    x: 0.6, y: 0.9, w: 9, h: 0.35,
    fontSize: 12, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  // Principle
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.45, w: 4.5, h: 1.6, fill: { color: C.cardBg }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.45, w: 0.06, h: 1.6, fill: { color: C.green } });
  slide.addText("基本原理", {
    x: 0.65, y: 1.55, w: 4, h: 0.35,
    fontSize: 15, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText([
    { text: "奇数个反相器组成自由振荡环路 → 频率由门延迟决定", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "热噪声导致每个边沿到达时间随机偏移（Jitter）", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "慢速时钟通过 DFF 采样快速 RO，捕获抖动不确定性", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "经典方案（Sunar 2007）：多 RO + XOR 输出放大熵", options: { bullet: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
  ], { x: 0.65, y: 1.95, w: 4, h: 1.0, paraSpaceAfter: 4, margin: 0 });

  // Issues
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.45, w: 4.5, h: 1.6, fill: { color: "FEF2F2" }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 1.45, w: 0.06, h: 1.6, fill: { color: C.red } });
  slide.addText("局限性", {
    x: 5.45, y: 1.55, w: 4, h: 0.35,
    fontSize: 15, fontFace: FONT, color: C.red, bold: true, margin: 0,
  });
  slide.addText([
    { text: "需累积数百至数千周期抖动 → 吞吐率仅 ~1 Mbps", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "等长环路容易相位锁定 → 振荡器间产生相关性", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "对温度/电压变化敏感 → 需校准", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "手工布局布线限制 → FPGA 可移植性差", options: { bullet: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
  ], { x: 5.45, y: 1.95, w: 4, h: 1.0, paraSpaceAfter: 4, margin: 0 });

  // ADPLL-based
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 3.3, w: 9.3, h: 1.1, fill: { color: C.cardBg }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 3.3, w: 0.06, h: 1.1, fill: { color: C.accent } });
  slide.addText("ADPLL-TRNG（2025）— FIR 数字滤波器全数字锁相环架构", {
    x: 0.65, y: 3.35, w: 8.7, h: 0.35,
    fontSize: 14, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText([
    { text: "平台：Artix-7 (XC7A35T)  |  功耗：0.072 W  |  吞吐率：200 Mbps  |  后处理：XOR-corrector", options: { breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "创新点：用 Kaiser 窗 FIR 低通滤波器替代传统模拟环路滤波器，抖动频谱整形效果优于经典 PLL 方案", options: { fontSize: 11, fontFace: FONT, color: C.textMuted } },
  ], { x: 0.65, y: 3.7, w: 8.7, h: 0.65, margin: 0 });

  // Classic RO comparison table
  slide.addText("经典环形振荡器 TRNG 方案对比", {
    x: 0.6, y: 4.65, w: 6, h: 0.35,
    fontSize: 14, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });

  const roTable = [
    [
      { text: "方案", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 10, fontFace: FONT, align: "center" } },
      { text: "熵源", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 10, fontFace: FONT, align: "center" } },
      { text: "吞吐率", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 10, fontFace: FONT, align: "center" } },
      { text: "资源", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 10, fontFace: FONT, align: "center" } },
      { text: "特点", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 10, fontFace: FONT, align: "center" } },
    ],
    [
      { text: "Sunar RO (2007)", options: { fontSize: 10, fontFace: FONT, color: C.textDark, bold: true } },
      { text: "多 RO 抖动", options: { fontSize: 10, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "~1 Mbps", options: { fontSize: 10, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "中", options: { fontSize: 10, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "XOR 合并多 RO 输出，经典基线", options: { fontSize: 10, fontFace: FONT, color: C.textMuted } },
    ],
    [
      { text: "高速 RO + LFSR (2025)", options: { fontSize: 10, fontFace: FONT, color: C.textDark, bold: true } },
      { text: "DFF 亚稳态采样 + LFSR 白化", options: { fontSize: 10, fontFace: FONT, color: C.textDark } },
      { text: "300 Mbps", options: { fontSize: 10, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "108 LUTs", options: { fontSize: 10, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "8/16/24/32 RO 可配置，Artix-7", options: { fontSize: 10, fontFace: FONT, color: C.textMuted } },
    ],
    [
      { text: "ADPLL-TRNG (2025)", options: { fontSize: 10, fontFace: FONT, color: C.textDark, bold: true } },
      { text: "全数字 PLL 噪声", options: { fontSize: 10, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "200 Mbps", options: { fontSize: 10, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "Artix-7", options: { fontSize: 10, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "FIR 环路滤波，0.072W 超低功耗", options: { fontSize: 10, fontFace: FONT, color: C.textMuted } },
    ],
  ];

  slide.addTable(roTable, {
    x: 0.35, y: 5.0, w: 9.3,
    colW: [2.0, 2.0, 1.2, 1.0, 3.1],
    rowH: [0.3, 0.28, 0.28, 0.28],
    border: { pt: 0.5, color: C.border },
  });
}

// ============================================================
// SLIDE 4 — TRNG 硬件架构（二）：亚稳态架构
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.green } });

  slide.addText("TRNG 硬件设计（二）：亚稳态架构", {
    x: 0.6, y: 0.35, w: 9, h: 0.6,
    fontSize: 26, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText("强制触发器进入亚稳态，利用热噪声决定最终输出 — 新一代主流方法", {
    x: 0.6, y: 0.9, w: 9, h: 0.35,
    fontSize: 12, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  // Principle card
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.45, w: 9.3, h: 1.1, fill: { color: C.cardBg }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.45, w: 0.06, h: 1.1, fill: { color: C.green } });
  slide.addText("亚稳态原理", {
    x: 0.65, y: 1.5, w: 3, h: 0.35,
    fontSize: 15, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText([
    { text: "当 DFF 的 Setup/Hold 时间被违反 → 输出进入亚稳态（不稳定平衡态）→ 热噪声决定最终稳定到 0 还是 1", options: { breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "关键挑战：自然亚稳态事件稀少。解决方案 → 主动构造电路迫使每次操作都产生亚稳态（故意违反对时序约束）", options: { breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "优势：仅需几个振荡周期即可产生随机位（vs. 抖动方案需数百/数千周期），大幅提升吞吐率", options: { fontSize: 11, fontFace: FONT, color: C.accent2 } },
  ], { x: 0.65, y: 1.85, w: 8.7, h: 0.7, margin: 0 });

  // Two classic architectures side by side
  // Meta-RO
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 2.8, w: 4.5, h: 2.4, fill: { color: C.cardBg }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 2.8, w: 4.5, h: 0.06, fill: { color: C.accent } });
  slide.addText("Meta-RO（CHES 2008）", {
    x: 0.6, y: 2.95, w: 4, h: 0.35,
    fontSize: 14, fontFace: MONO, color: C.accent, bold: true, margin: 0,
  });
  slide.addText("Samsung — Vasyltsov, Kim & Karpinskyy", {
    x: 0.6, y: 3.28, w: 4, h: 0.25,
    fontSize: 9, fontFace: FONT, color: C.textMuted, margin: 0,
  });
  slide.addText([
    { text: "结构：环中每个反相器可断开+短路 → 收敛到亚稳态电压", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: FONT, color: C.textDark } },
    { text: "工作两阶段：① 亚稳态模式（各节点独立噪声源）→ ② 生成模式（重连为 RO）", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: FONT, color: C.textDark } },
    { text: "仅需几个振荡周期 vs. 抖动法需数百周期", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: FONT, color: C.textDark } },
    { text: "吞吐率：35–140 Mbps  |  65nm ASIC  |  ~1 μm²", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: FONT, color: C.textDark } },
    { text: "通过 AIS.31 / FIPS 140 认证", options: { bullet: true, fontSize: 10, fontFace: FONT, color: C.green } },
  ], { x: 0.6, y: 3.55, w: 4, h: 1.5, paraSpaceAfter: 2, margin: 0 });

  // TERO
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 2.8, w: 4.5, h: 2.4, fill: { color: C.cardBg }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.15, y: 2.8, w: 4.5, h: 0.06, fill: { color: C.purple } });
  slide.addText("TERO（CHES 2010）", {
    x: 5.4, y: 2.95, w: 4, h: 0.35,
    fontSize: 14, fontFace: MONO, color: C.purple, bold: true, margin: 0,
  });
  slide.addText("Varchola & Drutarovsky — 瞬态效应环形振荡器", {
    x: 5.4, y: 3.28, w: 4, h: 0.25,
    fontSize: 9, fontFace: FONT, color: C.textMuted, margin: 0,
  });
  slide.addText([
    { text: "结构：XOR-AND-XOR-AND 双稳环路，受控边沿激励产生瞬态振荡", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: FONT, color: C.textDark } },
    { text: "每次激励的振荡次数因固有噪声随机变化 → T-FF 判断奇偶 → 输出 1-bit", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: FONT, color: C.textDark } },
    { text: "创新：熵源是亚稳态解析的振荡轨迹，非仅最终稳态", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: FONT, color: C.textDark } },
    { text: "比传统 RO 对全局扰动敏感度更低", options: { bullet: true, breakLine: true, fontSize: 10, fontFace: FONT, color: C.textDark } },
    { text: "极紧凑：仅需 1 CLB（Spartan-3E）", options: { bullet: true, fontSize: 10, fontFace: FONT, color: C.green } },
  ], { x: 5.4, y: 3.55, w: 4, h: 1.5, paraSpaceAfter: 2, margin: 0 });
}

// ============================================================
// SLIDE 5 — TRNG 硬件架构（三）：混合熵源与新型架构
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.green } });

  slide.addText("TRNG 硬件设计（三）：混合熵源与新型架构", {
    x: 0.6, y: 0.3, w: 9, h: 0.55,
    fontSize: 26, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText("2024-2025 年顶级会议/期刊发表的最新技术方案 — 双熵源融合 + 无需后处理", {
    x: 0.6, y: 0.8, w: 9, h: 0.3,
    fontSize: 12, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  const designs = [
    {
      name: "SSRO-TRNG",
      venue: "Integration VLSI Journal, 2025",
      resources: "16 LUTs + 13 DFFs",
      throughput: "400 Mbps",
      accent: C.accent,
      desc: "状态可切换环形振荡器。FF 亚稳态控制 SSRO 在「振荡」和「缓冲」状态间随机切换。反馈策略消除不动点，增强振荡复杂度。原始输出直接通过 NIST/AIS-31，无需后处理。",
    },
    {
      name: "DS-TRNG",
      venue: "Microelectronics Journal, 2024",
      resources: "33 LUTs + 4 FFs",
      throughput: "500 Mbps",
      accent: C.purple,
      desc: "基于 DEMUX 的双路径切换 TRNG。分离「抖动累积」和「亚稳态发生」到两条并行路径，缩短操作周期。RO 控制路径切换时序 → 切换本身也引入不确定性。min-entropy：0.9955。",
    },
    {
      name: "MRO-TRNG",
      venue: "Microelectronics Journal, 2024",
      resources: "10 LUTs + 2 DFFs + 1 MUX",
      throughput: "300 Mbps",
      accent: C.orange,
      desc: "MUX 环形振荡器 TRNG。利用 MUX + 路径延迟差异动态转换 MRO 频率。仅需 10 LUTs + 2 DFFs，是已知最紧凑的双熵源设计。XOR 后处理。",
    },
    {
      name: "Galois-LFSR",
      venue: "Microelectronics Journal, 2025",
      resources: "17 LUTs + 9 DFFs",
      throughput: "300 Mbps",
      accent: C.green,
      desc: "基于 Galois LFSR + 动态反馈路径。反相器 + 2-to-1 MUX 动态改变 LFSR 反馈路径，融合抖动+亚稳态。通过 NIST SP800-90B 和 AIS-31。",
    },
    {
      name: "Multi-Cell",
      venue: "Integration VLSI Journal, 2025",
      resources: "41 LUTs + 10 DFFs",
      throughput: "433 Mbps",
      accent: C.red,
      desc: "多单元选择器时钟驱动 + XOR 反馈。多个 RO 产生故意频率失配的时钟/数据 → 强制 DFF 亚稳态。级联 XOR 反馈放大熵。通过 TestU01。",
    },
  ];

  designs.forEach((d, i) => {
    const y = 1.25 + i * 0.78;
    const bgColor = i % 2 === 0 ? C.cardBg : C.lightBg2;

    slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 9.3, h: 0.68, fill: { color: bgColor } });
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.35, y, w: 0.05, h: 0.68, fill: { color: d.accent } });

    slide.addText(d.name, {
      x: 0.65, y: y + 0.02, w: 1.6, h: 0.3,
      fontSize: 13, fontFace: MONO, color: d.accent, bold: true, margin: 0,
    });
    slide.addText(d.venue, {
      x: 0.65, y: y + 0.32, w: 1.8, h: 0.28,
      fontSize: 7.5, fontFace: FONT, color: C.textMuted, margin: 0,
    });
    slide.addText(d.throughput, {
      x: 2.3, y: y + 0.02, w: 1.2, h: 0.28,
      fontSize: 13, fontFace: MONO, color: C.textDark, bold: true, align: "center", margin: 0,
    });
    slide.addText(d.resources, {
      x: 2.3, y: y + 0.32, w: 1.2, h: 0.28,
      fontSize: 7.5, fontFace: FONT, color: C.textMuted, align: "center", margin: 0,
    });
    slide.addText(d.desc, {
      x: 3.7, y: y + 0.08, w: 5.7, h: 0.55,
      fontSize: 9.5, fontFace: FONT, color: C.textDark, margin: 0,
    });
  });
}

// ============================================================
// SLIDE 6 — 后处理技术与量子/混沌 TRNG
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.green } });

  slide.addText("TRNG 后处理技术 & 新型熵源", {
    x: 0.6, y: 0.3, w: 9, h: 0.55,
    fontSize: 26, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });

  // ====== Post-processing section ======
  slide.addText("后处理（Post-Processing）技术", {
    x: 0.6, y: 1.05, w: 5, h: 0.35,
    fontSize: 16, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });

  const ppTable = [
    [
      { text: "技术", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 9, fontFace: FONT, align: "center" } },
      { text: "操作方式", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 9, fontFace: FONT, align: "center" } },
      { text: "偏差消除", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 9, fontFace: FONT, align: "center" } },
      { text: "吞吐率代价", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 9, fontFace: FONT, align: "center" } },
      { text: "适用场景", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 9, fontFace: FONT, align: "center" } },
    ],
    [
      { text: "von Neumann", options: { fontSize: 9, fontFace: MONO, color: C.textDark, bold: true } },
      { text: "01→1, 10→0, 00/11→丢弃", options: { fontSize: 9, fontFace: FONT, color: C.textDark } },
      { text: "完全消除", options: { fontSize: 9, fontFace: FONT, color: C.green, bold: true, align: "center" } },
      { text: "≥75% 损失", options: { fontSize: 9, fontFace: FONT, color: C.red, align: "center" } },
      { text: "低速率容忍场景", options: { fontSize: 9, fontFace: FONT, color: C.textMuted } },
    ],
    [
      { text: "XOR 校正", options: { fontSize: 9, fontFace: MONO, color: C.textDark, bold: true } },
      { text: "Z = X₁⊕X₂⊕...⊕Xw", options: { fontSize: 9, fontFace: FONT, color: C.textDark } },
      { text: "偏差 b→2b²", options: { fontSize: 9, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "无损失（需多源）", options: { fontSize: 9, fontFace: FONT, color: C.green, align: "center" } },
      { text: "多熵源组合（最常用）", options: { fontSize: 9, fontFace: FONT, color: C.textMuted } },
    ],
    [
      { text: "LFSR 白化", options: { fontSize: 9, fontFace: MONO, color: C.textDark, bold: true } },
      { text: "线性反馈移位寄存器", options: { fontSize: 9, fontFace: FONT, color: C.textDark } },
      { text: "良好", options: { fontSize: 9, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "低延迟", options: { fontSize: 9, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "高速 TRNG + FPGA", options: { fontSize: 9, fontFace: FONT, color: C.textMuted } },
    ],
    [
      { text: "密码学哈希", options: { fontSize: 9, fontFace: MONO, color: C.textDark, bold: true } },
      { text: "SHA-256 / AES / PRESENT", options: { fontSize: 9, fontFace: FONT, color: C.textDark } },
      { text: "鲁棒消除", options: { fontSize: 9, fontFace: FONT, color: C.green, bold: true, align: "center" } },
      { text: "1-10 pJ/bit", options: { fontSize: 9, fontFace: FONT, color: C.textDark, align: "center" } },
      { text: "高安全要求场景", options: { fontSize: 9, fontFace: FONT, color: C.textMuted } },
    ],
  ];

  slide.addTable(ppTable, {
    x: 0.35, y: 1.45, w: 9.3,
    colW: [1.5, 2.3, 1.2, 1.5, 2.0],
    rowH: [0.28, 0.3, 0.3, 0.3, 0.3],
    border: { pt: 0.5, color: C.border },
  });

  // Trend note
  slide.addText("2025 趋势：顶级设计（SSRO、DS-TRNG、Multi-Cell）已实现「原始输出直接通过 NIST 测试，无需后处理」", {
    x: 0.35, y: 3.05, w: 9.3, h: 0.3,
    fontSize: 10, fontFace: FONT, color: C.accent2, italic: true, margin: 0,
  });

  // ====== New entropy sources section ======
  slide.addText("新型熵源：超越传统电子噪声", {
    x: 0.6, y: 3.5, w: 5, h: 0.35,
    fontSize: 16, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });

  const newSources = [
    {
      name: "量子光学 QRNG", accent: C.purple,
      desc: "光子通过分束器 → 透射/反射由量子力学决定，理论上绝对不可预测。ID Quantique 已商用。IBM/Google 积极布局，预计 2026 年速率达 100 Mbps+。",
    },
    {
      name: "忆阻器 RTN 噪声", accent: C.orange,
      desc: "TiOx/Al₂O₃ 忆阻器的随机电报噪声（RTN），捕获/发射时间随机。2025 年实现偏置无关设计（边沿检测+N位计数器），解决传统偏置敏感问题。",
    },
    {
      name: "模拟混沌 Tent Map", accent: C.red,
      desc: "3 晶体管 Tent Map 单元 × 4 环形拓扑，65nm CMOS 仅 160 μm²，19.68 μW，0.787 pJ/bit。源跟随器消除级间干扰，无需精确同步时钟。",
    },
    {
      name: "ABN 自治布尔网络", accent: C.accent,
      desc: "5 节点耦合布尔网络 + 不平衡 RO + 3 输入 XOR，Lyapunov 指数为正（混沌态），熵率 0.9963 bps。当前 FPGA TRNG 吞吐率记录：750 Mbps（2025）。",
    },
  ];

  newSources.forEach((s, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.35 + col * 4.7;
    const y = 3.95 + row * 0.72;

    slide.addText(s.name, {
      x, y, w: 2.0, h: 0.28,
      fontSize: 11, fontFace: MONO, color: s.accent, bold: true, margin: 0,
    });
    slide.addText(s.desc, {
      x: x + 2.1, y, w: 2.5, h: 0.65,
      fontSize: 8.5, fontFace: FONT, color: C.textDark, margin: 0,
    });
  });
}

// ============================================================
// SLIDE 7 — PRNG 算法对比
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent } });

  slide.addText("伪随机数生成器 (PRNG)", {
    x: 0.6, y: 0.35, w: 8, h: 0.65,
    fontSize: 30, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText("确定性算法 + 种子 → 看似随机的序列 | Xₙ₊₁ = f(Xₙ, seed)", {
    x: 0.6, y: 0.95, w: 8, h: 0.4,
    fontSize: 13, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  const tableHeader = [
    { text: "算法", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 11, fontFace: FONT, align: "center" } },
    { text: "周期", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 11, fontFace: FONT, align: "center" } },
    { text: "速度", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 11, fontFace: FONT, align: "center" } },
    { text: "统计质量", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 11, fontFace: FONT, align: "center" } },
    { text: "适用场景", options: { fill: { color: C.darkBg }, color: C.white, bold: true, fontSize: 11, fontFace: FONT, align: "center" } },
  ];

  const rows = [
    ["LCG\n线性同余法", "≈ 2³²", "★★★★★", "★★☆☆☆", "教学 / 简单应用"],
    ["Mersenne Twister\nMT19937", "2¹⁹⁹³⁷−1", "★★★★☆", "★★★★★", "游戏 / Python/R/MATLAB"],
    ["Xorshift / Xoroshiro", "2¹²⁸−1", "★★★★★", "★★★★☆", "高性能场景"],
    ["PCG", "2⁶⁴~2¹²⁸", "★★★★☆", "★★★★★", "现代通用 PRNG"],
    ["LFSR", "2ⁿ−1", "★★★★★", "★★★☆☆", "硬件实现"],
    ["混沌映射类\nLogistic/Tent", "可变", "★★★☆☆", "★★★★☆", "2024-2025 研究热点"],
  ];

  const tableData = [tableHeader];
  rows.forEach((row, i) => {
    const bgColor = i % 2 === 0 ? C.white : C.lightBg;
    tableData.push(row.map((cell, j) => ({
      text: cell,
      options: {
        fontSize: 10, fontFace: FONT, color: C.textDark,
        fill: { color: bgColor },
        bold: j === 0, align: j === 0 ? "left" : "center", valign: "middle",
      },
    })));
  });

  slide.addTable(tableData, {
    x: 0.4, y: 1.6, w: 9.2,
    colW: [1.8, 1.4, 1.2, 1.3, 2.2],
    rowH: [0.45, 0.54, 0.54, 0.54, 0.54, 0.54, 0.54],
    border: { pt: 0.5, color: C.border },
  });

  slide.addText("选择建议：大规模模拟 → MT19937 | 高性能 → Xorshift | 现代通用 → PCG | 密码学 → 请使用 CSPRNG！", {
    x: 0.4, y: 5.2, w: 9.2, h: 0.35,
    fontSize: 11, fontFace: FONT, color: C.accent2, bold: true, margin: 0,
  });
}

// ============================================================
// SLIDE 8 — CSPRNG
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.purple } });

  slide.addText("密码学安全 PRNG (CSPRNG)", {
    x: 0.6, y: 0.35, w: 8, h: 0.65,
    fontSize: 30, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.3, w: 4.4, h: 1.1, fill: { color: C.cardBg }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.3, w: 0.06, h: 1.1, fill: { color: C.purple } });
  slide.addText("前向安全性", {
    x: 0.7, y: 1.35, w: 3.8, h: 0.35,
    fontSize: 16, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText("即使攻击者知道当前状态，也无法推算之前生成的随机数", {
    x: 0.7, y: 1.7, w: 3.8, h: 0.55,
    fontSize: 11, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 4.4, h: 1.1, fill: { color: C.cardBg }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.3, w: 0.06, h: 1.1, fill: { color: C.purple } });
  slide.addText("后向安全性", {
    x: 5.5, y: 1.35, w: 3.8, h: 0.35,
    fontSize: 16, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText("即使攻击者知道当前状态，也无法预测未来的随机数输出", {
    x: 5.5, y: 1.7, w: 3.8, h: 0.55,
    fontSize: 11, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  slide.addText("主流 CSPRNG 算法", {
    x: 0.6, y: 2.7, w: 4, h: 0.45,
    fontSize: 18, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });

  const algos = [
    { name: "ChaCha20", desc: "Google 力推，性能极佳，移动端友好", accent: C.accent },
    { name: "AES-CTR DRBG", desc: "硬件 AES 加速支持，NIST 标准", accent: C.green },
    { name: "BLAKE3", desc: "2025 趋势，速度比 SHA-2 快 50%", accent: C.orange },
    { name: "SHA-2/3 DRBG", desc: "通用性强，广泛验证", accent: C.purple },
    { name: "Fortuna", desc: "多熵源积累，操作系统广泛采用", accent: C.accent2 },
  ];

  algos.forEach((a, i) => {
    const y = 3.25 + i * 0.42;
    slide.addText(a.name, {
      x: 0.6, y, w: 1.8, h: 0.35,
      fontSize: 12, fontFace: MONO, color: a.accent, bold: true, margin: 0,
    });
    slide.addText(a.desc, {
      x: 2.5, y, w: 4, h: 0.35,
      fontSize: 11, fontFace: FONT, color: C.textDark, margin: 0,
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 2.7, w: 4.2, h: 2.5, fill: { color: "FEF2F2" }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 2.7, w: 0.06, h: 2.5, fill: { color: C.red } });
  slide.addText("⚠ 关键原则", {
    x: 5.7, y: 2.85, w: 3.7, h: 0.4,
    fontSize: 14, fontFace: FONT, color: C.red, bold: true, margin: 0,
  });
  slide.addText([
    { text: "PRNG 绝不能单独用于密码学", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "必须用 TRNG/熵源为 CSPRNG 播种", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "弱随机数是安全漏洞的根源", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "案例：Sony PS3 ECDSA nonce 复用导致签名密钥泄露", options: { bullet: true, breakLine: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
    { text: "标准：NIST SP 800-90A / BSI AIS-20", options: { bullet: true, fontSize: 11, fontFace: FONT, color: C.textDark } },
  ], { x: 5.7, y: 3.3, w: 3.7, h: 1.8, paraSpaceAfter: 5, margin: 0 });
}

// ============================================================
// SLIDE 9 — 2024-2025 趋势
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent3 } });

  slide.addText("2024-2025 前沿趋势", {
    x: 0.6, y: 0.35, w: 8, h: 0.65,
    fontSize: 30, fontFace: FONT, color: C.white, bold: true, margin: 0,
  });
  slide.addText("从学术论文到产业落地的关键技术方向", {
    x: 0.6, y: 0.95, w: 8, h: 0.4,
    fontSize: 13, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  const trends = [
    {
      num: "01", title: "混合双熵源 TRNG（Jitter + Metastability）",
      desc: "SSRO、DS-TRNG、MRO-TRNG 等设计融合两种熵源于一体，实现 300-500 Mbps 吞吐率，资源<50 LUTs，原始输出直接通过 NIST/AIS-31 测试。",
    },
    {
      num: "02", title: "无需后处理（Post-Processing-Free）",
      desc: "2025 年顶级 TRNG 设计（SSRO、DPRR、Multi-Cell）在原始输出层面即通过全部统计测试，消除后处理带来的吞吐率损失，是学界核心攻关目标。",
    },
    {
      num: "03", title: "神经网络驱动 PRNG",
      desc: "RNN/LSTM、GAN、深度强化学习三类方法。利用网络非线性与单向性难以逆向求种子。2025 综述论文系统分类。",
    },
    {
      num: "04", title: "量子随机数生成 (QRNG)",
      desc: "利用光子分束、纠缠态等量子本质不确定性。IBM/Google 投入研发，预计 2026 年大规模商用，速率有望达 100 Mbps+。",
    },
    {
      num: "05", title: "忆阻器 + 混沌电路",
      desc: "忆阻器 RTN 偏置无关设计（2025）；模拟混沌 Tent Map（0.787 pJ/bit，160μm²）；ABN 混沌网络（750 Mbps 纪录保持者）。",
    },
    {
      num: "06", title: "区块链可验证随机函数 (VRF)",
      desc: "Chainlink VRF 为数百个 DeFi 项目提供可验证链上随机数，解决预言机信任问题。",
    },
  ];

  trends.forEach((t, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.8;
    const y = 1.6 + row * 1.2;

    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.4, h: 1.0, fill: { color: C.darkBg2 } });
    slide.addText(t.num, {
      x: x + 0.25, y: y + 0.1, w: 0.5, h: 0.35,
      fontSize: 22, fontFace: MONO, color: C.accent, bold: true, margin: 0,
    });
    slide.addText(t.title, {
      x: x + 0.85, y: y + 0.1, w: 3.3, h: 0.35,
      fontSize: 13, fontFace: FONT, color: C.white, bold: true, margin: 0,
    });
    slide.addText(t.desc, {
      x: x + 0.85, y: y + 0.45, w: 3.3, h: 0.5,
      fontSize: 9.5, fontFace: FONT, color: C.textMuted, margin: 0,
    });
  });
}

// ============================================================
// SLIDE 10 — 质量评估
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.orange } });

  slide.addText("随机数质量评估标准", {
    x: 0.6, y: 0.35, w: 8, h: 0.65,
    fontSize: 30, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });

  const standards = [
    {
      name: "NIST SP 800-22", org: "美国 NIST", tests: "15 大项 188 小项",
      desc: "全球公认标准，含单比特、游程、DFT、矩阵秩、熵等测试",
    },
    {
      name: "NIST SP 800-90B", org: "美国 NIST", tests: "熵源验证",
      desc: "专用于熵源评估：独立/最近邻/马尔可夫/LZ 压缩等 min-entropy 估计方法",
    },
    {
      name: "TestU01 / BigCrush", org: "学术界", tests: "160+ 项测试",
      desc: "比 NIST 更严格。MT19937 也未全部通过（3 项失败）。TRNG 黄金标准。",
    },
    {
      name: "BSI AIS-20/31", org: "德国 BSI", tests: "K1-K4 功能等级",
      desc: "K1/K2：统计质量 | K3/K4：密码安全。可抵抗≤20年计算能力的攻击。",
    },
    {
      name: "GM/T 0005-2021", org: "中国国家密码管理局", tests: "随机性检测规范",
      desc: "适用于商用密码产品检测认证",
    },
  ];

  standards.forEach((s, i) => {
    const y = 1.4 + i * 0.78;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 0.06, h: 0.62, fill: { color: [C.accent, C.green, C.purple, C.orange, C.red][i] } });
    slide.addText(s.name, {
      x: 0.7, y, w: 2.5, h: 0.3,
      fontSize: 12, fontFace: MONO, color: C.textDark, bold: true, margin: 0,
    });
    slide.addText(`${s.org}  |  ${s.tests}`, {
      x: 0.7, y: y + 0.28, w: 3.5, h: 0.25,
      fontSize: 9, fontFace: FONT, color: C.textMuted, margin: 0,
    });
    slide.addText(s.desc, {
      x: 5.0, y, w: 4.6, h: 0.55,
      fontSize: 10.5, fontFace: FONT, color: C.textDark, margin: 0,
    });
    if (i < standards.length - 1) {
      slide.addShape(pres.shapes.LINE, { x: 0.4, y: y + 0.7, w: 9.2, h: 0, line: { color: C.border, width: 0.5 } });
    }
  });

  slide.addText("核心原则：统计测试只能拒绝「非随机」假设，不能证明随机性。理解熵源物理本质同样重要。TRNG 还需要温度/电压/重启鲁棒性测试。", {
    x: 0.4, y: 5.25, w: 9.2, h: 0.3,
    fontSize: 10, fontFace: FONT, color: C.textMuted, italic: true, margin: 0,
  });
}

// ============================================================
// SLIDE 11 — 应用领域
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.lightBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent } });

  slide.addText("应用领域", {
    x: 0.6, y: 0.35, w: 8, h: 0.65,
    fontSize: 30, fontFace: FONT, color: C.textDark, bold: true, margin: 0,
  });
  slide.addText("RNG 广泛渗透于现代计算机科学与工程的各个层面", {
    x: 0.6, y: 0.95, w: 8, h: 0.4,
    fontSize: 13, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  const apps = [
    { icon: "🔐", title: "密码学", items: "密钥生成 · IV/Nonce · 数字签名\nTLS 握手 · Salt 值" },
    { icon: "🎮", title: "游戏开发", items: "随机事件触发 · 怪物属性\n掉落概率 · 程序化地图生成" },
    { icon: "📊", title: "蒙特卡洛模拟", items: "金融风险评估 · 量子物理计算\n分子动力学 · 期权定价" },
    { icon: "🤖", title: "机器学习", items: "权重初始化 · Dropout 正则化\n数据增强 · 超参数搜索" },
    { icon: "⛓", title: "区块链 / Web3", items: "链上随机数 · DeFi 抽奖\nNFT 铸造 · 共识随机选择" },
    { icon: "📡", title: "通信仿真", items: "信道衰落模拟 · 误码率测试\n噪声生成 · 信号检测" },
    { icon: "📈", title: "统计学", items: "随机抽样 · Bootstrap\n假设检验 · A/B 测试" },
    { icon: "🔬", title: "科研计算", items: "可复现实验 · 随机化对照试验\n数值分析 · 优化算法" },
  ];

  apps.forEach((a, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const x = 0.35 + col * 2.35;
    const y = 1.6 + row * 1.9;

    slide.addShape(pres.shapes.RECTANGLE, { x, y, w: 2.2, h: 1.65, fill: { color: C.cardBg }, shadow: makeShadow() });
    slide.addText(a.icon, { x, y: y + 0.1, w: 2.2, h: 0.45, fontSize: 22, align: "center", margin: 0 });
    slide.addText(a.title, { x: x + 0.15, y: y + 0.55, w: 1.9, h: 0.3, fontSize: 13, fontFace: FONT, color: C.textDark, bold: true, align: "center", margin: 0 });
    slide.addText(a.items, { x: x + 0.15, y: y + 0.85, w: 1.9, h: 0.7, fontSize: 9, fontFace: FONT, color: C.textMuted, align: "center", margin: 0 });
  });
}

// ============================================================
// SLIDE 12 — 选择指南
// ============================================================
{
  const slide = pres.addSlide();
  slide.background = { color: C.darkBg };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: C.accent3 } });

  slide.addText("如何选择 RNG？", {
    x: 0.6, y: 0.35, w: 8, h: 0.65,
    fontSize: 30, fontFace: FONT, color: C.white, bold: true, margin: 0,
  });
  slide.addText("场景 × 硬件资源 × 安全需求 — 三维决策框架", {
    x: 0.6, y: 0.95, w: 8, h: 0.4,
    fontSize: 13, fontFace: FONT, color: C.textMuted, margin: 0,
  });

  const guide = [
    { scenario: "FPGA TRNG（轻量级）", choice: "SSRO / MRO-TRNG", note: "<50 LUTs，300-400 Mbps，无需后处理", accent: C.green },
    { scenario: "FPGA TRNG（高吞吐率）", choice: "ABN 混沌 / Multi-Cell", note: "433-750 Mbps，适合数据中心", accent: C.accent },
    { scenario: "ASIC/芯片 TRNG", choice: "Meta-RO / TERO", note: "极低面积（~1μm²），工艺成熟", accent: C.purple },
    { scenario: "IoT 超低功耗 TRNG", choice: "Tent Map 混沌电路", note: "0.787 pJ/bit，160μm²，65nm CMOS", accent: C.orange },
    { scenario: "高安全性密钥生成", choice: "QRNG / TRNG + CSPRNG", note: "量子真随机播种 → CSPRNG 扩展", accent: C.red },
    { scenario: "需要可复现结果", choice: "PRNG + 固定种子", note: "MT19937 / PCG / Xorshift", accent: C.accent3 },
    { scenario: "区块链链上随机数", choice: "Chainlink VRF", note: "可验证随机函数，防作弊", accent: C.accent2 },
  ];

  guide.forEach((g, i) => {
    const y = 1.5 + i * 0.52;
    slide.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.06, h: 0.42, fill: { color: g.accent } });
    slide.addText(g.scenario, {
      x: 0.85, y, w: 2.6, h: 0.42,
      fontSize: 11.5, fontFace: FONT, color: C.white, margin: 0, valign: "middle",
    });
    slide.addText(g.choice, {
      x: 3.6, y, w: 2.4, h: 0.42,
      fontSize: 11, fontFace: MONO, color: g.accent, bold: true, margin: 0, valign: "middle",
    });
    slide.addText(g.note, {
      x: 6.2, y, w: 3.4, h: 0.42,
      fontSize: 10, fontFace: FONT, color: C.textMuted, margin: 0, valign: "middle",
    });
  });

  slide.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 5.1, w: 9.2, h: 0.04, fill: { color: C.accent, transparency: 40 } });
  slide.addText("核心理念：理解熵源物理 → 选择匹配架构 → 后处理消除偏差 → 统计测试验证 → 鲁棒性评估", {
    x: 0.6, y: 5.18, w: 9, h: 0.4,
    fontSize: 13, fontFace: FONT, color: C.accent3, bold: true, margin: 0,
  });
}

// ============================================================
// Save
// ============================================================
pres.writeFile({ fileName: "随机数生成器_RNG_全面解析.pptx" }).then(() => {
  console.log("PPT saved: 随机数生成器_RNG_全面解析.pptx");
}).catch(err => {
  console.error("Error:", err);
});
