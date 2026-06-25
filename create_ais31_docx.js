const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        Header, Footer, AlignmentType, LevelFormat,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        PageNumber, PageBreak } = require('docx');
const fs = require('fs');

// ============================================================
// Helper functions
// ============================================================
const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

function p(text, opts = {}) {
  const runs = Array.isArray(text) ? text : [new TextRun({ text, ...opts })];
  return new Paragraph({
    spacing: { after: opts.after || 120 },
    ...opts.paraOpts,
    children: runs,
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, font: "Arial" })],
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    children: [new TextRun({ text, font: "Arial" })],
  });
}

function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    children: [new TextRun({ text, font: "Arial" })],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { after: 80 },
    children: Array.isArray(text)
      ? text.map(t => new TextRun(typeof t === 'string' ? { text: t } : t))
      : [new TextRun({ text, size: 21 })],
  });
}

function bulletBold(label, desc) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [
      new TextRun({ text: label, bold: true, size: 21 }),
      new TextRun({ text: desc, size: 21 }),
    ],
  });
}

function makeCell(text, opts = {}) {
  return new TableCell({
    borders,
    width: opts.width ? { size: opts.width, type: WidthType.DXA } : undefined,
    shading: opts.shading ? { fill: opts.shading, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children: [new Paragraph({
      children: [new TextRun({ text, size: 18, font: "Arial", bold: opts.bold, color: opts.color || "1A1A2E" })],
    })],
  });
}

function makeRow(cells, shading) {
  return new TableRow({
    children: cells.map(c => makeCell(c.text, { ...c, shading: c.shading || shading })),
  });
}

// ============================================================
// Document
// ============================================================
const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 21 } },
    },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: "1A3A5C" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2176A1" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "2C7A7B" },
        paragraph: { spacing: { before: 180, after: 120 }, outlineLevel: 2 },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } },
        }],
      },
    ],
  },
  sections: [
    // ============================================================
    // COVER PAGE
    // ============================================================
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      children: [
        new Paragraph({ spacing: { before: 3000 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "AIS 31 随机数生成器功能等级标准", size: 44, bold: true, font: "Arial", color: "1A3A5C" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "章节总结与关键内容摘要", size: 28, font: "Arial", color: "2176A1" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [new TextRun({ text: "A Proposal for Functionality Classes for Random Number Generators", size: 22, font: "Arial", color: "4A5568", italics: true })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "Version 2.35 — DRAFT (September 2022)", size: 20, font: "Arial", color: "4A5568" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "作者: Matthias Peter & Werner Schindler", size: 20, font: "Arial", color: "4A5568" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "Bundesamt für Sicherheit in der Informationstechnik (BSI)", size: 20, font: "Arial", color: "4A5568" })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 100 },
          children: [new TextRun({ text: "德国联邦信息安全办公室", size: 18, font: "Arial", color: "8899AA" })],
        }),
        new Paragraph({ children: [new PageBreak()] }),
      ],
    },

    // ============================================================
    // TABLE OF CONTENTS
    // ============================================================
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "1A3A5C", space: 4 } },
            children: [new TextRun({ text: "AIS 31 章节总结", size: 16, font: "Arial", color: "8899AA" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 4 } },
            children: [
              new TextRun({ text: "第 ", size: 16, color: "8899AA" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "8899AA" }),
              new TextRun({ text: " 页", size: 16, color: "8899AA" }),
            ],
          })],
        }),
      },
      children: [
        h1("目录"),
        new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
        new Paragraph({ children: [new PageBreak()] }),
      ],
    },

    // ============================================================
    // CHAPTER 1: INTRODUCTION
    // ============================================================
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "1A3A5C", space: 4 } },
            children: [new TextRun({ text: "AIS 31 章节总结", size: 16, font: "Arial", color: "8899AA" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 4 } },
            children: [
              new TextRun({ text: "第 ", size: 16, color: "8899AA" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "8899AA" }),
              new TextRun({ text: " 页", size: 16, color: "8899AA" }),
            ],
          })],
        }),
      },
      children: [
        h1("第一章 引言 (Introduction)"),

        h2("1.1 前言 (Foreword)"),
        p("随机数是大多数密码应用的基础需求。会话密钥、签名参数、Nonces、挑战值、盲化与掩码值等都需要高质量的随机数。弱随机数生成器(RNG)会决定性地削弱密码应用的安全性，因此对RNG进行可靠且可信的安全评估至关重要。"),
        p("在德国Common Criteria (CC)认证体系中，约二十年来AIS 20和AIS 31一直规定RNG的评估方法。它们定义了不同类型RNG的功能等级(Functionality Classes)，以及确定性和真随机数生成器的评估方法论。"),
        p("本文档是AIS 20和AIS 31的数学技术参考(Mathematical-Technical Reference)，面向开发者、评估者和认证机构。文档区分了DRNG、PTRNG和NPTRNG三类RNG，定义了六个功能等级: DRG.2、DRG.3、DRG.4、PTG.2、PTG.3、NTG.1。相比前版本[AIS2031An_11]，取消了DRG.1和PTG.1两个等级。"),

        h2("1.2 文档性质 (Character of this Document)"),
        p("功能等级的规范考虑了不同密码应用的安全需求差异，以及设备资源和限制的差异。本文档不将功能等级分配给具体密码应用——这属于密码应用安全评估的范畴。BSI技术指南TR-02102推荐了密码机制及相应的RNG功能等级。"),
        p("本文档不包含完整的评估方法论(那由AIS 20/31引用的独立文件规定)，而是为每个功能等级规定了申请人必须向评估者提供的交付物清单(包括安全证明)。这种方式使本文档可以轻松应用于CC以外的评估体系。"),

        h2("1.3 文档结构 (Structure of this Document)"),
        p("文档由五章组成:"),
        bulletBold("第一章: ", "介绍文档背景与整体定位。"),
        bulletBold("第二章: ", "解释AIS 20/31的范围与局限，RNG分类和功能等级的基本概念，PTRNG随机模型，并简要介绍其他RNG标准。"),
        bulletBold("第三章 (核心): ", "定义六个功能等级(DRG.2/3/4, PTG.2/3, NTG.1)，提供应用说明，解释每个要求的具体应用方法。"),
        bulletBold("第四章: ", "提供核心数学概念——概率论、熵理论、Guess Work、随机映射迭代、随机模型、在线测试/总失效测试/启动测试理论，以及评估者黑盒测试套件。"),
        bulletBold("第五章: ", "通过复杂实例说明第四章的概念，包括算法后处理、噪声源评估、在线测试实现、NIST SP800-90A一致性分析、Linux RNG分析。"),

        new Paragraph({ children: [new PageBreak()] }),

        // ============================================================
        // CHAPTER 2
        // ============================================================
        h1("第二章 AIS 20/31 — 范围、局限、RNG分类与概念"),

        h2("2.1 范围与局限 (Scope and Limits)"),
        p("本文档处理DRNG(确定性RNG)、PTRNG(物理RNG)和NPTRNG(非物理真RNG)。六种功能等级的技术要求是技术中立的，以鼓励研究和新的设计发展。RNG是否实际满足这些要求必须在安全评估中进行验证。"),
        p("以下方面超出本文档范围，但属于产品整体安全评估的一部分:"),
        bullet("实现攻击防御(硬件攻击、侧信道攻击、故障攻击)——适用于高安全级别设备(如智能卡)"),
        bullet("内存和数据通道保护——敏感数据(如DRNG内部状态、PTRNG缓存区)的未授权访问和操作防护"),
        bullet("已知答案测试(KAT)等自检测试——确保设备基本功能正常"),
        bullet("应用特定转换——从随机比特串到应用特定随机值(如ECDSA参数或RSA素数)的转换"),

        h2("2.2 RNG分类与功能等级 (RNG Classification and Functionality Classes)"),
        p("评估RNG的核心问题是: 哪些属性构成安全的RNG? 理想RNG应输出所有可取值等概率且独立于前后值。但现实中几乎不可能构建严格意义上的理想RNG，也无法证明或验证理想随机性。相反，目标是构建在特定意义上'接近'理想RNG的生成器。"),
        p("本文档将RNG分为三大类:"),
        bulletBold("DRNG (确定性RNG/伪随机数生成器): ", "通过确定性方式将短随机种子'扩展'为长随机序列。输出序列看起来随机，但总熵不超过种子熵。示例: NIST SP800-90A批准的DRBG设计。"),
        bulletBold("PTRNG (物理真RNG): ", "基于物理噪声源(物理实验或电子电路)产生高熵随机比特。通常可以精确估计熵值。示例: 环形振荡器(基于热噪声)、齐纳二极管。"),
        bulletBold("NPTRNG (非物理真RNG): ", "从非物理噪声源收集熵。通常难以精确估计熵值，因为NPTRNG可能运行在设计者无法控制的不同平台上。示例: Linux /dev/random。"),
        p("实际上不是总能严格区分这三类，因为RNG可能同时具有DRNG和TRNG的设计特征(如DRNG在生命周期中获得额外输入，TRNG使用密码学后处理)。本文档定义了六个功能等级(DRG.2/3/4, PTG.2/3, NTG.1)，其中DRG.4、PTG.3、NTG.1定义了混合型RNG(Hybrid RNG)。"),

        h2("2.3 PTRNG随机模型 (Stochastic Model for PTRNGs)"),
        p("TRNG的关键特性是能够在离开设备前产生含有一定熵量的随机数——任何外部实体，无论其设计知识、计算能力或密码分析能力如何，都对下一个随机数的值存在达到熵声明程度的不确定性。评估PTRNG比评估DRNG更为困难，因为相同设计在不同硬件上可能表现完全不同。"),
        p("PTRNG评估的核心任务: 验证每内部随机数比特的(平均)熵超过指定的下界。PTG.2和PTG.3都要求可验证的随机模型，该模型将原始随机数的随机行为追溯到物理现象。"),
        p("关键概念区分:"),
        bulletBold("原始随机数 (Raw Random Numbers / das-random-numbers): ", "物理噪声源的数字化输出。"),
        bulletBold("算法后处理 (Algorithmic Post-processing): ", "非密码学变换(如XOR、Von Neumann去偏)、可选。"),
        bulletBold("密码学后处理 (Cryptographic Post-processing): ", "带记忆的密码学变换(混合RNG需要)。"),
        bulletBold("内部随机数 (Internal Random Numbers): ", "经后处理、准备输出的数据。"),

        h2("2.4 其他RNG标准 (Other RNG Standards)"),
        p("文档简要介绍了NIST SP800-90系列等国际RNG标准，为后续第5.3节的一致性分析提供背景。"),

        new Paragraph({ children: [new PageBreak()] }),

        // ============================================================
        // CHAPTER 3: FUNCTIONALITY CLASSES (CORE)
        // ============================================================
        h1("第三章 功能等级 (Functionality Classes) — 核心章节"),

        h2("3.1 评估RNG的一般方面"),
        p("评估者需验证RNG设计是否满足所声明功能等级的所有要求。评估涵盖文档审查、设计分析、安全证明验证和(对PTRNG)随机模型的评估。RNG通常不是CC评估的目标对象(TOE)，而是更大设备(如智能卡或软件产品)的组成部分。"),

        h2("3.2 功能等级概览"),
        p("六个功能等级层次关系:"),
      ],
    },

    // Table for functionality classes needs its own section
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "1A3A5C", space: 4 } },
            children: [new TextRun({ text: "AIS 31 章节总结", size: 16, font: "Arial", color: "8899AA" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 4 } },
            children: [
              new TextRun({ text: "第 ", size: 16, color: "8899AA" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "8899AA" }),
              new TextRun({ text: " 页", size: 16, color: "8899AA" }),
            ],
          })],
        }),
      },
      children: [
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [1500, 1200, 2000, 1200, 3126],
          rows: [
            makeRow([
              { text: "类别", bold: true, color: "FFFFFF" },
              { text: "等级", bold: true, color: "FFFFFF" },
              { text: "类型", bold: true, color: "FFFFFF" },
              { text: "安全强度", bold: true, color: "FFFFFF" },
              { text: "核心特征", bold: true, color: "FFFFFF" },
            ], "1A3A5C"),
            makeRow([
              { text: "DRNG" }, { text: "DRG.2" }, { text: "确定性" }, { text: "基础级" },
              { text: "前后向保密, 密码学ϕ或ψ" },
            ]),
            makeRow([
              { text: "DRNG" }, { text: "DRG.3" }, { text: "确定性" }, { text: "增强级" },
              { text: "+增强后向保密, ϕ和ψ都密码学, 种子树" },
            ], "EDF2F7"),
            makeRow([
              { text: "DRNG" }, { text: "DRG.4" }, { text: "混合型" }, { text: "最高级(DRNG)" },
              { text: "+密码学后处理带记忆, 持续熵刷新" },
            ]),
            makeRow([
              { text: "PTRNG" }, { text: "PTG.2" }, { text: "物理真随机" }, { text: "基础级" },
              { text: "随机模型, 在线/总失效测试, 熵>0.9998/bit" },
            ], "EDF2F7"),
            makeRow([
              { text: "PTRNG" }, { text: "PTG.3" }, { text: "混合型" }, { text: "最强级(整体)" },
              { text: "+密码学后处理(强制), c_rate≥1" },
            ]),
            makeRow([
              { text: "NPTRNG" }, { text: "NTG.1" }, { text: "非物理真随机" }, { text: "—" },
              { text: "跨平台, min-entropy评估" },
            ], "EDF2F7"),
          ],
        }),

        new Paragraph({ spacing: { before: 200, after: 300 }, children: [new TextRun({ text: "注: DRG.1和PTG.1已在v2.35中取消。混合型RNG要求密码学后处理带记忆。", size: 18, italics: true, color: "4A5568" })] }),

        // DRG.2
        h2("3.3 DRNG: 功能等级"),
        h3("3.3.3 功能等级 DRG.2"),
        p("DRG.2是确定性RNG的基础等级，适用于不需要保证新鲜熵的密码应用。"),
        p("核心要求:"),
        bullet("(DRG.2.1) 种子材料由TRNG生成(推荐PTG.2/PTG.3/NTG.1合规的TRNG)"),
        bullet("(DRG.2.2) 两次种子过程间最多2^48次请求，单次请求≤2^19 bits (不需满足原子性条件)"),
        bullet("(DRG.2.3) 有效内部状态≥252 bits"),
        bullet("(DRG.2.4) 初始内部状态: min-entropy≥240 bits 或 Shannon熵≥250 bits"),
        bullet("(DRG.2.5) 前向保密 (内部随机数粒度)"),
        bullet("(DRG.2.6) 后向保密 (内部随机数粒度)"),
        bullet("(DRG.2.7) 如适用: 额外输入不削弱DRNG强度，即使攻击者能控制额外输入"),
        bullet("(DRG.2.8) 状态转移函数ϕ或输出函数ψ应是密码学的"),
        bullet("(DRG.2.9) 有强证据表明统计测试套件无法实际区分内部随机数与理想RNG输出"),

        h3("3.3.4 功能等级 DRG.3"),
        p("DRG.3在DRG.2基础上增加以下要求:"),
        bullet("(DRG.3.7) 增强后向保密 (Enhanced Backward Secrecy) — 以请求(request)为粒度，比DRG.2仅要求内部随机数粒度更强"),
        bullet("(DRG.3.9) 状态转移函数ϕ和输出函数ψ都必须是密码学的，且ϕ必须是单向函数(one-way function)"),
        bullet("(DRG.3.1) 种子树结构 — 支持DRNG种子传递: 根DRNG必须用TRNG种子，子DRNG可由根DRNG或其他DRNG种子，形成种子树(height限制)"),
        p("DRG.3适用于除需要保证新鲜熵之外的所有密码应用。种子树设计减轻了多DRNG实例的认证负担。"),

        h3("3.3.5 功能等级 DRG.4"),
        p("DRG.4是混合型DRNG(Hybrid DRNG)，结合了DRNG的计算安全性和TRNG的信息论安全性。"),
        p("在DRG.3基础上，DRG.4的核心特征是:"),
        bullet("密码学后处理算法必须带记忆(有内部状态)"),
        bullet("输出由熵输入持续刷新——防止内部状态因长时间运行而熵耗尽"),
        bullet("具备与PTG.2/PTG.3或NTG.1相当的熵源"),
        p("DRG.4是最强的DRNG等级。即使熵源暂时失效，DRG.4仍保持DRG.3的安全特性(因为密码学后处理)。"),

        // PTRNG
        h2("3.4 PTRNG: 功能等级"),
        h3("3.4.3 功能等级 PTG.2"),
        p("PTG.2是物理真随机数生成器的基础等级。核心要求:"),
        bullet("(PTG.2.1) 原始随机数满足时间局部平稳性(time-local stationarity)"),
        bullet("(PTG.2.2) 必须有正式的随机模型(Stochastic Model)，将随机行为追溯至物理现象"),
        bullet("(PTG.2.3) Shannon熵每比特>0.9998，或min-entropy每比特>0.98"),
        bullet("(PTG.2.4) 有效的在线测试(Online Test) — 实时监控噪声源是否偏离模型"),
        bullet("(PTG.2.5) 有效的总失效测试(Total Failure Test) — 检测噪声源完全失效"),
        bullet("(PTG.2.6) 启动测试(Start-up Test) — 上电/复位后验证噪声源正常"),
        p("PTG.2可选用非密码学后处理(如Von Neumann去偏、固定压缩率、XOR等)。但不推荐直接将PTG.2用于任意密码应用，而是用作DRNG的种子源或PTG.3的核心组件。原因是其允许的微小熵缺陷对某些应用(如ECDSA签名)可能存在理论上的安全风险。"),

        h3("3.4.4 功能等级 PTG.3"),
        p("PTG.3是AIS 20/31中整体安全性最强的功能等级。核心要求:"),
        bullet("(PTG.3.1-3) 继承PTG.2的所有要求(随机模型、在线/总失效/启动测试)"),
        bullet("(PTG.3.4) 必须使用密码学后处理(带记忆)"),
        bullet("(PTG.3.5) 压缩率c_rate ≥ 1 — 密码学后处理的输入比特数不能少于输出比特数"),
        bullet("(PTG.3.6) 支持多种熵声明方式: ①Shannon熵每比特>0.9998 ②min-entropy每比特>0.98 ③PTRNG特定熵界"),
        bullet("(PTG.3.7-8) 密码学后处理的设计必须使熵分析可行"),
        p("PTG.3的典型设计: PTG.2合规的PTRNG → 中间随机数 → DRG.3合规的密码学后处理 → 内部随机数。即使PTG.2核心部分突然输出可预测值，PTG.3整体仍保持DRG.3的安全性(DNRG后备安全层)。"),

        // NPTRNG
        h2("3.5 NPTRNG: 功能等级"),
        h3("3.5.3 功能等级 NTG.1"),
        p("NTG.1是非物理真随机数生成器。关键特征:"),
        bullet("熵源不在设计者控制之下(跨平台运行)"),
        bullet("使用min-entropy量化每比特熵下界(不使用Shannon熵，因为缺乏精确定量基础)"),
        bullet("同样需要在线测试和总失效测试"),
        bullet("不需要随机模型(因为无法建模)——这是与PTG.2/PTG.3的重要区别"),
        bullet("典型案例: Linux内核的/dev/random实现"),
        p("NTG.1可作为DRNG的种子源(在DRG.2/3/4中引用)。由于熵估计较为保守(只使用min-entropy)，NTG.1认证的安全性论证通常比PTG.2/3更依赖测试数据。"),

        h2("3.6 跨类主题 (Cross-class Topics)"),
        p("本章讨论了适用于多个功能等级的共同主题:"),
        bullet("密码学后处理的模棱两可性: 如何界定后处理是属于PTRNG还是属于DRNG部分"),
        bullet("有效内部状态的计算: 在攻击者已知大量内部随机数的假设下确定有效内部状态大小"),
        bullet("熵声明类型选择: min-entropy(保守但最可靠) vs Shannon熵(需要平稳性假设)的权衡"),
        bullet("种子传递链的安全性: 当DRNG种子另一个DRNG时，如何保证整条链的安全性不降级"),

        new Paragraph({ children: [new PageBreak()] }),

        // ============================================================
        // CHAPTER 4: MATHEMATICAL BACKGROUND
        // ============================================================
        h1("第四章 数学基础 (Mathematical Background)"),

        h2("4.1 随机性与随机实验"),
        p("介绍随机性的基本概念，定义随机实验、样本空间、事件等基础术语。理想RNG的数学定义: 输出独立且均匀分布的随机变量序列。本章为后续更深入的数学分析奠定概念基础。"),

        h2("4.2 概率论、随机过程、随机变量"),
        p("覆盖评估RNG所需的核心概率论知识:"),
        bullet("概率空间、条件概率、独立性"),
        bullet("随机变量及其分布(二项分布B(1,p)是噪声源建模中最基础的工具)"),
        bullet("大数定律与中心极限定理"),
        bullet("马尔可夫不等式、切比雪夫不等式等实用工具"),
        bullet("统计假设检验基础 — 第一类错误(误报)和第二类错误(漏报)"),

        h2("4.3 熵与猜测工作量 (Entropy and Guess Work)"),
        p("熵是度量随机性的核心概念。文档详细讨论了三种熵:"),
        bulletBold("Min-Entropy (最小熵): ", "H_min = -log2(max(p_i)) — 只考虑最可能出现的值，是最保守的熵度量，也是安全认证中最常用的指标。"),
        bulletBold("Shannon熵: ", "H = -Σ p_i·log2(p_i) — 信息论标准度量，反映平均不确定性。需要平稳分布假设。"),
        bulletBold("Collision熵: ", "H_coll = -log2(Σ p_i^2) — 基于碰撞概率的熵度量。"),
        p("Guess Work(猜测工作量): 攻击者以最优策略猜测随机值所需的期望工作量。对给定成功概率λ，工作因子w_λ量化了猜测难度。min-entropy与w_1/2直接相关。密钥结论: min-entropy是保守但最可靠的熵度量。"),

        h2("4.4 随机映射 (Random Mappings)"),
        p("随机映射理论是分析密码学后处理对熵影响的核心工具。文档详细分析了:"),
        bullet("随机映射迭代的统计特性: 在大小为n的集合A上，随机映射的期望循环长度、尾长等统计量"),
        bullet("随机映射与随机排列的比较: 两者的统计特性有显著差异"),
        bullet("对工作因子和熵的影响: 压缩率c_diff = (#输入比特 - #输出比特)比压缩率c_rate对输出熵的增加有更显著的影响——这是因为域和像空间的大小比为2^(c_diff)"),
        bullet("关键公式: 后处理后的min-entropy缺陷与输入熵、压缩参数之间的关系"),

        h2("4.5 随机模型、在线测试、总失效测试、启动测试"),
        p("这是PTRNG评估中最重要的理论基础部分:"),
        bulletBold("随机模型: ", "描述噪声源输出概率分布的数学模型。对硬币抛掷模型B(1,p)，参数p描述偏向程度。随机模型将原始随机数的统计特性追溯到物理现象。"),
        bulletBold("在线测试: ", "持续检测噪声源是否偏离随机模型预期。关键概念——参数空间划分为A_good(好区域)、A_bad(坏区域)和A_real(实际区域)。理想在线测试在A_good上从不失败但在A_bad上总是失败；实际在线测试在边界附近有显著失败概率。"),
        bulletBold("总失效测试: ", "检测噪声源的完全失效(如硬件故障)，比在线测试更严格。"),
        bulletBold("启动测试: ", "设备启动/复位后立即执行，在输出任何随机数之前验证噪声源正常工作。"),

        h2("4.6 评估者黑盒测试套件"),
        p("规定了评估者必须对PTRNG原始随机数应用的统计测试:"),
        bullet("测试套件T0_rrn (raw random numbers): 包括卡方拟合优度检验、Coron熵测试等"),
        bullet("测试套件T0_irn (internal random numbers): 针对内部随机数的测试"),
        bullet("详细规定了每个测试的参数、样本大小和判定准则"),
        bullet("强调黑盒测试不能替代随机模型分析，仅作为补充验证手段"),

        new Paragraph({ children: [new PageBreak()] }),

        // ============================================================
        // CHAPTER 5: EXAMPLES
        // ============================================================
        h1("第五章 实例 (Examples)"),

        h2("5.1 算法后处理示例"),
        p("展示了几种常见的非密码学后处理算法及其对熵的影响:"),
        bulletBold("固定压缩率 (Fixed Compression Rate): ", "将n比特输入映射到m比特输出(m<n)。关键: 压缩率必须>1才能增加每比特熵。"),
        bulletBold("Von Neumann去偏: ", "经典去偏算法——将比特对(0,1)→0，(1,0)→1，(0,0)和(1,1)→丢弃。消除偏向但显著降低吞吐率(理论上降至p(1-p)的输出率)。"),
        bulletBold("稀疏化 (Thinning Out): ", "丢弃部分比特以降低统计依赖性。"),

        h2("5.2 DRNG评估: 其他方面"),
        bullet("AES-OFB模式DRNG的详细安全评估——验证状态转移函数是否为单向函数"),
        bullet("纯DRNG与混合DRNG的区分，以及过于简单的状态转移函数的风险"),
        bullet("从AES分组密码导出的单向函数的安全性分析"),

        h2("5.3 NIST SP800-90A一致性分析"),
        p("文档第5.3节是开发者经常引用的关键章节，分析了NIST批准的DRBG设计与DRG.3和DRG.4的一致性:"),
        bulletBold("Hash DRBG安全评估: ", "详细验证了Hash DRBG满足DRG.3的所有算法要求——单向性、前后向保密、统计不可区分性。"),
        bulletBold("HMAC DRBG和CTR DRBG: ", "同样进行了符合性分析。"),
        bullet("开发者可直接引用第5.3节的分析作为安全证据，大大减轻认证负担。"),
        bullet("结论: 符合条件的NIST DRBG设计可以达到DRG.3或DRG.4等级。"),

        h2("5.4 噪声源与随机模型"),
        p("提供了丰富的噪声源设计和随机模型实例(本章最实用的部分):"),
        bulletBold("双噪声二极管PTRNG: ", "经典设计——利用两个齐纳二极管的噪声电压差。时间间隔服从Gamma分布，详细展示了模型建立、参数估计和符合性验证过程。"),
        bulletBold("等间隔采样事件(设计A): ", "基于独立同分布时间间隔的采样设计，正态分布建模T_j~N(μ,σ^2)。包含1000万次仿真实验数据。"),
        bulletBold("等间隔采样事件(设计B): ", "设计A的变体，不同的数字化策略。"),
        bulletBold("放射性衰变PTRNG: ", "利用放射性同位素衰变事件的随机时间间隔，物理随机性极好但实现复杂。"),
        bulletBold("PLL(锁相环)物理噪声源: ", "基于锁相环的抖动噪声，适合集成电路实现，展示了采样机制和随机模型。"),

        h2("5.5 在线测试"),
        bullet("单个统计测试分析: 卡方拟合优度检验在4比特字上的仿真结果"),
        bullet("更复杂的在线测试程序: 多层级测试流程(ER1→ER2→ER3)，结合多个基本测试形成完整的在线测试体系"),
        bullet("噪声报警触发概率的量化分析——给出了假阳性和假阴性的仿真数据"),

        h2("5.6 Linux /dev/random 和 /dev/urandom"),
        p("BSI委托的长期研究结果(可直接被开发者引用):"),
        bullet("分析了Linux内核RNG的架构(内核版本5.6+)，包括多种熵源(键盘/鼠标中断、磁盘I/O、网络中断、RDRAND等)"),
        bullet("熵池(Entropy Pool)→CRNG(ChaCha20-based DRNG)→/dev/random & /dev/urandom的数据流"),
        bullet("/dev/random与NTG.1的一致性验证"),
        bullet("/dev/urandom与DRG.3的一致性验证"),
        bullet("结论: Linux RNG的熵源部分可认证为NTG.1，CRNG部分可认证为DRG.3"),

        new Paragraph({ children: [new PageBreak()] }),

        // ============================================================
        // APPENDIX: GLOSSARY & KEY POINTS
        // ============================================================
        h1("附录: 核心术语与关键结论"),

        h2("术语表 (Glossary 精选)"),
      ],
    },

    // Final section with tables
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
        },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "1A3A5C", space: 4 } },
            children: [new TextRun({ text: "AIS 31 章节总结", size: 16, font: "Arial", color: "8899AA" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC", space: 4 } },
            children: [
              new TextRun({ text: "第 ", size: 16, color: "8899AA" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "8899AA" }),
              new TextRun({ text: " 页", size: 16, color: "8899AA" }),
            ],
          })],
        }),
      },
      children: [
        new Table({
          width: { size: 9026, type: WidthType.DXA },
          columnWidths: [2200, 6826],
          rows: [
            makeRow([
              { text: "术语", bold: true, color: "FFFFFF" },
              { text: "定义", bold: true, color: "FFFFFF" },
            ], "1A3A5C"),
            makeRow([
              { text: "DRNG" }, { text: "Deterministic RNG — 确定性随机数生成器，通过种子扩展生成伪随机序列" },
            ]),
            makeRow([
              { text: "PTRNG" }, { text: "Physical True RNG — 物理真随机数生成器，基于物理噪声源" },
            ], "EDF2F7"),
            makeRow([
              { text: "NPTRNG" }, { text: "Non-Physical True RNG — 非物理真随机数生成器，从非物理噪声源收集熵" },
            ]),
            makeRow([
              { text: "Hybrid RNG" }, { text: "混合型RNG — 具有DRNG和TRNG双重安全特性的RNG(DRG.4, PTG.3, NTG.1)" },
            ], "EDF2F7"),
            makeRow([
              { text: "Min-Entropy" }, { text: "最小熵: H_min = -log2(max(p_i)) — 最保守的熵度量" },
            ]),
            makeRow([
              { text: "Stochastic Model" }, { text: "随机模型 — 描述噪声源输出的概率分布模型，PTRNG评估的核心" },
            ], "EDF2F7"),
            makeRow([
              { text: "Online Test" }, { text: "在线测试 — 运行期间持续检测噪声源质量的测试" },
            ]),
            makeRow([
              { text: "Total Failure Test" }, { text: "总失效测试 — 检测噪声源完全失效的测试" },
            ], "EDF2F7"),
            makeRow([
              { text: "Forward Secrecy" }, { text: "前向保密 — 已知当前内部状态，无法计算之前的内部随机数" },
            ]),
            makeRow([
              { text: "Backward Secrecy" }, { text: "后向保密 — 已知当前内部状态，无法计算之后的内部随机数" },
            ], "EDF2F7"),
            makeRow([
              { text: "Enhanced Backward Secrecy" }, { text: "增强后向保密 — DRG.3特有，以请求为粒度的后向保密" },
            ]),
            makeRow([
              { text: "Seed Tree" }, { text: "种子树 — DRG.3中DRNG逐级种子传递的树形结构" },
            ], "EDF2F7"),
            makeRow([
              { text: "Compression Rate" }, { text: "压缩率 c_rate = n/m — 密码学后处理的输入比特数/输出比特数" },
            ]),
            makeRow([
              { text: "CC" }, { text: "Common Criteria — 信息技术安全评估通用准则" },
            ], "EDF2F7"),
            makeRow([
              { text: "BSI" }, { text: "Bundesamt für Sicherheit in der Informationstechnik — 德国联邦信息安全办公室" },
            ]),
            makeRow([
              { text: "TOE" }, { text: "Target of Evaluation — CC评估的目标对象" },
            ], "EDF2F7"),
          ],
        }),

        new Paragraph({ spacing: { before: 400 } }),

        h2("文档关键结论"),
        bullet("AIS 20/31是德国BSI发布的RNG安全评估权威标准，是Common Criteria认证中RNG评估的核心依据"),
        bullet("定义了三类RNG(DRNG/PTRNG/NPTRNG)和六个层次化功能等级(DRG.2→DRG.3→DRG.4, PTG.2→PTG.3, NTG.1)"),
        bullet("PTG.3是整体安全性最强的功能等级 — 结合了物理熵源(PTG.2核心)和密码学后处理(DRG.3级)"),
        bullet("三项核心技术支柱: ①随机模型(Stochastic Model) ②熵评估(Min-Entropy/Shannon熵) ③在线测试+总失效测试"),
        bullet("min-entropy是安全评估中最保守但最可靠的熵度量，当无法精确建模时推荐使用"),
        bullet("与NIST SP800-90A标准互补 — 文档提供了NIST DRBG设计的AIS 31一致性分析，可被开发者直接引用"),
        bullet("Linux /dev/random (NTG.1)和/dev/urandom (DRG.3)已经过BSI长期研究验证"),
        bullet("本文档技术中立、不限定具体实现，鼓励RNG设计领域的创新和研究"),

        new Paragraph({ spacing: { before: 400, after: 100 }, alignment: AlignmentType.CENTER, children: [
          new TextRun({ text: "— 全文完 —", size: 20, color: "8899AA", italics: true }),
        ]}),
      ],
    },
  ],
});

// ============================================================
// Generate file
// ============================================================
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("D:\\ClaudeCodeTest\\test5\\AIS31_章节总结.docx", buffer);
  console.log("Word document saved successfully! Size:", (buffer.length / 1024).toFixed(1), "KB");
}).catch(err => console.error(err));
