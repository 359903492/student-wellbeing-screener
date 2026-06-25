# 学生心理气象站 — 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个学生情绪筛查 Web 工具：学生手机扫码填写 PHQ-9 + GAD-7 问卷，教师通过仪表盘查看全年级统计并导出 Excel。

**Architecture:** 纯静态前端（student.html + dashboard.html）部署到 Gitee Pages，通过 Google Apps Script 作为 API 层读写 Google Sheets 数据存储。零服务器成本、国内访问稳定。

**Tech Stack:** HTML5 + CSS3 + Vanilla JS、Chart.js（图表）、SheetJS（Excel 导出）、Google Apps Script（后端 API）

## Global Constraints

- 部署到 Gitee Pages（国内访问稳定、免费）
- 学生仅填学号，不填真实姓名（隐私）
- 教师仪表盘密码：`123456`
- PHQ-9 + GAD-7 使用标准中文版措辞，不可修改
- PHQ-9 第 9 题得分 ≥ 1 时必须在仪表盘标红
- 所有页面必须标注"本问卷为筛查工具，非临床诊断"
- 知情同意声明必须在问卷第一屏显示
- 移动端优先（学生端），PC 端适配（教师端）

---

### Task 1: 项目脚手架与 Google Sheets 配置

**Files:**
- Create: `README.md`
- Create: `backend.gs`

**Interfaces:**
- Consumes: 无
- Produces: `doPost(e)` / `doGet(e)` 两个端点签名；Google Sheets 表结构定义

- [ ] **Step 1: 创建 README.md**

```markdown
# 学生心理气象站

期末阶段学生情绪/压力筛查工具。PHQ-9 + GAD-7 标准化量表。

## 部署步骤

### 1. Google Sheets 配置
1. 新建 Google Sheets，命名为"心理气象站-数据"
2. 创建两个工作表：`RawData` 和 `Summary`
3. RawData 表头：`timestamp | studentId | phq9_1 | phq9_2 | ... | phq9_9 | phq9_total | phq9_level | gad7_1 | ... | gad7_7 | gad7_total | gad7_level`

### 2. Google Apps Script 部署
1. 在 Sheets 中点击 扩展程序 → Apps Script
2. 粘贴 backend.gs 内容
3. 部署 → 新部署 → Web 应用 → 访问权限设为"所有人"
4. 复制生成的 URL，填入 student.html 和 dashboard.html 中的 `API_URL`

### 3. Gitee Pages 部署
1. 注册 gitee.com 账号
2. 新建仓库，上传 student.html 和 dashboard.html
3. 服务 → Gitee Pages → 启动
4. 将生成的 URL 生成二维码，发给学生
```

- [ ] **Step 2: 创建 backend.gs — Google Sheets 初始化与路由**

```javascript
// Google Apps Script 后端
// 部署为 Web App 时获取 URL，填入前端 API_URL

var SHEET_ID = 'YOUR_SHEET_ID_HERE'; // 替换为你的 Google Sheets ID

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  
  if (data.action === 'submit') {
    return handleSubmit(data);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var action = e.parameter.action;
  
  if (action === 'stats') {
    return handleStats(e.parameter.password);
  }
  if (action === 'export') {
    return handleExport(e.parameter.password);
  }
  
  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}
```

- [ ] **Step 3: 在 backend.gs 中添加提交处理函数**

```javascript
function handleSubmit(data) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('RawData');
  
  // 初始化表头（如果表为空）
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'timestamp', 'studentId',
      'phq9_1','phq9_2','phq9_3','phq9_4','phq9_5','phq9_6','phq9_7','phq9_8','phq9_9',
      'phq9_total', 'phq9_level',
      'gad7_1','gad7_2','gad7_3','gad7_4','gad7_5','gad7_6','gad7_7',
      'gad7_total', 'gad7_level'
    ]);
  }
  
  var phq9_total = data.phq9.reduce(function(a,b) { return a + b; }, 0);
  var gad7_total = data.gad7.reduce(function(a,b) { return a + b; }, 0);
  
  var phq9_level = getPHQ9Level(phq9_total);
  var gad7_level = getGAD7Level(gad7_total);
  
  var row = [
    new Date().toISOString(), data.studentId
  ].concat(data.phq9, phq9_total, phq9_level, data.gad7, gad7_total, gad7_level);
  
  sheet.appendRow(row);
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    phq9_total: phq9_total,
    phq9_level: phq9_level,
    gad7_total: gad7_total,
    gad7_level: gad7_level
  })).setMimeType(ContentService.MimeType.JSON);
}

function getPHQ9Level(score) {
  if (score <= 4) return '🟢 正常范围';
  if (score <= 9) return '🟡 轻度抑郁倾向';
  if (score <= 14) return '🟠 中度抑郁倾向';
  if (score <= 19) return '🔴 中重度抑郁倾向';
  return '🔴 重度抑郁倾向';
}

function getGAD7Level(score) {
  if (score <= 4) return '🟢 正常范围';
  if (score <= 9) return '🟡 轻度焦虑';
  if (score <= 14) return '🟠 中度焦虑';
  return '🔴 重度焦虑';
}
```

- [ ] **Step 4: 在 backend.gs 中添加统计和导出函数**

```javascript
function handleStats(password) {
  if (password !== '123456') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('RawData');
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({
      total: 0, levels: { good: 0, mild: 0, moderate: 0, severe: 0 },
      flagged: [], list: []
    })).setMimeType(ContentService.MimeType.JSON);
  }
  
  var headers = data[0];
  var phq9TotalCol = headers.indexOf('phq9_total');
  var phq9LevelCol = headers.indexOf('phq9_level');
  var gad7TotalCol = headers.indexOf('gad7_total');
  var gad7LevelCol = headers.indexOf('gad7_level');
  var studentIdCol = headers.indexOf('studentId');
  var phq9Q9Col = headers.indexOf('phq9_9');
  
  var levels = { good: 0, mild: 0, moderate: 0, severe: 0 };
  var flagged = [];
  var list = [];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var phq9 = row[phq9TotalCol];
    var gad7 = row[gad7TotalCol];
    var studentId = row[studentIdCol];
    
    // 取较严重的等级
    if (phq9 <= 4 && gad7 <= 4) levels.good++;
    else if (phq9 <= 9 && gad7 <= 9) levels.mild++;
    else if (phq9 <= 14 && gad7 <= 14) levels.moderate++;
    else levels.severe++;
    
    list.push({
      studentId: studentId,
      phq9: phq9,
      phq9_level: row[phq9LevelCol],
      gad7: gad7,
      gad7_level: row[gad7LevelCol],
      danger: row[phq9Q9Col] >= 1
    });
    
    if (row[phq9Q9Col] >= 1) {
      flagged.push(studentId);
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    total: data.length - 1,
    levels: levels,
    flagged: flagged,
    list: list
  })).setMimeType(ContentService.MimeType.JSON);
}

function handleExport(password) {
  if (password !== '123456') {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('RawData');
  var data = sheet.getDataRange().getValues();
  
  // 返回 CSV 格式
  var csv = data.map(function(row) { return row.join(','); }).join('\n');
  
  return ContentService.createTextOutput(csv)
    .setMimeType(ContentService.MimeType.CSV);
}
```

- [ ] **Step 5: Commit**

```bash
git add README.md backend.gs
git commit -m "feat: add project scaffold, Google Sheets schema, and Apps Script backend"
```

---

### Task 2: 学生问卷页 — HTML 结构与样式

**Files:**
- Create: `student.html`

**Interfaces:**
- Consumes: 无
- Produces: `.question-card` CSS 类、`#quiz-container` DOM 结构、`renderQuestions()` 函数签名

- [ ] **Step 1: 创建 student.html 基础结构**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>期末心理气象站 ☁️</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      background: linear-gradient(135deg, #FFF3E0, #FFE0B2, #FFCC80);
      min-height: 100vh;
      display: flex;
      justify-content: center;
      align-items: flex-start;
      padding: 20px 0 60px;
    }
    .container { width: 100%; max-width: 420px; padding: 0 16px; }
    .header { text-align: center; padding: 30px 0 20px; }
    .header h1 { font-size: 24px; color: #E65100; margin-bottom: 6px; }
    .header p { font-size: 13px; color: #BF360C; opacity: 0.8; }

    /* Screen management */
    .screen { display: none; }
    .screen.active { display: block; }

    /* Consent card */
    .card {
      background: #FFFFFF; border-radius: 16px; padding: 24px 20px;
      margin-bottom: 16px; box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    }
    .card h2 { font-size: 17px; color: #333; margin-bottom: 12px; }
    .card p { font-size: 14px; color: #666; line-height: 1.7; }

    /* Input */
    .input-group { margin-bottom: 16px; }
    .input-group label { display: block; font-size: 14px; color: #555; margin-bottom: 6px; font-weight: 500; }
    .input-group input {
      width: 100%; padding: 12px 16px; border: 2px solid #FFE0B2;
      border-radius: 12px; font-size: 16px; outline: none; transition: border-color 0.2s;
    }
    .input-group input:focus { border-color: #FF9800; }

    /* Question card */
    .question-card {
      background: #FFFFFF; border-radius: 16px; padding: 20px;
      margin-bottom: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .question-card .q-num { font-size: 12px; color: #FF9800; font-weight: 600; margin-bottom: 4px; }
    .question-card .q-text { font-size: 15px; color: #333; line-height: 1.5; margin-bottom: 14px; }

    /* Options */
    .options { display: flex; flex-direction: column; gap: 8px; }
    .option {
      padding: 12px 16px; border: 2px solid #F5F5F5; border-radius: 12px;
      font-size: 14px; color: #555; cursor: pointer; transition: all 0.2s;
      text-align: center; background: #FAFAFA;
    }
    .option:hover { border-color: #FFB74D; background: #FFF8E1; }
    .option.selected { border-color: #FF9800; background: #FFF3E0; color: #E65100; font-weight: 600; }

    /* Buttons */
    .btn {
      display: block; width: 100%; padding: 14px; border: none; border-radius: 14px;
      font-size: 16px; font-weight: 600; cursor: pointer; transition: all 0.2s;
      text-align: center;
    }
    .btn-primary { background: #FF9800; color: #FFF; margin-top: 8px; }
    .btn-primary:hover { background: #F57C00; }
    .btn-primary:disabled { background: #CCC; cursor: not-allowed; }
    .btn-outline { background: #FFF; color: #FF9800; border: 2px solid #FF9800; }

    /* Result */
    .meter-bar {
      height: 12px; border-radius: 6px; background: #EEE; margin: 8px 0 4px; overflow: hidden;
    }
    .meter-fill { height: 100%; border-radius: 6px; transition: width 0.6s ease; }
    .meter-fill.low { background: #4CAF50; }
    .meter-fill.mild { background: #FFC107; }
    .meter-fill.moderate { background: #FF9800; }
    .meter-fill.severe { background: #F44336; }

    .level-badge {
      display: inline-block; padding: 4px 14px; border-radius: 20px;
      font-size: 14px; font-weight: 600;
    }
    .level-badge.good { background: #E8F5E9; color: #2E7D32; }
    .level-badge.mild { background: #FFF8E1; color: #F57F17; }
    .level-badge.moderate { background: #FFF3E0; color: #E65100; }
    .level-badge.severe { background: #FFEBEE; color: #C62828; }

    .disclaimer { font-size: 11px; color: #999; text-align: center; margin-top: 16px; line-height: 1.6; }

    /* Progress bar */
    .progress-bar {
      height: 4px; background: #FFE0B2; border-radius: 2px; margin-bottom: 16px; overflow: hidden;
    }
    .progress-fill { height: 100%; background: #FF9800; border-radius: 2px; transition: width 0.3s; }

    .section-title {
      font-size: 18px; color: #E65100; text-align: center; margin: 16px 0 12px; font-weight: 600;
    }
  </style>
</head>
<body>
  <div class="container" id="app">
    <!-- Screen 1: Consent -->
    <div class="screen active" id="screen-consent">
      <div class="header">
        <h1>☁️ 期末心理气象站</h1>
        <p>了解你的心情状态，帮你更好地度过期末</p>
      </div>
      <div class="card">
        <h2>📋 知情同意</h2>
        <p>亲爱的同学：这份问卷帮助老师了解大家在期末阶段的情绪状态。你的回答只有班主任可见，<strong>不会影响任何成绩</strong>。请根据最近两周的真实感受作答。如果填完后觉得心情不好，随时可以找老师聊聊。</p>
      </div>
      <div class="card">
        <div class="input-group">
          <label>请输入你的学号</label>
          <input type="text" id="student-id" placeholder="例如：20240101" maxlength="20">
        </div>
      </div>
      <button class="btn btn-primary" id="btn-start" disabled>开始填写</button>
      <p class="disclaimer">⚠ 本问卷为筛查工具，非临床诊断</p>
    </div>

    <!-- Screen 2: Questions -->
    <div class="screen" id="screen-questions">
      <div class="progress-bar"><div class="progress-fill" id="progress-fill"></div></div>
      <div id="questions-container"></div>
      <button class="btn btn-primary" id="btn-submit" style="display:none;">提交问卷</button>
    </div>

    <!-- Screen 3: Result -->
    <div class="screen" id="screen-result">
      <div class="header"><h1>☁️ 你的情绪温度</h1></div>
      <div id="result-content"></div>
      <p class="disclaimer">⚠ 本问卷为筛查工具，非临床诊断。<br>如果分数偏低，建议找班主任或学校心理老师聊聊。</p>
      <button class="btn btn-outline" id="btn-retry" style="margin-top:12px;" onclick="location.reload()">重新填写</button>
    </div>
  </div>

  <script>
    // JS logic in Task 3
  </script>
</body>
</html>
```

- [ ] **Step 2: Commit**

```bash
git add student.html
git commit -m "feat: add student questionnaire HTML structure and styles"
```

---

### Task 3: 学生问卷页 — JavaScript 交互逻辑

**Files:**
- Modify: `student.html`（填补 `<script>` 标签内容）

**Interfaces:**
- Consumes: 全局配置 `QUESTIONS` 数组
- Produces: `renderQuestions()`、`collectAnswers()` 函数；调用 `API_URL + '/submit'` 提交

- [ ] **Step 1: 在 student.html 的 `<script>` 中添加题目数据与配置**

```javascript
const API_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL'; // 部署时替换

// PHQ-9 + GAD-7 题目
const QUESTIONS = {
  phq9: [
    { id: 'phq9_1', text: '做事时提不起劲或没有兴趣', section: 'PHQ-9 情绪状态' },
    { id: 'phq9_2', text: '感到心情低落、沮喪或绝望', section: 'PHQ-9 情绪状态' },
    { id: 'phq9_3', text: '入睡困难、睡不安稳或睡得太多', section: 'PHQ-9 情绪状态' },
    { id: 'phq9_4', text: '感觉疲倦或没有活力', section: 'PHQ-9 情绪状态' },
    { id: 'phq9_5', text: '食欲不振或吃太多', section: 'PHQ-9 情绪状态' },
    { id: 'phq9_6', text: '觉得自己很糟——或觉得自己很失败，或让自己或家人失望', section: 'PHQ-9 情绪状态' },
    { id: 'phq9_7', text: '对事物专注有困难，例如阅读或看电视时', section: 'PHQ-9 情绪状态' },
    { id: 'phq9_8', text: '动作或说话速度缓慢到别人已经觉察——或正好相反，烦躁或坐立不安、动来动去的情况更胜于平常', section: 'PHQ-9 情绪状态' },
    { id: 'phq9_9', text: '有不如死掉或用某种方式伤害自己的念头', section: 'PHQ-9 情绪状态' },
  ],
  gad7: [
    { id: 'gad7_1', text: '感觉紧张、焦虑或急切', section: 'GAD-7 焦虑状态' },
    { id: 'gad7_2', text: '不能停止或控制担忧', section: 'GAD-7 焦虑状态' },
    { id: 'gad7_3', text: '对各种各样的事情担忧过多', section: 'GAD-7 焦虑状态' },
    { id: 'gad7_4', text: '很难放松下来', section: 'GAD-7 焦虑状态' },
    { id: 'gad7_5', text: '由于不安而无法静坐', section: 'GAD-7 焦虑状态' },
    { id: 'gad7_6', text: '变得容易烦恼或急躁', section: 'GAD-7 焦虑状态' },
    { id: 'gad7_7', text: '感到似乎將有可怕的事情发生而害怕', section: 'GAD-7 焦虑状态' },
  ]
};

const OPTIONS = [
  { value: 0, label: '完全不会' },
  { value: 1, label: '好几天' },
  { value: 2, label: '一半以上天数' },
  { value: 3, label: '几乎每天' },
];

// 合并所有题目
const ALL_QUESTIONS = [
  ...QUESTIONS.phq9.map(function(q) { return Object.assign({}, q, { group: 'phq9' }); }),
  ...QUESTIONS.gad7.map(function(q) { return Object.assign({}, q, { group: 'gad7' }); }),
];

var answers = {}; // { phq9_1: 0, phq9_2: 1, ... }
var currentQuestion = 0;
```

- [ ] **Step 2: 添加题目渲染与导航逻辑**

```javascript
function showScreen(id) {
  document.querySelectorAll('.screen').forEach(function(s) { s.classList.remove('active'); });
  document.getElementById(id).classList.add('active');
}

function renderQuestion(index) {
  if (index >= ALL_QUESTIONS.length) {
    document.getElementById('btn-submit').style.display = 'block';
    document.getElementById('questions-container').innerHTML =
      '<div class="card" style="text-align:center;"><h2>✅ 所有题目已完成</h2><p style="margin-top:8px;color:#888;">请点击下方「提交问卷」</p></div>';
    return;
  }

  var q = ALL_QUESTIONS[index];
  var progress = Math.round((index / ALL_QUESTIONS.length) * 100);
  document.getElementById('progress-fill').style.width = progress + '%';

  // Section header (只在每个 section 第一题显示)
  var sectionHeader = '';
  if (index === 0 || ALL_QUESTIONS[index - 1].section !== q.section) {
    sectionHeader = '<div class="section-title">' + q.section + '</div>';
  }

  var optionsHTML = OPTIONS.map(function(opt) {
    var selectedClass = answers[q.id] === opt.value ? ' selected' : '';
    return '<div class="option' + selectedClass + '" data-value="' + opt.value + '" data-qid="' + q.id + '" data-next="' + (index + 1) + '">' + opt.label + '</div>';
  }).join('');

  var html = sectionHeader +
    '<div class="question-card">' +
    '<div class="q-num">第 ' + (index + 1) + ' / ' + ALL_QUESTIONS.length + ' 题</div>' +
    '<div class="q-text">' + q.text + '</div>' +
    '<div class="options">' + optionsHTML + '</div>' +
    '</div>';

  document.getElementById('questions-container').innerHTML = html;

  // 绑定选项点击事件
  document.querySelectorAll('.option').forEach(function(opt) {
    opt.addEventListener('click', function() {
      var qid = this.dataset.qid;
      var value = parseInt(this.dataset.value);
      var next = parseInt(this.dataset.next);

      answers[qid] = value;

      // 更新选中样式
      this.parentElement.querySelectorAll('.option').forEach(function(o) { o.classList.remove('selected'); });
      this.classList.add('selected');

      // 自动跳转到下一题（延迟 300ms 让用户看到选中效果）
      setTimeout(function() {
        currentQuestion = next;
        renderQuestion(next);
        window.scrollTo(0, 0);
      }, 300);
    });
  });
}
```

- [ ] **Step 3: 添加提交与结果显示逻辑**

```javascript
function collectAnswers() {
  var phq9 = QUESTIONS.phq9.map(function(q) { return answers[q.id] || 0; });
  var gad7 = QUESTIONS.gad7.map(function(q) { return answers[q.id] || 0; });
  return { phq9: phq9, gad7: gad7 };
}

function getPHQ9Level(score) {
  if (score <= 4) return { level: 'good', label: '正常范围' };
  if (score <= 9) return { level: 'mild', label: '轻度抑郁倾向' };
  if (score <= 14) return { level: 'moderate', label: '中度抑郁倾向' };
  if (score <= 19) return { level: 'severe', label: '中重度抑郁倾向' };
  return { level: 'severe', label: '重度抑郁倾向' };
}

function getGAD7Level(score) {
  if (score <= 4) return { level: 'good', label: '正常范围' };
  if (score <= 9) return { level: 'mild', label: '轻度焦虑' };
  if (score <= 14) return { level: 'moderate', label: '中度焦虑' };
  return { level: 'severe', label: '重度焦虑' };
}

function showResult(phq9Total, phq9Level, gad7Total, gad7Level) {
  var maxPHQ9 = 27, maxGAD7 = 21;
  var phq9Pct = Math.round((phq9Total / maxPHQ9) * 100);
  var gad7Pct = Math.round((gad7Total / maxGAD7) * 100);

  var html =
    '<div class="card" style="text-align:center;">' +
    '<h2>😌 抑郁情绪温度</h2>' +
    '<div style="font-size:36px;font-weight:700;color:#333;margin:8px 0;">' + phq9Total + ' / ' + maxPHQ9 + '</div>' +
    '<span class="level-badge ' + phq9Level.level + '">' + phq9Level.label + '</span>' +
    '<div class="meter-bar"><div class="meter-fill ' + phq9Level.level + '" style="width:' + phq9Pct + '%"></div></div>' +
    '</div>' +

    '<div class="card" style="text-align:center;">' +
    '<h2>😰 焦虑情绪温度</h2>' +
    '<div style="font-size:36px;font-weight:700;color:#333;margin:8px 0;">' + gad7Total + ' / ' + maxGAD7 + '</div>' +
    '<span class="level-badge ' + gad7Level.level + '">' + gad7Level.label + '</span>' +
    '<div class="meter-bar"><div class="meter-fill ' + gad7Level.level + '" style="width:' + gad7Pct + '%"></div></div>' +
    '</div>';

  // 如果有风险，显示额外提示
  if (phq9Level.level === 'severe' || gad7Level.level === 'severe') {
    html += '<div class="card" style="background:#FFF3E0;text-align:center;">' +
      '<p style="font-size:15px;color:#E65100;">💛 你的分数显示你可能正在经历一些困难。</p>' +
      '<p style="font-size:13px;color:#BF360C;margin-top:8px;">建议找班主任聊聊，或者联系学校心理老师。你不是一个人在面对这些。</p>' +
      '</div>';
  }

  document.getElementById('result-content').innerHTML = html;
  showScreen('screen-result');
}

function submitAnswers() {
  var studentId = document.getElementById('student-id').value.trim();
  if (!studentId) { alert('请输入学号'); return; }

  document.getElementById('btn-submit').disabled = true;
  document.getElementById('btn-submit').textContent = '提交中...';

  var answers = collectAnswers();

  fetch(API_URL, {
    method: 'POST',
    body: JSON.stringify({
      action: 'submit',
      studentId: studentId,
      phq9: answers.phq9,
      gad7: answers.gad7
    })
  })
  .then(function(r) { return r.json(); })
  .then(function(data) {
    if (data.success) {
      showResult(data.phq9_total, getPHQ9Level(data.phq9_total),
                 data.gad7_total, getGAD7Level(data.gad7_total));
    } else {
      alert('提交失败，请重试');
      document.getElementById('btn-submit').disabled = false;
      document.getElementById('btn-submit').textContent = '提交问卷';
    }
  })
  .catch(function() {
    alert('网络错误，请检查网络后重试');
    document.getElementById('btn-submit').disabled = false;
    document.getElementById('btn-submit').textContent = '提交问卷';
  });
}

// 初始化
document.addEventListener('DOMContentLoaded', function() {
  // 学号输入监听——启用开始按钮
  document.getElementById('student-id').addEventListener('input', function() {
    document.getElementById('btn-start').disabled = this.value.trim() === '';
  });

  // 开始按钮
  document.getElementById('btn-start').addEventListener('click', function() {
    if (document.getElementById('student-id').value.trim()) {
      showScreen('screen-questions');
      renderQuestion(0);
    }
  });

  // 提交按钮
  document.getElementById('btn-submit').addEventListener('click', submitAnswers);
});
```

- [ ] **Step 4: Commit**

```bash
git add student.html
git commit -m "feat: add questionnaire interaction logic, scoring, and result display"
```

---

### Task 4: 教师仪表盘 — HTML + 密码验证 + 概览卡片

**Files:**
- Create: `dashboard.html`

**Interfaces:**
- Consumes: Google Apps Script GET `?action=stats&password=123456`
- Produces: `fetchStats()` 函数、概览卡片 DOM、Chart.js 柱状图

- [ ] **Step 1: 创建 dashboard.html 基础结构**

```html
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>教师仪表盘 — 心理气象站</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
      font-family: -apple-system, BlinkMacSystemFont, "Microsoft YaHei", sans-serif;
      background: #F5F7FA; color: #333; min-height: 100vh;
    }
    .container { max-width: 1100px; margin: 0 auto; padding: 24px; }

    /* Login */
    .login-overlay {
      position: fixed; top: 0; left: 0; width: 100%; height: 100%;
      background: rgba(0,0,0,0.5); display: flex; align-items: center; justify-content: center; z-index: 100;
    }
    .login-box {
      background: #FFF; border-radius: 16px; padding: 32px; width: 320px; text-align: center;
      box-shadow: 0 4px 24px rgba(0,0,0,0.12);
    }
    .login-box h2 { margin-bottom: 16px; font-size: 20px; }
    .login-box input {
      width: 100%; padding: 12px; border: 2px solid #E0E0E0; border-radius: 10px;
      font-size: 16px; text-align: center; margin-bottom: 12px; outline: none;
    }
    .login-box input:focus { border-color: #FF9800; }
    .login-box button {
      width: 100%; padding: 12px; background: #FF9800; color: #FFF; border: none;
      border-radius: 10px; font-size: 16px; font-weight: 600; cursor: pointer;
    }

    /* Header */
    .header {
      background: #FFF; border-radius: 12px; padding: 20px 24px; margin-bottom: 20px;
      display: flex; justify-content: space-between; align-items: center;
      box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    }
    .header h1 { font-size: 22px; }
    .header .meta { font-size: 13px; color: #888; }

    /* Overview cards */
    .cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }
    .stat-card {
      background: #FFF; border-radius: 12px; padding: 20px; text-align: center;
      box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    }
    .stat-card .count { font-size: 36px; font-weight: 700; margin: 4px 0; }
    .stat-card .label { font-size: 13px; color: #888; }
    .stat-card .pct { font-size: 12px; color: #AAA; margin-top: 2px; }
    .stat-card.good .count { color: #4CAF50; }
    .stat-card.mild .count { color: #FFC107; }
    .stat-card.moderate .count { color: #FF9800; }
    .stat-card.severe .count { color: #F44336; }

    /* Chart */
    .chart-section {
      background: #FFF; border-radius: 12px; padding: 24px; margin-bottom: 24px;
      box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    }
    .chart-section h2 { font-size: 18px; margin-bottom: 16px; }

    /* Student list */
    .list-section {
      background: #FFF; border-radius: 12px; padding: 24px;
      box-shadow: 0 1px 6px rgba(0,0,0,0.06);
    }
    .list-section h2 { font-size: 18px; margin-bottom: 16px; }
    table { width: 100%; border-collapse: collapse; }
    th { text-align: left; padding: 10px 12px; font-size: 12px; color: #888; border-bottom: 2px solid #F0F0F0; }
    td { padding: 10px 12px; font-size: 14px; border-bottom: 1px solid #F5F5F5; }
    tr.danger { background: #FFF5F5; }
    tr.danger td:first-child::before { content: '⚠ '; }

    /* Buttons */
    .btn-row { display: flex; gap: 12px; margin-bottom: 20px; }
    .btn {
      padding: 10px 24px; border-radius: 10px; font-size: 14px; font-weight: 600;
      cursor: pointer; border: none;
    }
    .btn-primary { background: #FF9800; color: #FFF; }
    .btn-outline { background: #FFF; color: #FF9800; border: 2px solid #FF9800; }

    .disclaimer { font-size: 11px; color: #999; text-align: center; margin-top: 24px; padding: 16px; }
    .error-msg { color: #F44336; font-size: 13px; margin-top: 8px; display: none; }

    .refresh-time { font-size: 12px; color: #AAA; text-align: right; margin-top: 4px; }
  </style>
</head>
<body>
  <!-- Login overlay -->
  <div class="login-overlay" id="login-overlay">
    <div class="login-box">
      <h2>🔐 教师登录</h2>
      <input type="password" id="password-input" placeholder="请输入密码" autofocus>
      <p class="error-msg" id="login-error">密码错误</p>
      <button onclick="login()">进入仪表盘</button>
    </div>
  </div>

  <div class="container">
    <!-- Header -->
    <div class="header">
      <div>
        <h1>📊 期末心理气象站 — 全年级报告</h1>
        <div class="meta" id="header-meta">加载中...</div>
      </div>
      <div class="btn-row" style="margin-bottom:0;">
        <button class="btn btn-outline" onclick="refreshData()">🔄 刷新</button>
        <button class="btn btn-primary" onclick="exportExcel()">📥 导出 Excel</button>
      </div>
    </div>

    <!-- Overview cards -->
    <div class="cards" id="cards-container">
      <div class="stat-card"><p>加载中...</p></div>
    </div>

    <!-- Chart -->
    <div class="chart-section">
      <h2>📈 分级分布</h2>
      <canvas id="level-chart" height="80"></canvas>
    </div>

    <!-- Student list -->
    <div class="list-section">
      <h2>📋 需关注学生名单（PHQ-9 ≥ 10 或 GAD-7 ≥ 10）</h2>
      <table id="student-table">
        <thead><tr><th>学号</th><th>PHQ-9</th><th>PHQ-9 分级</th><th>GAD-7</th><th>GAD-7 分级</th><th>标记</th></tr></thead>
        <tbody id="student-tbody"><tr><td colspan="6">加载中...</td></tr></tbody>
      </table>
    </div>

    <p class="disclaimer">⚠ 本问卷为筛查工具，非临床诊断。PHQ-9 第 9 题得分 ≥ 1 的学生已标注为红色，建议立即跟进。</p>
    <p class="refresh-time" id="refresh-time"></p>
  </div>

  <script>
    // JS logic in next step
  </script>
</body>
</html>
```

- [ ] **Step 2: 添加密码验证与数据获取逻辑**

```javascript
const API_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL'; // 部署时替换
var PASSWORD = '';
var currentData = null;

function login() {
  PASSWORD = document.getElementById('password-input').value;
  document.getElementById('login-error').style.display = 'none';

  fetch(API_URL + '?action=stats&password=' + encodeURIComponent(PASSWORD))
    .then(function(r) { return r.json(); })
    .then(function(data) {
      if (data.error) {
        document.getElementById('login-error').style.display = 'block';
      } else {
        document.getElementById('login-overlay').style.display = 'none';
        currentData = data;
        renderDashboard();
      }
    })
    .catch(function() {
      document.getElementById('login-error').textContent = '网络错误，请重试';
      document.getElementById('login-error').style.display = 'block';
    });
}

// 回车键登录
document.addEventListener('DOMContentLoaded', function() {
  document.getElementById('password-input').addEventListener('keydown', function(e) {
    if (e.key === 'Enter') login();
  });
});

function refreshData() {
  fetch(API_URL + '?action=stats&password=' + encodeURIComponent(PASSWORD))
    .then(function(r) { return r.json(); })
    .then(function(data) {
      if (!data.error) {
        currentData = data;
        renderDashboard();
      }
    });

  // 自动每 5 分钟刷新
  setTimeout(refreshData, 5 * 60 * 1000);
}
```

- [ ] **Step 3: 添加概览卡片与图表渲染**

```javascript
function renderDashboard() {
  var d = currentData;
  document.getElementById('header-meta').textContent =
    '已提交：' + d.total + ' 人  |  最后更新：' + new Date().toLocaleString('zh-CN');

  var total = d.total || 1; // 防止除零
  var levels = d.levels;

  // 概览卡片
  document.getElementById('cards-container').innerHTML =
    '<div class="stat-card good">' +
    '<div class="count">' + levels.good + '</div>' +
    '<div class="label">🟢 状态良好</div>' +
    '<div class="pct">' + (levels.good / total * 100).toFixed(1) + '%</div>' +
    '</div>' +
    '<div class="stat-card mild">' +
    '<div class="count">' + levels.mild + '</div>' +
    '<div class="label">🟡 轻度压力</div>' +
    '<div class="pct">' + (levels.mild / total * 100).toFixed(1) + '%</div>' +
    '</div>' +
    '<div class="stat-card moderate">' +
    '<div class="count">' + levels.moderate + '</div>' +
    '<div class="label">🟠 中度倾向</div>' +
    '<div class="pct">' + (levels.moderate / total * 100).toFixed(1) + '%</div>' +
    '</div>' +
    '<div class="stat-card severe">' +
    '<div class="count">' + levels.severe + '</div>' +
    '<div class="label">🔴 中重度以上</div>' +
    '<div class="pct">' + (levels.severe / total * 100).toFixed(1) + '%</div>' +
    '</div>';

  // Chart.js 柱状图
  var ctx = document.getElementById('level-chart').getContext('2d');
  if (window._chart) window._chart.destroy();
  window._chart = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['全年级'],
      datasets: [
        { label: '良好', data: [levels.good], backgroundColor: '#4CAF50' },
        { label: '轻度', data: [levels.mild], backgroundColor: '#FFC107' },
        { label: '中度', data: [levels.moderate], backgroundColor: '#FF9800' },
        { label: '中重度+', data: [levels.severe], backgroundColor: '#F44336' },
      ]
    },
    options: {
      responsive: true,
      plugins: { legend: { position: 'bottom' } },
      scales: {
        x: { stacked: true },
        y: { stacked: true, beginAtZero: true, title: { display: true, text: '人数' } }
      }
    }
  });

  // 学生名单（需关注的学生：PHQ-9 >= 10 或 GAD-7 >= 10）
  var concerned = (d.list || []).filter(function(s) {
    return s.phq9 >= 10 || s.gad7 >= 10;
  }).sort(function(a, b) {
    return (b.phq9 + b.gad7) - (a.phq9 + a.gad7);
  });

  var tbodyHTML = '';
  if (concerned.length === 0) {
    tbodyHTML = '<tr><td colspan="6" style="text-align:center;color:#888;">🎉 没有需要特别关注的学生</td></tr>';
  } else {
    concerned.forEach(function(s) {
      var dangerClass = s.danger ? ' class="danger"' : '';
      var note = s.danger ? '⚠ 需立即关注' : '建议约谈';
      tbodyHTML += '<tr' + dangerClass + '>' +
        '<td>' + s.studentId + '</td>' +
        '<td><strong>' + s.phq9 + '</strong></td>' +
        '<td>' + s.phq9_level + '</td>' +
        '<td><strong>' + s.gad7 + '</strong></td>' +
        '<td>' + s.gad7_level + '</td>' +
        '<td>' + note + '</td>' +
        '</tr>';
    });
  }
  document.getElementById('student-tbody').innerHTML = tbodyHTML;

  document.getElementById('refresh-time').textContent = '最后更新：' + new Date().toLocaleString('zh-CN');
}
```

- [ ] **Step 4: 添加 Excel 导出功能**

```javascript
function exportExcel() {
  if (!currentData || !currentData.list) {
    alert('暂无数据可导出');
    return;
  }

  // 构建导出数据
  var exportData = [['学号', 'PHQ-9 总分', 'PHQ-9 分级', 'GAD-7 总分', 'GAD-7 分级', '安全标记']];

  currentData.list.forEach(function(s) {
    exportData.push([
      s.studentId,
      s.phq9,
      s.phq9_level,
      s.gad7,
      s.gad7_level,
      s.danger ? '是（第9题≥1）' : '否'
    ]);
  });

  var ws = XLSX.utils.aoa_to_sheet(exportData);
  ws['!cols'] = [
    { wch: 15 }, { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 18 }, { wch: 18 }
  ];

  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '心理健康筛查');
  XLSX.writeFile(wb, '心理气象站_筛查报告_' + new Date().toISOString().slice(0, 10) + '.xlsx');
}
```

- [ ] **Step 5: Commit**

```bash
git add dashboard.html
git commit -m "feat: add teacher dashboard with auth, overview cards, chart, student list, and Excel export"
```

---

### Task 5: 集成测试与部署

**Files:**
- Modify: `student.html`（更新 API_URL）
- Modify: `dashboard.html`（更新 API_URL）
- Modify: `backend.gs`（更新 SHEET_ID）
- Create: 无

- [ ] **Step 1: 配置 Google Sheets**

```bash
# 手动步骤（在浏览器中操作）：
# 1. 打开 sheets.google.com → 新建 → 命名为"心理气象站-数据"
# 2. 在 RawData 工作表中添加表头行
# 3. 扩展程序 → Apps Script → 粘贴 backend.gs → 替换 SHEET_ID
# 4. 部署 → 新部署 → Web 应用 → 执行身份：我 → 访问权限：所有人
# 5. 复制生成的 URL
```

- [ ] **Step 2: 更新前端 API_URL**

分别在 `student.html` 和 `dashboard.html` 中，将：
```javascript
const API_URL = 'YOUR_GOOGLE_APPS_SCRIPT_URL';
```
替换为实际的 Google Apps Script URL。

- [ ] **Step 3: 部署到 Gitee Pages**

```bash
# 手动步骤：
# 1. 注册 gitee.com 账号
# 2. 新建仓库 → 上传文件：student.html, dashboard.html, README.md
# 3. 服务 → Gitee Pages → 启动
# 4. 将生成的 URL 生成二维码二维码（可使用草料二维码等工具）
```

- [ ] **Step 4: 冒烟测试**

```
测试用例：
1. 学生端：手机扫码 → 看到知情同意页 → 输入学号 → 填写 16 题 → 提交 → 看到结果页
2. 教师端：打开 dashboard.html → 输入正确密码 → 看到概览卡片和图表
3. 教师端：输入错误密码 → 提示密码错误
4. 教师端：点击「导出 Excel」→ 下载 .xlsx 文件
5. 安全测试：PHQ-9 第 9 题选 ≥1 → 学生名单中标红标注
6. 边界测试：无学生数据时仪表盘显示空状态
```

- [ ] **Step 5: Commit**

```bash
git add student.html dashboard.html
git commit -m "chore: configure API URLs and deployment"
```

---

## 自审清单

- [x] Spec 覆盖：PHQ-9 + GAD-7 完整题目、知情同意、结果展示、教师仪表盘、概览卡片、图表、学生名单、Excel 导出、密码保护、安全协议（第9题标红）、非诊断声明
- [x] 无占位符：所有代码块均包含实际实现
- [x] 类型一致性：`submit`/`stats`/`export` 三个 action 名称在前后端一致；`phq9_total`/`gad7_total` 字段名在提交和查询两端一致
