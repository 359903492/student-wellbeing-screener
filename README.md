# 学生心理气象站

期末阶段学生情绪/压力筛查工具。PHQ-9 + GAD-7 标准化量表。

## 部署步骤

### 1. Google Sheets 配置

1. 打开 sheets.google.com，新建空白表格，命名为"心理气象站-数据"
2. 无需手动创建工作表或表头——后端代码首次接收提交时会自动创建 `RawData` 工作表并写入表头

### 2. Google Apps Script 部署

1. 在 Sheets 中点击 扩展程序 -> Apps Script
2. 复制 `backend.gs` 的全部内容，粘贴到 Apps Script 编辑器中
3. **重要：** 将代码中的 `YOUR_SHEET_ID_HERE` 替换为你的 Google Sheets ID
   - Sheets ID 可从浏览器地址栏提取：`https://docs.google.com/spreadsheets/d/<这里就是SHEET_ID>/edit`
4. **注意：** 默认教师密码为 `123456`（见 backend.gs 第 82 行和 143 行的 `password !== '123456'`），如需更改请修改这两处
5. 点击 部署 -> 新部署 -> 选择类型"Web 应用"
6. 执行身份设为"我"，访问权限设为"所有人"
7. 点击部署，复制弹出的 Web 应用 URL

### 3. 更新前端配置

1. 打开 `student.html`，将第 145 行的 `YOUR_GOOGLE_APPS_SCRIPT_URL` 替换为步骤 2 中复制的 URL
2. 打开 `dashboard.html`，将第 148 行的 `YOUR_GOOGLE_APPS_SCRIPT_URL` 替换为相同的 URL

### 4. Gitee Pages 部署

1. 注册 gitee.com 账号
2. 新建仓库，上传 `student.html`、`dashboard.html`（不要上传 `.superpowers/`、`backend.gs`、`node_modules/`）
3. 服务 -> Gitee Pages -> 启动
4. 将生成的 student.html 访问 URL 生成二维码（可使用草料二维码等工具），发给学生扫码填写

### 5. 冒烟测试

- 学生端：手机扫码 -> 知情同意 -> 输入学号 -> 填写 16 题 -> 提交 -> 查看结果
- 教师端：打开 dashboard.html -> 输入密码 -> 查看概览卡片和图表
- 教师端：错误密码应提示"密码错误"
- 教师端：导出 Excel 按钮下载 .xlsx 文件
- 安全测试：PHQ-9 第 9 题选 >= 1 -> 学生名单中标红
- 边界测试：无学生数据时仪表盘显示空状态

## 技术说明

- 前端：纯 HTML + CSS + JavaScript，无构建工具依赖
- 学生端 CDN 依赖：无（纯原生实现）
- 教师端 CDN 依赖：Chart.js 4.4.0（图表）+ SheetJS 0.18.5（Excel 导出）
- 后端：Google Apps Script（Web App 模式部署）
- 数据存储：Google Sheets（RawData 工作表）
- 安全：教师端密码验证（doGet 参数传递），第 9 题自伤风险自动标红
