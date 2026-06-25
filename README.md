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
