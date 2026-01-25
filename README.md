<div align="center">
  <img src="src/assets/logo.png" alt="ScoreMaster Logo" width="120" height="120" />
  <h1>ScoreMaster 成绩分析器</h1>
  <p>
    <strong>专业级成绩数据分析工具</strong><br>
    将复杂的原始成绩 JSON 一键重构成多维排行、自动分班统计的精美 Excel 报表。
  </p>

  <p>
    <a href="#功能特性">功能特性</a> •
    <a href="#技术栈">技术栈</a> •
    <a href="#快速开始">快速开始</a> •
    <a href="#使用指南">使用指南</a> •
    <a href="#常见问题">常见问题</a>
  </p>
  
  ![License](https://img.shields.io/badge/license-MIT-blue.svg)
  ![Angular](https://img.shields.io/badge/Angular-21.0-dd0031.svg)
  ![TypeScript](https://img.shields.io/badge/TypeScript-5.9-3178c6.svg)
  ![TailwindCSS](https://img.shields.io/badge/TailwindCSS-Latest-38bdf8.svg)
</div>

---

## ✨ 功能特性

- **🚀 秒级极速处理**：纯前端本地计算，无服务器交互，数据隐私绝对安全。
- **📊 多维报表生成**：
  - **全年级总分排行**：包含赋分、排名、总分统计。
  - **单科成绩排行**：自动识别各学科成绩并生成独立榜单。
  - **智能分班报表**：自动按班级拆分工作表，便于班主任查看。
  - **深度统计分析**：包含平均分、最高/最低分、优良率等多维指标。
- **🧠 智能语义识别**：自动推断考号、姓名、班级及各学科字段，无需手动映射。
- **🔗 知识库关联**：支持上传 Excel 名单库（如学籍表），自动补全缺失的姓名或班级信息。
- **🎨 品牌化输出**：支持上传图片作为 Excel 背景水印，自动调节透明度。
- **📝 自定义导出**：支持自定义输出文件名，默认智能追加 `.xlsx` 后缀。

## 🛠 技术栈

本项目基于现代前端技术栈构建，追求极致的性能与开发体验：

- **核心框架**: [Angular 21](https://angular.dev/) (Standalone Components)
- **构建工具**: [Vite 6](https://vitejs.dev/)
- **语言**: [TypeScript 5.9](https://www.typescriptlang.org/)
- **样式**: [Tailwind CSS](https://tailwindcss.com/)
- **Excel 处理**: [SheetJS (xlsx-js-style)](https://github.com/gitbrent/xlsx-js-style)
- **UI 设计**: Glassmorphism (毛玻璃风格)

## 🚀 快速开始

### 环境要求

- Node.js 18+
- npm 9+

### 安装与运行

1. **克隆项目**
   ```bash
   git clone https://github.com/YangZiyueZY/JsonToExcel_supper.git
   cd scoremaster
   ```

2. **安装依赖**
   > Windows PowerShell 用户如果遇到权限错误，请使用 `npm.cmd` 代替 `npm`
   ```bash
   npm install
   # 或者
   npm.cmd install
   ```

3. **启动开发服务器**
   ```bash
   npm run dev
   # 或者
   npm.cmd run dev
   ```
   访问 `http://localhost:3000` 即可预览。

4. **构建生产版本**
   ```bash
   npm run build
   # 或者
   npm.cmd run build
   ```
   构建产物位于 `dist/` 目录。

## 📖 使用指南

### 1. 准备数据

#### JSON 成绩单 (必须)
这是基础数据源，必须是 JSON 数组格式。
```json
[
  {
    "考号": "2024001",
    "姓名": "张三",
    "班级": "高一(1)班",
    "语文": 115,
    "数学": 138,
    "英语": 120,
    "物理": 88
  },
  {
    "id": "2024002",
    "name": "李四",
    "class": "高一(2)班",
    "chinese": 110,
    "math": 142,
    "english": 115,
    "physics": 92
  }
]
```
> **提示**：系统会自动识别中英文表头，支持 "考号/id", "姓名/name" 等多种常见命名。

#### Excel 知识库 (可选)
用于补充学生信息（例如 JSON 中只有考号，想补全姓名和班级）。
- 格式：`.xlsx` 或 `.xls`
- 规则：第一列必须是与 JSON 中匹配的 **唯一标识** (如考号/学号)。

### 2. 操作步骤

1. **上传 JSON**：点击主区域上传你的成绩单 JSON 文件。
2. **(可选) 关联知识库**：点击左下角 Excel 图标上传学生名单。
3. **(可选) 添加水印**：点击右下角图片图标上传校徽或 Logo。
4. **(可选) 命名文件**：在下方输入框填写期望的输出文件名。
5. **生成报表**：点击底部 "立即生成多维分析报表" 按钮。
6. **下载**：处理完成后点击 "立即下载"。

## ❓ 常见问题

### Windows 下无法运行 `npm` 命令？
如果在 PowerShell 中看到 `无法加载文件...npm.ps1，因为在此系统上禁止运行脚本` 的错误：
- **方法 A (推荐)**：直接使用 `npm.cmd` 代替 `npm`。
  ```powershell
  npm.cmd install
  npm.cmd run dev
  ```
- **方法 B**：以管理员身份修改 PowerShell 执行策略。
  ```powershell
  Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
  ```

### 生成的 Excel 打开乱码或格式不对？
本项目生成的 Excel 使用标准 OpenXML 格式，完全兼容 Microsoft Excel 2007+ 及 WPS Office。如果遇到问题，请确保使用现代版本的 Office 软件打开。

## 📄 License

MIT License © 2024 ScoreMaster Team
