<div align="center">

# SlideSCI

**一款面向科研绘图的 Microsoft PowerPoint 插件**（VSTO，Windows）

</div>

> 本仓库是 [Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI) 的个人精简分支：
> 移除了 AI 助手与素材库侧边栏，新增「局部放大」功能，并做了稳定性加固与安装器重构。

## ✨ 功能

### 🔍 局部放大（期刊 Figure 风格）
显微镜/电镜图片做"局部放大"效果（参考 Cell/Nature 等期刊样式）：

- **「插入选区框」**：在选中的图片中心放置方形选区框，可自由拖动、缩放，精确框出要放大的区域
- **「生成放大图」**：
  - **单击**：按上次设置直接生成/更新当前选区（不弹窗，快速迭代）；首次使用会自动打开设置
  - **Ctrl+单击**：打开「局部放大设置」——分段选择放大尺寸（与原图等宽[保持比例] / 指定放大倍数 / 自定义宽度 cm），
    期刊漏斗或交叉引线，连线/框线的颜色（含自定义色块）、粗细、实线/虚线，图间距，编组；
    实时示意图预览（原图/选区框/放大图/引线所见即所得），设置会自动记忆
- **连线自动吸附**：连线端点通过透明角锚点粘附到选区框与放大图四角，
  之后移动图片或调整位置，连线自动跟随拉伸，不需要重新计算
- **像素无损**：全程只做复制、裁剪、缩放、连线、编组等形状操作，不导出文件、不重新编码图像，
  原图像素原封不动
- 同一页支持多个选区；重复生成只覆盖当前选区的放大图与连线，其他实例保留

### 🖼️ 图片处理
- **图片自动排列**：3 种排序方式 × 3 种排列方式（列表格占位 / 统一高度 / 瀑布流），列数、行列间距、图宽图高（cm）实时生效
- **批量添加图片标题**：上/下标题、自定义字体/字号/距离/文本/是否居中/自动编组
- **批量添加图片标签**：A / a / A) / 1 / Ⅰ / ① 等 14 种模板，X/Y 偏移、加粗、编号自动更新
- **导出页面 / 导出原图 / 复制大图**：高 DPI 图片与 PDF 批量导出，原图直读（OpenXML）不降质

### 📋 格式与排版
- **复制/粘贴形状格式、文字格式**（字体/颜色/字号/效果细分）、**组格式**
- **复制/粘贴位置**（9 个基准点，支持多选批量）、**交换位置**
- **复制/粘贴宽、高、图片裁剪**（多选批量应用）
- **垂直/水平居中**（以第一个选中对象为参考，对齐 Illustrator 习惯）、**设置间距**

### 📝 Markdown / LaTeX / 代码
- **插入 Markdown**：整篇笔记一键粘贴为 PPT，支持标题/列表（保留悬挂缩进、任务列表）/表格/引述块/行内格式/行级与块级公式
- **文本转富文本**：把含有 Markdown 的文本框直接转为富文本
- **插入 LaTeX 文字**：PPT 原生公式（OMML），适合简单公式
- **插入 LaTeX SVG**：复杂公式走 MathJax（需 Node.js 环境，见下）
- **插入代码块**：8 种语言语法高亮（matlab/python/r/js/html/css/csharp/fortran），黑白背景切换

## 🪟 支持环境

- Windows + Microsoft PowerPoint（VSTO 加载项）
- WPS 兼容安装（不支持 Markdown/LaTeX 插入，强行使用可能卡死）
- 不支持 Mac

## 📥 安装方法

1. 下载 [Release](https://github.com/CaptainMusX/SlideSCI/releases) 中的安装包（exe 或 setup.exe）
2. **先退出 PowerPoint**，再双击安装
3. 依赖环境：.NET Framework 4.7.2、Microsoft Visual Studio 2010 Tools for Office Runtime（安装器会自动提示）

> 如果插件未显示：开发工具 → COM 加载项 → 勾选 `CaptainMusX.SlideSCI`，提示“未加载”请安装上述环境依赖后重启 PowerPoint。

## 🔧 开发与构建

- 开发环境：Visual Studio + Visual Studio Tools For Office，打开 `SlideSCI.sln`
- 命令行构建安装包：`pwsh -File .\build\Build-Installer.ps1`（详见 `build/README.md`）
- 「插入 LaTeX SVG」依赖 Node.js 运行时：
  ```
  cd <插件安装目录>/latex-converter
  npm install
  ```
  也可在构建时用 `-BundleLatexRuntime` 把 MathJax 打进安装包（默认不捆绑）

## ❓ 常见问题

- **如何把按钮加到快速访问工具栏？** 按钮右键 → 添加到快速访问工具栏
- **插入 LaTeX 显示不正常？** 简单公式用「插入LaTeX文字」，复杂公式用「插入LaTeX SVG」
- **连线不跟着图片走？** 连线已粘附到选区框与放大图四角；如果曾手动删除过选区框，请重新插入并生成

## 📄 协议与声明

- 本分支基于 [Achuan-2/SlideSCI](https://github.com/Achuan-2/SlideSCI)（AGPL-3.0）修改而来，保留原作者的全部版权与免责声明
- 本仓库代码仅用于学习与科研用途，禁止用于商业用途；任何使用风险自行承担

## 💬 问题反馈

- GitHub Issues：https://github.com/CaptainMusX/SlideSCI/issues
