# 导出与选择 分组（v2.0.21，2026-09-10）

v2.0.21 将「导出与选择」分组从 SciFigure 功能区移到 SciStudio 功能区（顺序：复制图片格式 / 排列 / 科研文本 / 导出与选择 / 关于）。

## 功能清单与按钮

| 按钮 | 功能 | 图标 |
| --- | --- | --- |
| 导出页面 | 把当前页 / 选中的页面 / 全部页面导出为 PNG / JPG / BMP / PDF，可选 DPI（96/150/300/600），PDF 可每页单独文件，导出后可选打开文件夹 | 自定义资源图标 |
| 导出原图 | 用 Open XML 直接取出选中图片嵌入的原始文件（不做重采样）保存到本地 | FileSave |
| 复制大图 | 把选中图片按幻灯片高度放大后复制到剪贴板；结束后恢复原尺寸/位置/锁定比 | Copy |
| 图文同缩 | 打开「图文同缩比例设置」窗口：形状缩放时文字字号按比例同步缩放 | PictureCompress |
| 全选文本框 | 按顺序选中当前幻灯片上的全部文本框与含文本占位符 | TextBoxInsert |

图标 ID 全部先用等效 customUI 渲染探测，确认有效（Copy / FileSave / PictureCompress / TextBoxInsert；FitToScreen、ExportSlides 在当前 PowerPoint 中无效，弃用）。

## 弹窗 DPI 修复（文字截断/压缩的根因）

三个弹窗均按 96 DPI 硬编码像素排版；在 200% 缩放显示器上字体按 DPI 放大、控件框尺寸不放大，导致文字溢出/被压缩。修复：

- `导出设置`（导出页面）弹窗；
- `图文同缩比例设置`（ScaleForm）；
- `对齐间距设置`（SpacingForm，同选项卡弹窗，一并加固）。

统一为 `ScientificUiTheme.ConfigureDialog + CompleteCodeBuiltLayout`（AutoScaleMode.Dpi + AutoScaleDimensions(96,96)，全部控件就位后装备缩放）——与「局部放大」弹窗（ZoomInsetForm）同款方案。Excel/系统对话框（另存为、MessageBox）不受影响。

## 验收

`tests/Run-TitleRibbonLayout.ps1` 125 项通过（选项卡归属、五个按钮图标/提示、序列化 XML）；真实 PowerPoint 渲染确认五个按钮图标与统一宽度。弹窗在用户 200% 缩放下的人工验收待安装后进行。
