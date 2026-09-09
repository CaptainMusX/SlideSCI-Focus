# Ribbon 2.0.13 修复与验证边界

## 启动问题

本轮在旧版本启动的 PowerPoint 中观察到 SciStudio 存在而 SciFigure 缺失。
本机 PowerPoint.officeUI 没有 SciFigure 隐藏条目，源代码没有主动隐藏选项卡的路径。
旧选项卡 ID 为通用的 tab2/tab1；新版本改为 CaptainMusX_SlideSCIFocus_SciFigure /
CaptainMusX_SlideSCIFocus_SciStudio，并在窗口激活后使 Office 重新查询 SciFigure 可见性。
没有修改用户的 Office 自定义文件，没有强制覆盖用户在 Office 中隐藏选项卡的选择。

这是一项针对选项卡标识和初始化时机的修复，尚未证明旧 ID 冲突就是本机启动缺失的唯一根因。
构造后的可见性及序列化 ID 已通过检查；新版本真实冷启动显示结果待人工确认。

## 标题和局部放大

- 图片标题保持三列。第三列第三行依次为完整字号 ComboBox、对齐菜单、编组。
- 所有标题行使用相同容器深度，避免仅最后一行嵌套盒的结构差异。
- 局部放大前两列按钮为框选区域、放大选区，保留图标；尺寸共用宽度字符串 00%，能预留单位字符。
- 百分比预设带 %；放大预设仅 1x、2x、3x、4x、5x、10x。裸数字和带单位输入共用经过有限值/范围检查的解析器。
- 默认 1x。旧的原图等宽/cm 模式进入 Ribbon 后转为 1x；已有倍数设置保留。旧文件不在加载时自动重写。
- 框线、连线是独立 RibbonLabel，右侧为 RibbonMenu。菜单使用 Office Ribbon Gallery 绘制主题色、标准色和最近色。
- 菜单包含无轮廓、其他颜色、取色器、粗细、虚线、箭头；草绘禁用。封闭选区框禁用箭头，防止 PowerPoint COM 越界错误。
- 取色器只在用户点击该菜单项后执行，不会在加载或测试中采集屏幕，不保存截图。

**未完成的视觉要求：** VSTO 的 SizeString 并非像素宽度，菜单也没有独立 Width 属性。
本轮未验证前两列三行总宽度都严格等于首行按钮、第三列前两行严格等宽，
也未验证标题第三行与其他分组的基线和上下框线精确重合。
统一容器深度及输入宽度不能代替实际 Office 渲染测量。

此菜单保留插件自身的样式预设语义，并非直接复用 PowerPoint 的 Shape Outline 命令。
因此不能声称它与原生菜单的外观、主题渐变色、取色体验完全相同。

## 检查

- Build-Installer.ps1：Release、VSTO 清单及安装引导程序。
- Run-CoreRegression.ps1：单位解析、无效输入、设置迁移/原子保存、扩展样式回写、转换进程超时。
- Run-TitleRibbonLayout.ps1：实际 Release Ribbon 对象、三列三行结构、顺序、菜单重复打开、唯一 XML ID。
- Run-ZoomNativeSmoke.ps1：实际 PowerPoint COM 创建与连续更新，编组、平行/交叉/无连线、颜色/粗细/虚线/箭头/无轮廓、连接点粘附以及失败回滚。
- Build-InnoInstaller.ps1：从当前 SlideSCI/bin/Release 打包。

用户要求禁止 computer-use 后未再使用该技能，也未使用其他 UI 自动化替代。
COM 测试只创建并关闭测试演示文稿，不操作用户文档。
安装/卸载、冷启动显示、Ribbon 像素对齐、菜单展开和取色交互仍需人工验收。
