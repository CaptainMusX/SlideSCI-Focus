# 图片标题 Ribbon 与四方向标题（v2.0.5，2026-09-08）

布局限定在“添加图片标题”分组，三个纵向列：

| 行 | 第一列 | 第二列 | 第三列 |
| --- | --- | --- | --- |
| 1 | 添加上标题（左图标、右文字） | 添加左标题（左图标、右文字） | 标题 |
| 2 | 添加下标题（左图标、右文字） | 添加右标题（左图标、右文字） | 字体 |
| 3 | 上下偏移 | 左右偏移 | 字号、编组切换按钮、图标对齐菜单 |

偏移输入使用相同窄宽度，和两行普通尺寸按钮组成紧凑列。字号、编组、对齐放在同一个水平 RibbonBox；对齐菜单折叠只显示当前 Office 原生图标，展开显示左对齐、居中、右对齐、两端对齐，并标记当前项。原生命令 ID 为 AlignLeft、AlignCenter、AlignRight、AlignJustify；参考 [Microsoft PowerPoint 控件 ID 清单](https://github.com/MicrosoftDocs/VBA-Docs/blob/main/Office-Mac/idMSOPowerPointMac.md)。

上下和左右偏移分别保存在新的 TitleOffsetY / TitleOffsetX 用户设置中，初始值均为 0 pt；正数向下/向右，负数向上/向左。两者叠加应用于新生成的所有方向标题，不会移动此前生成的其他标题。原有“图下距”保留为兼容字段，不迁移其上下相反的距离语义。旧“是否居中”偏好仅在尚未保存新的四项对齐设置时迁移。

上下标题和图片等宽；上标题按最终文字高度定位。左右标题保持横排、垂直居中，短标题按文字长度收窄，长标题在不超过图片宽度的文本框中换行。所有方向均设置中英文字体与段落对齐。批量添加保持排序；编组失败保留源对象和已生成的标题，不再执行旧的复制后删除补救流程。

验证结果：

- Release 构建和当前 Release 目录的 Inno 打包通过，版本元数据均为 2.0.5.0。
- tests/Run-CoreRegression.ps1：13 项通过。
- tests/Run-TitleSmoke.ps1：72 项通过。使用独立临时 PowerPoint 文稿验证四方向 × 四种段落对齐、双轴正负偏移叠加、实际多行高度、左右边界与垂直居中、无效参数不产生对象、新偏移默认值，以及 Office 命令 ID 可用。
- 已检查 PowerPoint 导出的四方向标题 PNG。测试图片、PPTX 与日志在 artifacts/title-smoke 和 artifacts/title-*.log，不提交二进制测试产物。
- 未修改正在打开的用户汇报文稿，未安装新版。真实 Ribbon 在用户缩放比例下的列宽、控件显示和点击交互，以及安装卸载、冷启动仍待人工验收。跨进程 GetImageMso 无法取得图像句柄，测试改为检查原生命令 ID；不将其视为图标渲染验收。

安装包：F:\SlideSCI\artifacts\inno\SlideSCI-Focus-2.0.5.0-Setup.exe。签名为本地开发证书，不是公共信任链证书。

## v2.0.6 视觉微调

- 三列之间加入 RibbonSeparator；第三列标题输入框为 `0000000000`，字体下拉框为 `00000000`，补偿下拉箭头后统一右边界。
- 第三行放入独立的垂直布局槽，避免字号/编组/对齐行紧贴字体输入框。
- 对齐菜单项改为 `RibbonButton`，不再使用带 `Checked` 状态的 `RibbonToggleButton`，因此展开菜单的四个图标保持相同外观。

v2.0.6 安装包：F:\SlideSCI\artifacts\inno\SlideSCI-Focus-2.0.6.0-Setup.exe。
