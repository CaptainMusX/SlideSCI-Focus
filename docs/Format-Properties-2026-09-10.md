# 格式属性 分组（v2.0.22，2026-09-10）

v2.0.22 将「复制格式」分组改名并重排为三个子分组：

| 子分组 | 行结构 |
| --- | --- |
| 1. 形状 / 文字 / 组合 | 每行 = 只读属性标签（图标+文字，禁用按钮呈现）+ 复制 + 粘贴 |
| 2. 位置 | 复制位置（九宫格菜单）/ 粘贴位置 / 交换位置 —— 原样保留 |
| 3. 宽度 / 高度 / 裁剪 | 与子分组 1 同构 |

- 复制/粘贴按钮统一为纯文本；完整功能名保留在 ScreenTip（如“复制形状格式”“粘贴文字格式”）。
- 形状 / 文字 / 组合 的复制按钮为 SplitButton，保留“全部 / 形状填充 / 形状轮廓”“全部 / 字体 / 颜色 / 大小 / 效果”等子菜单。
- 分组标题选「格式属性」：覆盖样式与尺寸/位置/裁剪等全部属性，且不只“复制”（还有粘贴与交换）。

## 弹窗加固（对齐间距设置）

- 状态文字标签限制最大宽度，防止超长加载信息把弹窗撑宽；lblInfo 仍保持 AutoSize 高度自适应。
- app.config 增加 `EnableWindowsFormsHighDpiAutoResizing=true` 与 `DpiAwareness=PerMonitorV2`，叠加各弹窗已有的 `AutoScaleMode.Dpi + AutoScaleDimensions(96,96)` 基准缩放。

## 已知工具链限制

静态 customUI 预览渲染（`artifacts/tools/render-xml.ps1` / `render-groups.ps1`）对 VSTO SplitButton 的序列化转换暂不完整（预览文件无法显示含 SplitButton 的分组，且会连带整个预览选项卡不加载）。生产加载项由 VSTO 运行时序列化 Ribbon——含 SplitButton 的结构（复制位置菜单）在 2.0.21 实装中已验证正常。涉及 SplitButton 的改动以对象级/XML 级回归为准，最终以实装人工验收为准。
