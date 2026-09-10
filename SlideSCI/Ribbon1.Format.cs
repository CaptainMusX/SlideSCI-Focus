using System.Drawing;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        /// <summary>
        /// 重排「格式属性」分组（原“复制格式”）：三个子分组按属性归组。
        /// 1) 形状 / 文字 / 组合：每行 = 属性标签（图标+文字）+ 复制 + 粘贴；
        /// 2) 位置：保留原有三个按钮不动；
        /// 3) 宽度 / 高度 / 裁剪：与子分组 1 同构。
        /// 复制/粘贴按钮统一为纯文本（完整功能名保留在 ScreenTip）；
        /// 属性标签用禁用的带图标按钮呈现（RibbonLabel 不支持图标，禁用按钮即“只读标签”）。
        /// </summary>
        private void InitializeFormatRibbon()
        {
            复制图片格式.Items.Clear();
            复制图片格式.Label = "格式属性";

            // ── 子分组 1：形状 / 文字 / 组合 ──
            var attributes = Factory.CreateRibbonBox();
            attributes.Name = "formatAttributeColumn";
            attributes.BoxStyle = RibbonBoxStyle.Vertical;
            attributes.Items.Add(CreateRibbonRow("formatShapeRow",
                FormatAttributeLabel("形状", "ShapeFillColor"), copyShapeStyle, pasteShapeStyle));
            attributes.Items.Add(CreateRibbonRow("formatTextRow",
                FormatAttributeLabel("文字", "FontProperties"), copyTextStyle, pasteTextStyle));
            attributes.Items.Add(CreateRibbonRow("formatGroupRow",
                FormatAttributeLabel("组合", "ObjectsGroup"), copyGroupStyle, pasteGroupStyle));
            复制图片格式.Items.Add(attributes);
            复制图片格式.Items.Add(CreateRibbonSeparator("formatPositionSeparator"));

            // ── 子分组 2：位置（保持原样）──
            复制图片格式.Items.Add(copyPosition);
            复制图片格式.Items.Add(pastePosition);
            复制图片格式.Items.Add(swapPosition);
            复制图片格式.Items.Add(CreateRibbonSeparator("formatSizeSeparator"));

            // ── 子分组 3：宽度 / 高度 / 裁剪 ──
            var sizes = Factory.CreateRibbonBox();
            sizes.Name = "formatSizeColumn";
            sizes.BoxStyle = RibbonBoxStyle.Vertical;
            sizes.Items.Add(CreateRibbonRow("formatWidthRow",
                FormatAttributeLabel("宽度", copyImgWidth.Image), copyImgWidth, pasteImgWidth));
            sizes.Items.Add(CreateRibbonRow("formatHeightRow",
                FormatAttributeLabel("高度", copyImgHeight.Image), copyImgHeight, pasteImgHeight));
            sizes.Items.Add(CreateRibbonRow("formatCropRow",
                FormatAttributeLabel("裁剪", copyCrop.Image), copyCrop, pasteCrop));
            复制图片格式.Items.Add(sizes);

            // 复制/粘贴按钮统一为纯文本，ScreenTip 保留完整功能名。
            ConfigureCopyPaste(copyShapeStyle, pasteShapeStyle);
            ConfigureCopyPaste(copyTextStyle, pasteTextStyle);
            ConfigureCopyPaste(copyGroupStyle, pasteGroupStyle);
            ConfigureCopyPaste(copyImgWidth, pasteImgWidth);
            ConfigureCopyPaste(copyImgHeight, pasteImgHeight);
            ConfigureCopyPaste(copyCrop, pasteCrop);
        }

        private static void ConfigureCopyPaste(RibbonControl copy, RibbonControl paste)
        {
            ConfigureTextOnlyButton(copy, "复制");
            ConfigureTextOnlyButton(paste, "粘贴");
        }

        private static void ConfigureTextOnlyButton(RibbonControl control, string text)
        {
            // 复制/粘贴中“形状”“文字”来自 SplitButton（带子菜单），其余为普通按钮。
            if (control is RibbonSplitButton split)
            {
                split.Label = text;
                split.OfficeImageId = null; // 与纯文本按钮一致，去掉图标
            }
            else if (control is RibbonButton button)
            {
                button.Label = text;
                button.ShowImage = false;
            }
        }

        /// <summary>属性标签：禁用的带图标按钮，呈现为“只读标签”。</summary>
        private RibbonButton FormatAttributeLabel(string text, string officeImageId)
        {
            var label = Factory.CreateRibbonButton();
            label.Name = "format" + text + "Label";
            label.Label = text;
            label.OfficeImageId = officeImageId;
            label.Enabled = false;
            label.ShowImage = true;
            label.ShowLabel = true;
            label.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            return label;
        }

        private RibbonButton FormatAttributeLabel(string text, Image image)
        {
            var label = Factory.CreateRibbonButton();
            label.Name = "format" + text + "Label";
            label.Label = text;
            label.Image = image;
            label.Enabled = false;
            label.ShowImage = true;
            label.ShowLabel = true;
            label.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            return label;
        }

        private RibbonSeparator CreateRibbonSeparator(string name)
        {
            var separator = Factory.CreateRibbonSeparator();
            separator.Name = name;
            return separator;
        }
    }
}
