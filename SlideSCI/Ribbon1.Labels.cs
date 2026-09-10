using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        // 第二列与「图片自动排列」的“列数量”使用同一输入宽度（4 位数字）。
        private const string LabelNumberSize = "0000";
        // 第一列第三行“字体”的输入宽度：与其它输入框一致，使该行总宽度等于上面两行的按钮宽度。
        private const string LabelFontNameSize = "0000";
        // 第三列前两行的偏移输入宽度：ComboBox 带下拉箭头，`0000` 使输入框右边界
        // 与第三行两个切换按钮的右边界对齐（真实渲染实测与第二列“列数量”同宽）。
        private const string LabelOffsetSize = "0000";
        private static readonly string[] LabelOffsetPresets =
        {
            "-20", "-10", "-5", "0", "5", "10", "20"
        };
        private const string LabelTextInset = LabelInset;
        private static readonly string[] LabelIndexPresets =
        {
            "1", "2", "3", "4", "5", "6", "7", "8", "9", "10",
            "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"
        };

        /// <summary>
        /// 重排「添加图片标签」分组：三列 × 三行。
        /// 第一列 添加标签 / 更新标签 / 字体；第二列 字号 / 模板 / 编号；
        /// 第三列 垂直偏移 / 水平偏移 / 加粗 + 编号自动更新。
        /// 行结构沿用「添加图片标题」分组的已验证做法：前两行保持裸控件以获得原生
        /// 行距，第三行用横向 RibbonBox 承载，使三列第三行落在同一条基线上。
        /// </summary>
        private void InitializeLabelRibbon()
        {
            group1.Items.Clear();

            // ── 第一列：添加标签 / 更新标签 / 字体 ──
            var actions = Factory.CreateRibbonBox();
            actions.Name = "labelActionColumn";
            actions.BoxStyle = RibbonBoxStyle.Vertical;
            ConfigureLabelButton(addLabelsButton, "添加标签", "为选中的图片添加编号标签，编号从“编号”输入框开始递增");
            ConfigureLabelButton(updateLabelsButton, "更新标签", "按当前的字体、字号、模板与偏移重新生成选中图片的标签");
            actions.Items.Add(addLabelsButton);
            actions.Items.Add(updateLabelsButton);
            labelFontNameEditBox.Label = LabelTextInset + "字体";
            labelFontNameEditBox.SizeString = LabelFontNameSize;
            actions.Items.Add(CreateRibbonRow("labelFontRow", labelFontNameEditBox));
            group1.Items.Add(actions);
            group1.Items.Add(CreateTitleColumnSeparator("labelActionSeparator"));

            // ── 第二列：字号 / 模板 / 编号：与“列数量”相同的可输入下拉框与输入宽度 ──
            var numbers = Factory.CreateRibbonBox();
            numbers.Name = "labelNumberColumn";
            numbers.BoxStyle = RibbonBoxStyle.Vertical;
            labelFontSizeEditBox.Label = LabelTextInset + "字号";
            labelFontSizeEditBox.SizeString = LabelNumberSize;
            labelFontSizeEditBox.ScreenTip = "标签字号 (pt)";
            labelFontSizeEditBox.SuperTip = "可直接输入字号，也可从下拉列表选择预设值。";
            labelTemplateComboBox.Label = LabelTextInset + "模板";
            labelTemplateComboBox.SizeString = LabelNumberSize;
            labelTemplateComboBox.ScreenTip = "编号模板";
            labelTemplateComboBox.SuperTip = "决定编号的显示形式，例如 1、1)、A、a、Ⅰ、①。";
            labelIndex.Label = LabelTextInset + "编号";
            labelIndex.SizeString = LabelNumberSize;
            labelIndex.ScreenTip = "起始编号";
            labelIndex.SuperTip = "下一个标签使用的编号；开启“编号自动更新”后每次添加标签自动递增。";
            PopulateLabelIndexPresets();
            numbers.Items.Add(labelFontSizeEditBox);
            numbers.Items.Add(labelTemplateComboBox);
            numbers.Items.Add(CreateRibbonRow("labelIndexRow", labelIndex));
            group1.Items.Add(numbers);
            group1.Items.Add(CreateTitleColumnSeparator("labelOffsetSeparator"));

            // ── 第三列：垂直偏移 / 水平偏移 / 加粗 + 编号自动更新 ──
            var offsets = Factory.CreateRibbonBox();
            offsets.Name = "labelOffsetColumn";
            offsets.BoxStyle = RibbonBoxStyle.Vertical;
            labelOffsetYEditBox.Label = LabelTextInset + "垂直偏移";
            labelOffsetYEditBox.SizeString = LabelOffsetSize;
            labelOffsetYEditBox.ScreenTip = "垂直偏移 (pt)";
            labelOffsetYEditBox.SuperTip = "正值向下，负值向上；与水平偏移叠加，可输入任意数值。";
            labelOffsetXEditBox.Label = LabelTextInset + "水平偏移";
            labelOffsetXEditBox.SizeString = LabelOffsetSize;
            labelOffsetXEditBox.ScreenTip = "水平偏移 (pt)";
            labelOffsetXEditBox.SuperTip = "正值向右，负值向左；与垂直偏移叠加，可输入任意数值。";
            PopulateLabelOffsetPresets();
            offsets.Items.Add(labelOffsetYEditBox);
            offsets.Items.Add(labelOffsetXEditBox);
            // 加粗与编号自动更新改为切换按钮，与「添加图片标题」的“编组”一致。
            ConfigureLabelToggle(labelBoldcheckBox, "加粗", "标签文字加粗");
            ConfigureLabelToggle(labelIndexUpdatecheckBox, "编号自动更新", "添加标签后自动把“编号”更新为下一个可用编号");
            offsets.Items.Add(CreateRibbonRow("labelFlagsRow", labelBoldcheckBox, labelIndexUpdatecheckBox));
            group1.Items.Add(offsets);
        }

        /// <summary>小按钮：左图标右文字，样式与「添加图片标题」的“添加上标题”一致。</summary>
        private void ConfigureLabelButton(RibbonButton button, string label, string superTip)
        {
            button.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            button.Label = label;
            button.ShowImage = true;
            button.ShowLabel = true;
            button.ScreenTip = label;
            button.SuperTip = superTip;
        }

        /// <summary>切换按钮：与「添加图片标题」的“编组”一致，无图标、按下表示开启。</summary>
        private void ConfigureLabelToggle(RibbonToggleButton toggle, string label, string superTip)
        {
            toggle.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            toggle.Label = label;
            toggle.ShowImage = false;
            toggle.ShowLabel = true;
            toggle.ScreenTip = label;
            toggle.SuperTip = superTip;
        }

        /// <summary>“垂直/水平偏移”下拉列表：与「添加图片标题」的偏移预设完全一致。</summary>
        private void PopulateLabelOffsetPresets()
        {
            labelOffsetYEditBox.Items.Clear();
            labelOffsetXEditBox.Items.Clear();
            foreach (string value in LabelOffsetPresets)
            {
                var itemY = Factory.CreateRibbonDropDownItem(); itemY.Label = value;
                var itemX = Factory.CreateRibbonDropDownItem(); itemX.Label = value;
                labelOffsetYEditBox.Items.Add(itemY);
                labelOffsetXEditBox.Items.Add(itemX);
            }
        }

        /// <summary>“编号”下拉列表：仍然可以手动输入列表以外的编号。</summary>
        private void PopulateLabelIndexPresets()
        {
            labelIndex.Items.Clear();
            foreach (string value in LabelIndexPresets)
            {
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = value;
                labelIndex.Items.Add(item);
            }
        }
    }
}
