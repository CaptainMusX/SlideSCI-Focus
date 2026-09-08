using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Office = Microsoft.Office.Core;
using P = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private const string TitleNumberSize = "00";
        private const string TitleWideSize = "00000000000000";
        private RibbonComboBox titleOffsetXCombo;
        private RibbonMenu titleAlignmentMenu;
        private RibbonMenu titleFontSizeMenu;
        private static readonly string[] TitleFontSizePresets =
        {
            "2", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "18", "20", "22", "24", "26", "28", "30", "40", "50", "60", "80", "100", "120", "150", "200"
        };
        private int titleAlignmentIndex = 1;
        private readonly string[] titleAlignmentIcons = { "AlignLeft", "AlignCenter", "AlignRight", "AlignJustify" };
        private readonly string[] titleAlignmentLabels = { "左对齐", "居中", "右对齐", "两端对齐" };

        private void InitializeTitleRibbon()
        {
            图片处理.Items.Clear();
            var vertical = Factory.CreateRibbonBox(); vertical.BoxStyle = RibbonBoxStyle.Vertical;
            图片上标题.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            图片上标题.Label = "添加上标题";
            AddTitleButton.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            AddTitleButton.Label = "添加下标题";
            vertical.Items.Add(图片上标题); vertical.Items.Add(AddTitleButton);
            distanceFromBottomEditBox.Label = "垂直偏移";
            distanceFromBottomEditBox.SizeString = TitleNumberSize;
            distanceFromBottomEditBox.ScreenTip = "垂直偏移 (pt)";
            distanceFromBottomEditBox.SuperTip = "正值向下，负值向上；与水平偏移叠加，对新生成的四个方向标题都生效。";
            vertical.Items.Add(distanceFromBottomEditBox);
            图片处理.Items.Add(vertical);
            图片处理.Items.Add(CreateTitleColumnSeparator("titleVerticalSeparator"));

            var horizontal = Factory.CreateRibbonBox(); horizontal.BoxStyle = RibbonBoxStyle.Vertical;
            horizontal.Items.Add(CreateSideTitleButton("添加左标题", PictureTitleSide.Left));
            horizontal.Items.Add(CreateSideTitleButton("添加右标题", PictureTitleSide.Right));
            titleOffsetXCombo = Factory.CreateRibbonComboBox();
            titleOffsetXCombo.Name = "titleOffsetXCombo";
            titleOffsetXCombo.Label = "水平偏移"; titleOffsetXCombo.SizeString = TitleNumberSize;
            titleOffsetXCombo.ScreenTip = "水平偏移 (pt)";
            titleOffsetXCombo.SuperTip = "正值向右，负值向左；与垂直偏移叠加，对新生成的四个方向标题都生效。";
            horizontal.Items.Add(titleOffsetXCombo);
            图片处理.Items.Add(horizontal);
            图片处理.Items.Add(CreateTitleColumnSeparator("titleHorizontalSeparator"));

            // 第三列：标题 / 字体 / 字号、编组与对齐方式。
            titleTextEditBox.SizeString = TitleWideSize;
            fontNameEditBox.Label = "字体"; fontNameEditBox.SizeString = TitleWideSize;
            fontSizeEditBox.Label = "字号"; fontSizeEditBox.SizeString = TitleNumberSize;
            var format = Factory.CreateRibbonBox(); format.BoxStyle = RibbonBoxStyle.Vertical;
            format.Name = "titleFormatColumn";
            format.Items.Add(titleTextEditBox);
            format.Items.Add(fontNameEditBox);
            图片处理.Items.Add(format);

            // 第三行避免把较高的 ComboBox 再嵌入水平盒。
            // 使用 EditBox + 独立预设菜单保留手输/预设字号，减少行高占用。
            autoGroupCheckBox.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            autoGroupCheckBox.ShowImage = false;
            autoGroupCheckBox.ShowLabel = true;
            titleAlignmentMenu = Factory.CreateRibbonMenu();
            titleAlignmentMenu.Name = "titleAlignmentMenu";
            titleAlignmentMenu.ShowLabel = false; titleAlignmentMenu.ShowImage = true;
            for (int i = 0; i < titleAlignmentIcons.Length; i++)
            {
                int index = i;
                // Use ordinary buttons for menu items. ToggleButton.Checked draws
                // the unwanted selection frame shown in the menu popup; the
                // current choice is represented solely by the menu's top icon.
                var choice = Factory.CreateRibbonButton();
                choice.Label = titleAlignmentLabels[i]; choice.OfficeImageId = titleAlignmentIcons[i]; choice.ShowImage = true;
                choice.Click += (s, e) => { SetTitleAlignment(index); SaveSettings(s, e); };
                titleAlignmentMenu.Items.Add(choice);
            }
            SetTitleAlignment(1);
            var options = Factory.CreateRibbonBox(); options.BoxStyle = RibbonBoxStyle.Horizontal;
            options.Name = "titleFormattingRow";
            fontSizeEditBox.ScreenTip = "标题字号 (pt)";
            titleFontSizeMenu = Factory.CreateRibbonMenu();
            titleFontSizeMenu.Name = "titleFontSizeMenu";
            titleFontSizeMenu.Label = "预设字号";
            titleFontSizeMenu.ShowLabel = false;
            titleFontSizeMenu.ShowImage = false;
            titleFontSizeMenu.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            titleFontSizeMenu.ScreenTip = "选择预设字号";
            options.Items.Add(fontSizeEditBox);
            options.Items.Add(titleFontSizeMenu);
            options.Items.Add(autoGroupCheckBox);
            options.Items.Add(titleAlignmentMenu);
            format.Items.Add(options);
            PopulateTitleFontSizePresets(TitleFontSizePresets);
        }

        private void PopulateTitleFontSizePresets(IEnumerable<string> values)
        {
            titleFontSizeMenu.Items.Clear();
            foreach (string value in values)
            {
                var choice = Factory.CreateRibbonButton();
                choice.Label = value;
                choice.Click += (sender, args) =>
                {
                    fontSizeEditBox.Text = value;
                    SaveSettings(sender, args);
                };
                titleFontSizeMenu.Items.Add(choice);
            }
        }

        private RibbonButton CreateSideTitleButton(string label, PictureTitleSide side)
        {
            var button = Factory.CreateRibbonButton(); button.Label = label;
            button.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            button.ShowImage = true;
            button.ShowLabel = true;
            button.Image = side == PictureTitleSide.Left ? 图片上标题.Image : AddTitleButton.Image;
            button.ScreenTip = "在图片" + (side == PictureTitleSide.Left ? "左" : "右") + "侧添加横排标题";
            button.Click += (s, e) => AddPictureTitles(side);
            return button;
        }

        private void SetTitleAlignment(int index)
        {
            titleAlignmentIndex = index >= 0 && index < 4 ? index : 1;
            titleAlignmentMenu.OfficeImageId = titleAlignmentIcons[titleAlignmentIndex];
            titleAlignmentMenu.Label = titleAlignmentLabels[titleAlignmentIndex];
            titleAlignmentMenu.ScreenTip = "标题对齐：" + titleAlignmentLabels[titleAlignmentIndex];
        }

        private RibbonSeparator CreateTitleColumnSeparator(string name)
        {
            var separator = Factory.CreateRibbonSeparator();
            separator.Name = name;
            return separator;
        }

        private void AddPictureTitles(PictureTitleSide side)
        {
            if (!TryGetActiveSlide(out P.Slide slide) || !PowerPointContext.TryGetActiveSelection(app, out P.Selection selection)
                || selection.Type != P.PpSelectionType.ppSelectionShapes)
            { MessageBox.Show("请选择要添加标题的图片或对象。", "添加图片标题"); return; }
            if (!TryParseFloat(fontSizeEditBox.Text, out float size) || !TitleFinite(size) || size <= 0 || size > 4000 ||
                !TryParseFloat(titleOffsetXCombo.Text, out float x) || !TitleFinite(x) ||
                !TryParseFloat(distanceFromBottomEditBox.Text, out float y) || !TitleFinite(y))
            { MessageBox.Show("请输入有效的字号和垂直/水平偏移量 (pt)。", "添加图片标题"); return; }
            var sources = new List<P.Shape>();
            foreach (P.Shape shape in (GetSortedSelection(selection, 10f) ?? selection.ShapeRange)) sources.Add(shape);
            var results = new List<P.Shape>();
            var errors = new List<string>();
            try { app.StartNewUndoEntry(); } catch { }
            foreach (var source in sources)
            {
                try
                {
                    var title = PictureTitleHelper.Add(slide, source, side, titleTextEditBox.Text, fontNameEditBox.Text, size, x, y,
                        (P.PpParagraphAlignment)(titleAlignmentIndex + 1));
                    if (autoGroupCheckBox.Checked)
                    {
                        try { results.Add(slide.Shapes.Range(new[] { source.Name, title.Name }).Group()); }
                        catch (Exception ex) { results.Add(title); errors.Add(source.Name + "：编组失败，图片和标题已保留。" + ex.Message); }
                    }
                    else results.Add(title);
                }
                catch (Exception ex) { errors.Add(source.Name + "：" + ex.Message); }
            }
            if (results.Count > 0) SelectMultipleShapes(results);
            if (errors.Count > 0) MessageBox.Show(string.Join(Environment.NewLine, errors), "添加图片标题");
        }
        private static bool TitleFinite(float value) { return !float.IsNaN(value) && !float.IsInfinity(value); }
    }
}
