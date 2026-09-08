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
            distanceFromBottomEditBox.Label = "上下偏移";
            distanceFromBottomEditBox.SizeString = TitleNumberSize;
            distanceFromBottomEditBox.ScreenTip = "上下偏移 (pt)";
            distanceFromBottomEditBox.SuperTip = "正值向下，负值向上；与左右偏移叠加，对新生成的四个方向标题都生效。";
            vertical.Items.Add(distanceFromBottomEditBox);
            图片处理.Items.Add(vertical);
            图片处理.Items.Add(CreateTitleColumnSeparator("titleVerticalSeparator"));

            var horizontal = Factory.CreateRibbonBox(); horizontal.BoxStyle = RibbonBoxStyle.Vertical;
            horizontal.Items.Add(CreateSideTitleButton("添加左标题", PictureTitleSide.Left));
            horizontal.Items.Add(CreateSideTitleButton("添加右标题", PictureTitleSide.Right));
            titleOffsetXCombo = Factory.CreateRibbonComboBox();
            titleOffsetXCombo.Name = "titleOffsetXCombo";
            titleOffsetXCombo.Label = "左右偏移"; titleOffsetXCombo.SizeString = TitleNumberSize;
            titleOffsetXCombo.ScreenTip = "左右偏移 (pt)";
            titleOffsetXCombo.SuperTip = "正值向右，负值向左；与上下偏移叠加，对新生成的四个方向标题都生效。";
            horizontal.Items.Add(titleOffsetXCombo);
            图片处理.Items.Add(horizontal);
            图片处理.Items.Add(CreateTitleColumnSeparator("titleHorizontalSeparator"));

            // The third column is one vertical layout item. Its three rows are
            // 标题, 字体, and the horizontal row 字号+编组+对齐方式. The horizontal
            // row stays a direct child of this column (single nesting level):
            // wrapping it in another vertical slot makes PowerPoint render it as
            // a new column, turning the group into four columns.
            titleTextEditBox.SizeString = TitleWideSize;
            fontNameEditBox.Label = "字体"; fontNameEditBox.SizeString = TitleWideSize;
            var format = Factory.CreateRibbonBox(); format.BoxStyle = RibbonBoxStyle.Vertical;
            format.Name = "titleFormatColumn";
            format.Items.Add(titleTextEditBox);
            format.Items.Add(fontNameEditBox);
            var row = Factory.CreateRibbonBox(); row.BoxStyle = RibbonBoxStyle.Horizontal;
            row.Name = "titleFormattingRow";
            fontSizeEditBox.Label = "字号"; fontSizeEditBox.SizeString = TitleNumberSize;
            autoGroupCheckBox.ControlSize = Office.RibbonControlSize.RibbonControlSizeRegular;
            autoGroupCheckBox.ShowImage = false;
            autoGroupCheckBox.ShowLabel = true;
            row.Items.Add(fontSizeEditBox); row.Items.Add(autoGroupCheckBox);
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
            row.Items.Add(titleAlignmentMenu);
            format.Items.Add(row);
            图片处理.Items.Add(format);
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
            { MessageBox.Show("请输入有效的字号和上下/左右偏移量 (pt)。", "添加图片标题"); return; }
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
