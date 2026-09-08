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
            distanceFromBottomEditBox.SizeString = "000000";
            distanceFromBottomEditBox.ScreenTip = "上下偏移 (pt)";
            distanceFromBottomEditBox.SuperTip = "正值向下，负值向上；与左右偏移叠加，对新生成的四个方向标题都生效。";
            vertical.Items.Add(distanceFromBottomEditBox);
            图片处理.Items.Add(vertical);
            图片处理.Items.Add(CreateTitleColumnSeparator("上、下标题与偏移"));

            var horizontal = Factory.CreateRibbonBox(); horizontal.BoxStyle = RibbonBoxStyle.Vertical;
            horizontal.Items.Add(CreateSideTitleButton("添加左标题", PictureTitleSide.Left));
            horizontal.Items.Add(CreateSideTitleButton("添加右标题", PictureTitleSide.Right));
            titleOffsetXCombo = Factory.CreateRibbonComboBox();
            titleOffsetXCombo.Name = "titleOffsetXCombo";
            titleOffsetXCombo.Label = "左右偏移"; titleOffsetXCombo.SizeString = "00";
            titleOffsetXCombo.ScreenTip = "左右偏移 (pt)";
            titleOffsetXCombo.SuperTip = "正值向右，负值向左；与上下偏移叠加，对新生成的四个方向标题都生效。";
            horizontal.Items.Add(titleOffsetXCombo);
            图片处理.Items.Add(horizontal);
            图片处理.Items.Add(CreateTitleColumnSeparator("左、右标题与偏移"));

            var format = Factory.CreateRibbonBox(); format.BoxStyle = RibbonBoxStyle.Vertical;
            // ComboBox has a built-in drop-down affordance. Give the plain title
            // edit box the equivalent extra character width so their right edges
            // line up in the Ribbon.
            titleTextEditBox.SizeString = "0000000000";
            fontNameEditBox.Label = "字体"; fontNameEditBox.SizeString = "00000000";
            format.Items.Add(titleTextEditBox); format.Items.Add(fontNameEditBox);
            var row = Factory.CreateRibbonBox(); row.BoxStyle = RibbonBoxStyle.Horizontal;
            fontSizeEditBox.Label = "字号"; fontSizeEditBox.SizeString = "000";
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
            // Keep the compact third row as its own vertical layout item. Office
            // then applies the same inter-row baseline spacing used by the other
            // stacked ComboBoxes instead of abutting it against the font box.
            var rowSlot = Factory.CreateRibbonBox(); rowSlot.BoxStyle = RibbonBoxStyle.Vertical;
            rowSlot.Items.Add(row);
            format.Items.Add(rowSlot);
            图片处理.Items.Add(format);
        }

        private RibbonButton CreateSideTitleButton(string label, PictureTitleSide side)
        {
            var button = Factory.CreateRibbonButton(); button.Label = label;
            button.ShowImage = true; button.OfficeImageId = "TextBoxInsert";
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

        private RibbonSeparator CreateTitleColumnSeparator(string title)
        {
            var separator = Factory.CreateRibbonSeparator();
            separator.Name = "titleSeparator_" + title.GetHashCode().ToString("X8");
            separator.Title = title;
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
