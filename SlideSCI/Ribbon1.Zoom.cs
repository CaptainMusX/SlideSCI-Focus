using System;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private const string ZoomNumberSize = "0000";
        // ComboBox 标签比普通菜单/切换按钮的文字少一个原生左边距。
        // 不换行空格仅用于显示，不能写入数值 Text 或解析输入。
        private const string ZoomLabelInset = "\u00A0";
        // 第一、二列的 ComboBox 与同列菜单的原生边距还差约 1–2 物理像素；
        // 加一个更细的 hair space（U+200A）做视觉微调，第三列原本已对齐故不加。
        private const string ZoomLabelFineInset = "\u200A";
        private ZoomSettings zoomRibbonSettings;
        private RibbonComboBox zoomSizeCombo, zoomGapCombo;
        private RibbonToggleButton zoomGroupCheck;
        private RibbonMenu zoomConnectionMenu, zoomLineMenu, zoomBoxMenu;

        private void InitializeZoomRibbon()
        {
            zoomRibbonSettings = ZoomSettings.LoadOrDefault();
            zoomGroup.Items.Clear();

            // 第一列：选区（生成选区框）→ 尺寸（选区框大小）→ 框线（选区框线条样式）。
            btnInsertZoomBox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            btnInsertZoomBox.Label = "选区";
            btnInsertZoomBox.ScreenTip = "为选中的图片添加局部放大选区框";
            btnInsertZoomBox.SuperTip = "选中一张图片后插入选区框。可在同一页创建多个选区，拖动或缩放到需要强调的区域。";
            var selection = Factory.CreateRibbonBox();
            selection.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            selection.Items.Add(btnInsertZoomBox);
            zoomBoxPercentCombo.Label = ZoomLabelInset + ZoomLabelFineInset + "尺寸";
            zoomBoxPercentCombo.SizeString = ZoomNumberSize;
            zoomBoxPercentCombo.Text = zoomRibbonSettings.BoxPercent.ToString(CultureInfo.CurrentCulture);
            zoomBoxPercentCombo.ScreenTip = "选区框边长占原图短边的百分比（5–90）";
            zoomBoxPercentCombo.SuperTip = "控制选区框大小；也可在幻灯片上拖动或缩放选区框。";
            foreach (string value in new[] { "10", "20", "25", "30", "40", "50", "60", "75", "90" }) AddZoomItem(zoomBoxPercentCombo, value);
            selection.Items.Add(zoomBoxPercentCombo);
            zoomBoxMenu = CreateZoomStrokeMenu("框线", true);
            selection.Items.Add(zoomBoxMenu);
            zoomGroup.Items.Add(selection);

            // 第二列：放大（生成放大图）→ 尺寸（放大图大小）→ 引线（引导线线条样式）。
            btnGenerateZoomInset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            btnGenerateZoomInset.Label = "放大";
            btnGenerateZoomInset.ScreenTip = "按本分组设置生成或更新放大图";
            btnGenerateZoomInset.SuperTip = "先选择图片并点击“选区”，拖动或缩放选区框，再点击“放大”。更新时选择选区框、放大图或引导线。原图不参与编组。";
            var generate = Factory.CreateRibbonBox();
            generate.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            generate.Items.Add(btnGenerateZoomInset);
            zoomSizeCombo = Factory.CreateRibbonComboBox();
            zoomSizeCombo.Label = ZoomLabelInset + ZoomLabelFineInset + "尺寸";
            zoomSizeCombo.SizeString = "原图等宽";
            zoomSizeCombo.ScreenTip = "放大图尺寸";
            zoomSizeCombo.SuperTip = "输入“原图等宽”、放大倍数（如 2x）或宽度（如 5cm）。始终保持选区宽高比。";
            foreach (string value in new[] { "原图等宽", "2x", "3x", "4x", "5x", "3cm", "5cm", "8cm" }) AddZoomItem(zoomSizeCombo, value);
            zoomSizeCombo.Text = zoomRibbonSettings.TargetMode == ZoomTargetMode.SameAsOriginal ? "原图等宽" :
                zoomRibbonSettings.TargetMode == ZoomTargetMode.Multiple ? zoomRibbonSettings.Magnification + "x" : zoomRibbonSettings.CustomWidthCm + "cm";
            generate.Items.Add(zoomSizeCombo);
            zoomLineMenu = CreateZoomStrokeMenu("引线", false);
            generate.Items.Add(zoomLineMenu);
            zoomGroup.Items.Add(generate);

            // 第三列：间距（原图与放大图距离）→ 连线（是否引线）→ 编组（开关）。
            var options = Factory.CreateRibbonBox();
            options.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            zoomGapCombo = Factory.CreateRibbonComboBox();
            zoomGapCombo.Label = ZoomLabelInset + "间距";
            zoomGapCombo.SizeString = ZoomNumberSize;
            zoomGapCombo.Text = ZoomInsetHelper.CmToPoints(zoomRibbonSettings.GapCm).ToString("0.##", CultureInfo.CurrentCulture);
            zoomGapCombo.ScreenTip = "原图与放大图之间的距离 (pt)";
            zoomGapCombo.SuperTip = "单位 pt，与图片自动排列的“列间距”一致；0 表示紧贴。";
            foreach (string value in new[] { "0", "5", "10", "15", "20", "30" }) AddZoomItem(zoomGapCombo, value);
            options.Items.Add(zoomGapCombo);
            zoomConnectionMenu = Factory.CreateRibbonMenu();
            zoomConnectionMenu.Label = "连线";
            zoomConnectionMenu.SuperTip = "选择放大图与选区框之间的引线方式。";
            AddZoomButton(zoomConnectionMenu, "平行引线", () => zoomRibbonSettings.LineStyle = ZoomLineStyle.JournalFunnel);
            AddZoomButton(zoomConnectionMenu, "交叉引线", () => zoomRibbonSettings.LineStyle = ZoomLineStyle.CrossedX);
            AddZoomButton(zoomConnectionMenu, "不连线", () => zoomRibbonSettings.LineStyle = ZoomLineStyle.None);
            options.Items.Add(zoomConnectionMenu);
            zoomGroupCheck = Factory.CreateRibbonToggleButton();
            zoomGroupCheck.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            zoomGroupCheck.Label = "编组";
            zoomGroupCheck.ShowImage = false;
            zoomGroupCheck.ShowLabel = true;
            zoomGroupCheck.Checked = zoomRibbonSettings.Group;
            zoomGroupCheck.SuperTip = "仅编组选区框、放大图和引导线；不包含原图或辅助形状。";
            zoomGroupCheck.Click += (s, e) => { zoomRibbonSettings.Group = zoomGroupCheck.Checked; ZoomSettings.Save(zoomRibbonSettings); };
            options.Items.Add(zoomGroupCheck);
            zoomGroup.Items.Add(options);

            RefreshZoomLabels();
            zoomBoxPercentCombo.TextChanged += ZoomPercentChanged;
            zoomSizeCombo.TextChanged += SaveZoomRibbonInputs;
            zoomGapCombo.TextChanged += SaveZoomRibbonInputs;
        }

        private void RefreshZoomLabels()
        {
            if (zoomConnectionMenu == null || zoomLineMenu == null || zoomBoxMenu == null) return;
            // 标签宽度保持恒定：切换选项时 Ribbon 不会重新排版，整组不会跳动。
            zoomConnectionMenu.Label = zoomRibbonSettings.LineStyle == ZoomLineStyle.None ? "连线 关闭" :
                zoomRibbonSettings.LineStyle == ZoomLineStyle.CrossedX ? "连线 交叉" : "连线 平行";
            string[] dashes = { "实线", "虚线", "点线", "点划线" };
            zoomLineMenu.Label = "引线 " + FormatZoomWeight(zoomRibbonSettings.LineWeight) + "pt";
            // 第一列菜单的“框线”与同列 ComboBox 的“尺寸”相差约 1 个物理像素；
            // 给菜单加一个不可见 hair space，使两行文字左端落在同一视觉基线上。
            zoomBoxMenu.Label = ZoomLabelFineInset + "框线 " + FormatZoomWeight(zoomRibbonSettings.BoxLineWeight) + "pt";
            zoomLineMenu.SuperTip = dashes[zoomRibbonSettings.LineDash] + "；颜色 " + ColorTranslator.ToHtml(ColorTranslator.FromOle(zoomRibbonSettings.LineColorRgb)) + "。点击设置颜色、粗细和线型；生成时应用。";
            zoomBoxMenu.SuperTip = dashes[zoomRibbonSettings.BoxLineDash] + "；颜色 " + ColorTranslator.ToHtml(ColorTranslator.FromOle(zoomRibbonSettings.BoxColorRgb)) + "。点击设置颜色、粗细和线型。";
        }

        /// <summary>固定两位小数，保证“框线/引线”菜单标签宽度一致。</summary>
        private static string FormatZoomWeight(float weight)
        {
            return weight.ToString("0.00", CultureInfo.InvariantCulture);
        }

        private void SaveZoomRibbonInputs(object sender, RibbonControlEventArgs e)
        {
            var settings = ReadZoomRibbonSettings();
            if (settings != null) ZoomSettings.Save(settings);
        }

        private void ZoomPercentChanged(object sender, RibbonControlEventArgs e)
        {
            if (!TryParseFloat(zoomBoxPercentCombo.Text, out float percent) || float.IsNaN(percent) || float.IsInfinity(percent) || percent < 5 || percent > 90)
            { MessageBox.Show("选区大小请输入 5–90。", "局部放大"); return; }
            zoomRibbonSettings.BoxPercent = percent;
            ZoomSettings.Save(zoomRibbonSettings);
            try
            {
                if (!TryGetActiveSlide(out Microsoft.Office.Interop.PowerPoint.Slide slide)) return;
                if (!PowerPointContext.TryGetActiveSelection(app, out Microsoft.Office.Interop.PowerPoint.Selection selection) ||
                    selection.Type != Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count != 1) return;
                var box = selection.ShapeRange[1];
                if (!ZoomInsetHelper.IsZoomBox(box)) return;
                var source = ZoomInsetHelper.FindSourcePicture(slide, box);
                if (source == null) return;
                float side = Math.Max(6f, Math.Min(source.Width, source.Height) * percent / 100f);
                float cx = box.Left + box.Width / 2, cy = box.Top + box.Height / 2;
                app.StartNewUndoEntry();
                box.Width = side; box.Height = side;
                box.Left = Math.Max(source.Left, Math.Min(cx - side / 2, source.Left + source.Width - side));
                box.Top = Math.Max(source.Top, Math.Min(cy - side / 2, source.Top + source.Height - side));
            }
            catch (Exception ex) { System.Diagnostics.Trace.TraceWarning("Unable to resize selection: {0}", ex.Message); }
        }

        private void AddZoomItem(RibbonComboBox combo, string value)
        {
            var item = Factory.CreateRibbonDropDownItem(); item.Label = value; combo.Items.Add(item);
        }

        private void AddZoomButton(RibbonMenu menu, string label, Action action)
        {
            var button = Factory.CreateRibbonButton(); button.Label = label;
            button.Click += (s, e) => { action(); RefreshZoomLabels(); ZoomSettings.Save(zoomRibbonSettings); };
            menu.Items.Add(button);
        }

        private RibbonMenu CreateZoomStrokeMenu(string label, bool box)
        {
            var menu = Factory.CreateRibbonMenu(); menu.Label = label;
            var colors = Factory.CreateRibbonMenu(); colors.Label = "颜色";
            string[] names = { "黑色", "白色", "红色", "蓝色", "绿色", "黄色" };
            Color[] values = { Color.Black, Color.White, Color.Red, Color.Blue, Color.Green, Color.Yellow };
            for (int i = 0; i < names.Length; i++)
            {
                Color color = values[i];
                AddZoomButton(colors, names[i], () => SetZoomColor(box, ColorTranslator.ToOle(color)));
            }
            AddZoomButton(colors, "其他颜色…", () =>
            {
                using (var dialog = new ColorDialog { FullOpen = true, Color = ColorTranslator.FromOle(box ? zoomRibbonSettings.BoxColorRgb : zoomRibbonSettings.LineColorRgb) })
                {
                    if (dialog.ShowDialog() == DialogResult.OK) SetZoomColor(box, ColorTranslator.ToOle(dialog.Color));
                }
            });
            menu.Items.Add(colors);
            var weights = Factory.CreateRibbonMenu(); weights.Label = "粗细 (pt)";
            foreach (float value in box ? new[] { 0.5f, 0.75f, 1f, 1.5f, 2f, 3f } : new[] { 0.5f, 0.75f, 1f, 1.5f, 2f })
            {
                float weight = value;
                AddZoomButton(weights, value.ToString(CultureInfo.CurrentCulture), () => { if (box) zoomRibbonSettings.BoxLineWeight = weight; else zoomRibbonSettings.LineWeight = weight; });
            }
            menu.Items.Add(weights);
            var dashes = Factory.CreateRibbonMenu(); dashes.Label = "线型";
            string[] dashNames = { "实线", "虚线", "点线", "点划线" };
            for (int i = 0; i < dashNames.Length; i++)
            {
                int dash = i;
                AddZoomButton(dashes, dashNames[i], () => { if (box) zoomRibbonSettings.BoxLineDash = dash; else zoomRibbonSettings.LineDash = dash; });
            }
            menu.Items.Add(dashes);
            return menu;
        }

        private void SetZoomColor(bool box, int color)
        {
            if (box) zoomRibbonSettings.BoxColorRgb = color; else zoomRibbonSettings.LineColorRgb = color;
        }

        private ZoomSettings ReadZoomRibbonSettings()
        {
            if (!TryParseFloat(zoomBoxPercentCombo.Text, out float percent) || float.IsNaN(percent) || float.IsInfinity(percent) || percent < 5 || percent > 90)
            { MessageBox.Show("选区大小请输入 5–90。", "局部放大"); return null; }
            if (!TryParseFloat(zoomGapCombo.Text, out float gap) || float.IsNaN(gap) || float.IsInfinity(gap) || gap < 0 || gap > ZoomInsetHelper.CmToPoints(50))
            { MessageBox.Show("图间距请输入 0–1417 pt。", "局部放大"); return null; }
            string size = (zoomSizeCombo.Text ?? "").Trim().ToLowerInvariant();
            if (size == "原图等宽") zoomRibbonSettings.TargetMode = ZoomTargetMode.SameAsOriginal;
            else
            {
                bool cm = size.EndsWith("cm", StringComparison.Ordinal);
                string number = cm ? size.Substring(0, size.Length - 2) : size.TrimEnd('x', '×');
                if (!TryParseFloat(number, out float value) || float.IsNaN(value) || float.IsInfinity(value) || value < 0.1f || value > (cm ? 200 : 100))
                { MessageBox.Show("尺寸请输入“原图等宽”、0.1–100x 或 0.1–200cm。", "局部放大"); return null; }
                zoomRibbonSettings.TargetMode = cm ? ZoomTargetMode.CustomWidthCm : ZoomTargetMode.Multiple;
                if (cm) zoomRibbonSettings.CustomWidthCm = value; else zoomRibbonSettings.Magnification = value;
            }
            zoomRibbonSettings.BoxPercent = percent;
            zoomRibbonSettings.GapCm = gap / ZoomInsetHelper.CmToPoints(1);
            zoomRibbonSettings.Group = zoomGroupCheck.Checked;
            return zoomRibbonSettings.NormalizedCopy();
        }
    }
}
