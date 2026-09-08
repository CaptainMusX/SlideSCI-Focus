using System;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private ZoomSettings zoomRibbonSettings;
        private RibbonComboBox zoomSizeCombo, zoomGapCombo;
        private RibbonCheckBox zoomGroupCheck;
        private RibbonMenu zoomConnectionMenu, zoomLineMenu, zoomBoxMenu;

        private void InitializeZoomRibbon()
        {
            zoomRibbonSettings = ZoomSettings.LoadOrDefault();
            zoomGroup.Items.Clear();
            btnInsertZoomBox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            btnInsertZoomBox.Label = "选择区域";
            btnGenerateZoomInset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            btnGenerateZoomInset.ScreenTip = "按本分组设置生成或更新放大图";
            btnGenerateZoomInset.SuperTip = "先选择图片并点击“选择区域”，拖动或缩放选区框，再点击“生成放大图”。更新时选择选区框、放大图或引导线。原图不参与编组。";
            var actions = Factory.CreateRibbonBox();
            actions.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            actions.Items.Add(btnInsertZoomBox);
            actions.Items.Add(btnGenerateZoomInset);
            zoomGroupCheck = Factory.CreateRibbonCheckBox();
            zoomGroupCheck.Label = "编组";
            zoomGroupCheck.Checked = zoomRibbonSettings.Group;
            zoomGroupCheck.SuperTip = "仅编组选区框、放大图和引导线；不包含原图或辅助形状。";
            zoomGroupCheck.Click += (s, e) => { zoomRibbonSettings.Group = zoomGroupCheck.Checked; ZoomSettings.Save(zoomRibbonSettings); };
            actions.Items.Add(zoomGroupCheck);
            zoomGroup.Items.Add(actions);

            var sizes = Factory.CreateRibbonBox();
            sizes.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            zoomBoxPercentCombo.Label = "选区 %";
            zoomBoxPercentCombo.SizeString = "原图等宽";
            zoomBoxPercentCombo.Text = zoomRibbonSettings.BoxPercent.ToString(CultureInfo.CurrentCulture);
            zoomBoxPercentCombo.ScreenTip = "边长占原图短边的百分比（5–90）；也可在幻灯片上拖动调整";
            foreach (string value in new[] { "10", "20", "25", "30", "40", "50", "60", "75", "90" }) AddZoomItem(zoomBoxPercentCombo, value);
            sizes.Items.Add(zoomBoxPercentCombo);
            zoomSizeCombo = Factory.CreateRibbonComboBox();
            zoomSizeCombo.Label = "尺寸";
            zoomSizeCombo.SizeString = "原图等宽";
            zoomSizeCombo.SuperTip = "输入“原图等宽”、放大倍数（如 2x）或宽度（如 5cm）。始终保持选区宽高比。";
            foreach (string value in new[] { "原图等宽", "2x", "3x", "4x", "5x", "3cm", "5cm", "8cm" }) AddZoomItem(zoomSizeCombo, value);
            zoomSizeCombo.Text = zoomRibbonSettings.TargetMode == ZoomTargetMode.SameAsOriginal ? "原图等宽" :
                zoomRibbonSettings.TargetMode == ZoomTargetMode.Multiple ? zoomRibbonSettings.Magnification + "x" : zoomRibbonSettings.CustomWidthCm + "cm";
            sizes.Items.Add(zoomSizeCombo);
            zoomGapCombo = Factory.CreateRibbonComboBox();
            zoomGapCombo.Label = "间距 pt";
            zoomGapCombo.SizeString = "原图等宽";
            zoomGapCombo.Text = ZoomInsetHelper.CmToPoints(zoomRibbonSettings.GapCm).ToString("0.##", CultureInfo.CurrentCulture);
            zoomGapCombo.SuperTip = "原图与放大图之间的距离；单位 pt，与图片排列的“列间距”一致。";
            foreach (string value in new[] { "0", "5", "10", "15", "20", "30" }) AddZoomItem(zoomGapCombo, value);
            sizes.Items.Add(zoomGapCombo);
            zoomGroup.Items.Add(sizes);

            var styles = Factory.CreateRibbonBox();
            styles.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            var connection = Factory.CreateRibbonMenu();
            zoomConnectionMenu = connection;
            connection.Label = "连线方式";
            AddZoomButton(connection, "平行引线", () => zoomRibbonSettings.LineStyle = ZoomLineStyle.JournalFunnel);
            AddZoomButton(connection, "交叉引线", () => zoomRibbonSettings.LineStyle = ZoomLineStyle.CrossedX);
            AddZoomButton(connection, "不连线", () => zoomRibbonSettings.LineStyle = ZoomLineStyle.None);
            styles.Items.Add(connection);
            zoomLineMenu = CreateZoomStrokeMenu("连线样式", false);
            zoomBoxMenu = CreateZoomStrokeMenu("选区框样式", true);
            styles.Items.Add(zoomLineMenu);
            styles.Items.Add(zoomBoxMenu);
            zoomGroup.Items.Add(styles);
            RefreshZoomLabels();
            zoomBoxPercentCombo.TextChanged += ZoomPercentChanged;
            zoomSizeCombo.TextChanged += SaveZoomRibbonInputs;
            zoomGapCombo.TextChanged += SaveZoomRibbonInputs;
        }

        private void RefreshZoomLabels()
        {
            if (zoomConnectionMenu == null || zoomLineMenu == null || zoomBoxMenu == null) return;
            zoomConnectionMenu.Label = zoomRibbonSettings.LineStyle == ZoomLineStyle.None ? "连线：无" :
                zoomRibbonSettings.LineStyle == ZoomLineStyle.CrossedX ? "连线：交叉" : "连线：平行";
            string[] dashes = { "实线", "虚线", "点线", "点划线" };
            zoomLineMenu.Label = "引线 " + zoomRibbonSettings.LineWeight + "pt";
            zoomBoxMenu.Label = "选框 " + zoomRibbonSettings.BoxLineWeight + "pt";
            zoomLineMenu.SuperTip = dashes[zoomRibbonSettings.LineDash] + "；颜色 " + ColorTranslator.ToHtml(ColorTranslator.FromOle(zoomRibbonSettings.LineColorRgb)) + "。点击设置颜色、粗细和线型；生成时应用。";
            zoomBoxMenu.SuperTip = dashes[zoomRibbonSettings.BoxLineDash] + "；颜色 " + ColorTranslator.ToHtml(ColorTranslator.FromOle(zoomRibbonSettings.BoxColorRgb)) + "。点击设置颜色、粗细和线型。";
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
