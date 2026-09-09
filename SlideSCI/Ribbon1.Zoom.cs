using System;
using System.Globalization;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        // Reserve the displayed unit as well as two digits (90% / 10x).
        private const string ZoomNumberSize = "00%";
        private ZoomSettings zoomRibbonSettings;
        private RibbonComboBox zoomSizeCombo, zoomGapCombo;
        private RibbonToggleButton zoomGroupCheck;
        private RibbonMenu zoomConnectionMenu, zoomLineMenu, zoomBoxMenu;

        private void InitializeZoomRibbon()
        {
            zoomRibbonSettings = ZoomSettings.LoadOrDefault();
            // Legacy physical-width modes are retained in the settings class for
            // old documents; the Ribbon now consistently edits magnification.
            if (zoomRibbonSettings.TargetMode != ZoomTargetMode.Multiple)
                zoomRibbonSettings.Magnification = 1f;
            zoomRibbonSettings.TargetMode = ZoomTargetMode.Multiple;
            zoomGroup.Items.Clear();

            // 第一列：选区（生成选区框）→ 尺寸（选区框大小）→ 框线（选区框线条样式）。
            btnInsertZoomBox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            btnInsertZoomBox.Label = "框选区域";
            btnInsertZoomBox.ScreenTip = "为选中的图片添加局部放大选区框";
            btnInsertZoomBox.SuperTip = "选中一张图片后插入选区框。可在同一页创建多个选区，拖动或缩放到需要强调的区域。";
            var selection = Factory.CreateRibbonBox();
            selection.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            selection.Items.Add(btnInsertZoomBox);
            zoomBoxPercentCombo.Label = "尺寸";
            zoomBoxPercentCombo.SizeString = ZoomNumberSize;
            zoomBoxPercentCombo.Text = zoomRibbonSettings.BoxPercent.ToString("0.##", CultureInfo.CurrentCulture) + "%";
            zoomBoxPercentCombo.ScreenTip = "选区框边长占原图短边的百分比（5–90）";
            zoomBoxPercentCombo.SuperTip = "控制选区框大小；也可在幻灯片上拖动或缩放选区框。";
            foreach (string value in new[] { "10%", "20%", "25%", "30%", "40%", "50%", "60%", "75%", "90%" }) AddZoomItem(zoomBoxPercentCombo, value);
            selection.Items.Add(zoomBoxPercentCombo);
            zoomBoxMenu = CreateZoomStrokeMenu("框线", true);
            WrapRibbonRows(selection, "zoomSelection");
            selection.Items.Add(CreateZoomStrokeRow("zoomBoxStrokeRow", "框线", zoomBoxMenu));
            zoomGroup.Items.Add(selection);

            // 第二列：放大（生成放大图）→ 尺寸（放大图大小）→ 引线（引导线线条样式）。
            btnGenerateZoomInset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            btnGenerateZoomInset.Label = "放大选区";
            btnGenerateZoomInset.ScreenTip = "按本分组设置生成或更新放大图";
            btnGenerateZoomInset.SuperTip = "先选择图片并点击“框选区域”，拖动或缩放选区框，再点击“放大选区”。更新时选择选区框、放大图或连线。原图不参与编组。";
            var generate = Factory.CreateRibbonBox();
            generate.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            generate.Items.Add(btnGenerateZoomInset);
            zoomSizeCombo = Factory.CreateRibbonComboBox();
            zoomSizeCombo.Label = "尺寸";
            zoomSizeCombo.SizeString = ZoomNumberSize;
            zoomSizeCombo.ScreenTip = "放大图尺寸";
            zoomSizeCombo.SuperTip = "输入放大倍数，如 2 或 2x；默认 1x，始终保持选区宽高比。";
            foreach (string value in new[] { "1x", "2x", "3x", "4x", "5x", "10x" }) AddZoomItem(zoomSizeCombo, value);
            zoomSizeCombo.Text = zoomRibbonSettings.Magnification.ToString("0.##", CultureInfo.CurrentCulture) + "x";
            generate.Items.Add(zoomSizeCombo);
            zoomLineMenu = CreateZoomStrokeMenu("连线", false);
            WrapRibbonRows(generate, "zoomGenerate");
            generate.Items.Add(CreateZoomStrokeRow("zoomLineStrokeRow", "连线", zoomLineMenu));
            zoomGroup.Items.Add(generate);

            // 第三列：间距（原图与放大图距离）→ 连线（是否引线）→ 编组（开关）。
            var options = Factory.CreateRibbonBox();
            options.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            zoomGapCombo = Factory.CreateRibbonComboBox();
            zoomGapCombo.Label = "间距";
            zoomGapCombo.SizeString = "000";
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
            WrapRibbonRows(options, "zoomOptions");
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
            string[] dashes = StrokeDashNames;
            zoomLineMenu.Label = FormatZoomWeight(zoomRibbonSettings.LineWeight);
            zoomBoxMenu.Label = FormatZoomWeight(zoomRibbonSettings.BoxLineWeight);
            zoomLineMenu.SuperTip = (zoomRibbonSettings.LineVisible ? dashes[zoomRibbonSettings.LineDash] : "无轮廓") + "；颜色 " + ColorTranslator.ToHtml(ColorTranslator.FromOle(zoomRibbonSettings.LineColorRgb)) + "；粗细 " + FormatZoomWeight(zoomRibbonSettings.LineWeight) + " pt。生成时应用。";
            zoomBoxMenu.SuperTip = (zoomRibbonSettings.BoxLineVisible ? dashes[zoomRibbonSettings.BoxLineDash] : "无轮廓") + "；颜色 " + ColorTranslator.ToHtml(ColorTranslator.FromOle(zoomRibbonSettings.BoxColorRgb)) + "；粗细 " + FormatZoomWeight(zoomRibbonSettings.BoxLineWeight) + " pt。创建或更新时应用。";
        }

        private RibbonBox CreateZoomStrokeRow(string name, string text, RibbonMenu menu)
        {
            var label = Factory.CreateRibbonLabel();
            label.Name = name + "Label";
            label.Label = text;
            menu.ShowImage = false;
            menu.ShowLabel = true;
            return CreateRibbonRow(name, label, menu);
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
            if (!ZoomSettings.TryParsePercent(zoomBoxPercentCombo.Text, out float percent))
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

        private ZoomSettings ReadZoomRibbonSettings()
        {
            if (!ZoomSettings.TryParsePercent(zoomBoxPercentCombo.Text, out float percent))
            { MessageBox.Show("选区大小请输入 5–90。", "局部放大"); return null; }
            if (!TryParseFloat(zoomGapCombo.Text, out float gap) || float.IsNaN(gap) || float.IsInfinity(gap) || gap < 0 || gap > ZoomInsetHelper.CmToPoints(50))
            { MessageBox.Show("图间距请输入 0–1417 pt。", "局部放大"); return null; }
            if (!ZoomSettings.TryParseMagnification(zoomSizeCombo.Text, out float value))
            { MessageBox.Show("尺寸请输入 0.1–100 的倍数，可省略 x。", "局部放大"); return null; }
            zoomRibbonSettings.TargetMode = ZoomTargetMode.Multiple;
            zoomRibbonSettings.Magnification = value;
            zoomRibbonSettings.BoxPercent = percent;
            zoomRibbonSettings.GapCm = gap / ZoomInsetHelper.CmToPoints(1);
            zoomRibbonSettings.Group = zoomGroupCheck.Checked;
            return zoomRibbonSettings.NormalizedCopy();
        }
    }
}
