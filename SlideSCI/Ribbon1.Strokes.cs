using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace SlideSCI
{
    public partial class Ribbon1
    {
        private static readonly string[] StrokeDashNames = { "实线", "短划线", "圆点", "短划线-点", "方点", "短划线-双点", "长划线", "长划线-点" };
        private static readonly float[] StrokeWeightPresets = { 0.25f, 0.5f, 0.75f, 1f, 1.5f, 2.25f, 3f, 4.5f, 6f };
        private static readonly string[] StrokeArrowNames = { "无箭头", "三角箭头", "开放箭头", "燕尾箭头", "菱形", "椭圆" };
        // 与 PowerPoint 原生“形状轮廓”菜单相同的 16pt 图标；32px 位图在高 DPI 下仍清晰。
        private const int StrokePreviewSize = 32;
        private readonly Dictionary<int, Bitmap> strokeSwatches = new Dictionary<int, Bitmap>();
        private readonly Dictionary<string, Bitmap> strokePreviews = new Dictionary<string, Bitmap>();

        private RibbonMenu CreateZoomStrokeMenu(string label, bool box)
        {
            var menu = Factory.CreateRibbonMenu();
            menu.Name = box ? "zoomBoxOutline" : "zoomLineOutline";
            menu.Label = label;
            menu.Dynamic = true;
            menu.ItemsLoading += (s, e) => PopulateStrokeMenu(menu, box);
            PopulateStrokeMenu(menu, box);
            return menu;
        }

        private void PopulateStrokeMenu(RibbonMenu menu, bool box)
        {
            while (menu.Items.Count > 0)
            {
                var control = menu.Items[0]; menu.Items.RemoveAt(0); control.Dispose();
            }
            var theme = GetStrokeThemeColors();
            var shades = new List<Color>(theme);
            foreach (float amount in new[] { 0.8f, 0.6f, 0.4f, -0.25f, -0.5f })
                shades.AddRange(theme.Select(c => TintStrokeColor(c, amount)));
            AddStrokeGallery(menu, "主题颜色", shades, box);
            AddStrokeGallery(menu, "标准色", new[] { Color.DarkRed, Color.Red, Color.Orange, Color.Yellow,
                Color.LightGreen, Color.Green, Color.Cyan, Color.DeepSkyBlue, Color.DarkBlue, Color.Purple }, box);
            var recent = ReadRecentStrokeColors();
            if (recent.Count > 0) AddStrokeGallery(menu, "最近使用的颜色", recent.Select(ColorTranslator.FromOle), box);
            menu.Items.Add(Factory.CreateRibbonSeparator());
            int strokeColor = box ? zoomRibbonSettings.BoxColorRgb : zoomRibbonSettings.LineColorRgb;

            // 与原生“形状轮廓”菜单逐项对应：无轮廓、其他轮廓颜色、取色器、粗细、草绘、虚线、箭头。
            var noOutline = AddZoomButton(menu, "无轮廓", () => { if (box) zoomRibbonSettings.BoxLineVisible = false; else zoomRibbonSettings.LineVisible = false; });
            noOutline.Image = GetStrokePreview("no-outline", strokeColor);
            noOutline.ShowImage = true;
            // 与原生菜单一致：Alt 打开菜单后可直接按键触发。
            noOutline.KeyTip = "N";

            var moreColors = AddZoomButton(menu, "其他轮廓颜色…", () =>
            {
                using (var dialog = new ColorDialog { FullOpen = true, Color = ColorTranslator.FromOle(box ? zoomRibbonSettings.BoxColorRgb : zoomRibbonSettings.LineColorRgb) })
                    if (dialog.ShowDialog() == DialogResult.OK) SetZoomColor(box, ColorTranslator.ToOle(dialog.Color));
            });
            // 原生菜单使用的调色盘图标（已在本机 PowerPoint 中核对可用）。
            moreColors.OfficeImageId = "ShapeOutlineColorPicker";
            moreColors.ShowImage = true;
            moreColors.KeyTip = "M";

            var eyedropper = AddZoomButton(menu, "取色器", () =>
            {
                // Defer until Office has dismissed its dropdown. This runs only
                // after the user explicitly chooses the eyedropper command.
                var timer = new Timer { Interval = 180 };
                timer.Tick += (s, e) =>
                {
                    timer.Stop(); timer.Dispose();
                    try
                    {
                        using (var picker = new StrokeColorPicker())
                            if (picker.ShowDialog() == DialogResult.OK)
                            {
                                SetZoomColor(box, ColorTranslator.ToOle(picker.SelectedColor));
                                RefreshZoomLabels(); ZoomSettings.Save(zoomRibbonSettings);
                            }
                    }
                    catch (Exception ex) { MessageBox.Show("取色失败：" + ex.Message, "局部放大"); }
                };
                timer.Start();
            });
            eyedropper.Image = GetStrokePreview("eyedropper", strokeColor);
            eyedropper.ShowImage = true;

            var weights = Factory.CreateRibbonMenu();
            weights.Label = "粗细"; weights.OfficeImageId = "LineStyle"; weights.ShowImage = true;
            foreach (float weight in StrokeWeightPresets)
            {
                float captured = weight;
                var item = AddZoomButton(weights, weight.ToString("0.##", CultureInfo.CurrentCulture) + " 磅", () =>
                {
                    if (box) { zoomRibbonSettings.BoxLineWeight = captured; zoomRibbonSettings.BoxLineVisible = true; }
                    else { zoomRibbonSettings.LineWeight = captured; zoomRibbonSettings.LineVisible = true; }
                });
                item.Image = GetStrokePreview("weight:" + weight.ToString(CultureInfo.InvariantCulture), strokeColor, weight);
                item.ShowImage = true;
            }
            menu.Items.Add(weights);

            // Matches the disabled Sketch entry shown in the reference. Office
            // exposes no portable sketch preset in the targeted VSTO line model.
            var sketch = Factory.CreateRibbonMenu(); sketch.Label = "草绘"; sketch.Enabled = false;
            sketch.OfficeImageId = "Scribble"; sketch.ShowImage = true;
            sketch.SuperTip = "当前线条预设不支持草绘。";
            var straight = Factory.CreateRibbonButton(); straight.Label = "直线"; sketch.Items.Add(straight);
            menu.Items.Add(sketch);

            var dashes = Factory.CreateRibbonMenu();
            dashes.Label = "虚线"; dashes.OfficeImageId = "LinePatternGallery"; dashes.ShowImage = true;
            dashes.KeyTip = "S";
            for (int i = 0; i < StrokeDashNames.Length; i++)
            {
                int dash = i;
                var item = AddZoomButton(dashes, StrokeDashNames[i], () => { if (box) zoomRibbonSettings.BoxLineDash = dash; else zoomRibbonSettings.LineDash = dash; });
                item.Image = GetStrokePreview("dash:" + dash.ToString(CultureInfo.InvariantCulture), strokeColor);
                item.ShowImage = true;
            }
            menu.Items.Add(dashes);

            var arrows = Factory.CreateRibbonMenu(); arrows.Label = "箭头"; arrows.Enabled = !box;
            arrows.Image = GetStrokePreview("arrow:2:end", strokeColor); arrows.ShowImage = true;
            foreach (bool begin in new[] { true, false })
            {
                var end = Factory.CreateRibbonMenu(); end.Label = begin ? "起点" : "终点";
                for (int i = 0; i < StrokeArrowNames.Length; i++)
                {
                    int value = i + 1; bool atBegin = begin;
                    var item = AddZoomButton(end, StrokeArrowNames[i], () =>
                    {
                        if (atBegin) zoomRibbonSettings.LineBeginArrow = value; else zoomRibbonSettings.LineEndArrow = value;
                    });
                    item.Image = GetStrokePreview("arrow:" + value.ToString(CultureInfo.InvariantCulture) + (begin ? ":begin" : ":end"), strokeColor);
                    item.ShowImage = true;
                }
                arrows.Items.Add(end);
            }
            menu.Items.Add(arrows);
        }

        private void AddStrokeGallery(RibbonMenu menu, string label, IEnumerable<Color> colors, bool box)
        {
            var heading = Factory.CreateRibbonSeparator(); heading.Title = label; menu.Items.Add(heading);
            var gallery = Factory.CreateRibbonGallery();
            gallery.Label = label; gallery.ColumnCount = 10; gallery.ShowItemLabel = false;
            gallery.ShowItemImage = true; gallery.ShowItemSelection = true; gallery.ItemImageSize = new Size(20, 20);
            foreach (Color color in colors)
            {
                int rgb = ColorTranslator.ToOle(color);
                if (!strokeSwatches.TryGetValue(rgb, out Bitmap swatch))
                {
                    swatch = new Bitmap(20, 20);
                    using (var graphics = Graphics.FromImage(swatch))
                    {
                        graphics.Clear(color);
                        graphics.DrawRectangle(Pens.Gray, 0, 0, 19, 19);
                    }
                    strokeSwatches.Add(rgb, swatch);
                }
                var item = Factory.CreateRibbonDropDownItem();
                item.Label = ColorTranslator.ToHtml(color); item.ScreenTip = item.Label; item.Image = swatch; item.Tag = rgb;
                gallery.Items.Add(item);
            }
            gallery.RowCount = Math.Max(1, (gallery.Items.Count + 9) / 10);
            gallery.Click += (s, e) =>
            {
                if (gallery.SelectedItem == null) return;
                SetZoomColor(box, (int)gallery.SelectedItem.Tag);
                RefreshZoomLabels(); ZoomSettings.Save(zoomRibbonSettings);
            };
            menu.Items.Add(gallery);
        }

        /// <summary>
        /// 绘制与原生“形状轮廓”菜单同风格的条目图标：无轮廓、取色器、粗细预览、
        /// 虚线预览与箭头预览。位图按颜色缓存，颜色切换后重新生成。
        /// </summary>
        private Bitmap GetStrokePreview(string key, int colorRgb, float weight = 1f)
        {
            string cacheKey = key + "@" + colorRgb.ToString("X6", CultureInfo.InvariantCulture);
            if (strokePreviews.TryGetValue(cacheKey, out Bitmap cached)) return cached;
            Color color = ColorTranslator.FromOle(colorRgb);
            var bitmap = new Bitmap(StrokePreviewSize, StrokePreviewSize);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                float left = 5f, right = StrokePreviewSize - 5f, middle = StrokePreviewSize / 2f;
                if (key.StartsWith("weight:", StringComparison.Ordinal))
                {
                    float thickness = Math.Min(10f, Math.Max(1f, weight * 1.6f));
                    using (var pen = new Pen(color, thickness))
                        graphics.DrawLine(pen, left, middle, right, middle);
                }
                else if (key.StartsWith("dash:", StringComparison.Ordinal))
                {
                    int dash = int.Parse(key.Substring(5), CultureInfo.InvariantCulture);
                    using (var pen = new Pen(color, 1.8f) { DashCap = DashCap.Flat })
                    {
                        ApplyPreviewDash(pen, dash);
                        graphics.DrawLine(pen, left, middle, right, middle);
                    }
                }
                else if (key.StartsWith("arrow:", StringComparison.Ordinal))
                {
                    string[] parts = key.Split(':');
                    int style = parts.Length > 1 && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0;
                    bool atBegin = parts.Length > 2 && parts[2] == "begin";
                    using (var pen = new Pen(color, 1.8f))
                    {
                        graphics.DrawLine(pen, left, middle, right, middle);
                        if (style > 0)
                            DrawPreviewArrowHead(graphics, color, new PointF(atBegin ? left : right, middle), atBegin, style);
                    }
                }
                else if (key == "no-outline")
                {
                    using (var pen = new Pen(Color.FromArgb(150, 150, 150), 1.6f))
                        graphics.DrawRectangle(pen, 6f, 6f, StrokePreviewSize - 12f, StrokePreviewSize - 12f);
                    using (var pen = new Pen(Color.FromArgb(214, 64, 64), 2.2f))
                        graphics.DrawLine(pen, 5f, StrokePreviewSize - 5f, StrokePreviewSize - 5f, 5f);
                }
                else if (key == "eyedropper")
                {
                    using (var pen = new Pen(Color.FromArgb(90, 90, 90), 2.2f))
                    {
                        graphics.DrawLine(pen, 7f, StrokePreviewSize - 7f, StrokePreviewSize - 11f, 11f);
                        using (var brush = new SolidBrush(color))
                            graphics.FillEllipse(brush, StrokePreviewSize - 12f, 5f, 7f, 7f);
                    }
                }
            }
            strokePreviews[cacheKey] = bitmap;
            return bitmap;
        }

        private static void ApplyPreviewDash(Pen pen, int dash)
        {
            switch (dash)
            {
                case 1: pen.DashPattern = new[] { 3f, 2f }; break;
                case 2: pen.DashPattern = new[] { 1f, 2f }; pen.DashCap = DashCap.Round; break;
                case 3: pen.DashPattern = new[] { 3f, 2f, 1f, 2f }; break;
                case 4: pen.DashPattern = new[] { 1f, 2f }; break;
                case 5: pen.DashPattern = new[] { 3f, 2f, 1f, 2f, 1f, 2f }; break;
                case 6: pen.DashPattern = new[] { 5f, 2f }; break;
                case 7: pen.DashPattern = new[] { 5f, 2f, 1f, 2f }; break;
                default: pen.DashStyle = DashStyle.Solid; break;
            }
        }

        private static void DrawPreviewArrowHead(Graphics graphics, Color color, PointF tip, bool atBegin, int style)
        {
            // dx 指向箭头尾部（沿线条向内），所有形状都画在图标范围内。
            float dx = atBegin ? 1f : -1f;
            using (var brush = new SolidBrush(color))
            using (var pen = new Pen(color, 1.6f))
            {
                switch (style)
                {
                    case 2: // 三角箭头
                        graphics.FillPolygon(brush, new[] { tip, new PointF(tip.X + dx * 8f, tip.Y - 5f), new PointF(tip.X + dx * 8f, tip.Y + 5f) });
                        break;
                    case 3: // 开放箭头
                        graphics.DrawLines(pen, new[] { new PointF(tip.X + dx * 8f, tip.Y - 5f), tip, new PointF(tip.X + dx * 8f, tip.Y + 5f) });
                        break;
                    case 4: // 燕尾箭头
                        graphics.FillPolygon(brush, new[] { tip, new PointF(tip.X + dx * 8f, tip.Y - 5f), new PointF(tip.X + dx * 4f, tip.Y), new PointF(tip.X + dx * 8f, tip.Y + 5f) });
                        break;
                    case 5: // 菱形
                        graphics.FillPolygon(brush, new[] { tip, new PointF(tip.X + dx * 4f, tip.Y - 4f), new PointF(tip.X + dx * 8f, tip.Y), new PointF(tip.X + dx * 4f, tip.Y + 4f) });
                        break;
                    case 6: // 椭圆
                        graphics.FillEllipse(brush, tip.X + dx * 5f - 4f, tip.Y - 4f, 8f, 8f);
                        break;
                }
            }
        }

        private Color[] GetStrokeThemeColors()
        {
            Color[] colors = { Color.White, Color.Black, Color.FromArgb(231,230,230), Color.FromArgb(68,84,106),
                Color.FromArgb(68,114,196), Color.FromArgb(237,125,49), Color.FromArgb(165,165,165),
                Color.FromArgb(255,192,0), Color.FromArgb(91,155,213), Color.FromArgb(112,173,71) };
            try
            {
                if (app == null || !PowerPointContext.TryGetActiveSlide(app, out var slide)) return colors;
                var scheme = slide.Design.SlideMaster.Theme.ThemeColorScheme;
                // Office scheme order: dark1/light1/dark2/light2/accent1..6.
                int[] indices = { 2, 1, 4, 3, 5, 6, 7, 8, 9, 10 };
                for (int i = 0; i < indices.Length; i++) colors[i] = ColorTranslator.FromOle(scheme.Colors((Microsoft.Office.Core.MsoThemeColorSchemeIndex)indices[i]).RGB);
            }
            catch (System.Runtime.InteropServices.COMException) { }
            return colors;
        }

        private static Color TintStrokeColor(Color color, float tint)
        {
            Func<byte, int> channel = value => (int)Math.Round(tint >= 0 ? value + (255 - value) * tint : value * (1 + tint));
            return Color.FromArgb(channel(color.R), channel(color.G), channel(color.B));
        }

        private List<int> ReadRecentStrokeColors()
        {
            var result = new List<int>();
            foreach (string text in (zoomRibbonSettings.RecentStrokeColors ?? "").Split(','))
                if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int color) && color >= 0 && color <= 0xFFFFFF && !result.Contains(color))
                { result.Add(color); if (result.Count == 10) break; }
            return result;
        }

        private void SetZoomColor(bool box, int color)
        {
            if (box) { zoomRibbonSettings.BoxColorRgb = color; zoomRibbonSettings.BoxLineVisible = true; }
            else { zoomRibbonSettings.LineColorRgb = color; zoomRibbonSettings.LineVisible = true; }
            var recent = ReadRecentStrokeColors(); recent.Remove(color); recent.Insert(0, color);
            zoomRibbonSettings.RecentStrokeColors = string.Join(",", recent.Take(10).Select(c => c.ToString(CultureInfo.InvariantCulture)));
        }

        private sealed class StrokeColorPicker : Form
        {
            private readonly Bitmap snapshot;
            public Color SelectedColor { get; private set; }

            public StrokeColorPicker()
            {
                Rectangle screen = SystemInformation.VirtualScreen;
                snapshot = new Bitmap(screen.Width, screen.Height);
                try
                {
                    using (var graphics = Graphics.FromImage(snapshot))
                        graphics.CopyFromScreen(screen.Location, Point.Empty, screen.Size);
                    AutoScaleMode = AutoScaleMode.None;
                    FormBorderStyle = FormBorderStyle.None;
                    StartPosition = FormStartPosition.Manual;
                    Bounds = screen;
                    ShowInTaskbar = false; TopMost = true; KeyPreview = true;
                    BackgroundImage = snapshot; BackgroundImageLayout = ImageLayout.Stretch;
                    Cursor = Cursors.Cross;
                    AccessibleName = "取色器：单击取色，Esc 或右键取消";
                }
                catch { snapshot.Dispose(); throw; }
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Right) { DialogResult = DialogResult.Cancel; Close(); return; }
                if (e.Button != MouseButtons.Left) return;
                int x = Math.Max(0, Math.Min(snapshot.Width - 1, e.X * snapshot.Width / Math.Max(1, ClientSize.Width)));
                int y = Math.Max(0, Math.Min(snapshot.Height - 1, e.Y * snapshot.Height / Math.Max(1, ClientSize.Height)));
                SelectedColor = snapshot.GetPixel(x, y);
                DialogResult = DialogResult.OK; Close();
            }

            protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
            {
                if (keyData == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); return true; }
                return base.ProcessCmdKey(ref msg, keyData);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing) { BackgroundImage = null; snapshot.Dispose(); }
                base.Dispose(disposing);
            }
        }
    }
}
