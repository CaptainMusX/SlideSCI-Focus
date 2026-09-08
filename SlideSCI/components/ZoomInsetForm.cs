using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>
    /// 「局部放大设置」v4。
    /// 完全重做交互与布局：分段选择 + 模式联动输入 + 实时示意图预览；
    /// 沿用 ScientificUiTheme 令牌与外部架构（多选区、等宽保比例、色块选择器、键盘焦点）。
    /// 构造时可用 ZoomSettings 初始化；确定后通过属性暴露结果。
    /// </summary>
    public sealed class ZoomInsetForm : Form
    {
        private const int ColorBlack = 0x000000;
        private const int ColorWhite = 0xFFFFFF;
        private const int ColorRed = 0x0000FF;
        private const int ColorBlue = 0xFF0000;
        private const int ColorGreen = 0x00A651;
        private const float PointsPerCm = 28.3464593f;

        private readonly float boxWidth;
        private readonly float boxHeight;
        private readonly float pictureWidth;
        private readonly float pictureHeight;

        private SegmentedControl segTarget;
        private SegmentedControl segLines;
        private Label lblActiveInput;
        private NumericUpDown numMagnification;
        private NumericUpDown numCustomWidth;
        private ComboBox cmbLineWeight;
        private ComboBox cmbBoxLineWeight;
        private ComboBox cmbLineColor;
        private ComboBox cmbBoxColor;
        private ComboBox cmbLineDash;
        private NumericUpDown numGap;
        private CheckBox chkGroup;
        private Label lblPreview;
        private PreviewDiagram diagram;
        private readonly ToolTip toolTip = new ToolTip();

        public ZoomTargetMode TargetMode { get; private set; }
        public float Magnification { get; private set; }
        public float CustomWidthCm { get; private set; }
        public float GapCm { get; private set; }
        public ZoomLineStyle LineStyle { get; private set; }
        public float LineWeight { get; private set; }
        public float BoxLineWeight { get; private set; }
        public int LineColorRgb { get; private set; }
        public int BoxColorRgb { get; private set; }
        public Office.MsoLineDashStyle LineDashStyle { get; private set; }
        public bool GroupEnabled { get; private set; }

        public ZoomInsetForm(float boxWidth, float boxHeight,
            float pictureWidth, float pictureHeight, ZoomSettings initial = null)
        {
            this.boxWidth = boxWidth;
            this.boxHeight = boxHeight;
            this.pictureWidth = pictureWidth;
            this.pictureHeight = pictureHeight;

            ZoomSettings s = (initial ?? ZoomSettings.CreateDefault()).NormalizedCopy();

            Text = "局部放大设置";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(520, 668);
            ScientificUiTheme.ConfigureDialog(this);

            BuildLayout(s);
            SyncInputsEnabled();
            UpdatePreview();
            ScientificUiTheme.CompleteCodeBuiltLayout(this);
        }

        private void BuildLayout(ZoomSettings s)
        {
            const int margin = 18;
            int contentWidth = ClientSize.Width - margin * 2; // 484
            int y = 16;

            // —— 头部 ——
            var title = new Label
            {
                Text = "生成局部放大图",
                Font = ScientificUiTheme.ChineseBodyFont(14f, FontStyle.Bold),
                ForeColor = ScientificUiTheme.Primary,
                AutoSize = true,
                Location = new Point(margin, y)
            };
            Controls.Add(title);

            var hint = new Label
            {
                Text = "Ctrl + 单击工具栏按钮可再次打开本设置",
                Font = ScientificUiTheme.BodyFont(8.5f),
                ForeColor = ScientificUiTheme.TextSecondary,
                AutoSize = true,
                Location = new Point(margin + contentWidth - 210, y + 6)
            };
            Controls.Add(hint);
            y += 24;

            var subtitle = new Label
            {
                Text = "精确裁剪选区，保持原始像素与宽高比；同页支持多个选区，各自独立更新。",
                Font = ScientificUiTheme.ChineseBodyFont(9f),
                ForeColor = ScientificUiTheme.TextSecondary,
                AutoSize = true,
                Location = new Point(margin, y)
            };
            Controls.Add(subtitle);
            y += 26;

            // —— 卡片一：放大尺寸 ——
            var cardTarget = new ScientificCard
            {
                Bounds = new Rectangle(margin, y, contentWidth, 148),
                Padding = new Padding(14, 10, 14, 10)
            };
            cardTarget.Controls.Add(SectionHeader("放大尺寸", 0, 0));

            segTarget = new SegmentedControl { Bounds = new Rectangle(14, 36, 456, 30) };
            segTarget.SetItems("与原图等宽", "指定放大倍数", "自定义宽度 (cm)");
            segTarget.SelectedIndex = (int)s.TargetMode;
            segTarget.SelectionChanged += (sender, args) =>
            {
                SyncInputsEnabled();
                UpdatePreview();
            };
            cardTarget.Controls.Add(segTarget);

            lblActiveInput = new Label
            {
                Font = ScientificUiTheme.ChineseBodyFont(9f),
                ForeColor = ScientificUiTheme.TextPrimary,
                AutoSize = true,
                Location = new Point(16, 104)
            };
            cardTarget.Controls.Add(lblActiveInput);

            numMagnification = CreateNumber(0.1m, 100m, (decimal)s.Magnification, 0.1m, 1, "放大倍数");
            numMagnification.Bounds = new Rectangle(132, 98, 110, 26);
            numMagnification.ValueChanged += (sender, args) => UpdatePreview();
            cardTarget.Controls.Add(numMagnification);

            numCustomWidth = CreateNumber(0.1m, 200m, (decimal)s.CustomWidthCm, 0.1m, 1, "放大图宽度（厘米）");
            numCustomWidth.Bounds = new Rectangle(132, 98, 110, 26);
            numCustomWidth.ValueChanged += (sender, args) => UpdatePreview();
            cardTarget.Controls.Add(numCustomWidth);

            var aspectNote = new Label
            {
                Text = "始终保持选区宽高比，图像不会被拉伸",
                Font = ScientificUiTheme.BodyFont(8.5f),
                ForeColor = ScientificUiTheme.TextSecondary,
                AutoSize = true,
                Location = new Point(256, 104)
            };
            cardTarget.Controls.Add(aspectNote);

            Controls.Add(cardTarget);
            y += 148 + 12;

            // —— 卡片二：连接方式 ——
            var cardLines = new ScientificCard
            {
                Bounds = new Rectangle(margin, y, contentWidth, 108),
                Padding = new Padding(14, 10, 14, 10)
            };
            cardLines.Controls.Add(SectionHeader("连接方式", 0, 0));

            segLines = new SegmentedControl { Bounds = new Rectangle(14, 36, 456, 30) };
            segLines.SetItems("期刊漏斗", "交叉引线");
            segLines.SelectedIndex = s.LineStyle == ZoomLineStyle.CrossedX ? 1 : 0;
            segLines.SelectionChanged += (sender, args) => UpdatePreview();
            cardLines.Controls.Add(segLines);

            var funDesc = new Label
            {
                Text = "期刊漏斗：选区下方两角 → 放大图上方两角，平行连接",
                Font = ScientificUiTheme.BodyFont(8.5f),
                ForeColor = ScientificUiTheme.TextSecondary,
                AutoSize = true,
                Location = new Point(16, 76)
            };
            cardLines.Controls.Add(funDesc);

            Controls.Add(cardLines);
            y += 108 + 12;

            // —— 卡片三：线条与布局 ——
            var cardAppearance = new ScientificCard
            {
                Bounds = new Rectangle(margin, y, contentWidth, 136),
                Padding = new Padding(14, 10, 14, 10)
            };
            cardAppearance.Controls.Add(SectionHeader("线条与布局", 0, 0));

            cmbLineWeight = CreateCombo(new object[] { 0.5f, 0.75f, 1f, 1.5f, 2f }, s.LineWeight, "连线粗细");
            cmbLineColor = CreateColorCombo(s.LineColorRgb, "连线颜色");
            cmbLineDash = CreateCombo(new object[] { "实线", "虚线" }, s.LineDash, "连线线型");
            cmbBoxLineWeight = CreateCombo(new object[] { 0.5f, 0.75f, 1f, 1.5f, 2f, 3f }, s.BoxLineWeight, "选区框粗细");
            cmbBoxColor = CreateColorCombo(s.BoxColorRgb, "选区框颜色");
            numGap = CreateNumber(0m, 50m, (decimal)s.GapCm, 0.1m, 1, "放大图与原图间距（厘米）");
            numGap.ValueChanged += (sender, args) => UpdatePreview();

            AddFieldRow(cardAppearance, 36, "连线粗细 (pt)", cmbLineWeight, "连线颜色", cmbLineColor);
            AddFieldRow(cardAppearance, 68, "选区框粗细 (pt)", cmbBoxLineWeight, "选区框颜色", cmbBoxColor);
            AddFieldRow(cardAppearance, 100, "连线线型", cmbLineDash, "图间距 (cm)", numGap);

            Controls.Add(cardAppearance);
            y += 136 + 12;

            // —— 预览卡（示意图 + 文本） ——
            var cardPreview = new ScientificCard
            {
                Bounds = new Rectangle(margin, y, contentWidth, 96),
                BackColor = ScientificUiTheme.SurfaceMuted,
                Padding = new Padding(12, 10, 14, 10)
            };
            diagram = new PreviewDiagram
            {
                Bounds = new Rectangle(10, 10, 118, 76),
                PicWidth = pictureWidth,
                PicHeight = Math.Max(1f, pictureHeight),
                BoxWidth = Math.Max(1f, boxWidth),
                BoxHeight = Math.Max(1f, boxHeight)
            };
            cardPreview.Controls.Add(diagram);

            lblPreview = new Label
            {
                Bounds = new Rectangle(136, 10, cardPreview.ClientSize.Width - 136 - 10, 76),
                ForeColor = ScientificUiTheme.TextSecondary,
                Font = ScientificUiTheme.ChineseBodyFont(9f),
                Text = ""
            };
            cardPreview.Controls.Add(lblPreview);

            Controls.Add(cardPreview);
            y += 96 + 14;

            // —— 编组开关 ——
            chkGroup = new CheckBox
            {
                Text = "生成后自动编组（选区框、放大图与连线）",
                AutoSize = true,
                Checked = s.Group,
                Location = new Point(margin, y),
                ForeColor = ScientificUiTheme.TextPrimary,
                AccessibleName = "生成后自动编组"
            };
            toolTip.SetToolTip(chkGroup, "编组后适合整体移动；重新生成时会保留当前选区框。");
            Controls.Add(chkGroup);
            y += 30;

            var note = new Label
            {
                Text = "提示：单击工具栏「放大」将直接沿用本次设置；同一页多个选区时，请先选中要更新的选区框或它所在的组。",
                Font = ScientificUiTheme.BodyFont(8.5f),
                ForeColor = ScientificUiTheme.TextSecondary,
                AutoSize = true,
                MaximumSize = new Size(contentWidth, 0),
                Location = new Point(margin, y)
            };
            Controls.Add(note);
            y += note.PreferredSize.Height + 10;

            // —— 底部按钮 ——
            var cancel = new ScientificButton
            {
                Text = "取消",
                Primary = false,
                Width = 96,
                Height = 34,
                DialogResult = DialogResult.Cancel,
                Location = new Point(margin + contentWidth - 96 - 8 - 120, y)
            };
            var ok = new ScientificButton
            {
                Text = "生成放大图",
                Primary = true,
                Width = 120,
                Height = 34,
                AccessibleName = "生成局部放大图",
                Location = new Point(margin + contentWidth - 120, y)
            };
            ok.Click += (sender, args) =>
            {
                ReadValues();
                DialogResult = DialogResult.OK;
            };
            Controls.Add(cancel);
            Controls.Add(ok);
            AcceptButton = ok;
            CancelButton = cancel;
            y += 44;

            // 实际高度以内容为准
            ClientSize = new Size(ClientSize.Width, Math.Max(ClientSize.Height, y));
        }

        private static Label SectionHeader(string text, int x, int y)
        {
            return new Label
            {
                Text = text,
                Font = ScientificUiTheme.ChineseBodyFont(10f, FontStyle.Bold),
                ForeColor = ScientificUiTheme.Primary,
                AutoSize = true,
                Location = new Point(x, y)
            };
        }

        private static void AddFieldRow(Control parent, int y, string labelText, Control control,
            string labelText2, Control control2)
        {
            parent.Controls.Add(new Label
            {
                Text = labelText,
                Font = ScientificUiTheme.ChineseBodyFont(9f),
                ForeColor = ScientificUiTheme.TextSecondary,
                AutoSize = true,
                Location = new Point(16, y + 4)
            });
            control.Bounds = new Rectangle(120, y, 96, 26);
            parent.Controls.Add(control);

            parent.Controls.Add(new Label
            {
                Text = labelText2,
                Font = ScientificUiTheme.ChineseBodyFont(9f),
                ForeColor = ScientificUiTheme.TextSecondary,
                AutoSize = true,
                Location = new Point(250, y + 4)
            });
            control2.Bounds = new Rectangle(348, y, 110, 26);
            parent.Controls.Add(control2);
        }

        private NumericUpDown CreateNumber(decimal min, decimal max, decimal value,
            decimal increment, int decimals, string accessibleName)
        {
            var number = new NumericUpDown
            {
                Minimum = min,
                Maximum = max,
                Value = Math.Max(min, Math.Min(max, value)),
                Increment = increment,
                DecimalPlaces = decimals,
                ThousandsSeparator = false,
                TextAlign = HorizontalAlignment.Right,
                BorderStyle = BorderStyle.FixedSingle,
                Font = ScientificUiTheme.ChineseBodyFont(9f),
                AccessibleName = accessibleName
            };
            return number;
        }

        private ComboBox CreateCombo(object[] values, object selected, string accessibleName)
        {
            var combo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = ScientificUiTheme.ChineseBodyFont(9f),
                AccessibleName = accessibleName
            };
            combo.Items.AddRange(values);
            int index = Array.IndexOf(values, selected);
            combo.SelectedIndex = index >= 0 ? index : 0;
            return combo;
        }

        private ComboBox CreateColorCombo(int initialRgb, string accessibleName)
        {
            var combo = CreateCombo(new object[]
            {
                new ColorChoice("黑色", ColorBlack),
                new ColorChoice("白色", ColorWhite),
                new ColorChoice("红色", ColorRed),
                new ColorChoice("蓝色", ColorBlue),
                new ColorChoice("绿色", ColorGreen),
                new ColorChoice("自定义…", initialRgb, true)
            }, initialRgb == ColorRed ? 2 : 0, accessibleName);
            combo.DrawMode = DrawMode.OwnerDrawFixed;
            combo.DrawItem += DrawColorItem;
            combo.SelectedIndexChanged += ColorSelectionChanged;
            return combo;
        }

        private void SyncInputsEnabled()
        {
            bool multiple = segTarget.SelectedIndex == (int)ZoomTargetMode.Multiple;
            bool custom = segTarget.SelectedIndex == (int)ZoomTargetMode.CustomWidthCm;

            lblActiveInput.Text = multiple ? "放大倍数" : custom ? "放大图宽度 (cm)" : "";
            numMagnification.Visible = multiple;
            numCustomWidth.Visible = custom;
        }

        private void DrawColorItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (!(sender is ComboBox combo) || e.Index < 0 || e.Index >= combo.Items.Count) return;

            var choice = combo.Items[e.Index] as ColorChoice;
            Rectangle swatch = new Rectangle(e.Bounds.Left + 4, e.Bounds.Top + 4, 15,
                Math.Max(8, e.Bounds.Height - 8));
            using (var brush = new SolidBrush(RgbToColor(choice?.Rgb ?? ColorBlack)))
            using (var border = new Pen(ScientificUiTheme.Border))
            {
                e.Graphics.FillRectangle(brush, swatch);
                e.Graphics.DrawRectangle(border, swatch);
            }
            TextRenderer.DrawText(e.Graphics, choice?.Label ?? string.Empty, combo.Font,
                new Rectangle(swatch.Right + 6, e.Bounds.Top, e.Bounds.Width - swatch.Width - 12, e.Bounds.Height),
                e.ForeColor, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            e.DrawFocusRectangle();
        }

        private void ColorSelectionChanged(object sender, EventArgs e)
        {
            if (!(sender is ComboBox combo) || !(combo.SelectedItem is ColorChoice choice) || !choice.Custom) return;

            using (var dialog = new ColorDialog
            {
                FullOpen = true,
                Color = RgbToColor(choice.Rgb)
            })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    choice.Rgb = ColorToRgb(dialog.Color);
                    combo.Invalidate();
                    UpdatePreview();
                }
            }
        }

        private void UpdatePreview()
        {
            if (lblPreview == null || boxWidth <= 0 || boxHeight <= 0) return;

            float targetWidth;
            float targetHeight;
            if (segTarget.SelectedIndex == (int)ZoomTargetMode.Multiple)
            {
                float factor = (float)numMagnification.Value;
                targetWidth = boxWidth * factor;
                targetHeight = boxHeight * factor;
            }
            else if (segTarget.SelectedIndex == (int)ZoomTargetMode.CustomWidthCm)
            {
                targetWidth = CmToPoints((float)numCustomWidth.Value);
                targetHeight = boxHeight * (targetWidth / boxWidth);
            }
            else
            {
                targetWidth = pictureWidth;
                targetHeight = boxHeight * (targetWidth / boxWidth);
            }

            float factorX = targetWidth / boxWidth;
            float widthCm = PointsToCm(targetWidth);
            float heightCm = PointsToCm(targetHeight);

            lblPreview.Text = string.Format(CultureInfo.CurrentCulture,
                "目标尺寸约 {0:0.0} × {1:0.0} cm\n实际放大 {2:0.0}×\n保持选区宽高比；下方空间不足时自动放到原图右侧。",
                widthCm, heightCm, factorX);
            lblPreview.ForeColor = widthCm > 40f || heightCm > 40f
                ? ScientificUiTheme.Warning
                : ScientificUiTheme.TextSecondary;

            if (diagram != null)
            {
                diagram.TargetWidth = targetWidth;
                diagram.TargetHeight = targetHeight;
                diagram.LineStyle = segLines.SelectedIndex == 1 ? ZoomLineStyle.CrossedX : ZoomLineStyle.JournalFunnel;
                var lineChoice = cmbLineColor.SelectedItem as ColorChoice;
                var boxChoice = cmbBoxColor.SelectedItem as ColorChoice;
                diagram.LineColorRgb = lineChoice?.Rgb ?? ColorBlack;
                diagram.BoxColorRgb = boxChoice?.Rgb ?? ColorBlack;
                diagram.Invalidate();
            }
        }

        private void ReadValues()
        {
            TargetMode = (ZoomTargetMode)segTarget.SelectedIndex;
            Magnification = (float)numMagnification.Value;
            CustomWidthCm = (float)numCustomWidth.Value;
            GapCm = (float)numGap.Value;
            LineStyle = segLines.SelectedIndex == 1 ? ZoomLineStyle.CrossedX : ZoomLineStyle.JournalFunnel;
            LineWeight = Convert.ToSingle(cmbLineWeight.SelectedItem, CultureInfo.InvariantCulture);
            BoxLineWeight = Convert.ToSingle(cmbBoxLineWeight.SelectedItem, CultureInfo.InvariantCulture);
            LineColorRgb = ((ColorChoice)cmbLineColor.SelectedItem).Rgb;
            BoxColorRgb = ((ColorChoice)cmbBoxColor.SelectedItem).Rgb;
            LineDashStyle = cmbLineDash.SelectedIndex == 1
                ? Office.MsoLineDashStyle.msoLineDash
                : Office.MsoLineDashStyle.msoLineSolid;
            GroupEnabled = chkGroup.Checked;
        }

        private static float CmToPoints(float cm) => cm * PointsPerCm;
        private static float PointsToCm(float points) => points / PointsPerCm;

        private static int ColorToRgb(Color color)
        {
            return (color.B << 16) | (color.G << 8) | color.R;
        }

        private static Color RgbToColor(int rgb)
        {
            return Color.FromArgb(rgb & 0xFF, (rgb >> 8) & 0xFF, (rgb >> 16) & 0xFF);
        }

        /// <summary>分段选择控件：药丸轨道 + 白色选中段，支持键盘左右键与焦点框。</summary>
        private sealed class SegmentedControl : Control
        {
            private string[] items = new string[0];
            private int selectedIndex;
            private int hoverIndex = -1;

            public event EventHandler SelectionChanged;

            public int SelectedIndex
            {
                get => selectedIndex;
                set
                {
                    if (value < 0 || value >= items.Length || value == selectedIndex) return;
                    selectedIndex = value;
                    SelectionChanged?.Invoke(this, EventArgs.Empty);
                    Invalidate();
                }
            }

            public SegmentedControl()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.SupportsTransparentBackColor, true);
                BackColor = Color.Transparent;
                Font = ScientificUiTheme.ChineseBodyFont(9f);
                TabStop = true;
                Height = 30;
            }

            public void SetItems(params string[] values)
            {
                Array.Resize(ref items, values.Length);
                Array.Copy(values, items, values.Length);
                selectedIndex = 0;
                Invalidate();
            }

            protected override void OnMouseLeave(EventArgs e) { hoverIndex = -1; Invalidate(); base.OnMouseLeave(e); }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                int hit = HitTest(e.X);
                if (hit != hoverIndex) { hoverIndex = hit; Invalidate(); }
                base.OnMouseMove(e);
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                if (e.Button == MouseButtons.Left)
                {
                    int hit = HitTest(e.X);
                    if (hit >= 0) { SelectedIndex = hit; Focus(); }
                }
                base.OnMouseDown(e);
            }

            protected override void OnKeyDown(KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Left && selectedIndex > 0) { SelectedIndex--; e.Handled = true; }
                else if (e.KeyCode == Keys.Right && selectedIndex < items.Length - 1) { SelectedIndex++; e.Handled = true; }
                base.OnKeyDown(e);
            }

            private int HitTest(int x)
            {
                if (items.Length == 0) return -1;
                int w = Width / items.Length;
                int index = x / w;
                return Math.Min(items.Length - 1, Math.Max(0, index));
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                if (items.Length == 0) return;
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                using (GraphicsPath track = ScientificUiTheme.RoundedPath(
                    new Rectangle(0, 0, Width - 1, Height - 1), Height / 2))
                using (var trackBrush = new SolidBrush(Color.FromArgb(226, 232, 240)))
                {
                    g.FillPath(trackBrush, track);
                }

                int w = Width / items.Length;
                for (int i = 0; i < items.Length; i++)
                {
                    Rectangle segment = new Rectangle(i * w, 0, w, Height);
                    if (i == selectedIndex)
                    {
                        using (GraphicsPath selected = ScientificUiTheme.RoundedPath(
                            new Rectangle(segment.Left + 2, 2, w - 4, Height - 5), Height / 2 - 2))
                        using (var fill = new SolidBrush(ScientificUiTheme.Surface))
                        using (var border = new Pen(ScientificUiTheme.Border))
                        {
                            g.FillPath(fill, selected);
                            g.DrawPath(border, selected);
                        }
                    }

                    Color text = i == selectedIndex
                        ? ScientificUiTheme.TextPrimary
                        : i == hoverIndex
                            ? ScientificUiTheme.Action
                            : ScientificUiTheme.TextSecondary;
                    TextRenderer.DrawText(g, items[i], Font, segment, text,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter |
                        TextFormatFlags.EndEllipsis);
                }

                if (Focused)
                {
                    ControlPaint.DrawFocusRectangle(g, new Rectangle(2, 2, Width - 5, Height - 5),
                        ScientificUiTheme.Focus, Color.Transparent);
                }
            }
        }

        /// <summary>实时示意图：原图 + 选区框 + 放大图 + 引线，按当前参数绘制。</summary>
        private sealed class PreviewDiagram : Control
        {
            public float PicWidth { get; set; }
            public float PicHeight { get; set; }
            public float BoxWidth { get; set; }
            public float BoxHeight { get; set; }
            public float TargetWidth { get; set; }
            public float TargetHeight { get; set; }
            public ZoomLineStyle LineStyle { get; set; } = ZoomLineStyle.JournalFunnel;
            public int LineColorRgb { get; set; } = ColorBlack;
            public int BoxColorRgb { get; set; } = ColorBlack;

            public PreviewDiagram()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer |
                         ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.SupportsTransparentBackColor, true);
                BackColor = Color.Transparent;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;

                const float layoutH = 68f;
                const float originalWidth = 58f;

                float picAspect = PicHeight / Math.Max(1f, PicWidth);
                float targetAspect = TargetHeight / Math.Max(1f, TargetWidth);
                float originalHeight = Math.Max(10f, originalWidth * picAspect);
                float insetHeight = Math.Max(10f, originalWidth * targetAspect);
                float total = originalHeight + 8f + insetHeight;
                float scale = total > layoutH ? layoutH / total : 1f;

                float ow = originalWidth * scale;
                float oh = originalHeight * scale;
                float iw = originalWidth * scale;
                float ih = insetHeight * scale;
                float x = (Width - ow) / 2f;
                float originalTop = (Height - total * scale) / 2f;
                float insetTop = originalTop + oh + 8f * scale;

                using (var fill = new SolidBrush(ScientificUiTheme.SurfaceMuted))
                using (var border = new Pen(ScientificUiTheme.Border))
                {
                    // 原图
                    g.FillRectangle(fill, x, originalTop, ow, oh);
                    g.DrawRectangle(border, x, originalTop, ow, oh);

                    // 放大图
                    g.FillRectangle(fill, x, insetTop, iw, ih);
                    g.DrawRectangle(border, x, insetTop, iw, ih);
                }

                // 选区框
                float bw = ow * (BoxWidth / Math.Max(1f, PicWidth));
                float bh = bw * (BoxHeight / Math.Max(1f, BoxWidth));
                float bx = x + (ow - bw) / 2f;
                float by = originalTop + (oh - bh) / 2f;
                using (var boxPen = new Pen(RgbToColor(BoxColorRgb), 1.4f))
                {
                    g.DrawRectangle(boxPen, bx, by, bw, bh);
                }

                // 引线
                using (var linePen = new Pen(RgbToColor(LineColorRgb), 1.2f))
                {
                    float boxBLx = bx, boxBLy = by + bh;
                    float boxBRx = bx + bw, boxBRy = by + bh;
                    float boxTLx = bx, boxTLy = by;
                    float boxTRx = bx + bw, boxTRy = by;
                    float insTLx = x, insTLy = insetTop;
                    float insTRx = x + iw, insTRy = insetTop;
                    float insBRx = x + iw;
                    float insBRy = insetTop + ih;

                    if (LineStyle == ZoomLineStyle.JournalFunnel)
                    {
                        g.DrawLine(linePen, boxBLx, boxBLy, insTLx, insTLy);
                        g.DrawLine(linePen, boxBRx, boxBRy, insTRx, insTRy);
                    }
                    else
                    {
                        g.DrawLine(linePen, boxTLx, boxTLy, insTRx, insTRy);
                        g.DrawLine(linePen, boxTRx, boxTRy, insTLx, insTLy);
                        g.DrawLine(linePen, boxBLx, boxBLy, insBRx, insBRy);
                        g.DrawLine(linePen, boxBRx, boxBRy, insTLx, insTLy);
                    }
                }
            }
        }

        private sealed class ColorChoice
        {
            internal string Label { get; }
            internal int Rgb { get; set; }
            internal bool Custom { get; }

            internal ColorChoice(string label, int rgb, bool custom = false)
            {
                Label = label;
                Rgb = rgb;
                Custom = custom;
            }

            public override string ToString() => Label;
        }
    }
}
