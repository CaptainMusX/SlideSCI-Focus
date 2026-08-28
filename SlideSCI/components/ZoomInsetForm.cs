using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>
    /// 「生成放大图」参数对话框（Apple 风格）。
    /// 选项通过属性暴露给调用方；尺寸计算由 ZoomInsetHelper.ComputeTargetSize 完成。
    /// 构造时传入当前选区框与原图尺寸（磅），用于实时预览放大图尺寸/放大倍数。
    /// </summary>
    public class ZoomInsetForm : Form
    {
        // —— Apple 风格调色板 ——
        private static readonly Color Accent = Color.FromArgb(0, 122, 255);        // iOS 蓝
        private static readonly Color AccentHover = Color.FromArgb(0, 113, 227);
        private static readonly Color TextPrimary = Color.FromArgb(29, 29, 31);    // #1D1D1F
        private static readonly Color TextSecondary = Color.FromArgb(110, 110, 115); // #6E6E73
        private static readonly Color WindowBackground = Color.FromArgb(245, 245, 247); // #F5F5F7
        private static readonly Color CardBackground = Color.White;
        private static readonly Color CardBorder = Color.FromArgb(229, 229, 234);  // #E5E5EA
        private static readonly Color PreviewBackground = Color.FromArgb(242, 242, 247); // #F2F2F7
        private static readonly Color ButtonBorder = Color.FromArgb(209, 209, 214); // #D1D1D6
        private static readonly Color WarningColor = Color.FromArgb(255, 149, 0);  // iOS 橙

        private const int ColorBlack = 0x000000;
        private const int ColorWhite = 0xFFFFFF;
        private const int ColorRed = 0x0000FF;   // PPT RGB 实为 BGR
        private const int ColorBlue = 0xFF0000;
        private const int ColorGreen = 0x00FF00;
        private const string CustomItem = "自定义…";
        private const float PointsPerCm = 28.3464593f;

        private readonly float boxWidth;
        private readonly float boxHeight;
        private readonly float pictureWidth;
        private readonly float pictureHeight;

        private RadioButton rbSameAsOriginal;
        private RadioButton rbMultiple;
        private TextBox txtMagnification;
        private RadioButton rbCustomWidth;
        private TextBox txtCustomWidth;
        private RadioButton rbJournalFunnel;
        private RadioButton rbCrossedX;
        private ComboBox cmbLineWeight;
        private ComboBox cmbBoxLineWeight;
        private ComboBox cmbLineColor;
        private ComboBox cmbBoxColor;
        private ComboBox cmbLineDash;
        private TextBox txtGap;
        private CheckBox chkGroup;
        private Label lblPreview;
        private int customLineColorRgb = ColorRed;
        private int customBoxColorRgb = ColorRed;

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

        public ZoomInsetForm(float boxWidth, float boxHeight, float pictureWidth, float pictureHeight)
        {
            this.boxWidth = boxWidth;
            this.boxHeight = boxHeight;
            this.pictureWidth = pictureWidth;
            this.pictureHeight = pictureHeight;

            this.Text = "生成局部放大图";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(420, 584);
            this.BackColor = WindowBackground;
            this.Font = new Font("Segoe UI", 9f);

            BuildLayout();
            UpdatePreview();
        }

        private void BuildLayout()
        {
            const int margin = 24;
            int contentWidth = ClientSize.Width - margin * 2; // 372
            int y = 18;

            // —— 标题区 ——
            var title = new Label
            {
                Text = "生成局部放大图",
                Font = new Font("Segoe UI Semibold", 14f, FontStyle.Bold),
                ForeColor = TextPrimary,
                Location = new Point(margin, y),
                AutoSize = true
            };
            Controls.Add(title);
            y += 28;

            var subtitle = new Label
            {
                Text = "把选区框内的画面放大，并放置在原图下方",
                Font = new Font("Segoe UI", 9f),
                ForeColor = TextSecondary,
                Location = new Point(margin, y),
                AutoSize = true
            };
            Controls.Add(subtitle);
            y += 28;

            // —— 卡片一：放大目标 ——
            var cardTarget = new RoundedPanel
            {
                Bounds = new Rectangle(margin, y, contentWidth, 136),
                Padding = new Padding(16)
            };
            var lblTargetHeader = SectionHeader("放大目标");
            lblTargetHeader.Location = new Point(4, 4);
            cardTarget.Controls.Add(lblTargetHeader);

            rbSameAsOriginal = MakeRadio("与原图相同尺寸", 4, 32, cardTarget, true);
            rbMultiple = MakeRadio("指定放大倍数", 4, 62, cardTarget, false);
            txtMagnification = MakeSmallInput("2", contentWidth - 16 - 64, 62, 64, cardTarget);
            rbCustomWidth = MakeRadio("自定义宽度 (cm)", 4, 92, cardTarget, false);
            txtCustomWidth = MakeSmallInput("10", contentWidth - 16 - 64, 92, 64, cardTarget);

            rbSameAsOriginal.CheckedChanged += (s, e) => { SyncInputsEnabled(); UpdatePreview(); };
            rbMultiple.CheckedChanged += (s, e) => { SyncInputsEnabled(); UpdatePreview(); };
            rbCustomWidth.CheckedChanged += (s, e) => { SyncInputsEnabled(); UpdatePreview(); };
            txtMagnification.TextChanged += (s, e) => UpdatePreview();
            txtCustomWidth.TextChanged += (s, e) => UpdatePreview();

            Controls.Add(cardTarget);
            y += 136 + 12;

            // —— 卡片二：框角连线样式 ——
            var cardLines = new RoundedPanel
            {
                Bounds = new Rectangle(margin, y, contentWidth, 100),
                Padding = new Padding(16)
            };
            var lblLinesHeader = SectionHeader("框角连线样式");
            lblLinesHeader.Location = new Point(4, 4);
            cardLines.Controls.Add(lblLinesHeader);
            rbJournalFunnel = MakeRadio("期刊漏斗：框下两角 → 放大图上两角", 4, 34, cardLines, true);
            rbCrossedX = MakeRadio("交叉 X 形：四角交叉连线", 4, 64, cardLines, false);
            Controls.Add(cardLines);
            y += 100 + 12;

            // —— 两列栅格行 ——
            // 行1：连线粗细 | 框线粗细
            cmbLineWeight = AddComboRow(y, "连线粗细 (pt)", new object[] { 0.5f, 1f, 1.5f, 2f }, 1);
            cmbBoxLineWeight = AddComboRow2(y, "框线粗细 (pt)", new object[] { 1f, 1.5f, 2f, 3f }, 1);
            y += 30;

            // 行2：连线颜色 | 框线颜色
            cmbLineColor = AddComboRow(y, "连线颜色", new object[] { "黑色", "白色", "红色", "蓝色", "绿色", CustomItem }, 0);
            cmbLineColor.SelectedIndexChanged += (s, e) => HandleColorSelection(cmbLineColor, ref customLineColorRgb);
            cmbBoxColor = AddComboRow2(y, "框线颜色", new object[] { "黑色", "红色", "蓝色", "绿色" }, 0);
            y += 30;

            // 行3：连线线型 | 下边距
            cmbLineDash = AddComboRow(y, "连线线型", new object[] { "实线", "虚线" }, 0);
            var lblGap = new Label
            {
                Text = "下边距 (cm)",
                ForeColor = TextPrimary,
                Location = new Point(200, y + 4),
                AutoSize = true
            };
            Controls.Add(lblGap);
            txtGap = new TextBox
            {
                Text = "0.5",
                Bounds = new Rectangle(298, y, 64, 24),
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f)
            };
            Controls.Add(txtGap);
            y += 34;

            // —— 预览卡 ——
            var previewCard = new RoundedPanel
            {
                Bounds = new Rectangle(margin, y, contentWidth, 66),
                BackColor = PreviewBackground,
                Padding = new Padding(14, 10, 14, 10)
            };
            lblPreview = new Label
            {
                ForeColor = TextSecondary,
                Font = new Font("Segoe UI", 9f),
                AutoSize = false,
                Dock = DockStyle.Fill,
                Text = ""
            };
            previewCard.Controls.Add(lblPreview);
            Controls.Add(previewCard);
            y += 66 + 14;

            // —— 编组开关 ——
            chkGroup = new CheckBox
            {
                Text = "生成后自动编组：框 + 放大图 + 连线组成一组，方便整体移动",
                ForeColor = TextPrimary,
                Location = new Point(margin, y),
                AutoSize = true,
                MaximumSize = new Size(contentWidth, 0),
                Checked = true
            };
            Controls.Add(chkGroup);
            y += 30;

            // —— 底部按钮 ——
            var btnCancel = new RoundedButton
            {
                Text = "取消",
                Primary = false,
                Font = new Font("Segoe UI", 9f),
                Bounds = new Rectangle(420 - margin - 100 - 10, y, 100, 32)
            };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            var btnOk = new RoundedButton
            {
                Text = "生成",
                Primary = true,
                Font = new Font("Segoe UI Semibold", 9.5f),
                Bounds = new Rectangle(420 - margin - 100, y, 100, 32)
            };
            btnOk.Click += (s, e) => { if (ValidateAndRead()) this.DialogResult = DialogResult.OK; };

            Controls.Add(btnCancel);
            Controls.Add(btnOk);
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
        }

        private static Label SectionHeader(string text)
        {
            return new Label
            {
                Text = text,
                Font = new Font("Segoe UI Semibold", 10f, FontStyle.Bold),
                ForeColor = TextPrimary,
                AutoSize = true
            };
        }

        private static RadioButton MakeRadio(string text, int x, int y, Control parent, bool checkedState)
        {
            var radio = new RadioButton
            {
                Text = text,
                ForeColor = TextPrimary,
                Font = new Font("Segoe UI", 9f),
                Location = new Point(x, y),
                AutoSize = true,
                MaximumSize = new Size(parent.ClientSize.Width - 90, 0),
                Checked = checkedState
            };
            parent.Controls.Add(radio);
            return radio;
        }

        private static TextBox MakeSmallInput(string text, int x, int y, int width, Control parent)
        {
            var input = new TextBox
            {
                Text = text,
                Bounds = new Rectangle(x, y, width, 24),
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Segoe UI", 9f)
            };
            parent.Controls.Add(input);
            return input;
        }

        private ComboBox AddComboRow(int y, string labelText, object[] items, int selectedIndex)
        {
            CurrentRowY = y;
            AddGridLabel(labelText, 0);
            return AddGridCombo(items, selectedIndex, 0, y);
        }

        private ComboBox AddComboRow2(int y, string labelText, object[] items, int selectedIndex)
        {
            CurrentRowY = y;
            AddGridLabel(labelText, 1);
            return AddGridCombo(items, selectedIndex, 1, y);
        }

        private void AddGridLabel(string text, int column)
        {
            Controls.Add(new Label
            {
                Text = text,
                ForeColor = TextPrimary,
                Location = new Point(column == 0 ? 24 : 200, CurrentRowY + 4),
                AutoSize = true
            });
        }

        private int CurrentRowY;

        private ComboBox AddGridCombo(object[] items, int selectedIndex, int column, int y)
        {
            var combo = new ComboBox
            {
                Bounds = new Rectangle(column == 0 ? 118 : 298, y, 66, 24),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9f),
                FlatStyle = FlatStyle.Flat
            };
            combo.Items.AddRange(items);
            combo.SelectedIndex = selectedIndex;
            Controls.Add(combo);
            return combo;
        }

        private void SyncInputsEnabled()
        {
            txtMagnification.Enabled = rbMultiple.Checked;
            txtCustomWidth.Enabled = rbCustomWidth.Checked;
        }

        private void HandleColorSelection(ComboBox combo, ref int customRgb)
        {
            if (!combo.Text.Equals(CustomItem, StringComparison.Ordinal)) return;

            using (var dialog = new ColorDialog { FullOpen = true })
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    customRgb = ColorToRgb(dialog.Color);
                }
                else
                {
                    combo.SelectedIndex = 0; // 取消则回到黑色
                }
            }
        }

        private static int ColorToRgb(Color color)
        {
            return (color.B << 16) | (color.G << 8) | color.R; // PPT RGB 实为 BGR
        }

        private static int GetColorRgb(ComboBox combo, int customRgb)
        {
            switch (combo.Text)
            {
                case "白色": return ColorWhite;
                case "红色": return ColorRed;
                case "蓝色": return ColorBlue;
                case "绿色": return ColorGreen;
                case CustomItem: return customRgb;
                default: return ColorBlack;
            }
        }

        private void UpdatePreview()
        {
            if (lblPreview == null) return;

            string text;
            bool stretched = false;

            if (rbMultiple.Checked && float.TryParse(txtMagnification.Text.Trim(), out float mag) && mag > 0)
            {
                float w = PointsToCm(boxWidth * mag);
                float h = PointsToCm(boxHeight * mag);
                text = $"放大图尺寸 ≈ {w:0.0} × {h:0.0} cm（{mag:0.#}×）";
            }
            else if (rbCustomWidth.Checked && float.TryParse(txtCustomWidth.Text.Trim(), out float wCm) && wCm > 0)
            {
                float h = boxWidth > 0 ? PointsToCm(boxHeight) * (wCm / PointsToCm(boxWidth)) : 0;
                text = $"放大图尺寸 ≈ {wCm:0.0} × {h:0.0} cm（按选区比例）";
            }
            else
            {
                float w = PointsToCm(pictureWidth);
                float h = PointsToCm(pictureHeight);
                text = $"放大图尺寸 ≈ {w:0.0} × {h:0.0} cm（原图大小）";
                if (boxWidth > 0) text += $"\n实际放大倍数 ≈ {pictureWidth / boxWidth:F1}×";
                if (boxWidth > 0 && boxHeight > 0 && pictureWidth > 0 && pictureHeight > 0)
                {
                    float boxAspect = boxWidth / boxHeight;
                    float picAspect = pictureWidth / pictureHeight;
                    if (Math.Abs(boxAspect - picAspect) / Math.Max(boxAspect, picAspect) > 0.01f)
                    {
                        text += "\n⚠ 选区与目标画面比例不一致，将按原图尺寸拉伸";
                        stretched = true;
                    }
                }
            }

            lblPreview.Text = text;
            lblPreview.ForeColor = stretched ? WarningColor : TextSecondary;
        }

        private static float PointsToCm(float points) => points / PointsPerCm;

        private bool ValidateAndRead()
        {
            if (rbSameAsOriginal.Checked)
            {
                TargetMode = ZoomTargetMode.SameAsOriginal;
            }
            else if (rbMultiple.Checked)
            {
                if (!float.TryParse(txtMagnification.Text.Trim(), out float mag) || mag <= 0 || mag > 100)
                {
                    MessageBox.Show("请输入有效的放大倍数（大于 0）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                TargetMode = ZoomTargetMode.Multiple;
                Magnification = mag;
            }
            else
            {
                if (!float.TryParse(txtCustomWidth.Text.Trim(), out float w) || w <= 0 || w > 200)
                {
                    MessageBox.Show("请输入有效的放大图宽度（cm，0–200）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }
                TargetMode = ZoomTargetMode.CustomWidthCm;
                CustomWidthCm = w;
            }

            if (!float.TryParse(txtGap.Text.Trim(), out float gap) || gap < 0 || gap > 50)
            {
                MessageBox.Show("请输入有效的下边距（cm，0–50）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            GapCm = gap;
            LineStyle = rbCrossedX.Checked ? ZoomLineStyle.CrossedX : ZoomLineStyle.JournalFunnel;
            LineWeight = (float)cmbLineWeight.SelectedItem;
            BoxLineWeight = (float)cmbBoxLineWeight.SelectedItem;
            LineColorRgb = GetColorRgb(cmbLineColor, customLineColorRgb);
            BoxColorRgb = GetColorRgb(cmbBoxColor, customBoxColorRgb);
            LineDashStyle = cmbLineDash.SelectedIndex == 1
                ? Office.MsoLineDashStyle.msoLineDash
                : Office.MsoLineDashStyle.msoLineSolid;
            GroupEnabled = chkGroup.Checked;
            return true;
        }

        /// <summary>白色圆角卡片面板。</summary>
        private sealed class RoundedPanel : Panel
        {
            public RoundedPanel()
            {
                SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                BackColor = CardBackground;
            }

            protected override void OnPaintBackground(PaintEventArgs e) { }

            protected override void OnPaint(PaintEventArgs e)
            {
                Graphics g = e.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                using (GraphicsPath path = GetRoundedPath(new Rectangle(0, 0, Width - 1, Height - 1), 12))
                using (var fill = new SolidBrush(BackColor))
                using (var pen = new Pen(CardBorder))
                {
                    g.FillPath(fill, path);
                    g.DrawPath(pen, path);
                }
            }

            private static GraphicsPath GetRoundedPath(Rectangle bounds, int radius)
            {
                var path = new GraphicsPath();
                int d = radius * 2;
                path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
                path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
                path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
                path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }
        }

        /// <summary>圆角按钮：主按钮为 iOS 蓝实心，次要按钮为白色描边。</summary>
        private sealed class RoundedButton : Button
        {
            private bool hovered;

            public bool Primary { get; set; } = true;

            public RoundedButton()
            {
                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                Cursor = Cursors.Hand;
            }

            protected override void OnMouseEnter(EventArgs e) { hovered = true; Invalidate(); base.OnMouseEnter(e); }
            protected override void OnMouseLeave(EventArgs e) { hovered = false; Invalidate(); base.OnMouseLeave(e); }

            protected override void OnPaint(PaintEventArgs pevent)
            {
                Graphics g = pevent.Graphics;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                var rect = new Rectangle(0, 0, Width - 1, Height - 1);
                using (GraphicsPath path = GetRoundedPath(rect, 8))
                {
                    Color fill = Primary
                        ? (hovered ? AccentHover : Accent)
                        : (hovered ? Color.FromArgb(242, 242, 247) : Color.White);
                    Color text = Primary ? Color.White : TextPrimary;

                    using (var brush = new SolidBrush(fill)) g.FillPath(brush, path);
                    if (!Primary)
                    {
                        using (var pen = new Pen(ButtonBorder)) g.DrawPath(pen, path);
                    }

                    TextRenderer.DrawText(g, Text, Font, rect, text,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
                }
            }

            private static GraphicsPath GetRoundedPath(Rectangle bounds, int radius)
            {
                var path = new GraphicsPath();
                int d = radius * 2;
                path.AddArc(bounds.Left, bounds.Top, d, d, 180, 90);
                path.AddArc(bounds.Right - d, bounds.Top, d, d, 270, 90);
                path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
                path.AddArc(bounds.Left, bounds.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                return path;
            }
        }
    }
}