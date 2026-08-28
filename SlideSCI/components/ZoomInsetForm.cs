using System;
using System.Drawing;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    /// <summary>
    /// 「生成放大图」参数对话框。
    /// 选项通过属性暴露给调用方；尺寸计算由 ZoomInsetHelper.ComputeTargetSize 完成。
    /// 构造时传入当前选区框与原图尺寸（磅），用于实时预览放大图尺寸/放大倍数。
    /// </summary>
    public class ZoomInsetForm : Form
    {
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
        private Button btnOk;
        private Button btnCancel;

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
            InitializeComponent();
            UpdatePreview();
        }

        private void InitializeComponent()
        {
            this.Text = "生成局部放大图";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(360, 512);

            // —— 放大目标 ——
            var groupTarget = new GroupBox { Text = "放大目标", Bounds = new Rectangle(12, 12, 336, 128) };
            rbSameAsOriginal = new RadioButton
            {
                Text = "与原图相同尺寸",
                Bounds = new Rectangle(16, 26, 300, 22),
                Checked = true
            };
            rbMultiple = new RadioButton { Text = "指定放大倍数", Bounds = new Rectangle(16, 52, 140, 22) };
            txtMagnification = new TextBox { Text = "2", Bounds = new Rectangle(160, 52, 56, 24), Enabled = false };
            rbCustomWidth = new RadioButton { Text = "自定义宽度(cm)", Bounds = new Rectangle(16, 80, 140, 22) };
            txtCustomWidth = new TextBox { Text = "10", Bounds = new Rectangle(160, 80, 56, 24), Enabled = false };

            rbSameAsOriginal.CheckedChanged += (s, e) => { SyncInputsEnabled(); UpdatePreview(); };
            rbMultiple.CheckedChanged += (s, e) => { SyncInputsEnabled(); UpdatePreview(); };
            rbCustomWidth.CheckedChanged += (s, e) => { SyncInputsEnabled(); UpdatePreview(); };
            txtMagnification.TextChanged += (s, e) => UpdatePreview();
            txtCustomWidth.TextChanged += (s, e) => UpdatePreview();

            groupTarget.Controls.Add(rbSameAsOriginal);
            groupTarget.Controls.Add(rbMultiple);
            groupTarget.Controls.Add(txtMagnification);
            groupTarget.Controls.Add(rbCustomWidth);
            groupTarget.Controls.Add(txtCustomWidth);

            // —— 连线样式 ——
            var groupLine = new GroupBox { Text = "框角连线样式", Bounds = new Rectangle(12, 150, 336, 96) };
            rbJournalFunnel = new RadioButton
            {
                Text = "期刊漏斗：框下两角 → 放大图上两角",
                Bounds = new Rectangle(16, 24, 310, 36),
                Checked = true
            };
            rbCrossedX = new RadioButton
            {
                Text = "交叉 X 形：四角交叉连线",
                Bounds = new Rectangle(16, 60, 300, 22)
            };
            groupLine.Controls.Add(rbJournalFunnel);
            groupLine.Controls.Add(rbCrossedX);

            // —— 线宽 ——
            var lblLineWeight = new Label { Text = "连线粗细(pt)", Bounds = new Rectangle(12, 258, 110, 24) };
            cmbLineWeight = new ComboBox
            {
                Bounds = new Rectangle(130, 256, 64, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbLineWeight.Items.AddRange(new object[] { 0.5f, 1f, 1.5f, 2f });
            cmbLineWeight.SelectedIndex = 1;

            var lblBoxWeight = new Label { Text = "框线粗细(pt)", Bounds = new Rectangle(210, 258, 100, 24) };
            cmbBoxLineWeight = new ComboBox
            {
                Bounds = new Rectangle(300, 256, 48, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbBoxLineWeight.Items.AddRange(new object[] { 1f, 1.5f, 2f, 3f });
            cmbBoxLineWeight.SelectedIndex = 1;

            // —— 颜色 ——
            var lblLineColor = new Label { Text = "连线颜色", Bounds = new Rectangle(12, 292, 100, 24) };
            cmbLineColor = new ComboBox
            {
                Bounds = new Rectangle(120, 290, 96, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbLineColor.Items.AddRange(new object[] { "黑色", "白色", "红色", "蓝色", "绿色", CustomItem });
            cmbLineColor.SelectedIndex = 0;
            cmbLineColor.SelectedIndexChanged += (s, e) => HandleColorSelection(cmbLineColor, ref customLineColorRgb);

            var lblBoxColor = new Label { Text = "框线颜色", Bounds = new Rectangle(230, 292, 80, 24) };
            cmbBoxColor = new ComboBox
            {
                Bounds = new Rectangle(304, 290, 44, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbBoxColor.Items.AddRange(new object[] { "黑色", "红色", "蓝色", "绿色" });
            cmbBoxColor.SelectedIndex = 0;

            // —— 线型 ——
            var lblDash = new Label { Text = "连线线型", Bounds = new Rectangle(12, 326, 100, 24) };
            cmbLineDash = new ComboBox
            {
                Bounds = new Rectangle(120, 324, 96, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbLineDash.Items.AddRange(new object[] { "实线", "虚线" });
            cmbLineDash.SelectedIndex = 0;

            // —— 下边距 ——
            var lblGap = new Label { Text = "放大图下边距(cm)", Bounds = new Rectangle(12, 358, 130, 24) };
            txtGap = new TextBox { Text = "0.5", Bounds = new Rectangle(150, 356, 56, 24) };

            // —— 实时预览 ——
            lblPreview = new Label
            {
                Bounds = new Rectangle(12, 388, 336, 46),
                ForeColor = Color.Gray,
                Text = ""
            };

            chkGroup = new CheckBox
            {
                Text = "生成后自动编组（框+放大图+连线整体移动）",
                Bounds = new Rectangle(12, 440, 320, 26),
                Checked = true
            };

            btnOk = new Button { Text = "生成", Bounds = new Rectangle(168, 472, 84, 30) };
            btnOk.Click += (s, e) => { if (ValidateAndRead()) this.DialogResult = DialogResult.OK; };

            btnCancel = new Button { Text = "取消", Bounds = new Rectangle(262, 472, 84, 30) };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            this.Controls.Add(groupTarget);
            this.Controls.Add(groupLine);
            this.Controls.Add(lblLineWeight);
            this.Controls.Add(cmbLineWeight);
            this.Controls.Add(lblBoxWeight);
            this.Controls.Add(cmbBoxLineWeight);
            this.Controls.Add(lblLineColor);
            this.Controls.Add(cmbLineColor);
            this.Controls.Add(lblBoxColor);
            this.Controls.Add(cmbBoxColor);
            this.Controls.Add(lblDash);
            this.Controls.Add(cmbLineDash);
            this.Controls.Add(lblGap);
            this.Controls.Add(txtGap);
            this.Controls.Add(lblPreview);
            this.Controls.Add(chkGroup);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
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
            lblPreview.ForeColor = stretched ? Color.FromArgb(200, 120, 0) : Color.Gray;
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
    }
}