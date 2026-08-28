using System;
using System.Drawing;
using System.Windows.Forms;

namespace SlideSCI
{
    /// <summary>
    /// 「生成放大图」参数对话框。
    /// 各选项通过属性暴露给调用方；尺寸计算由 ZoomInsetHelper.ComputeTargetSize 完成。
    /// </summary>
    public class ZoomInsetForm : Form
    {
        private RadioButton rbSameAsOriginal;
        private RadioButton rbMultiple;
        private TextBox txtMagnification;
        private RadioButton rbCustomWidth;
        private TextBox txtCustomWidth;
        private RadioButton rbJournalFunnel;
        private RadioButton rbCrossedX;
        private ComboBox cmbLineWeight;
        private ComboBox cmbBoxLineWeight;
        private TextBox txtGap;
        private CheckBox chkGroup;
        private Button btnOk;
        private Button btnCancel;

        public ZoomTargetMode TargetMode { get; private set; }
        public float Magnification { get; private set; }
        public float CustomWidthCm { get; private set; }
        public float GapCm { get; private set; }
        public ZoomLineStyle LineStyle { get; private set; }
        public float LineWeight { get; private set; }
        public float BoxLineWeight { get; private set; }
        public bool GroupEnabled { get; private set; }

        public ZoomInsetForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "生成局部放大图";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ClientSize = new Size(360, 430);

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

            rbSameAsOriginal.CheckedChanged += (s, e) => txtMagnification.Enabled = rbMultiple.Checked;
            rbSameAsOriginal.CheckedChanged += (s, e) => txtCustomWidth.Enabled = rbCustomWidth.Checked;

            groupTarget.Controls.Add(rbSameAsOriginal);
            groupTarget.Controls.Add(rbMultiple);
            groupTarget.Controls.Add(txtMagnification);
            groupTarget.Controls.Add(rbCustomWidth);
            groupTarget.Controls.Add(txtCustomWidth);

            var groupLine = new GroupBox { Text = "框角连线样式", Bounds = new Rectangle(12, 150, 336, 96) };
            rbJournalFunnel = new RadioButton
            {
                Text = "期刊漏斗：框下两角 → 放大图上两角（平行）",
                Bounds = new Rectangle(16, 24, 310, 40),
                Checked = true
            };
            rbCrossedX = new RadioButton
            {
                Text = "交叉 X 形：四角交叉连线",
                Bounds = new Rectangle(16, 64, 300, 22)
            };
            groupLine.Controls.Add(rbJournalFunnel);
            groupLine.Controls.Add(rbCrossedX);

            var lblLineWeight = new Label { Text = "连线粗细(pt)", Bounds = new Rectangle(12, 260, 110, 24) };
            cmbLineWeight = new ComboBox
            {
                Bounds = new Rectangle(130, 258, 64, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbLineWeight.Items.AddRange(new object[] { 0.5f, 1f, 1.5f, 2f });
            cmbLineWeight.SelectedIndex = 1;

            var lblBoxWeight = new Label { Text = "框线粗细(pt)", Bounds = new Rectangle(210, 260, 100, 24) };
            cmbBoxLineWeight = new ComboBox
            {
                Bounds = new Rectangle(300, 258, 48, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbBoxLineWeight.Items.AddRange(new object[] { 1f, 1.5f, 2f, 3f });
            cmbBoxLineWeight.SelectedIndex = 1;

            var lblGap = new Label { Text = "放大图下边距(cm)", Bounds = new Rectangle(12, 296, 130, 24) };
            txtGap = new TextBox { Text = "0.5", Bounds = new Rectangle(150, 294, 56, 24) };

            chkGroup = new CheckBox
            {
                Text = "生成后自动编组并选中",
                Bounds = new Rectangle(12, 330, 240, 26),
                Checked = true
            };

            btnOk = new Button { Text = "生成", Bounds = new Rectangle(168, 384, 84, 30) };
            btnOk.Click += (s, e) => { if (ValidateAndRead()) this.DialogResult = DialogResult.OK; };

            btnCancel = new Button { Text = "取消", Bounds = new Rectangle(262, 384, 84, 30) };
            btnCancel.Click += (s, e) => this.DialogResult = DialogResult.Cancel;

            this.Controls.Add(groupTarget);
            this.Controls.Add(groupLine);
            this.Controls.Add(lblLineWeight);
            this.Controls.Add(cmbLineWeight);
            this.Controls.Add(lblBoxWeight);
            this.Controls.Add(cmbBoxLineWeight);
            this.Controls.Add(lblGap);
            this.Controls.Add(txtGap);
            this.Controls.Add(chkGroup);
            this.Controls.Add(btnOk);
            this.Controls.Add(btnCancel);
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
        }

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
            GroupEnabled = chkGroup.Checked;
            return true;
        }
    }
}