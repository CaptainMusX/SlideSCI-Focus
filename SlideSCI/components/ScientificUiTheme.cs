using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace SlideSCI
{
    /// <summary>
    /// Shared visual tokens for SlideSCI Focus desktop dialogs. The palette is tuned
    /// for a dense biomedical research tool rather than a consumer/mobile UI.
    /// </summary>
    internal static class ScientificUiTheme
    {
        internal static readonly Color Primary = Color.FromArgb(30, 58, 95);
        internal static readonly Color PrimaryHover = Color.FromArgb(23, 49, 82);
        internal static readonly Color Action = Color.FromArgb(37, 99, 235);
        internal static readonly Color ActionHover = Color.FromArgb(29, 78, 216);
        internal static readonly Color Background = Color.FromArgb(248, 250, 252);
        internal static readonly Color Surface = Color.White;
        internal static readonly Color SurfaceMuted = Color.FromArgb(241, 245, 249);
        internal static readonly Color TextPrimary = Color.FromArgb(15, 23, 42);
        internal static readonly Color TextSecondary = Color.FromArgb(71, 85, 105);
        internal static readonly Color Border = Color.FromArgb(203, 213, 225);
        internal static readonly Color Focus = Color.FromArgb(37, 99, 235);
        internal static readonly Color Warning = Color.FromArgb(161, 98, 7);
        internal static readonly Color Error = Color.FromArgb(220, 38, 38);

        internal static Font BodyFont(float size = 9f, FontStyle style = FontStyle.Regular)
        {
            return new Font("Segoe UI", size, style, GraphicsUnit.Point);
        }

        internal static Font ChineseBodyFont(float size = 9f, FontStyle style = FontStyle.Regular)
        {
            return new Font("Microsoft YaHei UI", size, style, GraphicsUnit.Point);
        }

        internal static void ConfigureDialog(Form form)
        {
            form.AutoScaleMode = AutoScaleMode.Dpi;
            form.BackColor = Background;
            form.ForeColor = TextPrimary;
            form.Font = ChineseBodyFont();
            form.ShowIcon = false;
            form.ShowInTaskbar = false;
            form.StartPosition = FormStartPosition.CenterParent;
            form.KeyPreview = true;
        }

        /// <summary>
        /// Arms the first-show DPI scaling after a code-built dialog has added
        /// all controls and finished any text-dependent layout.  Calling this
        /// earlier lets WinForms consume the scale pass before child controls
        /// exist, which leaves fixed bounds at 96 DPI while fonts render at the
        /// monitor DPI.
        /// </summary>
        internal static void CompleteCodeBuiltLayout(Form form)
        {
            if (form == null) throw new ArgumentNullException(nameof(form));
            form.AutoScaleDimensions = new SizeF(96f, 96f);
        }

        internal static GraphicsPath RoundedPath(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            if (bounds.Width <= 0 || bounds.Height <= 0)
            {
                return path;
            }

            int diameter = Math.Max(2, radius * 2);
            diameter = Math.Min(diameter, Math.Min(bounds.Width, bounds.Height));
            path.AddArc(bounds.Left, bounds.Top, diameter, diameter, 180, 90);
            path.AddArc(bounds.Right - diameter, bounds.Top, diameter, diameter, 270, 90);
            path.AddArc(bounds.Right - diameter, bounds.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(bounds.Left, bounds.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    internal sealed class ScientificCard : Panel
    {
        internal int CornerRadius { get; set; } = 10;

        internal ScientificCard()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw |
                     ControlStyles.UserPaint, true);
            BackColor = ScientificUiTheme.Surface;
        }

        protected override void OnPaintBackground(PaintEventArgs e)
        {
            e.Graphics.Clear(Parent?.BackColor ?? ScientificUiTheme.Background);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath path = ScientificUiTheme.RoundedPath(
                new Rectangle(0, 0, Width - 1, Height - 1), CornerRadius))
            using (var fill = new SolidBrush(BackColor))
            using (var border = new Pen(ScientificUiTheme.Border))
            {
                e.Graphics.FillPath(fill, path);
                e.Graphics.DrawPath(border, path);
            }

            base.OnPaint(e);
        }
    }

    internal sealed class ScientificButton : Button
    {
        private bool hovered;
        private bool pressed;

        internal bool Primary { get; set; }

        internal ScientificButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.ResizeRedraw |
                     ControlStyles.UserPaint, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            Height = 36;
            Font = ScientificUiTheme.ChineseBodyFont(9f, FontStyle.Bold);
            AccessibleRole = AccessibleRole.PushButton;
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            hovered = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            hovered = false;
            pressed = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            pressed = true;
            Invalidate();
            base.OnMouseDown(mevent);
        }

        protected override void OnMouseUp(MouseEventArgs mevent)
        {
            pressed = false;
            Invalidate();
            base.OnMouseUp(mevent);
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            Cursor = Enabled ? Cursors.Hand : Cursors.Default;
            Invalidate();
            base.OnEnabledChanged(e);
        }

        protected override void OnGotFocus(EventArgs e)
        {
            Invalidate();
            base.OnGotFocus(e);
        }

        protected override void OnLostFocus(EventArgs e)
        {
            Invalidate();
            base.OnLostFocus(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Rectangle rect = new Rectangle(0, 0, Width - 1, Height - 1);
            Color fill;
            Color foreground;
            Color borderColor;

            if (!Enabled)
            {
                fill = ScientificUiTheme.SurfaceMuted;
                foreground = Color.FromArgb(148, 163, 184);
                borderColor = ScientificUiTheme.Border;
            }
            else if (Primary)
            {
                fill = pressed ? ScientificUiTheme.PrimaryHover
                    : hovered ? ScientificUiTheme.ActionHover
                    : ScientificUiTheme.Action;
                foreground = Color.White;
                borderColor = fill;
            }
            else
            {
                fill = pressed ? Color.FromArgb(226, 232, 240)
                    : hovered ? ScientificUiTheme.SurfaceMuted
                    : ScientificUiTheme.Surface;
                foreground = ScientificUiTheme.TextPrimary;
                borderColor = ScientificUiTheme.Border;
            }

            using (GraphicsPath path = ScientificUiTheme.RoundedPath(rect, 8))
            using (var brush = new SolidBrush(fill))
            using (var pen = new Pen(borderColor))
            {
                e.Graphics.FillPath(brush, path);
                e.Graphics.DrawPath(pen, path);
            }

            TextRenderer.DrawText(e.Graphics, Text, Font, rect, foreground,
                TextFormatFlags.HorizontalCenter |
                TextFormatFlags.VerticalCenter |
                TextFormatFlags.EndEllipsis |
                TextFormatFlags.NoPadding);

            if (Focused && ShowFocusCues)
            {
                Rectangle focusRect = Rectangle.Inflate(rect, -4, -4);
                ControlPaint.DrawFocusRectangle(e.Graphics, focusRect,
                    Primary ? Color.White : ScientificUiTheme.Focus, fill);
            }
        }
    }
}
