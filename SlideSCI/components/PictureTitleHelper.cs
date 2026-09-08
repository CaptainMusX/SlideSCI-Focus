using System;
using Microsoft.Office.Core;
using P = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    public enum PictureTitleSide { Top, Bottom, Left, Right }

    public static class PictureTitleHelper
    {
        public static P.Shape Add(P.Slide slide, P.Shape picture, PictureTitleSide side,
            string text, string font, float fontSize, float offsetX, float offsetY, P.PpParagraphAlignment alignment)
        {
            if (slide == null || picture == null) throw new ArgumentNullException("picture");
            if (!Enum.IsDefined(typeof(PictureTitleSide), side)) throw new ArgumentOutOfRangeException("side");
            if (!Finite(fontSize) || fontSize <= 0 || fontSize > 4000 || !Finite(offsetX) || !Finite(offsetY))
                throw new ArgumentOutOfRangeException("fontSize", "字号和偏移必须是有效数字。");
            if (alignment != P.PpParagraphAlignment.ppAlignLeft && alignment != P.PpParagraphAlignment.ppAlignCenter &&
                alignment != P.PpParagraphAlignment.ppAlignRight && alignment != P.PpParagraphAlignment.ppAlignJustify)
                throw new ArgumentOutOfRangeException("alignment");
            P.Shape title = null;
            try
            {
                title = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, picture.Left, picture.Top, picture.Width, fontSize * 2);
                title.Name = "SlideSCIFocus_PictureTitle_" + Guid.NewGuid().ToString("N");
                var frame = title.TextFrame;
                frame.MarginLeft = frame.MarginRight = frame.MarginTop = frame.MarginBottom = 0;
                frame.TextRange.Text = text ?? "";
                frame.TextRange.Font.Name = font;
                frame.TextRange.Font.NameFarEast = font;
                frame.TextRange.Font.Size = fontSize;
                frame.TextRange.ParagraphFormat.Alignment = alignment;
                frame.WordWrap = MsoTriState.msoTrue;
                frame.AutoSize = P.PpAutoSize.ppAutoSizeShapeToFitText;
                title.Width = picture.Width;
                if (side == PictureTitleSide.Left || side == PictureTitleSide.Right)
                    title.Width = Math.Max(fontSize, Math.Min(picture.Width, frame.TextRange.BoundWidth + 1f));
                // Measure after formatting and wrapping, then place the final text box.
                float left = picture.Left, top = picture.Top;
                switch (side)
                {
                    case PictureTitleSide.Top: top -= title.Height; break;
                    case PictureTitleSide.Bottom: top += picture.Height; break;
                    case PictureTitleSide.Left: left -= title.Width; top += (picture.Height - title.Height) / 2; break;
                    case PictureTitleSide.Right: left += picture.Width; top += (picture.Height - title.Height) / 2; break;
                }
                title.Left = left + offsetX;
                title.Top = top + offsetY;
                return title;
            }
            catch { if (title != null) { try { title.Delete(); } catch { } } throw; }
        }
        private static bool Finite(float value) { return !float.IsNaN(value) && !float.IsInfinity(value); }
    }
}
