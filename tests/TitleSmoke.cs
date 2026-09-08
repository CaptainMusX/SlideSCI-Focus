using System;
using System.IO;
using System.Drawing;
using System.Reflection;
using SlideSCI;
using Microsoft.Office.Core;
using P = Microsoft.Office.Interop.PowerPoint;
class TitleSmoke
{
    static int count;
    static void Check(bool ok, string name) { if (!ok) throw new Exception(name); count++; Console.WriteLine("PASS " + name); }
    static bool Near(float a, float b) { return Math.Abs(a-b) < 0.2f; }
    [STAThread] static int Main(string[] args)
    {
        P.Application app = new P.Application();
        P.Presentation doc = app.Presentations.Add(MsoTriState.msoFalse);
        try
        {
            foreach (string icon in new[] { "AlignLeft", "AlignCenter", "AlignRight", "AlignJustify", "TextBoxInsert" })
                Check(!string.IsNullOrEmpty(app.CommandBars.GetLabelMso(icon)), "native Office command " + icon);
            var settings = typeof(PictureTitleHelper).Assembly.GetType("SlideSCI.Properties.Settings");
            foreach (string name in new[] { "TitleOffsetX", "TitleOffsetY" })
            {
                var property = settings.GetProperty(name);
                var value = (System.Configuration.DefaultSettingValueAttribute)Attribute.GetCustomAttribute(property, typeof(System.Configuration.DefaultSettingValueAttribute));
                Check(value.Value == "0", "zero default " + name);
            }
            string file = Path.Combine(args[0], "source.png");
            using (var bitmap = new Bitmap(160, 100)) { using (var g = Graphics.FromImage(bitmap)) g.Clear(Color.SteelBlue); bitmap.Save(file); }
            var slide = doc.Slides.Add(1, P.PpSlideLayout.ppLayoutBlank);
            var source = slide.Shapes.AddPicture(file, MsoTriState.msoFalse, MsoTriState.msoTrue, 300, 220, 160, 100);
            foreach (PictureTitleSide side in Enum.GetValues(typeof(PictureTitleSide)))
                for (int alignment = 1; alignment <= 4; alignment++)
                {
                    string text = "第一行中文标题\nSecond line of title";
                    var baseline = PictureTitleHelper.Add(slide, source, side, text, "Arial", 18, 0, 0, (P.PpParagraphAlignment)alignment);
                    var moved = PictureTitleHelper.Add(slide, source, side, text, "Arial", 18, 13, -7, (P.PpParagraphAlignment)alignment);
                    Check(Near(moved.Left, baseline.Left + 13) && Near(moved.Top, baseline.Top - 7), side + " additive signed XY offsets " + alignment);
                    Check((int)moved.TextFrame.TextRange.ParagraphFormat.Alignment == alignment, side + " paragraph alignment " + alignment);
                    Check(Near(moved.Width, baseline.Width) && Near(moved.Height, baseline.Height), "offsets preserve text dimensions");
                    if (side == PictureTitleSide.Top) Check(Near(baseline.Top + baseline.Height, source.Top), "multiline top anchored to actual height");
                    if (side == PictureTitleSide.Bottom) Check(Near(baseline.Top, source.Top + source.Height), "bottom anchor");
                    if (side == PictureTitleSide.Left) Check(Near(baseline.Left + baseline.Width, source.Left) && Near(baseline.Top + baseline.Height/2, source.Top+source.Height/2), "left edge and vertical center");
                    if (side == PictureTitleSide.Right) Check(Near(baseline.Left, source.Left+source.Width) && Near(baseline.Top + baseline.Height/2, source.Top+source.Height/2), "right edge and vertical center");
                    baseline.Delete(); moved.Delete();
                }
            int before = slide.Shapes.Count;
            try { PictureTitleHelper.Add(slide, source, PictureTitleSide.Top, "bad", "Arial", 18, float.NaN, 0, P.PpParagraphAlignment.ppAlignCenter); throw new Exception("Expected rejection"); }
            catch (ArgumentOutOfRangeException) { Check(slide.Shapes.Count == before, "invalid offsets create no shapes"); }
            foreach (PictureTitleSide side in Enum.GetValues(typeof(PictureTitleSide)))
                PictureTitleHelper.Add(slide, source, side, "标题 " + side, "微软雅黑", 18, 0, 0, P.PpParagraphAlignment.ppAlignCenter);
            doc.SaveAs(Path.Combine(args[0], "four-side-titles.pptx"), P.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
            slide.Export(Path.Combine(args[0], "four-side-titles.png"), "PNG", 1280, (int)(1280 * doc.PageSetup.SlideHeight / doc.PageSetup.SlideWidth));
            Console.WriteLine("TOTAL PASS " + count);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { doc.Close(); }
    }
}
