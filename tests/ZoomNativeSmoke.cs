using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using SlideSCI;
using Microsoft.Office.Core;
using P = Microsoft.Office.Interop.PowerPoint;
class ZoomNativeSmoke
{
    static void Check(bool ok, string name) { if (!ok) throw new Exception(name); Console.WriteLine("PASS " + name); }
    static float[] Crop(P.Shape source, P.Shape box)
    {
        object[] args = { source, box, null, null };
        var method = typeof(ZoomInsetHelper).GetMethod("TryBuildCropGeometry", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
        Check((bool)method.Invoke(null, args), "crop geometry accepted");
        var values = new List<float>();
        foreach (string field in new[] { "Left", "Right", "Top", "Bottom" })
            values.Add((float)args[2].GetType().GetField(field, System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(args[2]));
        return values.ToArray();
    }
    static List<P.Shape> All(P.Slide slide)
    {
        var list = new List<P.Shape>();
        foreach (P.Shape shape in slide.Shapes) Collect(shape, list);
        return list;
    }
    static void Collect(P.Shape shape, List<P.Shape> list)
    {
        list.Add(shape);
        if (shape.Type == MsoShapeType.msoGroup)
            foreach (P.Shape child in shape.GroupItems) Collect(child, list);
    }
    [STAThread] static int Main(string[] args)
    {
        P.Application app = new P.Application();
        P.Presentation doc = app.Presentations.Add(MsoTriState.msoFalse);
        try
        {
            string image = Path.Combine(args[0], "zoom-test-source.png");
            using (var bitmap = new Bitmap(1600, 1600))
            { using (var g = Graphics.FromImage(bitmap)) { g.Clear(Color.White); g.FillEllipse(Brushes.Green, 400, 400, 600, 600); } bitmap.Save(image); }
            P.Slide slide = doc.Slides.Add(1, P.PpSlideLayout.ppLayoutBlank);
            P.Shape source = slide.Shapes.AddPicture(image, MsoTriState.msoFalse, MsoTriState.msoTrue, 100, 60, 140, 140);
            P.Shape box = ZoomInsetHelper.InsertZoomBox(slide, source, 40);
            string name = box.Name;
            foreach (float factor in new[] { 1f, 0.5f, 1f / 3f, 2f })
            {
                box.Height = 28;
                float width, height;
                ZoomInsetHelper.ComputeTargetSize(source, box, ZoomTargetMode.Multiple, factor, 10, out width, out height);
                Check(Math.Abs(width - source.Width * factor) < 0.01f && Math.Abs(height / width - 0.5f) < 0.001f, "original width preset and rectangular aspect " + factor);
            }
            box.Height = box.Width;
            box.Left = source.Left + 10; box.Top = source.Top + 15;
            float[] normalCrop = Crop(source, box);
            source.Flip(MsoFlipCmd.msoFlipHorizontal);
            source.Flip(MsoFlipCmd.msoFlipVertical);
            float[] flippedCrop = Crop(source, box);
            Check(Math.Abs(normalCrop[0] - flippedCrop[1]) < 0.01f && Math.Abs(normalCrop[2] - flippedCrop[3]) < 0.01f, "flipped crop follows displayed selection");
            source.Flip(MsoFlipCmd.msoFlipHorizontal); source.Flip(MsoFlipCmd.msoFlipVertical);
            source.PictureFormat.CropLeft = -10;
            box.Left = source.Left;
            Check(Math.Abs(Crop(source, box)[0] + 10) < 0.01f, "negative crop margin preserved");
            source.PictureFormat.CropLeft = 0;
            box.Left = source.Left + 10;
            for (int round = 0; round < 5; round++)
            {
                box = All(slide).Find(b => b.Name == name);
                var strokes = new ZoomSettings { BoxLineWeight = 4.5f, LineWeight = 2.25f,
                    BoxColorRgb = 0x123456, LineColorRgb = 0x654321, LineDash = 6,
                    BoxLineVisible = round != 2, LineVisible = round != 2, LineEndArrow = 2 };
                var result = ZoomInsetHelper.GenerateZoomInset(slide, box, source, 140, 140, 10, round == 3 ? ZoomLineStyle.CrossedX : round == 4 ? ZoomLineStyle.None : ZoomLineStyle.JournalFunnel,
                    1, 1.5f, 0, 0, MsoLineDashStyle.msoLineDash, round == 1, app, MsoLineDashStyle.msoLineDash, strokes);
                Check(result.Ok, "generation round " + round + ": " + result.Error);
                Check(result.Glued, "native glued endpoints");
                var shapes = All(slide);
                int lines = 0, pictures = 0, groups = 0, boxes = 0;
                P.Shape inset = null;
                foreach (P.Shape shape in shapes)
                {
                    if (shape.Type == MsoShapeType.msoGroup) groups++;
                    else if (shape.Connector == MsoTriState.msoTrue)
                    {
                        lines++;
                        Check(Math.Abs(shape.Line.Weight - 2.25f) < 0.01f && shape.Line.ForeColor.RGB == 0x654321, "connector weight and color applied");
                        Check(shape.Line.DashStyle == MsoLineDashStyle.msoLineLongDash && shape.Line.EndArrowheadStyle == MsoArrowheadStyle.msoArrowheadTriangle, "connector dash and arrow applied");
                        Check((shape.Line.Visible == MsoTriState.msoTrue) == (round != 2), "connector no-outline state applied");
                        Check(shape.ConnectorFormat.BeginConnected == MsoTriState.msoTrue && shape.ConnectorFormat.EndConnected == MsoTriState.msoTrue, "connector attachment");
                        Check(shape.ConnectorFormat.BeginConnectedShape.ConnectionSiteCount == 8 && shape.ConnectorFormat.EndConnectedShape.ConnectionSiteCount == 8, "eight native sites on both ends");
                        Console.WriteLine("SITES " + shape.ConnectorFormat.BeginConnectionSite + " -> " + shape.ConnectorFormat.EndConnectionSite);
                    }
                    else if (shape.Type == MsoShapeType.msoPicture) { pictures++; if (shape.Id != source.Id) inset = shape; }
                    else if (ZoomInsetHelper.IsZoomBox(shape))
                    {
                        boxes++;
                        Check(Math.Abs(shape.Line.Weight - 4.5f) < 0.01f && shape.Line.ForeColor.RGB == 0x123456, "box stroke applied on update");
                        Check((shape.Line.Visible == MsoTriState.msoTrue) == (round != 2), "box no-outline state applied");
                    }
                    else throw new Exception("Redundant shape " + shape.Name);
                }
                Check(lines == (round == 4 ? 0 : 2) && pictures == 2 && boxes == 1 && groups == (round == 1 ? 1 : 0), "only source, box, picture, two connectors and optional single group");
                Check(Math.Abs(inset.Width - 140) < 0.1f, "inset dimensions");
                if (round == 0)
                {
                    inset.Left += 30;
                    foreach (P.Shape shape in shapes)
                        if (shape.Connector == MsoTriState.msoTrue)
                            Check(shape.ConnectorFormat.EndConnectedShape.Id == inset.Id, "attachment survives picture movement");
                }
            }
            box = All(slide).Find(b => b.Name == name);
            int previousCount = All(slide).Count;
            string oldTemp = Environment.GetEnvironmentVariable("TEMP");
            string oldTmp = Environment.GetEnvironmentVariable("TMP");
            string blockedTemp = Path.Combine(args[0], "blocked-temp-file");
            File.WriteAllText(blockedTemp, "This is a file, not a temporary directory.");
            try
            {
                Environment.SetEnvironmentVariable("TEMP", blockedTemp);
                Environment.SetEnvironmentVariable("TMP", blockedTemp);
                var failed = ZoomInsetHelper.GenerateZoomInset(slide, box, source, 140, 140, 10, ZoomLineStyle.JournalFunnel,
                    1, 1.5f, 0, 0, MsoLineDashStyle.msoLineDash, false, app);
                Check(failed.Ok && !string.IsNullOrEmpty(failed.Warning), "staging save failure falls back with warning");
                Check(All(slide).Count == previousCount + 2 && All(slide).Exists(b => b.Name == name), "fallback replaces previous inset without draft debris");
            }
            finally { Environment.SetEnvironmentVariable("TEMP", oldTemp); Environment.SetEnvironmentVariable("TMP", oldTmp); }
            box = All(slide).Find(b => b.Name == name);
            previousCount = All(slide).Count;
            box.Left = source.Left - 20;
            var invalid = ZoomInsetHelper.GenerateZoomInset(slide, box, source, 140, 140, 10, ZoomLineStyle.JournalFunnel,
                1, 1.5f, 0, 0, MsoLineDashStyle.msoLineDash, false, app);
            Check(!invalid.Ok && All(slide).Count == previousCount, "invalid selection preserves existing inset");
            box.Left = source.Left + 20;
            doc.SaveAs(Path.Combine(args[0], "zoom-native-test.pptx"), P.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
            slide.Export(Path.Combine(args[0], "zoom-native-test.png"), "PNG", 1200, 900);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally { doc.Close(); }
    }
}
