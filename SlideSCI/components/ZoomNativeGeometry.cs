using System;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using System.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    /// <summary>Native DrawingML connection sites, without auxiliary shapes or rasterization.</summary>
    internal static class ZoomNativeGeometry
    {
        internal static void ReplaceWithNativeSites(PowerPoint.Application app, PowerPoint.Slide slide,
            ref PowerPoint.Shape box, ref PowerPoint.Shape picture)
        {
            string path = Path.Combine(Path.GetTempPath(), "SlideSCIFocus-sites-" + Guid.NewGuid().ToString("N") + ".pptx");
            PowerPoint.Presentation scratch = null;
            PowerPoint.ShapeRange pasted = null;
            string boxName = box.Name, pictureName = picture.Name;
            float boxLeft = box.Left, boxTop = box.Top, picLeft = picture.Left, picTop = picture.Top;
            try
            {
                scratch = app.Presentations.Add(MsoTriState.msoFalse);
                var sourcePresentation = (PowerPoint.Presentation)slide.Parent;
                scratch.PageSetup.SlideWidth = sourcePresentation.PageSetup.SlideWidth;
                scratch.PageSetup.SlideHeight = sourcePresentation.PageSetup.SlideHeight;
                var temporarySlide = scratch.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                box.Copy();
                temporarySlide.Shapes.Paste();
                picture.Copy();
                temporarySlide.Shapes.Paste();
                scratch.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, MsoTriState.msoFalse);
                scratch.Close(); scratch = null;
                AddConnectionSites(path);
                scratch = app.Presentations.Open(path, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
                scratch.Slides[1].Shapes.Range().Copy();
                pasted = slide.Shapes.Paste();
                PowerPoint.Shape newBox = null, newPicture = null;
                foreach (PowerPoint.Shape shape in pasted)
                {
                    if (shape.Type == MsoShapeType.msoPicture || shape.Type == MsoShapeType.msoLinkedPicture) newPicture = shape;
                    else newBox = shape;
                }
                if (newBox == null || newPicture == null || newBox.ConnectionSiteCount != 8 || newPicture.ConnectionSiteCount != 8)
                    throw new InvalidOperationException("PowerPoint 未保留原生连接点。原选区框已保留。");
                newBox.Left = boxLeft; newBox.Top = boxTop;
                newPicture.Left = picLeft; newPicture.Top = picTop;
                newBox.Name = boxName; newPicture.Name = pictureName;
                box = newBox; picture = newPicture;
                pasted = null;
            }
            finally
            {
                if (pasted != null) { try { pasted.Delete(); } catch { } }
                if (scratch != null) { try { scratch.Close(); } catch { } }
                try { File.Delete(path); } catch { }
            }
        }

        internal static void AddConnectionSites(string path)
        {
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                var entry = archive.GetEntry("ppt/slides/slide1.xml");
                XDocument document;
                using (var stream = entry.Open()) document = XDocument.Load(stream);
                foreach (var geometry in document.Descendants().Where(e => e.Name == a + "prstGeom" || e.Name == a + "custGeom").ToList())
                {
                    string[,] coordinates = { { "0", "0" }, { "hc", "0" }, { "w", "0" }, { "w", "vc" }, { "w", "h" }, { "hc", "h" }, { "0", "h" }, { "0", "vc" } };
                    var sites = new XElement(a + "cxnLst");
                    for (int i = 0; i < 8; i++)
                        sites.Add(new XElement(a + "cxn", new XAttribute("ang", i < 3 ? "16200000" : i < 5 ? "0" : i < 7 ? "5400000" : "10800000"),
                            new XElement(a + "pos", new XAttribute("x", coordinates[i, 0]), new XAttribute("y", coordinates[i, 1]))));
                    var outline = new XElement(a + "path", new XAttribute("w", "21600"), new XAttribute("h", "21600"));
                    int[,] corners = { { 0, 0 }, { 21600, 0 }, { 21600, 21600 }, { 0, 21600 } };
                    for (int i = 0; i < 4; i++)
                        outline.Add(new XElement(a + (i == 0 ? "moveTo" : "lnTo"), new XElement(a + "pt", new XAttribute("x", corners[i, 0]), new XAttribute("y", corners[i, 1]))));
                    outline.Add(new XElement(a + "close"));
                    geometry.ReplaceWith(new XElement(a + "custGeom", new XElement(a + "avLst"), new XElement(a + "gdLst"), new XElement(a + "ahLst"), sites,
                        new XElement(a + "rect", new XAttribute("l", "0"), new XAttribute("t", "0"), new XAttribute("r", "w"), new XAttribute("b", "h")),
                        new XElement(a + "pathLst", outline)));
                }
                entry.Delete();
                using (var stream = archive.CreateEntry("ppt/slides/slide1.xml").Open()) document.Save(stream);
            }
        }
    }
}
