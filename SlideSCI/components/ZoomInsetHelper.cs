using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace SlideSCI
{
    public enum ZoomTargetMode
    {
        /// <summary>Use the source picture width and preserve the selection aspect ratio.</summary>
        SameAsOriginal = 0,
        /// <summary>Scale source picture width and preserve selection aspect ratio.</summary>
        Multiple = 1,
        /// <summary>Use a width in centimetres and preserve the selection aspect ratio.</summary>
        CustomWidthCm = 2
    }

    public enum ZoomLineStyle
    {
        JournalFunnel = 0,
        CrossedX = 1,
        None = 2
    }

    public sealed class ZoomInsetResult
    {
        public bool Ok { get; set; }
        public string Error { get; set; }
        public string Warning { get; set; }
        public List<PowerPoint.Shape> Created { get; } = new List<PowerPoint.Shape>();
        public PowerPoint.Shape Group { get; set; }
        public bool Glued { get; set; } = true;
    }

    /// <summary>
    /// Non-destructive PowerPoint shape operations for journal-style zoom insets.
    /// Crop values are calculated in original-image points, matching the Office
    /// object model, and each selection box owns an independent artifact set.
    /// </summary>
    public static class ZoomInsetHelper
    {
        public const string BoxNamePrefix = "SlideSCIFocus_ZoomBox_";
        public const string InsetNamePrefix = "SlideSCIFocus_ZoomInset_";
        private const string LegacyBoxNamePrefix = "SlideSCI_ZoomBox_";
        private const string LegacyInsetNamePrefix = "SlideSCI_ZoomInset_";

        private const float PointsPerCm = 28.3464593f;
        private const float GeometryTolerance = 0.5f;
        private const float SlideMarginPoints = 6f;

        public static float CmToPoints(float cm) => cm * PointsPerCm;

        public static PowerPoint.Shape InsertZoomBox(PowerPoint.Slide slide,
            PowerPoint.Shape picture, float boxSidePercent, float boxLineWeight = 1.5f)
        {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            if (picture == null) throw new ArgumentNullException(nameof(picture));
            if (picture.Width <= 0 || picture.Height <= 0)
            {
                throw new InvalidOperationException("图片尺寸无效，无法插入选区框。");
            }

            float percent = Clamp(boxSidePercent, 5f, 90f);
            float side = Math.Max(6f, Math.Min(picture.Width, picture.Height) * percent / 100f);
            float left = picture.Left + (picture.Width - side) / 2f;
            float top = picture.Top + (picture.Height - side) / 2f;

            PowerPoint.Shape box = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle, left, top, side, side);
            string uniquePart = Guid.NewGuid().ToString("N").Substring(0, 8);
            box.Name = BoxNamePrefix + picture.Id.ToString(CultureInfo.InvariantCulture) + "_" + uniquePart;
            box.Fill.Visible = MsoTriState.msoFalse;
            box.Line.ForeColor.RGB = 0x000000;
            box.Line.Weight = boxLineWeight;
            return box;
        }

        public static PowerPoint.Shape FindZoomBox(PowerPoint.Slide slide)
        {
            List<PowerPoint.Shape> boxes = FindZoomBoxes(slide);
            return boxes.Count > 0 ? boxes[0] : null;
        }

        public static List<PowerPoint.Shape> FindZoomBoxes(PowerPoint.Slide slide)
        {
            var boxes = new List<PowerPoint.Shape>();
            if (slide == null) return boxes;

            foreach (PowerPoint.Shape shape in CollectAllShapes(slide))
            {
                if (IsZoomBox(shape)) boxes.Add(shape);
            }
            return boxes;
        }

        /// <summary>
        /// Prefer a box in the current selection. If none is selected, fall
        /// back only when the slide contains exactly one box.
        /// </summary>
        public static bool TryResolveZoomBox(PowerPoint.Slide slide,
            PowerPoint.Selection selection, out PowerPoint.Shape box, out string error)
        {
            box = null;
            error = null;
            var selectedBoxes = new List<PowerPoint.Shape>();
            var selectedIds = new HashSet<int>();

            try
            {
                if (selection != null && selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shape in selection.ShapeRange)
                    {
                        CollectZoomBoxesFromShape(shape, selectedBoxes, selectedIds);
                    }
                }
            }
            catch { }

            if (selectedBoxes.Count == 0 && selection != null)
            {
                try
                {
                    var all = FindZoomBoxes(slide);
                    foreach (PowerPoint.Shape selected in selection.ShapeRange)
                        foreach (PowerPoint.Shape candidate in all)
                            if ((selected.Name ?? "").StartsWith(InsetNamePrefix + GetArtifactKey(candidate) + "_", StringComparison.Ordinal)
                                && selectedIds.Add(candidate.Id)) selectedBoxes.Add(candidate);
                }
                catch { }
            }

            if (selectedBoxes.Count == 1)
            {
                box = selectedBoxes[0];
                return true;
            }
            if (selectedBoxes.Count > 1)
            {
                error = "当前选择中包含多个选区框。请只选中要更新的一个选区框或放大图组。";
                return false;
            }

            List<PowerPoint.Shape> allBoxes = FindZoomBoxes(slide);
            if (allBoxes.Count == 1)
            {
                box = allBoxes[0];
                return true;
            }
            if (allBoxes.Count == 0)
            {
                error = "未找到选区框。请先选中图片并点击「插入选区框」。";
                return false;
            }

            error = "当前页有多个选区框。请先选中要生成或更新的选区框。";
            return false;
        }

        public static bool IsZoomBox(PowerPoint.Shape shape)
        {
            if (shape == null) return false;
            string name = shape.Name ?? string.Empty;
            return name.StartsWith(BoxNamePrefix, StringComparison.Ordinal)
                || name.StartsWith(LegacyBoxNamePrefix, StringComparison.Ordinal);
        }

        public static bool IsZoomArtifact(PowerPoint.Shape shape)
        {
            string name = shape?.Name ?? string.Empty;
            return name.StartsWith(BoxNamePrefix, StringComparison.Ordinal)
                || name.StartsWith(InsetNamePrefix, StringComparison.Ordinal)
                || name.StartsWith(LegacyBoxNamePrefix, StringComparison.Ordinal)
                || name.StartsWith(LegacyInsetNamePrefix, StringComparison.Ordinal);
        }

        public static PowerPoint.Shape FindSourcePicture(PowerPoint.Slide slide,
            PowerPoint.Shape box)
        {
            if (!TryGetSourcePictureId(box, out int pictureId)) return null;

            foreach (PowerPoint.Shape shape in CollectAllShapes(slide))
            {
                try
                {
                    if (shape.Id == pictureId) return shape;
                }
                catch { }
            }
            return null;
        }

        public static List<PowerPoint.Shape> CollectAllShapes(PowerPoint.Slide slide)
        {
            var result = new List<PowerPoint.Shape>();
            if (slide == null) return result;

            var pending = new Queue<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes) pending.Enqueue(shape);

            while (pending.Count > 0)
            {
                PowerPoint.Shape shape = pending.Dequeue();
                result.Add(shape);
                try
                {
                    if (shape.Type == MsoShapeType.msoGroup)
                    {
                        foreach (PowerPoint.Shape child in shape.GroupItems)
                        {
                            pending.Enqueue(child);
                        }
                    }
                }
                catch { }
            }
            return result;
        }

        public static void ComputeTargetSize(PowerPoint.Shape picture,
            PowerPoint.Shape box, ZoomTargetMode mode, float magnification,
            float customWidthCm, out float targetWidth, out float targetHeight)
        {
            float boxWidth = Math.Max(0.1f, box?.Width ?? 0.1f);
            float boxHeight = Math.Max(0.1f, box?.Height ?? 0.1f);

            switch (mode)
            {
                case ZoomTargetMode.Multiple:
                    float factor = magnification > 0 ? magnification : 1f;
                    targetWidth = Math.Max(0.1f, picture?.Width ?? boxWidth) * factor;
                    targetHeight = boxHeight * targetWidth / boxWidth;
                    break;
                case ZoomTargetMode.CustomWidthCm:
                    targetWidth = CmToPoints(Math.Max(0.1f, customWidthCm));
                    targetHeight = boxHeight * targetWidth / boxWidth;
                    break;
                case ZoomTargetMode.SameAsOriginal:
                default:
                    targetWidth = Math.Max(0.1f, picture?.Width ?? boxWidth * 2f);
                    targetHeight = boxHeight * targetWidth / boxWidth;
                    break;
            }
        }

        public static ZoomInsetResult GenerateZoomInset(
            PowerPoint.Slide slide,
            PowerPoint.Shape box,
            PowerPoint.Shape picture,
            float targetWidth,
            float targetHeight,
            float gapPoints,
            ZoomLineStyle lineStyle,
            float lineWeight,
            float boxLineWeight,
            int lineColorRgb,
            int boxColorRgb,
            Office.MsoLineDashStyle lineDash,
            bool group,
            PowerPoint.Application app, Office.MsoLineDashStyle boxDash = Office.MsoLineDashStyle.msoLineSolid,
            ZoomSettings strokeSettings = null)
        {
            var result = new ZoomInsetResult();
            string artifactKey = null;

            try
            {
                if (slide == null || box == null || picture == null)
                {
                    result.Error = "未找到选区框或原图。请重新插入选区框后再试。";
                    return result;
                }
                if (targetWidth <= 0 || targetHeight <= 0)
                {
                    result.Error = "放大图尺寸无效。请检查放大倍数或自定义宽度。";
                    return result;
                }
                if (Math.Abs(picture.Rotation) > 0.01f || Math.Abs(box.Rotation) > 0.01f)
                {
                    result.Error = "原图或选区框带有旋转角度。请先将旋转角度设为 0，再生成放大图。";
                    return result;
                }

                if (!TryBuildCropGeometry(picture, box, out CropGeometry crop, out string cropError))
                {
                    result.Error = cropError;
                    return result;
                }
                if (!TryComputeInsetPosition(slide, picture, targetWidth, targetHeight,
                    Math.Max(0f, gapPoints), out float insetLeft, out float insetTop,
                    out string placementWarning, out string placementError))
                {
                    result.Error = placementError;
                    return result;
                }

                string previousKey = GetArtifactKey(box);
                string previousBoxName = box.Name;
                bool legacy = IsLegacyBoxName(box);
                int sourceId;
                TryGetSourcePictureId(box, out sourceId);
                string finalBoxName = legacy ? BoxNamePrefix + sourceId.ToString(CultureInfo.InvariantCulture)
                    + "_" + Guid.NewGuid().ToString("N").Substring(0, 8) : previousBoxName;
                string finalKey = finalBoxName.Substring(BoxNamePrefix.Length);
                artifactKey = "pending_" + Guid.NewGuid().ToString("N");
                bool cleanedLegacy = false;

                PowerPoint.Shape duplicate = picture.Duplicate()[1];
                duplicate.Name = MakeInsetName(artifactKey, "Picture");
                try
                {
                    duplicate.PictureFormat.CropLeft = crop.Left;
                    duplicate.PictureFormat.CropTop = crop.Top;
                    duplicate.PictureFormat.CropRight = crop.Right;
                    duplicate.PictureFormat.CropBottom = crop.Bottom;
                    duplicate.LockAspectRatio = MsoTriState.msoFalse;
                    duplicate.Width = targetWidth;
                    duplicate.Height = targetHeight;
                    duplicate.Left = insetLeft;
                    duplicate.Top = insetTop;
                }
                catch
                {
                    try { duplicate.Delete(); } catch { }
                    result.Error = "该对象不支持精确裁剪。请将它转换为普通图片后再试。";
                    return result;
                }

                PowerPoint.Shape draftPicture = duplicate;
                string connectionWarning = null;
                try
                {
                    ZoomNativeGeometry.ReplaceWithNativeSites(app, slide, ref box, ref duplicate);
                    try { draftPicture.Delete(); } catch { }
                }
                catch (Exception ex)
                {
                    // The native-site round trip is optional. Keep the exact crop
                    // and duplicate the box so the existing inset remains intact
                    // until the replacement is ready to commit.
                    System.Diagnostics.Trace.TraceWarning("Zoom connection sites: {0}", ex.Message);
                    box = box.Duplicate()[1];
                    connectionWarning = "原生角点连接不可用，已使用普通连接点生成放大图。";
                }
                box.Name = MakeInsetName(artifactKey, "Box");
                box.Line.ForeColor.RGB = boxColorRgb;
                box.Line.Weight = boxLineWeight;
                box.Line.DashStyle = boxDash;
                strokeSettings?.ApplyStroke(box, true);
                AnchoredUnit boxUnit = CreateAnchoredUnit(box);
                AnchoredUnit insetUnit = CreateAnchoredUnit(duplicate);
                result.Created.Add(boxUnit.Group);
                result.Created.Add(insetUnit.Group);

                var lines = new List<PowerPoint.Shape>();
                var endpoints = new List<Tuple<string, string>>();
                if (lineStyle != ZoomLineStyle.None)
                {
                    if (insetLeft >= picture.Left + picture.Width)
                    { endpoints.Add(Tuple.Create("TR", "TL")); endpoints.Add(Tuple.Create("BR", "BL")); }
                    else if (insetLeft + targetWidth <= picture.Left)
                    { endpoints.Add(Tuple.Create("TL", "TR")); endpoints.Add(Tuple.Create("BL", "BR")); }
                    else if (insetTop + targetHeight <= picture.Top)
                    { endpoints.Add(Tuple.Create("TL", "BL")); endpoints.Add(Tuple.Create("TR", "BR")); }
                    else
                    { endpoints.Add(Tuple.Create("BL", "TL")); endpoints.Add(Tuple.Create("BR", "TR")); }
                    if (lineStyle == ZoomLineStyle.CrossedX)
                    {
                        string firstTarget = endpoints[0].Item2;
                        endpoints[0] = Tuple.Create(endpoints[0].Item1, endpoints[1].Item2);
                        endpoints[1] = Tuple.Create(endpoints[1].Item1, firstTarget);
                    }
                }

                int lineIndex = 0;
                foreach (Tuple<string, string> endpoint in endpoints)
                {
                    AnchorTarget from = boxUnit.Anchors[endpoint.Item1];
                    AnchorTarget to = insetUnit.Anchors[endpoint.Item2];
                    PowerPoint.Shape line = CreateLine(slide, from, to, artifactKey,
                        ++lineIndex, lineWeight, lineColorRgb, lineDash, out bool glued);
                    result.Glued &= glued;
                    lines.Add(line);
                    result.Created.Add(line);
                    strokeSettings?.ApplyStroke(line, false);
                }

                var names = new List<string> { boxUnit.Group.Name, insetUnit.Group.Name };
                foreach (PowerPoint.Shape line in lines) names.Add(line.Name);

                if (group)
                {
                    PowerPoint.Shape outer = slide.Shapes.Range(names.ToArray()).Group();
                    outer.Name = MakeInsetName(artifactKey, "Group");
                    result.Group = outer;
                    result.Created.Add(outer);
                    SelectShape(app, outer);
                }
                else
                {
                    SelectShapes(app, slide, names.ToArray());
                }

                // Commit only after the replacement and native connections exist.
                CleanUpPreviousArtifacts(slide, previousKey);
                if (legacy) cleanedLegacy = CleanLegacyArtifacts(slide, sourceId) > 0;
                foreach (PowerPoint.Shape oldShape in CollectAllShapes(slide))
                    if (oldShape.Name == previousBoxName) { oldShape.Delete(); break; }
                string pendingPrefix = InsetNamePrefix + artifactKey + "_";
                foreach (PowerPoint.Shape created in CollectAllShapes(slide))
                    if (created.Name.StartsWith(pendingPrefix, StringComparison.Ordinal))
                        created.Name = InsetNamePrefix + finalKey + "_" + created.Name.Substring(pendingPrefix.Length);
                box.Name = finalBoxName;
                result.Warning = connectionWarning == null ? placementWarning : AppendWarning(placementWarning, connectionWarning);
                if (!result.Glued)
                {
                    result.Warning = AppendWarning(result.Warning,
                        "部分连接点无法粘附；整体移动仍正常，单独移动组件后请重新生成。");
                }
                if (cleanedLegacy)
                {
                    result.Warning = AppendWarning(result.Warning,
                        "已清理旧版本遗留的放大图与连线，并升级了选区框命名。");
                }
                result.Ok = true;
                return result;
            }
            catch (Exception ex)
            {
                if (!string.IsNullOrEmpty(artifactKey))
                {
                    try { CleanUpPreviousArtifacts(slide, artifactKey); } catch { }
                }
                result.Ok = false;
                result.Error = "生成放大图失败：" + ex.Message;
                return result;
            }
        }

        private static bool TryBuildCropGeometry(PowerPoint.Shape picture,
            PowerPoint.Shape box, out CropGeometry crop, out string error)
        {
            crop = null;
            error = null;
            float picLeft = picture.Left;
            float picTop = picture.Top;
            float picRight = picLeft + picture.Width;
            float picBottom = picTop + picture.Height;
            float boxLeft = box.Left;
            float boxTop = box.Top;
            float boxRight = boxLeft + box.Width;
            float boxBottom = boxTop + box.Height;

            if (boxLeft < picLeft - GeometryTolerance || boxTop < picTop - GeometryTolerance ||
                boxRight > picRight + GeometryTolerance || boxBottom > picBottom + GeometryTolerance)
            {
                error = "选区框有一部分超出图片。请把选区框完整移入图片内部后再生成。";
                return false;
            }
            if (picture.Width < GeometryTolerance || picture.Height < GeometryTolerance ||
                box.Width < GeometryTolerance || box.Height < GeometryTolerance)
            {
                error = "图片或选区框尺寸过小，无法生成有效放大图。";
                return false;
            }
            if (!TryGetOriginalPictureSize(picture, out float originalWidth, out float originalHeight))
            {
                error = "无法读取图片的原始尺寸。请将对象转换为普通图片后重试。";
                return false;
            }

            float existingLeft = picture.PictureFormat.CropLeft;
            float existingRight = picture.PictureFormat.CropRight;
            float existingTop = picture.PictureFormat.CropTop;
            float existingBottom = picture.PictureFormat.CropBottom;
            float visibleOriginalWidth = originalWidth - existingLeft - existingRight;
            float visibleOriginalHeight = originalHeight - existingTop - existingBottom;
            if (visibleOriginalWidth <= GeometryTolerance || visibleOriginalHeight <= GeometryTolerance)
            {
                error = "图片现有裁剪参数无效，无法继续裁剪。请先重置图片裁剪。";
                return false;
            }

            float leftRatio = Clamp((boxLeft - picLeft) / picture.Width, 0f, 1f);
            float rightRatio = Clamp((picRight - boxRight) / picture.Width, 0f, 1f);
            float topRatio = Clamp((boxTop - picTop) / picture.Height, 0f, 1f);
            float bottomRatio = Clamp((picBottom - boxBottom) / picture.Height, 0f, 1f);
            if (picture.HorizontalFlip == MsoTriState.msoTrue)
            { float swap = leftRatio; leftRatio = rightRatio; rightRatio = swap; }
            if (picture.VerticalFlip == MsoTriState.msoTrue)
            { float swap = topRatio; topRatio = bottomRatio; bottomRatio = swap; }

            crop = new CropGeometry
            {
                Left = existingLeft + visibleOriginalWidth * leftRatio,
                Right = existingRight + visibleOriginalWidth * rightRatio,
                Top = existingTop + visibleOriginalHeight * topRatio,
                Bottom = existingBottom + visibleOriginalHeight * bottomRatio
            };
            return true;
        }

        private static bool TryGetOriginalPictureSize(PowerPoint.Shape picture,
            out float width, out float height)
        {
            width = 0f;
            height = 0f;
            PowerPoint.ShapeRange duplicateRange = null;
            PowerPoint.Shape probe = null;
            try
            {
                duplicateRange = picture.Duplicate();
                probe = duplicateRange[1];
                probe.PictureFormat.CropLeft = 0f;
                probe.PictureFormat.CropRight = 0f;
                probe.PictureFormat.CropTop = 0f;
                probe.PictureFormat.CropBottom = 0f;
                probe.ScaleWidth(1f, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
                probe.ScaleHeight(1f, MsoTriState.msoTrue, MsoScaleFrom.msoScaleFromTopLeft);
                width = probe.Width;
                height = probe.Height;
                return width > GeometryTolerance && height > GeometryTolerance;
            }
            catch
            {
                return false;
            }
            finally
            {
                try { probe?.Delete(); } catch { }
                PowerPointContext.Release(probe);
                PowerPointContext.Release(duplicateRange);
            }
        }

        private static bool TryComputeInsetPosition(PowerPoint.Slide slide,
            PowerPoint.Shape picture, float width, float height, float gap,
            out float left, out float top, out string warning, out string error)
        {
            left = picture.Left;
            top = picture.Top + picture.Height + gap;
            warning = null;
            error = null;
            float slideWidth;
            float slideHeight;

            try
            {
                var presentation = (PowerPoint.Presentation)slide.Parent;
                slideWidth = presentation.PageSetup.SlideWidth;
                slideHeight = presentation.PageSetup.SlideHeight;
            }
            catch
            {
                return true;
            }

            if (width > slideWidth - SlideMarginPoints * 2f ||
                height > slideHeight - SlideMarginPoints * 2f)
            {
                error = "放大图尺寸超过当前幻灯片。请减小放大倍数或自定义宽度。";
                return false;
            }

            var candidates = new[]
            {
                new PlacementCandidate(picture.Left, picture.Top + picture.Height + gap, null),
                new PlacementCandidate(picture.Left + picture.Width + gap, picture.Top,
                    "原图下方空间不足，放大图已自动放到右侧。"),
                new PlacementCandidate(picture.Left - width - gap, picture.Top,
                    "原图下方空间不足，放大图已自动放到左侧。"),
                new PlacementCandidate(picture.Left, picture.Top - height - gap,
                    "原图下方空间不足，放大图已自动放到上方。")
            };

            foreach (PlacementCandidate candidate in candidates)
            {
                if (FitsSlide(candidate.Left, candidate.Top, width, height, slideWidth, slideHeight))
                {
                    left = candidate.Left;
                    top = candidate.Top;
                    warning = candidate.Warning;
                    return true;
                }
            }

            left = Clamp(picture.Left, SlideMarginPoints,
                Math.Max(SlideMarginPoints, slideWidth - width - SlideMarginPoints));
            top = Clamp(picture.Top + picture.Height + gap, SlideMarginPoints,
                Math.Max(SlideMarginPoints, slideHeight - height - SlideMarginPoints));
            warning = "页面可用空间不足，放大图已自动限制在幻灯片边界内，可能与现有内容重叠。";
            return true;
        }

        private static bool FitsSlide(float left, float top, float width, float height,
            float slideWidth, float slideHeight)
        {
            return left >= SlideMarginPoints && top >= SlideMarginPoints &&
                left + width <= slideWidth - SlideMarginPoints &&
                top + height <= slideHeight - SlideMarginPoints;
        }

        private static AnchoredUnit CreateAnchoredUnit(PowerPoint.Shape content)
        {
            var targets = new Dictionary<string, AnchorTarget>(StringComparer.Ordinal);
            foreach (string tag in new[] { "TL", "TR", "BL", "BR" })
            {
                GetCorner(content, tag, out float x, out float y);
                targets[tag] = new AnchorTarget(content, x, y);
            }
            return new AnchoredUnit(content, targets);
        }

        private static PowerPoint.Shape CreateLine(PowerPoint.Slide slide,
            AnchorTarget from, AnchorTarget to, string artifactKey, int index,
            float weight, int colorRgb, Office.MsoLineDashStyle dash,
            out bool glued)
        {
            PowerPoint.Shape line = slide.Shapes.AddConnector(MsoConnectorType.msoConnectorStraight,
                from.X, from.Y, to.X, to.Y);
            line.Name = MakeInsetName(artifactKey, "Line" + index.ToString(CultureInfo.InvariantCulture));
            line.Line.ForeColor.RGB = colorRgb;
            line.Line.Weight = weight;
            try { line.Line.DashStyle = dash; } catch { }

            bool beginGlued = ConnectCorner(line, from.Shape, from.X, from.Y,
                to.X, to.Y, true);
            DeriveGluedPoint(line, to.X, to.Y, out float actualX, out float actualY);
            bool endGlued = ConnectCorner(line, to.Shape, to.X, to.Y,
                actualX, actualY, false);
            glued = beginGlued && endGlued;
            return line;
        }

        private static bool ConnectCorner(PowerPoint.Shape connector, PowerPoint.Shape shape,
            float targetX, float targetY, float knownX, float knownY, bool isBegin)
        {
            int siteCount;
            try { siteCount = shape.ConnectionSiteCount; }
            catch { return false; }
            if (siteCount <= 0) return false;

            int bestSite = 0;
            double bestDistance = double.MaxValue;
            for (int site = 1; site <= siteCount; site++)
            {
                try
                {
                    Disconnect(connector, isBegin);
                    if (isBegin) connector.ConnectorFormat.BeginConnect(shape, site);
                    else connector.ConnectorFormat.EndConnect(shape, site);

                    DeriveGluedPoint(connector, knownX, knownY, out float siteX, out float siteY);
                    double dx = siteX - targetX;
                    double dy = siteY - targetY;
                    double distance = dx * dx + dy * dy;
                    if (distance < bestDistance)
                    {
                        bestDistance = distance;
                        bestSite = site;
                    }
                }
                catch { }
            }

            if (bestSite <= 0) return false;
            try
            {
                Disconnect(connector, isBegin);
                if (isBegin) connector.ConnectorFormat.BeginConnect(shape, bestSite);
                else connector.ConnectorFormat.EndConnect(shape, bestSite);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void Disconnect(PowerPoint.Shape connector, bool begin)
        {
            try
            {
                if (begin) connector.ConnectorFormat.BeginDisconnect();
                else connector.ConnectorFormat.EndDisconnect();
            }
            catch { }
        }

        private static void DeriveGluedPoint(PowerPoint.Shape connector,
            float knownX, float knownY, out float x, out float y)
        {
            float left = connector.Left;
            float top = connector.Top;
            float right = left + connector.Width;
            float bottom = top + connector.Height;
            x = Math.Abs(knownX - left) < Math.Abs(knownX - right) ? right : left;
            y = Math.Abs(knownY - top) < Math.Abs(knownY - bottom) ? bottom : top;
        }

        private static void GetCorner(PowerPoint.Shape shape, string tag,
            out float x, out float y)
        {
            x = tag.EndsWith("R", StringComparison.Ordinal)
                ? shape.Left + shape.Width : shape.Left;
            y = tag.StartsWith("B", StringComparison.Ordinal)
                ? shape.Top + shape.Height : shape.Top;
        }

        private static void SelectShape(PowerPoint.Application app, PowerPoint.Shape shape)
        {
            try
            {
                app.ActiveWindow.Selection.Unselect();
                shape.Select();
            }
            catch { }
        }

        private static void SelectShapes(PowerPoint.Application app,
            PowerPoint.Slide slide, object[] names)
        {
            try
            {
                app.ActiveWindow.Selection.Unselect();
                slide.Shapes.Range(names).Select();
            }
            catch { }
        }

        private static void CleanUpPreviousArtifacts(PowerPoint.Slide slide, string artifactKey)
        {
            if (slide == null || string.IsNullOrEmpty(artifactKey)) return;

            bool changed = true;
            while (changed)
            {
                changed = false;
                var groups = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        if (shape.Type == MsoShapeType.msoGroup &&
                            GroupContainsArtifact(shape, artifactKey))
                        {
                            groups.Add(shape);
                        }
                    }
                    catch { }
                }
                foreach (PowerPoint.Shape group in groups)
                {
                    try
                    {
                        group.Ungroup();
                        changed = true;
                    }
                    catch { }
                }
            }

            string prefix = InsetNamePrefix + artifactKey + "_";
            var toDelete = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if ((shape.Name ?? string.Empty).StartsWith(prefix, StringComparison.Ordinal))
                {
                    toDelete.Add(shape);
                }
            }
            foreach (PowerPoint.Shape shape in toDelete)
            {
                try { shape.Delete(); } catch { }
            }
        }

        private static bool GroupContainsArtifact(PowerPoint.Shape group,
            string artifactKey)
        {
            string prefix = InsetNamePrefix + artifactKey + "_";
            try
            {
                if ((group.Name ?? string.Empty).StartsWith(prefix, StringComparison.Ordinal))
                {
                    return true;
                }
                foreach (PowerPoint.Shape child in group.GroupItems)
                {
                    string name = child.Name ?? string.Empty;
                    if (name.StartsWith(prefix, StringComparison.Ordinal)) return true;
                    if (child.Type == MsoShapeType.msoGroup &&
                        GroupContainsArtifact(child, artifactKey)) return true;
                }
            }
            catch { }
            return false;
        }

        private static void CollectZoomBoxesFromShape(PowerPoint.Shape shape,
            List<PowerPoint.Shape> boxes, HashSet<int> ids)
        {
            if (shape == null) return;
            try
            {
                if (IsZoomBox(shape))
                {
                    if (ids.Add(shape.Id)) boxes.Add(shape);
                    return;
                }
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    foreach (PowerPoint.Shape child in shape.GroupItems)
                    {
                        CollectZoomBoxesFromShape(child, boxes, ids);
                    }
                }
            }
            catch { }
        }

        private static bool TryGetSourcePictureId(PowerPoint.Shape box, out int pictureId)
        {
            pictureId = 0;
            string name = box?.Name ?? string.Empty;
            string prefix = name.StartsWith(BoxNamePrefix, StringComparison.Ordinal)
                ? BoxNamePrefix
                : name.StartsWith(LegacyBoxNamePrefix, StringComparison.Ordinal)
                    ? LegacyBoxNamePrefix
                    : null;
            if (prefix == null) return false;
            string suffix = name.Substring(prefix.Length);
            int separator = suffix.IndexOf('_');
            string idText = separator >= 0 ? suffix.Substring(0, separator) : suffix;
            return int.TryParse(idText, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out pictureId);
        }

        private static string GetArtifactKey(PowerPoint.Shape box)
        {
            string name = box?.Name ?? string.Empty;
            string prefix = name.StartsWith(BoxNamePrefix, StringComparison.Ordinal)
                ? BoxNamePrefix
                : name.StartsWith(LegacyBoxNamePrefix, StringComparison.Ordinal)
                    ? LegacyBoxNamePrefix
                    : null;
            if (prefix != null)
            {
                string suffix = name.Substring(prefix.Length);
                if (suffix.IndexOf('_') >= 0) return suffix;
            }
            return box.Id.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 旧版本选区框名，例如 SlideSCI_ZoomBox_37 或其带 GUID 的变体。
        /// </summary>
        private static bool IsLegacyBoxName(PowerPoint.Shape box)
        {
            string name = box?.Name ?? string.Empty;
            if (name.StartsWith(LegacyBoxNamePrefix, StringComparison.Ordinal)) return true;
            if (!name.StartsWith(BoxNamePrefix, StringComparison.Ordinal)) return false;
            string suffix = name.Substring(BoxNamePrefix.Length);
            // Focus 格式 = <图片Id>_<8位hex>；其他当前前缀格式视为旧格式。
            int separator = suffix.IndexOf('_');
            if (separator < 0) return true;
            string idPart = suffix.Substring(0, separator);
            string remain = suffix.Substring(separator + 1);
            return !(remain.Length == 8 && IsHexString(remain));
        }

        private static bool IsHexString(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            foreach (char c in text)
            {
                if (!Uri.IsHexDigit(c)) return false;
            }
            return true;
        }

        /// <summary>
        /// 删除指定图片的旧格式生成物（SlideSCI_ZoomInset_&lt;图片Id&gt;_*），
        /// 逐层拆开包含它们的编组；不影响 Focus 版多实例的其他选区。返回清理数量。
        /// </summary>
        private static int CleanLegacyArtifacts(PowerPoint.Slide slide, int pictureId)
        {
            if (slide == null) return 0;
            string legacyPrefix = LegacyInsetNamePrefix + pictureId.ToString(CultureInfo.InvariantCulture) + "_";

            bool IsLegacyName(string name)
            {
                if (!name.StartsWith(legacyPrefix, StringComparison.Ordinal)) return false;
                string rest = name.Substring(legacyPrefix.Length);
                int separator = rest.IndexOf('_');
                string first = separator >= 0 ? rest.Substring(0, separator) : rest;
                // 新版 = <8位hex>_<suffix>；旧版 = Pic/LineN/Unit/Anchor_*
                return !(first.Length == 8 && IsHexString(first));
            }

            bool changed = true;
            while (changed)
            {
                changed = false;
                var groups = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        if (shape.Type == MsoShapeType.msoGroup && GroupContainsLegacy(shape, IsLegacyName))
                        {
                            groups.Add(shape);
                        }
                    }
                    catch { }
                }
                foreach (PowerPoint.Shape group in groups)
                {
                    try { group.Ungroup(); changed = true; } catch { }
                }
            }

            var toDelete = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (IsLegacyName(shape.Name ?? string.Empty)) toDelete.Add(shape);
            }
            foreach (PowerPoint.Shape shape in toDelete)
            {
                try { shape.Delete(); } catch { }
            }
            return toDelete.Count;
        }

        private static bool GroupContainsLegacy(PowerPoint.Shape group,
            Func<string, bool> isLegacyName)
        {
            try
            {
                if (isLegacyName(group.Name ?? string.Empty)) return true;
                foreach (PowerPoint.Shape child in group.GroupItems)
                {
                    if (isLegacyName(child.Name ?? string.Empty)) return true;
                    if (child.Type == MsoShapeType.msoGroup &&
                        GroupContainsLegacy(child, isLegacyName)) return true;
                }
            }
            catch { }
            return false;
        }

        private static string MakeInsetName(string artifactKey, string suffix)
        {
            return InsetNamePrefix + artifactKey + "_" + suffix;
        }

        private static string AppendWarning(string current, string next)
        {
            return string.IsNullOrWhiteSpace(current) ? next : current + Environment.NewLine + next;
        }

        private static float Clamp(float value, float min, float max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        private sealed class CropGeometry
        {
            internal float Left;
            internal float Top;
            internal float Right;
            internal float Bottom;
        }

        private sealed class AnchorTarget
        {
            internal PowerPoint.Shape Shape { get; }
            internal float X { get; }
            internal float Y { get; }

            internal AnchorTarget(PowerPoint.Shape shape, float x, float y)
            {
                Shape = shape;
                X = x;
                Y = y;
            }
        }

        private sealed class AnchoredUnit
        {
            internal PowerPoint.Shape Group { get; }
            internal Dictionary<string, AnchorTarget> Anchors { get; }

            internal AnchoredUnit(PowerPoint.Shape group,
                Dictionary<string, AnchorTarget> anchors)
            {
                Group = group;
                Anchors = anchors;
            }
        }

        private sealed class PlacementCandidate
        {
            internal float Left { get; }
            internal float Top { get; }
            internal string Warning { get; }

            internal PlacementCandidate(float left, float top, string warning)
            {
                Left = left;
                Top = top;
                Warning = warning;
            }
        }
    }
}
