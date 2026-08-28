using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlideSCI
{
    /// <summary>放大图的尺寸确定方式。</summary>
    public enum ZoomTargetMode
    {
        /// <summary>放大图与原图同尺寸（期刊常见布局；若选区比例不一致会产生拉伸）。</summary>
        SameAsOriginal = 0,
        /// <summary>按选区边长倍数放大，保持选区宽高比。</summary>
        Multiple = 1,
        /// <summary>自定义放大图宽度（cm），高度按选区比例缩放。</summary>
        CustomWidthCm = 2
    }

    /// <summary>框角与放大图角之间的连线样式。</summary>
    public enum ZoomLineStyle
    {
        /// <summary>期刊漏斗：框下两角 → 放大图上两角（平行、非同角直连）。</summary>
        JournalFunnel = 0,
        /// <summary>四角交叉 X 形连线。</summary>
        CrossedX = 1
    }

    /// <summary>生成结果。</summary>
    public class ZoomInsetResult
    {
        public bool Ok { get; set; }
        public string Error { get; set; }
        public List<PowerPoint.Shape> Created { get; } = new List<PowerPoint.Shape>();
        public PowerPoint.Shape Group { get; set; }
    }

    /// <summary>
    /// 「局部放大图」功能的纯 COM 实现：
    /// 在原图上放置选区框（InsertZoomBox），再把框内区域裁剪放大到目标尺寸、
    /// 用细线连接框角与放大图角（GenerateZoomInset）。
    /// 生成物统一以 SlideSCI_Zoom* 前缀命名，便于重新生成时定位与清理。
    /// </summary>
    public static class ZoomInsetHelper
    {
        public const string BoxNamePrefix = "SlideSCI_ZoomBox_";
        public const string InsetNamePrefix = "SlideSCI_ZoomInset_";

        private const float PointsPerCm = 28.3464593f;

        public static float CmToPoints(float cm) => cm * PointsPerCm;

        /// <summary>
        /// 在原图中心放置一个无填充的方形选区框。
        /// </summary>
        /// <param name="boxSidePercent">框边长占图片宽度的百分比（5–90）。</param>
        public static PowerPoint.Shape InsertZoomBox(PowerPoint.Slide slide, PowerPoint.Shape picture,
            float boxSidePercent, float boxLineWeight = 1.5f)
        {
            float side = Math.Max(1f, picture.Width * (boxSidePercent / 100f));
            float left = picture.Left + (picture.Width - side) / 2f;
            float top = picture.Top + (picture.Height - side) / 2f;

            PowerPoint.Shape box = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle, left, top, side, side);
            box.Name = BoxNamePrefix + picture.Id;
            box.Fill.Visible = MsoTriState.msoFalse;
            box.Line.ForeColor.RGB = 0x000000;
            box.Line.Weight = boxLineWeight;
            return box;
        }

        /// <summary>
        /// 在当前幻灯片上查找选区框（可能位于编组内部）。
        /// </summary>
        public static PowerPoint.Shape FindZoomBox(PowerPoint.Slide slide)
        {
            List<PowerPoint.Shape> all = CollectAllShapes(slide);
            foreach (PowerPoint.Shape shape in all)
            {
                if (IsZoomBox(shape)) return shape;
            }
            return null;
        }

        public static bool IsZoomBox(PowerPoint.Shape shape)
        {
            return shape != null
                && (shape.Name ?? string.Empty).StartsWith(BoxNamePrefix, StringComparison.Ordinal);
        }

        public static bool IsZoomArtifact(PowerPoint.Shape shape)
        {
            string name = shape?.Name ?? string.Empty;
            return name.StartsWith(BoxNamePrefix, StringComparison.Ordinal)
                || name.StartsWith(InsetNamePrefix, StringComparison.Ordinal);
        }

        /// <summary>按选区框名称中记录的原图 Id 找回原图片（可能位于编组内部）。</summary>
        public static PowerPoint.Shape FindSourcePicture(PowerPoint.Slide slide, PowerPoint.Shape box)
        {
            string name = box?.Name ?? string.Empty;
            int prefixLength = BoxNamePrefix.Length;
            if (name.Length <= prefixLength || !int.TryParse(name.Substring(prefixLength), out int pictureId))
            {
                return null;
            }

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

        /// <summary>收集当前幻灯片所有形状（含编组内部成员），广度优先。</summary>
        public static List<PowerPoint.Shape> CollectAllShapes(PowerPoint.Slide slide)
        {
            var result = new List<PowerPoint.Shape>();
            var pending = new Queue<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                pending.Enqueue(shape);
            }

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

        /// <summary>
        /// 计算放大图的目标尺寸。
        /// </summary>
        public static void ComputeTargetSize(PowerPoint.Shape picture, PowerPoint.Shape box,
            ZoomTargetMode mode, float magnification, float customWidthCm,
            out float targetWidth, out float targetHeight)
        {
            switch (mode)
            {
                case ZoomTargetMode.SameAsOriginal:
                    targetWidth = picture.Width;
                    targetHeight = picture.Height;
                    break;
                case ZoomTargetMode.Multiple:
                    float factor = magnification > 0 ? magnification : 2f;
                    targetWidth = box.Width * factor;
                    targetHeight = box.Height * factor;
                    break;
                case ZoomTargetMode.CustomWidthCm:
                default:
                    targetWidth = CmToPoints(Math.Max(0.1f, customWidthCm));
                    targetHeight = box.Width > 0
                        ? box.Height * (targetWidth / box.Width)
                        : targetWidth;
                    break;
            }
        }

        /// <summary>
        /// 生成局部放大图：
        /// 1) 清理旧生成物（仅保留选区框本身）；2) 复制原图并按框裁剪放大；
        /// 3) 画连接线；4) 可选整体编组并选中。
        /// </summary>
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
            bool group,
            PowerPoint.Application app)
        {
            var result = new ZoomInsetResult();
            try
            {
                // 0) 前置校验
                if (box == null || picture == null)
                {
                    result.Error = "未找到选区框或原图。请重新点击「插入选区框」后再试。";
                    return result;
                }

                if (Math.Abs(picture.Rotation) > 0.01f)
                {
                    result.Error = "图片带有旋转角度，裁剪放大会错位。请先取消图片旋转（旋转角度设为 0），再重新生成。";
                    return result;
                }

                // 1) 清理上一次生成物：先拆开包含生成物的编组，再删除放大图与连线，
                //    保留选区框（可能已被用户移动到目标区域）。
                CleanUpPreviousArtifacts(slide);

                // 2) 把选区框几何夹取到原图范围内
                float picLeft = picture.Left, picTop = picture.Top;
                float picRight = picLeft + picture.Width, picBottom = picTop + picture.Height;
                float boxLeft = Math.Max(box.Left, picLeft);
                float boxTop = Math.Max(box.Top, picTop);
                float boxRight = Math.Min(box.Left + box.Width, picRight);
                float boxBottom = Math.Min(box.Top + box.Height, picBottom);
                float boxW = boxRight - boxLeft;
                float boxH = boxBottom - boxTop;
                if (boxW < 0.5f || boxH < 0.5f)
                {
                    result.Error = "选区框与图片没有有效重叠，无法生成放大图。请把选区框拖回图片内部。";
                    return result;
                }

                // 3) 裁剪比例：框相对原图的百分比（PowerPoint Crop* 取值 0~1）
                float cropLeft = (boxLeft - picLeft) / picture.Width;
                float cropTop = (boxTop - picTop) / picture.Height;
                float cropRight = (picRight - boxRight) / picture.Width;
                float cropBottom = (picBottom - boxBottom) / picture.Height;

                // 4) 复制原图 → 先缩放到框大小 → 裁剪 → 再放大到目标尺寸
                PowerPoint.Shape dup = picture.Duplicate()[1];
                dup.Name = MakeInsetName(box.Id, "Pic");
                result.Created.Add(dup);
                try
                {
                    dup.LockAspectRatio = MsoTriState.msoFalse;
                    dup.Width = boxW;
                    dup.Height = boxH;
                    dup.PictureFormat.CropLeft = cropLeft;
                    dup.PictureFormat.CropTop = cropTop;
                    dup.PictureFormat.CropRight = cropRight;
                    dup.PictureFormat.CropBottom = cropBottom;
                }
                catch
                {
                    try { dup.Delete(); } catch { }
                    result.Created.Clear();
                    result.Error = "该图片不支持裁剪复制（可能是占位符粘贴的图片或链接图片）。请先在图片上右键 →「剪切」再原位「粘贴」为普通图片后重试。";
                    return result;
                }

                dup.Width = targetWidth;
                dup.Height = targetHeight;
                dup.Left = picLeft;
                dup.Top = picBottom + gapPoints;

                // 5) 画连接线
                float boxBLx = box.Left, boxBLy = box.Top + box.Height;
                float boxBRx = box.Left + box.Width, boxBRy = box.Top + box.Height;
                float boxTLx = box.Left, boxTLy = box.Top;
                float boxTRx = box.Left + box.Width, boxTRy = box.Top;
                float insTLx = dup.Left, insTLy = dup.Top;
                float insTRx = dup.Left + dup.Width, insTRy = dup.Top;
                float insBLx = dup.Left, insBLy = dup.Top + dup.Height;
                float insBRx = dup.Left + dup.Width, insBRy = dup.Top + dup.Height;

                var lineEnds = new List<(float x1, float y1, float x2, float y2)>();
                if (lineStyle == ZoomLineStyle.JournalFunnel)
                {
                    // 期刊漏斗：框下两角 → 放大图上两角（平行、非同角直连）
                    lineEnds.Add((boxBLx, boxBLy, insTLx, insTLy));
                    lineEnds.Add((boxBRx, boxBRy, insTRx, insTRy));
                }
                else
                {
                    // 交叉 X 形：四角交叉连接
                    lineEnds.Add((boxTLx, boxTLy, insTRx, insTRy));
                    lineEnds.Add((boxTRx, boxTRy, insTLx, insTLy));
                    lineEnds.Add((boxBLx, boxBLy, insBRx, insBRy));
                    lineEnds.Add((boxBRx, boxBRy, insBLx, insBLy));
                }

                int lineIndex = 0;
                foreach (var (x1, y1, x2, y2) in lineEnds)
                {
                    PowerPoint.Shape line = slide.Shapes.AddConnector(
                        MsoConnectorType.msoConnectorStraight, x1, y1, x2, y2);
                    line.Name = MakeInsetName(box.Id, "Line" + (++lineIndex));
                    line.Line.ForeColor.RGB = 0x000000;
                    line.Line.Weight = lineWeight;
                    result.Created.Add(line);
                }

                // 6) 更新框线样式（对话框可能改了线宽）
                try
                {
                    box.Line.ForeColor.RGB = 0x000000;
                    box.Line.Weight = boxLineWeight;
                }
                catch { }

                // 7) 可选编组
                if (group)
                {
                    var names = new List<string>();
                    foreach (PowerPoint.Shape shape in result.Created)
                    {
                        names.Add(shape.Name);
                    }
                    names.Add(box.Name);

                    PowerPoint.Shape groupShape = slide.Shapes.Range(names.ToArray()).Group();
                    result.Group = groupShape;
                    try
                    {
                        app.ActiveWindow.Selection.Unselect();
                        slide.Shapes.Range(new object[] { groupShape.Name }).Select();
                    }
                    catch { }
                }
                else
                {
                    try
                    {
                        app.ActiveWindow.Selection.Unselect();
                        slide.Shapes.Range(namesOf(result.Created)).Select();
                    }
                    catch { }
                }

                result.Ok = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Ok = false;
                result.Error = "生成放大图失败: " + ex.Message;
                return result;
            }
        }

        private static object[] namesOf(List<PowerPoint.Shape> shapes)
        {
            var names = new object[shapes.Count];
            for (int i = 0; i < shapes.Count; i++) names[i] = shapes[i].Name;
            return names;
        }

        private static string MakeInsetName(int boxShapeId, string suffix)
        {
            return InsetNamePrefix + boxShapeId + "_" + suffix;
        }

        /// <summary>
        /// 清理旧的放大图与连线：把包含生成物（选区框/放大图/连线）的编组逐层拆开，
        /// 再删除所有放大图与连线成员，选区框保留。
        /// </summary>
        private static void CleanUpPreviousArtifacts(PowerPoint.Slide slide)
        {
            // 逐层拆开包含生成物的顶层编组，直到顶层不再有生成物编组
            bool ungrouped = true;
            while (ungrouped)
            {
                ungrouped = false;
                var topGroups = new List<PowerPoint.Shape>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        if (shape.Type == MsoShapeType.msoGroup && GroupContainsArtifact(shape))
                        {
                            topGroups.Add(shape);
                        }
                    }
                    catch { }
                }

                foreach (PowerPoint.Shape groupShape in topGroups)
                {
                    try
                    {
                        groupShape.Ungroup();
                        ungrouped = true;
                    }
                    catch { }
                }
            }

            // 删除放大图与连线（保留选区框）
            var toDelete = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                string name = shape.Name ?? string.Empty;
                if (name.StartsWith(InsetNamePrefix, StringComparison.Ordinal))
                {
                    toDelete.Add(shape);
                }
            }
            foreach (PowerPoint.Shape shape in toDelete)
            {
                try { shape.Delete(); } catch { }
            }
        }

        private static bool GroupContainsArtifact(PowerPoint.Shape groupShape)
        {
            try
            {
                foreach (PowerPoint.Shape child in groupShape.GroupItems)
                {
                    if (IsZoomArtifact(child)) return true;
                    if (child.Type == MsoShapeType.msoGroup && GroupContainsArtifact(child)) return true;
                }
            }
            catch { }
            return false;
        }
    }
}