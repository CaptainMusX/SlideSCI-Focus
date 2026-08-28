using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

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
        public bool Glued { get; set; } = true;
    }

    /// <summary>
    /// 「局部放大图」功能的纯 COM 实现。
    ///
    /// 像素保真：所有操作只在形状层进行（复制、裁剪、缩放、连线、编组），
    /// 不导出文件、不重新编码图像，原图像素始终原封不动。
    ///
    /// 连线吸附：连线的端点通过 PowerPoint 原生连接点（BeginConnect/EndConnect）
    /// 粘附到「选区框的四角」和「放大图四角的隐形锚点」（锚点与放大图编组在一起）。
    /// 之后移动选区框或放大图，连线会自动跟随拉伸，无需重新计算坐标。
    /// 连接点索引因 PowerPoint 版本而异，实现通过试连 + 测距自动校准到最近的真实角点。
    /// </summary>
    public static class ZoomInsetHelper
    {
        public const string BoxNamePrefix = "SlideSCI_ZoomBox_";
        public const string InsetNamePrefix = "SlideSCI_ZoomInset_";

        private const float PointsPerCm = 28.3464593f;
        private const float SocketSizePoints = 6f;

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
            foreach (PowerPoint.Shape shape in CollectAllShapes(slide))
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
        /// 计算放大图的目标尺寸（纯几何换算，不触碰像素）。
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
        /// 1) 清理旧生成物（仅保留选区框）；2) 复制原图并按框裁剪放大（像素无损）；
        /// 3) 在放大图四角放置隐形锚点并与放大图编组；4) 画连接线并粘附到角点；
        /// 5) 可选整体编组。
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
            int lineColorRgb,
            int boxColorRgb,
            Office.MsoLineDashStyle lineDash,
            bool group,
            PowerPoint.Application app)
        {
            var result = new ZoomInsetResult();
            try
            {
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

                // 1) 清理：拆开包含生成物的编组，删除旧的放大图/锚点/连线，保留选区框
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

                // 3) 裁剪比例（PowerPoint Crop* 取值 0~1）
                float cropLeft = (boxLeft - picLeft) / picture.Width;
                float cropTop = (boxTop - picTop) / picture.Height;
                float cropRight = (picRight - boxRight) / picture.Width;
                float cropBottom = (picBottom - boxBottom) / picture.Height;

                // 4) 复制原图 → 缩放到框大小 → 裁剪 → 放大到目标尺寸
                //    （全部是矢量操作：像素数据原样随形状缩放，绝不重新编码）
                PowerPoint.Shape dup = picture.Duplicate()[1];
                dup.Name = MakeInsetName(box.Id, "Pic");
                float dupLeft, dupTop;
                try
                {
                    dup.LockAspectRatio = MsoTriState.msoFalse;
                    dup.Width = boxW;
                    dup.Height = boxH;
                    dup.PictureFormat.CropLeft = cropLeft;
                    dup.PictureFormat.CropTop = cropTop;
                    dup.PictureFormat.CropRight = cropRight;
                    dup.PictureFormat.CropBottom = cropBottom;

                    dup.Width = targetWidth;
                    dup.Height = targetHeight;
                    dupLeft = picLeft;
                    dupTop = picBottom + gapPoints;
                    dup.Left = dupLeft;
                    dup.Top = dupTop;
                }
                catch
                {
                    try { dup.Delete(); } catch { }
                    result.Error = "该图片不支持裁剪复制（可能是占位符粘贴的图片或链接图片）。请先在图片上右键 →「剪切」再原位「粘贴」为普通图片后重试。";
                    return result;
                }
                result.Created.Add(dup);

                // 5) 放大图四角的隐形锚点（与放大图编组，作为连线粘附点）
                var socketNames = new List<string>();
                var corners = new[] { "TL", "TR", "BL", "BR" };
                foreach (string tag in corners)
                {
                    (float cx, float cy) = GetCorner(dup, tag);
                    PowerPoint.Shape socket = CreateSocket(slide, cx, cy, box.Id, tag);
                    socketNames.Add(socket.Name);
                    result.Created.Add(socket);
                }

                // 6) 放大图 + 锚点 编组为「放大图单元」
                var unitNames = new List<string> { dup.Name };
                unitNames.AddRange(socketNames);
                PowerPoint.Shape unit = slide.Shapes.Range(unitNames.ToArray()).Group();
                unit.Name = MakeInsetName(box.Id, "Unit");
                result.Created.Add(unit);

                // 7) 连接线：粘附到 选区框角 ↔ 锚点角
                var lineShapes = new List<PowerPoint.Shape>();
                var lineEnds = new List<(string boxCorner, string insetCorner)>();
                if (lineStyle == ZoomLineStyle.JournalFunnel)
                {
                    // 期刊漏斗：框下两角 → 放大图上两角（平行、非同角直连）
                    lineEnds.Add(("BL", "TL"));
                    lineEnds.Add(("BR", "TR"));
                }
                else
                {
                    // 交叉 X 形：四角交叉连接
                    lineEnds.Add(("TL", "TR"));
                    lineEnds.Add(("TR", "TL"));
                    lineEnds.Add(("BL", "BR"));
                    lineEnds.Add(("BR", "BL"));
                }

                int lineIndex = 0;
                foreach (var (boxCorner, insetCorner) in lineEnds)
                {
                    PowerPoint.Shape line = CreateLine(slide, box, dup, boxCorner, insetCorner,
                        box.Id, ++lineIndex, lineWeight, lineColorRgb, lineDash);
                    lineShapes.Add(line);
                    result.Created.Add(line);
                }

                // 8) 更新框线样式
                try
                {
                    box.Line.ForeColor.RGB = boxColorRgb;
                    box.Line.Weight = boxLineWeight;
                }
                catch { }

                // 9) 可选整体编组（框 + 放大图单元 + 连线）
                var allNames = new List<string> { box.Name, unit.Name };
                foreach (PowerPoint.Shape line in lineShapes) allNames.Add(line.Name);

                if (group)
                {
                    PowerPoint.Shape groupShape = slide.Shapes.Range(allNames.ToArray()).Group();
                    result.Group = groupShape;
                    SelectShape(app, groupShape);
                }
                else
                {
                    SelectShapes(app, slide, allNames.ToArray());
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

        private static void SelectShape(PowerPoint.Application app, PowerPoint.Shape shape)
        {
            try
            {
                app.ActiveWindow.Selection.Unselect();
                shape.Select();
            }
            catch { }
        }

        private static void SelectShapes(PowerPoint.Application app, PowerPoint.Slide slide, object[] names)
        {
            try
            {
                app.ActiveWindow.Selection.Unselect();
                slide.Shapes.Range(names).Select();
            }
            catch { }
        }

        private static (float x, float y) GetCorner(PowerPoint.Shape shape, string tag)
        {
            switch (tag)
            {
                case "TR": return (shape.Left + shape.Width, shape.Top);
                case "BL": return (shape.Left, shape.Top + shape.Height);
                case "BR": return (shape.Left + shape.Width, shape.Top + shape.Height);
                default: return (shape.Left, shape.Top); // TL
            }
        }

        /// <summary>隐形矩形锚点：中心位于目标角点，作为连线粘附点。</summary>
        private static PowerPoint.Shape CreateSocket(PowerPoint.Slide slide, float centerX, float centerY,
            int boxId, string tag)
        {
            PowerPoint.Shape socket = slide.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                centerX - SocketSizePoints / 2f,
                centerY - SocketSizePoints / 2f,
                SocketSizePoints,
                SocketSizePoints);
            socket.Name = MakeInsetName(boxId, "Anchor_" + tag);
            socket.Fill.Visible = MsoTriState.msoFalse;
            socket.Line.Visible = MsoTriState.msoFalse;
            return socket;
        }

        private static PowerPoint.Shape CreateLine(PowerPoint.Slide slide,
            PowerPoint.Shape box, PowerPoint.Shape inset,
            string boxCorner, string insetCorner,
            int boxId, int index, float weight, int colorRgb, Office.MsoLineDashStyle dash)
        {
            // 先用角点坐标创建直线，再用连接点粘附
            (float bx, float by) = GetCorner(box, boxCorner);
            (float ix, float iy) = GetCorner(inset, insetCorner);

            PowerPoint.Shape line = slide.Shapes.AddConnector(
                MsoConnectorType.msoConnectorStraight, bx, by, ix, iy);
            line.Name = MakeInsetName(boxId, "Line" + index);
            line.Line.ForeColor.RGB = colorRgb;
            line.Line.Weight = weight;
            try { line.Line.DashStyle = dash; } catch { }

            // 粘附 begin → 选区框角（此时对侧端点仍位于 (ix, iy)，可直接作为已知点）
            ConnectCorner(line, box, bx, by, ix, iy, isBegin: true);

            // 反推 begin 实际吸附点，作为 end 校准的已知点
            (float ax, float ay) = DeriveGluedPoint(line, ix, iy);
            ConnectCorner(line, inset, ix, iy, ax, ay, isBegin: false);
            return line;
        }

        /// <summary>
        /// 直线连接器粘附一端后，其包围盒（Left/Top/Width/Height）即两端的包围盒。
        /// 给定已知的另一端坐标，可反推出被粘附端点的实际坐标。
        /// </summary>
        private static (float x, float y) DeriveGluedPoint(PowerPoint.Shape connector,
            float knownX, float knownY)
        {
            float L = connector.Left, T = connector.Top;
            float R = L + connector.Width, B = T + connector.Height;
            float x = Math.Abs(knownX - L) < Math.Abs(knownX - R) ? R : L;
            float y = Math.Abs(knownY - T) < Math.Abs(knownY - B) ? B : T;
            return (x, y);
        }

        /// <summary>
        /// 把连线端点粘附到形状上「最靠近目标角点」的连接点。
        /// 连接点索引因形状类型/PPT 版本而异，因此逐个试连，通过包围盒反推
        /// 每个连接点的实际坐标，选择距离目标角点最近者作为最终粘附点。
        /// 若形状没有连接点（ConnectionSiteCount &lt;= 0），保持绝对坐标不动。
        /// </summary>
        private static void ConnectCorner(PowerPoint.Shape connector, PowerPoint.Shape shape,
            float targetX, float targetY, float knownX, float knownY, bool isBegin)
        {
            int siteCount = 0;
            try { siteCount = shape.ConnectionSiteCount; } catch { }
            if (siteCount <= 0) return; // 无连接点 → 保持绝对坐标

            int bestSite = 1;
            double bestDist = double.MaxValue;

            for (int site = 1; site <= siteCount; site++)
            {
                try
                {
                    // 换连接点前先断开旧连接（从未连接过时忽略异常）
                    if (isBegin)
                    {
                        try { connector.ConnectorFormat.BeginDisconnect(); } catch { }
                        connector.ConnectorFormat.BeginConnect(shape, site);
                    }
                    else
                    {
                        try { connector.ConnectorFormat.EndDisconnect(); } catch { }
                        connector.ConnectorFormat.EndConnect(shape, site);
                    }

                    float L = connector.Left, T = connector.Top;
                    float R = L + connector.Width, B = T + connector.Height;
                    float siteX = Math.Abs(knownX - L) < Math.Abs(knownX - R) ? R : L;
                    float siteY = Math.Abs(knownY - T) < Math.Abs(knownY - B) ? B : T;

                    double dx = siteX - targetX;
                    double dy = siteY - targetY;
                    double dist = dx * dx + dy * dy;
                    if (dist < bestDist)
                    {
                        bestDist = dist;
                        bestSite = site;
                    }
                }
                catch { }
            }

            try
            {
                if (isBegin)
                {
                    try { connector.ConnectorFormat.BeginDisconnect(); } catch { }
                    connector.ConnectorFormat.BeginConnect(shape, bestSite);
                }
                else
                {
                    try { connector.ConnectorFormat.EndDisconnect(); } catch { }
                    connector.ConnectorFormat.EndConnect(shape, bestSite);
                }
            }
            catch { }
        }

        private static string MakeInsetName(int boxShapeId, string suffix)
        {
            return InsetNamePrefix + boxShapeId + "_" + suffix;
        }

        /// <summary>
        /// 清理旧的生成物：把包含生成物的编组（外层组、放大图单元）逐层拆开，
        /// 再删除所有放大图/锚点/连线成员，选区框保留。
        /// </summary>
        private static void CleanUpPreviousArtifacts(PowerPoint.Slide slide)
        {
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