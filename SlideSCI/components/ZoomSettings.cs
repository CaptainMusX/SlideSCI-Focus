using System;
using System.IO;
using System.Diagnostics;
using System.Xml.Serialization;

namespace SlideSCI
{
    /// <summary>
    /// 「局部放大」的持久化设置。
    /// 保存于 %APPDATA%\SlideSCI Focus\zoom_settings.xml，供「生成放大图」单击直出时复用。
    /// </summary>
    public class ZoomSettings
    {
        public ZoomTargetMode TargetMode { get; set; } = ZoomTargetMode.Multiple;
        public float Magnification { get; set; } = 1f;
        public float CustomWidthCm { get; set; } = 10f;
        /// <summary>仅在「与原图相同尺寸」时生效：保持选区比例、按原图范围等比放大，不拉伸。</summary>
        public bool KeepAspect { get; set; } = true;
        public float GapCm { get; set; } = 0.5f;
        public ZoomLineStyle LineStyle { get; set; } = ZoomLineStyle.JournalFunnel;
        public float LineWeight { get; set; } = 1f;
        public float BoxLineWeight { get; set; } = 1.5f;
        public int LineColorRgb { get; set; } = 0x000000;
        public int BoxColorRgb { get; set; } = 0x000000;
        /// <summary>0 = 实线，1 = 虚线。</summary>
        public int LineDash { get; set; }
        public bool Group { get; set; } = false;
        public int BoxLineDash { get; set; }
        public bool BoxLineVisible { get; set; } = true;
        public bool LineVisible { get; set; } = true;
        public int LineBeginArrow { get; set; } = 1;
        public int LineEndArrow { get; set; } = 1;
        public string RecentStrokeColors { get; set; } = "";
        public float BoxPercent { get; set; } = 40f;

        public static ZoomSettings CreateDefault() => new ZoomSettings();

        public static bool TryParsePercent(string text, out float value) => TryParseUnit(text, "%", 5f, 90f, out value);
        public static bool TryParseMagnification(string text, out float value) => TryParseUnit(text, "x", 0.1f, 100f, out value);

        private static bool TryParseUnit(string text, string unit, float min, float max, out float value)
        {
            string input = (text ?? string.Empty).Trim();
            if (input.EndsWith(unit, StringComparison.OrdinalIgnoreCase))
                input = input.Substring(0, input.Length - unit.Length).TrimEnd();
            bool parsed = float.TryParse(input, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.CurrentCulture, out value)
                || float.TryParse(input, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value);
            return parsed && !float.IsNaN(value) && !float.IsInfinity(value) && value >= min && value <= max;
        }

        /// <summary>读取设置；文件不存在时返回默认值。</summary>
        public static ZoomSettings LoadOrDefault()
        {
            return LoadFromPaths(GetConfigPath(), GetLegacyConfigPath());
        }

        internal static ZoomSettings LoadFromPaths(params string[] paths)
        {
            // Read the new location first. The old location is intentionally
            // retained as a read-only migration source for earlier releases.
            foreach (string path in paths)
            {
                if (!File.Exists(path)) continue;
                try
                {
                    using (var reader = new StreamReader(path))
                    {
                        var serializer = new XmlSerializer(typeof(ZoomSettings));
                        var loaded = serializer.Deserialize(reader) as ZoomSettings;
                        if (loaded != null) return loaded.NormalizedCopy();
                    }
                }
                catch
                {
                    // Try the next location, then fall back to defaults.
                }
            }
            return CreateDefault();
        }

        /// <summary>是否已存在已保存的设置（用于决定首次单击是否先弹设置）。</summary>
        public static bool HasSavedSettings()
        {
            return File.Exists(GetConfigPath()) || File.Exists(GetLegacyConfigPath());
        }

        public static bool Save(ZoomSettings settings)
        {
            return SaveToPath(settings, GetConfigPath());
        }

        internal static bool SaveToPath(ZoomSettings settings, string path)
        {
            string temporaryPath = null;
            try
            {
                string directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
                temporaryPath = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                using (var writer = new StreamWriter(temporaryPath, false, new System.Text.UTF8Encoding(false)))
                {
                    var serializer = new XmlSerializer(typeof(ZoomSettings));
                    serializer.Serialize(writer, (settings ?? CreateDefault()).NormalizedCopy());
                }
                // Commit only a complete document; a failed write preserves the
                // previous settings. The migration source is never modified.
                if (File.Exists(path)) File.Replace(temporaryPath, path, null);
                else File.Move(temporaryPath, path);
                return true;
            }
            catch (Exception ex)
            {
                Trace.TraceWarning("Unable to save zoom settings: {0}", ex.Message);
                return false;
            }
            finally
            {
                if (temporaryPath != null)
                {
                    try { File.Delete(temporaryPath); }
                    catch (Exception ex) { Trace.TraceWarning("Unable to remove temporary settings: {0}", ex.Message); }
                }
            }
        }

        internal ZoomSettings NormalizedCopy()
        {
            var copy = (ZoomSettings)MemberwiseClone();
            if (!Enum.IsDefined(typeof(ZoomTargetMode), copy.TargetMode)) copy.TargetMode = ZoomTargetMode.Multiple;
            if (!Enum.IsDefined(typeof(ZoomLineStyle), copy.LineStyle)) copy.LineStyle = ZoomLineStyle.JournalFunnel;
            copy.Magnification = ValidNumber(copy.Magnification, 0.1f, 100f, 1f);
            copy.CustomWidthCm = ValidNumber(copy.CustomWidthCm, 0.1f, 200f, 10f);
            copy.GapCm = ValidNumber(copy.GapCm, 0f, 50f, 0.5f);
            copy.LineWeight = ValidNumber(copy.LineWeight, 0.25f, 6f, 1f);
            copy.BoxLineWeight = ValidNumber(copy.BoxLineWeight, 0.25f, 6f, 1.5f);
            if (copy.LineColorRgb < 0 || copy.LineColorRgb > 0xFFFFFF) copy.LineColorRgb = 0;
            if (copy.BoxColorRgb < 0 || copy.BoxColorRgb > 0xFFFFFF) copy.BoxColorRgb = 0;
            if (copy.LineDash < 0 || copy.LineDash > 7) copy.LineDash = 0;
            if (copy.BoxLineDash < 0 || copy.BoxLineDash > 7) copy.BoxLineDash = 0;
            if (copy.LineBeginArrow < 1 || copy.LineBeginArrow > 6) copy.LineBeginArrow = 1;
            if (copy.LineEndArrow < 1 || copy.LineEndArrow > 6) copy.LineEndArrow = 1;
            copy.BoxPercent = ValidNumber(copy.BoxPercent, 5f, 90f, 40f);
            return copy;
        }

        public static Microsoft.Office.Core.MsoLineDashStyle GetDashStyle(int value)
        {
            switch (value)
            {
                case 1: return Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
                case 2: return Microsoft.Office.Core.MsoLineDashStyle.msoLineRoundDot;
                case 3: return Microsoft.Office.Core.MsoLineDashStyle.msoLineDashDot;
                case 4: return Microsoft.Office.Core.MsoLineDashStyle.msoLineSquareDot;
                case 5: return Microsoft.Office.Core.MsoLineDashStyle.msoLineDashDotDot;
                case 6: return Microsoft.Office.Core.MsoLineDashStyle.msoLineLongDash;
                case 7: return Microsoft.Office.Core.MsoLineDashStyle.msoLineLongDashDot;
                default: return Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid;
            }
        }

        internal void ApplyStroke(Microsoft.Office.Interop.PowerPoint.Shape shape, bool box)
        {
            var line = shape.Line;
            line.Weight = box ? BoxLineWeight : LineWeight;
            line.ForeColor.RGB = box ? BoxColorRgb : LineColorRgb;
            line.DashStyle = GetDashStyle(box ? BoxLineDash : LineDash);
            // PowerPoint rejects arrowhead assignments on closed selection
            // outlines, including the no-arrow value. Only connectors support it.
            if (!box)
            {
                line.BeginArrowheadStyle = (Microsoft.Office.Core.MsoArrowheadStyle)LineBeginArrow;
                line.EndArrowheadStyle = (Microsoft.Office.Core.MsoArrowheadStyle)LineEndArrow;
            }
            line.Visible = (box ? BoxLineVisible : LineVisible) ? Microsoft.Office.Core.MsoTriState.msoTrue : Microsoft.Office.Core.MsoTriState.msoFalse;
        }

        private static float ValidNumber(float value, float minimum, float maximum, float fallback)
        {
            return float.IsNaN(value) || float.IsInfinity(value) || value < minimum || value > maximum
                ? fallback : value;
        }

        private static string GetConfigPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "SlideSCI Focus", "zoom_settings.xml");
        }

        private static string GetLegacyConfigPath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "SlideSCI", "zoom_settings.xml");
        }
    }
}
