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
        public ZoomTargetMode TargetMode { get; set; } = ZoomTargetMode.SameAsOriginal;
        public float Magnification { get; set; } = 2f;
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
        public float BoxPercent { get; set; } = 40f;

        public static ZoomSettings CreateDefault() => new ZoomSettings();

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
            if (!Enum.IsDefined(typeof(ZoomTargetMode), copy.TargetMode)) copy.TargetMode = ZoomTargetMode.SameAsOriginal;
            if (!Enum.IsDefined(typeof(ZoomLineStyle), copy.LineStyle)) copy.LineStyle = ZoomLineStyle.JournalFunnel;
            copy.Magnification = ValidNumber(copy.Magnification, 0.1f, 100f, 2f);
            copy.CustomWidthCm = ValidNumber(copy.CustomWidthCm, 0.1f, 200f, 10f);
            copy.GapCm = ValidNumber(copy.GapCm, 0f, 50f, 0.5f);
            copy.LineWeight = ValidNumber(copy.LineWeight, 0.5f, 2f, 1f);
            copy.BoxLineWeight = ValidNumber(copy.BoxLineWeight, 0.5f, 3f, 1.5f);
            if (copy.LineColorRgb < 0 || copy.LineColorRgb > 0xFFFFFF) copy.LineColorRgb = 0;
            if (copy.BoxColorRgb < 0 || copy.BoxColorRgb > 0xFFFFFF) copy.BoxColorRgb = 0;
            if (copy.LineDash < 0 || copy.LineDash > 3) copy.LineDash = 0;
            if (copy.BoxLineDash < 0 || copy.BoxLineDash > 3) copy.BoxLineDash = 0;
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
                default: return Microsoft.Office.Core.MsoLineDashStyle.msoLineSolid;
            }
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
