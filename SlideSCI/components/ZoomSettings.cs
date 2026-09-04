using System;
using System.IO;
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
        public bool Group { get; set; } = true;

        public static ZoomSettings CreateDefault() => new ZoomSettings();

        /// <summary>读取设置；文件不存在时返回默认值。</summary>
        public static ZoomSettings LoadOrDefault()
        {
            // Read the new location first. The old location is intentionally
            // retained as a read-only migration source for earlier releases.
            foreach (string path in new[] { GetConfigPath(), GetLegacyConfigPath() })
            {
                if (!File.Exists(path)) continue;
                try
                {
                    using (var reader = new StreamReader(path))
                    {
                        var serializer = new XmlSerializer(typeof(ZoomSettings));
                        var loaded = serializer.Deserialize(reader) as ZoomSettings;
                        if (loaded != null) return loaded;
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

        public static void Save(ZoomSettings settings)
        {
            try
            {
                string directory = Path.GetDirectoryName(GetConfigPath());
                if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
                using (var writer = new StreamWriter(GetConfigPath(), false, new System.Text.UTF8Encoding(false)))
                {
                    var serializer = new XmlSerializer(typeof(ZoomSettings));
                    serializer.Serialize(writer, settings ?? CreateDefault());
                }
            }
            catch { }
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
