using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Threading;
using SlideSCI;

internal static class CoreRegression
{
    private static int passed;
    private static void Check(bool condition, string name)
    {
        if (!condition) throw new Exception(name);
        Console.WriteLine("PASS " + name);
        passed++;
    }

    private static object SettingsCall(string name, params object[] args)
    {
        return typeof(ZoomSettings).GetMethod(name, BindingFlags.Static | BindingFlags.NonPublic).Invoke(null, args);
    }

    private static ZoomSettings Load(params string[] paths)
    {
        return (ZoomSettings)SettingsCall("LoadFromPaths", new object[] { paths });
    }

    private static bool Save(ZoomSettings value, string path)
    {
        return (bool)SettingsCall("SaveToPath", value, path);
    }

    private static int Main(string[] args)
    {
        if (args.Length == 2 && args[0] == "--read-settings")
        {
            var restored = Load(args[1]);
            return restored.Magnification == 1f / 3f && restored.BoxPercent == 25f
                && restored.RecentStrokeColors == "123,456" ? 0 : 1;
        }
        // Act as a controllable child process for the real converter.
        if (args.Length > 0)
        {
            Console.InputEncoding = new UTF8Encoding(false);
            Console.OutputEncoding = new UTF8Encoding(false);
            string mode = Environment.GetEnvironmentVariable("SLIDESCI_TEST_CHILD");
            if (mode == "hang") { Thread.Sleep(90000); return 0; }
            if (mode == "error") { Console.Error.Write("intentional conversion failure"); return 7; }
            if (mode == "pipes") Console.Error.Write(new string('e', 200000));
            Console.Write(Console.In.ReadToEnd());
            return 0;
        }

        string directory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "case-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            Check(System.Configuration.ConfigurationManager.AppSettings["EnableWindowsFormsHighDpiAutoResizing"] == "true"
                && System.Configuration.ConfigurationManager.AppSettings["DpiAwareness"] == "PerMonitorV2",
                "deployed configuration initializes and retains DPI settings");
            Check(System.Configuration.ConfigurationManager.GetSection("userSettings/SlideSCI.Properties.Settings") != null,
                "deployed user settings section loads");
            float parsed;
            Check(ZoomSettings.TryParseMagnification("1/2x", out parsed) && parsed == 0.5f, "half original width preset");
            Check(ZoomSettings.TryParseMagnification("1/3", out parsed) && Math.Abs(parsed - 1f / 3f) < 0.000001f && ZoomSettings.FormatMagnification(parsed) == "1/3x", "third width survives display round trip");
            Check(ZoomSettings.TryParseMagnification("1×", out parsed) && parsed == 1, "multiplication sign supported");
            foreach (string fraction in new[] { "1/0", "1/2/3", "NaN/2", "1/Infinity", "1/20", "1/2xx" })
                Check(!ZoomSettings.TryParseMagnification(fraction, out parsed), "reject invalid fraction " + fraction);
            Check(ZoomSettings.TryParsePercent("25%", out parsed) && parsed == 25f && ZoomSettings.TryParsePercent("25", out parsed) && parsed == 25f, "percentage accepts explicit or omitted unit");
            Check(ZoomSettings.TryParseMagnification("10x", out parsed) && parsed == 10f && ZoomSettings.TryParseMagnification("1", out parsed) && parsed == 1f, "magnification accepts explicit or omitted unit");
            foreach (string invalidInput in new[] { "NaN", "Infinity", "0", "101x", "2xx", "5cm", "原图等宽", "" })
                Check(!ZoomSettings.TryParseMagnification(invalidInput, out parsed), "reject invalid magnification: " + invalidInput);
            Check(!ZoomSettings.TryParsePercent("25%%", out parsed) && !ZoomSettings.TryParsePercent("91%", out parsed), "percentage rejects repeated units and out of range");
            string current = Path.Combine(directory, "current.xml");
            string legacy = Path.Combine(directory, "legacy.xml");
            Check(Load(current).Magnification == 1f, "missing settings use defaults");
            Check(Save(new ZoomSettings { Magnification = 4f, LineColorRgb = 0x123456 }, legacy), "initial save");
            string legacyText = File.ReadAllText(legacy);
            File.WriteAllText(current, "<broken>");
            Check(Load(current, legacy).Magnification == 4f, "corrupt current falls back to legacy");
            Check(Save(new ZoomSettings { Magnification = 6f }, current), "replace existing settings");
            Check(Load(current, legacy).Magnification == 6f && File.ReadAllText(legacy) == legacyText,
                "current takes priority and legacy is unchanged");
            File.WriteAllText(current, "<ZoomSettings><Magnification>NaN</Magnification><CustomWidthCm>INF</CustomWidthCm><GapCm>-1</GapCm><LineDash>9</LineDash></ZoomSettings>");
            ZoomSettings normalized = Load(current);
            Check(normalized.Magnification == 1f && normalized.CustomWidthCm == 10f && normalized.GapCm == 0.5f && normalized.LineDash == 0,
                "nonfinite and out of range XML values normalized");
            ZoomSettings invalid = new ZoomSettings { TargetMode = (ZoomTargetMode)99, LineStyle = (ZoomLineStyle)99, LineColorRgb = -1 };
            Check(Save(invalid, current) && Load(current).TargetMode == ZoomTargetMode.Multiple && Load(current).LineColorRgb == 0,
                "invalid enums and colors normalized before serialization");
            Check((int)invalid.TargetMode == 99, "normalization does not mutate caller");
            Check(Save(new ZoomSettings { BoxLineVisible = false, LineVisible = false, LineEndArrow = 2,
                LineDash = 6, BoxLineWeight = 4.5f, LineWeight = 2.25f, RecentStrokeColors = "123,456" }, current), "extended stroke settings save");
            var stroke = Load(current);
            Check(!stroke.BoxLineVisible && !stroke.LineVisible && stroke.LineEndArrow == 2 && stroke.LineDash == 6
                && stroke.BoxLineWeight == 4.5f && stroke.LineWeight == 2.25f && stroke.RecentStrokeColors == "123,456", "extended stroke settings survive roundtrip without clamping");
            string before = File.ReadAllText(current);
            using (File.Open(current, FileMode.Open, FileAccess.Read, FileShare.None))
                Check(!Save(new ZoomSettings { Magnification = 9f }, current), "locked destination reports failure");
            Check(File.ReadAllText(current) == before && Directory.GetFiles(directory, "*.tmp").Length == 0,
                "failed replace preserves old file and removes temporary file");

            foreach (float factor in new[] { 1f, 0.5f, 1f / 3f })
                Check(Save(new ZoomSettings { Magnification = factor, BoxPercent = 25f, RecentStrokeColors = "123,456" }, current)
                    && Load(current).Magnification == factor, "consecutive preset save and reload " + factor);
            using (var reader = Process.Start(new ProcessStartInfo(Assembly.GetExecutingAssembly().Location,
                "--read-settings \"" + current + "\"") { UseShellExecute = false, CreateNoWindow = true }))
            {
                if (!reader.WaitForExit(10000)) { reader.Kill(); throw new Exception("Settings reader timed out"); }
                Check(reader.ExitCode == 0, "new process restores saved settings with deployed configuration");
            }

            string scriptDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "latex-converter");
            Directory.CreateDirectory(scriptDirectory);
            File.WriteAllText(Path.Combine(scriptDirectory, "latex-to-svg.js"), "// Test placeholder");
            var converter = new LatexToSvgConverter(Assembly.GetExecutingAssembly().Location);
            Environment.SetEnvironmentVariable("SLIDESCI_TEST_CHILD", "pipes");
            string large = "公式 α " + new string('x', 200000);
            Check(converter.ConvertLatexToSvg(large) == large, "large concurrent stdin and stderr plus UTF8 roundtrip");
            Environment.SetEnvironmentVariable("SLIDESCI_TEST_CHILD", "error");
            try { converter.ConvertLatexToSvg(large); throw new Exception("expected error"); }
            catch (InvalidOperationException ex) { Check(ex.Message.Contains("intentional conversion failure"), "early child exit retains diagnostic"); }
            Environment.SetEnvironmentVariable("SLIDESCI_TEST_CHILD", "hang");
            var watch = Stopwatch.StartNew();
            try { converter.ConvertLatexToSvg(large); throw new Exception("expected timeout"); }
            catch (InvalidOperationException ex)
            {
                Check(ex.Message.Contains("60") && watch.Elapsed.TotalSeconds >= 59 && watch.Elapsed.TotalSeconds < 70,
                    "blocked stdin respects 60 second deadline");
            }
            Console.WriteLine("TOTAL PASS " + passed);
            return 0;
        }
        catch (Exception ex) { Console.Error.WriteLine(ex); return 1; }
        finally
        {
            Environment.SetEnvironmentVariable("SLIDESCI_TEST_CHILD", null);
            // Only the unique directory created by this test is removed.
            Directory.Delete(directory, true);
        }
    }
}
