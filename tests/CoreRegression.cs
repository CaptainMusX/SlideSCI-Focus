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
            string current = Path.Combine(directory, "current.xml");
            string legacy = Path.Combine(directory, "legacy.xml");
            Check(Load(current).Magnification == 2f, "missing settings use defaults");
            Check(Save(new ZoomSettings { Magnification = 4f, LineColorRgb = 0x123456 }, legacy), "initial save");
            string legacyText = File.ReadAllText(legacy);
            File.WriteAllText(current, "<broken>");
            Check(Load(current, legacy).Magnification == 4f, "corrupt current falls back to legacy");
            Check(Save(new ZoomSettings { Magnification = 6f }, current), "replace existing settings");
            Check(Load(current, legacy).Magnification == 6f && File.ReadAllText(legacy) == legacyText,
                "current takes priority and legacy is unchanged");
            File.WriteAllText(current, "<ZoomSettings><Magnification>NaN</Magnification><CustomWidthCm>INF</CustomWidthCm><GapCm>-1</GapCm><LineDash>9</LineDash></ZoomSettings>");
            ZoomSettings normalized = Load(current);
            Check(normalized.Magnification == 2f && normalized.CustomWidthCm == 10f && normalized.GapCm == 0.5f && normalized.LineDash == 0,
                "nonfinite and out of range XML values normalized");
            ZoomSettings invalid = new ZoomSettings { TargetMode = (ZoomTargetMode)99, LineStyle = (ZoomLineStyle)99, LineColorRgb = -1 };
            Check(Save(invalid, current) && Load(current).TargetMode == ZoomTargetMode.SameAsOriginal && Load(current).LineColorRgb == 0,
                "invalid enums and colors normalized before serialization");
            Check((int)invalid.TargetMode == 99, "normalization does not mutate caller");
            string before = File.ReadAllText(current);
            using (File.Open(current, FileMode.Open, FileAccess.Read, FileShare.None))
                Check(!Save(new ZoomSettings { Magnification = 9f }, current), "locked destination reports failure");
            Check(File.ReadAllText(current) == before && Directory.GetFiles(directory, "*.tmp").Length == 0,
                "failed replace preserves old file and removes temporary file");

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
