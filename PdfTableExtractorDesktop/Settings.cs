using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Win32;

namespace PdfTableExtractorDesktop
{
    public static class Settings
    {
        private const string SETTING_PARALLEL_FILEPROCESS = "ParallelFileProcess";
        private const string SETTING_ROWS_PER_PAGE = "RowsPerPage";
        private const string SETTING_ROW_COMPARISON_METHOD = "RowComparisonMethod";
        private const string SETTING_COLUMNS_PER_PAGE = "ColumnsPerPage";
        private const string SETTING_COLUMN_COMPARISON_METHOD = "ColumnComparisonMethod";
        private const string SETTING_AUTOSIZE_COLUMNS = "AutosizeColumns";
        private const string SETTING_PAGENAMING_METHOD = "PageNamingMethod";
        private const string SETTING_EMPTY_COLUMN_SKIP_METHOD = "EmptyColumnSkipMethod";
        private const string SETTING_EMPTY_ROW_SKIP_METHOD = "EmptyRowSkipMethod";
        private const string SETTING_OUTPUT_DIRECTORY_METHOD = "OutputDirectoryMethod";
        private const string SETTING_OUTPUT_DIRECTORY_CUSTOM = "OutputDirectoryCustom";
        private const string SETTING_CONTEXT_MENU_OPTION_ENABLED = "ContextMenuOptionEnabled";
        private const string SETTING_VERSION_CHECKING_DISABLED = "VersionCheckingDisabled";

        public static readonly bool VersionCheckingDisabled;
        public static readonly bool ContextMenuOptionEnabled;
        public static readonly int RowsPerPage;
        public static readonly int RowComparisonMethod;
        public static readonly int ColumnsPerPage;
        public static readonly int ColumnComparisonMethod;
        public static readonly bool ParallelExtraction;
        public static readonly bool AutosizeColumns;
        public static readonly int PageNamingMethod;
        public static readonly int EmptyColumnSkipMethod;
        public static readonly int OutputDirectoryMethod;
        public static readonly string UserSelectedOutputDirectory;
        public static readonly int EmptyRowSkipMethod;

        static Settings()
        {
            var settings = ReadSettings();

            VersionCheckingDisabled = GetJsonValue(settings, SETTING_VERSION_CHECKING_DISABLED, true);
            RowsPerPage = GetJsonValue(settings, SETTING_ROWS_PER_PAGE, 1);
            RowComparisonMethod = GetJsonValue(settings, SETTING_ROW_COMPARISON_METHOD, 2);
            ColumnsPerPage = GetJsonValue(settings, SETTING_COLUMNS_PER_PAGE, 1);
            ColumnComparisonMethod = GetJsonValue(settings, SETTING_COLUMN_COMPARISON_METHOD, 2);
            ParallelExtraction = GetJsonValue(settings, SETTING_PARALLEL_FILEPROCESS, false);
            AutosizeColumns = GetJsonValue(settings, SETTING_AUTOSIZE_COLUMNS, true);
            PageNamingMethod = GetJsonValue(settings, SETTING_PAGENAMING_METHOD, 0);
            EmptyColumnSkipMethod = GetJsonValue(settings, SETTING_EMPTY_COLUMN_SKIP_METHOD, 0);
            OutputDirectoryMethod = GetJsonValue(settings, SETTING_OUTPUT_DIRECTORY_METHOD, 0);
            EmptyRowSkipMethod = GetJsonValue(settings, SETTING_EMPTY_ROW_SKIP_METHOD, 0);
            UserSelectedOutputDirectory = GetJsonValue(settings, SETTING_OUTPUT_DIRECTORY_CUSTOM, "");
            ContextMenuOptionEnabled = GetJsonValue(settings, SETTING_CONTEXT_MENU_OPTION_ENABLED, false);
        }

        private static JsonObject ReadSettings()
        {
            var settingsFilePath = "settings.json";

            if (!File.Exists(settingsFilePath))
            {
                Console.WriteLine("No settings file, using default settings!");
                File.WriteAllText(settingsFilePath, "{}");
                return [];
            }

            try
            {
                var jsonString = File.ReadAllText(settingsFilePath);
                return JsonSerializer.Deserialize<JsonObject>(jsonString) ?? [];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading settings: {ex.Message}");
                return [];
            }
        }

        private static T GetJsonValue<T>(JsonObject json, string key, T defaultValue)
        {
            try
            {
                var asd = json.TryGetPropertyValue(key, out var node) && node is JsonValue value && value.TryGetValue(out T? result)
                    ? result
                    : defaultValue;
                return asd;
            }
            catch
            {
                return defaultValue;
            }
        }

        public static void Save(JsonObject settings, bool createContextMenu)
        {
            try
            {
                File.WriteAllText("settings.json", JsonSerializer.Serialize(settings));

                if (IsWindows())
                {
                    string registryKeyPath = @"HKEY_CLASSES_ROOT\SystemFileAssociations\.pdf\shell\PDFTableExtractor";
                    bool keyExists = Registry.GetValue(registryKeyPath, null, null) != null;

                    if (createContextMenu)
                    {
                        if (!keyExists)
                        {
                            var regFileTemplate = ADD_CONTEXT_TEMPLATE;
                            var workDir = Directory.GetCurrentDirectory().Replace("\\", "\\\\");
                            var regFilePath = Path.Combine(Directory.GetCurrentDirectory(), "regToRun.reg");

                            File.WriteAllText(regFilePath, regFileTemplate.Replace("%W", workDir));

                            ExecuteRegistryCommand(regFilePath);
                        }
                    }
                    else
                    {
                        if (keyExists)
                        {
                            var regFileTemplate = REMOVE_CONTEXT_TEMPLATE;
                            var regFilePath = Path.Combine(Directory.GetCurrentDirectory(), "regToRun.reg");

                            File.WriteAllText(regFilePath, regFileTemplate);

                            ExecuteRegistryCommand(regFilePath);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error saving settings: {e.Message}");
            }
        }

        private static void ExecuteRegistryCommand(string regFilePath)
        {
            try
            {
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "cmd",
                        Arguments = $"/c regedit.exe /S \"{regFilePath}\"",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }
                };

                process.Start();
                process.WaitForExit();
                File.Delete(regFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing registry command: {ex.Message}");
                Console.WriteLine("To do it manually, run the 'regToRun.reg' file.");
            }
        }

        public static bool IsWindows() => Environment.OSVersion.Platform is PlatformID.Win32NT;

        private static readonly string ADD_CONTEXT_TEMPLATE = """
        Windows Registry Editor Version 5.00
        [HKEY_CLASSES_ROOT\SystemFileAssociations\.pdf\shell\PDFTableExtractor]
        @="Extract Tables to Excel"
        "Icon"="\"%W\\PDFTableExtractor.ico\""
        
        [HKEY_CLASSES_ROOT\SystemFileAssociations\.pdf\shell\PDFTableExtractor\Command]
        @="cmd /c cd \"%W\" && \"%W\\PdfTableExtractorDesktop.exe\" \"%1\""
    """;

        private static readonly string REMOVE_CONTEXT_TEMPLATE = """
        Windows Registry Editor Version 5.00
        [-HKEY_CLASSES_ROOT\SystemFileAssociations\.pdf\shell\PDFTableExtractor]
    """;
    }
}
