using System.Diagnostics;
using System.Net.Http.Headers;
using System.Security.Principal;
using System.Text.Json;
using PdfTableExtractorLib;

namespace PdfTableExtractorDesktop
{
    public static class Program
    {
        public const string VERSION = "1.0.0";


        private static readonly HttpClient httpClient = new() { Timeout = TimeSpan.FromSeconds(7) };


        [STAThread]
        public static async Task Main(string[] args)
        {
            if (args.Length is 0)
            {
                ShowSettingsUI();
                await ShowVersionCheckResult(GetVersionCheckResultMessage());
                return;
            }
            else
            {
                var rowComparisonFunction = PDFTableExtractor.GetComparisonFunction(Settings.RowComparisonMethod, Settings.RowsPerPage);
                var columnComparisonFunction = PDFTableExtractor.GetComparisonFunction(Settings.ColumnComparisonMethod, Settings.ColumnsPerPage);
                var pageNamingFunction = PDFTableExtractor.GetPageNamingFunction(Settings.PageNamingMethod);
                var outputPathCreatorFunction = GetOutputPathCreatorFunction(Settings.OutputDirectoryMethod);

                var pdfFiles = (Settings.ParallelExtraction ? args.AsParallel().WithDegreeOfParallelism(4) : args.AsEnumerable())
                            .Select(Path.GetFullPath)
                            .Where(IsPdfFile)
                            .ToList();

                if (pdfFiles.Count is 0)
                {
                    Console.WriteLine("No valid PDF files found.");
                    return;
                }

                pdfFiles.ForEach(filePath =>
                {
                    Console.WriteLine($"Opening file: {filePath}");
                    HandlePdfExtraction(filePath, pageNamingFunction, outputPathCreatorFunction, rowComparisonFunction, columnComparisonFunction);
                });

                await ShowVersionCheckResult(GetVersionCheckResultMessage());
                Console.Write("All done, press enter");
                Console.ReadLine();
            }
        }

        private static void ShowSettingsUI()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form settingsForm = new()
            { 
                Text = "PDF Table Extractor Settings",
                Width = 650, Height = 650,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath),
                MaximizeBox = false, // Prevent resizing
                MinimizeBox = false
            };

            // Create settings panel
            Components components = new Components();
            Panel settingsPanel = components.CreateSettingsPanel();
            settingsPanel.Location = new Point(10, 10);
            settingsPanel.Size = new Size(600, 600);

            // Add panel to form
            settingsForm.Controls.Add(settingsPanel);

            Application.Run(settingsForm);
        }

        private static bool IsPdfFile(string filePath)
        {
            if (!filePath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"{filePath} is not a PDF file");
                return false;
            }
            return true;
        }

        private static void HandlePdfExtraction(
            string inputFile,
            PageNamingFunction pageNamingFunction,
            Func<string, string> outputPathCreatorFunction,
            Predicate<int> rowComparisonFunction,
            Predicate<int> columnComparisonFunction)
        {
            Console.WriteLine($"Extracting data from: {inputFile}");

            try
            {
                using var excelOutput = PDFTableExtractor.Extract(
                    Settings.AutosizeColumns,
                    Settings.EmptyColumnSkipMethod,
                    Settings.EmptyRowSkipMethod,
                    rowComparisonFunction,
                    columnComparisonFunction,
                    pageNamingFunction,
                    inputFile);

                var numberOfSheets = excelOutput.NumberOfSheets;
                if (numberOfSheets > 0)
                {
                    var outputFile = Path.Combine(
                        outputPathCreatorFunction(inputFile) + ".xlsx"
                    );

                    Console.WriteLine($"Writing {numberOfSheets} sheets to file: {outputFile}");

                    try
                    {
                        using (var outputStream = new FileStream(outputFile, FileMode.Create, FileAccess.Write))
                        {
                            excelOutput.Write(outputStream);
                        }

                        Console.WriteLine($"Finished: {outputFile}\n");
                    }
                    catch (UnauthorizedAccessException)
                    {
                        Console.WriteLine($"Access denied for {outputFile}. Trying to elevate permissions...");

                        if (!IsAdministrator())
                        {
                            RequestAdminPrivileges();
                        }
                        else
                        {
                            Console.WriteLine("Even with admin rights, access is denied.");
                        }
                    }
                    catch (IOException ioEx)
                    {
                        Console.WriteLine($"File I/O error: {ioEx.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"No sheets were extracted from file: {inputFile}");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("An error occurred, check error.txt for details");
                try
                {
                    File.WriteAllText("error.txt", e.ToString());
                }
                catch (IOException)
                {
                    Console.WriteLine("Failed to write to error.txt");
                }
            }
        }

        private static async Task ShowVersionCheckResult(Task<string> versionCheckResultMessage)
        {
            try
            {
                string result = await versionCheckResultMessage;
                Console.WriteLine(result);
            }
            catch (Exception e)
            {
                Console.WriteLine("Version checking failed: " + e.Message);
            }
        }

        private static Func<string, string> GetOutputPathCreatorFunction(int outDirMethod)
        {
            return outDirMethod switch
            {
                0 => path => StripExtension(path),
                1 => path => Components.ShowExcelDirectoryPicker(Path.GetDirectoryName(path) ?? string.Empty) + "\\" + StripExtension(Path.GetFileName(path)),
                _ => path => Settings.UserSelectedOutputDirectory + "\\" + StripExtension(Path.GetFileName(path)),
            };
        }

        private static string StripExtension(string path)
        {
            return Path.HasExtension(path) ? path[..path.LastIndexOf('.')] : path;
        }

        private static async Task<string> GetVersionCheckResultMessage()
        {
            if (Settings.VersionCheckingDisabled)
            {
                return $"Local app version: {VERSION}, version checking was disabled in config";
            }

            const string apiUrl = "https://api.github.com/repos/ogunerkutay/PdfTableExtractorDesktop/releases/latest";
            httpClient.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("Mozilla", "5.0"));

            try
            {
                var response = await httpClient.GetStringAsync(apiUrl);
                using var doc = JsonDocument.Parse(response);

                if (doc.RootElement.TryGetProperty("tag_name", out JsonElement tagNameElement))
                {
                    if (tagNameElement.ValueKind is not JsonValueKind.String)
                        return "Error: 'tag_name' in the JSON response is not a string.";

                    string remoteVersion = tagNameElement.GetString()!;

                    return remoteVersion == VERSION
                        ? $"Local app version: {VERSION} is up to date!"
                        : $"Local app version: {VERSION} is out of date! Remote app version: {remoteVersion}\nCheck https://github.com/Degubi/PDFTableExtractor for new release!";
                }
                else
                {
                    return "Error: Could not find 'tag_name' in the JSON response.";
                }
            }
            catch (HttpRequestException ex)
            {
                return $"Network error: Unable to get app version from repository. {ex.Message}";
            }
            catch (JsonException ex)
            {
                return $"JSON parsing error: Invalid response from repository. {ex.Message}";
            }
            catch (Exception ex)
            {
                return $"Unexpected error: {ex.Message}";
            }
        }

        public static class JsonHelper
        {
            public static readonly JsonSerializerOptions JsonOptions = new()
            {
                WriteIndented = true // Equivalent to withFormatting(Boolean.TRUE)
            };
        }

        /// <summary>
        /// Checks if the application is running with administrator privileges.
        /// </summary>
        private static bool IsAdministrator()
        {
            using WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new(identity);
            return principal.IsInRole(WindowsBuiltInRole.Administrator);
        }

        /// <summary>
        /// Relaunches the application with administrator privileges.
        /// </summary>
        private static void RequestAdminPrivileges()
        {
            ProcessStartInfo psi = new()
            {
                FileName = Application.ExecutablePath, // Restart the same app
                UseShellExecute = true,
                Verb = "runas" // Request UAC elevation
            };

            try
            {
                Process.Start(psi);
                Environment.Exit(0); // Close the current instance
            }
            catch
            {
                Console.WriteLine("User denied UAC elevation.");
            }
        }
    }

}
