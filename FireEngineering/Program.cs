namespace FireEngineering
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string folderPath = null;
            DateTime startDate = DateTime.MinValue;
            int maxReaderConcurrency = 1;

            // Setup log file in Logs folder with timestamp
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            string logsDir = Path.Combine(exeDir, "Logs");
            Directory.CreateDirectory(logsDir);
            string logFileName = $"log_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
            string logPath = Path.Combine(logsDir, logFileName);

            void Log(string message)
            {
                try
                {
                    lock (logPath)
                    {
                        File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n");
                    }
                }
                catch { }
            }
            void LogInfo(string message) => Log($"INFO: {message}");
            void LogError(string message) => Log($"ERROR: {message}");

            AppDomain.CurrentDomain.UnhandledException += (_, eventArgs) =>
            {
                if (eventArgs.ExceptionObject is Exception exception)
                {
                    LogError($"FATAL unhandled exception: {exception}");
                }
                else
                {
                    LogError($"FATAL unhandled exception: {eventArgs.ExceptionObject}");
                }
            };

            // Always read from appsettings.json
            string configPath = Path.Combine(exeDir, "appsettings.json");
            if (!File.Exists(configPath))
            {
                Log("Missing appsettings.json with FolderPath and StartDate.");
                return;
            }
            string outputFile = null;
            try
            {
                var json = File.ReadAllText(configPath);
                var config = System.Text.Json.JsonDocument.Parse(json).RootElement;
                folderPath = config.GetProperty("FolderPath").GetString();
                string dateStr = config.GetProperty("StartDate").GetString();
                if (!DateTime.TryParseExact(dateStr, "dd-MM-yyyy", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out startDate))
                {
                    Log("Invalid StartDate in appsettings.json. Use dd-MM-yyyy.");
                    return;
                }
                outputFile = config.TryGetProperty("OutputFile", out var outProp) ? outProp.GetString() : null;
                if (config.TryGetProperty("MaxReaderConcurrency", out var concurrencyProp) &&
                    concurrencyProp.ValueKind == System.Text.Json.JsonValueKind.Number &&
                    concurrencyProp.TryGetInt32(out var configuredConcurrency))
                {
                    if (configuredConcurrency < 1)
                    {
                        Log("MaxReaderConcurrency must be at least 1.");
                        return;
                    }

                    maxReaderConcurrency = configuredConcurrency;
                }
            }
            catch (Exception ex)
            {
                Log($"Error reading appsettings.json: {ex.Message}");
                return;
            }

            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
            {
                Log($"Folder does not exist: {folderPath}");
                return;
            }
            if (string.IsNullOrWhiteSpace(outputFile))
            {
                Log("OutputFile not specified in appsettings.json.");
                return;
            }

            string temporaryImageFolder = Path.Combine(
                Path.GetTempPath(),
                $"FireEngineeringImages_{Guid.NewGuid():N}");

            Directory.CreateDirectory(temporaryImageFolder);

            try
            {
                List<ProtocolDataModel> rows = DataReader.Read(
                    folderPath,
                    startDate,
                    temporaryImageFolder,
                    maxReaderConcurrency,
                    LogInfo,
                    (current, total) => Console.WriteLine($"Reading {current}/{total}"));
                if(rows.Count == 0)
                {
                    Log("No records found matching the criteria.");
                    return;
                }

                var seenStickerIds = new HashSet<(int StickerId, DateTime ChangeDate)>();
                rows = rows
                    .Where(row =>
                        row.StickerId == 0 ||
                        (DateTime.TryParse(row.ChangeDate, out var changeDate) &&
                         seenStickerIds.Add((row.StickerId, changeDate.Date))))
                    .ToList();

                var groups = rows
                    .Where(x => DateTime.TryParse(x.ChangeDate, out _))
                    .OrderBy(x => DateTime.Parse(x.ChangeDate))
                    .GroupBy(x => DateTime.Parse(x.ChangeDate).ToString("yyyy-MM"))
                    .ToList();

                int groupNumber = 0;
                foreach (var group in groups)
                {
                    groupNumber++;
                    Console.WriteLine($"Writing {groupNumber}/{groups.Count}");

                    string debugFile = Path.Combine(
                        Path.GetDirectoryName(outputFile)!,
                        $"{Path.GetFileNameWithoutExtension(outputFile)}___{group.Key}{Path.GetExtension(outputFile)}");

                    if (File.Exists(outputFile))
                    {
                        File.Copy(outputFile, debugFile, true);
                    }

                    var targetFile = debugFile;
                    LogInfo($"Writing month {group.Key} to {targetFile}");

                    WordTableWriter.WriteCombinedDoc(targetFile, group.ToList());
                    LogInfo($"Successfully wrote output to {targetFile}");
                }
                
                
            }
            catch (Exception ex)
            {
                LogError($"Unhandled exception: {ex.Message}\r\n{ex.StackTrace}");
            }
            finally
            {
                try
                {
                    if (Directory.Exists(temporaryImageFolder))
                    {
                        Directory.Delete(temporaryImageFolder, true);
                    }
                }
                catch (Exception ex)
                {
                    LogError($"Failed to delete temporary image folder: {ex.Message}");
                }
            }
        }
    }
}
