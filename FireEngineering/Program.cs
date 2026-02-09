namespace FireEngineering
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string folderPath = null;
            DateTime startDate = DateTime.MinValue;

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
                    File.AppendAllText(logPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}\r\n");
                }
                catch { }
            }
            void LogInfo(string message) => Log($"INFO: {message}");
            void LogError(string message) => Log($"ERROR: {message}");

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

            try
            {
                List<ProtocolDataModel> rows = DataReader.Read(folderPath, startDate, LogInfo);
                if(rows.Count == 0)
                {
                    Log("No records found matching the criteria.");
                    return;
                }
                string targetFile = outputFile;
#if DEBUG
                string debugFile = Path.Combine(Path.GetDirectoryName(outputFile), Path.GetFileNameWithoutExtension(outputFile) + Guid.NewGuid() + Path.GetExtension(outputFile));
                if (File.Exists(outputFile))
                {
                    File.Copy(outputFile, debugFile, true);
                }
                targetFile = debugFile;
                LogInfo($"DEBUG: Writing to debug copy: {targetFile}");
#endif
                WordTableWriter.WriteCombinedDoc(targetFile, rows);
                LogInfo($"Successfully wrote output to {targetFile}");
            }
            catch (Exception ex)
            {
                LogError($"Unhandled exception: {ex.Message}\r\n{ex.StackTrace}");
            }
        }
    }
}
