using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;

namespace FireEngineering
{

    public class NumberingState
    {
        public Dictionary<(int numId, int ilvl), int> Counters = new();
    }

    public class DataReader
    {
        public static List<ProtocolDataModel> Read(
            string folderPath,
            DateTime startDate,
            string imageFolder,
            int maxReaderConcurrency = 1,
            Action<string> logInfo = null,
            Action<int, int> readProgress = null)
        {
            var recordsByFile = new System.Collections.Concurrent.ConcurrentDictionary<int, List<ProtocolDataModel>>();

            // Recursively get all .docx files in the folder and subfolders
            string[] files = Directory.GetFiles(folderPath, "*.docx", SearchOption.AllDirectories)
                .Where(f => File.GetLastWriteTime(f) >= startDate)
                .ToArray();

            string[] nonDocxFiles = Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories)
                .Where(f => !string.Equals(Path.GetExtension(f), ".docx", StringComparison.OrdinalIgnoreCase))
                .ToArray();

            foreach (var file in nonDocxFiles)
            {
                logInfo?.Invoke($"Skipping non-.docx file: {file}");
            }

            int totalFiles = files.Length;
            int fileNumber = 0;
            var parallelOptions = new ParallelOptions
            {
                MaxDegreeOfParallelism = Math.Max(1, maxReaderConcurrency)
            };

            Parallel.ForEach(files.Select((file, index) => (file, index)), parallelOptions, fileItem =>
            {
                string file = fileItem.file;
                int fileIndex = fileItem.index;
                var fileRecords = new List<ProtocolDataModel>();
                int currentFileNumber = Interlocked.Increment(ref fileNumber);
                readProgress?.Invoke(currentFileNumber, totalFiles);

                if (Path.GetFileName(file).StartsWith("~$", StringComparison.Ordinal))
                {
                    logInfo?.Invoke($"Skipping Word temporary lock file: {file}");
                    recordsByFile[fileIndex] = fileRecords;
                    return;
                }

                try
                {
                    logInfo?.Invoke($"Reading file: {file}");
                    var fileMonths = new HashSet<string>();

                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(file, false))
                    {
                        using (var reader = OpenXmlReader.Create(wordDoc.MainDocumentPart))
                        {
                            NumberingState state = null;
                            ProtocolDataModel lastRecord = null;
                            int tableDepth = -1;
                            int tableRowCount = 0;

                            while (reader.Read())
                            {
                                if (reader.ElementType == typeof(Table))
                                {
                                    if (reader.IsStartElement)
                                    {
                                        tableDepth = reader.Depth;
                                        tableRowCount = 0;
                                        state = new NumberingState();
                                        lastRecord = new ProtocolDataModel();
                                    }
                                    else if (reader.Depth == tableDepth)
                                    {
                                        tableDepth = -1;
                                    }

                                    continue;
                                }

                                if (!reader.IsStartElement ||
                                    reader.ElementType != typeof(TableRow) ||
                                    reader.Depth != tableDepth + 1)
                                {
                                    continue;
                                }

                                tableRowCount++;
                                var row = (TableRow)reader.LoadCurrentElement();
                                if (tableRowCount == 1) continue; // skip header row

                                var cells = row.Elements<TableCell>().ToList();
                                if (cells.Count < 11) continue; // now 11 columns expected

                            string changeDate = !string.IsNullOrWhiteSpace(cells[7].InnerText) ? cells[7].InnerText.Trim() : lastRecord.ChangeDate;
                            // Row-level filter: only include if changeDate >= startDate
                            DateTime rowChangeDate;
                            if (!DateTime.TryParse(changeDate, out rowChangeDate) || rowChangeDate < startDate)
                            {
                                continue;
                            }

                            var stickerCell = cells[10];
                            string stickerValue = null;

                            if (string.IsNullOrEmpty(stickerCell.InnerText?.Trim()))
                            {
                                var para = stickerCell.Elements<Paragraph>().FirstOrDefault();
                                if (para != null)
                                {
                                    stickerValue = GetNumberedParagraphText(para, wordDoc, state);
                                }
                            }
                            else
                            {
                                stickerValue = stickerCell.InnerText?.Trim();
                            }

                            int stickerId = 0;
                            int.TryParse(stickerValue, out stickerId);

                            // stable columns (carry forward if empty)
                            var (brand, serial) = !string.IsNullOrWhiteSpace(cells[1].InnerText) ? SplitBrandAndSerial(cells[1].InnerText.Trim()) : (lastRecord.BrandAndModel, lastRecord.Serial);
                            string category = !string.IsNullOrWhiteSpace(cells[2].InnerText) ? cells[2].InnerText.Trim() : lastRecord.Category;
                            double capacity = !string.IsNullOrWhiteSpace(cells[3].InnerText) ? (double.TryParse(cells[3].InnerText, out var cap) ? cap : 0) : lastRecord.Capacity;
                            string type = !string.IsNullOrWhiteSpace(cells[4].InnerText) ? cells[4].InnerText.Trim() : lastRecord.Type;
                            string changeType = !string.IsNullOrWhiteSpace(cells[5].InnerText) ? cells[5].InnerText.Trim() : lastRecord.ChangeType;
                            string changed = !string.IsNullOrWhiteSpace(cells[6].InnerText) ? cells[6].InnerText.Trim() : lastRecord.Changed;
                            string changeName = !string.IsNullOrWhiteSpace(cells[8].InnerText) ? cells[8].InnerText.Trim() : lastRecord.WorkerName;

                            logInfo?.Invoke($"Starting signature extraction: {file}");
                            var (data, rotation) = GetImageFromCell(cells[9], wordDoc, imageFolder);
                            logInfo?.Invoke($"Finished signature extraction: {file}");

                            if (data == null)
                            {
                                logInfo?.Invoke($"Signature image was skipped or unavailable in file: {file}");
                            }

                            var signature = data != null ? (data, rotation) : lastRecord.Signature;

                            var currentRecord = new ProtocolDataModel()
                            {
                                BrandAndModel = brand,
                                Serial = serial,
                                Category = category,
                                Capacity = capacity,
                                ChangeType = changeType,
                                Changed = changed,
                                ChangeDate = changeDate,
                                Type = type,
                                WorkerName = changeName,
                                Signature = signature,
                                StickerId = stickerId
                            };

                            // Add record
                            fileRecords.Add(currentRecord);

                            string month = rowChangeDate.ToString("yyyy-MM");
                            if (fileMonths.Add(month))
                            {
                                logInfo?.Invoke($"Month {month} received records from file: {file}");
                            }

                                lastRecord = currentRecord;
                            }
                        }
                    }

                    logInfo?.Invoke($"Finished reading file: {file}");
                }
                catch (Exception ex)
                {
                    logInfo?.Invoke(
                        $"ERROR: Failed to read file: {file}{Environment.NewLine}" +
                        $"Message: {ex.Message}{Environment.NewLine}" +
                        $"StackTrace: {ex.StackTrace}");
                }
                recordsByFile[fileIndex] = fileRecords;
            });

            return recordsByFile
                .OrderBy(item => item.Key)
                .SelectMany(item => item.Value)
                .ToList();
        }

        public static string GetNumberedParagraphText(
    Paragraph p,
    WordprocessingDocument doc,
    NumberingState state)
        {
            var text = string.Join("", p.Descendants<Text>().Select(t => t.Text)).Trim();

            var numPr = p.Descendants<NumberingProperties>().FirstOrDefault();
            if (numPr == null)
                return ""; // Not numbered

            var numId = numPr.NumberingId?.Val?.Value;
            var ilvl = numPr.NumberingLevelReference?.Val?.Value ?? 0;

            if (numId == null)
                return "";

            var key = (numId.Value, ilvl);

            // Load numbering definitions
            var numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
            var numbering = numberingPart.Numbering;

            // Get the abstract numbering ID for this numId
            var abstractNumId = numbering
                .Elements<NumberingInstance>()
                .First(inst => inst.NumberID == numId)
                .AbstractNumId.Val.Value;

            var abstractNum = numbering
                .Elements<AbstractNum>()
                .First(a => a.AbstractNumberId == abstractNumId);

            // Get this level definition
            var level = abstractNum
                .Elements<Level>()
                .First(l => l.LevelIndex == ilvl);

            // Check for custom start value
            int startValue = 1;
            var start = level.Elements<StartNumberingValue>().FirstOrDefault();
            if (start != null && start.Val != null)
                startValue = (int)start.Val.Value;

            // Initialize or increment counter
            if (!state.Counters.ContainsKey(key))
                state.Counters[key] = startValue;
            else
                state.Counters[key]++;

            int number = state.Counters[key];

            return number.ToString();
        }

        public static (string FilePath, int Rotation) GetImageFromCell(
            TableCell cell,
            WordprocessingDocument doc,
            string imageFolder)
        {
            var drawing = cell.Descendants<Drawing>().FirstOrDefault();
            if (drawing == null) return (null, 0);

            // Look for Blip reference
            var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
            if (blip == null) return (null, 0);

            string embedId = blip.Embed.Value;
            var imagePart = (ImagePart)doc.MainDocumentPart.GetPartById(embedId);

            int rotation = 0;
            var xfrm = drawing.Descendants<A.Transform2D>().FirstOrDefault();
            if (xfrm?.Rotation != null)
            {
                rotation = (int)xfrm.Rotation.Value; // in 1/60000 degree
            }

            Directory.CreateDirectory(imageFolder);
            string imagePath = Path.Combine(imageFolder, $"{Guid.NewGuid():N}.img");

            try
            {
                using var source = imagePart.GetStream();
                using var output = new FileStream(imagePath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                source.CopyTo(output);
                output.Flush();

                return (imagePath, rotation);
            }
            catch
            {
                if (File.Exists(imagePath))
                {
                    File.Delete(imagePath);
                }

                throw;
            }
        }

        private static (string BrandText, string Serial) SplitBrandAndSerial(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return ("", "");

            int idx = input.IndexOf("Сериен №");
            if (idx > -1)
            {
                string brand = input.Substring(0, idx).Trim();
                string serial = input.Substring(idx + 8).Trim();
                return (brand, serial);
            }

            return (input.Trim(), "");
        }
    }
}
