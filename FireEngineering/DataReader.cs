using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;

namespace FireEngineering
{

    public class NumberingState
    {
        public Dictionary<(int numId, int ilvl), int> Counters = new();
    }

    public class DataReader
    {

        public static List<ProtocolDataModel> Read(string folderPath, DateTime startDate, Action<string> logInfo = null)
        {
            List<ProtocolDataModel> allRecords = new List<ProtocolDataModel>();

            // Recursively get all .docx files in the folder and subfolders
            string[] files = Directory.GetFiles(folderPath, "*.docx", SearchOption.AllDirectories)
                .Where(f => File.GetLastWriteTime(f) >= startDate)
                .ToArray();

            foreach (var file in files)
            {
                logInfo?.Invoke($"Reading file: {file}");
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(file, false))
                {
                    var tables = wordDoc.MainDocumentPart.Document.Body.Elements<Table>();

                    foreach (var table in tables)
                    {
                        if (table == null) continue;

                        var rows = table.Elements<TableRow>().Skip(1); // skip header row

                        ProtocolDataModel lastRecord = new ProtocolDataModel();
                        foreach (var row in rows)
                        {
                            var state = new NumberingState();

                            foreach (var cell in row.Elements<TableCell>())
                            {
                                foreach (var p in cell.Elements<Paragraph>())
                                {
                                    string numberedText = GetNumberedParagraphText(p, wordDoc, state);
                                    // Logging removed, was: Console.WriteLine(numberedText);
                                }
                            }

                            var cells = row.Elements<TableCell>().ToList();
                            if (cells.Count < 11) continue; // now 11 columns expected

                            // stable columns (carry forward if empty)
                            var (brand, serial) = !string.IsNullOrWhiteSpace(cells[1].InnerText) ? SplitBrandAndSerial(cells[1].InnerText.Trim()) : (lastRecord.BrandAndModel, lastRecord.Serial);
                            string category = !string.IsNullOrWhiteSpace(cells[2].InnerText) ? cells[2].InnerText.Trim() : lastRecord.Category;
                            double capacity = !string.IsNullOrWhiteSpace(cells[3].InnerText) ? (double.TryParse(cells[3].InnerText, out var cap) ? cap : 0) : lastRecord.Capacity;
                            string type = !string.IsNullOrWhiteSpace(cells[4].InnerText) ? cells[4].InnerText.Trim() : lastRecord.Type;
                            string changeType = !string.IsNullOrWhiteSpace(cells[5].InnerText) ? cells[5].InnerText.Trim() : lastRecord.ChangeType;
                            string changed = !string.IsNullOrWhiteSpace(cells[6].InnerText) ? cells[6].InnerText.Trim() : lastRecord.Changed;
                            string changeDate = !string.IsNullOrWhiteSpace(cells[7].InnerText) ? cells[7].InnerText.Trim() : lastRecord.ChangeDate;
                            string changeName = !string.IsNullOrWhiteSpace(cells[8].InnerText) ? cells[8].InnerText.Trim() : lastRecord.WorkerName;

                            var (data, rotation) = GetImageFromCell(cells[9], wordDoc);
                            var signature = data != null ? (data, rotation) : lastRecord.Signature;
                            int stickerId = int.TryParse(cells[10].InnerText, out var id) ? id : 0;

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
                            allRecords.Add(currentRecord);

                            lastRecord = currentRecord;

                            var test = GetCellValues(cells[10]);
                        }
                    }
                }
            }

            return allRecords;
        }

        public static string GetNumberedParagraphText(
    Paragraph p,
    WordprocessingDocument doc,
    NumberingState state)
        {
            string text = string.Join("", p.Descendants<Text>().Select(t => t.Text));

            var numPr = p.Descendants<NumberingProperties>().FirstOrDefault();
            if (numPr == null)
                return text; // Not numbered

            var numId = numPr.NumberingId?.Val?.Value;
            var ilvl = numPr.NumberingLevelReference?.Val?.Value ?? 0;

            if (numId == null)
                return text;

            var key = (numId.Value, ilvl);

            // Increase our counter
            if (!state.Counters.ContainsKey(key))
                state.Counters[key] = 1;
            else
                state.Counters[key]++;

            int number = state.Counters[key];

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

            string lvlText = level.LevelText?.Val ?? "%1.";

            // Replace %1, %2, etc. with actual numbers
            string formattedNumber = lvlText.Replace("%" + (ilvl + 1).ToString(), number.ToString());

            return formattedNumber + " " + text;
        }

        private static List<string> GetCellValues(TableCell cell)
        {
            if (cell == null) return new List<string>();

            // Each paragraph inside a cell can represent a separate value
            var values = cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>()
                             .Select(p => string.Join(" ",
                                  p.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>()
                                   .Select(t => t.Text)))
                             .Where(v => !string.IsNullOrWhiteSpace(v))
                             .ToList();

            return values;
        }

        public static (byte[] ImageData, int Rotation) GetImageFromCell(TableCell cell, WordprocessingDocument doc)
        {
            var drawing = cell.Descendants<Drawing>().FirstOrDefault();
            if (drawing == null) return (null, 0);

            // Look for Blip reference
            var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
            if (blip == null) return (null, 0);

            string embedId = blip.Embed.Value;
            var imagePart = (ImagePart)doc.MainDocumentPart.GetPartById(embedId);

            byte[] imageBytes;
            using (var stream = new MemoryStream())
            {
                imagePart.GetStream().CopyTo(stream);
                imageBytes = stream.ToArray();
            }

            // Check rotation
            int rotation = 0;
            var xfrm = drawing.Descendants<A.Transform2D>().FirstOrDefault();
            if (xfrm != null && xfrm.Rotation != null)
            {
                rotation = (int)xfrm.Rotation.Value; // in 1/60000 degree
            }

            return (imageBytes, rotation);
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

    public class RowImage
    {
        public TableRow Row { get; set; }
        public byte[] ImageData { get; set; }
    }
}
