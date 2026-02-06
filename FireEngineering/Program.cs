namespace FireEngineering
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<ProtocolDataModel> rows = new List<ProtocolDataModel>();
            for (int i = 0; i < 1; i++)
            {
                rows.AddRange(DataReader.Read());
            }
            WordTableWriter.WriteCombinedDoc(@"D:\FireEngineering\Test123.docx", rows);
        }
    }
}
