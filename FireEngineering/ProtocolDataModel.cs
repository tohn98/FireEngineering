namespace FireEngineering
{
    public class ProtocolDataModel
    {
        public string BrandAndModel {  get; set; }
        public string Serial { get; set; }
        public string Category { get; set; }
        public double Capacity { get; set; }
        public string Type { get; set; }
        public string ChangeType { get; set; }
        public string Changed { get; set; }
        public string ChangeDate { get; set; }
        public string WorkerName { get; set; }
        public (byte[] data, int rotation) Signature { get; set; }
        public int StickerId { get; set; }

        public string FutureChangeDate
        {
            get
            {
                if (DateTime.TryParse(ChangeDate, out var dt))
                    return dt.AddYears(1).ToString("dd-MM-yyyy");
                return "";
            }
        }
    }
}
