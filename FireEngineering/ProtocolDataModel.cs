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
                {
                    int years = Changed switch
                    {
                        "Техническо обслужване" => 1,
                        _ when Changed?.StartsWith("Хидростатично", StringComparison.OrdinalIgnoreCase) == true => 10,
                        _ when Changed?.StartsWith("хидравлично", StringComparison.OrdinalIgnoreCase) == true => 10,
                        _ when Changed?.StartsWith("Презареждане", StringComparison.OrdinalIgnoreCase) == true => 5,
                        _ => 1
                    };

                    return dt.AddYears(years).ToString("dd-MM-yyyy");
                }

                return "";
            }
        }
    }
}
