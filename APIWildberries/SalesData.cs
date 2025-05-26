using Newtonsoft.Json;

namespace APIWildberries
{
    public class SalesData : IData
    {
        [JsonProperty("supplierArticle")] public string? VendoreCode { get; set; }
        [JsonProperty("barcode")] public string? Barcode { get; set; }
        [JsonProperty("techSize")] public string? TechSize { get; set; }
        public int Amount { get; set; } = 1;
    }
}
