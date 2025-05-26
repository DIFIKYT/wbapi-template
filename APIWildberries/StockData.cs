using Newtonsoft.Json;

namespace APIWildberries
{
    public class StockData : IData
    {
        [JsonProperty("vendorCode")] public string? VendoreCode { get; set; }
        [JsonProperty("nmId")] public int Articule { get; set; }
        [JsonProperty("techSize")] public string? TechSize { get; set; }
        [JsonProperty("barcode")] public string? Barcode { get; set; }
        [JsonProperty("warehouses")] private readonly List<Warehouse>? _warehouses;
        public List<Warehouse>? WareHouses => _warehouses;
    }

    public class Warehouse
    {
        [JsonProperty("warehouseName")] public string? Name { get; set; }
        [JsonProperty("quantity")] public int Quantity { get; set; }
    }
}