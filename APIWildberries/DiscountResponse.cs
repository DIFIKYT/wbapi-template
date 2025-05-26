using Newtonsoft.Json;

namespace APIWildberries
{
    public class DiscountResponse
    {
        [JsonProperty("data")] public DiscountData? Data { get; set; }
    }

    public class DiscountData
    {
        [JsonProperty("listGoods")] private List<Good>? _goods;
        public List<Good>? Goods => _goods;
    }

    public class Good
    {
        [JsonProperty("nmID")] public int Articule { get; private set; }
        [JsonProperty("discount")] public int Discount;
        [JsonProperty("sizes")] private List<Size>? _sizes;
        public List<Size>? Sizes => _sizes;
    }

    public class Size
    {
        [JsonProperty("discountedPrice")] public double DiscountedPrice { get; private set; }
    }
}