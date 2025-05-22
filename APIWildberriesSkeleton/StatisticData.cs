using Newtonsoft.Json;

namespace APIWildberries
{
    public class StatisticData
    {
        [JsonProperty("data")] public Data? Data { get; private set; }
    }

    public class Data
    {
        [JsonProperty("cards")] private List<Card>? _cards;
        public List<Card>? Cards => _cards;
    }

    public class Card
    {
        [JsonProperty("nmID")] public int Articule { get; private set; }
        [JsonProperty("statistics")] public Statistic? Statistics { get; private set; }
    }

    public class Statistic
    {
        [JsonProperty("selectedPeriod")] public SelectedPeriod? SelectedPeriod { get; private set; }
    }

    public class SelectedPeriod
    {
        [JsonProperty("openCardCount")] public int OpenCardCount { get; private set; }
        [JsonProperty("addToCartCount")] public int AddToCartCount { get; private set; }
        [JsonProperty("ordersCount")] public int OrdersCount { get; private set; }
        [JsonProperty("ordersSumRub")] public int OrdersSumRub { get; private set; }
        [JsonProperty("conversions")] public Conversion? Conversions { get; private set; }
    }

    public class Conversion
    {
        [JsonProperty("addToCartPercent")] public double AddToCartPercent { get; private set; }
        [JsonProperty("cartToOrderPercent")] public double CartToOrderPercent { get; private set; }
    }
}