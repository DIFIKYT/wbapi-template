using Newtonsoft.Json;

namespace APIWildberries
{
    public class CampaignResponse
    {
        [JsonProperty("days")] private List<Day>? _days;

        public List<Day>? Days => _days;
    }

    public class Day
    {
        [JsonProperty("apps")] private List<App>? _apps;

        public List<App>? Apps => _apps;
    }

    public class App
    {
        [JsonProperty("nm")] private List<Nm>? _nm;

        public List<Nm>? Nm => _nm;
    }

    public class Nm
    {
        [JsonProperty("views")] public int Views { get; set; }
        [JsonProperty("clicks")] public int Clicks { get; set; }
        [JsonProperty("ctr")] public double Ctr { get; set; }
        [JsonProperty("sum")] public double Expenses { get; set; }
        [JsonProperty("nmId")] public int Articule { get; set; }
    }
}