using Newtonsoft.Json;

namespace APIWildberries
{
    public class CompanyData
    {
        [JsonProperty("adverts")] private List<Adverts>? _adverts;

        public List<Adverts>? Adverts => _adverts;
    }

    public class Adverts
    {
        [JsonProperty("advert_list")] private List<AdvertList>? _advertLists;

        public List<AdvertList>? AdvertLists => _advertLists;
    }

    public class AdvertList
    {
        [JsonProperty("advertId")] public int AdvertId { get; set; }
    }
}