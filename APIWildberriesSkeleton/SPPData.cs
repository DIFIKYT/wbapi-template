namespace APIWildberries
{
    public class SPPData(int spp, double priceBeforeSpp)
    {
        public int SPP { get; private set; } = spp;
        public double PriceBeforeSpp { get; private set; } = priceBeforeSpp;
    }
}