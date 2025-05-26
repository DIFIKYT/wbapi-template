namespace APIWildberries
{
    public class ProductData(int viewsCount, int clicksCount, double clickConversion, double sum)
    {
        public int ViewsCount { get; set; } = viewsCount;
        public int ClicksCount { get; set; } = clicksCount;
        public double ClickConversion { get; set; } = clickConversion;
        public double Expenses { get; set; } = sum;
    }
}