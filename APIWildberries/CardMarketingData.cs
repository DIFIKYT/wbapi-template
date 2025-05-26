namespace APIWildberries
{
    class CardMarketingData(
        string imageUrl,
        double raiting, double margin,
        double ordersSumRub,
        double priceBeforeSPP, int spp,
        int openCardCount, int addToCartCount, int ordersCount,
        int viewsCount, int clicksCount, double expenses,
        double clickConversion, double addToCartPercent, double cartToOrderPercent,
        int turnoverInDays,
        int centerStockSum, int piterStockSum, int kazanStockSum, int ugStockSum, int uralStockSum, int sibirStockSum,
        int toClient, int fromClient, int totalStocksSum)
    {

        public string ImageUrl { get; private set; } = imageUrl;

        public double Raiting { get; private set; } = raiting;
        public double Margin { get; private set; } = margin;

        public double OrdersSumRub { get; private set; } = ordersSumRub;

        public double PriceBeforeSPP { get; private set; } = priceBeforeSPP;
        public int SPP { get; private set; } = spp;

        public int OpenCardCount { get; private set; } = openCardCount;
        public int AddToCartCount { get; private set; } = addToCartCount;
        public int OrdersCount { get; private set; } = ordersCount;

        public int ViewsCount { get; private set; } = viewsCount;
        public int ClicksCount { get; private set; } = clicksCount;
        public double Expenses { get; private set; } = expenses;

        public double ClickConversion { get; private set; } = clickConversion;
        public double AddToCartPercent { get; private set; } = addToCartPercent;
        public double CartToOrderPercent { get; private set; } = cartToOrderPercent;

        public int TurnoverInDays { get; private set; } = turnoverInDays;

        public int CenterStockSum { get; private set; } = centerStockSum;
        public int PiterStockSum { get; private set; } = piterStockSum;
        public int KazanStockSum { get; private set; } = kazanStockSum;
        public int UgStockSum { get; private set; } = ugStockSum;
        public int UralStockSum { get; private set; } = uralStockSum;
        public int SibirStockSum { get; private set; } = sibirStockSum;
        public int ToClient { get; private set; } = toClient;
        public int FromClient { get; private set; } = fromClient;
        public int TotalStocksSum { get; private set; } = totalStocksSum;
    }
}