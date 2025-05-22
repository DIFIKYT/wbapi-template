namespace APIWildberries
{
    public class DataForMargin(double purchasePrice, double commissionPercent, double acquiringPercent, double fullfilmentProcessingRub, double logisticsRub, double keepingRub, double acceptanceRub)
    {
        public double PurchasePrice { get; private set; } = purchasePrice;
        public double CommissionPercent { get; private set; } = commissionPercent;
        public double AcquiringPercent { get; private set; } = acquiringPercent;
        public double FullfilmentProcessingRub { get; private set; } = fullfilmentProcessingRub;
        public double LogisticsRub { get; private set; } = logisticsRub;
        public double KeepingRub { get; private set; } = keepingRub;
        public double AcceptanceRub { get; private set; } = acceptanceRub;
    }
}
