using Newtonsoft.Json;

namespace APIWildberries
{
    public class Config
    {
        [JsonProperty("summarySheetId")] public string? SummarySheetId { get; private set; }
        [JsonProperty("googleFolderId")] public string? GoogleFolderId { get; private set; }

        [JsonProperty("credentialsPath")] public string? CredentialsPath { get; private set; }
        [JsonProperty("sheetInfoFilePath")] public string? SheetInfoFilePath { get; private set; }
        [JsonProperty("spreadsheetTitlesPath")] public string? SpreadsheetTitlesPath { get; private set; }
        [JsonProperty("warehouseNamesPath")] public string? WarehouseNamesPath { get; private set; }

        [JsonProperty("startWordSpreadsheetName")] public string? StartWordSpreadsheetName { get; private set; }

        [JsonProperty("salesDataListName")] public string? SalesDataListName { get; private set; }
        [JsonProperty("generalSummaryListName")] public string? GeneralSummaryListName { get; private set; }
        [JsonProperty("salesFunnelListName")] public string? SalesFunnelListName { get; private set; }
        [JsonProperty("unitListName")] public string? UnitListName { get; private set; }

        [JsonProperty("salesUrl")] public string? SalesUrl { get; private set; }
        [JsonProperty("cardStatUrl")] public string? CardStatUrl { get; private set; }
        [JsonProperty("companiesURL")] public string? CompaniesURL { get; private set; }
        [JsonProperty("fullCompanyStatURL")] public string? FullCompanyStatURL { get; private set; }
        [JsonProperty("discountDataUrl")] public string? DiscountDataUrl { get; private set; }
        [JsonProperty("feedbacksUrl")] public string? FeedbacksUrl { get; private set; }
        [JsonProperty("warehouseRemainsUrl")] public string? WarehouseRemainsUrl { get; private set; }

        [JsonProperty("warehouseRemainsCreateReportEnding")] public string? WarehouseRemainsCreateReportEnding { get; private set; }
        [JsonProperty("warehouseRemainsCheckStatusEnding")] public string? WarehouseRemainsCheckStatusEnding { get; private set; }
        [JsonProperty("warehouseRemainsDownloadEnding")] public string? WarehouseRemainsDownloadEnding { get; private set; }
    }
}
