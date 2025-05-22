using APIWildberries;
using Newtonsoft.Json;

class Program
{
    static readonly ManualResetEvent shutdownEvent = new(false);

    static void Main()
    {
        try
        {
            Config? config = JsonConvert.DeserializeObject<Config>(File.ReadAllText("Config.json"));

            if (config == null)
            {
                Console.WriteLine("Конфиг не получен");
                return;
            }

            HttpClient client = new()
            {
                Timeout = new TimeSpan(0, 2, 0)
            };


            DailyMidnightTaskScheduler dailyMidnightTaskScheduler = new();
            WBApiHelper wBApiHelper = new(
                client: client,
                salesUrl: config.SalesUrl!,
                cardStatUrl: config.CardStatUrl!,
                companiesURL: config.CompaniesURL!,
                fullCompanyStatURL: config.FullCompanyStatURL!,
                discountDataUrl: config.DiscountDataUrl!,
                feedbacksUrl: config.FeedbacksUrl!,
                warehouseRemainsUrl: config.WarehouseRemainsUrl!,
                warehouseRemainsCreateReportEnding: config.WarehouseRemainsCreateReportEnding!,
                warehouseRemainsCheckStatusEnding: config.WarehouseRemainsCheckStatusEnding!,
                warehouseRemainsDownloadEnding: config.WarehouseRemainsDownloadEnding!
            );
            GoogleSheetsHelper sheetsHelper = new(
                wBApiHelper: wBApiHelper,
                dailyMidnightTaskScheduler: dailyMidnightTaskScheduler,
                summurySpreadsheetId: config.SummarySheetId!,
                googleFolderId: config.GoogleFolderId!,
                credentialsPath: config.CredentialsPath!,
                sheetInfoFilePath: config.SheetInfoFilePath!,
                spreadsheetTitlesPath: config.SpreadsheetTitlesPath!,
                warehouseNamesPath: config.WarehouseNamesPath!,
                startWordSpreadsheetName: config.StartWordSpreadsheetName!,
                dailyDataListName: config.SalesDataListName!,
                generalSummaryListName: config.GeneralSummaryListName!,
                salesFunnelListName: config.SalesFunnelListName!,
                unitListName: config.UnitListName!
            );

            Console.CancelKeyPress += (sender, e) =>
            {
                Console.WriteLine("Shutdown...");
                shutdownEvent.Set();
                e.Cancel = true;
            };

            Console.WriteLine("Press any key for stop...");

            shutdownEvent.WaitOne();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Произошла ошибка: {ex.Message}");
        }
    }
}