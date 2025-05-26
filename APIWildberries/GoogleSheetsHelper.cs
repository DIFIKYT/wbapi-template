using static Google.Apis.Sheets.v4.SpreadsheetsResource.ValuesResource;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Data;
using Google;
using System.Net;
using System.Text;
using Google.Apis.Drive.v3;

namespace APIWildberries
{
    public class GoogleSheetsHelper : DataHelper
    {
        private const int TaxPercentage = 6;
        private const int ThreeDays = 3;
        private const int SevenDays = 7;
        private const int TweentyOneDays = 21;
        private const int FirstDataRowTurnover = 6;
        private const int FirstDataRowSalesFunnel = 35;
        private const int BaseCountAddIteractions = 14;
        private const int ColumnsCountForNewSheet = 40;
        private const int ArticuleImageCellsCount = 10;
        private const int GrayRowSize = 36;
        private const int DistanceBetweenData = 2;
        private const string ArticuleColumn = "B";
        private const string BarcodeColumn = "C";
        private const string SizesColumn = "D";
        private const string RemainsWBColumn = "G";
        private const string StartColumnForAverageData = "K";
        private const string EndColumnForAverageData = "M";
        private const string ThreeDaysColumn = "AP";
        private const string SevenDaysColumn = "AQ";
        private const string TweentyOneDaysColumn = "AR";
        private const string ApiKeyShellAdress = "A1";
        private const string StartColumnForMainData = "A";
        private const string EndColumnForMainData = "J";
        private const string FirstDataColumnSalesFunnel = "A";
        private const string PeriodPlanColumnSalesFunnel = "B";
        private const string FactPeriodColumnSalesFunnel = "C";
        private const string IndicatorsColumnSalesFunnel = "D";
        private const string NormPerDayColumnSalesFunnel = "E";

        private static readonly string[] _sheetsScopes = [SheetsService.Scope.Spreadsheets];
        private static readonly string[] _driveScopes = [DriveService.Scope.DriveReadonly];

        private readonly Color _grayColor = new()
        {
            Red = 0.8f,
            Green = 0.8f,
            Blue = 0.8f,
            Alpha = 1.0f
        };

        private readonly Color _greenColor = new()
        {
            Red = 0.8f,
            Green = 0.9f,
            Blue = 0.8f,
            Alpha = 1.0f
        };

        private readonly string _googleFolderId;
        private readonly SheetsService _sheetsService;
        private readonly DriveService _driveService;
        private readonly WBApiHelper _wBApiHelper;
        private readonly DailyMidnightTaskScheduler _dailyMidnightTaskScheduler;
        private readonly Dictionary<int, string> _columnsAndDates = [];
        private readonly List<string> _columnsForSales = ["K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO"];
        private readonly List<string> _columnsForLastMonthSales = ["AW", "AX", "AY", "AZ", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "CA", "CB"];
        private readonly List<string> _columnsForSalesFunnelData = ["F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ"];
        private readonly string _unitListName;
        private readonly string _startWordSpreadsheetName;
        private readonly string _sheetInfoFilePath;
        private readonly string _spreadsheetTitlesPath;
        private readonly string _warehouseNamesPath;
        private readonly string _defaultRangeLastRow;
        private readonly string _turnoverListName;
        private readonly string _generalSummaryListName;
        private readonly string _RNPListName;
        private readonly string? _summarySpreadsheetId;
        private readonly string RNPCellAdressForCheckEmpty;
        private readonly CultureInfo _cultureInfo = new("ru-RU");
        private Spreadsheet _currentSpreadsheet;
        private int _startValueLastTwoWeeksData;
        private string? _currentSpreadsheetId;
        private int? _turnoverSheetid;
        private int? _RNPSheetId;
        private bool _isRNPEmpty;

        public GoogleSheetsHelper(WBApiHelper wBApiHelper, DailyMidnightTaskScheduler dailyMidnightTaskScheduler,
            string summurySpreadsheetId, string googleFolderId,
            string credentialsPath, string sheetInfoFilePath, string spreadsheetTitlesPath, string warehouseNamesPath,
            string startWordSpreadsheetName,
            string dailyDataListName, string generalSummaryListName, string salesFunnelListName, string unitListName)
        {
            GoogleCredential spreadsheetsCredential = GoogleCredential.FromFile(credentialsPath).CreateScoped(_sheetsScopes);
            GoogleCredential driveCredential = GoogleCredential.FromFile(credentialsPath).CreateScoped(_driveScopes);

            _sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = spreadsheetsCredential,
                ApplicationName = "WBApi",
            });
            _driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = driveCredential,
                ApplicationName = "WBApi",
            });

            _wBApiHelper = wBApiHelper;
            _dailyMidnightTaskScheduler = dailyMidnightTaskScheduler;

            _summarySpreadsheetId = summurySpreadsheetId;
            _googleFolderId = googleFolderId;

            _sheetInfoFilePath = sheetInfoFilePath;
            _spreadsheetTitlesPath = spreadsheetTitlesPath;
            _warehouseNamesPath = warehouseNamesPath;

            _startWordSpreadsheetName = startWordSpreadsheetName;

            _turnoverListName = dailyDataListName;
            _generalSummaryListName = generalSummaryListName;
            _RNPListName = salesFunnelListName;
            _unitListName = unitListName;

            _defaultRangeLastRow = $"{_turnoverListName}!{ArticuleColumn}:{ArticuleColumn}";
            RNPCellAdressForCheckEmpty = $"{_RNPListName}!D2";

            _dailyMidnightTaskScheduler.TaskExecuted += OnTaskExecuted;

            for (int i = 0; i < _columnsForSales.Count; i++)
            {
                _columnsAndDates.Add(i, _columnsForSales[i]);
            }
        }

        //Subscribe methods

        private async Task OnTaskExecuted()
        {
            await StartWork(-1);
        }

        //Main methods
        public async Task StartWork(int minusDaysCount)
        {
            Dictionary<string, string> dataToUpdate = [];
            Dictionary<string, SalesData> salesData = [];
            Dictionary<string, StockData> stocksData = [];
            Dictionary<int, Good> articulesToGoods = [];
            Dictionary<string, Card> urlToCardData = [];
            Dictionary<int, ProductData> articulesToMarketingData = [];
            DateTime todayDate = DateTime.Now.AddDays(minusDaysCount).Date;

            _startValueLastTwoWeeksData = todayDate.Day < BaseCountAddIteractions ? todayDate.Day - 1 : BaseCountAddIteractions - 1;

            string json = File.ReadAllText(_spreadsheetTitlesPath);
            JObject data = JObject.Parse(json);
            JArray titles = (JArray)data["titles"]!;

            foreach (string spreadSheetId in GetSpreadshetIds())
            {
                salesData.Clear();
                stocksData.Clear();
                dataToUpdate.Clear();
                articulesToGoods.Clear();
                urlToCardData.Clear();
                articulesToMarketingData.Clear();

                _currentSpreadsheetId = spreadSheetId;
                _currentSpreadsheet = await _sheetsService.Spreadsheets.Get(_currentSpreadsheetId).ExecuteAsync();
                _turnoverSheetid = GetSheetId(_currentSpreadsheet, _turnoverListName);
                _RNPSheetId = GetSheetId(_currentSpreadsheet, _RNPListName);
                string apiKey = GetApiKeyFromList();

                if (string.IsNullOrEmpty(apiKey))
                {
                    continue;
                }

                titles.Add(_currentSpreadsheet.Properties.Title);

                File.WriteAllText(_spreadsheetTitlesPath, data.ToString(Formatting.Indented));

                await _wBApiHelper.GetData(salesData, stocksData, apiKey, todayDate);

                _isRNPEmpty = GetValueByShellsRange(_currentSpreadsheetId!, RNPCellAdressForCheckEmpty).Values == null;

                if (GetLastRow(_currentSpreadsheetId, _defaultRangeLastRow) == FirstDataRowTurnover)
                {
                    await PrepareTurnoverSheet(salesData, stocksData, dataToUpdate, _startValueLastTwoWeeksData, todayDate);
                    await Task.Delay(1000);

                    if (_isRNPEmpty)
                    {
                        await PrepareRNPSheet(articulesToGoods, urlToCardData, articulesToMarketingData, stocksData, dataToUpdate, _RNPSheetId, todayDate, _startValueLastTwoWeeksData);
                    }

                    continue;
                }

                if (todayDate.Day == 1)
                {
                    await CopyAndClearData();
                }

                await UpdateTodayData(salesData, stocksData, dataToUpdate, urlToCardData, articulesToMarketingData, articulesToGoods, todayDate);
            }

            await Task.Delay(10000);
            await MergeSummarySheets(titles);

            titles.Clear();
            File.WriteAllText(_spreadsheetTitlesPath, data.ToString(Formatting.Indented));
        }

        private async Task MergeSummarySheets(JArray titles)
        {
            try
            {
                await ClearData(_summarySpreadsheetId!, $"{_generalSummaryListName}!A4:{GetColumnName(GetColumnIndex(EndColumnForAverageData) + 1)}{GetLastRow(_summarySpreadsheetId!, $"{_generalSummaryListName}!{ArticuleColumn}:{ArticuleColumn}")}");
                Spreadsheet spreadsheet = await _sheetsService.Spreadsheets.Get(_summarySpreadsheetId).ExecuteAsync();

                int currentRow = 4;
                var batchUpdateRequests = new List<ValueRange>();

                foreach (Sheet sheet in spreadsheet.Sheets)
                {
                    string sheetTitle = sheet.Properties.Title ?? string.Empty;

                    bool isMatch = titles.Any(title =>
                    {
                        string titleStr = title.ToString();
                        bool match = string.Equals(sheetTitle.Trim(), titleStr.Trim(), StringComparison.OrdinalIgnoreCase);
                        return match;
                    });

                    if (isMatch == false)
                    {
                        continue;
                    }

                    ValueRange valueRange = GetValueByShellsRange(_summarySpreadsheetId!, $"{sheet.Properties.Title}!A3:{EndColumnForAverageData}{GetLastRow(_summarySpreadsheetId!, $"{sheet.Properties.Title}!{ArticuleColumn}:{ArticuleColumn}")}");

                    string rangeToUpdate = $"{_generalSummaryListName}!A{currentRow}:{GetColumnName(GetColumnIndex(EndColumnForAverageData) + 1)}{currentRow + valueRange.Values.Count - 1}";

                    for (int i = 0; i < valueRange.Values.Count; i++)
                        valueRange.Values[i].Insert(0, sheet.Properties.Title);

                    batchUpdateRequests.Add(new ValueRange
                    {
                        Range = rangeToUpdate,
                        Values = valueRange.Values
                    });

                    currentRow += valueRange.Values.Count + 2;
                    await Task.Delay(5000);
                }

                if (batchUpdateRequests.Count != 0)
                {
                    BatchUpdateValuesRequest batchRequest = new()
                    {
                        ValueInputOption = "USER_ENTERED",
                        Data = batchUpdateRequests
                    };

                    await _sheetsService.Spreadsheets.Values.BatchUpdate(batchRequest, _summarySpreadsheetId).ExecuteAsync();
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Ошибка в общем заполнении сводной таблицы:\n{exception}");
            }
        }

        private async Task UpdateTodayData(Dictionary<string, SalesData> salesData, Dictionary<string, StockData> stocksData, Dictionary<string, string> dataToUpdate, Dictionary<string, Card> urlToCardData, Dictionary<int, ProductData> articulesToMarketingData, Dictionary<int, Good> articulesToGoods, DateTime todayDate)
        {
            if (TryGetList() == false)
            {
                Console.WriteLine("List not found");
                return;
            }

            try
            {
                await AddTodaySalesAndStocksData(salesData, stocksData, dataToUpdate, todayDate);
                await AddSummaryData();

                dataToUpdate.Clear();

                if (_isRNPEmpty)
                {
                    await PrepareRNPSheet(articulesToGoods, urlToCardData, articulesToMarketingData, stocksData, dataToUpdate, _RNPSheetId, todayDate, _startValueLastTwoWeeksData);
                }
                else
                {
                    await AddTodayRNPData(urlToCardData, articulesToMarketingData, articulesToGoods, stocksData, dataToUpdate, todayDate);
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Ошибка ежедневном заполнении {_currentSpreadsheet.Properties.Title} | {_currentSpreadsheetId}:\n{exception}");
            }
        }

        //Add methods

        private async Task AddTodaySalesAndStocksData(Dictionary<string, SalesData> salesData, Dictionary<string, StockData> stocksData, Dictionary<string, string> dataToUpdate, DateTime todayDate)
        {
            AddStocksData(dataToUpdate, stocksData);
            await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
            await Task.Delay(1000);

            dataToUpdate.Clear();
            AddSalesData(salesData, dataToUpdate, todayDate.Day);
            await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
            await Task.Delay(1000);

            dataToUpdate.Clear();
            FillEmptyDataRows(dataToUpdate, todayDate);
            await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
            await Task.Delay(1000);

            dataToUpdate.Clear();
            await AddAverageData(dataToUpdate, todayDate);
            await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
        }

        private async Task AddTodayRNPData(Dictionary<string, Card> urlToCardData, Dictionary<int, ProductData> articulesToMarketingData, Dictionary<int, Good> articulesToGoods, Dictionary<string, StockData> stocksData, Dictionary<string, string> dataToUpdate, DateTime todayDate)
        {
            try
            {
                string sourceSpreadsheetName = _currentSpreadsheet.Properties.Title;

                CompanyData? companyData = await _wBApiHelper.GetCompanyData();
                List<int> companyIds = [];
                List<Request> requests = [];
                List<string> notFoundArticules = [];
                Dictionary<int, string> rowsFoundArticules = [];
                Dictionary<string, CardMarketingData> articuleToCardMarketingData = [];
                Dictionary<int, int> articulesToTurnoverInDays = [];
                Dictionary<int, SPPData> articuleToSPP = [];
                Dictionary<int, double> articuleToPrice = [];
                Dictionary<int, DataForMargin> articuleToDataForMargin = [];

                ProductData productData;
                Good good;

                bool isMonthDataAverageAdded = false;
                bool isWeekDataAverageAdded = false;
                int lastFreeRow = GetLastRow(_currentSpreadsheetId!, $"{_RNPListName}!{IndicatorsColumnSalesFunnel}:{IndicatorsColumnSalesFunnel}");
                int articulesInfoRowNumber = lastFreeRow;
                int dataRowNumber = articulesInfoRowNumber + 2;
                string columnName = GetColumnName(GetMaxDataCount(await _sheetsService.Spreadsheets.Values.Get(_currentSpreadsheetId, _RNPListName).ExecuteAsync(), GetColumnIndex(_columnsForSalesFunnelData[0])));

                int totalStocksSum = 0;
                int centerStockSum = 0;
                int piterStockSum = 0;
                int kazanStockSum = 0;
                int ugStockSum = 0;
                int uralStockSum = 0;
                int sibirStockSum = 0;
                int toClient = 0;
                int fromClient = 0;

                int currentColumnIndex = GetColumnIndex(columnName);
                int baseColumnIndex = GetColumnIndex(_columnsForSalesFunnelData[0]);

                if (currentColumnIndex <= baseColumnIndex - 1)
                {
                    columnName = _columnsForSalesFunnelData[todayDate.Day - 1];
                }
                else
                {
                    columnName = GetColumnName(currentColumnIndex + 1);
                }

                foreach (Adverts adverts in companyData!.Adverts!)
                {
                    foreach (AdvertList list in adverts.AdvertLists!)
                    {
                        companyIds.Add(list.AdvertId);
                    }
                }

                Task discountTask = _wBApiHelper.GetDiscountData(articulesToGoods);
                Task statTask = _wBApiHelper.GetStatisticData(urlToCardData, todayDate);
                Task fullStatTask = _wBApiHelper.GetFullStat(articulesToMarketingData, companyIds, todayDate);
                Task turnoverTask = _wBApiHelper.GetTurnoverInDays(articulesToTurnoverInDays, todayDate);

                await Task.WhenAll(discountTask, statTask, fullStatTask, turnoverTask);

                articuleToPrice = articulesToGoods.Values.ToDictionary(good => good.Articule, good => good.Sizes![0].DiscountedPrice);
                GetDataForMargin(articuleToDataForMargin);

                await _wBApiHelper.GetSPP(articuleToSPP, articuleToPrice);

                string json = File.ReadAllText(_warehouseNamesPath, Encoding.UTF8);

                ILookup<string, string> warehouseRegions = JObject.Parse(json)
                    .Properties()
                    .SelectMany(p => p.Value.ToObject<string[]>()!
                        .Select(wh => new { Warehouse = wh, Region = p.Name }))
                    .ToLookup(x => x.Warehouse, x => x.Region);

                foreach (string url in urlToCardData.Keys)
                {
                    try
                    {
                        totalStocksSum = 0;
                        centerStockSum = 0;
                        piterStockSum = 0;
                        kazanStockSum = 0;
                        ugStockSum = 0;
                        uralStockSum = 0;
                        sibirStockSum = 0;
                        toClient = 0;
                        fromClient = 0;

                        int articuleFromCard = urlToCardData[url].Articule;

                        if (articulesToMarketingData.TryGetValue(articuleFromCard, out productData!) == false)
                        {
                            productData = new(0, 0, 0, 0);
                        }

                        if (articulesToGoods.TryGetValue(articuleFromCard, out good!) == false)
                        {
                            good = new();
                        }

                        Dictionary<string, int> sums = stocksData.Values
                            .Where(stock => stock.Articule == articuleFromCard)
                            .SelectMany(stock => stock.WareHouses!)
                            .GroupBy(wh =>
                            {
                                IEnumerable<string> regions = warehouseRegions[wh.Name!];
                                if (regions.Any())
                                {
                                    return regions.First();
                                }

                                return wh.Name switch
                                {
                                    "В пути до получателей" => "toClient",
                                    "В пути возвраты на склад WB" => "fromClient",
                                    "Всего находится на складах" => "totalStocks",
                                    _ => "other"
                                };
                            })
                            .ToDictionary(
                                g => g.Key,
                                g => g.Sum(wh => wh.Quantity)
                            );

                        centerStockSum = sums.GetValueOrDefault("центр", 0);
                        piterStockSum = sums.GetValueOrDefault("питер", 0);
                        kazanStockSum = sums.GetValueOrDefault("казань", 0);
                        ugStockSum = sums.GetValueOrDefault("юг", 0);
                        uralStockSum = sums.GetValueOrDefault("урал", 0);
                        sibirStockSum = sums.GetValueOrDefault("сибирь", 0);
                        toClient = sums.GetValueOrDefault("toClient", 0);
                        fromClient = sums.GetValueOrDefault("fromClient", 0);
                        totalStocksSum = sums.GetValueOrDefault("totalStocks", 0);

                        articuleToCardMarketingData[articuleFromCard.ToString()] = new CardMarketingData(
                            url + "images/big/1.webp",
                            await _wBApiHelper.GetRating(articuleFromCard.ToString(), todayDate, _currentSpreadsheetId!),
                            articuleToDataForMargin.TryGetValue(articuleFromCard, out DataForMargin? dataForMargin) ? CalculateMargin(dataForMargin, urlToCardData[url].Statistics!.SelectedPeriod!.OrdersCount, urlToCardData[url].Statistics!.SelectedPeriod!.OrdersSumRub, productData.Expenses, TaxPercentage)! : 0,
                            urlToCardData[url].Statistics!.SelectedPeriod!.OrdersSumRub,
                            articuleToSPP[articuleFromCard].PriceBeforeSpp,
                            articuleToSPP[articuleFromCard].SPP,
                            urlToCardData[url].Statistics!.SelectedPeriod!.OpenCardCount,
                            urlToCardData[url].Statistics!.SelectedPeriod!.AddToCartCount,
                            urlToCardData[url].Statistics!.SelectedPeriod!.OrdersCount,
                            productData!.ViewsCount,
                            productData.ClicksCount,
                            productData.Expenses,
                            productData.ClickConversion,
                            urlToCardData[url].Statistics!.SelectedPeriod!.Conversions!.AddToCartPercent,
                            urlToCardData[url].Statistics!.SelectedPeriod!.Conversions!.CartToOrderPercent,
                            articulesToTurnoverInDays.TryGetValue(articuleFromCard, out int turnoverCount) ? turnoverCount : 0,
                            centerStockSum,
                            piterStockSum,
                            kazanStockSum,
                            ugStockSum,
                            uralStockSum,
                            sibirStockSum,
                            toClient,
                            fromClient,
                            totalStocksSum
                            );
                    }
                    catch { }

                    await Task.Delay(500);
                }

                string lastDayColumn = GetColumnName(GetColumnIndex(columnName) - 1);

                await TakeMarketingDataRows(articuleToCardMarketingData, rowsFoundArticules, notFoundArticules);
                ValueRange lastDayColumnValueRange = GetValueByShellsRange(_currentSpreadsheetId!, $"{_RNPListName}!{lastDayColumn}:{lastDayColumn}");

                foreach (string articule in notFoundArticules)
                {
                    AddNewMarketingCardData(articule, columnName, articuleToCardMarketingData[articule], dataToUpdate, lastDayColumnValueRange, sourceSpreadsheetName, requests, _RNPSheetId, ref articulesInfoRowNumber, ref dataRowNumber);
                }

                foreach (int row in rowsFoundArticules.Keys)
                {
                    AddMarketingCardData(row + 1, columnName, articuleToCardMarketingData[rowsFoundArticules[row]], dataToUpdate, lastDayColumnValueRange);
                }

                await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);

                dataToUpdate.Clear();

                AddTotalArticulesDataRNPData(dataToUpdate, articulesInfoRowNumber, columnName);

                await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);

                dataToUpdate.Clear();

                if (todayDate.AddDays(1).Day == 1)
                {
                    await AddTotalArticuleDataRNPData(dataToUpdate, requests, articulesInfoRowNumber, columnName, todayDate.Day, todayDate);
                    isMonthDataAverageAdded = true;
                }

                if ((int)todayDate.DayOfWeek == 0)
                {
                    if (isMonthDataAverageAdded)
                    {
                        await AddTotalArticuleDataRNPData(dataToUpdate, requests, articulesInfoRowNumber, GetColumnName(GetColumnIndex(columnName) + 1), 7, todayDate);
                    }
                    else
                    {
                        await AddTotalArticuleDataRNPData(dataToUpdate, requests, articulesInfoRowNumber, columnName, 7, todayDate);
                    }

                    isWeekDataAverageAdded = true;
                }

                if (isMonthDataAverageAdded)
                {
                    await AddNewDates(GetColumnName(GetColumnIndex(columnName) + 1), dataToUpdate, todayDate.AddDays(1), _RNPSheetId, lastFreeRow - 1, requests);
                }
                else if (isMonthDataAverageAdded && isWeekDataAverageAdded)
                {
                    await AddNewDates(GetColumnName(GetColumnIndex(columnName) + 2), dataToUpdate, todayDate.AddDays(1), _RNPSheetId, lastFreeRow - 1, requests);
                }

                await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
                dataToUpdate.Clear();

                if (requests.Count > 0)
                {
                    await BatchUpdateCells(requests, _currentSpreadsheetId!);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка в обновлении ежедневной РНП - {_currentSpreadsheetId}\n{ex}");
            }
        }

        private async Task AddTotalArticuleDataRNPData(Dictionary<string, string> dataToUpdate, List<Request> requests, int lastDataRow, string columnName, int countColumnsData, DateTime todayDate)
        {
            double marginProfit;
            double ordersInRubs;
            double expenses;
            string marginCellAdress;

            int columnIndexForData = GetColumnIndex(columnName) - (countColumnsData - 1);
            string columnNameForTotalData = GetColumnName(columnIndexForData);

            if (columnIndexForData <= 2)
            {
                return;
            }

            int rowNumber = 3;
            string startColumn = columnNameForTotalData;
            string endColumn = GetColumnName(GetColumnIndex(columnName));

            string json = File.ReadAllText(_sheetInfoFilePath);
            JObject data = JObject.Parse(json);
            string sheetInfoListName = _currentSpreadsheetId!.ToString();

            if (data[sheetInfoListName] == null || data[sheetInfoListName]!.Type != JTokenType.Array)
            {
                data[sheetInfoListName] = new JArray();
            }

            JArray columns = (JArray)data[sheetInfoListName]!;
            List<ValueRange> weaklyColumns = [];

            if (countColumnsData == 7 && todayDate.AddDays(-countColumnsData).Month == todayDate.Month)
            {
                columns.Add(columnNameForTotalData);
                File.WriteAllText(_sheetInfoFilePath, data.ToString(Formatting.Indented));
            }
            else
            {
                columnNameForTotalData = GetColumnName(GetColumnIndex(columnNameForTotalData) - columns.Count);
                startColumn = GetColumnName(GetColumnIndex(startColumn) - columns.Count);

                foreach (JToken column in columns)
                {
                    ValueRange tempoRange = GetValueByShellsRange(_currentSpreadsheetId!, $"{_RNPListName}!{column}:{column}");
                    weaklyColumns.Add(tempoRange);
                    await Task.Delay(1000);
                }

                columns.Clear();
                File.WriteAllText(_sheetInfoFilePath, data.ToString(Formatting.Indented));
            }


            ValueRange valueRangeLastData = GetValueByShellsRange(_currentSpreadsheetId!, $"{_RNPListName}!{startColumn}1:{endColumn}{lastDataRow}");
            IList<IList<object>> values = valueRangeLastData.Values;

            if (countColumnsData == 7 && GetColumnIndex(columnNameForTotalData) >= GetColumnIndex(_columnsForSalesFunnelData[0]))
            {
                await AddNewColumns(_currentSpreadsheetId!, _RNPSheetId, GetColumnIndex(columnNameForTotalData) - 1, GetColumnIndex(columnNameForTotalData));
                dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}1"] = $"{todayDate.AddDays(-(countColumnsData - 1)):dd.MM.yyyy} - {todayDate:dd.MM.yyyy}";
                requests.Add(GetColorCellRequest(_RNPSheetId, 1, lastDataRow, GetColumnIndex(columnNameForTotalData) - 1, GetColumnIndex(columnNameForTotalData), _greenColor));

                while (rowNumber < values.Count)
                {
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    marginProfit = GetSumData(values, rowNumber);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber++}"] = marginProfit.ToString("0.#", _cultureInfo);
                    marginCellAdress = $"{_RNPListName}!{columnNameForTotalData}{rowNumber++}";
                    rowNumber++;
                    ordersInRubs = GetSumData(values, rowNumber);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber++}"] = ordersInRubs.ToString("0.#", _cultureInfo);
                    dataToUpdate[marginCellAdress] = (ordersInRubs <= 0 ? 0 : marginProfit / ordersInRubs).ToString("0.#", _cultureInfo);
                    rowNumber++;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++, true).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++, true).ToString("0.#", _cultureInfo);
                    rowNumber++;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetSumData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetSumData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetSumData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    rowNumber++;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetSumData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetSumData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    expenses = GetSumData(values, rowNumber);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber++}"] = expenses.ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber++}"] = (ordersInRubs <= 0 ? 0 : expenses / ordersInRubs * PercentBase).ToString("0.#", _cultureInfo);
                    rowNumber++;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    rowNumber++;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    rowNumber += DistanceBetweenData;
                }
            }
            else if (countColumnsData > 7)
            {
                await AddNewColumns(_currentSpreadsheetId!, _RNPSheetId, GetColumnIndex(columnNameForTotalData) - 1, GetColumnIndex(columnNameForTotalData));
                dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}1"] = $"{todayDate.AddDays(-(countColumnsData - 1)):dd.MM.yyyy} - {todayDate:dd.MM.yyyy}";
                requests.Add(GetColorCellRequest(_RNPSheetId, 1, lastDataRow, GetColumnIndex(columnNameForTotalData) - 1, GetColumnIndex(columnNameForTotalData), _greenColor));

                while (rowNumber < values.Count)
                {
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    marginProfit = (GetSumData(values, rowNumber) - GetSumWeeklyData(rowNumber, weaklyColumns));
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber++}"] = marginProfit.ToString("0.#", _cultureInfo);
                    marginCellAdress = $"{_RNPListName}!{columnNameForTotalData}{rowNumber++}";
                    requests.Add(GetColorCellRequest(_RNPSheetId, rowNumber - 1, rowNumber++, GetColumnIndex(columnName), GetColumnIndex(columnName) + GrayRowSize, _grayColor));
                    ordersInRubs = (GetSumData(values, rowNumber) - GetSumWeeklyData(rowNumber, weaklyColumns));
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber++}"] = ordersInRubs.ToString("0.#", _cultureInfo);
                    dataToUpdate[marginCellAdress] = (ordersInRubs <= 0 ? 0 : marginProfit / ordersInRubs * PercentBase).ToString("0.#", _cultureInfo);
                    requests.Add(GetColorCellRequest(_RNPSheetId, rowNumber - 1, rowNumber++, GetColumnIndex(columnName), GetColumnIndex(columnName) + GrayRowSize, _grayColor));
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++, true).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++, true).ToString("0.#", _cultureInfo);
                    requests.Add(GetColorCellRequest(_RNPSheetId, rowNumber - 1, rowNumber++, GetColumnIndex(columnName), GetColumnIndex(columnName) + GrayRowSize, _grayColor));
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = (GetSumData(values, rowNumber) - GetSumWeeklyData(rowNumber++, weaklyColumns)).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = (GetSumData(values, rowNumber) - GetSumWeeklyData(rowNumber++, weaklyColumns)).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = (GetSumData(values, rowNumber) - GetSumWeeklyData(rowNumber++, weaklyColumns)).ToString("0.#", _cultureInfo);
                    requests.Add(GetColorCellRequest(_RNPSheetId, rowNumber - 1, rowNumber++, GetColumnIndex(columnName), GetColumnIndex(columnName) + GrayRowSize, _grayColor));
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = (GetSumData(values, rowNumber) - GetSumWeeklyData(rowNumber++, weaklyColumns)).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = (GetSumData(values, rowNumber) - GetSumWeeklyData(rowNumber++, weaklyColumns)).ToString("0.#", _cultureInfo);
                    expenses = (GetSumData(values, rowNumber) - GetSumWeeklyData(rowNumber, weaklyColumns));
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber++}"] = expenses.ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = (ordersInRubs <= 0 ? 0 : expenses / ordersInRubs * PercentBase).ToString("0.#", _cultureInfo);
                    requests.Add(GetColorCellRequest(_RNPSheetId, rowNumber - 1, rowNumber++, GetColumnIndex(columnName), GetColumnIndex(columnName) + GrayRowSize, _grayColor));
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetAverageData(values, rowNumber++).ToString("0.#", _cultureInfo);
                    requests.Add(GetColorCellRequest(_RNPSheetId, rowNumber - 1, rowNumber++, GetColumnIndex(columnName), GetColumnIndex(columnName) + GrayRowSize, _grayColor));
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    dataToUpdate[$"{_RNPListName}!{columnNameForTotalData}{rowNumber}"] = GetLastValue(values, rowNumber++).ToString()!;
                    rowNumber += DistanceBetweenData;
                }
            }
        }

        private void AddTotalArticulesDataRNPData(Dictionary<string, string> dataToUpdate, int lastDataRow, string columnName)
        {
            ValueRange columnValueRange = GetValueByShellsRange(_currentSpreadsheetId!, $"{_RNPListName}!{columnName}{FirstDataRowSalesFunnel}:{columnName}{lastDataRow}");
            IList<IList<object>> columnValues = columnValueRange.Values;

            int raitingCount = 0;
            double sumRaiting = 0;
            double marginProfit = 0;
            double ordersSumRubTotal = 0;
            double priceBeforeSPP = 0;
            double spp = 0;
            double openCardCountTotal = 0;
            double addToCartCountTotal = 0;
            double ordersCountTotal = 0;
            double viewsCountTotal = 0;
            double clicksCountTotal = 0;
            double expenseTotal = 0;
            double clickConversionAverage = 0;
            double addToCartConversionTotal = 0;
            double cartToOrderConversionTotal = 0;

            double turnoverInDays = 0;
            double centerStockSumTotal = 0;
            double piterStockSumTotal = 0;
            double kazanStockSumTotal = 0;
            double ugStockSumTotal = 0;
            double uralStockSumTotal = 0;
            double sibirStockSumTotal = 0;
            double toClientTotal = 0;
            double fromClientTotal = 0;
            double totalStocksSumTotal = 0;

            int rowNumberData = 3;
            int rowNumberTotalData = 3;
            int interactionCount = 0;

            while (rowNumberData < columnValues.Count)
            {
                raitingCount += GetSumData(columnValues, rowNumberData) == 0 ? 0 : 1;
                sumRaiting += GetSumData(columnValues, rowNumberData++);
                marginProfit += GetSumData(columnValues, rowNumberData++);
                rowNumberData++;
                rowNumberData++;
                ordersSumRubTotal += GetSumData(columnValues, rowNumberData++);
                rowNumberData++;
                priceBeforeSPP += GetSumData(columnValues, rowNumberData++);
                spp += GetSumData(columnValues, rowNumberData++);
                rowNumberData++;
                openCardCountTotal += GetSumData(columnValues, rowNumberData++);
                addToCartCountTotal += GetSumData(columnValues, rowNumberData++);
                ordersCountTotal += GetSumData(columnValues, rowNumberData++);
                rowNumberData++;
                viewsCountTotal += GetSumData(columnValues, rowNumberData++);
                clicksCountTotal += GetSumData(columnValues, rowNumberData++);
                expenseTotal += GetSumData(columnValues, rowNumberData++);
                rowNumberData++;
                rowNumberData++;
                clickConversionAverage += GetSumData(columnValues, rowNumberData++);
                addToCartConversionTotal += GetSumData(columnValues, rowNumberData++);
                cartToOrderConversionTotal += GetSumData(columnValues, rowNumberData++);
                rowNumberData++;
                turnoverInDays += GetSumData(columnValues, rowNumberData++);
                totalStocksSumTotal += GetSumData(columnValues, rowNumberData++);
                centerStockSumTotal += GetSumData(columnValues, rowNumberData++);
                piterStockSumTotal += GetSumData(columnValues, rowNumberData++);
                kazanStockSumTotal += GetSumData(columnValues, rowNumberData++);
                ugStockSumTotal += GetSumData(columnValues, rowNumberData++);
                uralStockSumTotal += GetSumData(columnValues, rowNumberData++);
                sibirStockSumTotal += GetSumData(columnValues, rowNumberData++);
                toClientTotal += GetSumData(columnValues, rowNumberData++);
                fromClientTotal += GetSumData(columnValues, rowNumberData++);
                rowNumberData += DistanceBetweenData;
                interactionCount++;
            }

            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (sumRaiting / raitingCount).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (marginProfit).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (ordersSumRubTotal <= 0 ? 0 : marginProfit / ordersSumRubTotal).ToString("0.#", _cultureInfo);
            rowNumberTotalData++;
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (ordersSumRubTotal).ToString("0.#", _cultureInfo);
            rowNumberTotalData++;
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (priceBeforeSPP / interactionCount).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (spp / interactionCount).ToString("0.#", _cultureInfo);
            rowNumberTotalData++;
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (openCardCountTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (addToCartCountTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (ordersCountTotal).ToString("0.#", _cultureInfo);
            rowNumberTotalData++;
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (viewsCountTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (clicksCountTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (expenseTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (ordersSumRubTotal <= 0 ? 0 : expenseTotal / ordersSumRubTotal * PercentBase).ToString("0.#", _cultureInfo);
            rowNumberTotalData++;
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (clickConversionAverage / interactionCount).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (addToCartCountTotal / openCardCountTotal * PercentBase).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (ordersCountTotal / addToCartCountTotal * PercentBase).ToString("0.#", _cultureInfo);
            rowNumberTotalData++;
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (turnoverInDays / interactionCount).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (totalStocksSumTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (centerStockSumTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (piterStockSumTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (kazanStockSumTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (ugStockSumTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (uralStockSumTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (sibirStockSumTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (toClientTotal).ToString("0.#", _cultureInfo);
            dataToUpdate[$"{_RNPListName}!{columnName}{rowNumberTotalData++}"] = (fromClientTotal).ToString("0.#", _cultureInfo);
        }

        private void AddStocksData(Dictionary<string, string> dataToUpdate, Dictionary<string, StockData> stocksData)
        {
            Dictionary<int, StockData> rowsToData = [];
            List<StockData> stocks = [];
            List<StockData> DataNotFoundRowsProducts = [];

            foreach (StockData stock in stocksData.Values)
            {
                stocks.Add(stock);
            }

            TakeRowsSalesByBarcodes(stocks, rowsToData, DataNotFoundRowsProducts);

            int lastRow = GetLastRow(_currentSpreadsheetId!, _defaultRangeLastRow);

            foreach (StockData stock in DataNotFoundRowsProducts)
            {
                int quantity = stock.WareHouses!
                    .Where(warehouse => warehouse.Name == "Всего находится на складах")
                    .Select(warehouse => warehouse.Quantity)
                    .FirstOrDefault();

                dataToUpdate[$"{_turnoverListName}!{ArticuleColumn}{lastRow}"] = stock.VendoreCode!;
                dataToUpdate[$"{_turnoverListName}!{BarcodeColumn}{lastRow}"] = stock.Barcode!;
                dataToUpdate[$"{_turnoverListName}!{SizesColumn}{lastRow}"] = stock.TechSize!;
                dataToUpdate[$"{_turnoverListName}!{RemainsWBColumn}{lastRow}"] = quantity.ToString();
                lastRow++;
            }

            foreach (int row in rowsToData.Keys)
            {
                int quantity = rowsToData[row].WareHouses!
                    .Where(warehouse => warehouse.Name == "Всего находится на складах")
                    .Select(warehouse => warehouse.Quantity)
                    .FirstOrDefault();

                dataToUpdate[$"{_turnoverListName}!{RemainsWBColumn}{row}"] = quantity.ToString();
                dataToUpdate[$"{_turnoverListName}!{SizesColumn}{row}"] = rowsToData[row].TechSize!;
            }
        }

        private void AddSalesData(Dictionary<string, SalesData> salesData, Dictionary<string, string> dataToUpdate, int todayDay)
        {
            int lastRow;
            Dictionary<int, SalesData> rowsToData = [];
            List<SalesData> sales = [];
            List<SalesData> DataNotFoundRowsProducts = [];

            foreach (SalesData sale in salesData.Values)
            {
                sales.Add(sale);
            }

            TakeRowsSalesByBarcodes(sales, rowsToData, DataNotFoundRowsProducts);

            string dataColumn = _columnsAndDates[todayDay - 1];

            foreach (int row in rowsToData.Keys)
            {
                dataToUpdate[$"{_turnoverListName}!{dataColumn}{row}"] = rowsToData[row].Amount.ToString();
            }

            if (DataNotFoundRowsProducts.Count > 0)
            {
                foreach (SalesData sale in DataNotFoundRowsProducts)
                {
                    lastRow = GetLastRow(_currentSpreadsheetId!, _defaultRangeLastRow);

                    if (dataToUpdate.ContainsKey($"{_turnoverListName}!{ArticuleColumn}{lastRow}") == false && dataToUpdate.ContainsKey($"{_turnoverListName}!{RemainsWBColumn}{lastRow}") == false
                    && dataToUpdate.ContainsKey($"{_turnoverListName}!{SizesColumn}{lastRow}") == false && dataToUpdate.ContainsKey($"{_turnoverListName}!{dataColumn}{lastRow}") == false)
                    {
                        dataToUpdate[$"{_turnoverListName}!{ArticuleColumn}{lastRow}"] = sale.VendoreCode!;
                        dataToUpdate[$"{_turnoverListName}!{BarcodeColumn}{lastRow}"] = sale.Barcode!;
                        dataToUpdate[$"{_turnoverListName}!{SizesColumn}{lastRow}"] = sale.TechSize!;
                        dataToUpdate[$"{_turnoverListName}!{RemainsWBColumn}{lastRow}"] = "0";
                        dataToUpdate[$"{_turnoverListName}!{dataColumn}{lastRow}"] = sale.Amount.ToString();
                        lastRow++;
                    }
                }
            }
        }

        private async Task AddAverageData(Dictionary<string, string> dataToUpdate, DateTime todayDate)
        {
            int lastDataRow = GetLastRow(_currentSpreadsheetId!, _defaultRangeLastRow) - 1;
            ValueRange valueRange = new();

            try
            {
                if (todayDate.Day >= TweentyOneDays)
                {
                    valueRange = GetValueByShellsRange(_currentSpreadsheetId!, $"{_turnoverListName}!{_columnsAndDates[todayDate.Day - TweentyOneDays]}{FirstDataRowTurnover}:{_columnsAndDates[todayDate.Day - 1]}{lastDataRow}");
                }
                else
                {
                    ValueRange valueRange1 = GetValueByShellsRange(_currentSpreadsheetId!, $"{_turnoverListName}!{_columnsAndDates[0]}{FirstDataRowTurnover}:{_columnsAndDates[todayDate.Day - 1]}{lastDataRow}");
                    ValueRange valueRange2 = GetValueByShellsRange(_currentSpreadsheetId!, $"{_turnoverListName}!{_columnsForLastMonthSales[todayDate.AddDays(-TweentyOneDays).Day]}{FirstDataRowTurnover}:{_columnsForLastMonthSales[todayDate.AddDays(-todayDate.Day).Day - 1]}{lastDataRow}");

                    valueRange = new()
                    {
                        Values = MergeRows(valueRange1.Values, valueRange2.Values)
                    };
                }

                for (int i = 0; i < valueRange.Values.Count; i++)
                {
                    IList<object> row = valueRange.Values[i];

                    List<int> values = row
                        .Select(value => int.TryParse(value?.ToString(), out int parsedValue) ? parsedValue : 0)
                        .ToList();

                    if (todayDate.Day < TweentyOneDays)
                    {
                        if (values.Count < 21)
                        {

                        }
                    }

                    int rowNumber = FirstDataRowTurnover + i;

                    double avg3 = CalculateAverageData(ThreeDays, values);
                    double avg7 = CalculateAverageData(SevenDays, values);
                    double avg21 = CalculateAverageData(TweentyOneDays, values);

                    dataToUpdate[$"{_turnoverListName}!{ThreeDaysColumn}{rowNumber}"] = avg3.ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_turnoverListName}!{SevenDaysColumn}{rowNumber}"] = avg7.ToString("0.#", _cultureInfo);
                    dataToUpdate[$"{_turnoverListName}!{TweentyOneDaysColumn}{rowNumber}"] = avg21.ToString("0.#", _cultureInfo);
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Произошла ошибка в AddAverageData: {exception}");
            }

            await Task.Delay(10000);
        }

        private async Task AddSummaryData()
        {
            Spreadsheet destinationSpreadsheet = await _sheetsService.Spreadsheets.Get(_summarySpreadsheetId).ExecuteAsync();

            if (await TryAddNewList(_currentSpreadsheet, destinationSpreadsheet, _summarySpreadsheetId!) == false)
            {
                destinationSpreadsheet = await _sheetsService.Spreadsheets.Get(_summarySpreadsheetId).ExecuteAsync();
            }

            int lastDataRow = GetLastRow(_currentSpreadsheetId!, _defaultRangeLastRow) - 1;

            CopyPastData copyPastDataFirst = new(
                _currentSpreadsheetId!,
                _turnoverListName,
                _summarySpreadsheetId!,
                _currentSpreadsheet.Properties.Title,
                4,
                lastDataRow,
                StartColumnForMainData,
                EndColumnForMainData,
                1,
                lastDataRow,
                StartColumnForMainData,
                EndColumnForMainData
            );

            CopyPastData copyPastDataSecond = new(
                _currentSpreadsheetId!,
                _turnoverListName,
                _summarySpreadsheetId!,
                _currentSpreadsheet.Properties.Title,
                4,
                lastDataRow,
                ThreeDaysColumn,
                TweentyOneDaysColumn,
                1,
                lastDataRow,
                StartColumnForAverageData,
                EndColumnForAverageData
            );

            await ClearData(_summarySpreadsheetId!, $"{_currentSpreadsheet.Properties.Title}!A1:{EndColumnForAverageData}{lastDataRow}");
            await CopyPastSummaryData(_currentSpreadsheet, destinationSpreadsheet, copyPastDataFirst);
            await CopyPastSummaryData(_currentSpreadsheet, destinationSpreadsheet, copyPastDataSecond);
        }

        private void AddMarketingCardData(int articuleRow, string column, CardMarketingData cardMarketingData, Dictionary<string, string> dataToUpdate, ValueRange lastDayColumnValueRange)
        {
            string raiting = cardMarketingData.Raiting > 0
                ? cardMarketingData.Raiting.ToString("0.#", _cultureInfo)
                : "0";

            if (cardMarketingData.Raiting <= 0 &&
                column != _columnsForSalesFunnelData[0] &&
                lastDayColumnValueRange?.Values != null)
            {
                int previousRowIndex = articuleRow - 1;
                if (previousRowIndex >= 0 && previousRowIndex < lastDayColumnValueRange.Values.Count)
                {
                    var previousValue = lastDayColumnValueRange.Values[previousRowIndex];
                    if (previousValue?.Count > 0 && previousValue[0] != null)
                    {
                        raiting = previousValue[0].ToString()!;
                    }
                }
            }

            void AddData(string value) =>
                dataToUpdate[$"{_RNPListName}!{column}{articuleRow++}"] = value ?? string.Empty;

            AddData(raiting);
            AddData(cardMarketingData.Margin.ToString("0.#", _cultureInfo));
            AddData((cardMarketingData.OrdersSumRub <= 0 ? 0 : cardMarketingData.Margin / cardMarketingData.OrdersSumRub).ToString("0.#", _cultureInfo));
            articuleRow++;
            AddData(cardMarketingData.OrdersSumRub.ToString("0.#", _cultureInfo));
            articuleRow++;
            AddData(cardMarketingData.PriceBeforeSPP.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.SPP.ToString());
            articuleRow++;
            AddData(cardMarketingData.OpenCardCount.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.AddToCartCount.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.OrdersCount.ToString("0.#", _cultureInfo));
            articuleRow++;
            AddData(cardMarketingData.ViewsCount.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.ClicksCount.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.Expenses.ToString("0.#", _cultureInfo));
            AddData((cardMarketingData.OrdersSumRub <= 0 ? 0 : cardMarketingData.Expenses / cardMarketingData.OrdersSumRub * PercentBase).ToString("0.#", _cultureInfo));
            articuleRow++;
            AddData(cardMarketingData.ClickConversion.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.AddToCartPercent.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.CartToOrderPercent.ToString("0.#", _cultureInfo));
            articuleRow++;
            AddData(cardMarketingData.TurnoverInDays.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.TotalStocksSum.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.CenterStockSum.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.PiterStockSum.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.KazanStockSum.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.UgStockSum.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.UralStockSum.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.SibirStockSum.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.ToClient.ToString("0.#", _cultureInfo));
            AddData(cardMarketingData.FromClient.ToString("0.#", _cultureInfo));
        }

        private void AddNewMarketingCardData(string articule, string column, CardMarketingData cardMarketingData, Dictionary<string, string> dataToUpdate, ValueRange lastDayColumnValueRange, string sourceSpreadsheetName, List<Request> requests, int? sheetId, ref int articulesInfoRowNumber, ref int dataRowNumber)
        {

            int firstRowForBorders = articulesInfoRowNumber - 1;
            int indicatorRowNumber;

            dataToUpdate[$"{_RNPListName}!{FirstDataColumnSalesFunnel}{articulesInfoRowNumber++}"] = sourceSpreadsheetName;
            dataToUpdate[$"{_RNPListName}!{FirstDataColumnSalesFunnel}{articulesInfoRowNumber}"] = articule;

            indicatorRowNumber = articulesInfoRowNumber;

            dataToUpdate[$"{_RNPListName}!{PeriodPlanColumnSalesFunnel}{articulesInfoRowNumber}"] = "План периода";
            dataToUpdate[$"{_RNPListName}!{FactPeriodColumnSalesFunnel}{articulesInfoRowNumber}"] = "Факт периода";
            dataToUpdate[$"{_RNPListName}!{NormPerDayColumnSalesFunnel}{articulesInfoRowNumber++}"] = "Норма в день";
            dataToUpdate[$"{_RNPListName}!{FirstDataColumnSalesFunnel}{articulesInfoRowNumber}"] = $"=Image(\"{cardMarketingData.ImageUrl}\")";

            requests.Add(GetMergeCellsRequest(sheetId, articulesInfoRowNumber - 1, articulesInfoRowNumber + ArticuleImageCellsCount, 0, 1));

            articulesInfoRowNumber += ArticuleImageCellsCount + 2;

            dataToUpdate[$"{_RNPListName}!{FirstDataColumnSalesFunnel}{articulesInfoRowNumber++}"] = "Остаток в днях";
            articulesInfoRowNumber++;
            dataToUpdate[$"{_RNPListName}!{FirstDataColumnSalesFunnel}{articulesInfoRowNumber++}"] = "Остаток наш склад";
            articulesInfoRowNumber++;
            dataToUpdate[$"{_RNPListName}!{FirstDataColumnSalesFunnel}{articulesInfoRowNumber++}"] = "В пути к нам";
            articulesInfoRowNumber++;
            dataToUpdate[$"{_RNPListName}!{FirstDataColumnSalesFunnel}{articulesInfoRowNumber++}"] = "В производстве";

            int indicatorsColumnSalesFunnelIndex = GetColumnIndex(IndicatorsColumnSalesFunnel) - 1;

            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Показатель";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Рейтинг карточки";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Маржинальная прибыль, руб";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Маржа, %";
            requests.Add(GetColorCellRequest(sheetId, indicatorRowNumber - 1, indicatorRowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Заказы, руб";
            requests.Add(GetColorCellRequest(sheetId, indicatorRowNumber - 1, indicatorRowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Средний чек без СПП, руб";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "СПП, %";
            requests.Add(GetColorCellRequest(sheetId, indicatorRowNumber - 1, indicatorRowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Общий трафик, клики";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Корзины, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Заказы, шт";
            requests.Add(GetColorCellRequest(sheetId, indicatorRowNumber - 1, indicatorRowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Показы, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Клики с РК, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Расход РК, руб";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "ДРР, %";
            requests.Add(GetColorCellRequest(sheetId, indicatorRowNumber - 1, indicatorRowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "CTR, %";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Конверсия в корзину, %";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Конверсия в заказ, %";
            requests.Add(GetColorCellRequest(sheetId, indicatorRowNumber - 1, indicatorRowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Оборачиваемость, дни";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "Остаток FBO, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "- Центр";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "- Питер";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "- Казань";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "- Юг";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "- Урал";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "- Сибирь";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "В пути к клиенту, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{indicatorRowNumber++}"] = "В пути от клиента, шт";

            AddMarketingCardData(dataRowNumber, column, cardMarketingData, dataToUpdate, lastDayColumnValueRange);

            dataRowNumber = indicatorRowNumber + DistanceBetweenData;
            requests.Add(GetBordersRequest(sheetId, firstRowForBorders, dataRowNumber, 0, GrayRowSize));
            requests.Add(GetAlignmentRequest(sheetId, firstRowForBorders, dataRowNumber, 0, GrayRowSize));

            articulesInfoRowNumber = indicatorRowNumber;
        }

        private async Task AddNewDates(string columnName, Dictionary<string, string> dataToUpdate, DateTime todayDate, int? sheetId, int lastRowData, List<Request> requests)
        {
            int nextMonthDaysCount = DateTime.DaysInMonth(todayDate.Year, todayDate.Month);
            int startColumnIndex = GetColumnIndex(columnName) + 1;
            int endColumnIndex = startColumnIndex + nextMonthDaysCount;
            DateTime firstDayOfMonth = new(todayDate.Year, todayDate.Month, 1);

            await AddNewColumns(_currentSpreadsheetId!, sheetId, startColumnIndex, endColumnIndex);

            for (int i = 0; i < nextMonthDaysCount; i++)
            {
                int currentColumnIndex = startColumnIndex + i;
                string columnLetter = GetColumnName(currentColumnIndex);
                DateTime currentDate = firstDayOfMonth.AddDays(i);
                string formattedDate = currentDate.ToString("dd.MM.yyyy");
                dataToUpdate[$"{_RNPListName}!{columnLetter}1"] = formattedDate;
            }

            requests.Add(GetBordersRequest(sheetId, 0, lastRowData, startColumnIndex - 1, endColumnIndex - 1));
        }

        private async Task AddNewColumns(string spreadsheetId, int? sheetId, int startIndex, int endIndex)
        {
            if (sheetId == -1)
            {
                Console.WriteLine("Не удалось получить sheetId");
                return;
            }

            List<Request> requests =
            [
                new() {
                    InsertDimension = new InsertDimensionRequest
                    {
                        Range = new DimensionRange
                        {
                            SheetId = sheetId,
                            Dimension = "COLUMNS",
                            StartIndex = startIndex,
                            EndIndex = endIndex
                        },
                        InheritFromBefore = false
                    }
                }
            ];

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };

            await _sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).ExecuteAsync();
        }

        private async Task AddNewList(string spreadSheetIdForNewList, string newListName)
        {
            Request addSheetRequest = new()
            {
                AddSheet = new AddSheetRequest
                {
                    Properties = new SheetProperties
                    {
                        Title = newListName,
                        Index = 0,
                        GridProperties = new GridProperties
                        {
                            RowCount = 1000,
                            ColumnCount = 26
                        }
                    }
                }
            };

            BatchUpdateSpreadsheetRequest batchUpdateRequest = new()
            {
                Requests = [addSheetRequest]
            };

            await _sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetIdForNewList).ExecuteAsync();
        }

        private async Task AddStartRowsRNPData(Dictionary<string, string> dataToUpdate, int? sheetId)
        {
            List<Request> requests = [];
            int rowNumber = 2;
            int indicatorsColumnSalesFunnelIndex = GetColumnIndex(IndicatorsColumnSalesFunnel) - 1;

            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Показатель";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Рейтинг карточки";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Маржинальная прибыль, руб";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Маржа, %";
            requests.Add(GetColorCellRequest(sheetId, rowNumber - 1, rowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Заказы, руб";
            requests.Add(GetColorCellRequest(sheetId, rowNumber - 1, rowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Средний чек без СПП, руб";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "СПП, %";
            requests.Add(GetColorCellRequest(sheetId, rowNumber - 1, rowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Общий трафик, клики";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Корзины, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Заказы, шт";
            requests.Add(GetColorCellRequest(sheetId, rowNumber - 1, rowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Показы, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Клики с РК, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Расход РК, руб";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "ДРР, %";
            requests.Add(GetColorCellRequest(sheetId, rowNumber - 1, rowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "CTR, %";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Конверсия в корзину, %";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Конверсия в заказ, %";
            requests.Add(GetColorCellRequest(sheetId, rowNumber - 1, rowNumber++, indicatorsColumnSalesFunnelIndex, GrayRowSize, _grayColor));
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Оборачиваемость, дни";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "Остаток FBO, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "- Центр";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "- Питер";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "- Казань";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "- Юг";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "- Урал";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "- Сибирь";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "В пути к клиенту, шт";
            dataToUpdate[$"{_RNPListName}!{IndicatorsColumnSalesFunnel}{rowNumber++}"] = "В пути от клиента, шт";

            requests.Add(GetBordersRequest(sheetId, 0, rowNumber, 0, GrayRowSize));

            if (requests != null && requests.Count > 0)
            {
                await BatchUpdateCells(requests, _currentSpreadsheetId!);
            }
        }

        //Get methods

        public List<string> GetSpreadshetIds()
        {
            var request = _driveService.Files.List();
            request.Q = $"'{_googleFolderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";
            request.Fields = "files(id, name)";

            var result = request.Execute();

            return result.Files.Where(item => item.Name.StartsWith(_startWordSpreadsheetName)).Select(item => item.Id).ToList<string>();
        }

        private ValueRange GetValueByShellsRange(string spreadsheetId, string range)
        {
            return _sheetsService.Spreadsheets.Values.Get(spreadsheetId, range).Execute();
        }

        private string GetApiKeyFromList()
        {
            string range = $"{_turnoverListName}!{ApiKeyShellAdress}";

            try
            {
                ValueRange response = _sheetsService.Spreadsheets.Values.Get(_currentSpreadsheetId, range).Execute();

                if (response.Values != null && response.Values.Count > 0 && response.Values[0].Count > 0)
                {
                    return response.Values[0][0]?.ToString() ?? string.Empty;
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Empty;
        }

        private int GetLastRow(string spreadSheetId, string range)
        {
            try
            {
                ValueRange response = _sheetsService.Spreadsheets.Values.Get(spreadSheetId, range).Execute();
                return response.Values?.Count + 1 ?? 1;
            }
            catch (Exception exciption)
            {
                Console.WriteLine("Ошибка получения последний строки");
                Console.WriteLine(exciption);
                return 1;
            }
        }

        private void GetDataForMargin(Dictionary<int, DataForMargin> articuleToDataForMargin)
        {
            try
            {
                ValueRange unitValueRange = GetValueByShellsRange(_currentSpreadsheetId!, $"{_unitListName}!A2:I{GetLastRow(_currentSpreadsheetId!, $"{_unitListName}!A:A")}");

                if (unitValueRange == null || unitValueRange.Values == null || unitValueRange.Values.Count == 0)
                    return;

                foreach (var row in unitValueRange.Values)
                    articuleToDataForMargin[Convert.ToInt32(row[0])] = new(double.TryParse(row[1].ToString(), out double value) ? value : 0, double.TryParse(row[3].ToString(), out value) ? value : 0, double.TryParse(row[4].ToString(), out value) ? value : 0, double.TryParse(row[5].ToString(), out value) ? value : 0, double.TryParse(row[6].ToString(), out value) ? value : 0, double.TryParse(row[7].ToString(), out value) ? value : 0, double.TryParse(row[8].ToString(), out value) ? value : 0);

            }
            catch (Exception exception)
            {
                Console.WriteLine("Ошибка с получением маржи");
                Console.WriteLine(exception);
            }
        }

        //Take methods

        private async Task TakeMarketingDataRows(Dictionary<string, CardMarketingData> articuleToCardMarketingData, Dictionary<int, string> rowsFoundArticules, List<string> notFoundArticules)
        {
            string rangeArticules = $"{_RNPListName}!A:A";
            ValueRange responseArticules = await _sheetsService.Spreadsheets.Values.Get(_currentSpreadsheetId!, rangeArticules).ExecuteAsync();

            if (responseArticules.Values == null || responseArticules.Values.Count == 0)
            {
                notFoundArticules.AddRange(articuleToCardMarketingData.Keys);

                return;
            }

            foreach (string tempoArticule in articuleToCardMarketingData.Keys)
            {
                bool isFound = false;

                for (int i = 0; i < responseArticules.Values.Count; i++)
                {
                    string? articule = responseArticules.Values[i].FirstOrDefault()?.ToString();

                    if (tempoArticule == articule)
                    {
                        rowsFoundArticules[i + 1] = tempoArticule;
                        isFound = true;
                        break;
                    }
                }

                if (isFound == false)
                {
                    notFoundArticules.Add(tempoArticule);
                }
            }
        }

        private void TakeRowsSalesByBarcodes<T>(List<T> productsData, Dictionary<int, T> rowsToDataFoundProducts, List<T> DataNotFoundRowsProducts) where T : IData
        {
            string rangeBarcodes = $"{_turnoverListName}!{BarcodeColumn}:{BarcodeColumn}";
            ValueRange responseBarcodes = _sheetsService.Spreadsheets.Values.Get(_currentSpreadsheetId, rangeBarcodes).Execute();

            if (responseBarcodes.Values == null)
                return;

            foreach (T product in productsData)
            {
                bool isFound = false;

                for (int i = 0; i < responseBarcodes.Values.Count; i++)
                {
                    string? barcode = responseBarcodes.Values[i].FirstOrDefault()?.ToString();

                    if (product.Barcode == barcode)
                    {
                        rowsToDataFoundProducts[i + 1] = product;
                        isFound = true;
                        break;
                    }
                }

                if (isFound == false)
                {
                    DataNotFoundRowsProducts.Add(product);
                }
            }
        }

        //Try methods

        private bool TryGetList()
        {
            if (string.IsNullOrEmpty(_turnoverListName))
                return false;

            Spreadsheet spreadsheet = _sheetsService.Spreadsheets.Get(_currentSpreadsheetId).Execute();
            return spreadsheet.Sheets.Any(sheet => sheet.Properties.Title == _turnoverListName);
        }

        private async Task<bool> TryAddNewList(Spreadsheet sourceSpreadsheet, Spreadsheet destinationSpreadsheet, string spreadSheetIdForNewList)
        {
            foreach (Sheet sheet in destinationSpreadsheet.Sheets)
            {
                if (sheet.Properties.Title == sourceSpreadsheet.Properties.Title)
                {
                    return true;
                }
            }

            await AddNewList(spreadSheetIdForNewList, sourceSpreadsheet.Properties.Title);
            return false;
        }

        //Batch methods

        private async Task BatchUpdateCells(List<Request> requests, string spreadsheetId)
        {
            BatchUpdateSpreadsheetRequest batchUpdateRequest = new()
            {
                Requests = requests
            };

            await _sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).ExecuteAsync();
        }

        private async Task BatchUpdateValues(Dictionary<string, string> dataToUpdate, string spreadsheetId)
        {
            const int maxBatchSize = 1000;
            var batches = dataToUpdate
                .Select((kvp, index) => new { kvp, index })
                .GroupBy(x => x.index / maxBatchSize)
                .Select(g => g.Select(x => x.kvp).ToList())
                .ToList();

            foreach (var batch in batches)
            {
                List<ValueRange> valueRanges = batch
                    .Select(kvp => new ValueRange { Range = kvp.Key, Values = new List<IList<object>> { new List<object> { kvp.Value } } })
                    .ToList();

                var batchRequest = new BatchUpdateValuesRequest
                {
                    ValueInputOption = "USER_ENTERED",
                    Data = valueRanges,
                };

                bool success = false;
                int retryCount = 0;
                const int maxRetries = 3;

                while (!success && retryCount < maxRetries)
                {
                    try
                    {
                        var request = _sheetsService.Spreadsheets.Values.BatchUpdate(batchRequest, spreadsheetId);
                        await request.ExecuteAsync();
                        success = true;

                        if (batches.Count > 1)
                            await Task.Delay(1000);
                    }
                    catch (GoogleApiException ex) when (ex.HttpStatusCode == HttpStatusCode.TooManyRequests)
                    {
                        retryCount++;
                        if (retryCount >= maxRetries)
                        {
                            Console.WriteLine($"Failed after {maxRetries} retries: {ex.Message}");
                            throw;
                        }

                        int delay = (int)Math.Pow(2, retryCount) * 1000;
                        await Task.Delay(delay);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during batch update: {ex.Message}");
                        throw;
                    }
                }
            }
        }

        //Other

        private async Task PrepareTurnoverSheet(Dictionary<string, SalesData> salesData, Dictionary<string, StockData> stocksData, Dictionary<string, string> dataToUpdate, int startValue, DateTime todayDate)
        {
            try
            {
                await PrepareTurnoverSheet(salesData, stocksData, dataToUpdate, todayDate, startValue, _turnoverSheetid);

                dataToUpdate.Clear();
                FillEmptyDataRows(dataToUpdate, todayDate);
                await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
                dataToUpdate.Clear();
                await Task.Delay(1000);

                await AddSummaryData();
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Ошибка в подготовке листа оборачиваемости {_currentSpreadsheet.Properties.Title} | {_currentSpreadsheetId}:\n{exception}");
            }
        }

        private async Task PrepareRNPSheet(Dictionary<int, Good> articulesToGoods, Dictionary<string, Card> urlToCardData, Dictionary<int, ProductData> articulesToMarketingData, Dictionary<string, StockData> stocksData, Dictionary<string, string> dataToUpdate, int? sheetId, DateTime todayDate, int startValue)
        {
            try
            {
                await AddStartRowsRNPData(dataToUpdate, sheetId);
                await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
                dataToUpdate.Clear();

                for (int i = startValue; i >= 0; i--)
                {
                    urlToCardData.Clear();
                    articulesToMarketingData.Clear();
                    articulesToGoods.Clear();

                    await AddTodayRNPData(urlToCardData, articulesToMarketingData, articulesToGoods, stocksData, dataToUpdate, todayDate.AddDays(-i));

                    await Task.Delay(60000);
                }

                dataToUpdate.Clear();
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Ошибка в подготовке листа РНП {_currentSpreadsheet.Properties.Title} | {_currentSpreadsheetId}:\n{exception}");
            }
        }

        private async Task PrepareTurnoverSheet(Dictionary<string, SalesData> salesData, Dictionary<string, StockData> stocksData, Dictionary<string, string> dataToUpdate, DateTime todayDate, int startValue, int? sheetId)
        {
            await AddNewColumns(_currentSpreadsheetId!, sheetId, GetColumnIndex(TweentyOneDaysColumn), GetColumnIndex(TweentyOneDaysColumn) + ColumnsCountForNewSheet);

            AddStocksData(dataToUpdate, stocksData);
            await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
            dataToUpdate.Clear();

            for (int i = startValue; i >= 0; i--)
            {
                salesData.Clear();
                await _wBApiHelper.GetSalesData(salesData, todayDate.AddDays(-i));

                AddSalesData(salesData, dataToUpdate, todayDate.AddDays(-i).Day);

                await Task.Delay(60000);
            }

            await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);
            dataToUpdate.Clear();

            await AddAverageData(dataToUpdate, todayDate);
            await BatchUpdateValues(dataToUpdate, _currentSpreadsheetId!);

            dataToUpdate.Clear();
        }

        private async Task CopyAndClearData()
        {
            int lastDataRow = GetLastRow(_currentSpreadsheetId!, _defaultRangeLastRow) - 1;

            await CopyPastDataInCurrentTable(_currentSpreadsheetId!, _turnoverSheetid, FirstDataRowTurnover - 1, lastDataRow, GetColumnIndex(ArticuleColumn) - 1, GetColumnIndex(ArticuleColumn), FirstDataRowTurnover - 1, lastDataRow, GetColumnIndex(_columnsForLastMonthSales[0]) - 1, GetColumnIndex(_columnsForLastMonthSales[0]));
            await CopyPastDataInCurrentTable(_currentSpreadsheetId!, _turnoverSheetid, FirstDataRowTurnover - 1, lastDataRow, GetColumnIndex(_columnsForSales[0]) - 1, GetColumnIndex(_columnsForSales[^1]), FirstDataRowTurnover - 1, lastDataRow, GetColumnIndex(_columnsForLastMonthSales[1]) - 1, GetColumnIndex(_columnsForLastMonthSales[^1]));
            await ClearData(_currentSpreadsheetId!, $"{_turnoverListName}!{_columnsForSales[0]}{FirstDataRowTurnover}:{TweentyOneDaysColumn}{GetLastRow(_currentSpreadsheetId!, _defaultRangeLastRow)}");
        }

        private async Task ClearData(string spreadsheetId, string range)
        {
            ClearValuesRequest clearValuesRequest = new();

            await _sheetsService.Spreadsheets.Values.Clear(clearValuesRequest, spreadsheetId, range).ExecuteAsync();
        }

        private async Task CopyPastDataInCurrentTable(string spreadSheetId, int? sheetId,
            int startRowForCopy, int endRowForCopy, int startColumnForCopy, int endColumnForCopy,
            int startRowForPast, int endRowForPast, int startColumnForPast, int endColumnForPast)
        {

            List<Request> requests =
            [
                new Request
                {
                    CopyPaste = new CopyPasteRequest
                    {
                        Source = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = startRowForCopy,
                            EndRowIndex = endRowForCopy,
                            StartColumnIndex = startColumnForCopy,
                            EndColumnIndex = endColumnForCopy
                        },
                        Destination = new GridRange
                        {
                            SheetId = sheetId,
                            StartRowIndex = startRowForPast,
                            EndRowIndex = endRowForPast,
                            StartColumnIndex = startColumnForPast,
                            EndColumnIndex = endColumnForPast
                        },
                        PasteType = "PASTE_NORMAL",
                        PasteOrientation = "NORMAL"
                    }
                },
            ];

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };

            await _sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).ExecuteAsync();
        }

        private async Task CopyPastSummaryData(Spreadsheet sourceSpreadsheet, Spreadsheet destinationSpreadsheet, CopyPastData copyPastData)
        {
            int? listIdForCopy = GetSheetId(sourceSpreadsheet, copyPastData.ListNameCopy);
            int? listIdForPast = GetSheetId(destinationSpreadsheet, copyPastData.ListNamePast);

            if (listIdForCopy == null || listIdForPast == null)
            {
                Console.WriteLine("Лист не найден по id.");
                return;
            }

            string sourceRange = $"{copyPastData.ListNameCopy}!{copyPastData.StartColumnCopy}{copyPastData.StartRowIndexCopy}:{copyPastData.EndColumnCopy}{copyPastData.EndRowIndexCopy}";

            ValueRange sourceData = await _sheetsService.Spreadsheets.Values.Get(copyPastData.SpreadSheetCopyId, sourceRange).ExecuteAsync();

            if (sourceData.Values == null || sourceData.Values.Count == 0)
            {
                Console.WriteLine("Нет данных для копирования.");
                return;
            }

            string destinationRange = $"{copyPastData.ListNamePast}!{copyPastData.StartColumnPast}{copyPastData.StartRowIndexPast}:{copyPastData.EndColumnPast}{copyPastData.EndRowIndexPast}";

            ValueRange destinationData = new()
            {
                Values = sourceData.Values
            };

            UpdateRequest updateRequest = _sheetsService.Spreadsheets.Values.Update(destinationData, copyPastData.SpreadSheetPastId, destinationRange);
            updateRequest.ValueInputOption = UpdateRequest.ValueInputOptionEnum.RAW;
            await updateRequest.ExecuteAsync();
        }

        private void FillEmptyDataRows(Dictionary<string, string> dataToUpdate, DateTime todayDate)
        {
            string dataColumn = _columnsAndDates[todayDate.Day - 1];

            ValueRange lastDataColumnRange = GetValueByShellsRange(_currentSpreadsheetId!, $"{_turnoverListName}!{dataColumn}{FirstDataRowTurnover}:{dataColumn}");
            ValueRange articulesColumnRange = GetValueByShellsRange(_currentSpreadsheetId!, $"{_turnoverListName}!{ArticuleColumn}{FirstDataRowTurnover}:{ArticuleColumn}");

            if (lastDataColumnRange.Values == null || lastDataColumnRange.Values.Count == 0)
            {
                for (int i = 0; i < articulesColumnRange.Values.Count; i++)
                {
                    dataToUpdate[$"{_turnoverListName}!{dataColumn}{i + FirstDataRowTurnover}"] = "0";
                }
            }
            else
            {
                lastDataColumnRange.Values
                    .Select((row, index) => new { Row = row, OriginalIndex = index })
                    .Where(x => x.Row.Count == 0 || x.Row[0] == null || string.IsNullOrEmpty(x.Row[0].ToString()))
                    .ToList()
                    .ForEach(x =>
                    {
                        dataToUpdate[$"{_turnoverListName}!{dataColumn}{x.OriginalIndex + FirstDataRowTurnover}"] = "0";
                    });

                if (lastDataColumnRange.Values.Count < articulesColumnRange.Values.Count)
                {
                    for (int i = 0; i < articulesColumnRange.Values.Count - lastDataColumnRange.Values.Count; i++)
                    {
                        dataToUpdate[$"{_turnoverListName}!{dataColumn}{lastDataColumnRange.Values.Count + FirstDataRowTurnover + i}"] = "0";
                    }
                }
            }
        }
    }
}