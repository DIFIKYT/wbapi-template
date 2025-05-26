using Newtonsoft.Json;
using System.Text;
using System.Text.Json;

namespace APIWildberries
{
    public class WBApiHelper(HttpClient client, string salesUrl, string cardStatUrl, string companiesURL, string fullCompanyStatURL, string discountDataUrl, string feedbacksUrl,
        string warehouseRemainsUrl, string warehouseRemainsCreateReportEnding, string warehouseRemainsCheckStatusEnding, string warehouseRemainsDownloadEnding)
    {
        private readonly HttpClient _client = client;
        private readonly string _salesUrl = salesUrl;
        private readonly string _cardStatUrl = cardStatUrl;
        private readonly string _companiesURL = companiesURL;
        private readonly string _fullCompanyStatURL = fullCompanyStatURL;
        private readonly string _discountDataUrl = discountDataUrl;
        private readonly string _feedbacksUrl = feedbacksUrl;
        private readonly string _warehouseRemainsUrl = warehouseRemainsUrl;
        private readonly string _warehouseRemainsCreateReportEnding = warehouseRemainsCreateReportEnding;
        private readonly string _warehouseRemainsCheckStatusEnding = warehouseRemainsCheckStatusEnding;
        private readonly string _warehouseRemainsDownloadEnding = warehouseRemainsDownloadEnding;
        private string? _apiKey;

        public void GetApiKey(string apiKey)
        {
            _apiKey = apiKey;
        }

        public async Task GetSPP(Dictionary<int, SPPData> articuleToSPP, Dictionary<int, double> articuleToPrice)
        {
            if (articuleToPrice == null || articuleToPrice.Count == 0)
            {
                Console.WriteLine("Ошибка: Словарь articuleToPrice пуст или null");
                return;
            }

            HttpClientHandler handler = new();
            HttpClient client = new(handler);
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0");

            string url;
            double localPrice;
            double sitePrice;
            int article;
            int spp;

            foreach (var articleEntry in articuleToPrice)
            {
                article = articleEntry.Key;
                localPrice = articleEntry.Value;

                try
                {
                    url = $"https://card.wb.ru/cards/detail?appType=1&curr=rub&dest=-1257786&nm={article}";

                    HttpResponseMessage response = await client.GetAsync(url);

                    if (response.IsSuccessStatusCode == false)
                    {
                        continue;
                    }

                    string result = await response.Content.ReadAsStringAsync();
                    using JsonDocument document = JsonDocument.Parse(result);

                    JsonElement root = document.RootElement;

                    if (root.TryGetProperty("data", out JsonElement data) == false || data.TryGetProperty("products", out JsonElement products) == false || products.GetArrayLength() == 0 || products[0].TryGetProperty("salePriceU", out JsonElement salePriceU) == false)
                    {
                        articuleToSPP.Add(article, new SPPData(0, 0));
                        continue;
                    }

                    sitePrice = salePriceU.GetInt32() / 100.0;

                    spp = (int)Math.Round((1 - sitePrice / localPrice) * 100);

                    if (articuleToSPP.ContainsKey(article) == false)
                    {
                        articuleToSPP.Add(article, new SPPData(spp, localPrice));
                    }
                }
                catch (System.Text.Json.JsonException jsonEx)
                {
                    Console.WriteLine($"Ошибка: Парсинг JSON для артикула {article}: {jsonEx.Message}");
                }
                catch (HttpRequestException httpEx)
                {
                    Console.WriteLine($"Ошибка: HTTP запрос для артикула {article}: {httpEx.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: Неожиданная ошибка при обработке артикула {article}: {ex.Message}");
                }
            }
        }

        public async Task GetData(Dictionary<string, SalesData> salesData, Dictionary<string, StockData> stocksData, string apiKey, DateTime todayDate)
        {
            try
            {
                GetApiKey(apiKey);

                if (string.IsNullOrEmpty(_apiKey))
                    return;

                await GetSalesData(salesData, todayDate);
                await GetStocksData(stocksData);
            }
            catch (Exception exciption)
            {
                Console.WriteLine($"Ошибка получения основных данных по wb api {apiKey}:\n{exciption}");
            }
        }

        public async Task GetDiscountData(Dictionary<int, Good> articulesToGoods)
        {
            _client.DefaultRequestHeaders.Clear();
            _client.DefaultRequestHeaders.Add("Authorization", _apiKey);

            HttpResponseMessage response = await _client.GetAsync($"{_discountDataUrl}?limit=1000");
            string result = await response.Content.ReadAsStringAsync();
            DiscountResponse discount = JsonConvert.DeserializeObject<DiscountResponse>(result)!;

            if (response.IsSuccessStatusCode)
            {
                try
                {
                    foreach (var good in discount.Data!.Goods!)
                    {
                        articulesToGoods.Add(good.Articule, good);
                    }
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Ошибка запроса: {e.Message}");
                }
                catch (Exception e)
                {
                    Console.WriteLine($"Неожиданная ошибка: {e.Message}");
                }
            }
        }

        public async Task GetFullStat(Dictionary<int, ProductData> articulesToMarketingData, List<int> companyIds, DateTime todayDate)
        {
            try
            {
                _client.DefaultRequestHeaders.Clear();
                _client.DefaultRequestHeaders.Add("Authorization", _apiKey);

                var dates = new[]
                {
                $"{todayDate:yyyy-MM-dd}"
                };

                var requestBody = companyIds.Select(id => new
                {
                    id,
                    dates
                }).ToList();

                string jsonContent = JsonConvert.SerializeObject(requestBody);
                HttpContent content = new StringContent(jsonContent, Encoding.UTF8, "application/json");
                HttpResponseMessage response = await _client.PostAsync(_fullCompanyStatURL, content);
                string result = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    List<CampaignResponse> campaignResponses = JsonConvert.DeserializeObject<List<CampaignResponse>>(result)!;

                    foreach (CampaignResponse campaignResponse in campaignResponses)
                    {
                        foreach (Day day in campaignResponse.Days!)
                        {
                            foreach (App app in day.Apps!)
                            {
                                foreach (Nm nm in app.Nm!)
                                {
                                    int articule = nm.Articule;

                                    if (articulesToMarketingData.ContainsKey(articule) == false)
                                    {
                                        articulesToMarketingData[articule] = new ProductData(nm.Views, nm.Clicks, nm.Ctr, nm.Expenses);
                                    }
                                    else
                                    {
                                        articulesToMarketingData[articule].ViewsCount += nm.Views;
                                        articulesToMarketingData[articule].ClicksCount += nm.Clicks;
                                        articulesToMarketingData[articule].ClickConversion = nm.Ctr;
                                        articulesToMarketingData[articule].Expenses += nm.Expenses;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Ошибка:");
                    Console.WriteLine($"Status Code: {response.StatusCode}");
                    Console.WriteLine($"Error Content: {result}");
                }
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"Ошибка запроса: {e.Message}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Неожиданная ошибка: {e.Message}");
            }
        }

        public async Task<CompanyData?> GetCompanyData()
        {
            try
            {
                _client.DefaultRequestHeaders.Clear();
                _client.DefaultRequestHeaders.Add("Authorization", _apiKey);

                HttpResponseMessage response = await _client.GetAsync(_companiesURL);
                string result = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    CompanyData companyData = JsonConvert.DeserializeObject<CompanyData>(result)!;

                    return companyData;
                }
                else
                {
                    Console.WriteLine("Ошибка:");
                    Console.WriteLine($"Error Content: {result}");
                    return null;
                }
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"Ошибка запроса: {e.Message}");
                return null;
            }
            catch (Exception e)
            {
                Console.WriteLine($"Неожиданная ошибка: {e.Message}");
                return null;
            }
        }

        public async Task<double> GetRating(string articule, DateTime todayDate, string currentSpreadsheetId)
        {
            try
            {
                _client.DefaultRequestHeaders.Clear();
                _client.DefaultRequestHeaders.Add("Authorization", _apiKey);

                long unixDateFrom = new DateTimeOffset(todayDate).ToUnixTimeSeconds();
                long unixDateTo = new DateTimeOffset(todayDate.AddDays(1)).ToUnixTimeSeconds();

                HttpResponseMessage response = await _client.GetAsync(
                    $"{_feedbacksUrl}?isAnswered=true&nmId={articule}&take=5000&skip=0" +
                    $"&dateFrom={unixDateFrom}&dateTo={unixDateTo}"
                );

                if (response.IsSuccessStatusCode)
                {
                    double sum = 0;
                    int count = 0;

                    string result = await response.Content.ReadAsStringAsync();
                    JsonDocument jsonDocument = JsonDocument.Parse(result);
                    JsonElement data = jsonDocument.RootElement.GetProperty("data");
                    JsonElement feedbacks = data.GetProperty("feedbacks");

                    foreach (JsonElement feedback in feedbacks.EnumerateArray())
                    {
                        sum += feedback.GetProperty("productValuation").GetDouble();
                        count++;
                    }

                    return count > 0 ? sum / count : 0;
                }

                Console.Write("Не корректный ключ - ");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(currentSpreadsheetId);
                Console.ForegroundColor = ConsoleColor.White;

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка рейтинга у артикула {articule}");
                Console.WriteLine(ex);
                return 0;
            }
        }

        public async Task GetStatisticData(Dictionary<string, Card> urlToCard, DateTime todayDate)
        {
            try
            {
                _client.DefaultRequestHeaders.Clear();
                _client.DefaultRequestHeaders.Add("Authorization", _apiKey);

                var requestBody = new
                {
                    period = new
                    {
                        begin = $"{todayDate:yyyy-MM-dd} 00:00:00",
                        end = $"{todayDate:yyyy-MM-dd} 23:59:59"
                    },
                    page = 1
                };

                string jsonContent = System.Text.Json.JsonSerializer.Serialize(requestBody);
                HttpContent content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await _client.PostAsync(_cardStatUrl, content);
                string result = await response.Content.ReadAsStringAsync();

                if (response.IsSuccessStatusCode)
                {
                    StatisticData cardStatisticData = JsonConvert.DeserializeObject<StatisticData>(result)!;

                    foreach (Card card in cardStatisticData.Data!.Cards!)
                    {
                        urlToCard.Add($"{GenerateCardUrl(card.Articule)}", card);
                    }
                }
                else
                {
                    Console.WriteLine("Ошибка:");
                    Console.WriteLine($"Error Content: {result}");
                }
            }
            catch (HttpRequestException e)
            {
                Console.WriteLine($"Ошибка запроса: {e.Message}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Неожиданная ошибка: {e.Message}");
            }
        }

        public async Task GetSalesData(Dictionary<string, SalesData> salesData, DateTime todayDate)
        {
            _client.DefaultRequestHeaders.Clear();
            _client.DefaultRequestHeaders.Add("Authorization", _apiKey);

            HttpResponseMessage response = await _client.GetAsync($"{_salesUrl}?dateFrom={todayDate:yyyy-MM-dd}&flag=1");
            string result = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                try
                {
                    List<SalesData>? sales = JsonConvert.DeserializeObject<List<SalesData>>(result);

                    foreach (SalesData sale in sales!)
                    {
                        string saleKey = $"{sale.VendoreCode}_{sale.Barcode}";

                        if (salesData.TryGetValue(saleKey, out SalesData? saleValue))
                        {
                            if (sale.Barcode == saleValue.Barcode)
                            {
                                salesData[saleKey].Amount++;
                            }
                            else
                            {
                                salesData.Add(saleKey, sale);
                            }
                        }
                        else
                        {
                            salesData.Add(saleKey, sale);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Ошибка при разборе данных продажи: " + ex.Message);
                }
            }
            else
            {
                Console.WriteLine("Ошибка запроса: " + result);
            }
        }

        public async Task GetStocksData(Dictionary<string, StockData> stocksData)
        {
            _client.DefaultRequestHeaders.Clear();
            _client.DefaultRequestHeaders.Add("Authorization", _apiKey);

            HttpResponseMessage taskIdResponse = await _client.GetAsync(_warehouseRemainsUrl + "?" + _warehouseRemainsCreateReportEnding);
            string taskIdResult = await taskIdResponse.Content.ReadAsStringAsync();
            JsonDocument taskIdDocument = JsonDocument.Parse(taskIdResult);
            JsonElement taskIdData = taskIdDocument.RootElement.GetProperty("data");

            string warehouseRemainsUrl = _warehouseRemainsUrl + "/" + "tasks/" + taskIdData.GetProperty("taskId") + "/";

            while (true)
            {
                HttpResponseMessage statusResponse = await _client.GetAsync(warehouseRemainsUrl + _warehouseRemainsCheckStatusEnding);
                string statusResult = await statusResponse.Content.ReadAsStringAsync();
                JsonDocument statusDocument = JsonDocument.Parse(statusResult);
                JsonElement statusData = statusDocument.RootElement.GetProperty("data");

                if (statusData.GetProperty("status").ToString() == "done")
                {
                    break;
                }

                await Task.Delay(6000);
            }

            HttpResponseMessage stocksResponse = await _client.GetAsync(warehouseRemainsUrl + _warehouseRemainsDownloadEnding);
            stocksResponse.EnsureSuccessStatusCode();
            string stocksResult = await stocksResponse.Content.ReadAsStringAsync();

            List<StockData>? stocks = JsonConvert.DeserializeObject<List<StockData>>(stocksResult);

            foreach (StockData stock in stocks!)
            {
                stocksData.Add(stock.Barcode!, stock);
            }
        }

        public async Task GetTurnoverInDays(Dictionary<int, int> articulesToTurnoverInDays, DateTime todayDate)
        {
            _client.DefaultRequestHeaders.Clear();
            _client.DefaultRequestHeaders.Add("Authorization", _apiKey);

            var requestBody = new
            {
                currentPeriod = new
                {
                    start = $"{todayDate:yyyy-MM-dd}",
                    end = $"{todayDate.AddDays(1):yyyy-MM-dd}",
                },
                stockType = "",
                skipDeletedNm = true,
                orderBy = new
                {
                    field = "stockSum",
                    mode = "asc"
                },
                availabilityFilters = new[]
                {
                    "deficient",
                    "actual",
                    "balanced",
                    "nonActual",
                    "nonLiquid",
                    "invalidData"
                },
                offset = 0
            };

            string jsonContent = System.Text.Json.JsonSerializer.Serialize(requestBody);
            HttpContent content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

            HttpResponseMessage response = await _client.PostAsync("https://seller-analytics-api.wildberries.ru/api/v2/stocks-report/products/products", content);
            string result = await response.Content.ReadAsStringAsync();
            JsonDocument jsonDocument = JsonDocument.Parse(result);
            JsonElement data = jsonDocument.RootElement.GetProperty("data");
            JsonElement items = data.GetProperty("items");

            articulesToTurnoverInDays = items.EnumerateArray()
                .ToDictionary(
                    item => item.GetProperty("nmID").GetInt32(),
                    item => item.GetProperty("metrics").GetProperty("saleRate").GetProperty("days").GetInt32()
                );
        }

        private string GenerateCardUrl(int nm)
        {
            int vol = nm / 100000;
            int part = nm / 1000;

            string host = vol switch
            {
                <= 143 => "01",
                <= 287 => "02",
                <= 431 => "03",
                <= 719 => "04",
                <= 1007 => "05",
                <= 1061 => "06",
                <= 1115 => "07",
                <= 1169 => "08",
                <= 1313 => "09",
                <= 1601 => "10",
                <= 1655 => "11",
                <= 1919 => "12",
                <= 2045 => "13",
                <= 2189 => "14",
                <= 2405 => "15",
                <= 2621 => "16",
                <= 2837 => "17",
                <= 3053 => "18",
                <= 3269 => "19",
                <= 3485 => "20",
                <= 3701 => "21",
                <= 3845 => "22",
                >= 3845 => "23"
            };

            return $"https://basket-{host}.wbbasket.ru/vol{vol}/part{part}/{nm}/";
        }
    }
}