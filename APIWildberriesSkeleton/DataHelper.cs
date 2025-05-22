using Google.Apis.Sheets.v4.Data;
using System.Text;

namespace APIWildberries
{
    public class DataHelper
    {
        public readonly static int PercentBase = 100;

        public static double CalculateAverageData(int daysCount, List<int> values)
        {
            int sum;

            if (values.Count < daysCount)
            {
                sum = values.Sum();
            }
            else
            {
                List<int> subset = values.TakeLast(daysCount).ToList();
                sum = subset.Sum();
            }

            double average = (double)sum / daysCount;

            return Math.Round(average, 1, MidpointRounding.AwayFromZero);
        }

        public static Border CreateBorder()
        {
            return new Border
            {
                Style = "SOLID",
                Width = 1,
                Color = new Color
                {
                    Red = 0.0f,
                    Green = 0.0f,
                    Blue = 0.0f
                }
            };
        }

        public static Request GetAlignmentRequest(int? sheetId, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            return new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = startRowIndex,
                        EndRowIndex = endRowIndex,
                        StartColumnIndex = startColumnIndex,
                        EndColumnIndex = endColumnIndex
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            VerticalAlignment = "TOP"
                        }
                    },
                    Fields = "userEnteredFormat.verticalAlignment"
                }
            };
        }

        public static Request GetBordersRequest(int? sheetId, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            return new()
            {
                UpdateBorders = new UpdateBordersRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = startRowIndex,
                        EndRowIndex = endRowIndex,
                        StartColumnIndex = startColumnIndex,
                        EndColumnIndex = endColumnIndex
                    },
                    Top = CreateBorder(),
                    Bottom = CreateBorder(),
                    Left = CreateBorder(),
                    Right = CreateBorder(),
                    InnerHorizontal = CreateBorder(),
                    InnerVertical = CreateBorder()
                }
            };
        }

        public static Request GetColorCellRequest(int? sheetId, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex, Color color)
        {
            return new Request
            {
                RepeatCell = new RepeatCellRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = startRowIndex,
                        EndRowIndex = endRowIndex,
                        StartColumnIndex = startColumnIndex,
                        EndColumnIndex = endColumnIndex
                    },
                    Cell = new CellData
                    {
                        UserEnteredFormat = new CellFormat
                        {
                            BackgroundColor = color
                        }
                    },
                    Fields = "userEnteredFormat.backgroundColor"
                }
            };
        }

        public static int GetColumnIndex(string columnLetter)
        {
            int columnIndex = 0;
            columnLetter = columnLetter.ToUpper();

            for (int i = 0; i < columnLetter.Length; i++)
            {
                char letter = columnLetter[i];
                columnIndex = columnIndex * 26 + (letter - 'A' + 1);
            }

            return columnIndex;
        }

        public static string GetColumnName(int columnIndex)
        {
            StringBuilder columnName = new();

            while (columnIndex > 0)
            {
                int remainder = (columnIndex - 1) % 26;
                char letter = (char)('A' + remainder);
                columnName.Insert(0, letter);
                columnIndex = (columnIndex - 1) / 26;
            }

            return columnName.ToString();
        }

        public static double GetAverageData(IList<IList<object>> values, int rowNumber, bool withZeroValue = false)
        {
            rowNumber--;

            if (values == null || rowNumber < 0 || rowNumber >= values.Count)
            {
                Console.WriteLine($"Ошибка: Некорректный rowNumber {rowNumber} или values пуст.");
                return 0;
            }

            if (values[rowNumber] == null || values[rowNumber].Count == 0)
                return 0;

            List<double> nonZeroValues = [];

            if (withZeroValue)
            {
                nonZeroValues = values[rowNumber]
                    .Where(value => value != null && double.TryParse(value.ToString(), out double num))
                    .Select(value => Convert.ToDouble(value))
                    .ToList();
            }
            else
            {
                nonZeroValues = values[rowNumber]
                    .Where(value => value != null && double.TryParse(value.ToString(), out double num) && num != 0)
                    .Select(value => Convert.ToDouble(value))
                    .ToList();
            }

            return nonZeroValues.Count != 0 ? nonZeroValues.Average() : 0;
        }

        public static double GetSumData(IList<IList<object>> values, int rowNumber)
        {
            if (values == null || values.Count == 0 || rowNumber <= 0 || rowNumber > values.Count)
                return 0;

            int rowIndex = rowNumber - 1;
            IList<object> row = values[rowIndex];

            if (row == null || row.Count == 0)
                return 0;

            return row.Where(value => value != null && double.TryParse(value.ToString(), out double num) && num != 0)
                      .Select(value => Convert.ToDouble(value))
                      .Sum();
        }

        public static double GetLastValue(IList<IList<object>> values, int rowNumber)
        {
            if (values.Count == 0 || values.Count < --rowNumber || values[rowNumber].Count < 0 || values == null || values[rowNumber].Count == 0 || values[rowNumber] == null)
                return 0;
            else
                return double.TryParse(values[rowNumber].LastOrDefault()!.ToString(), out double value) ? value : 0;
        }

        public static int GetMaxDataCount(ValueRange valueRange, int startColumnIndex)
        {
            if (valueRange.Values == null || valueRange.Values.Count == 0)
            {
                Console.WriteLine("Таблица пуста.");
                return -1;
            }

            if (valueRange.Values.Count < startColumnIndex)
            {
                return startColumnIndex + 1;
            }

            return valueRange.Values
                .Where((_, index) => index >= 3 && (index - 3) % 23 == 0)
                .Max(row => row?.Count ?? 0);
        }

        public static Request GetMergeCellsRequest(int? sheetId, int startRowIndex, int endRowIndex, int startColumnIndex, int endColumnIndex)
        {
            return new()
            {
                MergeCells = new MergeCellsRequest
                {
                    Range = new GridRange
                    {
                        SheetId = sheetId,
                        StartRowIndex = startRowIndex,
                        EndRowIndex = endRowIndex,
                        StartColumnIndex = startColumnIndex,
                        EndColumnIndex = endColumnIndex
                    },

                    MergeType = "MERGE_ALL"
                }
            };
        }

        public static int? GetSheetId(Spreadsheet spreadSheetId, string listName)
        {
            foreach (Sheet sheet in spreadSheetId.Sheets)
            {
                if (sheet.Properties.Title == listName)
                {
                    return sheet.Properties.SheetId;
                }
            }

            return -1;
        }

        public static double GetSumWeeklyData(int rowNumber, List<ValueRange> weaklyColumns)
        {
            double sum = 0;

            foreach (ValueRange valueRange in weaklyColumns)
            {
                try
                {
                    if (valueRange.Values != null &&
                        rowNumber - 1 < valueRange.Values.Count &&
                        valueRange.Values[rowNumber - 1].Count > 0)
                    {
                        if (double.TryParse(valueRange.Values[rowNumber - 1][0]?.ToString(), out double num))
                        {
                            sum += num;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing row {rowNumber}: {ex.Message}");
                }
            }

            return sum;
        }

        public static List<IList<object>> MergeRows(IList<IList<object>> rows1, IList<IList<object>> rows2)
        {
            var mergedRows = new List<IList<object>>();

            int maxRows = Math.Max(rows1?.Count ?? 0, rows2?.Count ?? 0);

            for (int i = 0; i < maxRows; i++)
            {
                var row1 = i < (rows1?.Count ?? 0) ? rows1![i] : [];
                var row2 = i < (rows2?.Count ?? 0) ? rows2![i] : [];

                var mergedRow = row2.Concat(row1).ToList();

                mergedRows.Add(mergedRow);
            }

            return mergedRows;
        }

        public static double CalculateMargin(DataForMargin dataForMargin, int ordersSum, double ordersSumRub, double expanses, double taxPercentage)
        {
            return ordersSumRub -
                (ordersSumRub * (dataForMargin.CommissionPercent / PercentBase))
                -
                (ordersSumRub * (dataForMargin.AcquiringPercent / PercentBase))
                -
                (ordersSum * dataForMargin.LogisticsRub)
                -
                (ordersSum * dataForMargin.KeepingRub)
                -
                (ordersSum * dataForMargin.AcceptanceRub)
                -
                (ordersSum * (dataForMargin.PurchasePrice + dataForMargin.FullfilmentProcessingRub))
                -
                (ordersSumRub * (taxPercentage / PercentBase))
                -
                expanses;
        }
    }
}