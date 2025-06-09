using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace UI_Testing
{
    public class GoogleSheetsHelper
    {
        private readonly SheetsService service;

        public GoogleSheetsHelper()
        {
            try
            {
                UserCredential credential;

                using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
                {
                    string tokenPath = "token"; // тут сохраняется refresh/access token

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.FromStream(stream).Secrets,
                        new[] { SheetsService.Scope.Spreadsheets },
                        "user", // имя пользователя или уникальный ID
                        CancellationToken.None,
                        new FileDataStore(tokenPath, true)).Result;
                }

                service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Jira Sheet Exporter"
                });
            }
            catch (Exception ex)
            {
                MaterialMessageBox.Show($"Ошибка авторизации: {ex.Message}");
            }
        }

        public string GetSpreadsheetId(string sheetUrl)
        {
            var match = Regex.Match(sheetUrl, @"/spreadsheets/d/([a-zA-Z0-9-_]+)");
            if (!match.Success)
                throw new ArgumentException("Invalid Google Sheet URL");
            return match.Groups[1].Value;
        }

        public string GetSheetNameByGid(string spreadsheetId, string gid)
        {
            var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
            foreach (var sheet in spreadsheet.Sheets)
            {
                if (sheet.Properties.SheetId.ToString() == gid)
                    return sheet.Properties.Title;
            }
            throw new Exception("Sheet with specified gid not found");
        }

        public string GetWorksheetByGid(string sheetUrl)
        {
            var spreadsheetId = GetSpreadsheetId(sheetUrl);
            var gidMatch = Regex.Match(sheetUrl, @"gid=(\d+)");
            if (!gidMatch.Success)
                throw new ArgumentException("No gid found in URL");

            string gid = gidMatch.Groups[1].Value;
            string sheetName = GetSheetNameByGid(spreadsheetId, gid);

            return $"{spreadsheetId}|{sheetName}";
        }
        public int GetWorksheetGid(string sheetUrl)
        {
            var spreadsheetId = GetSpreadsheetId(sheetUrl);
            var gidMatch = Regex.Match(sheetUrl, @"gid=(\d+)");
            if (!gidMatch.Success)
                throw new ArgumentException("No gid found in URL");

            int gid = int.Parse(gidMatch.Groups[1].Value);
            return gid;
        }
        public async Task<IList<IList<object>>> GetSheetData(string spreadsheetId, string range)
        {
            if (service == null)
                throw new InvalidOperationException("Sheets service is not initialized. Call InitializeAsync first.");

            var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
            var response = await request.ExecuteAsync();

            return response.Values ?? new List<IList<object>>();
        }

        public void AddRowsToSheet(string worksheetId, List<List<object>> rows)
        {
            var parts = worksheetId.Split('|');
            string spreadsheetId = parts[0];
            string sheetName = parts[1];

            var valueRange = new ValueRange
            {
                Values = rows.Cast<IList<object>>().ToList()
            };

            var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, $"{sheetName}!A1");
            appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
            appendRequest.Execute();
            
        }
        public int FindFirstRowWithNonWhiteBackground(string spreadsheetId, int sheetId, int maxRows = 10)
        {
            var whiteColor = new Color
            {
                Red = 1f,
                Green = 1f,
                Blue = 1f
            };

            var gridRange = new GridRange
            {
                SheetId = sheetId,
                StartRowIndex = 0,
                EndRowIndex = maxRows,
                StartColumnIndex = 3,
                EndColumnIndex = 4
            };

            // Запрашиваем данные с форматами
            var request = service.Spreadsheets.Get(spreadsheetId);
            request.Ranges = new List<string> { $"B1:B{maxRows}" };
            request.IncludeGridData = true;
            var response = request.Execute();

            var rows = response.Sheets.First().Data.First().RowData;
            if (rows == null) return -1;

            for (int i = 0; i < rows.Count; i++)
            {
                var cell = rows[i].Values?.FirstOrDefault();
                if (cell == null) continue;

                var cellColor = cell.EffectiveFormat?.BackgroundColor;
                if (cellColor == null) continue;

                if (!AreColorsEqual(cellColor, whiteColor))
                    return i; // индекс строки, начиная с 0
            }

            return -1; // если не нашли отличий
        }

        // Сравнение цветов с погрешностью
        private bool AreColorsEqual(Color c1, Color c2, float epsilon = 0.01f)
        {
            float r1 = c1.Red ?? 0f;
            float g1 = c1.Green ?? 0f;
            float b1 = c1.Blue ?? 0f;

            float r2 = c2.Red ?? 0f;
            float g2 = c2.Green ?? 0f;
            float b2 = c2.Blue ?? 0f;

            return Math.Abs(r1 - r2) < epsilon &&
                   Math.Abs(g1 - g2) < epsilon &&
                   Math.Abs(b1 - b2) < epsilon;
        }
        public void CopyPasteDataValidation(
        string spreadsheetId,
        int sourceSheetId,
        int destinationSheetId,
        int sourceStartRow,
        int sourceEndRow,
        int sourceStartColumn,
        int sourceEndColumn,
        int destinationStartRow,
        int destinationEndRow)
        {
            try
            {
                var sourceRange = new GridRange
                {
                    SheetId = sourceSheetId,
                    StartRowIndex = sourceStartRow,
                    EndRowIndex = sourceEndRow,
                    StartColumnIndex = sourceStartColumn,
                    EndColumnIndex = sourceEndColumn
                };

                var destinationRange = new GridRange
                {
                    SheetId = destinationSheetId,
                    StartRowIndex = destinationStartRow,
                    EndRowIndex = destinationEndRow,
                    StartColumnIndex = sourceStartColumn,
                    EndColumnIndex = sourceEndColumn
                };

                var requests = new List<Request>();
                // 1. Копирование валидации
                requests.Add(new Request
                {
                    CopyPaste = new CopyPasteRequest
                    {
                        Source = sourceRange,
                        Destination = destinationRange,
                        PasteType = "PASTE_DATA_VALIDATION"
                    }
                });

                // 2. Копирование формата
                requests.Add(new Request
                {
                    CopyPaste = new CopyPasteRequest
                    {
                        Source = sourceRange,
                        Destination = destinationRange,
                        PasteType = "PASTE_FORMAT"
                    }
                });
                // 1. Установить ширину колонки D (индекс 3) в 400px
                requests.Add(new Request
                {
                    UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                    {
                        Range = new DimensionRange
                        {
                            SheetId = destinationSheetId, // Укажи правильный ID листа
                            Dimension = "COLUMNS",
                            StartIndex = 3,
                            EndIndex = 4
                        },
                        Properties = new DimensionProperties
                        {
                            PixelSize = 400
                        },
                        Fields = "pixelSize"
                    }
                });
                requests.Add(new Request
                {
                    UpdateDimensionProperties = new UpdateDimensionPropertiesRequest
                    {
                        Range = new DimensionRange
                        {
                            SheetId = destinationSheetId, // Укажи правильный ID листа
                            Dimension = "COLUMNS",
                            StartIndex = 2,
                            EndIndex = 3
                        },
                        Properties = new DimensionProperties
                        {
                            PixelSize = 130
                        },
                        Fields = "pixelSize"
                    }
                });
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = requests
                };

                service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).Execute();
            }
            catch (Exception ex) { MaterialMessageBox.Show(ex.Message); }
        }

        public void ClearColors(string spreadsheetId,
        int destinationSheetId,
        int sourceStartColumn,
        int sourceEndColumn,
        int destinationStartRow,
        int destinationEndRow) 
        {
            try
            {
                var destinationRange = new GridRange
                {
                    SheetId = destinationSheetId,
                    StartRowIndex = destinationStartRow,
                    EndRowIndex = destinationEndRow,
                    StartColumnIndex = sourceStartColumn,
                    EndColumnIndex = sourceEndColumn
                };
                var requests = new List<Request>();

                // 3. Очистка цвета (заливки)
                requests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = destinationRange,
                        Cell = new CellData
                        {
                            UserEnteredFormat = new CellFormat
                            {
                                BackgroundColor = new Color
                                {
                                    Red = 1f,
                                    Green = 1f,
                                    Blue = 1f
                                }
                            }
                        },
                        Fields = "userEnteredFormat.backgroundColor"
                    }
                });
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = requests
                };

                service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).Execute();
            }
            catch (Exception ex) { MaterialMessageBox.Show(ex.Message); }

        }
    }
}
