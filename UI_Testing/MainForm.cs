using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web.UI;
using System.Windows.Forms;
using static UI_Testing.JiraClient;

namespace UI_Testing
{
    public partial class MainForm : MaterialForm
    {
        private string? login;
        private string? password;
        public List<List<object>> rows;
        private ShowError showError = new ShowError();
        public GoogleSheetsHelper helper = new GoogleSheetsHelper();
        private readonly MaterialSkinManager materialSkinManager;
        public MainForm(string? _login, string? _password)
        {
            InitializeComponent();
            // Инициализация менеджера скинов
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT; // или DARK
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Blue600, Primary.Blue700,
                Primary.Blue200, Accent.LightBlue200,
                TextShade.WHITE
            );
            login = _login;
            password = _password;
            toolStripMenuItem1.Enabled = false;
            materialLabel4.Visible = false;
            tableLayoutPanel1.Visible = true;
            tableLayoutPanel2.Visible = false;
            tableLayoutPanel3.Visible = false;
        }
        public void SetCredentials(string login, string password)
        {
            this.login = login;
            this.password = password;
        }
        private async Task<List<List<object>>> LoadRowsFromJira(string version)
        {
            var jira = new JiraClient(login, password);

            var issues = await jira.GetIssuesByVersion(version);
            var processedCases = new HashSet<string>();
            var rows = new List<List<object>>();
            int taskNumber = 1;

            foreach (var issue in issues)
            {
                var issueData = await jira.GetIssueData(issue.Key);
                if (issueData == null) continue;
                ProcessIssueAndCases(jira, issueData, taskNumber, processedCases, rows);
                taskNumber++;
            }

            return rows;
        }
        private async void materialButton1_Click(object sender, EventArgs e)
        {
            if (!toolStripMenuItem1.Enabled)
            {
                string version = textBoxVersion.Text;
                if (version == "") { MaterialMessageBox.Show("Заполните поле \"Скоуп\""); return; }
                if (!materialCheckbox1.Checked)
                {
                    rows = await LoadRowsFromJira(version);
                }
                else 
                {
                    rows = await ExportTestCasesFromUrl(version);
                }
                if (checkBoxPreview.Checked)
                {
                    if (materialCheckbox.Checked)
                    {
                        FillPreviewTable(rows, true);
                    }
                    else
                    {
                        FillPreviewTable(rows, false);
                    }
                }
            }
            else 
            {
                string issueKey = ExtractIssueKey(textBoxVersion.Text.Trim());
                if (string.IsNullOrEmpty(issueKey))
                {
                    MessageBox.Show("Не удалось извлечь ключ задачи из ссылки.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var jira = new JiraClient(login, password);
                var html = await jira.GetIssueHtml(issueKey);
                if (string.IsNullOrWhiteSpace(html))
                {
                    MessageBox.Show("Не удалось получить HTML содержимое задачи. Проверь авторизацию или ссылку.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var (productName, scopeUrl, zni) = ExtractReportHeader(html);
                string gdUrl = textBoxSheetUrl.Text.Trim();

                var urls = ExtractUrlsFromDescription(html);
                var urls2 = ExtractCycleUrls(html);
                if (urls == null)
                {
                    MessageBox.Show("Не удалось найти ссылку на регрессионный цикл.");
                    return;
                }

                
                var regressionStats = await jira.GetCycleTestStatsForCycles(urls);

                string sheetUrl = textBoxSheetUrl.Text.Trim();

                if (string.IsNullOrEmpty(sheetUrl))
                {
                    MaterialMessageBox.Show("Введите URL Google таблицы.");
                    return;
                }
                string sourcesheetID = helper.GetSpreadsheetId(sheetUrl);

                // 3. Имя листа (можно сделать вводом в отдельное поле, сейчас хардкод)
                string worksheetId = helper.GetWorksheetByGid(sheetUrl);
                var parts = worksheetId.Split('|');
                string spreadsheetId = parts[0];
                string sheetName = parts[1];

                var data = await helper.GetSheetData(sourcesheetID, $"{sheetName}!A1:Z1000");

                var stats = AnalyzeSheetData(data);
                var date = ExtractIsoDateFromHtml(html);
                var (foundUrl, knownUrl) = BuildDefectUrls(issueKey, date);
                textBox1.Text = FormatReport(regressionStats, productName, scopeUrl, zni, urls2, gdUrl, stats, foundUrl, knownUrl);
            }
        }
        private (string foundBugsUrl, string knownBugsUrl) BuildDefectUrls(string issueKey, string date)
        {
            string baseUrl = "https://jira.mos.social/issues/?jql=";
            string commonPart = $"issuetype in (Bug, Improvement, Task) AND issue in linkedIssues({issueKey}) ORDER BY priority DESC";

            string foundUrl = baseUrl + Uri.EscapeDataString($"issuetype in (Bug, Improvement, Task) AND createdDate >= {date} AND issue in linkedIssues({issueKey}) ORDER BY priority DESC");
            string knownUrl = baseUrl + Uri.EscapeDataString($"issuetype in (Bug, Improvement, Task) AND createdDate <= {date} AND issue in linkedIssues({issueKey}) ORDER BY priority DESC");

            return (foundUrl, knownUrl);
        }
        private string ExtractIsoDateFromHtml(string html)
        {
            var match = Regex.Match(html, @"<dd class=""date user-tz""\s+title=""(\d{2})\.(\d{2})\.(\d{2})\s+\d{2}:\d{2}""");
            if (match.Success)
            {
                var year = "20" + match.Groups[3].Value;
                var month = match.Groups[2].Value;
                var day = match.Groups[1].Value;
                return $"{year}-{month}-{day}";
            }
            return null;
        }
        private List<string> ExtractUrlsFromDescription(string description)
        {
            var matches = Regex.Matches(description, "<a href=\"(https://jira\\.mos\\.social/secure/enav/[^\"]+)\"[^>]*>.*?(РЦ|Регрессионный цикл|Регрессионный ЦТ|Регресс).*?</a>");
            return matches.Cast<Match>().Select(m => m.Groups[1].Value).Distinct().ToList();
        }
        private List<string> ExtractCycleUrls(string description)
        {
            var matches = Regex.Matches(description,
                "<a href=\"(https://jira\\.mos\\.social/secure/enav/[^\"]+)\"[^>]*>.*?(ФЦ|Функциональный цикл|Функ\\.цикл).*?</a>",
                RegexOptions.IgnoreCase);

            return matches.Cast<Match>()
                          .Select(m => m.Groups[1].Value)
                          .Distinct()
                          .ToList();
        }
        private string ExtractIssueKey(string url)
        {
            var match = Regex.Match(url, @"browse/([A-Z]+-\d+)");
            return match.Success ? match.Groups[1].Value : null;
        }
        private (string productName, string scopeUrl, string zni) ExtractReportHeader(string html)
        {
            string productName = "";
            string scopeUrl = "";
            string zni = "";
            // Ищем название продукта внутри <b>Требуется провести тестирование релиза ...</b>
            // Вариант, где "продукта" может быть, а может и не быть
            var productMatch = Regex.Match(html,
                @"Требуется провести тестирование релиза\s+([\d\.]+\s*(?:\([^)]+\))?)\s*(?:продукта\s+)?(.+?)</b>",
                RegexOptions.IgnoreCase);

            if (productMatch.Success)
            {
                string version = productMatch.Groups[1].Value.Trim();
                string product = productMatch.Groups[2].Value.Trim();
                productName = $"{version} {product}";
            }

            // Ищем ссылку после "Состав:"
            var scopeMatch = Regex.Match(html, @"(?:Состав|Скоуп|Скоуп релиза)\s*:\s*<a[^>]+href\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase);
            if (scopeMatch.Success)
            {
                scopeUrl = scopeMatch.Groups[1].Value.Trim();
            }
            var zniMatch = Regex.Match(html, @"(?:ZNI|ЗНИ)\s*:\s*<a[^>]+href\s*=\s*""([^""]+)""", RegexOptions.IgnoreCase);
            if (zniMatch.Success)
            {
                zni = zniMatch.Groups[1].Value.Trim();
            }
            return (productName, scopeUrl, zni);
        }
        private string FormatReport(
            List<CycleTestStats> regressionStats,
            string productName,
            string scopeUrl,string zni, List<string> urls,
            string gdUrl, SheetStats sheetStats, string foundUrl, string knownUrl)
        {
            var sb = new StringBuilder();

            sb.AppendLine($"Отчет о тестировании продукта: {productName}");
            sb.AppendLine($"ЗНИ: {zni}");
            sb.AppendLine($"Скоуп: {scopeUrl}");
            sb.AppendLine($"Прогресс тестирования зафиксирован в ГД: {gdUrl}");
            sb.AppendLine();
            
            sb.AppendLine($"Всего задач: {sheetStats.TotalTasks}");
            sb.AppendLine($"Закрыто: {sheetStats.ClosedTasks}");
            sb.AppendLine($"Переоткрытых задач: {sheetStats.ReopenedTasks}");

            if (sheetStats.UntestedTaskLinks.Count > 0)
            {
                sb.AppendLine("Не протестированных задач:");
                foreach (var link in sheetStats.UntestedTaskLinks)
                    sb.AppendLine(link);
            }
            else {
                sb.AppendLine("Не протестированных задач: 0");
            }
            sb.AppendLine();

            if (urls.Count > 0)
            {
                if (urls.Count == 1)
                {
                    sb.AppendLine($"Функциональное тестирование в цикле: {urls[0]}");
                }
                else
                {
                    sb.AppendLine("Функциональное тестирование в циклах:");
                    foreach (var url in urls)
                    {
                        sb.AppendLine($"{url}");
                    }
                }
            }
            sb.AppendLine();
            sb.AppendLine($"Количество запланированных тест-кейсов: {sheetStats.PlannedTestCases}");
            sb.AppendLine();
            sb.AppendLine($"Успешно: {sheetStats.Passed}");
            sb.AppendLine($"Успешно с багом: {sheetStats.PassedWithBug}");
            sb.AppendLine($"Провалено: {sheetStats.Failed}");
            sb.AppendLine($"Проведено ревью тест-кейсов: {sheetStats.ReviewedTestCases}");
            sb.AppendLine();
            foreach (var stats in regressionStats)
            {
                sb.AppendLine($"Регрессионное тестирование в цикле: {stats.Url}");
                sb.AppendLine();
                sb.AppendLine($"Всего тест-кейсов: {stats.Total}");
                sb.AppendLine($"Успешно пройденных тест-кейсов: {stats.Passed}");
                sb.AppendLine($"Проваленных тест-кейсов: {stats.Failed}");
                sb.AppendLine($"Пройденных с ошибкой тест-кейсов: {stats.Errored}");
                sb.AppendLine($"Заблокированных тест-кейсов: {stats.Blocked}");
                sb.AppendLine();
            }
            sb.AppendLine($"Найденные дефекты: {foundUrl}");
            sb.AppendLine();
            sb.AppendLine($"Известные дефекты: {knownUrl}");
            return sb.ToString();
        }
        private void ProcessIssueAndCases(JiraClient jira,JiraIssue issue,int taskNumber,HashSet<string> processedCases, List<List<object>> rows)
        {
            string issueKey = issue.Key;
            string link = $"=ГИПЕРССЫЛКА(\"{JiraClient.BaseUrl}/browse/{issueKey}\"; \"{issueKey}\")";

            var fields = issue.Fields;
            string status = fields.Status == "Waiting for release" ? "Waiting for Release" : fields.Status;
            if (checkBoxIteration.Checked && (status == "Waiting for Release" || status == "Closed")) { return; }
            var testCases = jira.GetLinkedTestCases(issue);
            string test_count = (testCases == null || !testCases.Any()) ? "1" : "";
            var row = new List<object>
            {
                taskNumber.ToString(),
                fields.Type,
                link,
                fields.Summary,
                test_count, "", "", // Кол-во тестов, шагов, ревью ТК
                fields.PriorityValue,
                fields.Type != "Test" ? status : "",
                "", "" // Статус проверки, тестировщик
            };
            rows.Add(row);

            
            foreach (var test in testCases)
            {
                string testLink = $"=ГИПЕРССЫЛКА(\"{JiraClient.BaseUrl}/browse/{test.Key}\"; \"{test.Key}\")";
                bool isNew = processedCases.Add(test.Key);

                var caseRow = new List<object>
                {
                    "", "Test case", testLink, test.Fields.Summary,
                    isNew ? "1" : "0", "", // Кол-во тестов, шагов
                    test.Fields.Status == "In Review" ? "На ревью" : "",
                    test.Fields.PriorityValue, "", "", ""
                };
                rows.Add(caseRow);
            }
        }
        public async Task<List<List<object>>> ExportTestCasesFromUrl(string url)
        {
            var jira = new JiraClient(login, password);
            var parsed = jira.ParseZqlUrl(url);

            var testCases = await jira.GetTestCasesFromCycle(
                parsed.CycleName,
                parsed.Version == "Незапланированные" ? "Unscheduled" : parsed.Version,
                parsed.Project,
                parsed.FolderNames,
                parsed.ExecutionStatus
            );

            var rows = new List<List<object>>();
            int number = 1;

            foreach (var test in testCases)
            {
                string testLink = $"=ГИПЕРССЫЛКА(\"{JiraClient.BaseUrl}/browse/{test.Key}\"; \"{test.Key}\")";

                var row = new List<object>
                {
                    number.ToString(), // ← Номер строки
                    "Test case",
                    testLink,
                    test.Fields.Summary,
                    "1", "", // Кол-во тестов, шагов
                    "", // статус
                    test.Fields.PriorityValue,
                    "", "", "" // Статус проверки, тестировщик и т.д.
                };

                rows.Add(row);
                number++; // ← Увеличиваем номер
            }

            return rows;
        }
        private void FillPreviewTable(List<List<object>> rows, bool showPriority)
        {
            dataGridViewPreview.Columns.Clear();
            dataGridViewPreview.Rows.Clear();

            // Формируем список заголовков в зависимости от showPriority
            List<string> headers;

            if (showPriority)
            {
                headers = new List<string>
                {
                    "№", "Тип", "ТК", "Название", "Кол-во тестов",
                    "Кол-во шагов", "Ревью ТК", "Приоритет", "Статус задачи", "Статус проверки", "Тестировщик"
                };
            }
            else
            {
                // Приоритет удаляем, добавляем колонку "Комментарий"
                headers = new List<string>
                {
                    "№", "Тип", "ТК", "Название", "Кол-во тестов",
                    "Кол-во шагов", "Ревью ТК","Статус ТК", "Тестировщик", "Комментарий", "Статус задачи"
                };
            }

            // Добавляем колонки
            foreach (var header in headers)
            {
                var column = new DataGridViewTextBoxColumn
                {
                    HeaderText = header,
                    Name = header,
                    ReadOnly = false
                };
                dataGridViewPreview.Columns.Add(column);
            }

            // Добавляем строки
            foreach (var row in rows)
            {
                string[] cells;

                if (showPriority)
                {
                    // row должен содержать значения для всех колонок приоритетного варианта
                    cells = row.Select(cell => cell?.ToString() ?? "").ToArray();
                }
                else
                {
                    // Без "Приоритета" и с "Комментариями"

                    cells = new string[]
                    {
                        row.ElementAtOrDefault(0)?.ToString() ?? "",
                        row.ElementAtOrDefault(1)?.ToString() ?? "",
                        row.ElementAtOrDefault(2)?.ToString() ?? "",
                        row.ElementAtOrDefault(3)?.ToString() ?? "",
                        row.ElementAtOrDefault(4)?.ToString() ?? "",
                        row.ElementAtOrDefault(5)?.ToString() ?? "",
                        row.ElementAtOrDefault(6)?.ToString() ?? "",  // Статус задачи
                        row.ElementAtOrDefault(9)?.ToString() ?? "",  // Ревью ТК
                        row.ElementAtOrDefault(10)?.ToString() ?? "", // Тестировщик
                        "",                                          // Комментарий — пусто, можно редактировать
                        row.ElementAtOrDefault(8)?.ToString() ?? ""  // Статус проверки
                    };
                }

                dataGridViewPreview.Rows.Add(cells);
            }

            // Настройки DataGridView
            dataGridViewPreview.AllowUserToAddRows = true;
            dataGridViewPreview.AllowUserToDeleteRows = true;
            dataGridViewPreview.AllowUserToOrderColumns = true;
            dataGridViewPreview.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            dataGridViewPreview.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridViewPreview.MultiSelect = true;
            dataGridViewPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            foreach (DataGridViewColumn column in dataGridViewPreview.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void checkBoxNoStyle_CheckedChanged(object sender, EventArgs e)
        {
            textBoxStyleSheetUrl.Enabled = !checkBoxNoStyle.Checked;
        }

        private void materialCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            bool showPriority = materialCheckbox.Checked;
            if(rows != null && checkBoxPreview.Checked) { 
                FillPreviewTable(rows, showPriority); // currentRows — твои данные, которые ты передаёшь
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            materialLabel1.Visible = true;
            materialLabel2.Visible = true;
            materialLabel3.Visible = true;
            materialLabel3.Text = "Скоуп";
            textBoxSheetUrl.Visible = true;
            textBoxStyleSheetUrl.Visible = true;
            textBoxVersion.Visible = true;
            materialButton1.Visible = true;
            checkBoxIteration.Visible = true;
            checkBoxNoStyle.Visible = true;
            checkBoxPreview.Visible = true;
            materialCheckbox.Visible = true;
            dataGridViewPreview.Visible = true;
            toolStripMenuItem2.Enabled = true;
            toolStripMenuItem1.Enabled = false;
            toolStripMenuItem3.Enabled = true;
            tableLayoutPanel3.Visible = false;
            materialLabel4.Visible = false;
            checkBoxPreview.Enabled = true;
            tableLayoutPanel2.Visible = false;
            materialCheckbox1.Visible = true;
            tableLayoutPanel1.Visible = true;
            textBox1.Visible = false;
            textBox1.Enabled = false;
            materialButton2.Visible = true;
            materialButton2.Enabled = true;
            this.Text = "ГД";
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            materialLabel1.Visible = true;
            materialLabel2.Visible = false;
            materialLabel3.Visible = true;
            materialLabel3.Text = "Задача";
            textBoxSheetUrl.Visible = true;
            textBoxStyleSheetUrl.Visible = false;
            textBoxVersion.Visible = true;
            materialButton1.Visible = true;
            checkBoxIteration.Visible = false;
            checkBoxNoStyle.Visible = false;
            checkBoxPreview.Visible = true;
            tableLayoutPanel2.Visible = false;
            checkBoxPreview.Enabled = false;
            materialCheckbox.Visible = false;
            dataGridViewPreview.Visible = false;
            toolStripMenuItem2.Enabled = false;
            materialCheckbox1.Visible = false;
            tableLayoutPanel1.Controls.Add(textBox1, 0, 4);
            tableLayoutPanel1.SetColumnSpan(textBox1, 4);
            textBox1.Visible = true;
            textBox1.Enabled = true;
            materialButton2.Visible = false;
            materialButton2.Enabled = false;
            toolStripMenuItem1.Enabled = true;
            toolStripMenuItem3.Enabled = true;
            materialLabel4.Visible = true;
            tableLayoutPanel3.Visible = false;
            tableLayoutPanel1.Visible = true;
            this.Text = "Отчет";
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            tableLayoutPanel3.Visible = false;
            tableLayoutPanel1.Visible = false;
            tableLayoutPanel2.Visible = true;
            checkBoxNoStyle.Visible = false;
            checkBoxPreview.Visible = false;
            materialCheckbox.Visible = false;
            dataGridViewPreview.Visible = false;
            materialCheckbox1.Visible = false;
            toolStripMenuItem2.Enabled = true;
            toolStripMenuItem1.Enabled = true;
            toolStripMenuItem3.Enabled = false;
            materialLabel4.Visible = true;
            materialButton2.Visible = false;
            materialButton2.Enabled = false;
            checkBoxPreview.Enabled = false;
            textBox1.Visible= false;
            textBox1.Enabled= false;
            materialLabel4.Text = "Спасибо что поинтересовались, функционал \"Поиск\" еще не реализован :(";
            this.Text = "Поиск";
        }

        private async void materialButton2_Click(object sender, EventArgs e)
        {
            try
            {
                string version = textBoxVersion.Text;
                if (version == "") { MaterialMessageBox.Show("Заполните поле \"Скоуп\""); return; }
                if (!materialCheckbox1.Checked)
                {
                    rows = await LoadRowsFromJira(version);
                }
                else
                {
                    rows = await ExportTestCasesFromUrl(version);
                }

                // 1. Получаем URL таблицы из текстбокса
                string sheetUrl = textBoxSheetUrl.Text.Trim();
                if (string.IsNullOrEmpty(sheetUrl))
                {
                    MaterialMessageBox.Show("Введите URL Google таблицы.");
                    return;
                }

                // 2. Извлекаем ID таблицы
                string spreadsheetId = helper.GetSpreadsheetId(sheetUrl);

                // 3. Имя листа (можно сделать вводом в отдельное поле, сейчас хардкод)
                string worksheetId = helper.GetWorksheetByGid(sheetUrl);

                // 4. Формируем строки
                var lisrRows = GetRowsForExport();
                // 5. Добавляем строки
                helper.AddRowsToSheet($"{worksheetId}", lisrRows);

                string sourceUrl = textBoxStyleSheetUrl.Text.Trim();
                if (string.IsNullOrEmpty(sourceUrl))
                {
                    MaterialMessageBox.Show("Введите URL Google таблицы.");
                    return;
                }
                if (!checkBoxNoStyle.Checked)
                {
                    string sourcesheetID = helper.GetSpreadsheetId(sourceUrl);
                    int sourcesheetId = helper.GetWorksheetGid(sourceUrl);
                    int spreadId = helper.GetWorksheetGid(sheetUrl);
                    int endCol = lisrRows[0].Count;
                    int numRows = lisrRows.Count;
                    helper.CopyPasteDataValidation(spreadsheetId, sourcesheetId, spreadId, 2, 3, 0, endCol, 1, numRows);
                    helper.ClearColors(spreadsheetId, spreadId, 0, endCol, 1, numRows);
                    int i = helper.FindFirstRowWithNonWhiteBackground(sourcesheetID, sourcesheetId);
                    if (i == -1) { i = 0; }
                    helper.CopyPasteDataValidation(spreadsheetId, sourcesheetId, spreadId, i, i + 1, 0, endCol, 0, 1);
                }
            }
            catch (Exception ex)
            {
                showError.ShowErr("Ошибка выгрузки", ex.Message);
            }
        }
        private List<List<object>> GetRowsForExport()
        {
            var rowsList = new List<List<object>>();

            if (checkBoxPreview.Checked)
            {
                // Режим предпросмотра — собираем из dataGridViewPreview
                var headers = new List<object>();
                foreach (DataGridViewColumn column in dataGridViewPreview.Columns)
                {
                    headers.Add(column.HeaderText);
                }
                rowsList.Add(headers);

                foreach (DataGridViewRow row in dataGridViewPreview.Rows)
                {
                    if (row.IsNewRow) continue;

                    var rowData = new List<object>();
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        rowData.Add(cell.Value?.ToString() ?? "");
                    }
                    rowsList.Add(rowData);
                }
            }
            else
            {

                // Режим без предпросмотра
                if (!materialCheckbox.Checked)
                {
                    // Альтернативный экспорт с материалами
                    var headers = new List<object>
                    {
                        "№", "Тип", "ТК", "Название", "Кол-во тестов",
                        "Кол-во шагов", "Ревью ТК", "Статус ТК", "Тестировщик", "Комментарий", "Статус задачи"
                    };
                    rowsList.Add(headers);

                    foreach (var row in rows)
                    {
                        var cells = new List<object>
                        {
                            row.ElementAtOrDefault(0)?.ToString() ?? "",
                            row.ElementAtOrDefault(1)?.ToString() ?? "",
                            row.ElementAtOrDefault(2)?.ToString() ?? "",
                            row.ElementAtOrDefault(3)?.ToString() ?? "",
                            row.ElementAtOrDefault(4)?.ToString() ?? "",
                            row.ElementAtOrDefault(5)?.ToString() ?? "",
                            row.ElementAtOrDefault(6)?.ToString() ?? "",  // Ревью ТК
                            row.ElementAtOrDefault(9)?.ToString() ?? "",  // Статус задачи
                            row.ElementAtOrDefault(10)?.ToString() ?? "", // Тестировщик
                            "",                                           // Комментарий
                            row.ElementAtOrDefault(8)?.ToString() ?? ""   // Статус проверки
                        };
                        rowsList.Add(cells);
                    }
                }
                else
                {
                    // Стандартный экспорт
                    var headers = new List<object>
                    {
                        "№", "Тип", "ТК", "Название", "Кол-во тестов",
                        "Кол-во шагов", "Ревью ТК", "Приоритет",
                        "Статус задачи", "Статус проверки", "Тестировщик"
                    };
                    rowsList.Add(headers);

                    foreach (var row in rows)
                    {
                        rowsList.Add(row);
                    }
                }
            }

            return rowsList;
        }
        private SheetStats AnalyzeSheetData(IList<IList<object>> data)
        {
            if (data == null || data.Count < 2)
            {
                MaterialMessageBox.Show("Недостаточно данных для анализа.");
                return null;
            }

            var header = data[0].Select(c => c.ToString().Trim()).ToList();
            int statusIndex = header.FindIndex(h => h.IndexOf("Статус задачи", StringComparison.OrdinalIgnoreCase) >= 0);
            int testStatusIndex = header.FindIndex(h =>
                h.IndexOf("Статус теста", StringComparison.OrdinalIgnoreCase) >= 0 ||
                h.IndexOf("Статус ТК", StringComparison.OrdinalIgnoreCase) >= 0 || h.IndexOf("Статус проверки", StringComparison.OrdinalIgnoreCase) >= 0);

            int linkIndex = 2;               // "ID задачи" всегда колонка 0
            int testCountIndex = 4;          // "Количество тестов" всегда колонка 4
            int fifthIndex = header.Count > 4 ? 4 : -1;
            int reviewIndex = 6; // "Проведено ревью тест-кейсов" — всегда колонка 7
            if (statusIndex == -1 || testStatusIndex == -1)
            {
                header = data[1].Select(c => c.ToString().Trim()).ToList();
                statusIndex = header.FindIndex(h => h.IndexOf("Статус задачи", StringComparison.OrdinalIgnoreCase) >= 0);
                testStatusIndex = header.FindIndex(h =>
                    h.IndexOf("Статус теста", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    h.IndexOf("Статус ТК", StringComparison.OrdinalIgnoreCase) >= 0 || h.IndexOf("Статус проверки", StringComparison.OrdinalIgnoreCase) >= 0);
                if (statusIndex == -1 || testStatusIndex == -1)
                    throw new Exception("Не удалось определить колонки 'Статус задачи' или 'Статус теста/ТК/проверки'.");
            }
            var stats = new SheetStats();

            foreach (var row in data.Skip(1))
            {
                string status = row.ElementAtOrDefault(statusIndex)?.ToString()?.Trim() ?? "";
                string testStatus = row.ElementAtOrDefault(testStatusIndex)?.ToString()?.Trim() ?? "";
                string taskId = row.ElementAtOrDefault(linkIndex)?.ToString()?.Trim() ?? "";
                string fifthCol = row.ElementAtOrDefault(fifthIndex)?.ToString()?.Trim() ?? "";
                double testCount = double.TryParse(row.ElementAtOrDefault(testCountIndex)?.ToString(), out var val) ? val : 0;
                string reviewStatus = row.ElementAtOrDefault(reviewIndex)?.ToString()?.Trim() ?? "";
                
                if (reviewStatus == "Активный" || reviewStatus == "Требует обновления")
                    stats.ReviewedTestCases++;
                // Total tasks — max по первой колонке (если число)
                if (int.TryParse(row[0]?.ToString(), out int firstCol))
                    stats.TotalTasks = Math.Max(stats.TotalTasks, firstCol);

                // Closed
                if (status == "Waiting for Release" || (status == "Closed" && (string.IsNullOrWhiteSpace(fifthCol) || fifthCol == "0")))
                    stats.ClosedTasks++;

                // Reopened
                if (status == "Reopened" || status == "Reopen")
                    stats.ReopenedTasks++;

                // Resolved → не протестированные
                if (status == "Resolved" && !string.IsNullOrWhiteSpace(taskId))
                    stats.UntestedTaskLinks.Add($"https://jira.mos.social/browse/{taskId}");

                // Подсчёт тестов
                switch (testStatus)
                {
                    case "Pass":
                        stats.Passed += (int)testCount;
                        break;
                    case "Passed":
                        stats.Passed += (int)testCount;
                        break;
                    case "Pass with bug":
                        stats.PassedWithBug += (int)testCount;
                        break;
                    case "Failed":
                        stats.Failed += (int)testCount;
                        break;
                    case "Fail":
                        stats.Failed += (int)testCount;
                        break;
                }

                // Кол-во тест-кейсов
                stats.PlannedTestCases += (int)testCount;
            }

            return stats;
        }
        private void checkBoxPreview_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxPreview.Checked)
            {
                materialButton1.Enabled = true;
            }
            else {
                dataGridViewPreview.Columns.Clear();
                dataGridViewPreview.Rows.Clear();
                materialButton1.Enabled = false;
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(0);
        }
    }
    public class SheetStats
    {
        public int TotalTasks { get; set; }
        public int ClosedTasks { get; set; }
        public int ReopenedTasks { get; set; }
        public List<string> UntestedTaskLinks { get; set; } = new List<string>();
        public int PlannedTestCases { get; set; }
        public int Passed { get; set; }
        public int PassedWithBug { get; set; }
        public int ReviewedTestCases { get; set; }
        public int Failed { get; set; }
    }
}
