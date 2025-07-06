using MaterialSkin.Controls;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UI_Testing
{
    class JiraClient
    {
        public static string BaseUrl { get; private set; } = "https://jira.mos.social";
        private readonly HttpClient httpClient;
        private readonly CookieContainer cookieContainer = new CookieContainer();
        public JiraClient(string username, string passwordOrToken)
        {
            string credentials = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{username}:{passwordOrToken}"));

            var handler = new HttpClientHandler
            {
                CookieContainer = cookieContainer,
                UseCookies = true
            };

            httpClient = new HttpClient(handler);
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", credentials);

            // Сделаем предварительный запрос, чтобы получить cookies (например, к /rest/api/2/myself)
            var task = httpClient.GetAsync("https://jira.mos.social/rest/api/2/myself");
            task.Wait();

            if (!task.Result.IsSuccessStatusCode)
            {
                throw new Exception("Ошибка при авторизации: " + task.Result.StatusCode);
            }

        }
        public async Task<List<CycleTestStats>> GetCycleTestStatsForCycles(List<string> urls)
        {
            var results = new List<CycleTestStats>();

            foreach (var baseUrl in urls)
            {
                int index = baseUrl.IndexOf("#?query=");
                if (index == -1) continue;

                string queryPart = baseUrl.Substring(index + "#?query=".Length);
                string zqlDecoded = Uri.UnescapeDataString(queryPart);

                string cycleName = Regex.Match(zqlDecoded, @"cycleName\s*(=|in)\s*\(?""([^""]+)""\)?").Groups[2].Value;
                string version = Regex.Match(zqlDecoded, @"fixVersion\s*=\s*""([^""]+)""").Groups[1].Value;
                string project = Regex.Match(zqlDecoded, @"project\s*=\s*""([^""]+)""").Groups[1].Value;

                if (string.IsNullOrWhiteSpace(cycleName) || string.IsNullOrWhiteSpace(version) || string.IsNullOrWhiteSpace(project))
                    continue;

                var stats = new CycleTestStats
                {
                    Url = baseUrl,
                    CycleName = cycleName,
                    Version = version,
                    Project = project,
                    Total = await GetTestCountByStatus(cycleName, version, project),
                    Passed = await GetTestCountByStatus(cycleName, version, project, "%3D+PASS"),
                    Failed = await GetTestCountByStatus(cycleName, version, project, "%3D+FAIL"),
                    Errored = await GetTestCountByStatus(cycleName, version, project, "%3D+%22PASS+WITH+BUG%22"),
                    Blocked = await GetTestCountByStatus(cycleName, version, project, "%3D+BLOCKED")
                };

                results.Add(stats);
            }

            return results;
        }
        public async Task<string> GetAoToken()
        {

            // ⚠ Добавь сюда куки или авторизацию, если Jira требует логин
            var response = await httpClient.GetAsync("https://jira.mos.social/secure/enav/");
            if (!response.IsSuccessStatusCode)
            {
                MaterialMessageBox.Show($"Не удалось получить страницу: {(int)response.StatusCode} {response.ReasonPhrase}");
                return null;
            }

            var html = await response.Content.ReadAsStringAsync();

            // 🧠 Парсим токен zEncKeyVal из JS
            var match = Regex.Match(html, @"var\s+zEncKeyVal\s*=\s*""([^""]+)""");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            else
            {
                MaterialMessageBox.Show("Не удалось найти переменную zEncKeyVal в HTML.");
                return null;
            }
        }
        public async Task<int> GetTestCountByStatus(string cycleName, string version, string project, string status = null)
        {
            string token = await GetAoToken();

            string baseUrl = "https://jira.mos.social/rest/zephyr/latest/zql/executeSearch/";
            var builder = new StringBuilder();
            builder.Append($"?zqlQuery=cycleName=\"{Uri.EscapeDataString(cycleName)}\"");
            builder.Append($"+AND+project=\"{Uri.EscapeDataString(project)}\"");

            if (!string.IsNullOrEmpty(status))
                builder.Append($"+AND+executionStatus+{status}");

            string fullUrl = baseUrl + builder.ToString();
            var request = new HttpRequestMessage(HttpMethod.Get, fullUrl);

            // ЗАГОЛОВКИ
            request.Headers.Add("ao-7deabf", token);

            // ДОБАВЛЯЕМ КУКИ ЯВНО
            request.Headers.UserAgent.ParseAdd("PostmanRuntime/7.44.0");
            // [ОПЦИОНАЛЬНО] Для отладки

            var response = await httpClient.SendAsync(request);

            if (!response.IsSuccessStatusCode)
            {
                var content = await response.Content.ReadAsStringAsync();
                throw new Exception($"Ошибка запроса: {response.StatusCode}\n{content}");
            }

            string json = await response.Content.ReadAsStringAsync();
            var data = JObject.Parse(json);

            return data["totalCount"]?.Value<int>() ?? 0;
        }

        public async Task<List<JiraIssue>> GetIssuesByVersion(string version)
        {
            var issues = new List<JiraIssue>();

            string jql = $"fixVersion = \"{version}\" ORDER BY key ASC";
            string url = $"{BaseUrl}/rest/api/2/search?jql={Uri.EscapeDataString(jql)}&fields=key,summary,issuetype,status,priority,issuelinks";

            var response = await httpClient.GetAsync(url);
            if (!response.IsSuccessStatusCode)
            {
                string error = await response.Content.ReadAsStringAsync();
                MaterialMessageBox.Show($"Ошибка JIRA: {error}");
                return issues;
            }

            var json = await response.Content.ReadAsStringAsync();
            if (!json.TrimStart().StartsWith("{"))
            {
                MaterialMessageBox.Show("JIRA не вернула JSON:\n" + json);
                throw new Exception("JIRA вернула HTML или другой неожиданный формат.");
            }

            using var doc = JsonDocument.Parse(json);
            foreach (var issueElem in doc.RootElement.GetProperty("issues").EnumerateArray())
            {
                var issue = JsonSerializer.Deserialize<JiraIssue>(issueElem.GetRawText());
                if (issue != null) issues.Add(issue);
            }

            return issues;
        }
        public async Task<string> GetIssueHtml(string issueKey)
        {
            var url = $"{BaseUrl}/browse/{issueKey}";
            try
            {
                var response = await httpClient.GetAsync(url);
                if (!response.IsSuccessStatusCode)
                {
                    MessageBox.Show($"Ошибка загрузки задачи {issueKey}: {response.StatusCode}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return null;
                }

                string content = await response.Content.ReadAsStringAsync();
                return content;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при загрузке HTML задачи: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }
        public async Task<string> GetIssueDescription(string issueKey)
        {
            var response = await httpClient.GetAsync($"{BaseUrl}/rest/api/2/issue/{issueKey}?fields=description");
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            var descriptionMatch = Regex.Match(json, "\\\"description\\\":\\\"(.*?)\\\",\\\"", RegexOptions.Singleline);
            return descriptionMatch.Success ? Regex.Unescape(descriptionMatch.Groups[1].Value) : null;
        }
        public async Task<JiraIssue?> GetIssueData(string issueKey)
        {
            string url = $"{BaseUrl}/rest/api/2/issue/{issueKey}?fields=summary,issuetype,status,priority,issuelinks";
            var response = await httpClient.GetAsync(url);

            if (!response.IsSuccessStatusCode)
            {
                string error = await response.Content.ReadAsStringAsync();
                MaterialMessageBox.Show($"Ошибка получения задачи {issueKey}: {error}");
                return null;
            }

            var json = await response.Content.ReadAsStringAsync();
            return JsonSerializer.Deserialize<JiraIssue>(json);
        }

        public List<JiraIssue> GetLinkedTestCases(JiraIssue issue)
        {
            var testCases = new List<JiraIssue>();

            if (issue.Fields.IssueLinks == null)
                return testCases;

            foreach (var link in issue.Fields.IssueLinks)
            {
                // Проверяем outward-связь и тип "is tested by"
                if (link.OutwardIssue != null &&
                    link.Type != null &&
                    link.Type.Outward == "is tested by" &&
                    link.OutwardIssue.Fields?.IssueType?.Name == "Test")
                {
                    testCases.Add(link.OutwardIssue);
                }
                if (link.InwardIssue != null &&
                link.Type != null &&
                link.Type.Inward == "tested by" &&
                link.OutwardIssue.Fields?.IssueType?.Name == "Test")
                {
                    testCases.Add(link.InwardIssue);
                }
                if (link.InwardIssue != null &&
                link.Type != null &&
                link.Type.Inward == "is a test for" &&
                link.OutwardIssue.Fields?.IssueType?.Name == "Test")
                {
                    testCases.Add(link.InwardIssue);
                }
                // с типом связи "tests", можно добавить аналогично.
            }

            return testCases;
        }
        public class JiraIssue
        {
            [JsonPropertyName("key")]
            public string Key { get; set; }

            [JsonPropertyName("fields")]
            public JiraFields Fields { get; set; }
        }

        public class JiraFields
        {
            [JsonPropertyName("summary")]
            public string Summary { get; set; }

            [JsonPropertyName("issuetype")]
            public IssueType IssueType { get; set; }

            [JsonPropertyName("status")]
            public Status StatusObj { get; set; }

            [System.Text.Json.Serialization.JsonIgnore]
            public string Status => StatusObj?.Name;

            [JsonPropertyName("priority")]
            public Priority Priority { get; set; }

            [JsonPropertyName("issuelinks")]
            public List<IssueLink>? IssueLinks { get; set; }

            [System.Text.Json.Serialization.JsonIgnore]
            public string Type => IssueType?.Name;

            [System.Text.Json.Serialization.JsonIgnore]
            public string PriorityValue => Priority?.Name;
        }

        public class IssueType
        {
            [JsonPropertyName("name")]
            public string Name { get; set; }
        }

        public class Status
        {
            [JsonPropertyName("name")]
            public string Name { get; set; }
        }

        public class Priority
        {
            [JsonPropertyName("name")]
            public string Name { get; set; }
        }

        public class IssueLink
        {
            [JsonPropertyName("inwardIssue")]
            public JiraIssue? InwardIssue { get; set; }

            [JsonPropertyName("outwardIssue")]
            public JiraIssue? OutwardIssue { get; set; }

            [JsonPropertyName("type")]
            public IssueLinkType? Type { get; set; }
        }

        public class IssueLinkType
        {
            [JsonPropertyName("name")]
            public string? Name { get; set; }

            [JsonPropertyName("inward")]
            public string? Inward { get; set; }

            [JsonPropertyName("outward")]
            public string? Outward { get; set; }
        }
    }
}
