using MaterialSkin.Controls;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
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
                version = version == "Незапланированные" ? "Unscheduled" : version;
                string project = Regex.Match(zqlDecoded, @"project\s*=\s*""([^""]+)""").Groups[1].Value;
                string folderName = "";
                var folderMatch = Regex.Match(zqlDecoded, @"folderName\s*in\s*\(\s*""([^""]+)""\s*\)", RegexOptions.IgnoreCase);
                if (folderMatch.Success)
                {
                    folderName = folderMatch.Groups[1].Value;
                }

                if (string.IsNullOrWhiteSpace(cycleName) || string.IsNullOrWhiteSpace(version) || string.IsNullOrWhiteSpace(project))
                    continue;

                var stats = new CycleTestStats
                {
                    Url = baseUrl,
                    CycleName = cycleName,
                    Version = version,
                    Project = project,
                    Total = await GetTestCountByStatus(cycleName, version, project, null, folderName),
                    Passed = await GetTestCountByStatus(cycleName, version, project, "%3D+PASS", folderName),
                    Failed = await GetTestCountByStatus(cycleName, version, project, "%3D+FAIL", folderName),
                    Errored = await GetTestCountByStatus(cycleName, version, project, "%3D+%22PASS+WITH+BUG%22", folderName),
                    Blocked = await GetTestCountByStatus(cycleName, version, project, "%3D+BLOCKED", folderName)
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
        public async Task<int> GetTestCountByStatus(string cycleName, string version, string project, string status = null, string folderName = null)
        {
            string token = await GetAoToken();

            string baseUrl = "https://jira.mos.social/rest/zephyr/latest/zql/executeSearch/";
            var builder = new StringBuilder();
            builder.Append($"?zqlQuery=cycleName=\"{Uri.EscapeDataString(cycleName)}\"");
            if (!string.IsNullOrWhiteSpace(folderName))
                builder.Append($"+AND+folderName=\"{Uri.EscapeDataString(folderName)}\"");
            builder.Append($"+AND+fixVersion=\"{Uri.EscapeDataString(version)}\"");
            builder.Append($"+AND+project=\"{Uri.EscapeDataString(project)}\"");

            if (!string.IsNullOrEmpty(status))
                builder.Append($"+AND+executionStatus+{status}");

            string fullUrl = baseUrl + builder.ToString();
            var request = new HttpRequestMessage(HttpMethod.Get, fullUrl);
            request.Headers.Add("ao-7deabf", token);
            request.Headers.UserAgent.ParseAdd("PostmanRuntime/7.44.0");

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
            int startAt = 0;
            int maxResults = 1000;

            string jql = $"fixVersion = \"{version}\" ORDER BY priority DESC, key ASC";
            string url = $"{BaseUrl}/rest/api/2/search" +
                            $"?jql={Uri.EscapeDataString(jql)}" +
                            $"&fields=key,summary,issuetype,status,priority,issuelinks" +
                            $"&startAt={startAt}&maxResults={maxResults}";

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
            var root = doc.RootElement;

            var issueArray = root.GetProperty("issues");
            foreach (var issueElem in issueArray.EnumerateArray())
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

            if (issue?.Fields?.IssueLinks == null)
                return testCases;

            foreach (var link in issue.Fields.IssueLinks)
            {
                // Outward: "is tested by"
                if (link?.OutwardIssue != null &&
                    link.Type?.Outward == "is tested by" &&
                    link.OutwardIssue.Fields?.IssueType?.Name == "Test")
                {
                    testCases.Add(link.OutwardIssue);
                }
                if (link?.OutwardIssue != null &&
                    link.Type?.Outward == "tests" &&
                    link.OutwardIssue.Fields?.IssueType?.Name == "Test")
                {
                    testCases.Add(link.OutwardIssue);
                }
                // Inward: "tested by"
                if (link?.InwardIssue != null &&
                    link.Type?.Inward == "tested by" &&
                    link.InwardIssue.Fields?.IssueType?.Name == "Test")
                {
                    testCases.Add(link.InwardIssue);
                }

                // Inward: "is a test for"
                if (link?.InwardIssue != null &&
                    link.Type?.Inward == "is a test for" &&
                    link.InwardIssue.Fields?.IssueType?.Name == "Test")
                {
                    testCases.Add(link.InwardIssue);
                }
            }

            return testCases;
        }
        public (string CycleName, string Version, string Project, List<string> FolderNames, string ExecutionStatus) ParseZqlUrl(string url)
        {
            var result = (
                CycleName: "",
                Version: "",
                Project: "",
                FolderNames: new List<string>(),
                ExecutionStatus: ""
            );

            int queryIndex = url.IndexOf("#?query=");
            if (queryIndex == -1)
                return result;

            string zqlEncoded = url.Substring(queryIndex + "#?query=".Length);
            string zqlDecoded = Uri.UnescapeDataString(zqlEncoded);

            // Основные поля
            result.CycleName = Regex.Match(zqlDecoded, @"cycleName\s*=\s*""([^""]+)""").Groups[1].Value;
            result.Version = Regex.Match(zqlDecoded, @"fixVersion\s*=\s*""([^""]+)""").Groups[1].Value;
            result.Project = Regex.Match(zqlDecoded, @"project\s*=\s*""([^""]+)""").Groups[1].Value;

            // FolderNames (если есть)
            var folderMatch = Regex.Match(zqlDecoded, @"folderName\s+IN\s*\(([^)]+)\)", RegexOptions.IgnoreCase);
            if (folderMatch.Success)
            {
                string group = folderMatch.Groups[1].Value;
                var folderNames = Regex.Matches(group, @"""([^""]+)""")
                                       .Cast<Match>()
                                       .Select(m => m.Groups[1].Value)
                                       .ToList();
                result.FolderNames = folderNames;
            }

            // ExecutionStatus (если есть)
            var execStatusMatch = Regex.Match(zqlDecoded, @"executionStatus\s*=\s*(\w+)", RegexOptions.IgnoreCase);
            if (execStatusMatch.Success)
            {
                result.ExecutionStatus = execStatusMatch.Groups[1].Value;
            }

            return result;
        }
        public async Task<List<JiraIssue>> GetTestCasesFromCycle(string cycleName, string version, string project, List<string> folderNames, string executionStatus)
        {
            var token = await GetAoToken();
            var allTests = new List<JiraIssue>();
            int startAt = 0;
            const int maxResults = 1000;

            do
            {
                var builder = new StringBuilder();
                builder.Append($"?zqlQuery=cycleName=\"{Uri.EscapeDataString(cycleName)}\"");
                builder.Append($"+AND+fixVersion=\"{Uri.EscapeDataString(version)}\"");
                builder.Append($"+AND+project=\"{Uri.EscapeDataString(project)}\"");

                if (folderNames != null && folderNames.Count > 0)
                {
                    var joinedFolders = string.Join(",", folderNames.Select(f => $"\"{Uri.EscapeDataString(f)}\""));
                    builder.Append($"+AND+folderName+IN+({joinedFolders})");
                }
                if (executionStatus != null)
                {
                    builder.Append($"+AND+executionStatus+=+{executionStatus}");
                }
                builder.Append($"&startAt={startAt}&maxRecords={maxResults}");

                string url = "https://jira.mos.social/rest/zephyr/latest/zql/executeSearch/" + builder.ToString();
                MaterialMessageBox.Show(url);
                var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.Add("ao-7deabf", token);
                request.Headers.UserAgent.ParseAdd("PostmanRuntime/7.44.0");

                var response = await httpClient.SendAsync(request);
                if (!response.IsSuccessStatusCode)
                {
                    string content = await response.Content.ReadAsStringAsync();
                    throw new Exception($"Ошибка запроса: {response.StatusCode}\n{content}");
                }

                string json = await response.Content.ReadAsStringAsync();
                var data = JObject.Parse(json);
                var executions = data["executions"];
                if (executions == null) break;

                foreach (var exec in executions)
                {
                    string key = exec["issueKey"]?.ToString();
                    string summary = exec["issueSummary"]?.ToString();
                    string priority = exec["priority"]?.ToString();

                    allTests.Add(new JiraIssue
                    {
                        Key = key,
                        Fields = new JiraFields
                        {
                            Summary = summary,
                            Priority = new Priority { Name = priority },
                            IssueType = new IssueType { Name = "Test" }
                        }
                    });
                }

                int totalCount = data["totalCount"]?.Value<int>() ?? 0;
                if (startAt + maxResults >= totalCount)
                    break;

                startAt += maxResults;

            } while (true);

            return allTests;
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
