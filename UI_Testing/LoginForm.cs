using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using MaterialSkin.Controls;
using MaterialSkin;
using UI_Testing;
using Google.Apis.Auth.OAuth2;
using System.Threading;
using Google.Apis.Util.Store;
using Google.Apis.Util;

namespace UI_Testing
{
    public partial class LoginForm : MaterialForm
    {
        public string? Login { get; private set; }
        public string? Password { get; private set; }
        private ShowError showError = new ShowError();
        private const string CredentialsPath = "user.dat";
        private readonly MaterialSkinManager materialSkinManager;
        public LoginForm()
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
            // Дополнительно: установка размеров формы
            this.MaximizeBox = false;
            this.Sizable = false;
            CheckSavedCredentialsAsync();
        }
        public LoginForm(int i)
        {
            InitializeComponent();
            Application.Exit();
        }
        private async void CheckSavedCredentialsAsync()
        {
            if (File.Exists(CredentialsPath))
            {
                try
                {
                    var decrypted = Decrypt(File.ReadAllBytes(CredentialsPath));
                    var parts = decrypted.Split(':');
                    string username = parts[0];
                    string password = parts[1];

                    txtUsername.Text = username;
                    txtPassword.Text = password;
                    chkRememberMe.Checked = true;

                    var success = await TryLoginAsync(username, password);
                    if (success)
                    {
                        OpenMainForm();
                    }
                }
                catch (Exception ex)
                {
                    showError.ShowErr("Ошибка при чтении сохранённых данных", ex.Message, null);
                    MaterialMessageBox.Show($"Ошибка при чтении сохранённых данных: {ex.Message}","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private async Task<bool> TryLoginAsync(string username, string password)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var authToken = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{username}:{password}"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var response = await client.GetAsync("https://jira.mos.social/rest/api/2/myself");

                    if (response.IsSuccessStatusCode)
                    {
                        string token = await GetOAuth2Token(); // Получаем токен

                        if (string.IsNullOrEmpty(token))
                        {
                            // Токен не получен, показываем ошибку
                            showError.ShowErr("Авторизация Google", "Не удалось получить OAuth2 токен.");
                            return false;
                        }

                        Login = username;
                        Password = password;
                        return true; // Всё прошло успешно
                    }
                    else
                    {
                        string errorText = await response.Content.ReadAsStringAsync();
                        showError.ShowErr("Ошибка авторизации", errorText, response.StatusCode);
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                showError.ShowErr("Ошибка подключения", ex.Message);
                return false;
            }
        }

        private async Task<string> GetOAuth2Token()
        {
            try
            {
                string credPath = "credentials.json";
                string tokenPath = "token";

                if (!File.Exists(credPath))
                {
                    MaterialMessageBox.Show("Файл credentials.json не найден.", "Ошибка");
                    return null;
                }

                using (var stream = new FileStream(credPath, FileMode.Open, FileAccess.Read))
                {
                    var secrets = GoogleClientSecrets.FromStream(stream).Secrets;
                    string[] scopes = { "https://www.googleapis.com/auth/spreadsheets" };

                    var credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                        secrets,
                        scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(tokenPath, true)
                    );
                    if (credential.Token.IsStale)
                    {
                        // Попытка обновить токен
                        if (await credential.RefreshTokenAsync(CancellationToken.None))
                        {
                            return credential.Token.AccessToken;
                        }
                        else
                        {
                            // Удаляем старый token, если не удалось обновить
                            var dir = new DirectoryInfo(tokenPath);
                            if (dir.Exists) dir.Delete(true);

                            // Повторная авторизация
                            credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                                secrets,
                                scopes,
                                "user",
                                CancellationToken.None,
                                new FileDataStore(tokenPath, true)
                            );
                        }
                    }

                    return credential.Token.AccessToken;
                }
            }
            catch (Exception ex)
            {
                showError.ShowErr("Ошибка получения токена", ex.Message);
                return null;
            }
        }

        private void OpenMainForm()
        {
            //this.DialogResult = DialogResult.OK; // <- Обязательно!
            //BeginInvoke((MethodInvoker)delegate { this.Close(); }); // <- Хак, чтобы корректно закрыть MaterialForm
            MainForm mainForm = new MainForm(Login, Password);
            this.Visible = false;
            mainForm.Show();
            
        }

        // Шифрование и дешифрование данных
        private byte[] Encrypt(string text)
        {
            byte[] data = Encoding.UTF8.GetBytes(text);
            return ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
        }

        private string Decrypt(byte[] data)
        {
            byte[] decrypted = ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(decrypted);
        }

        private async void materialButton1_Click(object sender, EventArgs e)
        {
            string username = txtUsername.Text.Trim();
            string password = txtPassword.Text;

            var success = await TryLoginAsync(username, password);
            if (success)
            {
                if (chkRememberMe.Checked)
                {
                    var encrypted = Encrypt($"{username}:{password}");
                    File.WriteAllBytes(CredentialsPath, encrypted);
                }
                OpenMainForm();
            }
        }

        private void LoginForm_FormClosing(object sender, FormClosingEventArgs e)
        {
        }

        private void LoginForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(0);
        }
    }
}