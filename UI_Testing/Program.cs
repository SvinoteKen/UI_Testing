using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using UI_Testing.Properties;

namespace UI_Testing
{
    static class Status
    {
        public static bool Enter { get; set; }
    }
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            //DialogResult result;
            //using (var loginForm = new LoginForm())
            //{
            //    result = loginForm.ShowDialog();
            //    MessageBox.Show(result.ToString());
            //}

            //if (result == DialogResult.OK)
            //{
            //    Application.Run(new MainForm(null, null)); // Передай, если нужно
            //}
            //else
            //{
            //    Application.Exit();
            //}
            Application.Run(new LoginForm());
        }
    }
}
