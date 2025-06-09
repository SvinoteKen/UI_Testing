using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UI_Testing
{
    class ShowError
    {
        public void ShowErr(string context, string errorText, HttpStatusCode? statusCode = null)
        {
            string statusInfo = statusCode.HasValue ? $" ({(int)statusCode} {statusCode.Value})" : "";
            string message = $"{context}{statusInfo}\n{errorText}\n\nЕсли ошибка не проходит, свяжитесь по Telegram: https://t.me/DI_KEN9";

            MaterialMessageBox.Show(message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
