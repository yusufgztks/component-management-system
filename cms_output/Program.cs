using System;
using System.Windows.Forms;
using CMS.Data;
using CMS.Views;

namespace CMS
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Kullanıcı veritabanını başlat
            UserRepository.Initialize();

            // Login ekranını aç
            using var login = new LoginForm();
            if (login.ShowDialog() != DialogResult.OK)
                return; // Giriş yapılmadıysa kapat

            Application.Run(new MainForm());
        }
    }
}
