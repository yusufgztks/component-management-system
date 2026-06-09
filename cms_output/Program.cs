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
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Kullanıcı veritabanını başlat
            UserRepository.Initialize();

            // Login ekranını aç
            using var login = new LoginForm();
            if (login.ShowDialog() != DialogResult.OK)
                return; // Giriş yapılmadıysa kapat

            try
            {
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                System.IO.File.WriteAllText("crash.log", ex.ToString());
                MessageBox.Show(ex.Message, "Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
