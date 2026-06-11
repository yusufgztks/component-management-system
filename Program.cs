using System;
using System.Windows.Forms;
using CMS.Views;

namespace CMS
{
    internal static class Program
    {
        public static string CurrentUser { get; private set; } = "";
        public static string CurrentRole { get; private set; } = "";
        public static bool   ShouldRelogin { get; set; } = false;

        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();
            Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            while (true)
            {
                using var loginForm = new LoginForm();
                if (loginForm.ShowDialog() != DialogResult.OK)
                    break; // User closed login → exit app

                CurrentUser   = loginForm.LoggedInUser;
                CurrentRole   = loginForm.LoggedInRole;
                ShouldRelogin = false;

                Application.Run(new MainForm());

                if (!ShouldRelogin)
                    break; // Normal close → exit app

                // ShouldRelogin = true → loop back to login screen
                CurrentUser = "";
                CurrentRole = "";
            }
        }
    }
}