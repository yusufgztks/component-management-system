using System;
using System.Drawing;
using System.Windows.Forms;
using CMS.Controls;
using CMS.Data;

namespace CMS.Views
{
    public class LoginForm : Form
    {
        private TextBox _txtUsername = null!;
        private TextBox _txtPassword = null!;
        private Label   _lblError    = null!;

        public LoginForm()
        {
            BuildForm();
            Theme.UpdateForm(this);
        }

        private void BuildForm()
        {
            Text            = "CMS — Login";
            ClientSize      = new Size(420, 420);
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Theme.SidebarBg;
            Font            = Theme.FontBody;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox     = false;

            try
            {
                using var bmp = new Bitmap(Image.FromFile(System.IO.Path.Combine(Application.StartupPath, "Resources", "feka_icon.png")));
                this.Icon = Icon.FromHandle(bmp.GetHicon());
            }
            catch { }

            // ── Logo / Başlık ────────────────────────────────────────────
            var pnlLogo = new Panel { Width = 420, Height = 148, Left = 0, Top = 0, BackColor = Theme.SidebarBg, Parent = this };

            var logoPath = System.IO.Path.Combine(Application.StartupPath, "Resources", "feka_logo.png");
            if (System.IO.File.Exists(logoPath))
            {
                var pb = new PictureBox
                {
                    Image     = Image.FromFile(logoPath),
                    SizeMode  = PictureBoxSizeMode.Zoom,
                    Width     = 220, Height = 74,
                    Left      = 100, Top = 20,
                    Parent    = pnlLogo,
                };
                new Label
                {
                    Text      = "Component\nManagement System",
                    Font      = new Font("Century Gothic", 12f, FontStyle.Bold),
                    ForeColor = Theme.TextSecondary,
                    TextAlign = ContentAlignment.MiddleCenter,
                    AutoSize  = false,
                    Width     = 420, Height = 48,
                    Left      = 0, Top = 90,
                    Parent    = pnlLogo,
                };
            }
            else
            {
                new Label
                {
                    Text      = "CMS",
                    Font      = new Font("Segoe UI", 28f, FontStyle.Bold),
                    ForeColor = Theme.Accent,
                    TextAlign = ContentAlignment.MiddleCenter,
                    AutoSize  = false,
                    Width     = 420, Height = 60,
                    Left      = 0, Top = 20,
                    Parent    = pnlLogo,
                };
                new Label
                {
                    Text      = "Component\nManagement System",
                    Font      = new Font("Century Gothic", 12f, FontStyle.Bold),
                    ForeColor = Theme.TextSecondary,
                    TextAlign = ContentAlignment.MiddleCenter,
                    AutoSize  = false,
                    Width     = 420, Height = 48,
                    Left      = 0, Top = 90,
                    Parent    = pnlLogo,
                };
            }

            var card = new CardPanel
            {
                Radius    = 12,
                Left      = 30, Top = 148,
                Width     = 360, Height = 238,
                CardColor = Theme.BgSurface,
            };
            Controls.Add(card);

            // Kullanıcı adı
            new Label
            {
                Text      = "Username",
                Font      = Theme.FontBold,
                ForeColor = Theme.TextSecondary,
                AutoSize  = true,
                Left      = 24, Top = 20,
                Parent    = card,
            };
            _txtUsername = new TextBox
            {
                Left            = 24, Top = 42,
                Width           = 312, Height = 34,
                BorderStyle     = BorderStyle.FixedSingle,
                BackColor       = Theme.BgSurface2,
                ForeColor       = Theme.TextPrimary,
                Font            = Theme.FontBody,
                PlaceholderText = "Username",
                Parent          = card,
            };
            Theme.MakeRounded(_txtUsername, 6);
            _txtUsername.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) _txtPassword.Focus(); };

            // Şifre
            new Label
            {
                Text      = "Password",
                Font      = Theme.FontBold,
                ForeColor = Theme.TextSecondary,
                AutoSize  = true,
                Left      = 24, Top = 90,
                Parent    = card,
            };
            _txtPassword = new TextBox
            {
                Left            = 24, Top = 112,
                Width           = 312, Height = 34,
                BorderStyle     = BorderStyle.FixedSingle,
                BackColor       = Theme.BgSurface2,
                ForeColor       = Theme.TextPrimary,
                Font            = Theme.FontBody,
                PasswordChar    = '●',
                PlaceholderText = "Password",
                Parent          = card,
            };
            Theme.MakeRounded(_txtPassword, 6);
            _txtPassword.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) DoLogin(); };

            // Hata mesajı
            _lblError = new Label
            {
                Text      = "",
                Font      = Theme.FontSmall,
                ForeColor = Theme.Danger,
                AutoSize  = false,
                TextAlign = ContentAlignment.MiddleCenter,
                Left      = 24, Top = 148,
                Width     = 312, Height = 18,
                Parent    = card,
            };

            // Giriş butonu
            var btnLogin = new ModernButton
            {
                Text   = "Login",
                Style  = ModernButton.BtnStyle.Primary,
                Width  = 312, Height = 40,
                Left   = 24, Top = 168,
                CustomBgColor = Theme.BgSurface,
                Parent = card,
            };
            btnLogin.Click += (s, e) => DoLogin();
        }

        private void DoLogin()
        {
            string user = _txtUsername.Text.Trim();
            string pass = _txtPassword.Text;

            if (string.IsNullOrWhiteSpace(user) || string.IsNullOrWhiteSpace(pass))
            {
                _lblError.Text = "Username and password are required.";
                return;
            }

            if (UserRepository.Login(user, pass))
            {
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                _lblError.Text = "Invalid username or password.";
                _txtPassword.Clear();
                _txtPassword.Focus();
            }
        }
    }
}