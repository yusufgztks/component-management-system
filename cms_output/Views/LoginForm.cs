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
        }

        private void BuildForm()
        {
            Text            = "CMS — Giriş";
            Size            = new Size(420, 520);
            MinimumSize     = new Size(420, 520);
            MaximumSize     = new Size(420, 520);
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Theme.BgMain;
            Font            = Theme.FontBody;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox     = false;

            // ── Logo / Başlık ────────────────────────────────────────────
            var pnlLogo = new Panel
            {
                Left = 0, Top = 0, Width = 420, Height = 120,
                BackColor = Theme.SidebarBg,
            };
            Controls.Add(pnlLogo);

            new Label
            {
                Text      = "CMS",
                Font      = new Font("Segoe UI", 28f, FontStyle.Bold),
                ForeColor = Color.FromArgb(79, 124, 255),
                AutoSize  = true,
                Left      = 30, Top = 28,
                Parent    = pnlLogo,
            };
            new Label
            {
                Text      = "Component Management System",
                Font      = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(100, 116, 150),
                AutoSize  = true,
                Left      = 30, Top = 82,
                Parent    = pnlLogo,
            };

            // ── Kart ────────────────────────────────────────────────────
            var card = new Panel
            {
                Left      = 30, Top = 148,
                Width     = 360, Height = 310,
                BackColor = Theme.BgSurface,
            };
            card.Paint += (s, e) =>
            {
                using var pen = new Pen(Theme.Border, 1);
                e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1);
            };
            Controls.Add(card);

            new Label
            {
                Text      = "Giriş Yap",
                Font      = new Font("Segoe UI", 13f, FontStyle.Bold),
                ForeColor = Theme.TextPrimary,
                AutoSize  = true,
                Left      = 24, Top = 24,
                Parent    = card,
            };

            // Kullanıcı adı
            new Label
            {
                Text      = "User Name",
                Font      = Theme.FontBold,
                ForeColor = Theme.TextSecondary,
                AutoSize  = true,
                Left      = 24, Top = 72,
                Parent    = card,
            };
            _txtUsername = new TextBox
            {
                Left            = 24, Top = 94,
                Width           = 312, Height = 34,
                BorderStyle     = BorderStyle.FixedSingle,
                BackColor       = Theme.BgSurface2,
                ForeColor       = Theme.TextPrimary,
                Font            = Theme.FontBody,
                PlaceholderText = "User Name ",
                Parent          = card,
            };
            _txtUsername.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) _txtPassword.Focus(); };

            // Şifre
            new Label
            {
                Text      = "Password",
                Font      = Theme.FontBold,
                ForeColor = Theme.TextSecondary,
                AutoSize  = true,
                Left      = 24, Top = 142,
                Parent    = card,
            };
            _txtPassword = new TextBox
            {
                Left            = 24, Top = 164,
                Width           = 312, Height = 34,
                BorderStyle     = BorderStyle.FixedSingle,
                BackColor       = Theme.BgSurface2,
                ForeColor       = Theme.TextPrimary,
                Font            = Theme.FontBody,
                PasswordChar    = '●',
                PlaceholderText = "Password",
                Parent          = card,
            };
            _txtPassword.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) DoLogin(); };

            // Hata mesajı
            _lblError = new Label
            {
                Text      = "",
                Font      = Theme.FontSmall,
                ForeColor = Theme.Danger,
                AutoSize  = false,
                Left      = 24, Top = 210,
                Width     = 312, Height = 18,
                Parent    = card,
            };

            // Giriş butonu
            var btnLogin = new ModernButton
            {
                Text   = "Giriş Yap",
                Style  = ModernButton.BtnStyle.Primary,
                Width  = 312, Height = 40,
                Left   = 24, Top = 240,
                Parent = card,
            };
            btnLogin.Click += (s, e) => DoLogin();

            // Alt bilgi
            new Label
            {
                Text      = "Varsayılan: admin / admin123",
                Font      = new Font("Segoe UI", 7.5f),
                ForeColor = Color.FromArgb(80, 96, 120),
                AutoSize  = true,
                Left      = 30, Top = 472,
            }.Parent = this;
        }

        private void DoLogin()
        {
            string user = _txtUsername.Text.Trim();
            string pass = _txtPassword.Text;

            if (string.IsNullOrWhiteSpace(user) || string.IsNullOrWhiteSpace(pass))
            {
                _lblError.Text = "Kullanıcı adı ve şifre gerekli.";
                return;
            }

            if (UserRepository.Login(user, pass))
            {
                DialogResult = DialogResult.OK;
                Close();
            }
            else
            {
                _lblError.Text = "Kullanıcı adı veya şifre hatalı.";
                _txtPassword.Clear();
                _txtPassword.Focus();
            }
        }
    }
}