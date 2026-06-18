using System;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using CMS.Controls;

namespace CMS.Views
{
    public class LoginForm : Form
    {
        public string LoggedInUser { get; private set; } = "";
        public string LoggedInRole { get; private set; } = "";

        private TextBox      _txtUsername = null!;
        private TextBox      _txtPassword = null!;
        private Label        _lblError    = null!;
        private ModernButton _btnLogin    = null!;
        private bool         _loggingIn   = false;

        public LoginForm()
        {
            InitForm();
            BuildUI();
        }

        private void InitForm()
        {
            Text            = "CMS — Sign In";
            Size            = new Size(420, 540);
            MinimumSize     = new Size(420, 540);
            MaximumSize     = new Size(420, 540);
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Theme.BgMain;
            Font            = Theme.FontBody;
            AutoScaleMode   = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox     = false;
        }

        private void BuildUI()
        {
            // Accent stripe
            Controls.Add(new Panel { Dock = DockStyle.Top, Height = 4, BackColor = Color.FromArgb(79, 124, 255) });

            // Logo panel
            var logoPanel = new Panel { Left = 0, Top = 4, Width = ClientSize.Width, Height = 120, BackColor = Theme.BgMain };

            var lblTitle = new Label
            {
                Text      = "CMS",
                Font      = new Font("Segoe UI", 28f, FontStyle.Bold),
                ForeColor = Color.FromArgb(79, 124, 255),
                AutoSize  = true,
                Top       = 28,
            };
            lblTitle.Left = (ClientSize.Width - lblTitle.PreferredWidth) / 2;

            var lblSub = new Label
            {
                Text      = "Component Management System",
                Font      = new Font("Segoe UI", 9f),
                ForeColor = Color.FromArgb(100, 116, 150),
                AutoSize  = true,
                Top       = 74,
            };
            lblSub.Left = (ClientSize.Width - lblSub.PreferredWidth) / 2;

            logoPanel.Controls.Add(lblTitle);
            logoPanel.Controls.Add(lblSub);
            Controls.Add(logoPanel);

            // Card
            var card = new Panel { Left = 30, Top = 136, Width = ClientSize.Width - 60, Height = 320, BackColor = Theme.BgSurface };
            card.Paint += (s, e) => { using var pen = new Pen(Theme.Border, 1); e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1); };
            Controls.Add(card);

            int y = 28, lh = 20, fh = 36, gap = 48;

            // Username
            new Label { Text = "Username", Font = Theme.FontBold, ForeColor = Theme.TextSecondary, Left = 24, Top = y, Width = 200, Height = lh, Parent = card };
            y += lh + 4;
            _txtUsername = new TextBox { Left = 24, Top = y, Width = card.Width - 48, Height = fh, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = new Font("Segoe UI", 10.5f), PlaceholderText = "Enter your username", Parent = card };
            _txtUsername.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; _txtPassword.Focus(); } };
            y += fh + gap - lh - 4;

            // Password
            new Label { Text = "Password", Font = Theme.FontBold, ForeColor = Theme.TextSecondary, Left = 24, Top = y, Width = 200, Height = lh, Parent = card };
            y += lh + 4;
            _txtPassword = new TextBox { Left = 24, Top = y, Width = card.Width - 48, Height = fh, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = new Font("Segoe UI", 10.5f), PlaceholderText = "Enter your password", PasswordChar = '●', Parent = card };
            _txtPassword.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; DoLogin(); } };
            y += fh + gap - lh - 4 + 4;

            // Error label
            _lblError = new Label { Text = "", Font = new Font("Segoe UI", 8.5f), ForeColor = Color.FromArgb(220, 60, 60), Left = 24, Top = y, Width = card.Width - 48, Height = 18, Parent = card };
            y += 22;

            // Login button
            _btnLogin = new ModernButton { Text = "Sign In", Style = ModernButton.BtnStyle.Primary, Width = card.Width - 48, Height = 40, Left = 24, Top = y, Parent = card, Font = new Font("Segoe UI", 10f, FontStyle.Bold) };
            _btnLogin.Click += (s, e) => DoLogin();

            // Footer
            var lblVer = new Label { Text = "©Component Management System", Font = new Font("Segoe UI", 7.5f), ForeColor = Color.FromArgb(80, 96, 120), Left = 0, Top = ClientSize.Height - 46, Width = ClientSize.Width, TextAlign = ContentAlignment.MiddleCenter };
            Controls.Add(lblVer);

            AcceptButton = _btnLogin;
        }

        private async void DoLogin()
        {
            if (_loggingIn) return;

            string username = _txtUsername.Text.Trim();
            string password = _txtPassword.Text;

            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            { ShowError("Username and password cannot be empty."); return; }

            _loggingIn        = true;
            _btnLogin.Enabled = false;
            _btnLogin.Text    = "Signing in...";
            _lblError.Text    = "";

            try
            {
                var (success, role) = await Task.Run(() => TryLogin(username, password));

                if (success)
                {
                    LoggedInUser = username;
                    LoggedInRole = role;
                    DialogResult = DialogResult.OK;
                    Close();
                }
                else
                {
                    ShowError("Invalid username or password.");
                    _txtPassword.Clear();
                    _txtPassword.Focus();
                }
            }
            catch (Exception ex)
            {
                ShowError($"Connection error: {ex.Message}");
            }
            finally
            {
                _loggingIn        = false;
                _btnLogin.Enabled = true;
                _btnLogin.Text    = "Sign In";
            }
        }

        private static (bool success, string role) TryLogin(string username, string password)
        {
            using var conn = new SqlConnection(CMS.Data.ComponentRepository.ConnectionString);
            conn.Open();

            const string sql = @"
                SELECT Role FROM dbo.Users
                WHERE Username = @u
                  AND PasswordHash = @p
                  AND IsActive = 1";

            using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@u", username);
            cmd.Parameters.AddWithValue("@p", password);

            var result = cmd.ExecuteScalar();
            return result != null ? (true, result.ToString() ?? "viewer") : (false, "");
        }

        private void ShowError(string msg) { _lblError.Text = msg; _lblError.Invalidate(); }
    }
}


