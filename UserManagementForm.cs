using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using CMS.Controls;

namespace CMS.Views
{
    public class UserManagementForm : Form
    {
        private DataGridView _grid        = null!;
        private Label        _lblStatus   = null!;
        private Panel        _detailPanel = null!;

        private TextBox  _txtNewUser = null!;
        private TextBox  _txtNewPass = null!;
        private ComboBox _cmbRole    = null!;

        private TextBox  _txtChgPass    = null!;
        private TextBox  _txtChgConfirm = null!;
        private Label    _lblSelTarget  = null!;
        private ComboBox _cmbChgRole    = null!;

        private int    _selectedUserId   = -1;
        private string _selectedUsername = "";

        public UserManagementForm()
        {
            InitForm();
            BuildLayout();
            LoadUsers();
        }

        private void InitForm()
        {
            Text            = "User Management";
            Size            = new Size(980, 640);
            MinimumSize     = new Size(860, 560);
            StartPosition   = FormStartPosition.CenterParent;
            BackColor       = Theme.BgMain;
            Font            = Theme.FontBody;
            AutoScaleMode   = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
        }

        private void BuildLayout()
        {
            Controls.Add(new Panel { Dock = DockStyle.Top, Height = 4, BackColor = Theme.Accent });

            var titleBar = new Panel { Dock = DockStyle.Top, Height = 52, BackColor = Theme.BgSurface };
            titleBar.Paint += (s, e) => e.Graphics.DrawLine(new Pen(Theme.Border, 1), 0, 51, titleBar.Width, 51);
            new Label { Text = "User Management", Font = Theme.FontH2, ForeColor = Theme.TextPrimary, AutoSize = true, Left = 20, Top = 16, Parent = titleBar };
            Controls.Add(titleBar);

            _lblStatus = new Label { Dock = DockStyle.Bottom, Height = 26, BackColor = Color.FromArgb(241, 245, 249), ForeColor = Theme.TextMuted, Font = Theme.FontSmall, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(12, 0, 0, 0) };
            Controls.Add(_lblStatus);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(16) };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            // ── LEFT: grid ────────────────────────────────────────────────
            var leftPanel = new Panel { Dock = DockStyle.Fill };
            root.Controls.Add(leftPanel, 0, 0);

            new Label { Text = "Users", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Theme.TextSecondary, Left = 0, Top = 0, Height = 24, AutoSize = true, Parent = leftPanel };

            var btnRefresh = new ModernButton { Text = "Refresh", Style = ModernButton.BtnStyle.Secondary, Width = 80, Height = 26, Left = 0, Top = 30, Parent = leftPanel };
            btnRefresh.Click += (s, e) => LoadUsers();

            _grid = new DataGridView
            {
                Left = 0, Top = 62,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                RowHeadersVisible = false, SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false, BackgroundColor = Theme.BgSurface, BorderStyle = BorderStyle.None,
                GridColor = Theme.Border, Font = Theme.FontBody,
                ColumnHeadersHeight = 34,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                Parent = leftPanel,
            };
            _grid.AutoGenerateColumns = false;
            _grid.DefaultCellStyle                = new DataGridViewCellStyle { BackColor = Theme.BgSurface, ForeColor = Theme.TextPrimary, SelectionBackColor = Theme.AccentLight, SelectionForeColor = Theme.Accent, Padding = new Padding(4, 0, 4, 0) };
            _grid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(248, 250, 252), ForeColor = Theme.TextPrimary, SelectionBackColor = Theme.AccentLight, SelectionForeColor = Theme.Accent };
            _grid.ColumnHeadersDefaultCellStyle   = new DataGridViewCellStyle { BackColor = Color.FromArgb(243, 246, 250), ForeColor = Theme.TextSecondary, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), SelectionBackColor = Color.FromArgb(243, 246, 250), SelectionForeColor = Theme.TextSecondary };
            _grid.EnableHeadersVisualStyles = false;
            _grid.RowTemplate.Height = 32;

            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Id",        HeaderText = "ID",         DataPropertyName = "Id",        Width = 45  });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Username",  HeaderText = "Username",   DataPropertyName = "Username",  AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Role",      HeaderText = "Role",       DataPropertyName = "Role",      Width = 90  });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "IsActive",  HeaderText = "Status",     DataPropertyName = "IsActive",  Width = 80  });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "CreatedAt", HeaderText = "Created At", DataPropertyName = "CreatedAt", Width = 120 });

            _grid.CellFormatting += (s, e) =>
            {
                if (e.ColumnIndex < 0 || e.RowIndex < 0 || e.Value == null) return;
                var colName = _grid.Columns[e.ColumnIndex].Name;
                if (colName == "IsActive")
                {
                    bool active = e.Value.ToString() == "True" || e.Value.ToString() == "1";
                    e.Value = active ? "Active" : "Inactive";
                    e.CellStyle.ForeColor = active ? Theme.AccentGreen : Theme.Danger;
                    e.CellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);
                    e.FormattingApplied = true;
                }
                if (colName == "Role")
                {
                    e.CellStyle.ForeColor = e.Value.ToString() == "admin" ? Theme.Accent : Theme.TextSecondary;
                    e.CellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);
                }
            };

            _grid.SelectionChanged += Grid_SelectionChanged;
            leftPanel.Resize += (s, e) => { _grid.Width = leftPanel.ClientSize.Width; _grid.Height = leftPanel.ClientSize.Height - 62; };

            // ── RIGHT: detail panel ───────────────────────────────────────
            _detailPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 0, 0, 0) };
            root.Controls.Add(_detailPanel, 1, 0);
            BuildDetailPanel();
        }

        private void BuildDetailPanel()
        {
            _detailPanel.Controls.Clear();
            int y = 0;

            // ── Add User ──────────────────────────────────────────────────
            var cardAdd = MakeCard("Add New User", ref y, 230);
            int cy = 36;
            FieldLabel("Username *", cardAdd, cy);
            _txtNewUser = FieldBox(cardAdd, cy); cy += 44;
            FieldLabel("Password *", cardAdd, cy);
            _txtNewPass = FieldBox(cardAdd, cy, isPassword: true); cy += 44;
            FieldLabel("Role *", cardAdd, cy);
            _cmbRole = new ComboBox { Left = 120, Top = cy, Width = cardAdd.Width - 136, DropDownStyle = ComboBoxStyle.DropDownList, Font = Theme.FontBody, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, FlatStyle = FlatStyle.Flat, Parent = cardAdd };
            _cmbRole.Items.AddRange(new object[] { "viewer", "editor", "admin" });
            _cmbRole.SelectedIndex = 0;
            cy += 44;
            var btnAdd = new ModernButton { Text = "Add User", Style = ModernButton.BtnStyle.Success, Left = 0, Top = cy, Width = cardAdd.Width, Height = 34, Parent = cardAdd };
            btnAdd.Click += BtnAddUser_Click;

            // ── Manage Selected User ────────────────────────────────────────
            y += 12;
            var cardMng = MakeCard("Manage Selected User", ref y, 400);
            int my = 36;
            _lblSelTarget = new Label { Text = "← Select a user from the list", Font = new Font("Segoe UI", 8.5f, FontStyle.Italic), ForeColor = Theme.TextMuted, Left = 0, Top = my, AutoSize = true, Parent = cardMng };
            my += 30;

            // Role change
            FieldLabel("Role", cardMng, my);
            _cmbChgRole = new ComboBox { Left = 120, Top = my, Width = cardMng.Width - 136, DropDownStyle = ComboBoxStyle.DropDownList, Font = Theme.FontBody, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, FlatStyle = FlatStyle.Flat, Parent = cardMng };
            _cmbChgRole.Items.AddRange(new object[] { "viewer", "editor", "admin" });
            my += 44;
            var btnChgRole = new ModernButton { Text = "Update Role", Style = ModernButton.BtnStyle.Primary, Left = 0, Top = my, Width = cardMng.Width, Height = 32, Parent = cardMng };
            btnChgRole.Click += BtnChangeRole_Click;
            my += 44;

            // Password change
            FieldLabel("New Password", cardMng, my);
            _txtChgPass = FieldBox(cardMng, my, isPassword: true); my += 44;
            FieldLabel("Confirm", cardMng, my);
            _txtChgConfirm = FieldBox(cardMng, my, isPassword: true); my += 44;
            var btnChgPwd = new ModernButton { Text = "Update Password", Style = ModernButton.BtnStyle.Primary, Left = 0, Top = my, Width = cardMng.Width, Height = 32, Parent = cardMng };
            btnChgPwd.Click += BtnChangePwd_Click;
            my += 44;

            // Status + delete
            var btnToggle = new ModernButton { Text = "Toggle Active / Inactive", Style = ModernButton.BtnStyle.Secondary, Left = 0, Top = my, Width = cardMng.Width, Height = 32, Parent = cardMng };
            btnToggle.Click += BtnToggleActive_Click;
            my += 40;
            var btnDelete = new ModernButton { Text = "Delete User", Style = ModernButton.BtnStyle.Danger, Left = 0, Top = my, Width = cardMng.Width, Height = 32, Parent = cardMng };
            btnDelete.Click += BtnDeleteUser_Click;
        }

        private void BtnChangeRole_Click(object? sender, EventArgs e)
        {
            if (_selectedUserId < 0) { Warn("Please select a user from the list."); return; }
            var role = _cmbChgRole.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(role)) { Warn("Please select a role."); return; }

            if (MessageBox.Show($"Set role of '{_selectedUsername}' to '{role}'?", "Update Role", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            try
            {
                using var conn = new SqlConnection(CMS.Data.ComponentRepository.ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand("UPDATE dbo.Users SET Role = @r WHERE Id = @id", conn);
                cmd.Parameters.AddWithValue("@r",  role);
                cmd.Parameters.AddWithValue("@id", _selectedUserId);
                cmd.ExecuteNonQuery();
                LoadUsers();
                SetStatus($"Role of '{_selectedUsername}' set to '{role}'.");
            }
            catch (Exception ex) { Warn($"Error: {ex.Message}"); }
        }

        // ── Helpers ───────────────────────────────────────────────────────
        private Panel MakeCard(string title, ref int y, int height)
        {
            var card = new Panel { Left = 0, Top = y, Width = _detailPanel.ClientSize.Width - 12, Height = height, BackColor = Theme.BgSurface, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, Padding = new Padding(16, 0, 16, 12), Parent = _detailPanel };
            card.Paint += (s, e) => { using var pen = new Pen(Theme.Border, 1); e.Graphics.DrawRectangle(pen, 0, 0, card.Width - 1, card.Height - 1); using var brush = new SolidBrush(Theme.Accent); e.Graphics.FillRectangle(brush, 0, 0, card.Width, 3); };
            new Label { Text = title, Font = new Font("Segoe UI", 9f, FontStyle.Bold), ForeColor = Theme.TextPrimary, Left = 16, Top = 12, AutoSize = true, Parent = card };
            _detailPanel.Resize += (s, e) => card.Width = _detailPanel.ClientSize.Width - 12;
            y += height;
            return card;
        }

        private static void FieldLabel(string text, Panel parent, int top) =>
            new Label { Text = text, Font = Theme.FontBold, ForeColor = Theme.TextSecondary, Left = 0, Top = top + 6, Width = 114, Height = 24, TextAlign = ContentAlignment.MiddleLeft, Parent = parent };

        private static TextBox FieldBox(Panel parent, int top, bool isPassword = false)
        {
            var tb = new TextBox { Left = 120, Top = top, Width = parent.Width - 136, Height = 30, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, Parent = parent, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            if (isPassword) tb.PasswordChar = '●';
            return tb;
        }

        // ══════════════════════════════════════════════════════════════════
        private void LoadUsers()
        {
            try
            {
                var table = new DataTable();
                using var conn = new SqlConnection(CMS.Data.ComponentRepository.ConnectionString);
                conn.Open();
                const string sql = "SELECT Id, Username, Role, IsActive, CONVERT(varchar(16), CreatedAt, 120) AS CreatedAt FROM dbo.Users ORDER BY Id";
                using var adapter = new SqlDataAdapter(sql, conn);
                adapter.Fill(table);
                _grid.DataSource = table;
                SetStatus($"{table.Rows.Count} user(s) loaded.");
            }
            catch (Exception ex) { SetStatus($"Error: {ex.Message}"); }
        }

        private void Grid_SelectionChanged(object? sender, EventArgs e)
        {
            if (_grid.SelectedRows.Count == 0) { _selectedUserId = -1; _selectedUsername = ""; return; }
            var row = _grid.SelectedRows[0];
            _selectedUserId   = Convert.ToInt32(row.Cells["Id"].Value);
            _selectedUsername = row.Cells["Username"].Value?.ToString() ?? "";
            var currentRole   = row.Cells["Role"].Value?.ToString() ?? "";
            if (_lblSelTarget != null) _lblSelTarget.Text = $"Target: {_selectedUsername}";
            if (_cmbChgRole != null)
            {
                int idx = _cmbChgRole.Items.IndexOf(currentRole);
                _cmbChgRole.SelectedIndex = idx >= 0 ? idx : 0;
            }
        }

        private void BtnAddUser_Click(object? sender, EventArgs e)
        {
            var username = _txtNewUser.Text.Trim();
            var password = _txtNewPass.Text;
            var role     = _cmbRole.SelectedItem?.ToString() ?? "viewer";

            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            { Warn("Username and password cannot be empty."); return; }
            if (username.Length < 3) { Warn("Username must be at least 3 characters."); return; }
            if (password.Length < 4) { Warn("Password must be at least 4 characters."); return; }

            try
            {
                using var conn = new SqlConnection(CMS.Data.ComponentRepository.ConnectionString);
                conn.Open();
                using (var chk = new SqlCommand("SELECT COUNT(*) FROM dbo.Users WHERE Username = @u", conn))
                {
                    chk.Parameters.AddWithValue("@u", username);
                    if ((int)chk.ExecuteScalar() > 0) { Warn($"'{username}' already exists."); return; }
                }
                const string ins = "INSERT INTO dbo.Users (Username, PasswordHash, Role, IsActive, CreatedAt) VALUES (@u, @p, @r, 1, GETDATE())";
                using var cmd = new SqlCommand(ins, conn);
                cmd.Parameters.AddWithValue("@u", username);
                cmd.Parameters.AddWithValue("@p", password);
                cmd.Parameters.AddWithValue("@r", role);
                cmd.ExecuteNonQuery();
                _txtNewUser.Clear(); _txtNewPass.Clear(); _cmbRole.SelectedIndex = 0;
                LoadUsers();
                SetStatus($"User '{username}' added successfully.");
            }
            catch (Exception ex) { Warn($"Error: {ex.Message}"); }
        }

        private void BtnChangePwd_Click(object? sender, EventArgs e)
        {
            if (_selectedUserId < 0) { Warn("Please select a user from the list."); return; }
            var newPwd = _txtChgPass.Text; var confirmPwd = _txtChgConfirm.Text;
            if (string.IsNullOrWhiteSpace(newPwd)) { Warn("New password cannot be empty."); return; }
            if (newPwd.Length < 4) { Warn("Password must be at least 4 characters."); return; }
            if (newPwd != confirmPwd) { Warn("Passwords do not match."); return; }

            if (MessageBox.Show($"Update password for '{_selectedUsername}'?", "Change Password", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            try
            {
                using var conn = new SqlConnection(CMS.Data.ComponentRepository.ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand("UPDATE dbo.Users SET PasswordHash = @p WHERE Id = @id", conn);
                cmd.Parameters.AddWithValue("@p",  newPwd);
                cmd.Parameters.AddWithValue("@id", _selectedUserId);
                cmd.ExecuteNonQuery();
                _txtChgPass.Clear(); _txtChgConfirm.Clear();
                SetStatus($"Password updated for '{_selectedUsername}'.");
                MessageBox.Show("Password updated successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { Warn($"Error: {ex.Message}"); }
        }

        private void BtnToggleActive_Click(object? sender, EventArgs e)
        {
            if (_selectedUserId < 0) { Warn("Please select a user from the list."); return; }
            if (_selectedUsername.Equals(Program.CurrentUser, StringComparison.OrdinalIgnoreCase))
            { Warn("You cannot deactivate your own account."); return; }

            try
            {
                using var conn = new SqlConnection(CMS.Data.ComponentRepository.ConnectionString);
                conn.Open();
                bool isActive;
                using (var q = new SqlCommand("SELECT IsActive FROM dbo.Users WHERE Id = @id", conn))
                { q.Parameters.AddWithValue("@id", _selectedUserId); isActive = Convert.ToBoolean(q.ExecuteScalar()); }

                string action = isActive ? "inactive" : "active";
                if (MessageBox.Show($"Set '{_selectedUsername}' to {action}?", "Toggle Status", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

                using var cmd = new SqlCommand("UPDATE dbo.Users SET IsActive = @v WHERE Id = @id", conn);
                cmd.Parameters.AddWithValue("@v",  isActive ? 0 : 1);
                cmd.Parameters.AddWithValue("@id", _selectedUserId);
                cmd.ExecuteNonQuery();
                LoadUsers();
                SetStatus($"'{_selectedUsername}' set to {action}.");
            }
            catch (Exception ex) { Warn($"Error: {ex.Message}"); }
        }

        private void BtnDeleteUser_Click(object? sender, EventArgs e)
        {
            if (_selectedUserId < 0) { Warn("Please select a user from the list."); return; }
            if (_selectedUsername.Equals(Program.CurrentUser, StringComparison.OrdinalIgnoreCase))
            { Warn("You cannot delete your own account."); return; }

            if (MessageBox.Show($"Permanently delete user '{_selectedUsername}'?\n\nThis action cannot be undone.", "Delete User", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            try
            {
                using var conn = new SqlConnection(CMS.Data.ComponentRepository.ConnectionString);
                conn.Open();
                using var cmd = new SqlCommand("DELETE FROM dbo.Users WHERE Id = @id", conn);
                cmd.Parameters.AddWithValue("@id", _selectedUserId);
                cmd.ExecuteNonQuery();
                _selectedUserId = -1; _selectedUsername = "";
                LoadUsers();
                SetStatus("User deleted.");
            }
            catch (Exception ex) { Warn($"Error: {ex.Message}"); }
        }

        private void SetStatus(string msg) => _lblStatus.Text = $"  {msg}";
        private static void Warn(string msg) => MessageBox.Show(msg, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}