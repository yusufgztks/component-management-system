using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using CMS.Controls;
using CMS.Data;

namespace CMS.Views
{
    public class UserManagerForm : Form
    {
        private DataGridView _grid     = null!;
        private TextBox      _txtUser  = null!;
        private TextBox      _txtPass  = null!;
        private CheckBox     _chkAdmin = null!;

        public UserManagerForm()
        {
            BuildForm();
            RefreshGrid();
        }

        private void BuildForm()
        {
            Text            = "User Management";
            Size            = new Size(780, 560);
            MinimumSize     = new Size(700, 480);
            StartPosition   = FormStartPosition.CenterParent;
            BackColor       = Theme.BgMain;
            Font            = Theme.FontBody;

            // ── Üst başlık çizgisi ───────────────────────────────────────
            var topBar = new Panel
            {
                Dock      = DockStyle.Top,
                Height    = 52,
                BackColor = Theme.TopbarBg,
            };
            topBar.Paint += (s, e) =>
                e.Graphics.DrawLine(new Pen(Theme.TopbarBorder, 1), 0, 51, topBar.Width, 51);
            Controls.Add(topBar);

            new Label
            {
                Text      = "User Management",
                Font      = Theme.FontH2,
                ForeColor = Theme.TextPrimary,
                AutoSize  = true,
                Left      = 20, Top = 14,
                Parent    = topBar,
            };

            // ── Sol: Kullanıcı Listesi ───────────────────────────────────
            var leftPanel = new Panel
            {
                Left      = 0, Top = 52,
                Width     = 500,
                Anchor    = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right,
                BackColor = Theme.BgMain,
            };
            Controls.Add(leftPanel);
            Resize += (s, e) =>
            {
                leftPanel.Width  = ClientSize.Width - 260;
                leftPanel.Height = ClientSize.Height - 52;
            };

            new Label
            {
                Text      = "Registered Users",
                Font      = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Theme.TextSecondary,
                AutoSize  = true,
                Left      = 16, Top = 14,
                Parent    = leftPanel,
            };

            _grid = new DataGridView
            {
                Left                  = 8, Top = 42,
                ReadOnly              = true,
                AllowUserToAddRows    = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible     = false,
                MultiSelect           = false,
                SelectionMode         = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor       = Theme.BgSurface,
                BorderStyle           = BorderStyle.None,
                GridColor             = Theme.Border,
                Font                  = Theme.FontBody,
                ColumnHeadersHeight   = 32,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                Anchor                = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                Parent                = leftPanel,
            };
            _grid.AutoGenerateColumns = false;
            _grid.DefaultCellStyle              = new DataGridViewCellStyle { BackColor = Theme.BgSurface, ForeColor = Theme.TextPrimary, SelectionBackColor = Theme.AccentLight, SelectionForeColor = Theme.Accent, Padding = new Padding(4, 0, 4, 0) };
            _grid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(248, 250, 252), ForeColor = Theme.TextPrimary, SelectionBackColor = Theme.AccentLight, SelectionForeColor = Theme.Accent };
            _grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(243, 246, 250), ForeColor = Theme.TextSecondary, Font = new Font("Segoe UI", 8.5f, FontStyle.Bold) };
            _grid.EnableHeadersVisualStyles     = false;

            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Id",        HeaderText = "ID",          Width = 40,  DataPropertyName = "Id",       SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Username",  HeaderText = "Username",    Width = 160, DataPropertyName = "Username", SortMode = DataGridViewColumnSortMode.NotSortable, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "IsAdmin",   HeaderText = "Role",        Width = 100, DataPropertyName = "IsAdmin",  SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "IsActive",  HeaderText = "Status",      Width = 80,  DataPropertyName = "IsActive", SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "CreatedAt", HeaderText = "Created At",  Width = 120, DataPropertyName = "CreatedAt",SortMode = DataGridViewColumnSortMode.NotSortable });

            _grid.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (_grid.Columns[e.ColumnIndex].Name == "IsAdmin")
                {
                    bool isAdmin = e.Value is true;
                    e.Value = isAdmin ? "Admin" : "User";
                    e.CellStyle.ForeColor = isAdmin ? Theme.Accent : Theme.TextSecondary;
                    e.CellStyle.Font      = new Font("Segoe UI", 8.5f, isAdmin ? FontStyle.Bold : FontStyle.Regular);
                }
                if (_grid.Columns[e.ColumnIndex].Name == "IsActive")
                {
                    bool active = e.Value is true;
                    e.Value = active ? "Active" : "Inactive";
                    e.CellStyle.ForeColor = active ? Theme.AccentGreen : Theme.Danger;
                    e.CellStyle.Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold);
                }
                if (_grid.Columns[e.ColumnIndex].Name == "CreatedAt" && e.Value is DateTime dt)
                    e.Value = dt.ToString("dd.MM.yyyy HH:mm");
            };

            leftPanel.Resize += (s, e) =>
            {
                _grid.Width  = leftPanel.ClientSize.Width - 8;
                _grid.Height = leftPanel.ClientSize.Height - 42;
            };

            // ── Eylem butonları (grid altına kontekst menüsü yerine butonlar) ──
            var ctxMenu = new ContextMenuStrip();
            ctxMenu.Items.Add("Make Admin",          null, (s, e) => SetAdminForSelected(true));
            ctxMenu.Items.Add("Remove Admin",        null, (s, e) => SetAdminForSelected(false));
            ctxMenu.Items.Add("-");
            ctxMenu.Items.Add("Toggle Active/Inactive", null, (s, e) => ToggleActive());
            ctxMenu.Items.Add("Reset Password",         null, (s, e) => ResetPassword());
            ctxMenu.Items.Add("-");
            ctxMenu.Items.Add("Delete User",            null, (s, e) => DeleteSelected());
            _grid.ContextMenuStrip = ctxMenu;

            // ── Sağ: Yeni Kullanıcı Ekle ────────────────────────────────
            var rightCard = new Panel
            {
                Left      = ClientSize.Width - 252, Top = 52,
                Width     = 244,
                Anchor    = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                BackColor = Theme.BgSurface,
            };
            rightCard.Paint += (s, e) =>
            {
                using var pen = new Pen(Theme.Border, 1);
                e.Graphics.DrawRectangle(pen, 0, 0, rightCard.Width - 1, rightCard.Height - 1);
            };
            Controls.Add(rightCard);
            Resize += (s, e) =>
            {
                rightCard.Left   = ClientSize.Width - 252;
                rightCard.Height = ClientSize.Height - 52;
            };

            new Label { Text = "Add New User", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Theme.TextSecondary, AutoSize = true, Left = 16, Top = 16, Parent = rightCard };

            new Label { Text = "Username", Font = Theme.FontBold, ForeColor = Theme.TextSecondary, AutoSize = true, Left = 16, Top = 52, Parent = rightCard };
            _txtUser = new TextBox { Left = 16, Top = 72, Width = 212, Height = 32, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, Parent = rightCard };

            new Label { Text = "Password", Font = Theme.FontBold, ForeColor = Theme.TextSecondary, AutoSize = true, Left = 16, Top = 114, Parent = rightCard };
            _txtPass = new TextBox { Left = 16, Top = 134, Width = 212, Height = 32, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, PasswordChar = '●', Parent = rightCard };

            _chkAdmin = new CheckBox { Text = "Grant Admin Privileges", Font = Theme.FontBody, ForeColor = Theme.TextPrimary, Left = 16, Top = 182, AutoSize = true, Parent = rightCard };

            var btnAdd = new ModernButton { Text = "Add User", Style = ModernButton.BtnStyle.Success, Width = 212, Height = 36, Left = 16, Top = 216, Parent = rightCard };
            btnAdd.Click += BtnAddUser_Click;

            // Seçili kullanıcı için hızlı eylemler
            new Panel { Left = 16, Top = 270, Width = 212, Height = 1, BackColor = Theme.Border, Parent = rightCard };
            new Label { Text = "Selected User", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Theme.TextSecondary, AutoSize = true, Left = 16, Top = 280, Parent = rightCard };

            var btnToggleAdmin = new ModernButton { Text = "Toggle Admin Privileges", Style = ModernButton.BtnStyle.Primary, Width = 212, Height = 34, Left = 16, Top = 310, Parent = rightCard };
            btnToggleAdmin.Click += (s, e) => ToggleAdmin();

            var btnToggleActive = new ModernButton { Text = "Toggle Active/Inactive", Style = ModernButton.BtnStyle.Secondary, Width = 212, Height = 34, Left = 16, Top = 352, Parent = rightCard };
            btnToggleActive.Click += (s, e) => ToggleActive();

            var btnResetPw = new ModernButton { Text = "Reset Password", Style = ModernButton.BtnStyle.Secondary, Width = 212, Height = 34, Left = 16, Top = 394, Parent = rightCard };
            btnResetPw.Click += (s, e) => ResetPassword();

            var btnDelete = new ModernButton { Text = "Delete User", Style = ModernButton.BtnStyle.Danger, Width = 212, Height = 34, Left = 16, Top = 436, Parent = rightCard };
            btnDelete.Click += (s, e) => DeleteSelected();
        }

        private void BtnAddUser_Click(object? sender, EventArgs e)
        {
            string user = _txtUser.Text.Trim();
            string pass = _txtPass.Text;

            if (string.IsNullOrWhiteSpace(user) || string.IsNullOrWhiteSpace(pass))
            {
                MessageBox.Show("Username and password are required.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool ok = UserRepository.AddUser(user, pass, _chkAdmin.Checked);
            if (!ok)
            {
                MessageBox.Show($"The username '{user}' already exists.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _txtUser.Clear();
            _txtPass.Clear();
            _chkAdmin.Checked = false;
            RefreshGrid();
        }

        private AppUser? GetSelectedUser()
        {
            if (_grid.SelectedRows.Count == 0) { MessageBox.Show("Please select a user.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return null; }
            var row = _grid.SelectedRows[0];
            int id  = (int)row.Cells["Id"].Value!;
            return UserRepository.GetAll().FirstOrDefault(u => u.Id == id);
        }

        private void ToggleAdmin()
        {
            var u = GetSelectedUser();
            if (u == null) return;
            if (u.Id == UserRepository.CurrentUser?.Id) { MessageBox.Show("You cannot change your own admin privileges.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            UserRepository.SetAdmin(u.Id, !u.IsAdmin);
            RefreshGrid();
        }

        private void SetAdminForSelected(bool isAdmin)
        {
            var u = GetSelectedUser();
            if (u == null) return;
            UserRepository.SetAdmin(u.Id, isAdmin);
            RefreshGrid();
        }

        private void ToggleActive()
        {
            var u = GetSelectedUser();
            if (u == null) return;
            if (u.Id == UserRepository.CurrentUser?.Id) { MessageBox.Show("You cannot deactivate your own account.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            UserRepository.SetActive(u.Id, !u.IsActive);
            RefreshGrid();
        }

        private void ResetPassword()
        {
            var u = GetSelectedUser();
            if (u == null) return;
            string newPw = Microsoft.VisualBasic.Interaction.InputBox($"New password for '{u.Username}':", "Reset Password", "");
            if (string.IsNullOrWhiteSpace(newPw)) return;
            UserRepository.ChangePassword(u.Id, newPw);
            MessageBox.Show("Password updated successfully.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void DeleteSelected()
        {
            var u = GetSelectedUser();
            if (u == null) return;
            if (u.Id == UserRepository.CurrentUser?.Id) { MessageBox.Show("You cannot delete your own account.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (MessageBox.Show($"Are you sure you want to delete the user '{u.Username}'?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                UserRepository.DeleteUser(u.Id);
                RefreshGrid();
            }
        }

        private void RefreshGrid()
        {
            var table = new System.Data.DataTable();
            table.Columns.Add("Id",        typeof(int));
            table.Columns.Add("Username",  typeof(string));
            table.Columns.Add("IsAdmin",   typeof(bool));
            table.Columns.Add("IsActive",  typeof(bool));
            table.Columns.Add("CreatedAt", typeof(DateTime));

            foreach (var u in UserRepository.GetAll().OrderBy(x => x.Id))
                table.Rows.Add(u.Id, u.Username, u.IsAdmin, u.IsActive, u.CreatedAt);

            _grid.DataSource = table;
        }
    }
}