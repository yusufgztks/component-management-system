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
            Theme.UpdateForm(this);
        }

        private void BuildForm()
        {
            Text            = "User Management";
            Size            = new Size(1000, 700);
            MinimumSize     = new Size(900, 600);
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

            var table = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent,
                Padding = new Padding(16, 16, 16, 16)
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 340f));
            Controls.Add(table);
            table.BringToFront();

            // ── Sol: Kullanıcı Listesi ───────────────────────────────────
            var leftCard = new CardPanel
            {
                Dock      = DockStyle.Fill,
                Margin    = new Padding(0, 0, 16, 0),
                CardColor = Theme.BgSurface,
                Radius    = 12
            };
            table.Controls.Add(leftCard, 0, 0);

            var gridHeader = new Panel { Dock = DockStyle.Top, Height = 56, BackColor = Color.Transparent };
            new Label { Text = "Registered Users", Font = Theme.FontH2, ForeColor = Theme.TextSecondary, AutoSize = true, Left = 16, Top = 16, Parent = gridHeader };
            leftCard.Controls.Add(gridHeader);

            _grid = new DataGridView
            {
                Dock                  = DockStyle.Fill,
                ReadOnly              = true,
                AllowUserToAddRows    = false,
                AllowUserToDeleteRows = false,
                RowHeadersVisible     = false,
                MultiSelect           = false,
                SelectionMode         = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor       = Theme.BgSurface,
                BorderStyle           = BorderStyle.None,
                CellBorderStyle       = DataGridViewCellBorderStyle.SingleHorizontal,
                GridColor             = Theme.Border,
                Font                  = Theme.FontBody,
                ColumnHeadersHeight   = 44,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None,
                RowTemplate           = { Height = 48 }
            };
            typeof(DataGridView).InvokeMember("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty, null, _grid, new object[] { true });
            _grid.AutoGenerateColumns = false;
            _grid.DefaultCellStyle              = new DataGridViewCellStyle { BackColor = Theme.BgSurface, ForeColor = Theme.TextPrimary, SelectionBackColor = Theme.AccentLight, SelectionForeColor = Theme.Accent, Padding = new Padding(16, 0, 16, 0) };
            _grid.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, SelectionBackColor = Theme.AccentLight, SelectionForeColor = Theme.Accent, Padding = new Padding(16, 0, 16, 0) };
            _grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Theme.BgSurface, ForeColor = Theme.TextSecondary, Font = new Font("Segoe UI", 9f, FontStyle.Bold), Padding = new Padding(16, 0, 16, 0) };
            _grid.EnableHeadersVisualStyles     = false;

            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Id",        HeaderText = "ID",          Width = 60, DataPropertyName = "Id",       SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "Username",  HeaderText = "Username",    Width = 160, DataPropertyName = "Username", SortMode = DataGridViewColumnSortMode.NotSortable, AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "IsAdmin",   HeaderText = "Role",        Width = 100, DataPropertyName = "IsAdmin",  SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "IsActive",  HeaderText = "Status",      Width = 100,  DataPropertyName = "IsActive", SortMode = DataGridViewColumnSortMode.NotSortable });
            _grid.Columns.Add(new DataGridViewTextBoxColumn { Name = "CreatedAt", HeaderText = "Created At",  Width = 140, DataPropertyName = "CreatedAt",SortMode = DataGridViewColumnSortMode.NotSortable });

            _grid.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (_grid.Columns[e.ColumnIndex].Name == "IsAdmin")
                {
                    bool isAdmin = e.Value is true;
                    e.Value = isAdmin ? "Admin" : "User";
                    e.CellStyle.ForeColor = isAdmin ? Theme.Accent : Theme.TextSecondary;
                    e.CellStyle.Font      = new Font("Segoe UI", 9f, isAdmin ? FontStyle.Bold : FontStyle.Regular);
                }
                if (_grid.Columns[e.ColumnIndex].Name == "IsActive")
                {
                    bool active = e.Value is true;
                    e.Value = active ? "Active" : "Inactive";
                    e.CellStyle.ForeColor = active ? Theme.AccentGreen : Theme.Danger;
                    e.CellStyle.Font      = new Font("Segoe UI", 9f, FontStyle.Bold);
                }
                if (_grid.Columns[e.ColumnIndex].Name == "CreatedAt" && e.Value is DateTime dt)
                    e.Value = dt.ToString("dd.MM.yyyy HH:mm");
            };
            leftCard.Controls.Add(_grid);
            _grid.BringToFront(); // Ensure grid is below header due to Docking

            // ── Sağ: Yeni Kullanıcı Ekle ────────────────────────────────
            var rightCard = new CardPanel
            {
                Dock      = DockStyle.Fill,
                CardColor = Theme.BgSurface,
                Radius    = 12
            };
            table.Controls.Add(rightCard, 1, 0);

            new Label { Text = "Add New User", Font = Theme.FontH2, ForeColor = Theme.TextSecondary, AutoSize = true, Left = 24, Top = 16, Parent = rightCard };

            new Label { Text = "Username", Font = Theme.FontBold, ForeColor = Theme.TextSecondary, AutoSize = true, Left = 24, Top = 64, Parent = rightCard };
            _txtUser = new TextBox { Left = 24, Top = 88, Width = 292, Height = 34, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, Parent = rightCard };
            Theme.MakeRounded(_txtUser, 6);

            new Label { Text = "Password", Font = Theme.FontBold, ForeColor = Theme.TextSecondary, AutoSize = true, Left = 24, Top = 140, Parent = rightCard };
            _txtPass = new TextBox { Left = 24, Top = 164, Width = 292, Height = 34, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, PasswordChar = '●', Parent = rightCard };
            Theme.MakeRounded(_txtPass, 6);

            _chkAdmin = new CheckBox { Text = "Grant Admin Privileges", Font = Theme.FontBody, ForeColor = Theme.TextPrimary, Left = 24, Top = 216, AutoSize = true, Parent = rightCard };

            var btnAdd = new ModernButton { Text = "Add User", Style = ModernButton.BtnStyle.Success, Width = 292, Height = 40, Left = 24, Top = 252, CustomBgColor = Theme.BgSurface, Parent = rightCard };
            btnAdd.Click += BtnAddUser_Click;

            // Seçili kullanıcı için hızlı eylemler
            var sep = new Panel { Left = 24, Top = 320, Width = 292, Height = 1, BackColor = Theme.Border, Parent = rightCard };
            new Label { Text = "Selected User Actions", Font = Theme.FontH2, ForeColor = Theme.TextSecondary, AutoSize = true, Left = 24, Top = 340, Parent = rightCard };

            var btnMakeAdmin = new ModernButton { Text = "Make Admin", Style = ModernButton.BtnStyle.Primary, Width = 292, Height = 36, Left = 24, Top = 380, CustomBgColor = Theme.BgSurface, Parent = rightCard };
            btnMakeAdmin.Click += (s, e) => SetAdminForSelected(true);

            var btnRevokeAdmin = new ModernButton { Text = "Revoke Admin", Style = ModernButton.BtnStyle.Secondary, Width = 292, Height = 36, Left = 24, Top = 424, CustomBgColor = Theme.BgSurface, Parent = rightCard };
            btnRevokeAdmin.Click += (s, e) => SetAdminForSelected(false);

            var btnResetPw = new ModernButton { Text = "Reset Password", Style = ModernButton.BtnStyle.Secondary, Width = 292, Height = 36, Left = 24, Top = 468, CustomBgColor = Theme.BgSurface, Parent = rightCard };
            btnResetPw.Click += (s, e) => ResetPassword();

            var btnDelete = new ModernButton { Text = "Delete User", Style = ModernButton.BtnStyle.Danger, Width = 292, Height = 36, Left = 24, Top = 512, CustomBgColor = Theme.BgSurface, Parent = rightCard };
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