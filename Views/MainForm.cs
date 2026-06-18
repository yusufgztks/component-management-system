using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Packaging;
using CMS.Controls;
using CMS.Data;
using CMS.Models;
using ClosedXML.Excel;

namespace CMS.Views
{
    public class MainForm : Form
    {
        private bool _dataLoaded = false;

        private Panel _sidebar = null!;
        private Panel _topbar  = null!;
        private Panel _content = null!;

        private readonly List<SidebarItem> _navItems = new();
        private string _activePage = "dashboard";

        private Label _lblPageTitle = null!;

        private Panel _pageDashboard  = null!;
        private Panel _pageCatalog    = null!;
        private Panel _pageAdd        = null!;
        private Panel _pageApproval   = null!;
        private Panel _pageExport     = null!;
        private Panel _pageCatManager = null!;

        // ── Catalog ────────────────────────────────────────────────────────
        private TextBox      _searchBox     = null!;
        private ComboBox     _filterCat     = null!;
        private ComboBox     _filterRating  = null!;
        private ComboBox     _filterProc    = null!;
        private ComboBox     _filterMfr     = null!;
        private DataGridView _grid          = null!;
        private Label        _lblCount      = null!;

        // ── Catalog Dynamic Filters ────────────────────────────────────────
        private Panel    _filterBarRef   = null!;
        private readonly List<ComboBox> _extraFilters = new();

        private TextBox  _addFekaNo  = null!;
        private ComboBox _addCat     = null!;
        private TextBox  _addDesc    = null!;
        private TextBox  _addPart    = null!;
        private ComboBox _addMfr     = null!;
        private ComboBox _addFoot    = null!;
        private ComboBox _addLibRef  = null!;
        private TextBox  _addFootRef = null!;
        private ComboBox _addRating  = null!;
        private ComboBox _addProcess = null!;
        private Panel    _extraPanel = null!;
        private readonly List<(string Key, Control Input)> _extraBoxes = new();

        private DataGridView _approvalGrid = null!;

        private StatCard _scTotal    = null!;
        private StatCard _scApproved = null!;
        private StatCard _scPending  = null!;
        private StatCard _scCats     = null!;
        private Panel    _catChart   = null!;
        private Panel    _mfrChart   = null!;

        private ComboBox _exportCat = null!;

        private ListBox _catList     = null!;
        private ListBox _fieldList   = null!;
        private Label   _lblCatTitle = null!;

        // ── Field Detail Panel ─────────────────────────────────────────────
        private ListBox _globalFieldList = null!;

        private SidebarItem? _navApproval;
        private SidebarItem? _navAdd;
        private SidebarItem? _navCatManager;
        private SidebarItem? _navUserMgmt;

        // ── Oturum bilgisi ─────────────────────────────────────────────────
        private string _currentUser = Program.CurrentUser;
        private string _currentRole = Program.CurrentRole;

        // ── Category → Excel Sheet name mapping ───────────────────────────
        private static readonly Dictionary<string, string> CategorySheetMap = new(StringComparer.OrdinalIgnoreCase)
        {
            ["RES"]     = "RESISTOR",
            ["CAP"]     = "CAPACITOR",
            ["IND"]     = "INDUCTOR",
            ["FB"]      = "FERRITE BEAD",
            ["CMC"]     = "CMC",
            ["DIO-REC"] = "DIODE RECTIFIER",
            ["DIO-SCH"] = "DIODE SWITCHING",
            ["DIO-ZEN"] = "DIODE ZENER",
            ["DIO-SW"]  = "DIODE SCHOTTKY",
            ["DIO-TVS"] = "DIODE TVS",
            ["NTC"]     = "NTC",
            ["PTC"]     = "PTC",
            ["LED"]     = "LED",
            ["TRAN"]    = "TRANSISTOR",
            ["MOS"]     = "MOSFET",
            ["DC-DC"]   = "DC-DC CONVERTER",
            ["LDO"]     = "LINEAR VOLTAGE REGULATOR",
            ["LED-DRV"] = "LED DRIVER",
            ["MCU"]     = "MICROCONTROLLER",
            ["OTHER-IC"]= "OTHER IC",
            ["CON"]     = "CONNECTOR",
            ["TERM"]    = "TERMINAL",
            ["PAD"]     = "PAD",
        };

        // ── Category → specific fields ────────────────────────────────────
        private static readonly Dictionary<string, List<string>> CategoryExcelFields = new(StringComparer.OrdinalIgnoreCase)
        {
            ["RES"]     = new() { "Resistance(Ohm)", "Power(W)", "Tolerance(%)" },
            ["CAP"]     = new() { "Capacitance(F)", "Voltage(V)", "Tolerance(%)", "Temperature Coefficient", "Failure Mode" },
            ["IND"]     = new() { "Inductance(H)", "Irms Current(A)", "Tolerance(%)", "Max DCR(Ohm)" },
            ["FB"]      = new() { "Impedance(Ohm)@Frequency(Hz)", "Max Rated Current(A)", "Max DCR(Ohm)" },
            ["CMC"]     = new() { "Inductance(H)@Frequency(Hz)", "Impedance(Ohm)@Frequency(Hz)", "Max Rated Current(A)", "Max DCR(Ohm)" },
            ["DIO-REC"] = new() { "Max Peak Reverse Voltage(V)", "Max Forward Voltage(V)", "Max Forward Current(A)" },
            ["DIO-SCH"] = new() { "Max Peak Reverse Voltage(V)", "Max Forward Voltage(V)", "Max Forward Current(A)" },
            ["DIO-ZEN"] = new() { "Nom Zener Voltage(V)", "Tolerance(%)" },
            ["DIO-SW"]  = new() { "Max Reverse Voltage(V)", "Max Forward Voltage(V)", "Avg Forward Current(V)" },
            ["DIO-TVS"] = new() { "Reverse Standoff Voltage(V)", "Peak Pulse Power(W)", "Max Peak Pulse Current(A)", "Polarization" },
            ["NTC"]     = new() { "Resistance@25°C(Ohm)", "Resistance Tolerance(%)", "Beta(K)", "Beta Tolerance(%)" },
            ["PTC"]     = new() { "Resistance@25°C(Ohm)", "Percentage R. Tolerance(%)" },
            ["LED"]     = new() { "Nom Forward Voltage(V)", "Nom Forward Current(A)", "BIN" },
            ["TRAN"]    = new() { "Type", "Max Collector-Emitter Voltage(V)", "Max Collector Current(I)", "Max Base-Emitter Voltage(V)", "Power(W)" },
            ["MOS"]     = new() { "Type", "Max Drain-Source Voltage(V)", "Max Drain Current(I)", "Max Gate-Source Voltage(V)", "Power(W)" },
            ["DC-DC"]   = new() { "Type", "Max Output Voltage(V)", "Max Output Current(I)" },
            ["LDO"]     = new() { "Output Type", "Max Output Voltage(V)", "Max Output Current(I)" },
            ["LED-DRV"] = new() { "Type", "Total Channel Count", "Output Current/Channel(A)" },
            ["MCU"]     = new() { "Core Processor", "Core Size(Bit)", "Clock Speed(Hz)", "Program Memory Size(B)", "Program Memory Type", "ASIL Category" },
            ["OTHER-IC"]= new() { },
            ["CON"]     = new() { "Gender", "Pin Count", "Pitch(mm)", "Color" },
            ["TERM"]    = new() { "Gender", "Base Material", "Plating Material" },
            ["PAD"]     = new() { "PAD X-Size(mm)", "PAD Y-Size(mm)", "Hole X-Size(mm)", "Hole Y-Size(mm)", "PAD Shape", "Hole Shape" },
        };

        public MainForm()
        {
            InitForm();
            BuildLayout();
            BuildSidebar();
            BuildTopbar();
            BuildDashboard();
            BuildCatalog();
            BuildAddForm();
            BuildApproval();
            BuildExport();
            BuildCategoryManager();
            Navigate("dashboard");

            this.Shown += async (s, e) => await LoadDataAsync();
        }

        // ══════════════════════════════════════════════════════════════════
        //  LOAD
        // ══════════════════════════════════════════════════════════════════
        private async Task LoadDataAsync()
        {
            var overlay = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(200, 30, 30, 40) };
            var lbl = new Label
            {
                Text = "  Loading data...",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 14f, FontStyle.Bold),
                AutoSize = true,
            };
            overlay.Controls.Add(lbl);
            overlay.Resize += (s, e) =>
            {
                lbl.Left = (overlay.Width  - lbl.Width)  / 2;
                lbl.Top  = (overlay.Height - lbl.Height) / 2;
            };
            _content.Controls.Add(overlay);
            overlay.BringToFront();
            lbl.Left = (_content.Width  - lbl.PreferredWidth)  / 2;
            lbl.Top  = (_content.Height - lbl.PreferredHeight) / 2;
            Application.DoEvents();

            await CMS.Data.ComponentRepository.InitializeAsync();

            _content.Controls.Remove(overlay);
            overlay.Dispose();
            _dataLoaded = true;

            RefreshDashboard();
            RefreshApprovalBadge();
        }

        // ══════════════════════════════════════════════════════════════════
        //  FORM INIT
        // ══════════════════════════════════════════════════════════════════
        private void InitForm()
        {
            Text            = "Component Management System";
            Size            = new Size(1300, 820);
            MinimumSize     = new Size(1100, 700);
            StartPosition   = FormStartPosition.CenterScreen;
            BackColor       = Theme.BgMain;
            Font            = Theme.FontBody;
            AutoScaleMode   = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(96F, 96F);
             WindowState     = FormWindowState.Maximized;
             WindowState = FormWindowState.Maximized;
FormBorderStyle = FormBorderStyle.FixedSingle;
MaximizeBox = false;
        }

        private void BuildLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1,
                Padding = new Padding(0), Margin = new Padding(0),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,  100));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            Controls.Add(root);

            _sidebar = new Panel { Dock = DockStyle.Fill, BackColor = Theme.SidebarBg, Margin = new Padding(0) };
            root.Controls.Add(_sidebar, 0, 0);

            var rightCol = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2,
                Padding = new Padding(0), Margin = new Padding(0),
            };
            rightCol.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            rightCol.RowStyles.Add(new RowStyle(SizeType.Absolute, 56));
            rightCol.RowStyles.Add(new RowStyle(SizeType.Percent,  100));
            root.Controls.Add(rightCol, 1, 0);

            _topbar  = new Panel { Dock = DockStyle.Fill, BackColor = Theme.TopbarBg, Margin = new Padding(0) };
            _content = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgMain, AutoScroll = true, Margin = new Padding(0), Padding = new Padding(24) };
            rightCol.Controls.Add(_topbar,  0, 0);
            rightCol.Controls.Add(_content, 0, 1);

            _topbar.Paint  += (s, e) => e.Graphics.DrawLine(new Pen(Theme.TopbarBorder, 1), 0, _topbar.Height - 1, _topbar.Width, _topbar.Height - 1);
            _sidebar.Paint += (s, e) => e.Graphics.DrawLine(new Pen(Color.FromArgb(36, 47, 66), 1), _sidebar.Width - 1, 0, _sidebar.Width - 1, _sidebar.Height);
        }

        private void BuildSidebar()
        {
            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown,
                WrapContents = false, AutoScroll = false,
                Padding = new Padding(0), Margin = new Padding(0),
            };
            _sidebar.Controls.Add(flow);

            var logoPanel = new Panel { Width = 220, Height = 68, BackColor = Theme.SidebarBg };
            logoPanel.Controls.Add(new Label { Text = " CMS", ForeColor = Color.FromArgb(79,124,255), Font = new Font("Segoe UI", 14f, FontStyle.Bold), AutoSize = false, Width = 200, Height = 36, Left = 16, Top = 14 });
            logoPanel.Controls.Add(new Label { Text = "Component Library", ForeColor = Color.FromArgb(100,116,150), Font = new Font("Segoe UI", 8f), AutoSize = false, Width = 200, Height = 16, Left = 18, Top = 48 });
            logoPanel.Paint += (s, e) => e.Graphics.DrawLine(new Pen(Color.FromArgb(36, 47, 66), 1), 0, 67, 220, 67);
            flow.Controls.Add(logoPanel);

            AddNavSection(flow, "");
            AddNavItem(flow, "",   "dashboard", "Dashboard");
            AddNavItem(flow, "",   "catalog",   "Catalog");

            AddNavSection(flow, "MANAGE");
            AddNavItem(flow, "➕", "add",        "Add Component");
            _navApproval = AddNavItem(flow, "", "approval", "Approvals");
            RefreshApprovalBadge();

            AddNavSection(flow, "SETTINGS");
            _navCatManager = AddNavItem(flow, "", "catmanager", "Category Manager");
            _navUserMgmt   = AddNavItem(flow, "", "usermgmt",  "User Management");

            AddNavSection(flow, "TOOLS");
            AddNavItem(flow, "", "export",     "Export / Import");

            // Role-based visibility
            ApplyRoleRestrictions();

            var versionPanel = new Panel { Width = 220, Height = 40, BackColor = Theme.SidebarBg };
            versionPanel.Paint += (s, e) =>
            {
                e.Graphics.DrawLine(new Pen(Color.FromArgb(36, 47, 66), 1), 0, 0, 220, 0);
                e.Graphics.DrawString("CMS", new Font("Segoe UI", 7.5f),
                    new SolidBrush(Color.FromArgb(80, 96, 120)), 16, 12);
            };
            flow.Controls.Add(versionPanel);
        }

        private void AddNavSection(FlowLayoutPanel flow, string title)
        {
            flow.Controls.Add(new Label
            {
                Text = title, Width = 220, Height = 28,
                Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(80, 96, 120),
                BackColor = Theme.SidebarBg,
                Padding = new Padding(16, 10, 0, 0),
            });
        }

        private SidebarItem AddNavItem(FlowLayoutPanel flow, string icon, string key, string label)
        {
            var item = new SidebarItem { ItemText = $"{icon}  {label}", PageKey = key, Width = 220 };
            item.Clicked += (s, e) => Navigate(key);
            flow.Controls.Add(item);
            _navItems.Add(item);
            if (key == "add") _navAdd = item;
            return item;
        }

        private void ApplyRoleRestrictions()
        {
            bool isAdmin  = _currentRole.Equals("admin",   StringComparison.OrdinalIgnoreCase);
            bool isEditor = _currentRole.Equals("editor",  StringComparison.OrdinalIgnoreCase) || isAdmin;

            // viewer → sadece dashboard + catalog; editor → + add; admin → hepsi
            if (_navAdd        != null) _navAdd.Visible        = isEditor;
            if (_navApproval   != null) _navApproval.Visible   = isAdmin;
            if (_navCatManager != null) _navCatManager.Visible = isAdmin;
            if (_navUserMgmt   != null) _navUserMgmt.Visible   = isAdmin;
        }

        private void BuildTopbar()
        {
            _lblPageTitle = new Label { Font = Theme.FontH2, ForeColor = Theme.TextPrimary, AutoSize = true, Left = 24, Top = 17 };
            _topbar.Controls.Add(_lblPageTitle);

            // Logout button
            var btnLogout = new ModernButton
            {
                Text   = "Sign Out",
                Style  = ModernButton.BtnStyle.Secondary,
                Width  = 70,
                Height = 28,
                Top    = 14,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
            };
            btnLogout.Click += (s, e) =>
            {
                if (MessageBox.Show("Are you sure you want to sign out?", "Sign Out",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Application.Restart();
                }
            };
            _topbar.Controls.Add(btnLogout);

            // User info
            var lblUser = new Label
            {
                Text      = $"  {_currentRole.ToUpper()}  |  {_currentUser}",
                Font      = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 116, 150),
                AutoSize  = true,
                Top       = 20,
                Anchor    = AnchorStyles.Top | AnchorStyles.Right,
            };
            _topbar.Controls.Add(lblUser);

            // Export CSV butonu
            var btnExport = new ModernButton
            {
                Text   = "Export CSV",
                Style  = ModernButton.BtnStyle.Primary,
                Width  = 110,
                Height = 32,
                Top    = 12,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
            };
            btnExport.Click += (s, e) => Navigate("export");
            _topbar.Controls.Add(btnExport);

            // Topbar resize
            _topbar.Resize += (s, e) =>
            {
                btnLogout.Left = _topbar.Width - 86;
                btnExport.Left = _topbar.Width - 210;
                lblUser.Left   = _topbar.Width - 210 - lblUser.PreferredWidth - 12;
            };
            _topbar.Width = _topbar.Width; // trigger resize
        }

        private void Navigate(string key)
        {
            // Role check
            bool isAdmin  = _currentRole.Equals("admin",  StringComparison.OrdinalIgnoreCase);
            bool isEditor = _currentRole.Equals("editor", StringComparison.OrdinalIgnoreCase) || isAdmin;

            if (key == "add" && !isEditor)
            { MessageBox.Show("Access denied.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if ((key == "approval" || key == "catmanager" || key == "usermgmt") && !isAdmin)
            { MessageBox.Show("Access denied.", "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            // User Management → open as modal dialog
            if (key == "usermgmt")
            {
                using var dlg = new UserManagementForm();
                dlg.ShowDialog(this);
                // Keep current page active
                foreach (var item in _navItems)
                { item.IsActive = item.PageKey == _activePage; item.Invalidate(); }
                return;
            }

            _activePage = key;

            foreach (var item in _navItems)
            { item.IsActive = item.PageKey == key; item.Invalidate(); }

            foreach (Control c in _content.Controls)
                c.Visible = false;

            Panel? page = key switch
            {
                "dashboard"  => _pageDashboard,
                "catalog"    => _pageCatalog,
                "add"        => _pageAdd,
                "approval"   => _pageApproval,
                "export"     => _pageExport,
                "catmanager" => _pageCatManager,
                _ => null
            };

            _lblPageTitle.Text = key switch
            {
                "dashboard"  => "Dashboard",
                "catalog"    => "Component Catalog",
                "add"        => "Add Component",
                "approval"   => "Pending Approvals",
                "export"     => "Export / Import",
                "catmanager" => "Category Manager",
                _ => ""
            };

            if (page != null) page.Visible = true;

            if (key == "dashboard")  { try { if (_dataLoaded) RefreshDashboard(); } catch { } }
            if (key == "approval")   RefreshApproval();
            if (key == "catmanager") RefreshCategoryManager();
            if (key == "catalog")
            {
                if (!_dataLoaded)
                {
                    MessageBox.Show("Data is still loading, please wait a few seconds and try again.");
                    return;
                }
                RebuildExtraFilters();
                RefreshCatalog();
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  DASHBOARD
        // ══════════════════════════════════════════════════════════════════
        private void BuildDashboard()
        {
            _pageDashboard = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgMain, Visible = false, AutoScroll = false };
            _content.Controls.Add(_pageDashboard);

            var statsRow = new TableLayoutPanel
            {
                Left = 0, Top = 0, Width = 900, Height = 110,
                ColumnCount = 4, RowCount = 1,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent, Padding = new Padding(0),
            };
            for (int i = 0; i < 4; i++) statsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            statsRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            _scTotal    = new StatCard { StatLabel = "Total Components", AccentColor = Color.FromArgb(79,124,255),  Dock = DockStyle.Fill, Margin = new Padding(0,0,12,0) };
            _scApproved = new StatCard { StatLabel = "Approved",         AccentColor = Color.FromArgb(32,193,148),  Dock = DockStyle.Fill, Margin = new Padding(0,0,12,0) };
            _scPending  = new StatCard { StatLabel = "Pending Review",   AccentColor = Color.FromArgb(245,158,11),  Dock = DockStyle.Fill, Margin = new Padding(0,0,12,0) };
            _scCats     = new StatCard { StatLabel = "Categories",       AccentColor = Color.FromArgb(168,85,247),  Dock = DockStyle.Fill, Margin = new Padding(0) };

            statsRow.Controls.Add(_scTotal, 0, 0); statsRow.Controls.Add(_scApproved, 1, 0);
            statsRow.Controls.Add(_scPending, 2, 0); statsRow.Controls.Add(_scCats, 3, 0);
            _pageDashboard.Controls.Add(statsRow);

            var chartsRow = new TableLayoutPanel
            {
                Left = 0, Top = 118, Width = 900, Height = 300,
                ColumnCount = 2, RowCount = 1,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                BackColor = Color.Transparent, Padding = new Padding(0),
            };
            chartsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            chartsRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            chartsRow.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            _catChart = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgSurface, Margin = new Padding(0,0,12,0) };
            _mfrChart = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgSurface, Margin = new Padding(0) };
            chartsRow.Controls.Add(_catChart, 0, 0);
            chartsRow.Controls.Add(_mfrChart, 1, 0);
            _pageDashboard.Controls.Add(chartsRow);

            bool catChartResizing = false;
            _catChart.Resize += (s, e) =>
            {
                if (catChartResizing || !_dataLoaded) return;
                catChartResizing = true;
                try { BuildCategoryChart(ComponentRepository.GetCategoryCounts()); }
                finally { catChartResizing = false; }
            };

            void ResizeDashboard()
            {
                int w = _pageDashboard.ClientSize.Width;
                int h = _pageDashboard.ClientSize.Height;
                statsRow.Width   = w;
                chartsRow.Width  = w;
                chartsRow.Height = Math.Max(200, h - 118 - 8);
            }

            _pageDashboard.Resize += (s, e) => ResizeDashboard();
        }

        private void RefreshDashboard()
        {
            var all      = ComponentRepository.GetAll();
            int total    = all.Count;
            int approved = all.Count(c => c.Status == ComponentStatus.Approved);
            int pending  = all.Count(c => c.Status == ComponentStatus.Pending);
            var catCounts = ComponentRepository.GetCategoryCounts();

            _scTotal.StatValue    = total.ToString("N0");    _scTotal.StatSub    = "in library";
            _scApproved.StatValue = approved.ToString("N0"); _scApproved.StatSub = $"{(total>0?approved*100/total:0)}% of total";
            _scPending.StatValue  = pending.ToString("N0");  _scPending.StatSub  = "awaiting review";
            _scCats.StatValue     = catCounts.Count.ToString(); _scCats.StatSub  = "unique categories";

            var allList = all.ToList();
            AttachStatCardClick(_scTotal,    () => BuildStatDetail("Total Components", total, allList));
            AttachStatCardClick(_scApproved, () => BuildStatDetail("Approved Components", approved, allList.Where(c => c.Status == ComponentStatus.Approved).ToList()));
            AttachStatCardClick(_scPending,  () => BuildStatDetail("Pending Components",  pending,  allList.Where(c => c.Status == ComponentStatus.Pending).ToList()));
            AttachStatCardClick(_scCats,     () => BuildCatDetail(catCounts));

            _scTotal.Invalidate(); _scApproved.Invalidate(); _scPending.Invalidate(); _scCats.Invalidate();
            BuildCategoryChart(catCounts);
            BuildManufacturerChart(ComponentRepository.GetManufacturerCounts());
        }

        private void AttachStatCardClick(StatCard card, Action onClick)
        {
            card.Click -= StatCard_Click;
            card.Tag = onClick;
            card.Click += StatCard_Click;
            card.Cursor = Cursors.Hand;
        }

        private void StatCard_Click(object? sender, EventArgs e)
        {
            if (sender is StatCard sc && sc.Tag is Action action)
                action();
        }

        private void BuildStatDetail(string title, int count, List<Component> items)
        {
            var dlg = new Form
            {
                Text = title,
                Size = new Size(520, 420),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Theme.BgMain,
                Font = Theme.FontBody,
                MinimizeBox = false, MaximizeBox = false,
            };

            var lblTitle = new Label
            {
                Text = $"{title}: {count:N0} items",
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Theme.TextPrimary,
                AutoSize = true,
                Left = 16, Top = 12,
                Parent = dlg,
            };

            var btnCopy = new ModernButton
            {
                Text = "Copy",
                Style = ModernButton.BtnStyle.Primary,
                Width = 90, Height = 28,
                Left = dlg.ClientSize.Width - 110, Top = 10,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Parent = dlg,
            };

            var grid = new DataGridView
            {
                Left = 8, Top = 48,
                Width = dlg.ClientSize.Width - 16,
                Height = dlg.ClientSize.Height - 100,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                ReadOnly = true,
                AllowUserToAddRows = false,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Theme.BgSurface,
                BorderStyle = BorderStyle.None,
                Font = Theme.FontBody,
                Parent = dlg,
            };
            grid.AutoGenerateColumns = false;
            grid.DefaultCellStyle = new DataGridViewCellStyle { BackColor = Theme.BgSurface, ForeColor = Theme.TextPrimary, SelectionBackColor = Theme.AccentLight, SelectionForeColor = Theme.Accent };
            grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor = Color.FromArgb(243,246,250), ForeColor = Theme.TextSecondary, Font = new Font("Segoe UI",8.5f,FontStyle.Bold) };
            grid.EnableHeadersVisualStyles = false;
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="Category",    HeaderText="Category",    Width=80,  DataPropertyName="Category"    });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="Description", HeaderText="Description", Width=200, AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill, DataPropertyName="Description" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="PartNumber",  HeaderText="Part No",     Width=130, DataPropertyName="PartNumber"   });
            grid.Columns.Add(new DataGridViewTextBoxColumn { Name="Status",      HeaderText="Status",      Width=90,  DataPropertyName="Status"       });

            var table = new System.Data.DataTable();
            table.Columns.Add("Category"); table.Columns.Add("Description");
            table.Columns.Add("PartNumber"); table.Columns.Add("Status");
            foreach (var c in items)
                table.Rows.Add(c.Category, c.Description, c.PartNumber, c.Status.ToString());
            grid.DataSource = table;

            btnCopy.Click += (s, e) =>
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"{title}: {count:N0} items");
                sb.AppendLine("Category\tDescription\tPart No\tStatus");
                foreach (DataGridViewRow r in grid.Rows)
                    sb.AppendLine($"{r.Cells[0].Value}\t{r.Cells[1].Value}\t{r.Cells[2].Value}\t{r.Cells[3].Value}");
                Clipboard.SetText(sb.ToString());
                MessageBox.Show("Copied to clipboard!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            dlg.Resize += (s, e) => { btnCopy.Left = dlg.ClientSize.Width - 110; };
            dlg.ShowDialog(this);
        }

        private void BuildCatDetail(Dictionary<string, int> catCounts)
        {
            var dlg = new Form
            {
                Text = "Categories",
                Size = new Size(360, 400),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Theme.BgMain,
                Font = Theme.FontBody,
                MinimizeBox = false, MaximizeBox = false,
            };

            new Label { Text = $"Total: {catCounts.Count} categories", Font = new Font("Segoe UI",10f,FontStyle.Bold), ForeColor = Theme.TextPrimary, AutoSize = true, Left = 16, Top = 12, Parent = dlg };

            var btnCopy = new ModernButton { Text="Copy", Style=ModernButton.BtnStyle.Primary, Width=90, Height=28, Left=dlg.ClientSize.Width-110, Top=10, Anchor=AnchorStyles.Top|AnchorStyles.Right, Parent=dlg };

            var grid = new DataGridView
            {
                Left=8, Top=48, Width=dlg.ClientSize.Width-16, Height=dlg.ClientSize.Height-68,
                Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right|AnchorStyles.Bottom,
                ReadOnly=true, AllowUserToAddRows=false, RowHeadersVisible=false,
                BackgroundColor=Theme.BgSurface, BorderStyle=BorderStyle.None, Font=Theme.FontBody, Parent=dlg,
            };
            grid.AutoGenerateColumns = false;
            grid.DefaultCellStyle = new DataGridViewCellStyle { BackColor=Theme.BgSurface, ForeColor=Theme.TextPrimary, SelectionBackColor=Theme.AccentLight, SelectionForeColor=Theme.Accent };
            grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle { BackColor=Color.FromArgb(243,246,250), ForeColor=Theme.TextSecondary, Font=new Font("Segoe UI",8.5f,FontStyle.Bold) };
            grid.EnableHeadersVisualStyles = false;
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText="Category", Width=120, DataPropertyName="Cat" });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText="Count",    Width=80,  DataPropertyName="Cnt", AutoSizeMode=DataGridViewAutoSizeColumnMode.Fill });

            var table = new System.Data.DataTable();
            table.Columns.Add("Cat"); table.Columns.Add("Cnt");
            foreach (var kvp in catCounts.OrderByDescending(k=>k.Value))
                table.Rows.Add(kvp.Key, kvp.Value);
            grid.DataSource = table;

            btnCopy.Click += (s, e) =>
            {
                var sb = new System.Text.StringBuilder();
                sb.AppendLine("Category\tCount");
                foreach (var kvp in catCounts.OrderByDescending(k=>k.Value))
                    sb.AppendLine($"{kvp.Key}\t{kvp.Value}");
                Clipboard.SetText(sb.ToString());
                MessageBox.Show("Copied to clipboard!", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            };

            dlg.Resize += (s, e) => { btnCopy.Left = dlg.ClientSize.Width - 110; grid.Width = dlg.ClientSize.Width-16; };
            dlg.ShowDialog(this);
        }

        private void BuildCategoryChart(Dictionary<string, int> data)
        {
            _catChart.Controls.Clear();
            _catChart.Paint -= null!;

            const int SCROLL_PAD_LEFT = 6;
            const int CAT_W           = 60;
            const int NUM_W           = 50;
            const int NUM_GAP         = 4;
            const int trackLeft       = CAT_W + 6;
            const int barH            = 8;
            const int rowH            = 40;

            int numColsWidth = (NUM_W + NUM_GAP) * 3 + 16 + 2;
            int panelW = Math.Max(300, _catChart.ClientSize.Width);
            int BAR_MAX = Math.Max(60, panelW - SCROLL_PAD_LEFT - trackLeft - 8 - numColsWidth);
            int numStartInRow = trackLeft + BAR_MAX + 8;

            var title = new Label
            {
                Text = "Components by Category",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Theme.TextSecondary,
                AutoSize = false, Dock = DockStyle.Top, Height = 36,
                Padding = new Padding(8, 10, 0, 0),
                BackColor = Theme.BgSurface
            };

            var hdrPanel = new Panel { Dock = DockStyle.Top, Height = 32, BackColor = Theme.BgSurface };
            hdrPanel.Controls.Add(new Label { Text = "Total", Left = numStartInRow + SCROLL_PAD_LEFT, Top = 20, Width = NUM_W, Height = 14, Font = new Font("Segoe UI", 7f, FontStyle.Bold), ForeColor = Theme.Accent, AutoSize = false, TextAlign = ContentAlignment.MiddleCenter });
            hdrPanel.Controls.Add(new Label { Text = "App",   Left = numStartInRow + SCROLL_PAD_LEFT + NUM_W + NUM_GAP, Top = 20, Width = NUM_W, Height = 14, Font = new Font("Segoe UI", 7f, FontStyle.Bold), ForeColor = Theme.AccentGreen, AutoSize = false, TextAlign = ContentAlignment.MiddleCenter });
            hdrPanel.Controls.Add(new Label { Text = "Pend",  Left = numStartInRow + SCROLL_PAD_LEFT + (NUM_W + NUM_GAP) * 2, Top = 20, Width = NUM_W, Height = 14, Font = new Font("Segoe UI", 7f, FontStyle.Bold), ForeColor = Color.FromArgb(245, 158, 11), AutoSize = false, TextAlign = ContentAlignment.MiddleCenter });

            var scroll = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgSurface, AutoScroll = true, Padding = new Padding(SCROLL_PAD_LEFT, 2, 2, 8) };
            _catChart.Controls.Add(scroll);
            _catChart.Controls.Add(hdrPanel);
            _catChart.Controls.Add(title);

            var allComps = ComponentRepository.GetAll();
            var fullData = ComponentRepository.AllCategories
                .ToDictionary(cat => cat, cat => data.TryGetValue(cat, out int v) ? v : 0);
            var sorted = fullData.OrderByDescending(x => x.Value).ThenBy(x => x.Key)
                .Select(x => (cat: x.Key, total: x.Value)).ToList();

            int max = sorted.Count > 0 ? sorted[0].total : 1;
            if (max == 0) max = 1;

            int yPos = 0;
            foreach (var (cat, total) in sorted)
            {
                int approved = allComps.Count(c => c.Category != null && c.Category.Trim().ToUpper() == cat.Trim().ToUpper() && c.Status == ComponentStatus.Approved);
                int pending  = allComps.Count(c => c.Category != null && c.Category.Trim().ToUpper() == cat.Trim().ToUpper() && c.Status == ComponentStatus.Pending);

                int rowWidth = numStartInRow + (NUM_W + NUM_GAP) * 3 + 16;
                var row = new Panel { Left = 0, Top = yPos, Width = rowWidth, Height = rowH, BackColor = Color.Transparent };
                row.Controls.Add(new Label { Text = cat, Left = 0, Top = (rowH - 16) / 2, Width = CAT_W, Height = 16, ForeColor = Theme.TextSecondary, Font = Theme.FontSmall, AutoSize = false, TextAlign = ContentAlignment.MiddleRight });

                int totalFillW = total    > 0 ? Math.Max(10, (int)(BAR_MAX * (double)total    / max)) : 0;
                int apprFillW  = approved > 0 ? Math.Max(10, (int)(BAR_MAX * (double)approved / max)) : 0;
                int pendFillW  = pending  > 0 ? Math.Max(10, (int)(BAR_MAX * (double)pending  / max)) : 0;

                var trackTotal = new Panel { Left = trackLeft, Top = 4,  Width = BAR_MAX, Height = barH, BackColor = Theme.BgSurface3 };
                if (totalFillW > 0) trackTotal.Controls.Add(new Panel { Width = totalFillW, Height = barH, BackColor = Theme.Accent });
                row.Controls.Add(trackTotal);

                var trackAppr = new Panel { Left = trackLeft, Top = 15, Width = BAR_MAX, Height = barH, BackColor = Theme.BgSurface3 };
                if (apprFillW > 0) trackAppr.Controls.Add(new Panel { Width = apprFillW, Height = barH, BackColor = Theme.AccentGreen });
                row.Controls.Add(trackAppr);

                var trackPend = new Panel { Left = trackLeft, Top = 26, Width = BAR_MAX, Height = barH, BackColor = Theme.BgSurface3 };
                if (pendFillW > 0) trackPend.Controls.Add(new Panel { Width = pendFillW, Height = barH, BackColor = Color.FromArgb(245, 158, 11) });
                row.Controls.Add(trackPend);

                int numY = (rowH / 2) - 7;
                row.Controls.Add(new Label { Text = total.ToString(),    Left = numStartInRow,                          Top = numY, Width = NUM_W, Height = 14, ForeColor = total    > 0 ? Theme.Accent                 : Theme.TextMuted, Font = Theme.FontSmall, AutoSize = false, TextAlign = ContentAlignment.MiddleCenter });
                row.Controls.Add(new Label { Text = approved.ToString(), Left = numStartInRow + NUM_W + NUM_GAP,        Top = numY, Width = NUM_W, Height = 14, ForeColor = approved > 0 ? Theme.AccentGreen            : Theme.TextMuted, Font = Theme.FontSmall, AutoSize = false, TextAlign = ContentAlignment.MiddleCenter });
                row.Controls.Add(new Label { Text = pending.ToString(),  Left = numStartInRow + (NUM_W + NUM_GAP) * 2, Top = numY, Width = NUM_W, Height = 14, ForeColor = pending  > 0 ? Color.FromArgb(245, 158, 11) : Theme.TextMuted, Font = Theme.FontSmall, AutoSize = false, TextAlign = ContentAlignment.MiddleCenter });

                scroll.Controls.Add(row);
                yPos += rowH;
            }

            _catChart.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                using var pen = new Pen(Theme.Border, 1);
                e.Graphics.DrawRectangle(pen, 0, 0, _catChart.Width - 1, _catChart.Height - 1);
            };
        }

        private void BuildManufacturerChart(Dictionary<string, int> data)
        {
            _mfrChart.Controls.Clear();
            _mfrChart.Paint -= null!;
            var title = new Label { Text = "Top Manufacturers", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Theme.TextSecondary, AutoSize = false, Dock = DockStyle.Top, Height = 40, Padding = new Padding(16,12,0,0), BackColor = Theme.BgSurface };
            _mfrChart.Controls.Add(title);
            var scroll = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgSurface, AutoScroll = true, Padding = new Padding(16,4,16,12) };
            _mfrChart.Controls.Add(scroll);
            var sorted = data.OrderByDescending(kvp => kvp.Value).ToList();
            int max = sorted.Count > 0 ? sorted[0].Value : 1;
            Color[] barColors = { Theme.AccentGreen, Theme.Accent, Color.FromArgb(168,85,247), Theme.Warning, Theme.Danger };
            var flow = new FlowLayoutPanel { Dock = DockStyle.Top, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, BackColor = Color.Transparent, Padding = new Padding(0) };
            scroll.Controls.Add(flow);
            int idx = 0;
            foreach (var kvp in sorted.Take(10))
            {
                var color = barColors[idx % barColors.Length];
                var row   = new Panel { Width = 600, Height = 28, BackColor = Color.Transparent };
                var lbl   = new Label { Text = kvp.Key, Left = 0, Top = 6, Width = 100, ForeColor = Theme.TextSecondary, Font = Theme.FontSmall, AutoSize = false };
                var track = new Panel { Left = 104, Top = 10, Width = 180, Height = 8, BackColor = Theme.BgSurface3 };
                var fill  = new Panel { Left = 0, Top = 0, Width = Math.Max(4, max>0?(int)(180.0*kvp.Value/max):0), Height = 8, BackColor = color };
                track.Controls.Add(fill);
                var val = new Label { Text = kvp.Value.ToString(), Left = 292, Top = 6, Width = 40, ForeColor = Theme.TextMuted, Font = Theme.FontSmall, AutoSize = false };
                row.Controls.Add(lbl); row.Controls.Add(track); row.Controls.Add(val);
                flow.Controls.Add(row);
                idx++;
            }
            _mfrChart.Paint += (s, e) => { e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias; using var pen = new Pen(Theme.Border,1); e.Graphics.DrawRectangle(pen,0,0,_mfrChart.Width-1,_mfrChart.Height-1); };
        }

        // ══════════════════════════════════════════════════════════════════
        //  CATALOG
        // ══════════════════════════════════════════════════════════════════
        private void BuildCatalog()
        {
            _pageCatalog = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgMain, Visible = false };
            _content.Controls.Add(_pageCatalog);

            var filterBar = new Panel { Left = 0, Top = 0, Height = 52, BackColor = Theme.BgSurface, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right };
            filterBar.Paint += (s, e) => e.Graphics.DrawLine(new Pen(Theme.Border, 1), 0, filterBar.Height-1, filterBar.Width, filterBar.Height-1);
            _pageCatalog.Controls.Add(filterBar);
            _filterBarRef = filterBar;

            _ = new Label { Text = "", Left = 10, Top = 14, AutoSize = true, Font = new Font("Segoe UI", 10f), BackColor = Color.Transparent, Parent = filterBar };

            _searchBox = new TextBox { Left = 34, Top = 11, Width = 210, Height = 28, BorderStyle = BorderStyle.FixedSingle, Font = Theme.FontBody, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, PlaceholderText = "Search by description, part no, manufacturer...", Parent = filterBar };
            _searchBox.TextChanged += (s, e) => RefreshCatalog();

            _filterCat = MakeCombo(252, 11, 120, filterBar, "All Categories");
            var sortedCats = ComponentRepository.AllCategories.OrderBy(c => c).ToArray();
            _filterCat.Items.AddRange(sortedCats.Cast<object>().ToArray());
            _filterCat.SelectedIndexChanged += (s, e) => { RebuildExtraFilters(); RefreshCatalog(); };

            _filterRating = MakeCombo(380, 11, 95, filterBar, "All Ratings");
            _filterRating.Items.AddRange(new object[] { "STD", "AECQ" });
            _filterRating.SelectedIndexChanged += (s, e) => RefreshCatalog();

            _filterProc = MakeCombo(483, 11, 90, filterBar, "All Processes");
            _filterProc.Items.AddRange(new object[] { "SMD", "THT" });
            _filterProc.SelectedIndexChanged += (s, e) => RefreshCatalog();

            _filterMfr = MakeCombo(581, 11, 150, filterBar, "All Manufacturers");
            _filterMfr.SelectedIndexChanged += (s, e) => RefreshCatalog();

            var btnClear = new ModernButton { Text = "Clear Filters", Style = ModernButton.BtnStyle.Secondary, Width = 100, Height = 28, Left = 739, Top = 12, Parent = filterBar };
            btnClear.Click += (s, e) =>
            {
                _searchBox.Text = "";
                _filterCat.SelectedIndex = _filterRating.SelectedIndex = _filterProc.SelectedIndex = _filterMfr.SelectedIndex = 0;
                foreach (var ef in _extraFilters) ef.SelectedIndex = 0;
                RefreshCatalog();
            };

            filterBar.Width = _pageCatalog.ClientSize.Width;
            _pageCatalog.Resize += (s, e) => filterBar.Width = _pageCatalog.ClientSize.Width;

            _lblCount = new Label { Left = 8, Top = 58, AutoSize = true, Font = Theme.FontSmall, ForeColor = Theme.TextMuted, BackColor = Color.Transparent, Parent = _pageCatalog };

            _grid = BuildGrid();
            _grid.Left = 0; _grid.Top = 76;
            _grid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            _grid.Width  = _pageCatalog.ClientSize.Width;
            _grid.Height = _pageCatalog.ClientSize.Height - 76;
            _grid.MultiSelect = true;
            _grid.KeyDown += Grid_KeyDown;
            _pageCatalog.Controls.Add(_grid);

            _pageCatalog.Resize += (s, e) =>
            {
                _grid.Height = _pageCatalog.ClientSize.Height - 76;
            };

            var ctxMenu = new ContextMenuStrip();
            ctxMenu.Items.Add(" Approve", null, (s, e) => GridApprove());
            ctxMenu.Items.Add(" Delete",  null, (s, e) => GridDelete());
            _grid.ContextMenuStrip = ctxMenu;
            _grid.MouseClick += (s, e) => { if (e.Button == MouseButtons.Right) { var hi = _grid.HitTest(e.X,e.Y); if (hi.RowIndex>=0) _grid.Rows[hi.RowIndex].Selected=true; } };

            _grid.ColumnHeaderMouseClick += (s, e) =>
            {
                _grid.ClearSelection();
                int colIdx = e.ColumnIndex;
                foreach (DataGridViewRow row in _grid.Rows)
                    if (!row.IsNewRow) row.Cells[colIdx].Selected = true;
            };
        }

        private ComboBox MakeCombo(int x, int y, int w, Control parent, string placeholder)
        {
            var cb = new ComboBox { Left=x, Top=y, Width=w, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextSecondary, FlatStyle=FlatStyle.Flat, Parent=parent };
            cb.Items.Insert(0, placeholder);
            cb.SelectedIndex = 0;
            return cb;
        }

        private DataGridView BuildGrid()
        {
            var g = new DataGridView
            {
                ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false,
                RowHeadersVisible = false, MultiSelect = true,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                RowTemplate = { Height = 30 },
                BackgroundColor = Theme.BgSurface, BorderStyle = BorderStyle.None,
                GridColor = Theme.Border, Font = Theme.FontBody,
                ColumnHeadersHeight = 34,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText,
            };
            g.AutoGenerateColumns = false;
            g.DefaultCellStyle                = new DataGridViewCellStyle { BackColor=Theme.BgSurface, ForeColor=Theme.TextPrimary, SelectionBackColor=Theme.AccentLight, SelectionForeColor=Theme.Accent, Padding=new Padding(4,0,4,0) };
            g.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle { BackColor=Color.FromArgb(248,250,252), ForeColor=Theme.TextPrimary, SelectionBackColor=Theme.AccentLight, SelectionForeColor=Theme.Accent, Padding=new Padding(4,0,4,0) };
            g.ColumnHeadersDefaultCellStyle   = new DataGridViewCellStyle { BackColor=Color.FromArgb(243,246,250), ForeColor=Theme.TextSecondary, Font=new Font("Segoe UI",8.5f,FontStyle.Bold), SelectionBackColor=Color.FromArgb(243,246,250), SelectionForeColor=Theme.TextSecondary, Padding=new Padding(4,0,4,0) };

            var cols = new[]
            {
                ("FekaNo",       "FekaNo",       110, false),
                ("ID",           "ID",            55,  false),
                ("Category",     "Category",      90,  false),
                ("Description",  "Description",   300, true),
                ("PartNumber",   "Part Number",   165, false),
                ("Manufacturer", "Manufacturer",  130, false),
                ("Footprint",    "Footprint",     125, false),
                ("Rating",       "Rating",        70,  false),
                ("Process",      "Process",       70,  false),
                ("Status",       "Status",        90,  false),
            };
            foreach (var (name, hdr, w, fill) in cols)
                g.Columns.Add(new DataGridViewTextBoxColumn { Name=name, HeaderText=hdr, DataPropertyName=name, Width=w, AutoSizeMode=fill?DataGridViewAutoSizeColumnMode.Fill:DataGridViewAutoSizeColumnMode.None, SortMode=DataGridViewColumnSortMode.NotSortable });

            g.EnableHeadersVisualStyles = false;
            g.CellFormatting += Grid_CellFormatting;
            return g;
        }

        private void Grid_CellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
        {
            var grid = sender as DataGridView;
            if (grid == null) return;
            if (grid.Columns[e.ColumnIndex].Name == "Status" && e.Value != null)
            { e.CellStyle.ForeColor = e.Value.ToString()=="Approved" ? Theme.AccentGreen : Theme.Warning; e.CellStyle.Font = new Font("Segoe UI",8.5f,FontStyle.Bold); }
            if (grid.Columns[e.ColumnIndex].Name == "Rating" && e.Value != null)
            { e.CellStyle.ForeColor = e.Value.ToString()=="AECQ" ? Theme.Accent : Theme.TextSecondary; e.CellStyle.Font = new Font("Segoe UI",8.5f,FontStyle.Bold); }
        }

        private void RebuildExtraFilters()
        {
            // Remove old extra filter combos from filterBar
            foreach (var ef in _extraFilters)
                _filterBarRef.Controls.Remove(ef);
            _extraFilters.Clear();

            // Remove extra columns from grid (keep base cols)
            var baseCols = new[] { "FekaNo","ID","Category","Description","PartNumber","Manufacturer","Footprint","Rating","Process","Status" };
            var toRemove = _grid.Columns.Cast<DataGridViewColumn>()
                .Where(c => !baseCols.Contains(c.Name)).Select(c => c.Name).ToList();
            foreach (var col in toRemove) _grid.Columns.Remove(col);

            // Collect fields: if a category is selected use its fields, else collect ALL unique fields
            List<(string Key, string Label)> fields;
            if (_filterCat.SelectedIndex > 0)
            {
                var cat = _filterCat.SelectedItem?.ToString() ?? "";
                fields = ComponentRepository.CategoryFields.TryGetValue(cat, out var cf) ? cf : new();
            }
            else
            {
                // All unique fields across every category, preserving first-seen label
                var seen = new Dictionary<string, string>();
                foreach (var cf in ComponentRepository.CategoryFields.Values)
                    foreach (var (k, l) in cf)
                        if (!seen.ContainsKey(k)) seen[k] = l;
                fields = seen.Select(kv => (kv.Key, kv.Value)).ToList();
            }

            if (fields.Count == 0) return;

            int x = 749;
            foreach (var (key, label) in fields)
            {
                var cb = MakeCombo(x, 11, 110, _filterBarRef, "All " + label);
                // Populate with distinct values from current (filtered) data
                string? selCat = _filterCat.SelectedIndex > 0 ? _filterCat.SelectedItem?.ToString() : null;
                var vals = ComponentRepository.GetAll()
                    .Where(c => (selCat == null || c.Category == selCat)
                             && c.ExtraFields.ContainsKey(key)
                             && !string.IsNullOrWhiteSpace(c.ExtraFields[key]))
                    .Select(c => c.ExtraFields[key]).Distinct().OrderBy(v => v).ToList();
                foreach (var v in vals) cb.Items.Add(v);
                cb.Tag = key;
                cb.SelectedIndexChanged += (s, e) => RefreshCatalog();
                _extraFilters.Add(cb);
                x += 118;

                // Add grid column
                _grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = key, HeaderText = label,
                    DataPropertyName = key, Width = 90,
                    AutoSizeMode = DataGridViewAutoSizeColumnMode.None,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                });
            }
        }

                private bool _refreshing = false;
        private void RefreshCatalog()
        {
            if (_refreshing) return;
            _refreshing = true;
            try
            {
                var query  = _searchBox?.Text ?? "";
                var cat    = _filterCat?.SelectedIndex    > 0 ? _filterCat.SelectedItem?.ToString()    : "";
                var rating = _filterRating?.SelectedIndex > 0 ? _filterRating.SelectedItem?.ToString() : "";
                var proc   = _filterProc?.SelectedIndex   > 0 ? _filterProc.SelectedItem?.ToString()   : "";
                var mfr    = _filterMfr?.SelectedIndex    > 0 ? _filterMfr.SelectedItem?.ToString()    : "";

                if (_filterMfr != null)
                {
                    var prevSel = _filterMfr.SelectedItem?.ToString();
                    _filterMfr.Items.Clear();
                    _filterMfr.Items.Insert(0, "All Manufacturers");
                    foreach (var m in ComponentRepository.GetManufacturers()) _filterMfr.Items.Add(m);
                    int idx = _filterMfr.Items.IndexOf(prevSel ?? "");
                    _filterMfr.SelectedIndex = idx > 0 ? idx : 0;
                }

                // Apply extra field filters
                var extraFieldFilters = _extraFilters
                    .Where(ef => ef.SelectedIndex > 0)
                    .Select(ef => (Key: ef.Tag?.ToString() ?? "", Value: ef.SelectedItem?.ToString() ?? ""))
                    .Where(kv => !string.IsNullOrEmpty(kv.Key))
                    .ToList();

                var results = ComponentRepository.Search(query!, cat!, rating!, proc!, mfr!);
                // Filter by extra fields
                if (extraFieldFilters.Count > 0)
                    results = results.Where(c => extraFieldFilters.All(kv =>
                        c.ExtraFields.TryGetValue(kv.Key, out var v) && v == kv.Value)).ToList();

                if (_lblCount != null) _lblCount.Text = $"{results.Count:N0} component{(results.Count!=1?"s":"")} found";

                // Build DataTable with base + extra columns
                var table = new System.Data.DataTable();
                var baseCols = new[]{"FekaNo","ID","Category","Description","PartNumber","Manufacturer","Footprint","Rating","Process","Status"};
                foreach (var col in baseCols) table.Columns.Add(col);
                // Add extra field columns if a category is selected
                List<string> extraKeys = new();
                if (_filterCat?.SelectedIndex > 0 && _filterCat.SelectedItem?.ToString() is string selCat
                    && ComponentRepository.CategoryFields.TryGetValue(selCat, out var catFields))
                {
                    extraKeys = catFields.Select(f => f.Key).ToList();
                    foreach (var key in extraKeys)
                        if (!table.Columns.Contains(key)) table.Columns.Add(key);
                }

                foreach (var c in results)
                {
                    var row = table.NewRow();
                    row["FekaNo"] = c.FekaNo; row["ID"] = c.Id; row["Category"] = c.Category;
                    row["Description"] = c.Description; row["PartNumber"] = c.PartNumber;
                    row["Manufacturer"] = c.Manufacturer; row["Footprint"] = c.Footprint;
                    row["Rating"] = c.Rating; row["Process"] = c.Process; row["Status"] = c.Status;
                    foreach (var key in extraKeys)
                        row[key] = c.ExtraFields.TryGetValue(key, out var v) ? v : "";
                    table.Rows.Add(row);
                }

                if (_grid != null)
                {
                    _grid.DataSource = table;
                    void SetW(string col, int w) { if (_grid.Columns.Contains(col) && _grid.Columns[col] != null) _grid.Columns[col]!.Width = w; }
                    SetW("FekaNo",110); SetW("ID",45); SetW("Category",80); SetW("PartNumber",165);
                    SetW("Manufacturer",130); SetW("Footprint",125); SetW("Rating",70); SetW("Process",70); SetW("Status",90);
                }
            }
            catch (Exception ex) { MessageBox.Show($"Failed to load catalog:\n\n{ex.Message}\n\nStack:\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
            finally { _refreshing = false; }
        }

        private void GridApprove()
        {
            if (_grid.SelectedCells.Count == 0) return;
            var row = _grid.SelectedCells[0].OwningRow;
            if (row == null) return;
            if (int.TryParse(row.Cells["ID"].Value?.ToString(), out int id))
            { ComponentRepository.Approve(id); RefreshCatalog(); RefreshApprovalBadge(); }
        }

        private void GridDelete()
        {
            if (_grid.SelectedCells.Count == 0) return;
            var row = _grid.SelectedCells[0].OwningRow;
            if (row == null) return;
            string desc = row.Cells["Description"].Value?.ToString() ?? "this component";
            if (MessageBox.Show($"Delete \"{desc}\"?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                if (int.TryParse(row.Cells["ID"].Value?.ToString(), out int id))
                { ComponentRepository.Delete(id); RefreshCatalog(); RefreshApprovalBadge(); }
        }

        private void Grid_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                var grid = sender as DataGridView;
                if (grid == null || grid.SelectedCells.Count == 0) return;
                DataObject? d = grid.GetClipboardContent();
                if (d != null) Clipboard.SetDataObject(d);
                e.Handled = true;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  ADD FORM
        // ══════════════════════════════════════════════════════════════════
        private void BuildAddForm()
        {
            _pageAdd = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgMain, Visible = false, AutoScroll = true };
            _content.Controls.Add(_pageAdd);

            var card = new Panel { Top=16, Width=720, BackColor=Theme.BgSurface, AutoSize=false, Anchor=AnchorStyles.Top, Padding=new Padding(32,28,32,28) };
            card.Left = Math.Max(0, (_pageAdd.ClientSize.Width - card.Width) / 2);
            card.Paint += (s, e) => { using var pen = new Pen(Theme.Border,1); e.Graphics.DrawRectangle(pen,0,0,card.Width-1,card.Height-1); };
            _pageAdd.Controls.Add(card);
            _pageAdd.Resize += (s, e) => { card.Width = Math.Min(720, _pageAdd.ClientSize.Width); card.Left = Math.Max(0, (_pageAdd.ClientSize.Width-card.Width)/2); };

            int y=28, lw=140, fw=340, lh=28, fh=32;
            Label L(string text, int top) => new Label { Text=text, Left=32, Top=top, Width=lw, Height=lh, Font=Theme.FontBold, ForeColor=Theme.TextSecondary, TextAlign=ContentAlignment.MiddleLeft, Parent=card };
            TextBox TB(int top) => new TextBox { Left=32+lw+8, Top=top, Width=fw, Height=fh, BorderStyle=BorderStyle.FixedSingle, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, Font=Theme.FontBody, Parent=card };

            L("Feka No", y); _addFekaNo = TB(y); y += 44;

            L("Category *", y);
            _addCat = new ComboBox { Left=32+lw+8, Top=y, Width=fw-90, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=card };
            foreach (var c in ComponentRepository.AllCategories.OrderBy(x => x)) _addCat.Items.Add(c);
            _addCat.SelectedIndexChanged += (s, e) => RebuildExtraFields();

            y += 44;

            L("Description *", y); _addDesc = TB(y); y += 44;
            L("Part Number *", y); _addPart = TB(y); y += 44;

            L("Manufacturer", y);
            _addMfr = new ComboBox { Left=32+lw+8, Top=y, Width=fw, DropDownStyle=ComboBoxStyle.DropDown, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=card, AutoCompleteMode=AutoCompleteMode.SuggestAppend, AutoCompleteSource=AutoCompleteSource.ListItems };
            foreach (var m in ComponentRepository.AllManufacturers) _addMfr.Items.Add(m);
            y += 44;

            L("Footprint", y);
            _addFoot = new ComboBox { Left=32+lw+8, Top=y, Width=fw, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=card };
            foreach (var f in ComponentRepository.AllFootprints) _addFoot.Items.Add(f);
            y += 44;

            L("Library Ref", y);
            _addLibRef = new ComboBox { Left=32+lw+8, Top=y, Width=fw, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=card };
            foreach (var l in ComponentRepository.AllLibraryRefs) _addLibRef.Items.Add(l);
            y += 44;

            L("Footprint Ref", y); _addFootRef = TB(y); y += 44;

            L("Rating", y);
            _addRating = new ComboBox { Left=32+lw+8, Top=y, Width=120, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=card };
            _addRating.Items.AddRange(new object[]{"STD","AECQ"});
            _addRating.SelectedIndex = 0;
            y += 44;

            L("Process", y);
            _addProcess = new ComboBox { Left=32+lw+8, Top=y, Width=120, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=card };
            _addProcess.Items.AddRange(new object[]{"SMD","THT"});
            _addProcess.SelectedIndex = 0;
            y += 44;

            int extraTop = y;
            _extraPanel = new Panel { Left=0, Top=extraTop, Width=card.Width, Height=0, BackColor=Color.Transparent, Parent=card };

            var sep      = new Panel { Left=32, Width=560, Height=1, BackColor=Theme.Border, Parent=card };
            var btnAdd   = new ModernButton { Text="Add Component", Style=ModernButton.BtnStyle.Success, Width=160, Height=36, Left=32, Parent=card };

            void Reposition()
            {
                int baseY = extraTop + _extraPanel.Height + 12;
                sep.Top = baseY; btnAdd.Top = baseY+12;
                card.Height = baseY+70; _extraPanel.Width = card.Width;
            }

            _addCat.SelectedIndexChanged += (s, e) => { RebuildExtraFields(); Reposition(); };
            btnAdd.Click   += BtnAddComponent_Click;
            Reposition();
        }

        private static readonly Dictionary<string, string[]> FieldComboOptions = new()
        {
            ["Tolerance(%)"]            = new[]{"0.5%","1%","5%","10%","20%"},
            ["Power(W)"]                = new[]{"0.0625W","0.1W","0.125W","0.2W","0.25W","0.33W","0.5W","0.75W","1W","2W"},
            ["Capacitance(F)"]          = new string[0],
            ["Voltage(V)"]              = new[]{"6.3V","10V","16V","25V","35V","50V","63V","100V","200V","450V"},
            ["Temperature Coefficient"] = new[]{"C0G","NP0","X5R","X7R","X7S","X7T","X8L","X8R"},
            ["Failure Mode"]            = new[]{"NFM","OM","ST","FT"},
            ["Polarization"]            = new[]{"Uni-Polar","Bi-Polar"},
            ["Type"]                    = new[]{"NPN","PNP","NPN&PNP","N-Ch","P-Ch","Buck","Buck-Boost","Linear","Fixed","Adjustable"},
            ["Gender"]                  = new[]{"Male","Female","Both"},
            ["Color"]                   = new[]{"Natural","Black","White","Red","Blue","Green","Yellow"},
            ["Base Material"]           = new[]{"CuZn","Cu","Brass","Steel"},
            ["Plating Material"]        = new[]{"Sn","Au","Ni","Ag"},
            ["PAD Shape"]               = new[]{"Rectangular","Round","Oval","Custom"},
            ["Hole Shape"]              = new[]{"Rectangular","Round","Oval","Slot"},
            ["Output Type"]             = new[]{"Fixed","Adjustable"},
        };

        private void RebuildExtraFields()
        {
            _extraPanel.Controls.Clear();
            _extraBoxes.Clear();

            int y=4; const int lw=140,fw=340,fh=32,rowH=44,lx=32,fx=32+lw+8;
            bool hasAny = false;

            // Category-specific fields
            if (_addCat.SelectedItem != null)
            {
                var cat = _addCat.SelectedItem.ToString()!;
                if (ComponentRepository.CategoryFields.TryGetValue(cat, out var fields) && fields.Count > 0)
                {
                    new Label { Text="▸  Category Parameters", Font=new Font("Segoe UI",8.5f,FontStyle.Bold), ForeColor=Theme.Accent, Left=lx, Top=y, Width=300, Height=24, BackColor=Color.Transparent, Parent=_extraPanel };
                    y += 28;
                    new Panel { Left=lx, Top=y, Width=fw+lw+8, Height=1, BackColor=Theme.AccentLight, Parent=_extraPanel };
                    y += 8;

                    foreach (var (key, label) in fields)
                    {
                        new Label { Text=label, Left=lx, Top=y+2, Width=lw, Height=fh, Font=Theme.FontBold, ForeColor=Theme.TextSecondary, TextAlign=ContentAlignment.MiddleLeft, BackColor=Color.Transparent, Parent=_extraPanel };
                        bool hasOptions = FieldComboOptions.TryGetValue(key, out var opts) && opts!=null && opts.Length>0;
                        if (key == "Type")
                        {
                            opts = cat switch { "TRAN"=>"NPN,PNP,NPN&PNP".Split(','), "MOS"=>"N-Ch,P-Ch".Split(','), "DC-DC"=>"Buck,Buck-Boost,Flyback,SEPIC,Cuk".Split(','), "LED-DRV"=>"Linear,Switching".Split(','), _=>new[]{"N/A"} };
                            hasOptions = true;
                        }
                        if (hasOptions)
                        {
                            var cb = new ComboBox { Left=fx, Top=y, Width=fw, Height=fh, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=_extraPanel };
                            foreach (var o in opts!) cb.Items.Add(o);
                            _extraBoxes.Add((key, cb));
                        }
                        else
                        {
                            var tb = new TextBox { Left=fx, Top=y, Width=fw, Height=fh, BorderStyle=BorderStyle.FixedSingle, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, Font=Theme.FontBody, Parent=_extraPanel };
                            _extraBoxes.Add((key, tb));
                        }
                        y += rowH;
                    }
                    hasAny = true;
                }
            }

            // Global fields
            if (ComponentRepository.GlobalFields.Count > 0)
            {
                new Label { Text="▸  Global Fields", Font=new Font("Segoe UI",8.5f,FontStyle.Bold), ForeColor=Theme.Accent, Left=lx, Top=y, Width=300, Height=24, BackColor=Color.Transparent, Parent=_extraPanel };
                y += 28;
                new Panel { Left=lx, Top=y, Width=fw+lw+8, Height=1, BackColor=Theme.AccentLight, Parent=_extraPanel };
                y += 8;

                foreach (var gf in ComponentRepository.GlobalFields)
                {
                    new Label { Text=gf.Label, Left=lx, Top=y+2, Width=lw, Height=fh, Font=Theme.FontBold, ForeColor=Theme.TextSecondary, TextAlign=ContentAlignment.MiddleLeft, BackColor=Color.Transparent, Parent=_extraPanel };
                    var tb = new TextBox { Left=fx, Top=y, Width=fw, Height=fh, BorderStyle=BorderStyle.FixedSingle, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, Font=Theme.FontBody, Parent=_extraPanel };
                    _extraBoxes.Add((gf.Key, tb));
                    y += rowH;
                }
                hasAny = true;
            }

            _extraPanel.Height = hasAny ? y+4 : 0;
        }

        private void BtnAddComponent_Click(object? sender, EventArgs e)
        {
            if (_addCat.SelectedItem==null || string.IsNullOrWhiteSpace(_addDesc.Text) || string.IsNullOrWhiteSpace(_addPart.Text))
            { MessageBox.Show("Category, Description and Part Number are required.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

            var comp = new Component
            {
                FekaNo=_addFekaNo.Text.Trim(),
                Category=_addCat.SelectedItem.ToString()!, Description=_addDesc.Text.Trim(), PartNumber=_addPart.Text.Trim(),
                Manufacturer=_addMfr.Text.Trim(), Footprint=_addFoot.Text.Trim(), LibraryRef=_addLibRef.Text.Trim(),
                FootprintRef=_addFootRef.Text.Trim(), IsCustom=true, Status=ComponentStatus.Pending,
            };
            if (Enum.TryParse<ComponentRating>(_addRating.SelectedItem?.ToString(), out var r)) comp.Rating = r;
            if (Enum.TryParse<ComponentProcess>(_addProcess.SelectedItem?.ToString(), out var p)) comp.Process = p;
            foreach (var (key, ctrl) in _extraBoxes)
            {
                string val = ctrl switch { ComboBox cb=>cb.Text.Trim(), TextBox tb=>tb.Text.Trim(), _=>"" };
                if (!string.IsNullOrWhiteSpace(val)) comp.ExtraFields[key] = val;
            }
            ComponentRepository.Add(comp);
            RefreshApprovalBadge();
            MessageBox.Show($"Component '{comp.Description}' added successfully!", "Added", MessageBoxButtons.OK, MessageBoxIcon.Information);
            ClearAddForm();
        }

        private void ClearAddForm()
        {
            _addFekaNo.Text = "";
            _addCat.SelectedIndex = -1;
            _addDesc.Text = _addPart.Text = _addFootRef.Text = "";
            _addMfr.Text = _addFoot.Text = "";
            if (_addLibRef.Items.Count > 0) _addLibRef.SelectedIndex = -1;
            _addRating.SelectedIndex = _addProcess.SelectedIndex = 0;
            _extraPanel.Controls.Clear(); _extraBoxes.Clear(); _extraPanel.Height = 0;
        }

        // ══════════════════════════════════════════════════════════════════
        //  APPROVAL
        // ══════════════════════════════════════════════════════════════════
        private void BuildApproval()
        {
            _pageApproval = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgMain, Visible = false };
            _content.Controls.Add(_pageApproval);

            var infoBar = new Panel { Left=0, Top=0, Height=40, BackColor=Color.FromArgb(254,243,199), Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right };
            new Label { Text="Double-click cells to edit. Pending components require FekaNo and PartNumber before approval.", ForeColor=Color.FromArgb(120,80,4), Font=Theme.FontBody, AutoSize=true, Left=12, Top=10, Parent=infoBar };
            _pageApproval.Controls.Add(infoBar);
            _pageApproval.Resize += (s, e) => infoBar.Width = _pageApproval.ClientSize.Width;

            // Kontrol paneli
            var btnPanel = new Panel { Left=0, Top=50, Height=32, BackColor=Theme.BgMain, Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right };
            _pageApproval.Controls.Add(btnPanel);

            var btnApproveAll = new ModernButton { Text="Approve All", Style=ModernButton.BtnStyle.Success, Width=120, Height=32, Left=0, Top=0, Parent=btnPanel };
            btnApproveAll.Click += (s, e) =>
            {
                var pending = ComponentRepository.GetAll().Where(c => c.IsCustom && c.Status == ComponentStatus.Pending).ToList();
                if (pending.Count == 0) { MessageBox.Show("No components pending approval.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
                if (MessageBox.Show($"Are you sure you want to approve {pending.Count} component(s)?", "Approve All", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    foreach (var c in pending) ComponentRepository.Approve(c.Id);
                    RefreshApproval(); RefreshApprovalBadge();
                    MessageBox.Show($"{pending.Count} component(s) approved and moved to catalog.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            var btnApproveSelected = new ModernButton { Text="Approve Selected", Style=ModernButton.BtnStyle.Success, Width=140, Height=32, Left=130, Top=0, Parent=btnPanel };
            btnApproveSelected.Click += (s, e) =>
            {
                var selected = _approvalGrid.SelectedRows.Cast<DataGridViewRow>().ToList();
                if (selected.Count == 0) { MessageBox.Show("Please select components to approve.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

                // FekaNo ve PartNumber validasyon
                var invalidRows = new List<int>();
                foreach (var row in selected)
                {
                    var fekaNo = row.Cells["FekaNo"].Value?.ToString()?.Trim() ?? "";
                    var partNum = row.Cells["PartNumber"].Value?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(fekaNo) || string.IsNullOrEmpty(partNum))
                        invalidRows.Add(_approvalGrid.Rows.IndexOf(row));
                }

                if (invalidRows.Count > 0)
                {
                    MessageBox.Show($"Cannot approve components with empty FekaNo or PartNumber.\n\nRow(s): {string.Join(", ", invalidRows.Select(i => i + 1))}\n\nPlease fill in missing values first.", 
                        "Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (MessageBox.Show($"Approve {selected.Count} component(s)?", "Approve Selected", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    int approvedCount = 0;
                    foreach (var row in selected)
                    {
                        if (int.TryParse(row.Cells["ID"].Value?.ToString(), out int id))
                        {
                            ComponentRepository.Approve(id);
                            approvedCount++;
                        }
                    }
                    RefreshApproval(); RefreshApprovalBadge();
                    MessageBox.Show($"{approvedCount} component(s) approved and moved to catalog.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            var btnDeleteSelected = new ModernButton { Text="Delete Selected", Style=ModernButton.BtnStyle.Danger, Width=130, Height=32, Left=280, Top=0, Parent=btnPanel };
            btnDeleteSelected.Click += (s, e) =>
            {
                var selected = _approvalGrid.SelectedRows.Cast<DataGridViewRow>().ToList();
                if (selected.Count == 0) { MessageBox.Show("Please select components to delete.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }

                if (MessageBox.Show($"Delete {selected.Count} component(s)? This cannot be undone.", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    int deletedCount = 0;
                    foreach (var row in selected)
                    {
                        if (int.TryParse(row.Cells["ID"].Value?.ToString(), out int id))
                        {
                            ComponentRepository.Delete(id);
                            deletedCount++;
                        }
                    }
                    RefreshApproval(); RefreshApprovalBadge();
                    MessageBox.Show($"{deletedCount} component(s) deleted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            };

            _approvalGrid = BuildGrid();
            _approvalGrid.ReadOnly = false;  // ✅ Düzenleme için ReadOnly kapalı!
            
            // Düzenlenebilir sütunları false yap
            foreach (DataGridViewColumn col in _approvalGrid.Columns)
            {
                if (col.Name == "FekaNo" || col.Name == "PartNumber" || 
                    col.Name == "Manufacturer" || col.Name == "Rating" || col.Name == "Process")
                    col.ReadOnly = false;
                else
                    col.ReadOnly = true;
            }
            
            _approvalGrid.Left=0; _approvalGrid.Top=90;
            _approvalGrid.Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right|AnchorStyles.Bottom;
            _approvalGrid.Width=_pageApproval.ClientSize.Width;
            _approvalGrid.Height=_pageApproval.ClientSize.Height-90;
            _approvalGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _approvalGrid.MultiSelect = true;  // ✅ Çoklu seçim açık!
            _approvalGrid.AllowUserToAddRows = false;

            RefreshApproval();

            // Çift tıklama ile inline düzenleme
            _approvalGrid.CellDoubleClick += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                var row = _approvalGrid.Rows[e.RowIndex];
                var colName = _approvalGrid.Columns[e.ColumnIndex].Name;

                // Düzenlenebilir sütunlar
                if (colName == "FekaNo" || colName == "PartNumber" || colName == "Manufacturer" || colName == "Rating" || colName == "Process")
                {
                    _approvalGrid.BeginEdit(false);
                }
            };

            // Cell düzenlemesi bittikten sonra kaydet
            _approvalGrid.CellEndEdit += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                var row = _approvalGrid.Rows[e.RowIndex];
                if (!int.TryParse(row.Cells["ID"].Value?.ToString(), out int id)) return;

                var component = ComponentRepository.GetAll().FirstOrDefault(c => c.Id == id);
                if (component == null) return;

                var colName = _approvalGrid.Columns[e.ColumnIndex].Name;
                var newValue = row.Cells[e.ColumnIndex].Value?.ToString() ?? "";

                // Güncelle
                if (colName == "FekaNo") component.FekaNo = newValue;
                else if (colName == "PartNumber") component.PartNumber = newValue;
                else if (colName == "Manufacturer") component.Manufacturer = newValue;
                else if (colName == "Rating" && Enum.TryParse<ComponentRating>(newValue, out var rating)) component.Rating = rating;
                else if (colName == "Process" && Enum.TryParse<ComponentProcess>(newValue, out var process)) component.Process = process;

                // Veritabanına kaydet
                ComponentRepository.Update(component);
                
                // Grid yenilemeyi defer et (event tamamlandıktan sonra çalışacak)
                this.BeginInvoke(new Action(() => RefreshApproval()));
            };

            _pageApproval.Controls.Add(_approvalGrid);
            _pageApproval.Resize += (s, e) =>
            {
                _approvalGrid.Width=_pageApproval.ClientSize.Width;
                _approvalGrid.Height=_pageApproval.ClientSize.Height-90;
                btnPanel.Width = _pageApproval.ClientSize.Width;
            };
        }

        private void RefreshApproval()
        {
            // Only pending items
            var pendingOnly = ComponentRepository.GetAll()
                .Where(c => c.IsCustom && c.Status == ComponentStatus.Pending)
                .ToList();

            var table = new System.Data.DataTable();
            foreach (var col in new[] { "FekaNo", "ID", "Category", "Description", "PartNumber", "Manufacturer", "Footprint", "Rating", "Process", "Status" })
                table.Columns.Add(col);

            foreach (var c in pendingOnly)
                table.Rows.Add(c.FekaNo, c.Id, c.Category, c.Description, c.PartNumber, c.Manufacturer, c.Footprint, c.Rating, c.Process, c.Status);

            _approvalGrid.DataSource = table;
        }

        private void RefreshApprovalBadge()
        {
            if (_navApproval != null)
            {
                _navApproval.BadgeCount = ComponentRepository.GetAll().Count(c => c.Status == ComponentStatus.Pending);
                _navApproval.Invalidate();
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  EXPORT
        // ══════════════════════════════════════════════════════════════════
        private void BuildExport()
        {
            _pageExport = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgMain, Visible = false };
            _content.Controls.Add(_pageExport);

            var card = new Panel { Left=0, Top=0, Width=620, Height=400, BackColor=Theme.BgSurface, Anchor=AnchorStyles.Top|AnchorStyles.Left };
            card.Paint += (s, e) => { using var pen=new Pen(Theme.Border,1); e.Graphics.DrawRectangle(pen,0,0,card.Width-1,card.Height-1); };
            _pageExport.Controls.Add(card);

            new Label { Text="Export Options", Font=Theme.FontH2, ForeColor=Theme.TextPrimary, Left=24, Top=20, AutoSize=true, Parent=card };
            new Label { Text="Category filter (optional):", Font=Theme.FontBold, ForeColor=Theme.TextSecondary, Left=24, Top=68, AutoSize=true, Parent=card };

            _exportCat = new ComboBox { Left=24, Top=92, Width=200, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=card };
            _exportCat.Items.Add("All Categories");
            foreach (var cat in ComponentRepository.AllCategories.OrderBy(c => c)) _exportCat.Items.Add(cat);
            _exportCat.SelectedIndex = 0;

            new Label { Text="Excel / Spreadsheet:", Font=Theme.FontBold, ForeColor=Theme.TextSecondary, Left=24, Top=136, AutoSize=true, Parent=card };

            // E-BOM export button
            var btnExportEBom = new ModernButton { Text="Export E-BOM (xlsx,xls,ods,csv,tsv,json)", Style=ModernButton.BtnStyle.Primary, Width=200, Height=70, Left=24,  Top=180, Parent=card };

            btnExportEBom.Click += BtnExportEBom_Click;
            new Panel { Left=24, Top=292, Width=572, Height=1, BackColor=Theme.Border, Parent=card };
            new Label { Text="Import CSV", Font=Theme.FontH2, ForeColor=Theme.TextPrimary, Left=24, Top=306, AutoSize=true, Parent=card };
            var btnImportCsv = new ModernButton { Text="Import CSV File…", Style=ModernButton.BtnStyle.Secondary, Width=160, Height=34, Left=24, Top=338, Parent=card };
            btnImportCsv.Click += BtnImportCsv_Click;
        }

        private enum ExportFormat { Csv, Tsv, Json, Xml, Xls, Ods }

        // E-BOM unified export dialog
        private void BtnExportEBom_Click(object? sender, EventArgs e)
        {
            var cat = _exportCat.SelectedIndex > 0 ? _exportCat.SelectedItem?.ToString() ?? "" : "";
            var components = ComponentRepository.Search("", cat, "", "", "");

            using var dlg = new SaveFileDialog
            {
                Title      = "Export E-BOM",
                Filter     = "Excel Workbook (*.xlsx)|*.xlsx" +
                             "|Excel 97-2003 (*.xls)|*.xls" +
                             "|CSV (*.csv)|*.csv" +
                             "|Tab Separated (*.tsv)|*.tsv" +
                             "|JSON (*.json)|*.json" +
                             "|XML (*.xml)|*.xml",
                FilterIndex = 1,
                FileName   = "E-BOM_List",
            };

            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                switch (dlg.FilterIndex)
                {
                    case 1: // xlsx - E-BOM template
                        ExportToEBomExcel(dlg.FileName, components);
                        break;
                    case 2: // xls
                        ExportDelimitedToPath(ExportFormat.Xls, dlg.FileName, components);
                        break;
                    case 3: // csv
                        ExportDelimitedToPath(ExportFormat.Csv, dlg.FileName, components);
                        break;
                    case 4: // tsv
                        ExportDelimitedToPath(ExportFormat.Tsv, dlg.FileName, components);
                        break;
                    case 5: // json
                        ExportDelimitedToPath(ExportFormat.Json, dlg.FileName, components);
                        break;
                    case 6: // xml
                        ExportDelimitedToPath(ExportFormat.Xml, dlg.FileName, components);
                        break;
                }
                MessageBox.Show($"Exported {components.Count} records successfully.", "Export Complete",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Separate format buttons
        private void ExportDelimited(ExportFormat fmt)
        {
            var cat = _exportCat.SelectedIndex > 0 ? _exportCat.SelectedItem?.ToString() ?? "" : "";
            var components = ComponentRepository.Search("", cat, "", "", "");

            string filter = fmt switch
            {
                ExportFormat.Tsv  => "TSV files|*.tsv",
                ExportFormat.Json => "JSON files|*.json",
                ExportFormat.Xml  => "XML files|*.xml",
                ExportFormat.Xls  => "Excel 97-2003|*.xls",
                ExportFormat.Ods  => "OpenDocument Spreadsheet|*.ods",
                _                 => "CSV files|*.csv",
            };
            string defaultName = fmt switch
            {
                ExportFormat.Tsv  => "components_export.tsv",
                ExportFormat.Json => "components_export.json",
                ExportFormat.Xml  => "components_export.xml",
                ExportFormat.Xls  => "components_export.xls",
                ExportFormat.Ods  => "components_export.ods",
                _                 => "components_export.csv",
            };

            using var dlg = new SaveFileDialog { Filter = filter, FileName = defaultName };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                ExportDelimitedToPath(fmt, dlg.FileName, components);
                MessageBox.Show($"Exported {components.Count} records successfully.", "Export Complete",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Export failed:\n{ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Core write logic
        private static void ExportDelimitedToPath(ExportFormat fmt, string filePath, List<Component> components)
        {
            string Esc(string? s) => $"\"{(s ?? "").Replace("\"", "\"\"")}\"";

            switch (fmt)
            {
                case ExportFormat.Csv:
                case ExportFormat.Xls:
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("FekaNo,Category,Description,PartNumber,Manufacturer,Footprint,LibraryRef,FootprintRef,Rating,Process,Status");
                    foreach (var c in components)
                        sb.AppendLine($"{Esc(c.FekaNo)},{Esc(c.Category)},{Esc(c.Description)},{Esc(c.PartNumber)},{Esc(c.Manufacturer)},{Esc(c.Footprint)},{Esc(c.LibraryRef)},{Esc(c.FootprintRef)},{Esc(c.Rating.ToString())},{Esc(c.Process.ToString())},{Esc(c.Status.ToString())}");
                    System.IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8);
                    break;
                }
                case ExportFormat.Tsv:
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("FekaNo\tCategory\tDescription\tPartNumber\tManufacturer\tFootprint\tLibraryRef\tFootprintRef\tRating\tProcess\tStatus");
                    foreach (var c in components)
                        sb.AppendLine($"{c.FekaNo}\t{c.Category}\t{c.Description}\t{c.PartNumber}\t{c.Manufacturer}\t{c.Footprint}\t{c.LibraryRef}\t{c.FootprintRef}\t{c.Rating}\t{c.Process}\t{c.Status}");
                    System.IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8);
                    break;
                }
                case ExportFormat.Json:
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("[");
                    for (int i = 0; i < components.Count; i++)
                    {
                        var c = components[i];
                        string comma = i < components.Count - 1 ? "," : "";
                        sb.AppendLine($"  {{\"FekaNo\":{System.Text.Json.JsonSerializer.Serialize(c.FekaNo)},\"Category\":{System.Text.Json.JsonSerializer.Serialize(c.Category)},\"Description\":{System.Text.Json.JsonSerializer.Serialize(c.Description)},\"PartNumber\":{System.Text.Json.JsonSerializer.Serialize(c.PartNumber)},\"Manufacturer\":{System.Text.Json.JsonSerializer.Serialize(c.Manufacturer)},\"Footprint\":{System.Text.Json.JsonSerializer.Serialize(c.Footprint)},\"LibraryRef\":{System.Text.Json.JsonSerializer.Serialize(c.LibraryRef)},\"FootprintRef\":{System.Text.Json.JsonSerializer.Serialize(c.FootprintRef)},\"Rating\":\"{c.Rating}\",\"Process\":\"{c.Process}\",\"Status\":\"{c.Status}\"}}{comma}");
                    }
                    sb.AppendLine("]");
                    System.IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8);
                    break;
                }
                case ExportFormat.Xml:
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
                    sb.AppendLine("<Components>");
                    foreach (var c in components)
                    {
                        string X(string? v) => System.Security.SecurityElement.Escape(v ?? "");
                        sb.AppendLine($"  <Component><FekaNo>{X(c.FekaNo)}</FekaNo><Category>{X(c.Category)}</Category><Description>{X(c.Description)}</Description><PartNumber>{X(c.PartNumber)}</PartNumber><Manufacturer>{X(c.Manufacturer)}</Manufacturer><Footprint>{X(c.Footprint)}</Footprint><LibraryRef>{X(c.LibraryRef)}</LibraryRef><FootprintRef>{X(c.FootprintRef)}</FootprintRef><Rating>{c.Rating}</Rating><Process>{c.Process}</Process><Status>{c.Status}</Status></Component>");
                    }
                    sb.AppendLine("</Components>");
                    System.IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8);
                    break;
                }
                case ExportFormat.Ods:
                {
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine("FekaNo,Category,Description,PartNumber,Manufacturer,Footprint,LibraryRef,FootprintRef,Rating,Process,Status");
                    foreach (var c in components)
                        sb.AppendLine($"{Esc(c.FekaNo)},{Esc(c.Category)},{Esc(c.Description)},{Esc(c.PartNumber)},{Esc(c.Manufacturer)},{Esc(c.Footprint)},{Esc(c.LibraryRef)},{Esc(c.FootprintRef)},{Esc(c.Rating.ToString())},{Esc(c.Process.ToString())},{Esc(c.Status.ToString())}");
                    System.IO.File.WriteAllText(filePath, sb.ToString(), System.Text.Encoding.UTF8);
                    break;
                }
            }
        }

        private void BtnExportCsv_Click(object? sender, EventArgs e) => ExportDelimited(ExportFormat.Csv);

        private void BtnImportCsv_Click(object? sender, EventArgs e)
        {
            using var dlg = new OpenFileDialog { Filter = "CSV files|*.csv" };
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                var csv = System.IO.File.ReadAllText(dlg.FileName);
                int count = ComponentRepository.ImportCsv(csv);
                RefreshApprovalBadge();
                MessageBox.Show($"Imported {count} components (pending review).", "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BtnExportExcel_Click(object? sender, EventArgs e)
        {
            var cat = _exportCat.SelectedIndex > 0 ? _exportCat.SelectedItem?.ToString() ?? "" : "";
            var components = ComponentRepository.Search("", cat, "", "", "");

            using var dlg = new SaveFileDialog { Filter = "Excel files|*.xlsx", FileName = "E-BOM_List.xlsx" };
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                ExportToEBomExcel(dlg.FileName, components);
                MessageBox.Show($"Exported {components.Count} records successfully.", "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show($"Export failed:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        }

        private static void ExportToEBomExcel(string filePath, List<Component> components)
        {
            WriteEBomXlsx(filePath, components);
        }

       private static void WriteEBomXlsx(string path, List<Component> data)
{
    string templatePath = System.IO.Path.Combine(
        Application.StartupPath, "Data2", "AR-GE_EE_Birimi_E-BOM_Listesi_Formu.xlsx");

    if (!System.IO.File.Exists(templatePath))
    {
        MessageBox.Show("Şablon bulunamadı:\n" + templatePath, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
        return;
    }

    try
    {
        System.IO.File.Copy(templatePath, path, true);

        using var workbook = new XLWorkbook(path);
        var ws = workbook.Worksheet("E-BOM");

        // Tüm hyperlinkleri temizle
        foreach (var hyperlink in ws.Hyperlinks.ToList())
            ws.Hyperlinks.Delete(hyperlink);

       
     

        // Veriyi yaz
        int currentRow = 6;
        foreach (var item in data)
        {
            // Stil kopyala (sadece ilk veri satırından)
            if (currentRow > 6)
                CopyRowFormat(ws, 6, currentRow, 2, 13);

            // Hücreleri sıfırla
            for (int col = 2; col <= 13; col++)
            {
                ws.Cell(currentRow, col).Style.Font.FontColor = XLColor.Black;
                ws.Cell(currentRow, col).Style.Font.Underline = XLFontUnderlineValues.None;
            }

            // NO kolonu boş (col 2)
            ws.Cell(currentRow, 2).Value = Blank.Value;
            ws.Cell(currentRow, 3).Value  = item.FekaNo       ?? "";
            ws.Cell(currentRow, 4).Value  = item.Category     ?? "";
            ws.Cell(currentRow, 5).Value  = item.Description  ?? "";
            ws.Cell(currentRow, 6).Value  = item.Footprint    ?? "";
            ws.Cell(currentRow, 7).Value  = item.PartNumber   ?? "";
            ws.Cell(currentRow, 8).Value  = item.Manufacturer ?? "";
            ws.Cell(currentRow, 9).Value  = item.Rating  == ComponentRating.AECQ ? "AECQ" : "STD";
            ws.Cell(currentRow, 10).Value = item.Process == ComponentProcess.THT  ? "THT"  : "SMD";
            // col 11 (Layer), 12 (Quantity), 13 (Designator) boş bırak

            ws.Row(currentRow).Height = 18;
            currentRow++;
        }

        ws.Column(5).Width = 35;
        ws.Column(7).Width = 22;
        ws.Column(8).Width = 16;

        workbook.Save();
    }
    catch (Exception ex)
    {
        MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
private static void CopyRowFormat(IXLWorksheet ws, int srcRow, int dstRow, int c1, int c2)
{
    for (int col = c1; col <= c2; col++)
    {
        var src = ws.Cell(srcRow, col);
        var dst = ws.Cell(dstRow, col);
        dst.Style.Font.FontName            = src.Style.Font.FontName;
        dst.Style.Font.FontSize            = src.Style.Font.FontSize;
        dst.Style.Font.Bold                = src.Style.Font.Bold;
        dst.Style.Alignment.Horizontal     = src.Style.Alignment.Horizontal;
        dst.Style.Alignment.Vertical       = src.Style.Alignment.Vertical;
        dst.Style.Border.LeftBorder        = src.Style.Border.LeftBorder;
        dst.Style.Border.RightBorder       = src.Style.Border.RightBorder;
        dst.Style.Border.TopBorder         = src.Style.Border.TopBorder;
        dst.Style.Border.BottomBorder      = src.Style.Border.BottomBorder;
        dst.Style.Border.LeftBorderColor   = src.Style.Border.LeftBorderColor;
        dst.Style.Border.RightBorderColor  = src.Style.Border.RightBorderColor;
        dst.Style.Border.TopBorderColor    = src.Style.Border.TopBorderColor;
        dst.Style.Border.BottomBorderColor = src.Style.Border.BottomBorderColor;
        dst.Style.Fill.BackgroundColor     = src.Style.Fill.BackgroundColor;
    }
}

        // ══════════════════════════════════════════════════════════════════
        //  CATEGORY MANAGER
        // ══════════════════════════════════════════════════════════════════
        private void BuildCategoryManager()
{
    _pageCatManager = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgMain, Visible = false };
    _content.Controls.Add(_pageCatManager);

    // --- LEFT CARD (Kategoriler) ---
    var leftCard = new Panel
    {
        Left = 0, Top = 0, Width = 220,
        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom,
        BackColor = Theme.BgSurface,
    };
    leftCard.Paint += (s, e) => { using var pen = new Pen(Theme.Border, 1); e.Graphics.DrawRectangle(pen, 0, 0, leftCard.Width - 1, leftCard.Height - 1); };
    _pageCatManager.Controls.Add(leftCard);
    _pageCatManager.Resize += (s, e) => leftCard.Height = _pageCatManager.ClientSize.Height;

    new Label { Text = "Categories", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Theme.TextSecondary, Left = 16, Top = 16, AutoSize = true, Parent = leftCard };

    _catList = new ListBox
    {
        Left = 8, Top = 44, Width = 204,
        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom,
        BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary,
        BorderStyle = BorderStyle.None, Font = Theme.FontBody, Parent = leftCard,
    };
    // Height adjusted to avoid overlap with bottom inputs
    leftCard.Resize += (s, e) => _catList.Height = leftCard.Height - 100;

    _catList.SelectedIndexChanged += (s, e) =>
    {
        RefreshFieldList();
    };

    var txtNewCat = new TextBox { Left = 8, Width = 152, Height = 28, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, PlaceholderText="New category...", Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = leftCard };
    var btnAddCat = new ModernButton { Text = "+", Style = ModernButton.BtnStyle.Primary, Width = 44, Height = 28, Left = 164, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = leftCard };
    
    // Vertical alignment for txtNewCat and btnAddCat
    leftCard.Resize += (s, e) => { 
        txtNewCat.Top = leftCard.Height - 40; 
        btnAddCat.Top = leftCard.Height - 40; 
    };

    btnAddCat.Click += (s, e) =>
    {
        var name = txtNewCat.Text.Trim().ToUpper();
        if (string.IsNullOrWhiteSpace(name)) return;
        if (ComponentRepository.CategoryFields.ContainsKey(name))
        { MessageBox.Show($"'{name}' already exists.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        ComponentRepository.AddCategory(name);
        txtNewCat.Text = "";
        RefreshCategoryManager();
        var idx = _catList.Items.IndexOf(name);
        if (idx >= 0) _catList.SelectedIndex = idx;
    };

    // --- MIDDLE CARD (Alanlar) ---
    const int MID_LEFT = 228;
    const int MID_W = 280;

    var midCard = new Panel
    {
        Left = MID_LEFT, Top = 0, Width = MID_W,
        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom,
        BackColor = Theme.BgSurface,
    };
    midCard.Paint += (s, e) => { using var pen = new Pen(Theme.Border, 1); e.Graphics.DrawRectangle(pen, 0, 0, midCard.Width - 1, midCard.Height - 1); };
    _pageCatManager.Controls.Add(midCard);
    _pageCatManager.Resize += (s, e) => midCard.Height = _pageCatManager.ClientSize.Height;

    // Title width narrowed to avoid button overlap
    _lblCatTitle = new Label { Text = "Select a category", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Theme.TextSecondary, Left = 16, Top = 16, Width = MID_W - 125, AutoSize = false, Parent = midCard };

    // Delete Cat button aligned to the right
    var btnDelCat = new ModernButton { Text = "Delete Cat", Style = ModernButton.BtnStyle.Danger, Width = 90, Height = 26, Left = MID_W - 100, Top = 14, Parent = midCard };
    btnDelCat.Click += (s, e) =>
    {
        var sel = _catList.SelectedItem?.ToString();
        if (sel == null) return;
        if (MessageBox.Show($"Delete category '{sel}' and ALL its field definitions?\n(Components are NOT deleted)", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
        { ComponentRepository.DeleteCategory(sel); RefreshCategoryManager(); }
    };

    new Label { Text = "Fields", Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), ForeColor = Theme.TextMuted, Left = 16, Top = 50, AutoSize = true, Parent = midCard };

    _fieldList = new ListBox
    {
        Left = 8, Top = 72, Width = MID_W - 16,
        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom,
        BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary,
        BorderStyle = BorderStyle.None, Font = Theme.FontBody, Parent = midCard,
    };
    // List height adjusted to not overlap Remove Field button
    midCard.Resize += (s, e) => _fieldList.Height = midCard.Height - 210;

    _fieldList.SelectedIndexChanged += (s, e) =>
    {
        var catSel = _catList.SelectedItem?.ToString();
        var fieldSel = _fieldList.SelectedItem?.ToString();
        if (catSel == null || fieldSel == null) return;
    };

    // Remove Field button positioned below list
    var btnDelField = new ModernButton { Text = "Remove Field", Style = ModernButton.BtnStyle.Secondary, Width = 130, Height = 28, Left = 8, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = midCard };
    midCard.Resize += (s, e) => btnDelField.Top = midCard.Height - 135; 
    
    btnDelField.Click += (s, e) =>
    {
        var catSel = _catList.SelectedItem?.ToString();
        var fieldSel = _fieldList.SelectedItem?.ToString();
        if (catSel == null || fieldSel == null) return;
        ComponentRepository.RemoveFieldFromCategory(catSel, fieldSel.Trim());
        RefreshFieldList();
    };

    var sepLine = new Panel { Left = 8, Height = 1, BackColor = Theme.Border, Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right, Parent = midCard };
    
    var lblAddField = new Label { Text = "Add Field:", Font = Theme.FontBold, ForeColor = Theme.TextSecondary, Left = 8, Width = 70, Height = 28, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = midCard };
    var txtFieldKey = new TextBox { Width = MID_W - 90, Height = 28, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, PlaceholderText = "Field key (e.g. Voltage(V))", Left = 82, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = midCard };
    var btnAddField = new ModernButton { Text = "Add", Style = ModernButton.BtnStyle.Success, Width = MID_W - 16, Height = 28, Left = 8, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = midCard };

    // Bottom panel element positions clarified
    midCard.Resize += (s, e) =>
    {
        sepLine.Top = midCard.Height - 145;
        sepLine.Width = midCard.Width - 16;
        lblAddField.Top = midCard.Height - 100;
        txtFieldKey.Top = midCard.Height - 100;
        btnAddField.Top = midCard.Height - 60;
    };

    btnAddField.Click += (s, e) =>
    {
        var catSel = _catList.SelectedItem?.ToString();
        if (catSel == null) { MessageBox.Show("Please select a category first.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
        var key = txtFieldKey.Text.Trim();
        if (string.IsNullOrWhiteSpace(key)) { MessageBox.Show("Field key is required.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
        ComponentRepository.AddFieldToCategory(catSel, key, key);
        txtFieldKey.Text = "";
        RefreshFieldList();
    };

    // --- RIGHT CARD (Global Fields) ---
    const int RIGHT_LEFT = MID_LEFT + MID_W + 8;
    const int RIGHT_W = 280;

    var rightCard = new Panel
    {
        Left = RIGHT_LEFT, Top = 0, Width = RIGHT_W,
        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
        BackColor = Theme.BgSurface,
    };
    rightCard.Paint += (s, e) =>
    {
        using var pen = new Pen(Theme.Border, 1);
        e.Graphics.DrawRectangle(pen, 0, 0, rightCard.Width - 1, rightCard.Height - 1);
        using var brush = new SolidBrush(Theme.Accent);
        e.Graphics.FillRectangle(brush, 0, 0, rightCard.Width, 4);
    };
    _pageCatManager.Controls.Add(rightCard);
    _pageCatManager.Resize += (s, e) =>
    {
        rightCard.Width = Math.Max(200, _pageCatManager.ClientSize.Width - RIGHT_LEFT - 10);
        rightCard.Height = _pageCatManager.ClientSize.Height;
    };

    new Label { Text = "Global Fields", Left = 16, Top = 14, Width = 200, Height = 22,
        Font = new Font("Segoe UI", 10.5f, FontStyle.Bold), ForeColor = Theme.TextPrimary,
        BackColor = Color.Transparent, Parent = rightCard };
    new Label { Text = "Applied to all components across all categories", Left = 16, Top = 36, Width = 260, Height = 16,
        Font = Theme.FontSmall, ForeColor = Theme.TextMuted, BackColor = Color.Transparent, Parent = rightCard };

    _globalFieldList = new ListBox
    {
        Left = 16, Top = 60, Width = RIGHT_W - 32, Height = 260,
        BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary,
        Font = Theme.FontBody, BorderStyle = BorderStyle.FixedSingle,
        Parent = rightCard,
    };
    rightCard.Resize += (s, e) => _globalFieldList.Width = rightCard.Width - 32;

    void RefreshGlobalFieldList()
    {
        _globalFieldList.Items.Clear();
        foreach (var f in ComponentRepository.GlobalFields)
            _globalFieldList.Items.Add($"{f.Key}  ({f.Label})");
    }
    RefreshGlobalFieldList();

    // Add field row
    int ay = 336;
    new Label { Text = "Key", Left = 16, Top = ay + 4, Width = 50, Height = 28,
        Font = Theme.FontBold, ForeColor = Theme.TextSecondary, BackColor = Color.Transparent, Parent = rightCard };
    var tbGKey = new TextBox { Left = 70, Top = ay, Width = RIGHT_W - 86, Height = 30,
        BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary,
        Font = Theme.FontBody, PlaceholderText = "e.g. Voltage Rating", Parent = rightCard };
    rightCard.Resize += (s, e) => tbGKey.Width = rightCard.Width - 86;
    ay += 38;

    new Label { Text = "Label", Left = 16, Top = ay + 4, Width = 50, Height = 28,
        Font = Theme.FontBold, ForeColor = Theme.TextSecondary, BackColor = Color.Transparent, Parent = rightCard };
    var tbGLabel = new TextBox { Left = 70, Top = ay, Width = RIGHT_W - 86, Height = 30,
        BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary,
        Font = Theme.FontBody, PlaceholderText = "optional display name", Parent = rightCard };
    rightCard.Resize += (s, e) => tbGLabel.Width = rightCard.Width - 86;
    ay += 42;

    var btnAddGField = new ModernButton { Text = "Add Global Field", Style = ModernButton.BtnStyle.Primary,
        Left = 16, Top = ay, Width = RIGHT_W - 32, Height = 32, Parent = rightCard };
    rightCard.Resize += (s, e) => btnAddGField.Width = rightCard.Width - 32;
    btnAddGField.Click += (s, e) =>
    {
        var key = tbGKey.Text.Trim();
        var lbl = string.IsNullOrWhiteSpace(tbGLabel.Text) ? key : tbGLabel.Text.Trim();
        if (string.IsNullOrWhiteSpace(key)) return;
        ComponentRepository.AddGlobalField(key, lbl);
        tbGKey.Text = ""; tbGLabel.Text = "";
        RefreshGlobalFieldList();
    };
    ay += 40;

    var btnRemoveGField = new ModernButton { Text = "Remove Selected", Style = ModernButton.BtnStyle.Danger,
        Left = 16, Top = ay, Width = RIGHT_W - 32, Height = 32, Parent = rightCard };
    rightCard.Resize += (s, e) => btnRemoveGField.Width = rightCard.Width - 32;
    btnRemoveGField.Click += (s, e) =>
    {
        if (_globalFieldList.SelectedIndex < 0) return;
        var selected = _globalFieldList.SelectedItem?.ToString() ?? "";
        var key = selected.Contains("  (") ? selected[..selected.IndexOf("  (")] : selected;
        if (MessageBox.Show("Are you sure you want to remove the field '"+key+"'?", "Confirm Remove", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
        ComponentRepository.RemoveGlobalField(key);
        RefreshGlobalFieldList();
    };
}
        private void RefreshCategoryManager()
        {
            if (_catList == null) return;
            var prevSel = _catList.SelectedItem?.ToString();
            _catList.Items.Clear();

            var cats = ComponentRepository.AllCategories.OrderBy(c => c).ToList();
            foreach (var cat in cats)
                _catList.Items.Add(cat);

            if (prevSel != null)
            {
                var idx = _catList.Items.IndexOf(prevSel);
                if (idx >= 0) _catList.SelectedIndex = idx;
            }
            RefreshFieldList();
    
            if (_addCat != null)
            {
                var addCatSel = _addCat.SelectedItem?.ToString();
                _addCat.Items.Clear();
                foreach (var c in ComponentRepository.AllCategories.OrderBy(x => x)) _addCat.Items.Add(c);
                if (addCatSel != null) { var i = _addCat.Items.IndexOf(addCatSel); if (i>=0) _addCat.SelectedIndex=i; }
            }

            if (_exportCat != null)
            {
                var prevExportSel = _exportCat.SelectedItem?.ToString();
                _exportCat.Items.Clear();
                _exportCat.Items.Add("All Categories");
                foreach (var c in ComponentRepository.AllCategories.OrderBy(x => x)) _exportCat.Items.Add(c);
                int ei = _exportCat.Items.IndexOf(prevExportSel ?? "");
                _exportCat.SelectedIndex = ei >= 0 ? ei : 0;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        //  NEW CATEGORY DIALOG
        // ══════════════════════════════════════════════════════════════════

                private void ShowNewCategoryDialog()
        {
            var dlg = new Form
            {
                Text = "New Category",
                Size = new Size(480, 520),
                MinimumSize = new Size(400, 400),
                StartPosition = FormStartPosition.CenterParent,
                BackColor = Theme.BgMain,
                Font = Theme.FontBody,
                MinimizeBox = false, MaximizeBox = false,
            };

            // header stripe
            var stripe = new Panel { Left = 0, Top = 0, Width = dlg.ClientSize.Width, Height = 4, BackColor = Theme.Accent, Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right, Parent = dlg };
            dlg.Resize += (s, e) => stripe.Width = dlg.ClientSize.Width;

            int y = 20, lw = 120, fw = 280, fh = 32, lh = 28;

            // Category name
            new Label { Text = "Category Code *", Left = 20, Top = y, Width = lw, Height = lh, Font = Theme.FontBold, ForeColor = Theme.TextSecondary, TextAlign = ContentAlignment.MiddleLeft, Parent = dlg };
            var tbCatName = new TextBox { Left = 20 + lw + 8, Top = y, Width = fw, Height = fh, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, PlaceholderText = "e.g. RELAY, FUSE, OPTO", Parent = dlg };
            y += 44;

            // Separator + fields header
            new Panel { Left = 20, Top = y, Width = lw + fw + 8, Height = 1, BackColor = Theme.Border, Parent = dlg };
            y += 12;
            new Label { Text = "▸  Category Fields", Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), ForeColor = Theme.Accent, Left = 20, Top = y, Width = 300, Height = 24, BackColor = Color.Transparent, Parent = dlg };
            y += 28;
            new Label { Text = "Field name (e.g. Voltage(V), Current(A))", Font = Theme.FontSmall, ForeColor = Theme.TextMuted, Left = 20 + lw + 8, Top = y, Width = fw, Height = 18, Parent = dlg };
            y += 22;

            // Scrollable field list panel
            int fieldPanelTop = y;
            var fieldScroll = new Panel
            {
                Left = 20 + lw + 8, Top = fieldPanelTop, Width = fw,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom,
                BackColor = Theme.BgSurface2, AutoScroll = true,
                BorderStyle = BorderStyle.FixedSingle, Parent = dlg,
            };
            dlg.Resize += (s, e) => fieldScroll.Height = Math.Max(80, dlg.ClientSize.Height - fieldPanelTop - 100);
            fieldScroll.Height = Math.Max(80, dlg.ClientSize.Height - fieldPanelTop - 100);

            var fieldNames = new List<TextBox>();

            void AddFieldRow(string initial = "")
            {
                int rowH = 34, rowY = fieldNames.Count * rowH;
                var rowPanel = new Panel { Left = 0, Top = rowY, Width = fieldScroll.ClientSize.Width, Height = rowH, BackColor = Color.Transparent, Parent = fieldScroll };
                fieldScroll.Resize += (s, e) => rowPanel.Width = fieldScroll.ClientSize.Width;

                var tb = new TextBox { Left = 4, Top = 2, Width = rowPanel.Width - 40, Height = 28, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, Text = initial, Parent = rowPanel };
                rowPanel.Resize += (s, e) => tb.Width = rowPanel.Width - 40;

                var btnDel = new ModernButton { Text = "×", Style = ModernButton.BtnStyle.Danger, Width = 28, Height = 28, Left = rowPanel.Width - 32, Top = 2, Anchor = AnchorStyles.Top | AnchorStyles.Right, Parent = rowPanel };
                btnDel.Click += (s2, e2) =>
                {
                    int idx = fieldNames.IndexOf(tb);
                    if (idx < 0) return;
                    fieldNames.RemoveAt(idx);
                    fieldScroll.Controls.Remove(rowPanel);
                    rowPanel.Dispose();
                    for (int i = 0; i < fieldScroll.Controls.Count; i++)
                        fieldScroll.Controls[i].Top = i * rowH;
                };

                fieldNames.Add(tb);
            }

            // Start with 3 empty rows
            AddFieldRow(); AddFieldRow(); AddFieldRow();

            // "+ Add Field" button (left of the scroll panel)
            var btnAddField = new ModernButton
            {
                Text = "+ Add Field", Style = ModernButton.BtnStyle.Secondary,
                Width = lw, Height = 28, Left = 20,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = dlg,
            };
            dlg.Resize += (s, e) => btnAddField.Top = dlg.ClientSize.Height - 95;
            btnAddField.Top = dlg.ClientSize.Height - 95;
            btnAddField.Click += (s, e) => AddFieldRow();

            // Bottom buttons
            var sep2 = new Panel { Left = 0, Height = 1, BackColor = Theme.Border, Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right, Parent = dlg };
            dlg.Resize += (s, e) => { sep2.Top = dlg.ClientSize.Height - 58; sep2.Width = dlg.ClientSize.Width; };
            sep2.Top = dlg.ClientSize.Height - 58; sep2.Width = dlg.ClientSize.Width;

            var btnSave = new ModernButton { Text = "Create Category", Style = ModernButton.BtnStyle.Success, Width = 150, Height = 34, Left = 20, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = dlg };
            var btnCancel = new ModernButton { Text = "Cancel", Style = ModernButton.BtnStyle.Secondary, Width = 80, Height = 34, Left = 178, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = dlg };
            dlg.Resize += (s, e) => { btnSave.Top = dlg.ClientSize.Height - 46; btnCancel.Top = dlg.ClientSize.Height - 46; };
            btnSave.Top = dlg.ClientSize.Height - 46; btnCancel.Top = dlg.ClientSize.Height - 46;

            btnCancel.Click += (s, e) => dlg.Close();

            btnSave.Click += (s, e) =>
            {
                var catCode = tbCatName.Text.Trim().ToUpper();
                if (string.IsNullOrWhiteSpace(catCode))
                { MessageBox.Show("Category code is required.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                if (ComponentRepository.CategoryFields.ContainsKey(catCode))
                { MessageBox.Show($"'{catCode}' already exists.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                ComponentRepository.AddCategory(catCode);

                foreach (var tb in fieldNames)
                {
                    var key = tb.Text.Trim();
                    if (!string.IsNullOrWhiteSpace(key))
                        ComponentRepository.AddFieldToCategory(catCode, key, key);
                }

                // Refresh _addCat combobox
                var prevSel = _addCat.SelectedItem?.ToString();
                _addCat.Items.Clear();
                foreach (var c in ComponentRepository.AllCategories.OrderBy(x => x)) _addCat.Items.Add(c);
                var newIdx = _addCat.Items.IndexOf(catCode);
                _addCat.SelectedIndex = newIdx >= 0 ? newIdx : (prevSel != null ? _addCat.Items.IndexOf(prevSel) : -1);

                // Refresh _filterCat (Catalog filter)
                if (_filterCat != null)
                {
                    var prevFilter = _filterCat.SelectedItem?.ToString();
                    _filterCat.Items.Clear();
                    _filterCat.Items.Add("All Categories");
                    foreach (var c in ComponentRepository.AllCategories.OrderBy(x => x)) _filterCat.Items.Add(c);
                    int fi = _filterCat.Items.IndexOf(prevFilter ?? "");
                    _filterCat.SelectedIndex = fi >= 0 ? fi : 0;
                }

                RefreshCategoryManager();
                MessageBox.Show($"Category '{catCode}' created with {fieldNames.Count(f => !string.IsNullOrWhiteSpace(f.Text))} field(s).", "Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dlg.Close();
            };

            dlg.ShowDialog(this);
        }

        private void RefreshFieldList()
        {
            if (_fieldList == null || _catList == null) return;
            _fieldList.Items.Clear();
            var sel = _catList.SelectedItem?.ToString();
            if (sel == null) { if (_lblCatTitle!=null) _lblCatTitle.Text="Select a category"; return; }
            if (_lblCatTitle != null) _lblCatTitle.Text = $"Fields: {sel}";
            if (ComponentRepository.CategoryFields.TryGetValue(sel, out var fields))
                foreach (var (key, label) in fields)
                    _fieldList.Items.Add(key);
        }
    }
}