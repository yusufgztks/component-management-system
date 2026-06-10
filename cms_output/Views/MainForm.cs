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

        // ── Detail Panel ───────────────────────────────────────────────────
        private Panel       _detailPanel   = null!;
        private bool        _detailVisible = false;
        private TextBox     _dtTitle       = null!;
        private RichTextBox _dtPartNo      = null!;
        private RichTextBox _dtMfr         = null!;
        private RichTextBox _dtFoot        = null!;
        private RichTextBox _dtLibRef      = null!;
        private RichTextBox _dtFootRef     = null!;
        private RichTextBox _dtRating      = null!;
        private RichTextBox _dtProcess     = null!;
        private RichTextBox _dtStatus      = null!;
        private Panel       _dtExtraPanel  = null!;

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
        private Panel    _globalPanel = null!;
        private readonly List<(string Key, Control Input)> _globalBoxes = new();
        private Action   _reposition  = () => { };

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
        private Panel _fieldDetailPanel    = null!;
        private Label _lblFieldDetailTitle = null!;
        private Panel _fieldDetailContent  = null!;

        private SidebarItem? _navApproval;

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
            AddNavItem(flow, "", "catmanager", "Category Manager");

            AddNavSection(flow, "TOOLS");
            AddNavItem(flow, "", "export",     "Export / Import");

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
            return item;
        }

        private void BuildTopbar()
        {
            _lblPageTitle = new Label { Font = Theme.FontH2, ForeColor = Theme.TextPrimary, AutoSize = true, Left = 24, Top = 17 };
            _topbar.Controls.Add(_lblPageTitle);

            var btnExport = new ModernButton
            {
                Text = "Export CSV", Style = ModernButton.BtnStyle.Primary,
                Width = 110, Height = 32, Left = _topbar.Width - 148, Top = 12,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
            };
            btnExport.Click += (s, e) => Navigate("export");
            _topbar.Controls.Add(btnExport);
        }

        private void Navigate(string key)
        {
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

            _ = new Label { Text = "", Left = 10, Top = 14, AutoSize = true, Font = new Font("Segoe UI", 10f), BackColor = Color.Transparent, Parent = filterBar };

            _searchBox = new TextBox { Left = 34, Top = 11, Width = 210, Height = 28, BorderStyle = BorderStyle.FixedSingle, Font = Theme.FontBody, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, PlaceholderText = "Search by description, part no, manufacturer...", Parent = filterBar };
            _searchBox.TextChanged += (s, e) => RefreshCatalog();

            _filterCat = MakeCombo(252, 11, 120, filterBar, "All Categories");
            var sortedCats = ComponentRepository.AllCategories.OrderBy(c => c).ToArray();
            _filterCat.Items.AddRange(sortedCats.Cast<object>().ToArray());
            _filterCat.SelectedIndexChanged += (s, e) => RefreshCatalog();

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

            BuildDetailPanel();

            _pageCatalog.Resize += (s, e) =>
            {
                _grid.Height = _pageCatalog.ClientSize.Height - 76;
                RepositionDetailPanel();
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

        // ══════════════════════════════════════════════════════════════════
        //  DETAIL PANEL
        // ══════════════════════════════════════════════════════════════════
        private void BuildDetailPanel()
        {
            const int PW = 320;

            _detailPanel = new Panel
            {
                Width      = PW,
                BackColor  = Theme.BgSurface,
                Anchor     = AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom,
                Visible    = false,
                AutoScroll = true,
            };
            _detailPanel.Paint += (s, e) =>
            {
                using var pen = new Pen(Theme.Border, 1);
                e.Graphics.DrawRectangle(pen, 0, 0, _detailPanel.Width - 1, _detailPanel.Height - 1);
            };
            _pageCatalog.Controls.Add(_detailPanel);

            var topStripe = new Panel { Left = 0, Top = 0, Width = PW, Height = 4, BackColor = Theme.Accent };
            _detailPanel.Controls.Add(topStripe);

            var btnClose = new Label
            {
                Text = "✕", ForeColor = Theme.TextMuted,
                Font = new Font("Segoe UI", 11f), AutoSize = true,
                Cursor = Cursors.Hand, Left = PW - 28, Top = 12,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
            };
            btnClose.MouseEnter += (s, e) => ((Label)s!).ForeColor = Theme.Danger;
            btnClose.MouseLeave += (s, e) => ((Label)s!).ForeColor = Theme.TextMuted;
            btnClose.Click += (s, e) => CloseDetail();
            _detailPanel.Controls.Add(btnClose);

            var lblTag = new TextBox
            {
                Name = "lblTag", Left = 16, Top = 12, Width = PW - 52, Height = 18,
                Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
                ForeColor = Theme.Accent, BackColor = Theme.BgSurface,
                BorderStyle = BorderStyle.None, ReadOnly = true, TabStop = false, Cursor = Cursors.IBeam,
            };
            lblTag.MouseDown += (s, e) => ((TextBox)s!).SelectAll();
            _detailPanel.Controls.Add(lblTag);

            _dtTitle = new TextBox
            {
                Left = 16, Top = 34, Width = PW - 32, Height = 52,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Theme.TextPrimary, BackColor = Theme.BgSurface,
                BorderStyle = BorderStyle.None, ReadOnly = true, TabStop = false,
                Multiline = true, ScrollBars = ScrollBars.None, Cursor = Cursors.IBeam,
            };
            _dtTitle.MouseDown += (s, e) => _dtTitle.SelectAll();
            _detailPanel.Controls.Add(_dtTitle);

            _detailPanel.Controls.Add(new Panel { Left = 0, Top = 92, Width = PW, Height = 1, BackColor = Theme.Border });

            AddDetailSectionHeader(_detailPanel, "CORE INFORMATION", 100);

            int y = 122;
            _dtPartNo  = AddDetailRowNew(_detailPanel, "Part Number",   ref y);
            _dtMfr     = AddDetailRowNew(_detailPanel, "Manufacturer",  ref y);
            _dtFoot    = AddDetailRowNew(_detailPanel, "Footprint",     ref y);
            _dtLibRef  = AddDetailRowNew(_detailPanel, "Library Ref",   ref y);
            _dtFootRef = AddDetailRowNew(_detailPanel, "Footprint Ref", ref y);
            _dtRating  = AddDetailRowNew(_detailPanel, "Rating",        ref y);
            _dtProcess = AddDetailRowNew(_detailPanel, "Process",       ref y);
            _dtStatus  = AddDetailRowNew(_detailPanel, "Status",        ref y);

            _detailPanel.Controls.Add(new Panel { Left = 16, Top = y + 4, Width = PW - 32, Height = 1, BackColor = Theme.Border });

            AddDetailSectionHeader(_detailPanel, "TECHNICAL SPECS", y + 16);

            _dtExtraPanel = new Panel { Left = 0, Top = y + 40, Width = PW, Height = 0, BackColor = Color.Transparent };
            _detailPanel.Controls.Add(_dtExtraPanel);

            _grid.SelectionChanged += (s, e) =>
            {
                if (_grid.SelectedCells.Count > 0 && _detailVisible)
                {
                    var row = _grid.SelectedCells[0].OwningRow;
                    if (row != null) FillDetail(row);
                }
            };
            _grid.CellClick += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                ShowDetail(_grid.Rows[e.RowIndex]);
            };

            RepositionDetailPanel();
        }

        private RichTextBox AddDetailRowNew(Panel parent, string labelText, ref int y)
        {
            var rtb = new RichTextBox
            {
                Left = 16, Top = y, Width = 288, Height = 20,
                BorderStyle = BorderStyle.None, ReadOnly = true, TabStop = false,
                ScrollBars = RichTextBoxScrollBars.None,
                BackColor = Theme.BgSurface, Cursor = Cursors.IBeam, Tag = labelText,
            };
            rtb.AppendText(labelText + ": ");
            rtb.Select(0, labelText.Length + 2);
            rtb.SelectionColor = Theme.TextMuted;
            rtb.SelectionFont  = Theme.FontSmall;
            rtb.AppendText("—");
            rtb.Select(labelText.Length + 2, 1);
            rtb.SelectionColor = Theme.TextPrimary;
            rtb.SelectionFont  = Theme.FontBody;
            rtb.SelectionLength = 0;
            parent.Controls.Add(rtb);
            y += 24;
            return rtb;
        }

        private void DetailTextBox_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                if (sender is TextBox tb && !string.IsNullOrWhiteSpace(tb.Text) && tb.Text != "—")
                { Clipboard.SetText(tb.Text); e.Handled = true; }
            }
        }

        private static void SetDetailRow(RichTextBox rtb, string value, Color? valueColor = null, Font? valueFont = null)
        {
            string label = rtb.Tag as string ?? "";
            rtb.Clear();
            rtb.SelectionColor = Theme.TextMuted;
            rtb.SelectionFont  = Theme.FontSmall;
            rtb.AppendText(label + ": ");
            rtb.SelectionColor = valueColor ?? Theme.TextPrimary;
            rtb.SelectionFont  = valueFont  ?? Theme.FontBody;
            rtb.AppendText(string.IsNullOrEmpty(value) ? "—" : value);
            rtb.SelectionLength = 0;
        }

        private static void CopyLabelValue(string label, string value)
        {
            if (!string.IsNullOrWhiteSpace(value) && value != "—")
                Clipboard.SetText($"{label}: {value}");
        }

        private void AddDetailSectionHeader(Panel parent, string text, int top)
        {
            var tb = new TextBox
            {
                Text = text,
                Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 116, 150),
                BackColor = Theme.BgSurface,
                Left = 16, Top = top, Width = 280, Height = 18,
                BorderStyle = BorderStyle.None, ReadOnly = true, TabStop = false, Cursor = Cursors.IBeam,
            };
            tb.MouseDown += (s, e) => ((TextBox)s!).SelectAll();
            parent.Controls.Add(tb);
        }

        private void RepositionDetailPanel()
        {
            const int PW = 320;
            int pageW = _pageCatalog.ClientSize.Width;
            int pageH = _pageCatalog.ClientSize.Height;
            _detailPanel.Left   = pageW - PW;
            _detailPanel.Top    = 52;
            _detailPanel.Width  = PW;
            _detailPanel.Height = pageH - 52;
            _grid.Width = _detailVisible ? pageW - PW : pageW;
        }

        private void ShowDetail(DataGridViewRow row)
        {
            _detailVisible       = true;
            _detailPanel.Visible = true;
            RepositionDetailPanel();
            FillDetail(row);
        }

        private void FillDetail(DataGridViewRow row)
        {
            string Get(string col) => row.DataGridView?.Columns.Contains(col) == true
                ? row.Cells[col]?.Value?.ToString() ?? "—" : "—";

            var lblTag = _detailPanel.Controls.OfType<TextBox>().FirstOrDefault(l => l.Name == "lblTag");
            if (lblTag != null)
                lblTag.Text = $"{Get("Category")} – {GetCategoryFullName(Get("Category"))}";

            _dtTitle.Text = Get("Description");

            int.TryParse(Get("ID"), out int id);
            var comp = ComponentRepository.GetAll().FirstOrDefault(c => c.Id == id);

            SetDetailRow(_dtPartNo,  Get("PartNumber"));
            SetDetailRow(_dtMfr,     Get("Manufacturer"));
            SetDetailRow(_dtFoot,    Get("Footprint"));
            SetDetailRow(_dtLibRef,  comp?.LibraryRef  is { Length: > 0 } lr ? lr : "—");
            SetDetailRow(_dtFootRef, comp?.FootprintRef is { Length: > 0 } fr ? fr : "—");

            string rating = Get("Rating");
            SetDetailRow(_dtRating, rating, rating == "AECQ" ? Theme.Accent : Theme.TextPrimary);
            SetDetailRow(_dtProcess, Get("Process"));

            string status = Get("Status");
            SetDetailRow(_dtStatus, status,
                status == "Approved" ? Theme.AccentGreen : Theme.Warning,
                new Font("Segoe UI", 8.5f, FontStyle.Bold));

            _dtExtraPanel.Controls.Clear();
            int ey = 0;
            if (comp != null && comp.ExtraFields.Count > 0)
            {
                foreach (var kvp in comp.ExtraFields)
                {
                    var rtb = new RichTextBox
                    {
                        Left = 16, Top = ey, Width = 288, Height = 20,
                        BorderStyle = BorderStyle.None, ReadOnly = true, TabStop = false,
                        ScrollBars = RichTextBoxScrollBars.None,
                        BackColor = Theme.BgSurface, Cursor = Cursors.IBeam, Parent = _dtExtraPanel,
                    };
                    rtb.SelectionColor = Theme.TextMuted;
                    rtb.SelectionFont  = Theme.FontSmall;
                    rtb.AppendText(kvp.Key + ": ");
                    rtb.SelectionColor = Theme.Accent;
                    rtb.SelectionFont  = Theme.FontBody;
                    rtb.AppendText(string.IsNullOrEmpty(kvp.Value) ? "—" : kvp.Value);
                    rtb.SelectionLength = 0;
                    ey += 24;
                }
            }
            _dtExtraPanel.Height = ey + 8;
        }

        private void CloseDetail()
        {
            _detailVisible       = false;
            _detailPanel.Visible = false;
            RepositionDetailPanel();
        }

        private static string GetCategoryFullName(string shortName) => shortName switch
        {
            "RES"     => "Resistor",
            "CAP"     => "Capacitor",
            "IND"     => "Inductor",
            "FB"      => "Ferrite Bead",
            "CMC"     => "Common Mode Choke",
            "DIO-REC" => "Diode Rectifier",
            "DIO-SCH" => "Diode Schottky",
            "DIO-ZEN" => "Diode Zener",
            "DIO-SW"  => "Diode Switching",
            "DIO-TVS" => "Diode TVS",
            "TRAN"    => "Transistor",
            "MOS"     => "MOSFET",
            "DIODE"   => "Diode",
            "LED"     => "LED",
            "IC"      => "Integrated Circuit",
            "CON"     => "Connector",
            "RELAY"   => "Relay",
            "FUSE"    => "Fuse",
            "DC-DC"   => "DC-DC Converter",
            "LED-DRV" => "LED Driver",
            "NTC"     => "NTC Thermistor",
            "PTC"     => "PTC Thermistor",
            "LDO"     => "Linear Voltage Regulator",
            "MCU"     => "Microcontroller",
            "OTHER-IC"=> "Other IC",
            "TERM"    => "Terminal",
            "PAD"     => "PCB Pad",
            _         => shortName,
        };

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

                var results = ComponentRepository.Search(query!, cat!, rating!, proc!, mfr!);
                if (_lblCount != null) _lblCount.Text = $"{results.Count:N0} component{(results.Count!=1?"s":"")} found";

                var table = new System.Data.DataTable();
                foreach (var col in new[]{"FekaNo","ID","Category","Description","PartNumber","Manufacturer","Footprint","Rating","Process","Status"})
                    table.Columns.Add(col);
                foreach (var gf in ComponentRepository.GlobalFields)
                    if (!table.Columns.Contains(gf.Key)) table.Columns.Add(gf.Key);
                foreach (var c in results)
                {
                    var row = table.NewRow();
                    row["FekaNo"] = c.FekaNo; row["ID"] = c.Id; row["Category"] = c.Category;
                    row["Description"] = c.Description; row["PartNumber"] = c.PartNumber;
                    row["Manufacturer"] = c.Manufacturer; row["Footprint"] = c.Footprint;
                    row["Rating"] = c.Rating; row["Process"] = c.Process; row["Status"] = c.Status;
                    foreach (var gf in ComponentRepository.GlobalFields)
                        if (c.ExtraFields.TryGetValue(gf.Key, out var val)) row[gf.Key] = val;
                    table.Rows.Add(row);
                }

                if (_grid != null)
                {
                    _grid.DataSource = table;
                    void SetW(string col, int w) { if (_grid.Columns[col] != null) _grid.Columns[col]!.Width = w; }
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

            var btnNewCat = new ModernButton { Text = "+ New Cat", Style = ModernButton.BtnStyle.Secondary, Width = 82, Height = fh, Left = 32+lw+8+(fw-90)+8, Top = y, Parent = card };
            btnNewCat.Click += (s, e) => ShowNewCategoryDialog();
            y += 44;

            L("Global Fields", y);
            var btnNewField = new ModernButton { Text = "+ New Field", Style = ModernButton.BtnStyle.Secondary, Width = 110, Height = fh, Left = 32+lw+8, Top = y, Parent = card };
            btnNewField.Click += (s, e) => { ShowNewGlobalFieldDialog(); RebuildGlobalFields(); };
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

            _globalPanel = new Panel { Left=0, Top=extraTop, Width=card.Width, Height=0, BackColor=Color.Transparent, Parent=card };

            var sep      = new Panel { Left=32, Width=560, Height=1, BackColor=Theme.Border, Parent=card };
            var btnAdd   = new ModernButton { Text="Add Component", Style=ModernButton.BtnStyle.Success, Width=160, Height=36, Left=32, Parent=card };
            var btnReset = new ModernButton { Text="Reset", Style=ModernButton.BtnStyle.Secondary, Width=90, Height=36, Left=200, Parent=card };

            void Reposition()
            {
                _globalPanel.Top = extraTop + _extraPanel.Height;
                int baseY = extraTop + _extraPanel.Height + _globalPanel.Height + 12;
                sep.Top = baseY; btnAdd.Top = baseY+12; btnReset.Top = baseY+12;
                card.Height = baseY+70; _extraPanel.Width = card.Width; _globalPanel.Width = card.Width;
            }
            _reposition = Reposition;

            _addCat.SelectedIndexChanged += (s, e) => { RebuildExtraFields(); RebuildGlobalFields(); Reposition(); };
            btnAdd.Click   += BtnAddComponent_Click;
            btnReset.Click += (s, e) => ClearAddForm();
            RebuildGlobalFields();
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
            if (_addCat.SelectedItem == null) { _extraPanel.Height = 0; return; }
            var cat = _addCat.SelectedItem.ToString()!;
            if (!ComponentRepository.CategoryFields.TryGetValue(cat, out var fields) || fields.Count == 0) { _extraPanel.Height = 0; return; }

            int y=4; const int lw=140,fw=340,fh=32,rowH=44,lx=32,fx=32+lw+8;
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
            _extraPanel.Height = y+4;
            _reposition();
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
            foreach (var (key, ctrl) in _globalBoxes)
            {
                string val = ctrl switch { ComboBox cb=>cb.Text.Trim(), TextBox tb=>tb.Text.Trim(), _=>"" };
                if (!string.IsNullOrWhiteSpace(val))
                {
                    comp.ExtraFields[key] = val;
                    ComponentRepository.RecordGlobalFieldValue(key, val);
                }
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
            RebuildGlobalFields();
        }

        private void RebuildGlobalFields()
        {
            _globalPanel.Controls.Clear();
            _globalBoxes.Clear();

            var fields = ComponentRepository.GlobalFields;
            if (fields.Count == 0) { _globalPanel.Height = 0; return; }

            const int lw=140, fw=340, fh=32, rowH=44, lx=32, fx=32+lw+8;
            int y = 4;

            new Label
            {
                Text = "▸  Global Parameters",
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(99, 102, 241),
                Left = lx, Top = y, Width = 300, Height = 24,
                BackColor = Color.Transparent, Parent = _globalPanel
            };
            y += 28;
            new Panel { Left=lx, Top=y, Width=fw+lw+8, Height=1, BackColor=Color.FromArgb(99,102,241), Parent=_globalPanel };
            y += 8;

            foreach (var field in fields)
            {
                new Label
                {
                    Text = field.Label, Left=lx, Top=y+2, Width=lw, Height=fh,
                    Font=Theme.FontBold, ForeColor=Theme.TextSecondary,
                    TextAlign=ContentAlignment.MiddleLeft,
                    BackColor=Color.Transparent, Parent=_globalPanel
                };

                var values = ComponentRepository.GlobalFieldValues.TryGetValue(field.Key, out var v) ? v : new();

                if (field.Type == "combo")
                {
                    var cb = new ComboBox
                    {
                        Left=fx, Top=y, Width=fw, Height=fh,
                        DropDownStyle=ComboBoxStyle.DropDown,
                        Font=Theme.FontBody, BackColor=Theme.BgSurface2,
                        ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat,
                        AutoCompleteMode=AutoCompleteMode.SuggestAppend,
                        AutoCompleteSource=AutoCompleteSource.ListItems,
                        Parent=_globalPanel
                    };
                    foreach (var val in values) cb.Items.Add(val);
                    _globalBoxes.Add((field.Key, cb));
                }
                else
                {
                    var tb = new TextBox
                    {
                        Left=fx, Top=y, Width=fw, Height=fh,
                        BorderStyle=BorderStyle.FixedSingle,
                        BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary,
                        Font=Theme.FontBody, Parent=_globalPanel
                    };
                    // Autocomplete from remembered values
                    if (values.Count > 0)
                    {
                        tb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                        tb.AutoCompleteSource = AutoCompleteSource.CustomSource;
                        var src = new AutoCompleteStringCollection();
                        src.AddRange(values.ToArray());
                        tb.AutoCompleteCustomSource = src;
                    }
                    _globalBoxes.Add((field.Key, tb));
                }
                y += rowH;
            }

            _globalPanel.Height = y + 4;
            _reposition();
        }

        private void ShowNewGlobalFieldDialog()
        {
            var dlg = new Form
            {
                Text = "New Global Field", Width = 360, Height = 200,
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false,
                BackColor = Theme.BgSurface
            };

            new Label { Text="Field Name:", Left=20, Top=20, Width=100, Height=24, ForeColor=Theme.TextPrimary, Font=Theme.FontBold, Parent=dlg };
            var tbLabel = new TextBox { Left=130, Top=18, Width=180, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, BorderStyle=BorderStyle.FixedSingle, Font=Theme.FontBody, Parent=dlg };

            new Label { Text="Input Type:", Left=20, Top=58, Width=100, Height=24, ForeColor=Theme.TextPrimary, Font=Theme.FontBold, Parent=dlg };
            var cbType = new ComboBox { Left=130, Top=56, Width=180, DropDownStyle=ComboBoxStyle.DropDownList, Font=Theme.FontBody, BackColor=Theme.BgSurface2, ForeColor=Theme.TextPrimary, FlatStyle=FlatStyle.Flat, Parent=dlg };
            cbType.Items.AddRange(new object[]{ "Text (free input)", "Dropdown (remembered values)" });
            cbType.SelectedIndex = 0;

            var btnOk = new ModernButton { Text="Add", Style=ModernButton.BtnStyle.Success, Width=80, Height=32, Left=130, Top=100, Parent=dlg };
            var btnCancel = new ModernButton { Text="Cancel", Style=ModernButton.BtnStyle.Secondary, Width=80, Height=32, Left=220, Top=100, Parent=dlg };

            btnCancel.Click += (s, e) => dlg.DialogResult = DialogResult.Cancel;
            btnOk.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(tbLabel.Text)) return;
                string label = tbLabel.Text.Trim();
                string key   = label.Replace(" ", "_").ToLower();
                string type  = cbType.SelectedIndex == 1 ? "combo" : "text";
                ComponentRepository.AddGlobalField(key, label, type);
                dlg.DialogResult = DialogResult.OK;
            };

            dlg.ShowDialog(this);
        }

        // ══════════════════════════════════════════════════════════════════
        //  APPROVAL
        // ══════════════════════════════════════════════════════════════════
        private void BuildApproval()
        {
            _pageApproval = new Panel { Dock = DockStyle.Fill, BackColor = Theme.BgMain, Visible = false };
            _content.Controls.Add(_pageApproval);

            var infoBar = new Panel { Left=0, Top=0, Height=40, BackColor=Color.FromArgb(254,243,199), Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right };
            new Label { Text="Pending components appear here for review. Approved components are automatically moved to catalog.", ForeColor=Color.FromArgb(120,80,4), Font=Theme.FontBody, AutoSize=true, Left=12, Top=10, Parent=infoBar };
            _pageApproval.Controls.Add(infoBar);
            _pageApproval.Resize += (s, e) => infoBar.Width = _pageApproval.ClientSize.Width;

            var btnApproveAll = new ModernButton { Text="Approve All", Style=ModernButton.BtnStyle.Success, Width=120, Height=32, Left=0, Top=50, Parent=_pageApproval };
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

            _approvalGrid = BuildGrid();
            _approvalGrid.Left=0; _approvalGrid.Top=92;
            _approvalGrid.Anchor=AnchorStyles.Top|AnchorStyles.Left|AnchorStyles.Right|AnchorStyles.Bottom;
            _approvalGrid.Width=_pageApproval.ClientSize.Width;
            _approvalGrid.Height=_pageApproval.ClientSize.Height-92;
            _approvalGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _approvalGrid.MultiSelect = false;

            RefreshApproval();

            var btnColApprove = new DataGridViewButtonColumn { Name="BtnApprove", HeaderText="Approve", Text="Approve", UseColumnTextForButtonValue=true, Width=90, FlatStyle=FlatStyle.Flat };
            var btnColDelete  = new DataGridViewButtonColumn { Name="BtnDelete",  HeaderText="Delete",  Text="Delete",  UseColumnTextForButtonValue=true, Width=80, FlatStyle=FlatStyle.Flat };
            
            if (_approvalGrid.Columns["Status"] != null)
            {
                int statusIdx = _approvalGrid.Columns["Status"]!.Index;
                _approvalGrid.Columns.Insert(statusIdx + 1, btnColApprove);
                _approvalGrid.Columns.Insert(statusIdx + 2, btnColDelete);
            }

            _approvalGrid.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                if (_approvalGrid.Columns[e.ColumnIndex].Name == "BtnApprove")
                {
                    e.CellStyle.BackColor = Color.FromArgb(32, 193, 148);
                    e.CellStyle.ForeColor = Color.White;
                    e.CellStyle.Font = new Font("Segoe UI", 8f, FontStyle.Bold);
                }
                else if (_approvalGrid.Columns[e.ColumnIndex].Name == "BtnDelete")
                {
                    e.CellStyle.BackColor = Color.FromArgb(220, 80, 80);
                    e.CellStyle.ForeColor = Color.White;
                    e.CellStyle.Font = new Font("Segoe UI", 8f, FontStyle.Bold);
                }
            };

            _approvalGrid.CellClick += (s, e) =>
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                var row = _approvalGrid.Rows[e.RowIndex];
                if (!int.TryParse(row.Cells["ID"].Value?.ToString(), out int id)) return;
                string desc = row.Cells["Description"].Value?.ToString() ?? "this component";

                if (_approvalGrid.Columns[e.ColumnIndex].Name == "BtnApprove")
                {
                    if (MessageBox.Show($"Approve '{desc}'?", "Approve", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    { 
                        ComponentRepository.Approve(id); 
                        RefreshApproval(); // Burada liste yenilenince onaylanan kaybolacak
                        RefreshApprovalBadge(); 
                    }
                }
                else if (_approvalGrid.Columns[e.ColumnIndex].Name == "BtnDelete")
                {
                    if (MessageBox.Show($"Delete '{desc}'?", "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    { 
                        ComponentRepository.Delete(id); 
                        RefreshApproval(); 
                        RefreshApprovalBadge(); 
                    }
                }
            };

            _pageApproval.Controls.Add(_approvalGrid);
            _pageApproval.Resize += (s, e) =>
            {
                _approvalGrid.Width=_pageApproval.ClientSize.Width;
                _approvalGrid.Height=_pageApproval.ClientSize.Height-92;
            };
        }

        private void RefreshApproval()
        {
            // SADECE Pending olanları getiriyoruz, onaylananlar otomatik listeden düşer
            var pendingOnly = ComponentRepository.GetAll()
                .Where(c => c.IsCustom && c.Status == ComponentStatus.Pending)
                .ToList();

            var table = new System.Data.DataTable();
            foreach (var col in new[] { "FekaNo", "ID", "Category", "Description", "PartNumber", "Manufacturer", "Footprint", "Rating", "Process", "Status" })
                table.Columns.Add(col);
            foreach (var gf in ComponentRepository.GlobalFields)
                if (!table.Columns.Contains(gf.Key)) table.Columns.Add(gf.Key);

            foreach (var c in pendingOnly)
            {
                var row = table.NewRow();
                row["FekaNo"] = c.FekaNo; row["ID"] = c.Id; row["Category"] = c.Category;
                row["Description"] = c.Description; row["PartNumber"] = c.PartNumber;
                row["Manufacturer"] = c.Manufacturer; row["Footprint"] = c.Footprint;
                row["Rating"] = c.Rating; row["Process"] = c.Process; row["Status"] = c.Status;
                foreach (var gf in ComponentRepository.GlobalFields)
                    if (c.ExtraFields.TryGetValue(gf.Key, out var val)) row[gf.Key] = val;
                table.Rows.Add(row);
            }

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

            // ── E-BOM butonu artık unified dialog açıyor ──────────────────
            var btnExportEBom = new ModernButton { Text="Export E-BOM (xlsx,xls,ods,csv,tsv,json)", Style=ModernButton.BtnStyle.Primary, Width=200, Height=70, Left=24,  Top=180, Parent=card };

            btnExportEBom.Click += BtnExportEBom_Click;
            new Panel { Left=24, Top=292, Width=572, Height=1, BackColor=Theme.Border, Parent=card };
            new Label { Text="Import CSV", Font=Theme.FontH2, ForeColor=Theme.TextPrimary, Left=24, Top=306, AutoSize=true, Parent=card };
            var btnImportCsv = new ModernButton { Text="Import CSV File…", Style=ModernButton.BtnStyle.Secondary, Width=160, Height=34, Left=24, Top=338, Parent=card };
            btnImportCsv.Click += BtnImportCsv_Click;
        }

        private enum ExportFormat { Csv, Tsv, Json, Xml, Xls, Ods }

        // ── YENİ: E-BOM unified export dialog ────────────────────────────
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
                    case 1: // xlsx – şablonlu E-BOM
                        ExportToEBomExcel(dlg.FileName, components);
                        break;
                    case 2: // xls – CSV içeriği .xls uzantısıyla
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

        // ── Mevcut ayrı format butonları için (kendi SaveFileDialog'ları var) ──
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

        // ── Asıl yazma mantığı (path + components dışarıdan gelir) ────────
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
                MessageBox.Show("Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                using (var workbook = new XLWorkbook(templatePath))
                {
                    var ws = workbook.Worksheet("E-BOM");

                    foreach (var hyperlink in ws.Hyperlinks.ToList())
                        ws.Hyperlinks.Delete(hyperlink);

                    foreach (var cell in ws.Range(6, 2, 50, 13).Cells())
                    {
                        cell.Style.Font.FontColor = XLColor.Black;
                        cell.Style.Font.Underline = XLFontUnderlineValues.None;
                        cell.Value = "";
                    }

                    int templateRow = 6;
                    int currentRow  = 6;

                    foreach (var item in data)
                    {
                        if (currentRow > templateRow)
                        {
                            CopyRowFormat(ws, templateRow, currentRow, 2, 13);
                            ws.Cell(currentRow, 2).FormulaA1 = "=ROW()-5";
                        }

                        for (int col = 2; col <= 13; col++)
                        {
                            ws.Cell(currentRow, col).Style.Font.FontColor = XLColor.Black;
                            ws.Cell(currentRow, col).Style.Font.Underline = XLFontUnderlineValues.None;
                        }

                        ws.Cell(currentRow, 3).Value  = item.FekaNo       ?? "";
                        ws.Cell(currentRow, 4).Value  = item.Category     ?? "";
                        ws.Cell(currentRow, 5).Value  = item.Description  ?? "";
                        ws.Cell(currentRow, 6).Value  = item.Footprint    ?? "";
                        ws.Cell(currentRow, 7).Value  = item.PartNumber   ?? "";
                        ws.Cell(currentRow, 8).Value  = item.Manufacturer ?? "";
                        ws.Cell(currentRow, 9).Value  = item.Rating  == ComponentRating.AECQ ? "AECQ" : "STD";
                        ws.Cell(currentRow, 10).Value = item.Process == ComponentProcess.THT  ? "THT"  : "SMD";

                        currentRow++;
                    }

                    for (int r = 6; r < currentRow; r++)
                        ws.Row(r).Height = 18;

                    ws.Column(5).Width = 35;
                    ws.Column(7).Width = 22;
                    ws.Column(8).Width = 16;

                    var table = ws.Tables.FirstOrDefault();
                    if (table != null)
                        table.Resize(ws.Range(5, 2, currentRow - 1, 13));

                    workbook.SaveAs(path);
                }
            }
            catch (Exception ex) { MessageBox.Show("Hata: " + ex.Message, "Sistem Hatası", MessageBoxButtons.OK, MessageBoxIcon.Error); }
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
    // Düzeltme: Alt kısımdaki input alanlarına çarpmaması için yükseklik payı artırıldı
    leftCard.Resize += (s, e) => _catList.Height = leftCard.Height - 100;

    _catList.SelectedIndexChanged += (s, e) =>
    {
        CloseFieldDetail();
        RefreshFieldList();
    };

    var txtNewCat = new TextBox { Left = 8, Width = 152, Height = 28, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, PlaceholderText="New category...", Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = leftCard };
    var btnAddCat = new ModernButton { Text = "+", Style = ModernButton.BtnStyle.Primary, Width = 44, Height = 28, Left = 164, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = leftCard };
    
    // Düzeltme: Dikey hizalama txtNewCat ile btnAddCat için eşitlendi
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

    // Düzeltme: Başlık genişliği butonla çakışmaması için daraltıldı
    _lblCatTitle = new Label { Text = "Select a category", Font = new Font("Segoe UI", 9.5f, FontStyle.Bold), ForeColor = Theme.TextSecondary, Left = 16, Top = 16, Width = MID_W - 125, AutoSize = false, Parent = midCard };

    // Düzeltme: "Delete Cat" butonu konumu başlığın üzerine binmeyecek şekilde sağa yaslandı
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
    // Düzeltme: Liste yüksekliği, alttaki "Remove Field" butonuyla çakışmayacak şekilde ayarlandı
    midCard.Resize += (s, e) => _fieldList.Height = midCard.Height - 210;

    _fieldList.SelectedIndexChanged += (s, e) =>
    {
        var catSel = _catList.SelectedItem?.ToString();
        var fieldSel = _fieldList.SelectedItem?.ToString();
        if (catSel == null || fieldSel == null) { CloseFieldDetail(); return; }
        // Detail panel disabled
        // ShowFieldDetail(catSel, fieldSel);
    };

    // Düzeltme: "Remove Field" butonu liste bittiğinde hemen altında düzgün duracak şekilde konumu güncellendi
    var btnDelField = new ModernButton { Text = "Remove Field", Style = ModernButton.BtnStyle.Secondary, Width = 130, Height = 28, Left = 8, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = midCard };
    midCard.Resize += (s, e) => btnDelField.Top = midCard.Height - 135; 
    
    btnDelField.Click += (s, e) =>
    {
        var catSel = _catList.SelectedItem?.ToString();
        var fieldSel = _fieldList.SelectedItem?.ToString();
        if (catSel == null || fieldSel == null) return;
        ComponentRepository.RemoveFieldFromCategory(catSel, fieldSel.Trim());
        CloseFieldDetail();
        RefreshFieldList();
    };

    var sepLine = new Panel { Left = 8, Height = 1, BackColor = Theme.Border, Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right, Parent = midCard };
    
    var lblAddField = new Label { Text = "Add Field:", Font = Theme.FontBold, ForeColor = Theme.TextSecondary, Left = 8, Width = 70, Height = 28, AutoSize = false, TextAlign = ContentAlignment.MiddleLeft, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = midCard };
    var txtFieldKey = new TextBox { Width = MID_W - 90, Height = 28, BorderStyle = BorderStyle.FixedSingle, BackColor = Theme.BgSurface2, ForeColor = Theme.TextPrimary, Font = Theme.FontBody, PlaceholderText = "Field key (e.g. Voltage(V))", Left = 82, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = midCard };
    var btnAddField = new ModernButton { Text = "Add", Style = ModernButton.BtnStyle.Success, Width = MID_W - 16, Height = 28, Left = 8, Anchor = AnchorStyles.Bottom | AnchorStyles.Left, Parent = midCard };

    // Düzeltme: Alt panel elemanlarının birbirine girmemesi için Top değerleri netleştirildi
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

    // --- RIGHT CARD (Detaylar) ---
    const int RIGHT_LEFT = MID_LEFT + MID_W + 8;

    _fieldDetailPanel = new Panel
    {
        Left = RIGHT_LEFT, Top = 0,
        Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
        BackColor = Theme.BgSurface,
        Visible = false,
        AutoScroll = true,
    };
    _fieldDetailPanel.Paint += (s, e) =>
    {
        using var pen = new Pen(Theme.Border, 1);
        e.Graphics.DrawRectangle(pen, 0, 0, _fieldDetailPanel.Width - 1, _fieldDetailPanel.Height - 1);
        using var brush = new SolidBrush(Theme.Accent);
        e.Graphics.FillRectangle(brush, 0, 0, _fieldDetailPanel.Width, 4);
    };
    _pageCatManager.Controls.Add(_fieldDetailPanel);

    _pageCatManager.Resize += (s, e) =>
    {
        _fieldDetailPanel.Width = Math.Max(200, _pageCatManager.ClientSize.Width - RIGHT_LEFT - 10); // Padding eklendi
        _fieldDetailPanel.Height = _pageCatManager.ClientSize.Height;
    };

    var btnCloseDetail = new Label
    {
        Text = "✕", ForeColor = Theme.TextMuted,
        Font = new Font("Segoe UI", 11f), AutoSize = true,
        Cursor = Cursors.Hand, Top = 12,
        Anchor = AnchorStyles.Top | AnchorStyles.Right,
        Parent = _fieldDetailPanel,
    };
    
    // Düzeltme: Kapatma butonu sağa yaslama mantığı güncellendi
    _fieldDetailPanel.Resize += (s, e) => btnCloseDetail.Left = _fieldDetailPanel.Width - 30;
    
    btnCloseDetail.MouseEnter += (s, e) => ((Label)s!).ForeColor = Theme.Danger;
    btnCloseDetail.MouseLeave += (s, e) => ((Label)s!).ForeColor = Theme.TextMuted;
    btnCloseDetail.Click += (s, e) => { CloseFieldDetail(); _fieldList.ClearSelected(); };

    _lblFieldDetailTitle = new Label
    {
        Text = "", Left = 16, Top = 14,
        Width = 300, Height = 22,
        Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
        ForeColor = Theme.TextPrimary,
        AutoSize = false, Parent = _fieldDetailPanel,
    };

    _fieldDetailContent = new Panel
    {
        Left = 0, Top = 44,
        Width = _fieldDetailPanel.Width,
        Height = 0,
        BackColor = Color.Transparent,
        Parent = _fieldDetailPanel,
        AutoScroll = false,
    };
    _fieldDetailPanel.Resize += (s, e) => _fieldDetailContent.Width = _fieldDetailPanel.Width;
}
        // ══════════════════════════════════════════════════════════════════
        //  FIELD DETAIL
        // ══════════════════════════════════════════════════════════════════
        private void ShowFieldDetail(string categoryKey, string fieldKey)
        {
            _fieldDetailPanel.Visible = true;
            _lblFieldDetailTitle.Text = fieldKey;
            _fieldDetailContent.Controls.Clear();

            int y = 0;
            var allComps = ComponentRepository.GetAll();

            y += 4;
            y = AddDetailSection(_fieldDetailContent, "USAGE STATISTICS", y);

            var catComps   = allComps.Where(c => string.Equals(c.Category?.Trim(), categoryKey.Trim(), StringComparison.OrdinalIgnoreCase)).ToList();
            var allWithField = allComps.Where(c => c.ExtraFields.ContainsKey(fieldKey) && !string.IsNullOrWhiteSpace(c.ExtraFields[fieldKey])).ToList();

            int catTotal     = catComps.Count;
            int usingInCat   = catComps.Count(c => c.ExtraFields.ContainsKey(fieldKey) && !string.IsNullOrWhiteSpace(c.ExtraFields[fieldKey]));
            int usingGlobal  = allWithField.Count;
            int catsUsingIt  = ComponentRepository.CategoryFields.Count(kvp => kvp.Value.Any(f => f.Key == fieldKey));

            string catStatText = catTotal > 0
                ? $"{usingInCat} / {catTotal} components in [{categoryKey}] use this field"
                : $"No components in [{categoryKey}] yet";
            AddDetailInfoLine(_fieldDetailContent, catStatText, Theme.TextSecondary, ref y);

            if (catTotal > 0)
            {
                int barW  = Math.Max(60, (_fieldDetailPanel.Width - 48));
                int fillW = catTotal > 0 ? (int)((double)usingInCat / catTotal * barW) : 0;
                var track = new Panel { Left = 16, Top = y, Width = barW, Height = 8, BackColor = Theme.BgSurface3, Parent = _fieldDetailContent };
                if (fillW > 0) track.Controls.Add(new Panel { Left = 0, Top = 0, Width = fillW, Height = 8, BackColor = Theme.AccentGreen });
                _fieldDetailPanel.Resize += (s, e) =>
                {
                    int bw = Math.Max(60, _fieldDetailPanel.Width - 48);
                    track.Width = bw;
                    int fw2 = catTotal > 0 ? (int)((double)usingInCat / catTotal * bw) : 0;
                    if (track.Controls.Count > 0) track.Controls[0].Width = fw2;
                };
                y += 18;
                int pct = catTotal > 0 ? (int)((double)usingInCat / catTotal * 100) : 0;
                AddDetailInfoLine(_fieldDetailContent, $"{pct}% coverage in [{categoryKey}]", Theme.TextMuted, ref y);
            }

            y += 6;

            AddDetailInfoLine(_fieldDetailContent, $"Total usage across all categories: {usingGlobal} component(s)", Theme.TextSecondary, ref y);
            AddDetailInfoLine(_fieldDetailContent, $"Used in {catsUsingIt} category(s) in total", Theme.TextMuted, ref y);

            int approvedCount = allWithField.Count(c => c.Status == ComponentStatus.Approved);
            int pendingCount  = allWithField.Count(c => c.Status == ComponentStatus.Pending);
            y += 4;
            AddDetailInfoLine(_fieldDetailContent, $"Approved: {approvedCount}   |   Pending: {pendingCount}", Theme.TextSecondary, ref y);

            y += 4;

            y = AddDetailSection(_fieldDetailContent, "PREDEFINED OPTIONS", y);

            string[]? typeOpts = null;
            if (fieldKey == "Type")
            {
                typeOpts = categoryKey switch
                {
                    "TRAN"    => new[] { "NPN", "PNP", "NPN&PNP" },
                    "MOS"     => new[] { "N-Ch", "P-Ch" },
                    "DC-DC"   => new[] { "Buck", "Buck-Boost", "Flyback", "SEPIC", "Cuk" },
                    "LED-DRV" => new[] { "Linear", "Switching" },
                    _         => new[] { "N/A" },
                };
            }

            var effectiveOpts = typeOpts ?? (FieldComboOptions.TryGetValue(fieldKey, out var fo) ? fo : null);

            if (effectiveOpts != null && effectiveOpts.Length > 0)
            {
                var optFlow = new FlowLayoutPanel
                {
                    Left = 16, Top = y,
                    Width = _fieldDetailContent.Width - 32,
                    AutoSize = true, FlowDirection = FlowDirection.LeftToRight,
                    WrapContents = true, BackColor = Color.Transparent,
                    Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                    Parent = _fieldDetailContent,
                };
                _fieldDetailContent.Resize += (s, e2) => optFlow.Width = _fieldDetailContent.Width - 32;

                foreach (var opt in effectiveOpts)
                {
                    new Label
                    {
                        Text = opt,
                        Font = Theme.FontSmall,
                        ForeColor = Theme.TextSecondary,
                        BackColor = Theme.BgSurface3,
                        AutoSize = false,
                        Width = Math.Max(40, TextRenderer.MeasureText(opt, Theme.FontSmall).Width + 16),
                        Height = 20,
                        TextAlign = ContentAlignment.MiddleCenter,
                        Margin = new Padding(0, 0, 4, 4),
                        Parent = optFlow,
                    };
                }
                optFlow.PerformLayout();
                y += optFlow.PreferredSize.Height + 8;
            }
            else
            {
                var freeLabel = new Label
                {
                    Text = "Free-text (no predefined options)",
                    Font = Theme.FontSmall, ForeColor = Theme.TextMuted, AutoSize = false,
                    Left = 16, Top = y, Width = _fieldDetailContent.Width - 32, Height = 20,
                    Parent = _fieldDetailContent,
                };
                _fieldDetailContent.Resize += (s, e2) => freeLabel.Width = _fieldDetailContent.Width - 32;
                y += 28;
            }

            y += 4;

            y = AddDetailSection(_fieldDetailContent, "VALUE DISTRIBUTION (this category)", y);

            var valueCounts = catComps
                .Where(c => c.ExtraFields.ContainsKey(fieldKey) && !string.IsNullOrWhiteSpace(c.ExtraFields[fieldKey]))
                .GroupBy(c => c.ExtraFields[fieldKey])
                .OrderByDescending(g => g.Count())
                .Take(10)
                .ToList();

            if (valueCounts.Count == 0)
            {
                AddDetailInfoLine(_fieldDetailContent, "No data yet", Theme.TextMuted, ref y);
            }
            else
            {
                int maxCount = valueCounts[0].Count();
                Color[] barColors = { Theme.Accent, Theme.AccentGreen, Color.FromArgb(168,85,247), Theme.Warning };
                int colorIdx = 0;

                foreach (var grp in valueCounts)
                {
                    var color  = barColors[colorIdx % barColors.Length];
                    int availW = Math.Max(80, _fieldDetailPanel.Width - 48 - 80 - 8 - 30);

                    var rowPanel = new Panel { Left = 16, Top = y, Width = _fieldDetailContent.Width - 32, Height = 22, BackColor = Color.Transparent, Parent = _fieldDetailContent };
                    _fieldDetailContent.Resize += (s, e2) => rowPanel.Width = _fieldDetailContent.Width - 32;

                    rowPanel.Controls.Add(new Label { Text = grp.Key, Left = 0, Top = 3, Width = 80, Height = 16, ForeColor = Theme.TextSecondary, Font = Theme.FontSmall, AutoSize = false });

                    int barMaxW = Math.Max(20, availW);
                    int fillPx  = maxCount > 0 ? Math.Max(4, (int)((double)grp.Count() / maxCount * barMaxW)) : 0;

                    var barTrack = new Panel { Left = 84, Top = 7, Width = barMaxW, Height = 8, BackColor = Theme.BgSurface3 };
                    barTrack.Controls.Add(new Panel { Left = 0, Top = 0, Width = fillPx, Height = 8, BackColor = color });
                    rowPanel.Controls.Add(barTrack);

                    rowPanel.Resize += (s, e2) =>
                    {
                        int aw2 = Math.Max(20, rowPanel.Width - 80 - 8 - 30);
                        barTrack.Width = aw2;
                        int fp2 = maxCount > 0 ? Math.Max(4, (int)((double)grp.Count() / maxCount * aw2)) : 0;
                        if (barTrack.Controls.Count > 0) barTrack.Controls[0].Width = fp2;
                    };

                    var lblCnt = new Label { Text = grp.Count().ToString(), Left = barMaxW + 88, Top = 3, Width = 28, Height = 16, ForeColor = Theme.TextMuted, Font = Theme.FontSmall, AutoSize = false, TextAlign = ContentAlignment.MiddleRight };
                    rowPanel.Controls.Add(lblCnt);
                    rowPanel.Resize += (s, e2) => lblCnt.Left = rowPanel.Width - 30;

                    y += 24;
                    colorIdx++;
                }
            }

            y += 8;
            y = AddDetailSection(_fieldDetailContent, "COMPONENTS USING THIS FIELD", y);

            var compsUsingField = catComps
                .Where(c => c.ExtraFields.ContainsKey(fieldKey) && !string.IsNullOrWhiteSpace(c.ExtraFields[fieldKey]))
                .OrderBy(c => c.Description)
                .ToList();

            if (compsUsingField.Count == 0)
            {
                AddDetailInfoLine(_fieldDetailContent, "No components yet", Theme.TextMuted, ref y);
            }
            else
            {
                foreach (var comp in compsUsingField.Take(15))
                {
                    var rowPanel = new Panel
                    {
                        Left = 16, Top = y,
                        Width = _fieldDetailContent.Width - 32,
                        Height = 36,
                        BackColor = Theme.BgSurface2,
                        Margin = new Padding(0, 0, 0, 2),
                        Parent = _fieldDetailContent,
                    };
                    _fieldDetailContent.Resize += (s, e2) => rowPanel.Width = _fieldDetailContent.Width - 32;

                    new Label
                    {
                        Text = comp.Description,
                        Font = new Font("Segoe UI", 8f, FontStyle.Bold),
                        ForeColor = Theme.TextPrimary,
                        Left = 8, Top = 2, Width = rowPanel.Width - 100, Height = 16,
                        AutoSize = false, BackColor = Color.Transparent,
                        Parent = rowPanel,
                    };
                    new Label
                    {
                        Text = comp.ExtraFields[fieldKey],
                        Font = Theme.FontSmall,
                        ForeColor = Theme.Accent,
                        Left = 8, Top = 18, Width = rowPanel.Width - 100, Height = 14,
                        AutoSize = false, BackColor = Color.Transparent,
                        Parent = rowPanel,
                    };

                    var statusBadge = new Label
                    {
                        Text = comp.Status.ToString(),
                        Font = new Font("Segoe UI", 7f, FontStyle.Bold),
                        ForeColor = comp.Status == ComponentStatus.Approved ? Theme.AccentGreen : Theme.Warning,
                        Left = rowPanel.Width - 90, Top = 10, Width = 82, Height = 16,
                        AutoSize = false, BackColor = Color.Transparent,
                        TextAlign = ContentAlignment.MiddleRight,
                        Parent = rowPanel,
                    };
                    rowPanel.Resize += (s, e2) => statusBadge.Left = rowPanel.Width - 90;

                    y += 40;
                }

                if (compsUsingField.Count > 15)
                    AddDetailInfoLine(_fieldDetailContent, $"... and {compsUsingField.Count - 15} more", Theme.TextMuted, ref y);
            }

            _fieldDetailContent.Height = y + 16;
        }

        private void CloseFieldDetail()
        {
            _fieldDetailPanel.Visible = false;
            _fieldDetailContent?.Controls.Clear();
        }

        private static int AddDetailSection(Panel parent, string title, int y)
        {
            var sep = new Panel { Left = 0, Top = y, Width = parent.Width, Height = 1, BackColor = Theme.Border, Parent = parent };
            parent.Resize += (s, e) => sep.Width = parent.Width;

            new Label
            {
                Text = title,
                Font = new Font("Segoe UI", 7.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(100, 116, 150),
                AutoSize = false, Left = 16, Top = y + 8,
                Width = 300, Height = 16,
                BackColor = Color.Transparent,
                Parent = parent,
            };
            return y + 32;
        }

        private static void AddDetailInfoLine(Panel parent, string text, Color color, ref int y)
        {
            var lbl = new Label
            {
                Text = text, Font = Theme.FontSmall, ForeColor = color, AutoSize = false,
                Left = 16, Top = y, Width = parent.Width - 32, Height = 18,
                BackColor = Color.Transparent, Parent = parent,
            };
            parent.Resize += (s, e) => lbl.Width = parent.Width - 32;
            y += 22;
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
            CloseFieldDetail();

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