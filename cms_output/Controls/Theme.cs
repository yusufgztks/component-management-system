using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace CMS.Controls
{
    public static class Theme
    {
        public static bool IsDarkMode { get; private set; } = true;

        // ── Background layers ──────────────────────────────────────────────
        public static Color BgMain;
        public static Color BgSurface;
        public static Color BgSurface2;
        public static Color BgSurface3;

        // ── Sidebar ───────────────────────────────────────────────────────
        public static Color SidebarBg;
        public static Color SidebarHover;
        public static Color SidebarActive;
        public static Color SidebarText;

        // ── Accent / action colours ───────────────────────────────────────
        public static Color Accent;
        public static Color AccentHover;
        public static Color AccentLight;
        public static Color AccentGreen;
        public static Color Danger;
        public static Color DangerHover;
        public static Color Warning;

        // ── Text ──────────────────────────────────────────────────────────
        public static Color TextPrimary;
        public static Color TextSecondary;
        public static Color TextMuted;
        public static Color TextDisabled;

        // ── Borders ───────────────────────────────────────────────────────
        public static Color Border;
        public static Color BorderFocus;

        // ── Topbar ────────────────────────────────────────────────────────
        public static Color TopbarBg;
        public static Color TopbarBorder;

        public static Font SafeFont(float size, FontStyle style = FontStyle.Regular)
        {
            string[] preferred = { "Inter", "Roboto", "Segoe UI Variable Display", "Segoe UI" };
            foreach (var f in preferred)
            {
                try
                {
                    using var fontFamily = new FontFamily(f);
                    return new Font(fontFamily, size, style);
                }
                catch { }
            }
            return new Font("Microsoft Sans Serif", size, style);
        }

        // ── Fonts ─────────────────────────────────────────────────────────
        public static Font FontBody    = SafeFont(9.5f, FontStyle.Regular);
        public static Font FontBold    = SafeFont(9.5f, FontStyle.Bold);
        public static Font FontSmall   = SafeFont(8.5f, FontStyle.Regular);
        public static Font FontNav     = SafeFont(10f,  FontStyle.Regular);
        public static Font FontNavBold = SafeFont(10f,  FontStyle.Bold);
        public static Font FontTitle   = SafeFont(14f,  FontStyle.Bold);
        public static Font FontH2      = SafeFont(12f,  FontStyle.Bold);
        public static Font FontHuge    = SafeFont(28f,  FontStyle.Regular);

        private static string ThemeFilePath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "theme.txt");

        static Theme()
        {
            LoadSettings();
            Apply();
        }

        public static void Toggle()
        {
            IsDarkMode = !IsDarkMode;
            File.WriteAllText(ThemeFilePath, IsDarkMode ? "dark" : "light");
            Apply();
        }

        private static void LoadSettings()
        {
            if (File.Exists(ThemeFilePath))
            {
                var content = File.ReadAllText(ThemeFilePath).Trim().ToLower();
                IsDarkMode = content != "light";
            }
            else
            {
                IsDarkMode = true; // Default
            }
        }

        [System.Runtime.InteropServices.DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        public static void UpdateForm(Form f)
        {
            try
            {
                var h = f.Handle; // Force handle creation
                int isDark = IsDarkMode ? 1 : 0;
                DwmSetWindowAttribute(h, 20, ref isDark, sizeof(int));
                DwmSetWindowAttribute(h, 19, ref isDark, sizeof(int)); // for older Windows 10 builds
            }
            catch { }
        }

        public static void MakeRounded(Control c, int radius = 6)
        {
            var p = new Panel { Bounds = c.Bounds, Parent = c.Parent, BackColor = c.BackColor, Cursor = c.Cursor, Dock = c.Dock, Margin = c.Margin, Anchor = c.Anchor };
            c.Dock = DockStyle.None;
            c.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            c.Parent = p;
            if (c is TextBox tb) tb.BorderStyle = BorderStyle.None;
            
            void UpdateLayout()
            {
                if (c is ComboBox cb) 
                {
                    cb.FlatStyle = FlatStyle.Flat;
                    c.Location = new Point(8, Math.Max(0, (p.Height - c.Height) / 2));
                    c.Width = p.Width - 6;
                }
                else
                {
                    c.Location = new Point(8, Math.Max(0, (p.Height - c.Height) / 2) + 2);
                    c.Width = p.Width - 16;
                }
            }
            
            p.Resize += (s, e) => UpdateLayout();
            UpdateLayout();
            
            // To ensure clicking panel focuses the input
            p.Click += (s, e) => c.Focus();
            
            p.Paint += (s, e) => {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                var r = new Rectangle(0, 0, p.Width - 1, p.Height - 1);
                var path = new System.Drawing.Drawing2D.GraphicsPath();
                int d = radius * 2;
                if (d > r.Width) d = r.Width;
                if (d > r.Height) d = r.Height;
                if (d > 0)
                {
                    path.AddArc(r.X, r.Y, d, d, 180, 90);
                    path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
                    path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
                    path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
                    path.CloseFigure();
                    p.Region = new Region(path);
                    using var pen = new Pen(Border, 1);
                    e.Graphics.DrawPath(pen, path);
                }
            };
        }

        private static void Apply()
        {
            try { Application.SetColorMode(IsDarkMode ? SystemColorMode.Dark : SystemColorMode.Classic); } catch { }

            try
            {
                if (Application.OpenForms != null)
                {
                    foreach (Form f in Application.OpenForms)
                    {
                        UpdateForm(f);
                        f.Invalidate(true);
                    }
                }
            }
            catch { }

            if (IsDarkMode)
            {
                BgMain        = Color.FromArgb(18, 16, 36);      // Deep purple background
                BgSurface     = Color.FromArgb(30, 26, 51);      // Glass cards
                BgSurface2    = Color.FromArgb(38, 33, 64);      // Input fields
                BgSurface3    = Color.FromArgb(46, 40, 77);      // Hover surfaces
                
                SidebarBg     = Color.FromArgb(18, 16, 36);
                SidebarHover  = Color.FromArgb(30, 26, 51);
                SidebarActive = Color.FromArgb(0, 160, 154);     // FEKA Turquoise
                SidebarText   = Color.FromArgb(160, 150, 180);

                Accent        = Color.FromArgb(0, 160, 154);     // FEKA Turquoise
                AccentHover   = Color.FromArgb(0, 180, 174);
                AccentLight   = Color.FromArgb(0, 60, 60);
                AccentGreen   = Color.FromArgb(32, 193, 148);
                Danger        = Color.FromArgb(239, 68, 68);
                DangerHover   = Color.FromArgb(220, 38, 38);
                Warning       = Color.FromArgb(245, 158, 11);

                TextPrimary   = Color.FromArgb(240, 240, 245);
                TextSecondary = Color.FromArgb(180, 175, 200);
                TextMuted     = Color.FromArgb(120, 115, 140);
                TextDisabled  = Color.FromArgb(80, 75, 100);

                Border        = Color.FromArgb(46, 40, 77);
                BorderFocus   = Color.FromArgb(0, 160, 154);

                TopbarBg      = Color.FromArgb(18, 16, 36);
                TopbarBorder  = Color.FromArgb(46, 40, 77);
            }
            else
            {
                BgMain        = Color.FromArgb(243, 244, 246);   // Light gray
                BgSurface     = Color.FromArgb(255, 255, 255);   // White cards
                BgSurface2    = Color.FromArgb(249, 250, 251);
                BgSurface3    = Color.FromArgb(229, 231, 235);
                
                SidebarBg     = Color.FromArgb(255, 255, 255);
                SidebarHover  = Color.FromArgb(243, 244, 246);
                SidebarActive = Color.FromArgb(0, 160, 154);     // FEKA Turquoise
                SidebarText   = Color.FromArgb(75, 85, 99);

                Accent        = Color.FromArgb(0, 160, 154);     // FEKA Turquoise
                AccentHover   = Color.FromArgb(0, 140, 134);
                AccentLight   = Color.FromArgb(224, 244, 243);
                AccentGreen   = Color.FromArgb(16, 185, 129);
                Danger        = Color.FromArgb(239, 68, 68);
                DangerHover   = Color.FromArgb(220, 38, 38);
                Warning       = Color.FromArgb(245, 158, 11);

                TextPrimary   = Color.FromArgb(17, 24, 39);
                TextSecondary = Color.FromArgb(75, 85, 99);
                TextMuted     = Color.FromArgb(156, 163, 175);
                TextDisabled  = Color.FromArgb(209, 213, 219);

                Border        = Color.FromArgb(229, 231, 235);
                BorderFocus   = Color.FromArgb(0, 160, 154);

                TopbarBg      = Color.FromArgb(255, 255, 255);
                TopbarBorder  = Color.FromArgb(229, 231, 235);
            }
        }
    }
}
