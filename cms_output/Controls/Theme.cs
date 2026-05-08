using System.Drawing;

namespace CMS.Controls
{
    public static class Theme
    {
        // ── Background layers ──────────────────────────────────────────────
        public static Color BgMain      = Color.FromArgb(245, 247, 250);   // page background
        public static Color BgSurface   = Color.White;                      // cards / panels
        public static Color BgSurface2  = Color.FromArgb(248, 250, 252);   // input fields
        public static Color BgSurface3  = Color.FromArgb(237, 242, 247);   // hover surfaces

        // ── Sidebar ───────────────────────────────────────────────────────
        public static Color SidebarBg     = Color.FromArgb(24, 32, 48);
        public static Color SidebarHover  = Color.FromArgb(36, 47, 66);
        public static Color SidebarActive = Color.FromArgb(79, 124, 255);
        public static Color SidebarText   = Color.FromArgb(160, 174, 200);

        // ── Accent / action colours ───────────────────────────────────────
        public static Color Accent       = Color.FromArgb(79, 124, 255);
        public static Color AccentHover  = Color.FromArgb(61, 106, 238);
        public static Color AccentLight  = Color.FromArgb(219, 229, 255);
        public static Color AccentGreen  = Color.FromArgb(32, 193, 148);
        public static Color Danger       = Color.FromArgb(239, 68,  68);
        public static Color DangerHover  = Color.FromArgb(220, 38,  38);
        public static Color Warning      = Color.FromArgb(245, 158,  11);

        // ── Text ──────────────────────────────────────────────────────────
        public static Color TextPrimary   = Color.FromArgb(17,  24, 39);
        public static Color TextSecondary = Color.FromArgb(75,  85, 99);
        public static Color TextMuted     = Color.FromArgb(107, 114, 128);
        public static Color TextDisabled  = Color.FromArgb(156, 163, 175);

        // ── Borders ───────────────────────────────────────────────────────
        public static Color Border       = Color.FromArgb(226, 232, 240);
        public static Color BorderFocus  = Color.FromArgb(79,  124, 255);

        // ── Topbar ────────────────────────────────────────────────────────
        public static Color TopbarBg     = Color.White;
        public static Color TopbarBorder = Color.FromArgb(226, 232, 240);

        // ── Fonts ─────────────────────────────────────────────────────────
        public static Font FontBody    = new Font("Segoe UI", 9f,  FontStyle.Regular);
        public static Font FontBold    = new Font("Segoe UI", 9f,  FontStyle.Bold);
        public static Font FontSmall   = new Font("Segoe UI", 8f,  FontStyle.Regular);
        public static Font FontNav     = new Font("Segoe UI", 9.5f, FontStyle.Regular);
        public static Font FontNavBold = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        public static Font FontTitle   = new Font("Segoe UI", 13f, FontStyle.Bold);
        public static Font FontH2      = new Font("Segoe UI", 11f, FontStyle.Bold);
    }
}
