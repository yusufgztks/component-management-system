using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace CMS.Controls
{
    // ===== MODERN BUTTON =====
    public class ModernButton : Button
    {
        public enum BtnStyle { Primary, Secondary, Success, Danger }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public BtnStyle Style { get; set; } = BtnStyle.Primary;

        private bool _hovered;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color CustomBgColor { get; set; } = Color.Empty;

        public ModernButton()
        {
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Height = 34;
            Font = Theme.FontBold;
            Cursor = Cursors.Hand;
        }

        protected override bool ShowFocusCues => false;

        protected override void OnMouseEnter(EventArgs e) { _hovered = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hovered = false; Invalidate(); base.OnMouseLeave(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

            Color baseBg = CustomBgColor;
            if (baseBg == Color.Empty)
            {
                baseBg = Theme.BgMain;
                Control? p = this.Parent;
                while (p != null) { if (p.BackColor != Color.Transparent) { baseBg = p.BackColor; break; } p = p.Parent; }
            }
            g.Clear(baseBg);

            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            var path = RoundedRect(r, 16);

            Color bg1, bg2; Color fg;
            switch (Style)
            {
                case BtnStyle.Success:
                    bg1 = _hovered ? Theme.AccentHover : Theme.Accent;
                    bg2 = _hovered ? Color.FromArgb(0, 140, 134) : Color.FromArgb(0, 125, 120);
                    fg = Color.White; break;
                case BtnStyle.Danger:
                    bg1 = _hovered ? Theme.DangerHover : Theme.Danger;
                    bg2 = _hovered ? Color.FromArgb(200, 20,  20) : Color.FromArgb(210, 40,  40);
                    fg = Color.White; break;
                case BtnStyle.Secondary:
                    bg1 = _hovered ? Theme.BgSurface3 : Theme.BgSurface2;
                    bg2 = _hovered ? Theme.BgSurface2 : Theme.BgSurface;
                    fg = Theme.TextPrimary; break;
                default:
                    bg1 = _hovered ? Theme.AccentHover : Theme.Accent;
                    bg2 = _hovered ? Color.FromArgb(0, 140, 134) : Color.FromArgb(0, 125, 120);
                    fg = Color.White; break;
            }

            if (r.Width > 0 && r.Height > 0)
            {
                using var brush = new LinearGradientBrush(r, bg1, bg2, 135f);
                g.FillPath(brush, path);
            }

            if (Style == BtnStyle.Secondary)
            {
                using var pen = new Pen(Color.FromArgb(40, 255, 255, 255), 1);
                g.DrawPath(pen, path);
            }

            using var tb = new SolidBrush(fg);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(Text, Font, tb, new RectangleF(0, 0, Width, Height), sf);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 0 || r.Width <= 0 || r.Height <= 0) { path.AddRectangle(r); return path; }
            int d = radius * 2;
            if (d > r.Width) d = r.Width;
            if (d > r.Height) d = r.Height;
            if (d <= 0) { path.AddRectangle(r); return path; }
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    // ===== STAT CARD =====
    public class StatCard : Panel
    {
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string StatLabel { get; set; } = "";

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string StatValue { get; set; } = "0";

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string StatSub { get; set; } = "";

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color AccentColor { get; set; } = Color.FromArgb(79, 124, 255);

        public StatCard()
        {
            Height = 100;
            BackColor = Color.Transparent;
            DoubleBuffered = true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using var path = RoundedRect(r, 24);
            
            // Frosted Glass Fill
            using var bgBrush = new SolidBrush(Color.FromArgb(15, Theme.TextPrimary));
            g.FillPath(bgBrush, path);

            // Glass Highlight Border
            using var borderPen = new Pen(Color.FromArgb(25, Theme.TextPrimary), 1f);
            g.DrawPath(borderPen, path);
            
            // Inner glow to simulate thickness
            var innerR = new Rectangle(1, 1, Width - 3, Height - 3);
            using var innerPath = RoundedRect(innerR, 23);
            using var innerPen = new Pen(Color.FromArgb(10, Theme.TextPrimary), 1f);
            g.DrawPath(innerPen, innerPath);

            int px = 24;
            using var labelBrush = new SolidBrush(Color.FromArgb(160, Theme.TextPrimary));
            using var titleFont = Theme.SafeFont(8.5f, FontStyle.Regular);
            g.DrawString(StatLabel.ToUpperInvariant(), titleFont, labelBrush, px, 20);

            using var valueBrush = new SolidBrush(AccentColor);
            using var valueFont = Theme.SafeFont(28f, FontStyle.Bold);
            g.DrawString(StatValue, valueFont, valueBrush, px - 4, 34);

            using var subBrush = new SolidBrush(Theme.TextSecondary);
            using var subFont = Theme.SafeFont(9f, FontStyle.Regular);
            g.DrawString(StatSub, subFont, subBrush, px, 82);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 0 || r.Width <= 0 || r.Height <= 0) { path.AddRectangle(r); return path; }
            int d = radius * 2;
            if (d > r.Width) d = r.Width;
            if (d > r.Height) d = r.Height;
            if (d <= 0) { path.AddRectangle(r); return path; }
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    // ===== SIDEBAR ITEM =====
    public class SidebarItem : Panel
    {
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string ItemText { get; set; } = "";

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string ItemIcon { get; set; } = "";

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string PageKey { get; set; } = "";

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int BadgeCount { get; set; } = -1;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public bool IsActive { get; set; }

        public event EventHandler? Clicked;

        private bool _hovered;

        public SidebarItem()
        {
            Height = 40;
            Cursor = Cursors.Hand;
            BackColor = Theme.SidebarBg;
        }

        protected override void OnMouseEnter(EventArgs e) { _hovered = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hovered = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnClick(EventArgs e) { Clicked?.Invoke(this, e); base.OnClick(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

            Color bgColor = IsActive ? Theme.SidebarHover : (_hovered ? Theme.SidebarHover : Theme.SidebarBg);
            g.FillRectangle(new SolidBrush(bgColor), 0, 0, Width, Height);

            if (IsActive && Height - 16 > 0)
            {
                g.FillRectangle(new SolidBrush(Theme.SidebarActive), 0, 8, 3, Height - 16);
                using var glow = new LinearGradientBrush(new Rectangle(3, 8, 10, Height - 16), Color.FromArgb(40, Theme.SidebarActive), Color.Transparent, 0f);
                g.FillRectangle(glow, 3, 8, 10, Height - 16);
            }

            var textColor = IsActive ? Theme.SidebarActive : Theme.SidebarText;
            using var tb = new SolidBrush(textColor);
            var font = IsActive ? Theme.FontNavBold : Theme.FontNav;
            
            int textX = 16;
            if (!string.IsNullOrEmpty(ItemIcon))
            {
                using var iconFont = new Font("Segoe Fluent Icons", 14f, FontStyle.Regular);
                g.DrawString(ItemIcon, iconFont, tb, 16, (Height - iconFont.Height) / 2f + 1);
                textX = 46;
            }

            g.DrawString(ItemText, font, tb, textX, (Height - font.Height) / 2f);

            if (BadgeCount >= 0)
            {
                string badgeText = BadgeCount.ToString();
                var badgeFont = Theme.FontSmall;
                var sz = g.MeasureString(badgeText, badgeFont);
                int bw = Math.Max(22, (int)sz.Width + 10);
                int bx = Width - bw - 12;
                int by = (Height - 18) / 2;
                using var badgeBrush = new SolidBrush(IsActive ? Theme.SidebarActive : Color.FromArgb(55, 65, 88));
                g.FillRectangle(badgeBrush, bx, by, bw, 18);
                using var badgeTextBrush = new SolidBrush(Color.White);
                var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                g.DrawString(badgeText, badgeFont, badgeTextBrush, new RectangleF(bx, by, bw, 18), sf);
            }
        }
    }

    // ===== CARD PANEL =====
    public class CardPanel : Panel
    {
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int Radius { get; set; } = 24;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color BorderColor { get; set; } = Color.Transparent;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color CardColor { get; set; } = Theme.BgSurface;

        public CardPanel()
        {
            BackColor = Color.Transparent;
            DoubleBuffered = true;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using var path = RoundedRect(r, Radius);
            
            using var baseBrush = new SolidBrush(CardColor);
            g.FillPath(baseBrush, path);

            // Glass Highlight Border
            var borderColor = BorderColor == Color.Transparent ? Color.FromArgb(25, Theme.TextPrimary) : BorderColor;
            using var pen = new Pen(borderColor, 1);
            g.DrawPath(pen, path);

            // Inner glow
            var innerR = new Rectangle(1, 1, Width - 3, Height - 3);
            using var innerPath = RoundedRect(innerR, Radius - 1);
            using var innerPen = new Pen(Color.FromArgb(10, Theme.TextPrimary), 1f);
            g.DrawPath(innerPen, innerPath);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 0 || r.Width <= 0 || r.Height <= 0) { path.AddRectangle(r); return path; }
            int d = radius * 2;
            if (d > r.Width) d = r.Width;
            if (d > r.Height) d = r.Height;
            if (d <= 0) { path.AddRectangle(r); return path; }
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            var path = new GraphicsPath();
            var r = new Rectangle(0, 0, Width, Height);
            if (Radius <= 0 || r.Width <= 0 || r.Height <= 0) { path.AddRectangle(r); Region = new Region(path); return; }
            int d = Radius * 2;
            if (d > r.Width) d = r.Width;
            if (d > r.Height) d = r.Height;
            if (d <= 0) { path.AddRectangle(r); Region = new Region(path); return; }
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            Region = new Region(path);
        }
    }

    // ===== MODERN TEXTBOX =====
    public class ModernTextBox : Panel
    {
        public TextBox Inner { get; }
        private bool _focused;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public new string Text
        {
            get => Inner.Text;
            set => Inner.Text = value;
        }

        public new event EventHandler? TextChanged;

        public ModernTextBox()
        {
            Height = 36;
            BackColor = Theme.BgSurface2;
            Padding = new Padding(1);

            Inner = new TextBox
            {
                BorderStyle = BorderStyle.None,
                BackColor = Theme.BgSurface2,
                ForeColor = Theme.TextPrimary,
                Font = Theme.FontBody,
                Dock = DockStyle.Fill,
            };
            Inner.GotFocus += (s, e) => { _focused = true; Invalidate(); };
            Inner.LostFocus += (s, e) => { _focused = false; Invalidate(); };
            Inner.TextChanged += (s, e) => TextChanged?.Invoke(this, e);

            var wrapper = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8, 0, 8, 0), BackColor = Theme.BgSurface2 };
            wrapper.Controls.Add(Inner);
            Controls.Add(wrapper);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var pen = _focused
                ? new Pen(Theme.Accent, 1.5f)
                : new Pen(Theme.Border, 1f);
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            g.DrawRectangle(pen, r);
            pen.Dispose();
        }
    }

    // ===== BADGE LABEL =====
    public class BadgeLabel : Label
    {
        public enum BadgeType { Blue, Green, Yellow, Red }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public BadgeType Badge { get; set; } = BadgeType.Blue;

        public BadgeLabel()
        {
            AutoSize = false;
            Height = 20;
            TextAlign = ContentAlignment.MiddleCenter;
            Font = new Font("Segoe UI", 7.5f, FontStyle.Bold);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            
            Control? p = this.Parent;
            Color baseBg = Theme.BgMain;
            while (p != null) { if (p.BackColor != Color.Transparent) { baseBg = p.BackColor; break; } p = p.Parent; }
            g.Clear(baseBg);

            Color bg; Color fg;
            switch (Badge)
            {
                case BadgeType.Green:
                    bg = Color.FromArgb(40, 16, 185, 129); fg = Color.FromArgb(52, 211, 153); break; // Approved
                case BadgeType.Yellow:
                    bg = Color.FromArgb(40, 245, 158, 11); fg = Color.FromArgb(252, 211, 77); break; // Pending
                case BadgeType.Red:
                    bg = Color.FromArgb(40, 239, 68, 68); fg = Color.FromArgb(248, 113, 113); break;  // Rejected
                default:
                    bg = Color.FromArgb(40, 59, 130, 246); fg = Color.FromArgb(96, 165, 250); break;  // Blue
            }
            using var brush = new SolidBrush(bg);
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using var path = RoundedRect(r, Height / 2); // Fully pill shaped
            g.FillPath(brush, path);
            using var borderPen = new Pen(Color.FromArgb(80, fg), 1);
            g.DrawPath(borderPen, path);
            using var tb = new SolidBrush(fg);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(Text, Font, tb, new RectangleF(0, 0, Width, Height), sf);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            if (radius <= 0 || r.Width <= 0 || r.Height <= 0) { path.AddRectangle(r); return path; }
            int d = radius * 2;
            if (d > r.Width) d = r.Width;
            if (d > r.Height) d = r.Height;
            if (d <= 0) { path.AddRectangle(r); return path; }
            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }

    // ===== MODERN COMBOBOX =====
    public class ModernComboBox : ComboBox
    {
        public ModernComboBox()
        {
            FlatStyle = FlatStyle.Flat;
            Font = Theme.FontBody;
            BackColor = Theme.BgSurface2;
            ForeColor = Theme.TextPrimary;
            DrawMode = DrawMode.OwnerDrawFixed;
            ItemHeight = 24;
            DropDownStyle = ComboBoxStyle.DropDownList;
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            e.DrawBackground();
            var g = e.Graphics;
            bool selected = (e.State & DrawItemState.Selected) != 0;
            var bg = selected ? Theme.AccentLight : Theme.BgSurface2;
            var fg = selected ? Theme.Accent : Theme.TextPrimary;
            g.FillRectangle(new SolidBrush(bg), e.Bounds);
            g.DrawString(Items[e.Index]?.ToString() ?? "", Theme.FontBody, new SolidBrush(fg),
                e.Bounds.X + 6, e.Bounds.Y + (e.Bounds.Height - Theme.FontBody.Height) / 2);
        }
    }
}
