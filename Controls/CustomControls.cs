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

        public ModernButton()
        {
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Height = 34;
            Font = Theme.FontBold;
            Cursor = Cursors.Hand;
        }

        protected override void OnMouseEnter(EventArgs e) { _hovered = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { _hovered = false; Invalidate(); base.OnMouseLeave(e); }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            var path = RoundedRect(r, 6);

            Color bg; Color fg;
            switch (Style)
            {
                case BtnStyle.Success:
                    bg = _hovered ? Color.FromArgb(22, 163, 120) : Theme.AccentGreen;
                    fg = Color.White; break;
                case BtnStyle.Danger:
                    bg = _hovered ? Theme.DangerHover : Theme.Danger;
                    fg = Color.White; break;
                case BtnStyle.Secondary:
                    bg = _hovered ? Theme.BgSurface3 : Theme.BgSurface2;
                    fg = Theme.TextSecondary; break;
                default:
                    bg = _hovered ? Theme.AccentHover : Theme.Accent;
                    fg = Color.White; break;
            }

            using var brush = new SolidBrush(bg);
            g.FillPath(brush, path);

            if (Style == BtnStyle.Secondary)
            {
                using var pen = new Pen(Theme.Border, 1);
                g.DrawPath(pen, path);
            }

            using var tb = new SolidBrush(fg);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(Text, Font, tb, new RectangleF(0, 0, Width, Height), sf);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(r.Right - radius * 2, r.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(r.Right - radius * 2, r.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(r.X, r.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
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
            BackColor = Theme.BgSurface;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using var bgBrush = new SolidBrush(Theme.BgSurface);
            using var path = RoundedRect(r, 10);
            g.FillPath(bgBrush, path);

            using var borderPen = new Pen(Theme.Border, 1f);
            g.DrawPath(borderPen, path);

            using var accentBrush = new SolidBrush(AccentColor);
            g.FillRectangle(accentBrush, 0, 14, 4, Height - 28);

            int px = 18;
            using var labelBrush = new SolidBrush(Theme.TextMuted);
            g.DrawString(StatLabel.ToUpper(), new Font("Segoe UI", 8f, FontStyle.Bold), labelBrush, px, 16);

            using var valueBrush = new SolidBrush(AccentColor);
            using var valueFont = new Font("Segoe UI", 24f, FontStyle.Bold);
            g.DrawString(StatValue, valueFont, valueBrush, px, 32);

            using var subBrush = new SolidBrush(Theme.TextMuted);
            g.DrawString(StatSub, Theme.FontSmall, subBrush, px, 72);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(r.Right - radius * 2, r.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(r.Right - radius * 2, r.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(r.X, r.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
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

            Color bgColor = IsActive ? Theme.SidebarHover : (_hovered ? Theme.SidebarHover : Theme.SidebarBg);
            g.FillRectangle(new SolidBrush(bgColor), 0, 0, Width, Height);

            if (IsActive)
                g.FillRectangle(new SolidBrush(Theme.SidebarActive), 0, 8, 3, Height - 16);

            var textColor = IsActive ? Color.White : Theme.SidebarText;
            using var tb = new SolidBrush(textColor);
            var font = IsActive ? Theme.FontNavBold : Theme.FontNav;
            g.DrawString(ItemText, font, tb, 16, (Height - font.Height) / 2f);

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
        public int Radius { get; set; } = 10;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public Color BorderColor { get; set; } = Color.Transparent;

        public CardPanel()
        {
            BackColor = Theme.BgSurface;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            var color = BorderColor == Color.Transparent ? Theme.Border : BorderColor;
            using var pen = new Pen(color, 1);
            using var path = RoundedRect(r, Radius);
            g.DrawPath(pen, path);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(r.Right - radius * 2, r.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(r.Right - radius * 2, r.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(r.X, r.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            var path = new GraphicsPath();
            var r = new Rectangle(0, 0, Width, Height);
            path.AddArc(r.X, r.Y, Radius * 2, Radius * 2, 180, 90);
            path.AddArc(r.Right - Radius * 2, r.Y, Radius * 2, Radius * 2, 270, 90);
            path.AddArc(r.Right - Radius * 2, r.Bottom - Radius * 2, Radius * 2, Radius * 2, 0, 90);
            path.AddArc(r.X, r.Bottom - Radius * 2, Radius * 2, Radius * 2, 90, 90);
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
            Color bg; Color fg;
            switch (Badge)
            {
                case BadgeType.Green:
                    bg = Color.FromArgb(209, 250, 229); fg = Color.FromArgb(6, 95, 70); break;
                case BadgeType.Yellow:
                    bg = Color.FromArgb(254, 243, 199); fg = Color.FromArgb(120, 80, 4); break;
                case BadgeType.Red:
                    bg = Color.FromArgb(254, 226, 226); fg = Color.FromArgb(153, 27, 27); break;
                default:
                    bg = Color.FromArgb(219, 234, 254); fg = Color.FromArgb(29, 78, 216); break;
            }
            using var brush = new SolidBrush(bg);
            var r = new Rectangle(0, 0, Width - 1, Height - 1);
            using var path = RoundedRect(r, 4);
            g.FillPath(brush, path);
            using var tb = new SolidBrush(fg);
            var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
            g.DrawString(Text, Font, tb, new RectangleF(0, 0, Width, Height), sf);
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            path.AddArc(r.X, r.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(r.Right - radius * 2, r.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(r.Right - radius * 2, r.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(r.X, r.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
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
