using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin
{
    public class ValidationMessage
    {
        public int Index { get; set; }
        public string MsgId { get; set; }
        public string Category { get; set; }
        public string Severity { get; set; }
        public string Message { get; set; }
        public string Excerpt { get; set; }
        public int ParaIndex { get; set; }
        public bool Dismissed { get; set; }
    }

    public class TaskPaneWinForms : UserControl
    {
        private static List<ValidationMessage> _messages = new List<ValidationMessage>();
        private static int _msgCounter = 0;
        private static TaskPaneWinForms _instance;

        private static readonly Color C_HeaderBg = Color.FromArgb(21, 101, 192);
        private static readonly Color C_AmsBadge = Color.FromArgb(13, 71, 161);
        private static readonly Color C_NavBg = Color.FromArgb(240, 244, 255);
        private static readonly Color C_CardBg = Color.White;
        private static readonly Color C_CardBorder = Color.FromArgb(224, 228, 234);
        private static readonly Color C_Excerpt = Color.FromArgb(100, 116, 139);
        private static readonly Color C_JumpLink = Color.FromArgb(37, 99, 235);
        private static readonly Color C_PanelBg = Color.FromArgb(245, 247, 250);
        private static readonly Color C_Error = Color.FromArgb(192, 57, 43);
        private static readonly Color C_Warning = Color.FromArgb(230, 126, 34);
        private static readonly Color C_Info = Color.FromArgb(39, 174, 96);

        private Panel pnlHeader;
        private Panel pnlToolbar;
        private ComboBox cmbFilter;
        private Panel pnlNavBar;
        private Button btnNavPrev, btnNavNext;
        private Label lblIssueCount, lblShortcut;
        private FlowLayoutPanel pnlBadgeBar;
        private Panel pnlBadgeBorder;
        private Panel pnlEmpty;
        private Panel pnlCardContainer;
        private FlowLayoutPanel pnlCards;
        private Panel pnlStatus;
        private Label lblStatus, lblVersion;

        private List<ValidationMessage> _shown = new List<ValidationMessage>();
        private int _navIndex = -1;
        private string _activeFilter = "All";

        public TaskPaneWinForms()
        {
            _instance = this;
            this.Dock = DockStyle.Fill;
            this.BackColor = C_PanelBg;
            BuildUI();
            RefreshList();
        }

        private void BuildUI()
        {
            this.SuspendLayout();

            pnlHeader = new Panel { Dock = DockStyle.Top, Height = 52, BackColor = C_HeaderBg };
            var lblTitle = new Label { Text = "MechEdit Error Checker", ForeColor = Color.White, Font = new Font("Segoe UI", 11, FontStyle.Bold), AutoSize = false, Left = 12, Top = 9, Width = 200, Height = 22 };
            var lblSub = new Label { Text = "AMS Manuscript Style Validator", ForeColor = Color.FromArgb(144, 202, 249), Font = new Font("Segoe UI", 8), AutoSize = false, Left = 12, Top = 30, Width = 220, Height = 16 };
            var lblAMS = new Label { Text = "AMS", ForeColor = Color.FromArgb(187, 222, 251), BackColor = C_AmsBadge, Font = new Font("Segoe UI", 8, FontStyle.Bold), AutoSize = false, Width = 34, Height = 18, TextAlign = ContentAlignment.MiddleCenter, Left = 268, Top = 17 };
            pnlHeader.Controls.AddRange(new Control[] { lblTitle, lblSub, lblAMS });

            pnlToolbar = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = Color.White, Padding = new Padding(8, 7, 8, 7) };
            var toolBorder = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = C_CardBorder };
            var lblFilter = new Label { Text = "Filter:", Font = new Font("Segoe UI", 9), ForeColor = Color.FromArgb(55, 65, 81), AutoSize = true, Left = 8, Top = 13 };
            cmbFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 9), Width = 110, Left = 50, Top = 9, FlatStyle = FlatStyle.Flat };
            cmbFilter.Items.AddRange(new object[] { "All", "ERROR", "WARNING", "INFO" });
            cmbFilter.SelectedIndex = 0;
            cmbFilter.SelectedIndexChanged += (s, e) => { _activeFilter = cmbFilter.SelectedItem.ToString(); RefreshList(); };
            pnlToolbar.Controls.AddRange(new Control[] { toolBorder, lblFilter, cmbFilter });

            pnlNavBar = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = C_NavBg, Visible = false };
            var navBorder = new Panel { Dock = DockStyle.Bottom, Height = 1, BackColor = Color.FromArgb(209, 213, 219) };
            btnNavPrev = MakeNavBtn("▲"); btnNavPrev.Left = 8; btnNavPrev.Top = 4; btnNavPrev.Click += (s, e) => NavPrev();
            btnNavNext = MakeNavBtn("▼"); btnNavNext.Left = 36; btnNavNext.Top = 4; btnNavNext.Click += (s, e) => NavNext();
            lblIssueCount = new Label { Text = "", ForeColor = Color.FromArgb(55, 65, 81), Font = new Font("Segoe UI", 8.5f, FontStyle.Bold), AutoSize = true, Left = 68, Top = 8 };
            lblShortcut = new Label { Text = "Ctrl+Shift+S = Scan", ForeColor = Color.FromArgb(156, 163, 175), Font = new Font("Segoe UI", 7.5f), AutoSize = false, Width = 120, Height = 16, TextAlign = ContentAlignment.MiddleRight, Anchor = AnchorStyles.Right | AnchorStyles.Top };
            pnlNavBar.Controls.AddRange(new Control[] { btnNavPrev, btnNavNext, lblIssueCount, lblShortcut, navBorder });
            pnlNavBar.Resize += (s, e) => { lblShortcut.Left = pnlNavBar.Width - 128; lblShortcut.Top = 8; };

            pnlBadgeBar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34, BackColor = Color.FromArgb(249, 250, 251), FlowDirection = FlowDirection.LeftToRight, WrapContents = false, AutoSize = false, Padding = new Padding(8, 6, 8, 4), Visible = false };
            pnlBadgeBorder = new Panel { Dock = DockStyle.Top, Height = 1, BackColor = Color.FromArgb(229, 231, 235), Visible = false };

            pnlEmpty = new Panel { Dock = DockStyle.Fill, BackColor = C_PanelBg, Visible = true };
            var lblEmptyIcon = new Label { Text = "◎", Font = new Font("Segoe UI", 32), ForeColor = Color.FromArgb(203, 213, 225), AutoSize = false, Width = 310, Height = 52, Top = 60, TextAlign = ContentAlignment.MiddleCenter };
            var lblEmptyText = new Label { Text = "No issues found yet.", Font = new Font("Segoe UI", 11, FontStyle.Bold), ForeColor = Color.FromArgb(148, 163, 184), AutoSize = false, Width = 310, Height = 26, Top = 118, TextAlign = ContentAlignment.MiddleCenter };
            var lblEmptyHint = new Label { Text = "Click a button in the AMS Style ribbon to scan.", Font = new Font("Segoe UI", 8.5f), ForeColor = Color.FromArgb(203, 213, 225), AutoSize = false, Width = 310, Height = 20, Top = 148, TextAlign = ContentAlignment.MiddleCenter };
            pnlEmpty.Controls.AddRange(new Control[] { lblEmptyIcon, lblEmptyText, lblEmptyHint });

            pnlCardContainer = new Panel { Dock = DockStyle.Fill, BackColor = C_PanelBg, AutoScroll = true, Visible = false };
            pnlCards = new FlowLayoutPanel { Dock = DockStyle.Top, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, BackColor = C_PanelBg, Padding = new Padding(8, 6, 8, 8) };
            pnlCardContainer.Controls.Add(pnlCards);

            pnlStatus = new Panel { Dock = DockStyle.Bottom, Height = 26, BackColor = C_HeaderBg };
            lblStatus = new Label { Text = "Ready", ForeColor = Color.FromArgb(186, 230, 253), Font = new Font("Segoe UI", 8f), AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(10, 0, 0, 0) };
            lblVersion = new Label { Text = "v1.0", ForeColor = Color.FromArgb(96, 165, 250), Font = new Font("Segoe UI", 7.5f), AutoSize = false, Width = 32, Dock = DockStyle.Right, TextAlign = ContentAlignment.MiddleCenter };
            pnlStatus.Controls.AddRange(new Control[] { lblVersion, lblStatus });

            this.Controls.Add(pnlCardContainer);
            this.Controls.Add(pnlEmpty);
            this.Controls.Add(pnlStatus);
            this.Controls.Add(pnlBadgeBorder);
            this.Controls.Add(pnlBadgeBar);
            this.Controls.Add(pnlNavBar);
            this.Controls.Add(pnlToolbar);
            this.Controls.Add(pnlHeader);
            this.ResumeLayout(false);
        }

        private Panel BuildCard(ValidationMessage msg)
        {
            Color badgeColor = GetCategoryColor(msg.Category);
            int cardWidth = pnlCardContainer.ClientSize.Width > 40 ? pnlCardContainer.ClientSize.Width - 20 : 300;
            var card = new Panel { BackColor = C_CardBg, Margin = new Padding(0, 0, 0, 6), Padding = new Padding(14, 8, 10, 8), Cursor = Cursors.Hand, Tag = msg, Width = cardWidth, AutoSize = false };
            card.Paint += (s, e) =>
            {
                var g = e.Graphics; var rc = card.ClientRectangle;
                g.FillRectangle(new SolidBrush(badgeColor), new Rectangle(0, 0, 4, rc.Height));
                using (var pen = new Pen(C_CardBorder)) g.DrawRectangle(pen, new Rectangle(0, 0, rc.Width - 1, rc.Height - 1));
            };
            int y = 8;
            var lblBadge = new Label { Text = msg.Category, BackColor = badgeColor, ForeColor = Color.White, Font = new Font("Segoe UI", 7.5f, FontStyle.Bold), AutoSize = false, Height = 18, Width = TextRenderer.MeasureText(msg.Category, new Font("Segoe UI", 7.5f, FontStyle.Bold)).Width + 14, Left = 12, Top = y, TextAlign = ContentAlignment.MiddleCenter, Cursor = Cursors.Hand };
            var lblNum = new Label { Text = $"#{msg.Index}", ForeColor = Color.FromArgb(156, 163, 175), Font = new Font("Segoe UI", 8f), AutoSize = true, Left = lblBadge.Left + lblBadge.Width + 6, Top = y + 1, Cursor = Cursors.Hand };
            var btnX = new Label { Text = "✕", ForeColor = Color.FromArgb(200, 200, 200), Font = new Font("Segoe UI", 8f), AutoSize = true, Cursor = Cursors.Hand };
            btnX.Click += (s, e) => { msg.Dismissed = true; RefreshList(); };
            btnX.MouseEnter += (s, e) => btnX.ForeColor = C_Error;
            btnX.MouseLeave += (s, e) => btnX.ForeColor = Color.FromArgb(200, 200, 200);
            y += 22;
            var lblMsg = new Label { Text = msg.Message, ForeColor = Color.FromArgb(30, 41, 59), Font = new Font("Segoe UI", 8.5f), AutoSize = false, Width = card.Width - 28, Height = 34, Left = 12, Top = y, Cursor = Cursors.Hand };
            y += 36;
            Label lblExcerpt = null;
            if (!string.IsNullOrEmpty(msg.Excerpt))
            {
                lblExcerpt = new Label { Text = Truncate(msg.Excerpt, 60), ForeColor = C_Excerpt, Font = new Font("Segoe UI", 8f, FontStyle.Italic), AutoSize = false, Width = cardWidth - 28, Height = 16, Left = 18, Top = y, Cursor = Cursors.Hand };
                y += 18;
            }
            var lblJump = new LinkLabel { Text = "→ Jump to location in Word", Font = new Font("Segoe UI", 8f), AutoSize = true, Left = 12, Top = y, LinkColor = C_JumpLink, ActiveLinkColor = C_JumpLink, Cursor = Cursors.Hand };
            lblJump.LinkClicked += (s, e) => JumpToMessage(msg);
            y += 20;
            card.Height = y + 6;
            card.Controls.AddRange(new Control[] { lblBadge, lblNum, lblMsg, lblJump, btnX });
            if (lblExcerpt != null) card.Controls.Add(lblExcerpt);
            card.Layout += (s, e) =>
            {
                btnX.Left = card.Width - btnX.Width - 10; btnX.Top = 10;
                lblMsg.Width = card.Width - 28;
                if (lblExcerpt != null) lblExcerpt.Width = card.Width - 28;
            };
            Action jump = () => JumpToMessage(msg);
            card.Click += (s, e) => jump(); lblBadge.Click += (s, e) => jump();
            lblNum.Click += (s, e) => jump(); lblMsg.Click += (s, e) => jump();
            if (lblExcerpt != null) lblExcerpt.Click += (s, e) => jump();
            return card;
        }

        private Color GetCategoryColor(string cat)
        {
            switch ((cat ?? "").ToUpper())
            {
                case "CASING": return Color.FromArgb(124, 58, 237);   // purple
                case "HEADING": return Color.FromArgb(37, 99, 235);    // blue
                case "PARATEXT": return Color.FromArgb(16, 185, 129);   // teal
                case "REF-ORD": return Color.FromArgb(192, 57, 43);    // red
                case "REF-WARN": return Color.FromArgb(180, 150, 0);    // dark yellow
                case "REF-BLUE": return Color.FromArgb(2, 119, 189);    // ← NEW: blue badge for same-surname/count/year/diff-initial rule
                case "REF-FMT": return Color.FromArgb(220, 38, 38);    // red
                case "FIG-CITE": return Color.FromArgb(5, 150, 105);    // green
                case "REF-PUNCT": return Color.FromArgb(234, 88, 12);    // orange
                case "CITATION": return Color.FromArgb(2, 132, 199);    // sky blue
                case "INFO": return C_Info;
                case "WARNING": return C_Warning;
                default: return C_Error;
            }
        }

        private void OnContainerSizeChanged(object sender, EventArgs e)
        {
            int newWidth = pnlCardContainer.ClientSize.Width - 20;
            if (newWidth < 40) return;
            foreach (Control c in pnlCards.Controls)
            {
                if (c is Panel card)
                {
                    card.Width = newWidth;
                    foreach (Control child in card.Controls)
                        if (child is Label lbl && lbl.AutoSize == false && lbl.Width > 50)
                            lbl.Width = newWidth - 28;
                }
            }
        }

        private string Truncate(string s, int max) => s == null ? "" : s.Length <= max ? s : s.Substring(0, max) + "…";

        private Button MakeNavBtn(string text)
        {
            var btn = new Button { Text = text, Width = 24, Height = 22, FlatStyle = FlatStyle.Flat, BackColor = Color.Transparent, ForeColor = Color.FromArgb(55, 65, 81), Font = new Font("Segoe UI", 8, FontStyle.Bold), Cursor = Cursors.Hand };
            btn.FlatAppearance.BorderSize = 0;
            return btn;
        }

        private void NavPrev()
        {
            if (_shown.Count == 0) return;
            _navIndex = (_navIndex - 1 + _shown.Count) % _shown.Count;
            ScrollToCard(_navIndex); JumpToMessage(_shown[_navIndex]); UpdateNavCounter();
        }

        private void NavNext()
        {
            if (_shown.Count == 0) return;
            _navIndex = (_navIndex + 1) % _shown.Count;
            ScrollToCard(_navIndex); JumpToMessage(_shown[_navIndex]); UpdateNavCounter();
        }

        private void ScrollToCard(int index)
        {
            if (index < 0 || index >= pnlCards.Controls.Count) return;
            pnlCardContainer.ScrollControlIntoView(pnlCards.Controls[index]);
        }

        private void UpdateNavCounter()
        {
            int total = _shown.Count;
            lblIssueCount.Text = total == 0 ? "No issues" : _navIndex >= 0 ? $"{_navIndex + 1} of {total} issue(s)" : $"{total} issue(s)";
            btnNavPrev.Enabled = total > 0;
            btnNavNext.Enabled = total > 0;
        }

        private void JumpToMessage(ValidationMessage msg)
        {
            try
            {
                if (msg.ParaIndex <= 0) return;
                Word.Document doc = Globals.ThisAddIn.Application.ActiveDocument;
                int end = Math.Min(msg.ParaIndex + 1, doc.Content.End);
                Word.Range range = doc.Range(msg.ParaIndex, end);
                range.Select();
                Globals.ThisAddIn.Application.ActiveWindow.ScrollIntoView(range, true);
            }
            catch { }
        }

        public static void SetProgress(string text)
        {
            var inst = _instance;
            if (inst == null) return;
            if (inst.InvokeRequired)
            {
                try { inst.BeginInvoke(new Action(() => { if (inst.lblStatus != null) inst.lblStatus.Text = text ?? ""; })); }
                catch { }
            }
            else { if (inst.lblStatus != null) inst.lblStatus.Text = text ?? ""; }
        }

        public void SetStatus()
        {
            _batchMode = false;
            RefreshList();
            int count = _messages.Count(m => !m.Dismissed);
            lblStatus.Text = count == 0 ? "No issues found." : $"{count} issue(s) found.";
        }

        public void BeginScan()
        {
            _batchMode = true;
            ClearMessages();
        }

        public void RunScan(bool scanBody, bool scanRefs)
        {
            lblStatus.Text = "Scanning…";
            Application.DoEvents();
            try
            {
                _batchMode = true;
                ClearMessages();
                if (scanRefs)
                {
                    new Checkers.ReferenceReorderingChecker().Run();
                    new Checkers.ReferenceFormatChecker().Run();
                }
                if (scanBody)
                {
                    new Checkers.ParatextChecker().Run();
                    new Checkers.HeadingAnatomyChecker().Run();
                    new Checkers.CasingErrorChecker().Run();
                }
                _batchMode = false;
                RefreshList();
                int count = _messages.Count(m => !m.Dismissed);
                lblStatus.Text = count == 0 ? "No issues found." : $"{count} issue(s) found.";
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Scan error.";
                MessageBox.Show("Scan failed:\n" + ex.Message, "AMS Style", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static bool _batchMode = false;

        public static void AddMessage(string category, string severity, string message, string excerpt = "", int paraIndex = 0)
        {
            _msgCounter++;
            _messages.Add(new ValidationMessage { Index = _msgCounter, MsgId = "Msg" + _msgCounter, Category = category.ToUpper(), Severity = severity.ToUpper(), Message = message, Excerpt = excerpt, ParaIndex = paraIndex, Dismissed = false });
            if (!_batchMode) _instance?.RefreshList();
        }

        public static void ClearMessages()
        {
            _messages.Clear();
            _msgCounter = 0;
            _instance?.RefreshList();
        }

        private void RefreshList()
        {
            if (this.InvokeRequired) { this.Invoke(new Action(RefreshList)); return; }
            _shown = _messages.Where(m => !m.Dismissed).Where(m => _activeFilter == "All" || m.Severity == _activeFilter).ToList();
            bool hasItems = _shown.Count > 0;
            pnlEmpty.Visible = !hasItems; pnlCardContainer.Visible = hasItems;
            pnlNavBar.Visible = hasItems; pnlBadgeBar.Visible = hasItems; pnlBadgeBorder.Visible = hasItems;
            pnlCardContainer.PerformLayout();
            pnlCards.SuspendLayout();
            pnlCards.Controls.Clear();
            foreach (var m in _shown) pnlCards.Controls.Add(BuildCard(m));
            pnlCards.ResumeLayout(true);
            pnlCardContainer.SizeChanged -= OnContainerSizeChanged;
            pnlCardContainer.SizeChanged += OnContainerSizeChanged;
            RebuildBadges(); _navIndex = -1; UpdateNavCounter();
        }

        private void RebuildBadges()
        {
            pnlBadgeBar.Controls.Clear();
            if (_messages.Count == 0) return;
            var active = _messages.Where(m => !m.Dismissed).ToList();
            foreach (var g in active.GroupBy(m => m.Category).OrderBy(g => g.Key))
            {
                string cat = g.Key; Color color = GetCategoryColor(cat);
                var badge = new Button { Text = $"{cat} {g.Count()}", BackColor = color, ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 7.5f, FontStyle.Bold), Height = 22, Width = TextRenderer.MeasureText($"{cat} {g.Count()}", new Font("Segoe UI", 7.5f, FontStyle.Bold)).Width + 16, Margin = new Padding(0, 0, 4, 0), Cursor = Cursors.Hand };
                badge.FlatAppearance.BorderSize = 0;
                badge.Click += (s, e) => { var first = _shown.FirstOrDefault(m => m.Category == cat); if (first != null) { ScrollToCard(_shown.IndexOf(first)); JumpToMessage(first); } };
                pnlBadgeBar.Controls.Add(badge);
            }
        }
    }

    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        public static TaskPaneWinForms TaskPaneUI;
        private static ThisAddIn _instance;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            _instance = this;
            try
            {
                TaskPaneUI = new TaskPaneWinForms();
                _taskPane = this.CustomTaskPanes.Add(TaskPaneUI, "MechEdit Error Checker");
                _taskPane.Width = 340;
                _taskPane.Visible = false;
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("AMS Addin startup: " + ex.Message); }
        }

        public static void ShowPane() { if (_instance?._taskPane == null) return; _instance._taskPane.Visible = true; }
        public static void ShowAndScan(bool scanBody, bool scanRefs) { ShowPane(); TaskPaneUI?.RunScan(scanBody, scanRefs); }
        public static void UpdateStatus() { TaskPaneUI?.SetStatus(); }

        private void ThisAddIn_Shutdown(object sender, EventArgs e) { }
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject() => new Ribbon1();

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}