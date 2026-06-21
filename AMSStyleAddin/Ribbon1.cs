using System;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Office = Microsoft.Office.Core;

namespace AMSStyleAddin
{
    [System.Runtime.InteropServices.ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1() { }

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("AMSStyleAddin.Ribbon1.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        // ── Button 1: Reference Reordering ───────────────────────────────────
        public void OnReferenceReordering(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.ShowPane();
                ThisAddIn.TaskPaneUI.BeginScan();
                new Checkers.ReferenceReorderingChecker().Run();
                ThisAddIn.TaskPaneUI.SetStatus();
            }
            catch (Exception ex) { ShowError("Reference Reordering", ex); }
        }

        // ── Button 2: Paratext Checker ───────────────────────────────────────
        public void OnParatextChecker(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.ShowPane();
                ThisAddIn.TaskPaneUI.BeginScan();
                new Checkers.ParatextChecker().Run();
                ThisAddIn.TaskPaneUI.SetStatus();
            }
            catch (Exception ex) { ShowError("Paratext Checker", ex); }
        }

        // ── Button 3: Heading Anatomy ────────────────────────────────────────
        public void OnHeadingAnatomy(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.ShowPane();
                ThisAddIn.TaskPaneUI.BeginScan();
                new Checkers.HeadingAnatomyChecker().Run();
                ThisAddIn.TaskPaneUI.SetStatus();
            }
            catch (Exception ex) { ShowError("Heading Anatomy", ex); }
        }

        // ── Button 4: Casing Error Checker ───────────────────────────────────
        public void OnCasingChecker(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.ShowPane();
                ThisAddIn.TaskPaneUI.BeginScan();
                new Checkers.CasingErrorChecker().Run();
                ThisAddIn.TaskPaneUI.SetStatus();
            }
            catch (Exception ex) { ShowError("Casing Checker", ex); }
        }

        // ── Button 5: Reference Format ───────────────────────────────────────
        public void OnReferenceFormat(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.ShowPane();
                ThisAddIn.TaskPaneUI.BeginScan();
                new Checkers.ReferenceFormatChecker().Run();
            }
            catch (Exception ex) { ShowError("Reference Format", ex); }
        }

        // ── Button 6: Figure and Table Citation Checker ──────────────────────
        public void OnFigTableCitation(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.ShowPane();
                ThisAddIn.TaskPaneUI.BeginScan();
                new Checkers.FigureTableCitationChecker().Run();
                ThisAddIn.TaskPaneUI.SetStatus();
            }
            catch (Exception ex) { ShowError("Fig/Table Citation", ex); }
        }

        // ── Button 7: Reference Punctuation Checker ──────────────────────────
        public void OnReferencePunctuation(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.ShowPane();
                ThisAddIn.TaskPaneUI.BeginScan();
                new Checkers.ReferencePunctuationChecker().Run();
                ThisAddIn.TaskPaneUI.SetStatus();
            }
            catch (Exception ex) { ShowError("Reference Punctuation", ex); }
        }

        // ── Button 8: Citation Checker ───────────────────────────────────────
        public void OnCitationChecker(Office.IRibbonControl control)
        {
            try
            {
                ThisAddIn.ShowPane();
                ThisAddIn.TaskPaneUI.BeginScan();
                new Checkers.CitationChecker().Run();
                ThisAddIn.TaskPaneUI.SetStatus();
            }
            catch (Exception ex) { ShowError("Citation Checker", ex); }
        }

        // ── Helper ───────────────────────────────────────────────────────────
        private void ShowError(string checker, Exception ex)
        {
            MessageBox.Show($"{checker} error:\n{ex.Message}",
                "AMS Style", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static string GetResourceText(string resourceName)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream(resourceName))
            using (var reader = new StreamReader(stream))
                return reader.ReadToEnd();
        }
    }
}