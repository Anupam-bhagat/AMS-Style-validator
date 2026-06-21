using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin.Checkers
{
    public class FigureTableCitationChecker
    {
        private const int ChunkSize = 30;
        private const int SleepMs = 20;

        // =====================================================================
        // SCOPED STYLES
        // =====================================================================
        private static readonly HashSet<string> ScopedStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "figurecaption", "tablecaption", "tablebody", "tablehead",
            "paratext", "acknowledgementtext", "abstracttext", "synopsis",
            "paranoindent", "numberedlistitem", "bulletedlistitem",
            "blockquot", "formalarg", "formalargend"
        };

        private static string NK(string s) =>
            (s ?? "").Replace(" ", "").Replace("-", "")
                     .Replace("_", "").ToLowerInvariant();

        private static readonly string FigCaptionStyleKey = NK("figurecaption");
        private static readonly string TblCaptionStyleKey = NK("tablecaption");

        // =====================================================================
        // NON-TERMINAL ABBREVIATIONS
        //
        // Words ending in "." that must NOT be treated as sentence terminators
        // when deciding whether a following "Figure"/"Fig." citation is at the
        // start of a sentence (Rule F2). Without this, something like
        // "(cf. Figures 3 and 9)" is misread as starting a new sentence right
        // after "cf.", so the spelled-out form is wrongly left unflagged.
        // =====================================================================
        private static readonly HashSet<string> NonTerminalAbbrevs =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "cf.", "vs.", "e.g.", "i.e.", "etc.", "approx.", "cont.",
            "viz.", "resp.", "ca.", "cca."
        };

        // =====================================================================
        // FIGURE NUMBER PATTERNS
        // =====================================================================
        private const string FigNum =
            @"[A-Z]?\d+[a-z]?(?:\(\w+\))*";

        private const string FigList =
            FigNum + @"(?:\s*[,\u2013\-]\s*" + FigNum + @"|\s+and\s+" + FigNum + @")*";

        // =====================================================================
        // TABLE NUMBER PATTERNS
        //
        // Table numbers mirror figure numbers:
        //   Regular     : plain integer          e.g. 1, 12
        //   Supplemental: S-prefix + digit       e.g. S1, S12
        //   Appendix    : other letter + digit   e.g. A1, B3, C2
        //
        // TblNum covers all three forms (no trailing panel letter for tables).
        // =====================================================================
        private const string TblNum =
            @"[A-Z]?\d+";

        private const string TblList =
            TblNum + @"(?:\s*[,\u2013\-]\s*" + TblNum + @"|\s+and\s+" + TblNum + @")*";

        // =====================================================================
        // FIGURE RULES — REGEXES
        // =====================================================================

        // RULE F1 — Wrong capitalisation
        private static readonly Regex RxFigBadCaps = new Regex(
            @"(?<![A-Za-z])(fig(?:s)?\.?|figure(?:s)?)(?=\s+[A-Z]?\d)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // RULE F2 — "Figure" mid-sentence / "Fig." at sentence start
        private static readonly Regex RxFigAnyCitation = new Regex(
            @"(?<![A-Za-z])(Fig\.?|Figs\.?|Figure|Figures)\s+(" + FigList + @")",
            RegexOptions.Compiled);

        private static readonly Regex RxSentenceEnd = new Regex(
            @"[.!?][""')\]]*\s*$",
            RegexOptions.Compiled);

        // RULE F3a — Singular "Fig."/"Figure" with multiple figures
        private static readonly Regex RxFigSingularShouldBePlural = new Regex(
            @"(?<![A-Za-z])(?:Fig\.|Figure)\s+" + FigNum +
            @"(?:\s*[,\u2013\-]\s*" + FigNum + @"|\s+and\s+" + FigNum + @")+",
            RegexOptions.Compiled);

        // RULE F3b — Plural "Figs."/"Figures" with only one figure
        // Atomic group (?>...) prevents [a-z]? backtracking that caused
        // "Figures 5a and 5b" to be falsely reported as singular.
        private static readonly Regex RxFigPluralShouldBeSingular = new Regex(
            @"(?<![A-Za-z])(?:Figs\.|Figures)\s+((?>[A-Z]?\d+[a-z]?(?:\(\w+\))*))" +
            @"(?!\s*[,\u2013\-]|\s+and\s+[A-Z]?\d)",
            RegexOptions.Compiled);

        // RULE F13 — Panel / sub-image list formatting
        //
        // Handles citations that list multiple panels (or panel sub-images)
        // of one figure, e.g. "Fig. 1b and c", "Figs. 8b,e,c,f",
        // "(Figs. 10b,d and f)", "[Figs. 11a(1)-a(3), b(1)-b(3), and
        // d(1)-d(3)]". A "panel token" is a bare letter, a letter with a
        // parenthesized sub-image number, or a dash-range of those.
        private const string PanelToken =
            @"[a-z](?![a-z])(?:\(?\d+\)?)?(?:\s*[\u2013\-]\s*[a-z](?![a-z])(?:\(?\d+\)?)?)?";

        private static readonly Regex RxFigPanelList = new Regex(
            @"(?<![A-Za-z])(Fig\.?|Figs\.?|Figure|Figures)\s+(\d+)(" + PanelToken + @")" +
            @"((?:\s*,\s*" + PanelToken + @")*)" +
            @"(?:\s*,?\s*and\s+(" + PanelToken + @"))?",
            RegexOptions.Compiled);

        // Splits one panel token into letter/sub-image parts, used both to
        // decide whether an internal dash-range collapses to a comma and to
        // rebuild a corrected token.
        private static readonly Regex RxSplitPanelToken = new Regex(
            @"^([a-z])(?:\(?(\d+)\)?)?(?:\s*[\u2013\-]\s*([a-z])(?:\(?(\d+)\)?)?)?$",
            RegexOptions.Compiled);

        // RULE F5 — Missing dot: "Fig 2" instead of "Fig. 2"
        private static readonly Regex RxFigMissingDot = new Regex(
            @"(?<![A-Za-z])(Figs?)(?=\s+[A-Z]?\d)",
            RegexOptions.Compiled);

        // RULE F6 / F8 — Caption number extraction and broad citation scan
        private static readonly Regex RxFigCaptionNum = new Regex(
            @"^(?:FIG(?:URE)?|Fig(?:ure)?)\s*\.?\s*([A-Z]?\d+)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex RxFigAnyCitationBroad = new Regex(
            @"(?<![A-Za-z])(?:Fig(?:s)?\.?|Figure(?:s)?)\s+(" + FigList + @")",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex RxExtractFigBase = new Regex(
            @"([A-Z]?\d+)[a-z]?(?:\(\w+\))*",
            RegexOptions.Compiled);

        // RULE F10 — Sequential single-letter panels joined by commas
        // ("Figs. 9g,h,i") should be collapsed into an en-dash range
        // ("Figs. 9g–i"). Requires the base figure/table number once, then
        // three or more bare panel letters chained with commas (no "and").
        // Only fires when the letters are genuinely consecutive — a list
        // like "8b,e,c,f" is NOT a range and must fall through to Rule F13
        // for full expansion instead of being collapsed into a misleading
        // "8b–f".
        private static readonly Regex RxFigPanelCommaShouldBeDash = new Regex(
            @"(?<![A-Za-z])(Figs?\.|Figures?)\s+(\d+)([a-z])((?:\s*,\s*[a-z])+)(?!\w)",
            RegexOptions.Compiled);

        // RULE F11 — Missing connector between figure numbers
        //
        // Catches "Fig. S1  S2" — two figure numbers separated only by
        // whitespace, with no comma, dash, or "and" between them. This
        // typically happens when an "and" is deleted (e.g. during editing
        // or a tracked-change acceptance) and the surrounding spaces are
        // left behind. Without this rule the citation is silently parsed
        // as a single reference to the first number, and the second number
        // goes unrecognized entirely — so no singular/plural mismatch is
        // ever flagged either.
        private static readonly Regex RxFigMissingConnector = new Regex(
            @"(?<![A-Za-z])(Fig\.?|Figs\.?|Figure|Figures)\s+(" + FigNum + @")\s+(" +
            FigNum + @")(?!\.\d)(?!\s*[,\u2013\-]|\s+and\b)",
            RegexOptions.Compiled);

        // RULE F12 — Self-contained sub-image dash range
        //
        // Catches a dash-range where BOTH sides restate their own base
        // figure number, e.g. the second half of "Figs. 14a(2)-a(3) and
        // 14d(2)-d(3)" — the "14d..." portion isn't adjacent to a
        // Fig./Figs. keyword (it follows "and"), so it falls outside Rule
        // F13's list grammar and needs its own pass. Only a 2-value span
        // (consecutive sub-images, e.g. 2 and 3) is converted to a comma;
        // a 3+-value span (e.g. 1 to 3) stays a dash, since that's a real
        // range worth abbreviating.
        private static readonly Regex RxFigSubimageRangeStandalone = new Regex(
            @"(?<![A-Za-z\d])(\d+)([a-z])(?![a-z])\(?(\d+)\)?\s*[\u2013\-]\s*" +
            @"([a-z])(?![a-z])\(?(\d+)\)?(?!\d)",
            RegexOptions.Compiled);

        // RULE F15 — Missing "and" before the final item in a 3+ list of
        // distinct figure numbers. Each item here carries its OWN digit
        // (unlike Rule F13's bare-letter panel continuations), so these are
        // separate figures, not panels of one figure — e.g. "Figs. 8b, 9b,
        // 10b" should read "Figs. 8b, 9b, and 10b".
        private static readonly Regex RxFigListMissingOxfordAnd = new Regex(
            @"(?<![A-Za-z])(Figs\.|Figures)\s+(" + FigNum + @")" +
            @"((?:\s*,\s*" + FigNum + @"){2,})(?!\s*,)(?!\s+and\b)",
            RegexOptions.Compiled);

        // RULE F16 — Semicolon-separated Fig./Figs. citations should be
        // merged into one list joined by "and", e.g. "Figs. 6c,g,k,o;
        // Fig. 7c" should read "Figs. 6c,g,k,o, and 7c".
        private static readonly Regex RxFigSemicolonShouldMerge = new Regex(
            @"(?<![A-Za-z])(Figs?\.|Figures?)\s+([^;]+?)\s*;\s*(Figs?\.|Figures?)\s+(" +
            FigList + @")",
            RegexOptions.Compiled);

        // =====================================================================
        // TABLE RULES — REGEXES
        //
        // AMS style for tables:
        //   • Always spell out in full: "Table" / "Tables" — never abbreviate
        //   • Regular tables        : "Table 1",  "Tables 2 and 3"
        //   • Supplemental tables   : "Table S1", "Tables S2 and S3"
        //   • Appendix tables       : "Table A1", "Tables B2 and B3"
        //   • Types must NOT be mixed in one citation (Rule T9)
        //   • Tables defined in a caption must be cited in body text (Rule T6)
        //   • Every citation must have a matching caption (Rule T8)
        // =====================================================================

        // RULE T1 — Wrong capitalisation ("table", "TABLE", etc.)
        private static readonly Regex RxTblBadCaps = new Regex(
            @"(?<![A-Za-z])(table(?:s)?)(?=\s+[A-Z]?\d)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // RULE T2 — Abbreviation used: "Tbl.", "Tbls.", "Tab.", "Tabs." etc.
        private static readonly Regex RxTblAbbrev = new Regex(
            @"(?<![A-Za-z])(Tbl\.?|Tbls\.?|Tab\.?|Tabs?\.?)(?=\s+[A-Z]?\d)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // RULE T3a — Singular "Table" with multiple table numbers
        private static readonly Regex RxTblSingularShouldBePlural = new Regex(
            @"(?<![A-Za-z])Table\s+" + TblNum +
            @"(?:\s*[,\u2013\-]\s*" + TblNum + @"|\s+and\s+" + TblNum + @")+",
            RegexOptions.Compiled);

        // RULE T3b — Plural "Tables" with only one table number
        // Atomic group prevents backtracking (mirrors F3b fix).
        private static readonly Regex RxTblPluralShouldBeSingular = new Regex(
            @"(?<![A-Za-z])Tables\s+((?>[A-Z]?\d+))(?!\s*[,\u2013\-]|\s+and\s+[A-Z]?\d)",
            RegexOptions.Compiled);

        // RULE T10 — Missing connector between table numbers
        // Mirrors Rule F11 for tables: catches "Table S1  S2" where the
        // connector word ("and") was lost, leaving only whitespace between
        // two table numbers.
        private static readonly Regex RxTblMissingConnector = new Regex(
            @"(?<![A-Za-z])(Table|Tables)\s+(" + TblNum + @")\s+(" +
            TblNum + @")(?!\.\d)(?!\s*[,\u2013\-]|\s+and\b)",
            RegexOptions.Compiled);

        // Broad citation scan for coverage and orphan checks (T6, T8, T9).
        private static readonly Regex RxTblAnyCitationBroad = new Regex(
            @"(?<![A-Za-z])(?:Tables?|Tbls?\.?|Tabs?\.?)\s+(" + TblList + @")\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // RULE T6 / T8 — Caption number extraction.
        private static readonly Regex RxTblCaptionNum = new Regex(
            @"^(?:TABLE|Table)\s*\.?\s*([A-Z]?\d+)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // Extracts base table numbers from a table-list string.
        // Stricter than the figure equivalent: requires the number to be
        // standalone so decimal values inside table cells are not picked up.
        private static readonly Regex RxExtractTblBase = new Regex(
            @"(?<![A-Za-z\d])([A-Z]?\d+)(?!\d)", RegexOptions.Compiled);

        // =====================================================================
        // RUN
        // =====================================================================
        public void Run()
        {
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch { return; }

            int totalParas = 0;
            try { totalParas = doc.Paragraphs.Count; } catch { }

            var scopedParas = new List<(string Text, int WordStart, string StyleKey)>();
            var figCaptionParas = new List<(string Text, int WordStart)>();
            var tblCaptionParas = new List<(string Text, int WordStart)>();
            var smallCapsIssues = new List<(string Msg, string Snippet, int Pos)>();
            var captionedFigNums = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var captionedTblNums = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            bool screenWasOn = true;
            try
            {
                screenWasOn = doc.Application.ScreenUpdating;
                doc.Application.ScreenUpdating = false;
                int paraIndex = 0;

                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    paraIndex++;
                    if (paraIndex % ChunkSize == 0)
                    {
                        TaskPaneWinForms.SetProgress(totalParas > 0
                            ? $"Scanning paragraph {paraIndex} of {totalParas}…"
                            : $"Scanning paragraph {paraIndex}…");
                        doc.Application.ScreenUpdating = screenWasOn;
                        Thread.Sleep(SleepMs);
                        doc.Application.ScreenUpdating = false;
                    }

                    Word.Range rng;
                    try { rng = para.Range; } catch { continue; }

                    string txt = "";
                    int wordStart = 0;
                    try { txt = rng.Text ?? ""; wordStart = rng.Start; } catch { continue; }

                    string sk = "";
                    try { sk = NK(para.get_Style().NameLocal); } catch { continue; }

                    string trimmed = txt.TrimEnd('\r', '\n');
                    if (string.IsNullOrWhiteSpace(trimmed)) continue;
                    if (!ScopedStyles.Contains(sk)) continue;

                    scopedParas.Add((trimmed, wordStart, sk));

                    if (sk == FigCaptionStyleKey)
                    {
                        figCaptionParas.Add((trimmed, wordStart));
                        // Rule F7: small-caps check must run on the COM thread.
                        CollectSmallCapsIssue(doc, trimmed, wordStart, smallCapsIssues);
                        // Rule F8: record captioned figure numbers.
                        var capM = RxFigCaptionNum.Match(trimmed);
                        if (capM.Success) captionedFigNums.Add(capM.Groups[1].Value);
                    }

                    if (sk == TblCaptionStyleKey)
                    {
                        tblCaptionParas.Add((trimmed, wordStart));
                        // Rule T8: record captioned table numbers.
                        var capM = RxTblCaptionNum.Match(trimmed);
                        if (capM.Success) captionedTblNums.Add(capM.Groups[1].Value);
                    }
                }
            }
            finally { try { doc.Application.ScreenUpdating = screenWasOn; } catch { } }

            TaskPaneWinForms.SetProgress("Analysing figure and table citations…");

            if (scopedParas.Count == 0)
            {
                TaskPaneWinForms.AddMessage("FIGTABLE", "WARNING",
                    "No scoped paragraphs found. Make sure the document uses AMS styles.");
                return;
            }

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    int found = 0;

                    // Per-paragraph style checks (Rules F1–F5, F9, F10, F11, T1–T3, T9, T10).
                    foreach (var (text, wordStart, sk) in scopedParas)
                        found += CheckParagraph(text, wordStart, sk);

                    // Rule F6: every captioned figure must be cited in the body.
                    found += CheckFigureCaptionCoverage(scopedParas, figCaptionParas);

                    // Rule F8: every cited figure must have a matching caption.
                    found += CheckOrphanFigCitations(scopedParas, captionedFigNums);

                    // Rule T6: every captioned table must be cited in the body.
                    found += CheckTableCaptionCoverage(scopedParas, tblCaptionParas);

                    // Rule T8: every cited table must have a matching caption.
                    found += CheckOrphanTblCitations(scopedParas, captionedTblNums);

                    // Rule F7: small-caps issues collected on the COM thread.
                    foreach (var (msg, snippet, pos) in smallCapsIssues)
                    {
                        TaskPaneWinForms.AddMessage("FIGTABLE", "WARNING", msg, snippet, pos);
                        found++;
                    }

                    if (found == 0)
                        TaskPaneWinForms.AddMessage("FIGTABLE", "INFO",
                            "Figure and table citation check passed — no issues found.");
                }
                catch (Exception ex)
                {
                    TaskPaneWinForms.AddMessage("FIGTABLE", "ERROR",
                        "Figure/table citation checker error: " + ex.Message);
                }
            });
        }

        // =====================================================================
        // RULE F7 — Small-caps check (runs on the COM thread)
        //
        // Only checks the LOWERCASE letters in the prefix (e.g. "ig" in "Fig").
        // The leading uppercase "F" often has no explicit SmallCaps flag set
        // even when the caption displays correctly, so including it would cause
        // false "partially applied" positives.
        // =====================================================================
        private static void CollectSmallCapsIssue(
            Word.Document doc,
            string text,
            int wordStart,
            List<(string Msg, string Snippet, int Pos)> issues)
        {
            int letterEnd = 0;
            while (letterEnd < text.Length && char.IsLetter(text[letterEnd]))
                letterEnd++;

            if (letterEnd < 2 || letterEnd > 10) return;

            int firstLower = -1;
            for (int i = 0; i < letterEnd; i++)
                if (char.IsLower(text[i])) { firstLower = i; break; }
            if (firstLower < 0) return;

            try
            {
                Word.Range lowerRange = doc.Range(wordStart + firstLower,
                                                  wordStart + letterEnd);
                int sc = (int)lowerRange.Font.SmallCaps;
                string prefix = text.Substring(0, letterEnd);
                string snippet = text.Length > 60 ? text.Substring(0, 60) + "…" : text;

                if (sc == 0)
                    issues.Add((
                        $"Figure caption label \"{prefix}.\" is missing small caps " +
                        $"formatting — select \"{prefix}.\" and enable " +
                        $"Format → Font → Small caps so it displays as \"FIG.\"",
                        snippet, wordStart));
                else if (sc == -1)
                    issues.Add((
                        $"Figure caption label \"{prefix}.\" has small caps partially " +
                        $"applied — select \"{prefix}.\" and re-apply " +
                        $"Format → Font → Small caps uniformly.",
                        snippet, wordStart));
            }
            catch { /* COM failure — skip silently */ }
        }

        // =====================================================================
        // RULE F6 — Figure caption-coverage check
        //
        // Every FigureCaption paragraph must be cited at least once in a
        // non-caption scoped paragraph.  Supports direct citations, comma/dash
        // lists, "and" connectors, and range citations ("Figures 2–4" covers 3).
        // =====================================================================
        private int CheckFigureCaptionCoverage(
            List<(string Text, int WordStart, string StyleKey)> allScoped,
            List<(string Text, int WordStart)> captions)
        {
            int found = 0;
            foreach (var (captionText, captionWordStart) in captions)
            {
                var capMatch = RxFigCaptionNum.Match(captionText);
                if (!capMatch.Success) continue;

                string figNum = capMatch.Groups[1].Value;
                bool isCited = false;

                foreach (var (paraText, paraStart, paraStyle) in allScoped)
                {
                    if (paraStart == captionWordStart) continue;
                    if (paraStyle == FigCaptionStyleKey || paraStyle == TblCaptionStyleKey)
                        continue;

                    foreach (Match cm in RxFigAnyCitationBroad.Matches(paraText))
                        if (IsFigNumCitedInList(figNum, cm.Groups[1].Value))
                        { isCited = true; break; }

                    if (isCited) break;
                }

                if (!isCited)
                {
                    string snippet = captionText.Length > 60
                        ? captionText.Substring(0, 60) + "…" : captionText;
                    TaskPaneWinForms.AddMessage("FIGTABLE", "WARNING",
                        $"Figure {figNum} is defined in a caption but is never cited in the " +
                        $"body text — add a citation (e.g. \"Fig. {figNum}\").",
                        snippet, captionWordStart);
                    found++;
                }
            }
            return found;
        }

        // =====================================================================
        // RULE T6 — Table caption-coverage check
        //
        // Every TableCaption paragraph must be cited at least once in a
        // non-caption scoped paragraph.  Handles supplemental (S-prefix),
        // appendix (letter-prefix), and regular (no prefix) tables.
        // =====================================================================
        private int CheckTableCaptionCoverage(
            List<(string Text, int WordStart, string StyleKey)> allScoped,
            List<(string Text, int WordStart)> captions)
        {
            int found = 0;
            foreach (var (captionText, captionWordStart) in captions)
            {
                var capMatch = RxTblCaptionNum.Match(captionText);
                if (!capMatch.Success) continue;

                string tblNum = capMatch.Groups[1].Value;
                bool isCited = false;

                foreach (var (paraText, paraStart, paraStyle) in allScoped)
                {
                    if (paraStart == captionWordStart) continue;
                    if (paraStyle == FigCaptionStyleKey || paraStyle == TblCaptionStyleKey)
                        continue;

                    foreach (Match cm in RxTblAnyCitationBroad.Matches(paraText))
                        if (IsTblNumCitedInList(tblNum, cm.Groups[1].Value))
                        { isCited = true; break; }

                    if (isCited) break;
                }

                if (!isCited)
                {
                    string snippet = captionText.Length > 60
                        ? captionText.Substring(0, 60) + "…" : captionText;
                    TaskPaneWinForms.AddMessage("FIGTABLE", "WARNING",
                        $"Table {tblNum} is defined in a caption but is never cited in the " +
                        $"body text — add a citation (e.g. \"Table {tblNum}\").",
                        snippet, captionWordStart);
                    found++;
                }
            }
            return found;
        }

        // =====================================================================
        // RULE F8 — Orphan figure citations
        //
        // A figure number cited in body text with no matching FigureCaption.
        // Supplemental figures (S-prefix) are skipped — they have captions in
        // the supplement, not in the main document.
        // Each orphan is reported only once (deduplicated by base number).
        // =====================================================================
        private int CheckOrphanFigCitations(
            List<(string Text, int WordStart, string StyleKey)> allScoped,
            HashSet<string> captionedFigNums)
        {
            int found = 0;
            var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var (text, wordStart, sk) in allScoped)
            {
                // Only scan body paragraphs — not captions themselves.
                if (sk == FigCaptionStyleKey || sk == TblCaptionStyleKey) continue;

                foreach (Match m in RxFigAnyCitationBroad.Matches(text))
                {
                    foreach (Match nm in RxExtractFigBase.Matches(m.Groups[1].Value))
                    {
                        string baseNum = nm.Groups[1].Value;

                        // Skip supplemental figures — no main-document caption expected.
                        if (FigCategory(baseNum) == "supplemental") continue;

                        if (captionedFigNums.Contains(baseNum)) continue;
                        if (!reported.Add(baseNum)) continue;

                        TaskPaneWinForms.AddMessage("FIGTABLE", "WARNING",
                            $"Figure {baseNum} is cited in the text but has no matching " +
                            $"FigureCaption paragraph — add a caption or check the number.",
                            Excerpt(text, m.Index, m.Length),
                            wordStart + m.Index);
                        found++;
                    }
                }
            }
            return found;
        }

        // =====================================================================
        // RULE T8 — Orphan table citations
        //
        // A table number cited in body text with no matching TableCaption.
        // Supplemental tables (S-prefix) are skipped for the same reason as
        // supplemental figures.  Each orphan is reported only once.
        // =====================================================================
        private int CheckOrphanTblCitations(
            List<(string Text, int WordStart, string StyleKey)> allScoped,
            HashSet<string> captionedTblNums)
        {
            int found = 0;
            var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var (text, wordStart, sk) in allScoped)
            {
                if (sk == FigCaptionStyleKey || sk == TblCaptionStyleKey) continue;

                foreach (Match m in RxTblAnyCitationBroad.Matches(text))
                {
                    foreach (Match nm in RxExtractTblBase.Matches(m.Groups[1].Value))
                    {
                        string baseNum = nm.Groups[1].Value;

                        // Skip supplemental tables.
                        if (TblCategory(baseNum) == "supplemental") continue;

                        if (captionedTblNums.Contains(baseNum)) continue;
                        if (!reported.Add(baseNum)) continue;

                        string cat = TblCategory(baseNum);
                        string typeLabel = cat == "supplemental" ? " supplemental"
                                         : cat.StartsWith("appendix") ? " appendix"
                                         : "";

                        TaskPaneWinForms.AddMessage("FIGTABLE", "WARNING",
                            $"Table {baseNum} is cited in the text but has no matching " +
                            $"TableCaption paragraph — add a{typeLabel} caption or check the number.",
                            Excerpt(text, m.Index, m.Length),
                            wordStart + m.Index);
                        found++;
                    }
                }
            }
            return found;
        }

        // =====================================================================
        // HELPERS — figure list / table list membership tests
        // =====================================================================

        /// <summary>
        /// Returns true if figNum appears in figListStr either explicitly
        /// ("7" in "6 and 7") or as part of a range ("3" in "2–4").
        /// </summary>
        private static bool IsFigNumCitedInList(string figNum, string figListStr)
        {
            var direct = new Regex(
                @"(?<![A-Za-z\d])" + Regex.Escape(figNum) + @"(?![A-Z\d])");
            if (direct.IsMatch(figListStr)) return true;

            var rangePat = new Regex(@"([A-Z]?\d+)\s*[\u2013\-]\s*([A-Z]?\d+)");
            foreach (Match rm in rangePat.Matches(figListStr))
            {
                string pfxS = LetterPrefix(rm.Groups[1].Value);
                string pfxE = LetterPrefix(rm.Groups[2].Value);
                string pfxT = LetterPrefix(figNum);
                if (!string.Equals(pfxS, pfxT, StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.Equals(pfxE, pfxT, StringComparison.OrdinalIgnoreCase)) continue;

                if (int.TryParse(DigitPart(rm.Groups[1].Value), out int nS) &&
                    int.TryParse(DigitPart(rm.Groups[2].Value), out int nE) &&
                    int.TryParse(DigitPart(figNum), out int nT))
                {
                    int lo = Math.Min(nS, nE), hi = Math.Max(nS, nE);
                    if (nT >= lo && nT <= hi) return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Returns true if tblNum appears in tblListStr either explicitly or
        /// as part of a range, with prefix matching enforced.
        /// </summary>
        private static bool IsTblNumCitedInList(string tblNum, string tblListStr)
        {
            var direct = new Regex(
                @"(?<![A-Za-z\d])" + Regex.Escape(tblNum) + @"(?!\d)");
            if (direct.IsMatch(tblListStr)) return true;

            var rangePat = new Regex(@"([A-Z]?\d+)\s*[\u2013\-]\s*([A-Z]?\d+)");
            foreach (Match rm in rangePat.Matches(tblListStr))
            {
                string pfxS = LetterPrefix(rm.Groups[1].Value);
                string pfxE = LetterPrefix(rm.Groups[2].Value);
                string pfxT = LetterPrefix(tblNum);
                if (!string.Equals(pfxS, pfxT, StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.Equals(pfxE, pfxT, StringComparison.OrdinalIgnoreCase)) continue;

                if (int.TryParse(DigitPart(rm.Groups[1].Value), out int nS) &&
                    int.TryParse(DigitPart(rm.Groups[2].Value), out int nE) &&
                    int.TryParse(DigitPart(tblNum), out int nT))
                {
                    int lo = Math.Min(nS, nE), hi = Math.Max(nS, nE);
                    if (nT >= lo && nT <= hi) return true;
                }
            }
            return false;
        }

        private static string LetterPrefix(string s)
        {
            int i = 0;
            while (i < s.Length && char.IsLetter(s[i])) i++;
            return s.Substring(0, i);
        }

        private static string DigitPart(string s)
        {
            var sb = new System.Text.StringBuilder();
            bool started = false;
            foreach (char c in s)
            {
                if (char.IsDigit(c)) { sb.Append(c); started = true; }
                else if (started) break;
            }
            return sb.ToString();
        }

        // =====================================================================
        // FigCategory — classifies a figure base number by type
        //   "regular"      — no leading letter: 1, 12
        //   "supplemental" — S/s prefix:        S1, S12
        //   "appendix:X"   — other letter:      A1, B3  (X = upper-cased)
        // =====================================================================
        private static string FigCategory(string figBase)
        {
            if (string.IsNullOrEmpty(figBase)) return "regular";
            char first = figBase[0];
            if (char.IsDigit(first)) return "regular";
            if ((first == 'S' || first == 's') && figBase.Length > 1 && char.IsDigit(figBase[1]))
                return "supplemental";
            return "appendix:" + char.ToUpperInvariant(first);
        }

        // =====================================================================
        // TblCategory — mirrors FigCategory for table numbers
        // =====================================================================
        private static string TblCategory(string tblBase)
        {
            if (string.IsNullOrEmpty(tblBase)) return "regular";
            char first = tblBase[0];
            if (char.IsDigit(first)) return "regular";
            if ((first == 'S' || first == 's') && tblBase.Length > 1 && char.IsDigit(tblBase[1]))
                return "supplemental";
            return "appendix:" + char.ToUpperInvariant(first);
        }

        // =====================================================================
        // CHECK ONE PARAGRAPH — Figure rules F1–F5, F9–F13, F15, F16  +  Table rules T1–T3, T9, T10
        // =====================================================================
        private int CheckParagraph(string text, int wordStart, string styleKey)
        {
            bool isFigureCaption = (styleKey == FigCaptionStyleKey);
            int found = 0;

            // =================================================================
            // FIGURE RULES
            // =================================================================

            // RULE F5 — Missing dot ("Fig 2")
            // Run first; record ranges so Rule F2 cannot double-fire.
            var missingDotRanges = new List<(int Start, int End)>();
            foreach (Match m in RxFigMissingDot.Matches(text))
            {
                string word = m.Groups[1].Value;
                Add("WARNING",
                    $"Missing dot in figure citation: \"{word}\" → \"{word}.\".",
                    text, m, wordStart);
                found++;
                missingDotRanges.Add((m.Index, m.Index + m.Length));
            }

            // RULE F1 — Wrong capitalisation
            foreach (Match m in RxFigBadCaps.Matches(text))
            {
                string raw = m.Groups[1].Value;
                if (raw == "Fig." || raw == "Figs." ||
                    raw == "Figure" || raw == "Figures" ||
                    raw == "Fig" || raw == "Figs") continue;

                string rawLower = raw.ToLowerInvariant();
                bool hasPlural = rawLower.StartsWith("figs");
                bool isSpelled = rawLower == "figure" || rawLower == "figures";
                string sug = isSpelled
                    ? (hasPlural ? "Figures" : "Figure")
                    : (hasPlural ? "Figs." : "Fig.");

                Add("WARNING", $"Wrong capitalisation \"{raw}\" — use \"{sug}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE F2 — "Figure" mid-sentence / "Fig." at sentence start
            // Caption label at position 0 is exempt (its caps are set by
            // the small-caps character format checked in Rule F7).
            foreach (Match m in RxFigAnyCitation.Matches(text))
            {
                if (isFigureCaption && m.Index == 0) continue;
                if (IsInRanges(m.Index, missingDotRanges)) continue;

                string word = m.Groups[1].Value;
                bool isSpelledOut = word == "Figure" || word == "Figures";
                bool isAbbrev = word == "Fig." || word == "Figs.";
                if (!isSpelledOut && !isAbbrev) continue;

                bool isSentenceStart = IsSentenceStart(text, m.Index);
                if (isSentenceStart && isAbbrev)
                {
                    string plural = word == "Figs." ? "s" : "";
                    Add("WARNING",
                        $"\"{word}\" at sentence start — spell out as \"Figure{plural}\".",
                        text, m, wordStart);
                    found++;
                }
                else if (!isSentenceStart && isSpelledOut)
                {
                    string plural = word == "Figures" ? "s" : "";
                    Add("WARNING",
                        $"\"{word}\" mid-sentence — abbreviate as \"Fig{plural}.\".",
                        text, m, wordStart);
                    found++;
                }
            }

            // RULE F3a — Singular with multiple figures
            foreach (Match m in RxFigSingularShouldBePlural.Matches(text))
            {
                Add("WARNING",
                    $"Multiple figures cited with singular — use \"Figs.\" or \"Figures\": " +
                    $"\"{Truncate(m.Value, 50)}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE F3b — Plural with only one figure
            foreach (Match m in RxFigPluralShouldBeSingular.Matches(text))
            {
                Add("WARNING",
                    $"Single figure cited with plural — use \"Fig.\" or \"Figure\": " +
                    $"\"{Truncate(m.Value, 50)}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE F13 — Panel / sub-image list formatting.
            //
            // Tracks every span it considers so Rule F12 (below) can skip
            // re-examining a dash-range that's already part of one of these
            // lists (e.g. the first group in "14a(2)-a(3) and 14d(2)-d(3)").
            var panelListRanges = new List<(int Start, int End)>();
            foreach (Match m in RxFigPanelList.Matches(text))
            {
                panelListRanges.Add((m.Index, m.Index + m.Length));

                string prefix = m.Groups[1].Value;
                string baseNum = m.Groups[2].Value;
                string firstTok = m.Groups[3].Value;
                string midRaw = m.Groups[4].Value;
                string andTok = m.Groups[5].Success ? m.Groups[5].Value : null;

                var midTokens = new List<string>();
                foreach (Match tm in Regex.Matches(midRaw, PanelToken))
                    midTokens.Add(tm.Value);

                var allTokens = new List<string> { firstTok };
                allTokens.AddRange(midTokens);
                if (!string.IsNullOrEmpty(andTok)) allTokens.Add(andTok);

                string suggestion = null;

                if (allTokens.Count < 2)
                {
                    // Single citation — only fix it if it's itself a
                    // 2-value dash-range (Rule F12's job otherwise).
                    var dm = RxSplitPanelToken.Match(firstTok);
                    if (dm.Success && !string.IsNullOrEmpty(dm.Groups[3].Value) &&
                        int.TryParse(dm.Groups[2].Value, out int fn1) &&
                        int.TryParse(dm.Groups[4].Value, out int fn2) &&
                        Math.Abs(fn2 - fn1) == 1)
                    {
                        suggestion = $"{prefix} {baseNum}{FixPanelToken(firstTok)}";
                    }
                }
                else
                {
                    bool allBare = allTokens.TrueForAll(IsBarePanelLetter);
                    bool isConsecutiveBareList = andTok == null && allBare &&
                                                  AreConsecutiveLetters(allTokens);
                    if (!isConsecutiveBareList)
                    {
                        var fixedTokens = allTokens.ConvertAll(FixPanelToken);
                        bool inside = IsInsideEnclosure(text, m.Index);
                        bool isSpelled = prefix.StartsWith("Figure", StringComparison.OrdinalIgnoreCase);
                        string pluralPrefix = isSpelled ? "Figures" : "Figs.";

                        if (inside)
                        {
                            var tail = fixedTokens.GetRange(1, fixedTokens.Count - 1)
                                                  .ConvertAll(f => "," + f);
                            suggestion = pluralPrefix + " " + baseNum + fixedTokens[0] +
                                         string.Concat(tail);
                        }
                        else
                        {
                            var full = fixedTokens.ConvertAll(f => baseNum + f);
                            suggestion = full.Count == 1
                                ? pluralPrefix + " " + full[0]
                                : pluralPrefix + " " +
                                  string.Join(", ", full.GetRange(0, full.Count - 1)) +
                                  " and " + full[full.Count - 1];
                        }
                    }
                    // else: leave to Rule F10's dash-collapse, below.
                }

                if (suggestion != null && suggestion != m.Value)
                {
                    Add("WARNING",
                        $"Panel/sub-image citation needs reformatting — " +
                        $"\"{Truncate(m.Value, 60)}\" should be \"{suggestion}\".",
                        text, m, wordStart);
                    found++;
                }
            }

            // RULE F10 — Sequential panel letters joined by commas should be
            // an en-dash range, e.g. "Figs. 9g,h,i" → "Figs. 9g–i". Only
            // fires when the letters are genuinely consecutive; a
            // non-sequential list like "8b,e,c,f" is left to Rule F13.
            foreach (Match m in RxFigPanelCommaShouldBeDash.Matches(text))
            {
                string prefix = m.Groups[1].Value;
                string num = m.Groups[2].Value;
                string firstLetter = m.Groups[3].Value;
                string lettersStr = m.Groups[4].Value;

                var letters = new List<string> { firstLetter };
                foreach (Match lm in Regex.Matches(lettersStr, @"[a-z]"))
                    letters.Add(lm.Value);

                // Only flag genuine sequential lists (3+ panels). Two panels
                // joined by a comma ("9g,h") are ambiguous and left alone —
                // AMS style permits "9g, h" or "9g and h" equally there.
                if (letters.Count < 3) continue;
                if (!AreConsecutiveLetters(letters)) continue;

                string lastLetter = letters[letters.Count - 1];
                string suggestion = $"{prefix} {num}{firstLetter}\u2013{lastLetter}";

                Add("WARNING",
                    $"Sequential figure panels listed with commas — use an en-dash range: " +
                    $"\"{Truncate(m.Value, 50)}\" → \"{suggestion}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE F11 — Missing connector ("and") between figure numbers.
            // Catches "Fig. S1  S2" — two figure numbers with only whitespace
            // between them, no comma/dash/and. The second number would
            // otherwise go completely unrecognized as part of the citation.
            foreach (Match m in RxFigMissingConnector.Matches(text))
            {
                string word = m.Groups[1].Value;
                string num1 = m.Groups[2].Value;
                string num2 = m.Groups[3].Value;
                bool isSpelled = word.Equals("Figure", StringComparison.OrdinalIgnoreCase) ||
                                  word.Equals("Figures", StringComparison.OrdinalIgnoreCase);
                string pluralWord = isSpelled ? "Figures" : "Figs.";

                Add("WARNING",
                    $"Missing connector between figure numbers — \"{Truncate(m.Value, 50)}\" " +
                    $"reads as two citations with no \"and\" between them; use " +
                    $"\"{pluralWord} {num1} and {num2}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE F12 — Self-contained sub-image dash range (digit repeated
            // on both sides), e.g. the second half of "Figs. 14a(2)-a(3) and
            // 14d(2)-d(3)". Skips any span already covered by Rule F13 so
            // the two rules don't double-report the same range.
            foreach (Match m in RxFigSubimageRangeStandalone.Matches(text))
            {
                if (IsInRanges(m.Index, panelListRanges)) continue;

                string baseNum = m.Groups[1].Value;
                string panel1 = m.Groups[2].Value;
                string sub1 = m.Groups[3].Value;
                string panel2 = m.Groups[4].Value;
                string sub2 = m.Groups[5].Value;

                if (!int.TryParse(sub1, out int n1) || !int.TryParse(sub2, out int n2)) continue;
                if (Math.Abs(n2 - n1) != 1) continue;

                string suggestion = $"{baseNum}{panel1}({sub1}),{panel2}({sub2})";

                Add("WARNING",
                    $"Sub-image range written with a dash — \"{Truncate(m.Value, 60)}\" " +
                    $"should list each sub-image explicitly with its number in " +
                    $"parentheses: \"{suggestion}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE F15 — Missing "and" before the final item in a 3+ list of
            // distinct figure numbers (each item carries its own digit, so
            // these are separate figures, not panels of one figure), e.g.
            // "Figs. 8b, 9b, 10b" should read "Figs. 8b, 9b, and 10b".
            foreach (Match m in RxFigListMissingOxfordAnd.Matches(text))
            {
                string prefix = m.Groups[1].Value;
                string first = m.Groups[2].Value;
                string restRaw = m.Groups[3].Value;

                var items = new List<string> { first };
                foreach (var part in restRaw.Split(','))
                {
                    string trimmed = part.Trim();
                    if (!string.IsNullOrEmpty(trimmed)) items.Add(trimmed);
                }
                if (items.Count < 3) continue;

                string suggestion = prefix + " " +
                    string.Join(", ", items.GetRange(0, items.Count - 1)) +
                    ", and " + items[items.Count - 1];

                Add("WARNING",
                    $"List of figures missing \"and\" before the last item — " +
                    $"\"{Truncate(m.Value, 60)}\" should be \"{suggestion}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE F16 — Semicolon-separated Fig./Figs. citations should be
            // merged into one list joined by "and", e.g. "Figs. 6c,g,k,o;
            // Fig. 7c" should read "Figs. 6c,g,k,o, and 7c".
            foreach (Match m in RxFigSemicolonShouldMerge.Matches(text))
            {
                string prefix1 = m.Groups[1].Value;
                string list1 = m.Groups[2].Value.Trim();
                string list2 = m.Groups[4].Value.Trim();

                string suggestion = $"{prefix1} {list1}, and {list2}";

                Add("WARNING",
                    $"Separate \"Fig.\"/\"Figs.\" citations joined by a semicolon should be " +
                    $"merged into one list — \"{Truncate(m.Value, 60)}\" should be " +
                    $"\"{suggestion}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE F9 — Mixed figure types in one citation
            // Supplemental (S-prefix), appendix (letter-prefix), and regular
            // figures must not appear together in one Fig./Figs. citation.
            foreach (Match m in RxFigAnyCitationBroad.Matches(text))
            {
                string figListStr = m.Groups[1].Value;
                bool hasReg = false, hasSupp = false, hasApp = false;
                foreach (Match nm in RxExtractFigBase.Matches(figListStr))
                {
                    switch (FigCategory(nm.Groups[1].Value))
                    {
                        case "regular": hasReg = true; break;
                        case "supplemental": hasSupp = true; break;
                        default: hasApp = true; break;
                    }
                }
                int typeCount = (hasReg ? 1 : 0) + (hasSupp ? 1 : 0) + (hasApp ? 1 : 0);
                if (typeCount < 2) continue;

                var parts = new List<string>();
                if (hasReg) parts.Add("regular (no-prefix)");
                if (hasSupp) parts.Add("supplemental (S-prefix)");
                if (hasApp) parts.Add("appendix (letter-prefix)");

                Add("WARNING",
                    $"Mixed figure types in one citation ({string.Join(" + ", parts)}): " +
                    $"\"{Truncate(m.Value, 60)}\" — supplemental (S-prefix) and appendix " +
                    $"figures must be cited separately from regular body figures.",
                    text, m, wordStart);
                found++;
            }

            // =================================================================
            // TABLE RULES
            // =================================================================

            // RULE T2 — Abbreviation used ("Tbl.", "Tab.", etc.)
            // Run first; record ranges so T1 cannot double-fire on the same span.
            var tblAbbrevRanges = new List<(int Start, int End)>();
            foreach (Match m in RxTblAbbrev.Matches(text))
            {
                string raw = m.Groups[1].Value;
                string plural = raw.ToLowerInvariant().TrimEnd('.').EndsWith("s") ? "s" : "";
                Add("WARNING",
                    $"Table abbreviation \"{raw}\" is not permitted in AMS style " +
                    $"— always spell out as \"Table{plural}\".",
                    text, m, wordStart);
                found++;
                tblAbbrevRanges.Add((m.Index, m.Index + m.Length));
            }

            // RULE T1 — Wrong capitalisation ("table 1", "TABLE 1", etc.)
            foreach (Match m in RxTblBadCaps.Matches(text))
            {
                if (IsInRanges(m.Index, tblAbbrevRanges)) continue;
                string raw = m.Groups[1].Value;
                if (raw == "Table" || raw == "Tables") continue;

                bool hasPlural = raw.ToLowerInvariant() == "tables";
                Add("WARNING",
                    $"Wrong capitalisation \"{raw}\" — use \"{(hasPlural ? "Tables" : "Table")}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE T3a — Singular "Table" with multiple table numbers
            foreach (Match m in RxTblSingularShouldBePlural.Matches(text))
            {
                Add("WARNING",
                    $"Multiple tables cited with singular — use \"Tables\": " +
                    $"\"{Truncate(m.Value, 50)}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE T3b — Plural "Tables" with only one table number
            foreach (Match m in RxTblPluralShouldBeSingular.Matches(text))
            {
                Add("WARNING",
                    $"Single table cited with plural — use \"Table\": " +
                    $"\"{Truncate(m.Value, 50)}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE T10 — Missing connector ("and") between table numbers.
            // Mirrors Rule F11 for tables.
            foreach (Match m in RxTblMissingConnector.Matches(text))
            {
                string num1 = m.Groups[2].Value;
                string num2 = m.Groups[3].Value;

                Add("WARNING",
                    $"Missing connector between table numbers — \"{Truncate(m.Value, 50)}\" " +
                    $"reads as two citations with no \"and\" between them; use " +
                    $"\"Tables {num1} and {num2}\".",
                    text, m, wordStart);
                found++;
            }

            // RULE T9 — Mixed table types in one citation
            foreach (Match m in RxTblAnyCitationBroad.Matches(text))
            {
                string tblListStr = m.Groups[1].Value;
                bool hasReg = false, hasSupp = false, hasApp = false;
                foreach (Match nm in RxExtractTblBase.Matches(tblListStr))
                {
                    switch (TblCategory(nm.Groups[1].Value))
                    {
                        case "regular": hasReg = true; break;
                        case "supplemental": hasSupp = true; break;
                        default: hasApp = true; break;
                    }
                }
                int typeCount = (hasReg ? 1 : 0) + (hasSupp ? 1 : 0) + (hasApp ? 1 : 0);
                if (typeCount < 2) continue;

                var parts = new List<string>();
                if (hasReg) parts.Add("regular (no-prefix)");
                if (hasSupp) parts.Add("supplemental (S-prefix)");
                if (hasApp) parts.Add("appendix (letter-prefix)");

                Add("WARNING",
                    $"Mixed table types in one citation ({string.Join(" + ", parts)}): " +
                    $"\"{Truncate(m.Value, 60)}\" — supplemental (S-prefix) and appendix " +
                    $"tables must be cited separately from regular body tables.",
                    text, m, wordStart);
                found++;
            }

            return found;
        }

        // =====================================================================
        // HELPERS
        // =====================================================================
        private static bool IsInRanges(int index, List<(int Start, int End)> ranges)
        {
            foreach (var (s, e) in ranges)
                if (index >= s && index < e) return true;
            return false;
        }

        // =====================================================================
        // IsInsideEnclosure
        //
        // Used by Rule F13 to decide formatting: panel-list style differs
        // depending on whether the citation sits inside an open, not-yet-
        // closed "(" or "[" — i.e. it's itself a parenthetical citation —
        // versus appearing as plain running text. Computed as a simple
        // bracket balance over the paragraph text up to the match start.
        // =====================================================================
        private static bool IsInsideEnclosure(string text, int index)
        {
            int balance = 0;
            int limit = Math.Min(index, text.Length);
            for (int i = 0; i < limit; i++)
            {
                char c = text[i];
                if (c == '(' || c == '[') balance++;
                else if (c == ')' || c == ']') balance--;
            }
            return balance > 0;
        }

        // =====================================================================
        // FixPanelToken
        //
        // Normalizes one panel token for Rule F13/F12: a bare letter is
        // returned unchanged; a letter+sub-image number gets parentheses
        // added if missing; a dash-range collapses to a comma only when it
        // spans exactly two consecutive sub-images (e.g. "a(2)-a(3)" →
        // "a(2),a(3)"), and otherwise keeps its dash (e.g. "a(1)-a(3)").
        // =====================================================================
        private static string FixPanelToken(string token)
        {
            var m = RxSplitPanelToken.Match(token);
            if (!m.Success) return token;

            string l1 = m.Groups[1].Value;
            string s1 = m.Groups[2].Value;
            string l2 = m.Groups[3].Value;
            string s2 = m.Groups[4].Value;

            if (string.IsNullOrEmpty(l2))
                return string.IsNullOrEmpty(s1) ? l1 : $"{l1}({s1})";

            if (int.TryParse(s1, out int n1) && int.TryParse(s2, out int n2) &&
                Math.Abs(n2 - n1) == 1)
                return $"{l1}({s1}),{l2}({s2})";

            return $"{l1}({s1})\u2013{l2}({s2})";
        }

        private static bool IsBarePanelLetter(string token) =>
            Regex.IsMatch(token, @"^[a-z]$");

        private static bool AreConsecutiveLetters(List<string> letters)
        {
            for (int i = 0; i < letters.Count - 1; i++)
                if (letters[i][0] + 1 != letters[i + 1][0]) return false;
            return true;
        }

        // =====================================================================
        // IsSentenceStart
        //
        // A citation is "at the start of a sentence" only if it follows real
        // terminal punctuation. Abbreviations like "cf.", "e.g.", "i.e." end
        // in a period but do NOT end the sentence, so text like
        // "(cf. Figures 3 and 9)" must still be treated as mid-sentence —
        // otherwise the spelled-out "Figures" there is wrongly left
        // unflagged instead of being corrected to "Figs.".
        // =====================================================================
        private static bool IsSentenceStart(string text, int index)
        {
            if (index == 0) return true;
            int firstNonSpace = 0;
            while (firstNonSpace < text.Length && char.IsWhiteSpace(text[firstNonSpace]))
                firstNonSpace++;
            if (index == firstNonSpace) return true;

            string before = text.Substring(0, index);
            if (!RxSentenceEnd.IsMatch(before)) return false;

            // Don't treat known non-terminal abbreviations as sentence enders.
            var wm = Regex.Match(before.TrimEnd(), @"(\S+)$");
            if (wm.Success && NonTerminalAbbrevs.Contains(wm.Groups[1].Value))
                return false;

            return true;
        }

        private void Add(string sev, string msg, string text, Match m, int paraStart)
        {
            TaskPaneWinForms.AddMessage("FIGTABLE", sev, msg,
                Excerpt(text, m.Index, m.Length), paraStart + m.Index);
        }

        private static string Excerpt(string text, int index, int length)
        {
            if (string.IsNullOrEmpty(text)) return "";
            index = Math.Max(0, Math.Min(index, text.Length - 1));
            length = Math.Max(0, Math.Min(length, text.Length - index));
            int s = Math.Max(0, index - 15);
            int e = Math.Min(text.Length, index + length + 15);
            string snip = text.Substring(s, e - s).Trim();
            return snip.Length > 70 ? snip.Substring(0, 70) + "…" : snip;
        }

        private static string Truncate(string s, int max) =>
            s == null ? "" : s.Length <= max ? s : s.Substring(0, max) + "…";
    }
}