using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin.Checkers
{
    public class HeadingAnatomyChecker
    {
        // ── Issue counter (reset each Run) ────────────────────────────────────
        private int _issueNum;

        private void Msg(string cat, string level, string msg, string snippet = "", int start = 0)
        {
            _issueNum++;
            TaskPaneWinForms.AddMessage(cat, level, $"#{_issueNum} {msg}", snippet, start);
        }

        private static string NK(string s) =>
            (s ?? "").Replace(" ", "").Replace("-", "").Replace("_", "").ToLowerInvariant();

        private static readonly HashSet<string> ScopedStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "figurecaption","tablecaption","tablebody","tablehead",
            "paratext","acknowledgementtext","abstracttext","synopsis",
            "paranoindent","numberedlistitem","bulletedlistitem",
            "blockquot","formalarg","formalargend"
        };

        private static bool IsH1(ParaInfo p) => p.SK == "heading01" || p.SK == "heading1";
        private static bool IsH2(ParaInfo p) => p.SK == "heading02" || p.SK == "heading2";
        private static bool IsH3(ParaInfo p) => p.SK == "heading03" || p.SK == "heading3";
        private static bool IsH4(ParaInfo p) => p.SK == "heading04" || p.SK == "heading4";
        private static bool IsHeading(ParaInfo p) => IsH1(p) || IsH2(p) || IsH3(p) || IsH4(p);

        private static int HeadingLevel(ParaInfo p)
        {
            if (IsH1(p)) return 1; if (IsH2(p)) return 2;
            if (IsH3(p)) return 3; if (IsH4(p)) return 4;
            return 0;
        }

        private static bool IsAckTitle(ParaInfo p) =>
            p.SK == NK("acknowledgement_title") || p.SK == NK("acknowledgment_title");

        private static readonly (string Style, string Label, bool Required)[] FrontMatter =
        {
            ("right_rh","Right running head",true),("doi_number","DOI number",true),
            ("article_title","Article title",true),("authors","Authors",true),
            ("affiliation","Affiliation",true),("received","History info (received)",true),
            ("abstract_title","Abstract title",true),("abstract_text","Abstract text",true),
            ("synopsis","Synopsis / significance",true),("keyword_title","Keywords title",true),
            ("keywords","Keywords",true),("footnote","Footnote (supplementary info)",true),
            ("correspondence","Correspondence",true),("footnote_aff","Current affiliation",true),
        };

        private static readonly (string Style, string Label, bool Required)[] BackMatter =
        {
            ("acknowledgement_title","Acknowledgments title",true),
            ("acknowledgement_text","Acknowledgments text",true),
            ("appendix_title","Appendix title",true),
            ("reference_title","References title",true),
            ("figure_caption","Figure caption(s)",true),
            ("table_caption","Table caption(s)",true),
        };

        // ── Heading pattern regexes ───────────────────────────────────────────
        private static readonly Regex RxH1Pattern = new Regex(@"^\d+\.\s", RegexOptions.Compiled);
        private static readonly Regex RxH2Pattern = new Regex(@"^[a-z]\.\s", RegexOptions.Compiled);
        private static readonly Regex RxH3Pattern = new Regex(@"^\d+\)\s", RegexOptions.Compiled);
        private static readonly Regex RxH4Pattern = new Regex(@"^\([ivxlcdm]+\)\s", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxH2Label = new Regex(@"^([a-z])\.", RegexOptions.Compiled);
        private static readonly Regex RxH3Label = new Regex(@"^(\d+)\)", RegexOptions.Compiled);
        private static readonly Regex RxH4Label = new Regex(@"^\(([ivxlcdm]+)\)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxEndPunct = new Regex(@"[.:]$", RegexOptions.Compiled);
        private static readonly Regex RxAllCaps = new Regex(@"[A-Z]{3,}", RegexOptions.Compiled);

        // ── Appendix regexes ─────────────────────────────────────────────────
        private static readonly Regex RxAppendixWord = new Regex(@"\bappendix\b|\bappendices\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        // Case-insensitive so titles/citations styled as "APPENDIX A.", "Appendix A.",
        // or "appendix a." are all recognised the same way regardless of casing.
        private static readonly Regex RxAppendixLetter = new Regex(@"\bAppendix\s+([A-Za-z])\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // ── Equation regexes ─────────────────────────────────────────────────
        // Two forms are supported for every numbering scheme:
        //   "Paren" — the classic form, numbering wrapped in parentheses,
        //             e.g. "(1)", "(1a)", "(2b)", "(1.2)", "(A.1)", "(A1a)"
        //   "Bare"  — no parentheses, but a mandatory trailing period so the
        //             number isn't confused with ordinary digits inside the
        //             equation text, e.g. "03.", "1.2.", "A1.", "A2 .", "B1."
        // ParseEquationNumber() below tries Appendix → Sectional → Simple,
        // Paren before Bare. Anything matching none of them is flagged via
        // the EqKind.Unknown fail-safe rather than silently skipped.
        private static readonly Regex RxEqApxParen = new Regex(
            @"\(\s*([A-Za-z])\.?\s*(\d+)\s*([a-z])?\s*\)\s*\.?\s*$", RegexOptions.Compiled);
        private static readonly Regex RxEqApxBare = new Regex(
            @"(?<![A-Za-z0-9])([A-Z])\.?\s*(\d+)\s*([a-z])?\s*\.\s*$", RegexOptions.Compiled);

        private static readonly Regex RxEqSectParen = new Regex(
            @"\(\s*(\d+)\.(\d+)\s*([a-z])?\s*\)\s*\.?\s*$", RegexOptions.Compiled);
        private static readonly Regex RxEqSectBare = new Regex(
            @"(?<![A-Za-z0-9])(\d+)\.(\d+)\s*([a-z])?\s*\.\s*$", RegexOptions.Compiled);

        private static readonly Regex RxEqSimpleParen = new Regex(
            @"\(\s*(\d+)\s*([a-z])?\s*\)\s*\.?\s*$", RegexOptions.Compiled);
        private static readonly Regex RxEqSimpleBare = new Regex(
            @"(?<![A-Za-z0-9.])(\d+)\s*([a-z])?\s*\.\s*$", RegexOptions.Compiled);

        private static readonly Regex RxTEQuery = new Regex(@"<+\?TE:.*?>>", RegexOptions.Compiled | RegexOptions.Singleline);

        // ── Section citation regexes ─────────────────────────────────────────
        private static readonly Regex RxSectionCite = new Regex(
            @"\bsections?\s+(?:[A-Z]?\d+[a-z]?(?:\(\d+\))?(?:\([ivxlcdm]+\))?)" +
            @"(?:\s*[-\u2013\u2014]\s*(?:[A-Z]?\d+[a-z]?(?:\(\d+\))?(?:\([ivxlcdm]+\))?))?" +
            @"(?:\s+and\s+(?:[A-Z]?\d+[a-z]?(?:\(\d+\))?(?:\([ivxlcdm]+\))?))*",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxSectionToken = new Regex(
            @"[A-Z]?\d+[a-z]?(?:\(\d+\))?(?:\([ivxlcdm]+\))?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxSectionRange = new Regex(
            @"^([A-Z]?\d+[a-z]?(?:\(\d+\))?(?:\([ivxlcdm]+\))?)\s*[-\u2013\u2014]\s*([A-Z]?\d+[a-z]?(?:\(\d+\))?(?:\([ivxlcdm]+\))?)$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // ── Section citation — NEW rules regexes ─────────────────────────────

        // § symbol used instead of spelling out "section"
        private static readonly Regex RxSectionSymbol = new Regex(
            @"§\s*\d+", RegexOptions.Compiled);

        // "subsection" followed by a label (only flag when a label is present)
        private static readonly Regex RxSubsectionLabelled = new Regex(
            @"\bsubsection\s+[A-Z]?\d+[a-z]?(?:\(\d+\))?(?:\([ivxlcdm]+\))?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // H4 cited with NO parentheses at all — e.g. "4a2ii"
        private static readonly Regex RxH4CiteNoParens = new Regex(
            @"\b([A-Z]?\d+[a-z])(\d+)([ivxlcdm]+)\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // H4 cited with first parens but missing second — e.g. "4a(2)ii"
        private static readonly Regex RxH4CiteMissingSecondParens = new Regex(
            @"\b([A-Z]?\d+[a-z]?)\((\d+)\)([ivxlcdm]+)\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // H4 cited with second parens but missing first — e.g. "4a2(ii)"
        private static readonly Regex RxH4CiteMissingFirstParens = new Regex(
            @"\b([A-Z]?\d+[a-z]?)(\d+)\(([ivxlcdm]+)\)\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // ── History info regexes ─────────────────────────────────────────────
        // Flags zero-padded month/day numbers like 01., 02. … 09. inside history info paragraphs
        private static readonly Regex RxHistoryZeroPadded = new Regex(
            @"\b0[1-9]\.", RegexOptions.Compiled);

        // ── Appendix citation regexes ─────────────────────────────────────────
        // Matches "Appendix A", "APPENDIX B", "appendix a", etc. in body text (case-insensitive)
        private static readonly Regex RxAppendixCiteWithLetter = new Regex(
            @"\bAppendix\s+([A-Za-z])\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        // Matches bare "Appendix" (no letter) in body text, any case
        private static readonly Regex RxAppendixCiteBare = new Regex(
            @"\bAppendix\b(?!\s+[A-Za-z])", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // ── Known-good section token — Layer 1 strict match ──────────────────
        // Valid forms: 3 | 3a | 3a(1) | 3a(1)(ii) | A3 | A3a | A3a(1)(ii)
        // Anything that doesn't match this is flagged for manual review.
        private static readonly Regex RxKnownGoodToken = new Regex(
            @"^[A-Z]?\d+[a-z]?(?:\(\d+\))?(?:\([ivxlcdm]+\))?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // ── Author parsing regex ─────────────────────────────────────────────
        private static readonly Regex RxAuthorToken = new Regex(
            @"(?<![A-Za-z])([A-Z]\.(?:\s+[A-Z]\.)*)\s+([A-Z][a-zA-Z\u00C0-\u024F\-]{1,30})",
            RegexOptions.Compiled);

        // =====================================================================
        // RUN
        // =====================================================================
        public void Run()
        {
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; } catch { return; }

            _issueNum = 0;

            bool screenWasOn = true;
            var allParas = new List<ParaInfo>();
            try
            {
                screenWasOn = doc.Application.ScreenUpdating;
                doc.Application.ScreenUpdating = false;
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    string sn = "", sk = "";
                    try { sn = para.get_Style().NameLocal; sk = NK(sn); } catch { continue; }
                    string text = "";
                    try { text = para.Range.Text.Trim(); } catch { continue; }
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    int start = 0;
                    try { start = para.Range.Start; } catch { }
                    allParas.Add(new ParaInfo(sk, sn, text, start));
                }
            }
            finally { try { doc.Application.ScreenUpdating = screenWasOn; } catch { } }

            int firstH1Idx = allParas.FindIndex(IsH1);
            int ackIdx = allParas.FindIndex(IsAckTitle);

            int found = 0;
            found += CheckFrontMatter(allParas, firstH1Idx);
            found += CheckBackMatter(allParas, ackIdx);
            found += CheckBackMatterRelativeOrder(allParas);
            found += CheckHeadingRules(doc, allParas, ackIdx);
            found += CheckHeadingSequence(allParas);
            found += CheckHeadingLabelSequence(allParas);
            found += CheckDisplayEquationFollowUp(allParas);
            found += CheckSpecialTitleStyles(allParas);
            found += CheckNonHeadingStyles(allParas);
            found += CheckAppendixReferences(allParas);
            found += CheckAppendixCitations(allParas);
            found += CheckSectionCitations(allParas);
            found += CheckCurrentAffiliation(allParas);
            found += CheckDisplayEquationNumbering(allParas);
            found += CheckHistoryInfoContent(allParas);

            if (found == 0)
                TaskPaneWinForms.AddMessage("HEADING", "INFO",
                    "Heading anatomy check passed — article structure and all headings follow AMS rules.");
        }

        // =====================================================================
        // CHECK — SECTION CITATIONS
        // =====================================================================
        private int CheckSectionCitations(List<ParaInfo> paras)
        {
            int found = 0;
            var h1Numbers = new HashSet<int>();
            var h2Letters = new HashSet<char>();
            var h3Numbers = new HashSet<int>();
            var h4Romans = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var p in paras)
            {
                if (IsH1(p)) { var m = RxH1Pattern.Match(p.Text); if (m.Success) { string ns = p.Text.Substring(0, p.Text.IndexOf('.')).Trim(); if (int.TryParse(ns, out int n)) h1Numbers.Add(n); } }
                else if (IsH2(p)) { var m = RxH2Label.Match(p.Text); if (m.Success) h2Letters.Add(m.Groups[1].Value[0]); }
                else if (IsH3(p)) { var m = RxH3Label.Match(p.Text); if (m.Success && int.TryParse(m.Groups[1].Value, out int n)) h3Numbers.Add(n); }
                else if (IsH4(p)) { var m = RxH4Label.Match(p.Text); if (m.Success) h4Romans.Add(m.Groups[1].Value.ToLowerInvariant()); }
            }

            foreach (var p in paras)
            {
                if (!ScopedStyles.Contains(p.SK)) continue;
                string text = p.Text;

                // ── NEW Rule 1: § symbol must not be used ─────────────────────
                if (RxSectionSymbol.IsMatch(text))
                {
                    Msg("HEADING", "ERROR",
                        "Do not use the § symbol for section references — spell out \"section\" in full.",
                        Truncate(text, 60), p.Start);
                    found++;
                }

                // ── NEW Rule 2: "subsection N" with a label ───────────────────
                // "subsection" alone (e.g. "the next subsection") is allowed.
                // Only flag when it is followed by an actual section label.
                foreach (Match m in RxSubsectionLabelled.Matches(text))
                {
                    Msg("HEADING", "ERROR",
                        "\"" + Truncate(m.Value, 40) + "\" — do not use \"subsection\" with a label. " +
                        "Use \"section\" instead (e.g. \"section 3a\" not \"subsection 3a\"). " +
                        "\"subsection\" may only be used alone without a label.",
                        Truncate(text, 60), p.Start);
                    found++;
                }

                // ── NEW Rule 3: H4 refs must use two sets of parentheses ──────
                // Case A: both sets of parentheses missing — e.g. 4a2ii
                foreach (Match m in RxH4CiteNoParens.Matches(text))
                {
                    string correct = m.Groups[1].Value + "(" + m.Groups[2].Value + ")(" + m.Groups[3].Value + ")";
                    Msg("HEADING", "ERROR",
                        "H4 section reference \"" + m.Value + "\" must use two sets of parentheses — " +
                        "use \"" + correct + "\" instead (e.g. \"4a(2)(ii)\").",
                        Truncate(text, 60), p.Start);
                    found++;
                }

                // Case B: second set of parentheses missing — e.g. 4a(2)ii
                foreach (Match m in RxH4CiteMissingSecondParens.Matches(text))
                {
                    string correct = m.Groups[1].Value + "(" + m.Groups[2].Value + ")(" + m.Groups[3].Value + ")";
                    Msg("HEADING", "ERROR",
                        "H4 section reference \"" + m.Value + "\" is missing the second set of parentheses — " +
                        "use \"" + correct + "\" instead (e.g. \"4a(2)(ii)\").",
                        Truncate(text, 60), p.Start);
                    found++;
                }

                // Case C: first set of parentheses missing — e.g. 4a2(ii)
                foreach (Match m in RxH4CiteMissingFirstParens.Matches(text))
                {
                    string correct = m.Groups[1].Value + "(" + m.Groups[2].Value + ")(" + m.Groups[3].Value + ")";
                    Msg("HEADING", "ERROR",
                        "H4 section reference \"" + m.Value + "\" is missing the first set of parentheses — " +
                        "use \"" + correct + "\" instead (e.g. \"4a(2)(ii)\").",
                        Truncate(text, 60), p.Start);
                    found++;
                }

                // ── Existing: singular/plural + label existence checks ─────────
                foreach (Match cite in RxSectionCite.Matches(text))
                {
                    string citeVal = cite.Value;
                    var tokensToCheck = new List<string>();
                    var rangeMatch = RxSectionRange.Match(citeVal.Substring(citeVal.IndexOf(' ') + 1).Trim());
                    if (rangeMatch.Success)
                    { tokensToCheck.Add(rangeMatch.Groups[1].Value); tokensToCheck.Add(rangeMatch.Groups[2].Value); }
                    else
                    { foreach (Match t in RxSectionToken.Matches(citeVal)) tokensToCheck.Add(t.Value); }

                    bool usedPlural = citeVal.TrimStart().StartsWith("sections ", StringComparison.OrdinalIgnoreCase);
                    int tokenCount = tokensToCheck.Count;
                    if (tokenCount > 1 && !usedPlural)
                    { Msg("HEADING", "WARNING", "\"" + Truncate(citeVal, 40) + "\" cites multiple sections but uses singular \"section\" — should be \"sections\".", Truncate(text, 60), p.Start); found++; }
                    else if (tokenCount == 1 && usedPlural)
                    { Msg("HEADING", "WARNING", "\"" + Truncate(citeVal, 40) + "\" cites one section but uses plural \"sections\" — should be \"section\".", Truncate(text, 60), p.Start); found++; }

                    foreach (string token in tokensToCheck)
                    {
                        // ── Layer 1: is this a format we can validate at all? ──
                        if (!RxKnownGoodToken.IsMatch(token))
                        {
                            Msg("HEADING", "WARNING",
                                "Section reference \"" + token + "\" in \"" + Truncate(citeVal, 40) + "\" has an unrecognised format — please check manually.",
                                Truncate(text, 60), p.Start);
                            found++;
                            continue;
                        }

                        // ── Layer 2: known format — validate against actual headings ──
                        bool missing = false; string reason = "";
                        var tm = Regex.Match(token, @"^([A-Z])?(\d+)([a-z])?(?:\((\d+)\))?(?:\(([ivxlcdm]+)\))?$", RegexOptions.IgnoreCase);
                        if (!tm.Success)
                        {
                            // Passed Layer 1 but failed detailed parse — should not happen,
                            // but flag for manual check rather than silently skipping.
                            Msg("HEADING", "WARNING",
                                "Section reference \"" + token + "\" in \"" + Truncate(citeVal, 40) + "\" could not be parsed — please check manually.",
                                Truncate(text, 60), p.Start);
                            found++;
                            continue;
                        }
                        bool hasH1Num = tm.Groups[2].Success; bool hasH2Letter = tm.Groups[3].Success;
                        bool hasH3Num = tm.Groups[4].Success; bool hasH4Roman = tm.Groups[5].Success;
                        if (hasH4Roman) { string roman = tm.Groups[5].Value.ToLowerInvariant(); if (!h4Romans.Contains(roman)) { missing = true; reason = "heading-04 \"(" + roman + ")\""; } }
                        else if (hasH3Num) { if (int.TryParse(tm.Groups[4].Value, out int h3n) && !h3Numbers.Contains(h3n)) { missing = true; reason = "heading-03 \"" + h3n + ")\""; } }
                        else if (hasH2Letter) { char letter = tm.Groups[3].Value.ToLowerInvariant()[0]; if (!h2Letters.Contains(letter)) { missing = true; reason = "heading-02 \"" + letter + ".\""; } }
                        else if (hasH1Num) { if (int.TryParse(tm.Groups[2].Value, out int h1n) && !h1Numbers.Contains(h1n)) { missing = true; reason = "heading-01 \"" + h1n + ".\""; } }
                        if (missing) { Msg("HEADING", "WARNING", "\"" + token + "\" is cited (\"" + Truncate(citeVal, 40) + "\") but the corresponding " + reason + " heading was not found — please check and confirm.", Truncate(text, 60), p.Start); found++; }
                    }
                }
            }
            return found;
        }

        // =====================================================================
        // CHECK — CURRENT AFFILIATION
        // =====================================================================
        private int CheckCurrentAffiliation(List<ParaInfo> paras)
        {
            int found = 0;
            string affKey = NK("footnote_aff"), authorsKey = NK("authors");
            var affParas = paras.Where(p => p.SK == affKey).ToList();
            if (affParas.Count == 0) return 0;
            var authorsPara = paras.FirstOrDefault(p => p.SK == authorsKey);
            var authorList = new List<(string SurnameLower, string Initial, string SurnameRaw)>();
            if (authorsPara != null)
            {
                string authText = Regex.Replace(authorsPara.Text, @"\s+and\s+", ", ", RegexOptions.IgnoreCase);
                foreach (Match m in RxAuthorToken.Matches(authText))
                    authorList.Add((m.Groups[2].Value.ToLowerInvariant(), m.Groups[1].Value[0].ToString().ToUpperInvariant(), m.Groups[2].Value));
            }
            var surnameGroups = authorList.GroupBy(a => a.SurnameLower).Where(g => g.Count() > 1)
                .ToDictionary(g => g.Key, g => g.Select(a => a.Initial).ToList());
            foreach (var affPara in affParas)
            {
                if (surnameGroups.Any())
                {
                    var needInitial = surnameGroups.Select(kvp => { string raw = authorList.First(a => a.SurnameLower == kvp.Key).SurnameRaw; return string.Join(". and ", kvp.Value.Select(i => i + ".")) + " " + raw; }).ToList();
                    Msg("HEADING", "WARNING", "Current affiliation detected — use author surname only. Note: the following surname(s) are shared by multiple authors and must include the initial: " + string.Join("; ", needInitial) + " (e.g. \"Current affiliation of " + needInitial[0] + ": ...\").", Truncate(affPara.Text, 60), affPara.Start);
                }
                else
                    Msg("HEADING", "WARNING", "Current affiliation detected — use author surname only (e.g. \"Current affiliation of Smith: ...\"). Do not include full name or affiliation number.", Truncate(affPara.Text, 60), affPara.Start);
                found++;
            }
            return found;
        }

        // =====================================================================
        // CHECK 1 — FRONT MATTER
        // =====================================================================
        private int CheckFrontMatter(List<ParaInfo> all, int firstH1Idx)
        {
            int found = 0;
            var front = firstH1Idx >= 0 ? all.Take(firstH1Idx).ToList() : all.ToList();
            var styles = DeduplicatedStyleKeys(front);
            foreach (var (style, label, required) in FrontMatter)
            {
                string key = NK(style), alt = style == "received" ? NK("history_info") : null;
                if (!styles.Contains(key) && (alt == null || !styles.Contains(alt)))
                { Msg("HEADING", "ERROR", $"{label} is missing — it should appear before the first heading-01."); found++; }
            }
            int lastPos = -1; string lastLabel = "";
            foreach (var (style, label, required) in FrontMatter)
            {
                string key = NK(style), alt = style == "received" ? NK("history_info") : null;
                int pos = styles.IndexOf(key); if (pos < 0 && alt != null) pos = styles.IndexOf(alt); if (pos < 0) continue;
                if (pos < lastPos) { Msg("HEADING", "ERROR", $"{label} is out of order — it should appear after {lastLabel}.", "", GetFirstStart(front, pos == styles.IndexOf(key) ? key : alt)); found++; }
                else { lastPos = pos; lastLabel = label; }
            }
            return found;
        }

        // =====================================================================
        // CHECK 2 — BACK MATTER
        // =====================================================================
        private int CheckBackMatter(List<ParaInfo> all, int ackIdx)
        {
            int found = 0;
            if (ackIdx < 0) { Msg("HEADING", "ERROR", "Acknowledgments is missing — it should appear after the body sections."); return ++found; }
            var back = all.Skip(ackIdx).ToList();
            var styles = DeduplicatedStyleKeys(back);
            string ResolveKey(string style) { string k1 = NK(style); if (styles.Contains(k1)) return k1; string k2 = NK(style.Replace("acknowledgement", "acknowledgment")); return styles.Contains(k2) ? k2 : k1; }
            foreach (var (style, label, required) in BackMatter)
                if (!styles.Contains(ResolveKey(style))) { Msg("HEADING", "ERROR", $"{label} is missing — it should appear after the body sections."); found++; }
            int bmLastPos = -1; string bmLastLabel = "";
            foreach (var (style, label, required) in BackMatter)
            {
                int pos = styles.IndexOf(ResolveKey(style)); if (pos < 0) continue;
                if (pos < bmLastPos) { Msg("HEADING", "ERROR", $"{label} is out of order — it should appear after {bmLastLabel}.", "", GetFirstStart(back, ResolveKey(style))); found++; }
                else { bmLastPos = pos; bmLastLabel = label; }
            }
            return found;
        }

        // =====================================================================
        // CHECK 3 — BACK-MATTER RELATIVE ORDER
        // =====================================================================
        private int CheckBackMatterRelativeOrder(List<ParaInfo> paras)
        {
            int found = 0;
            int ackPos = FirstPos(paras, IsAckTitle);
            int refPos = FirstPos(paras, p => p.SK == NK("reference_title"));
            int figPos = FirstPos(paras, p => p.SK == NK("figure_caption"));
            int tblPos = FirstPos(paras, p => p.SK == NK("table_caption"));
            if (ackPos >= 0 && refPos >= 0 && ackPos > refPos) { Msg("HEADING", "ERROR", "Acknowledgments is out of order — it should appear before References.", "", paras[ackPos].Start); found++; }
            if (refPos >= 0 && figPos >= 0 && refPos > figPos) { Msg("HEADING", "ERROR", "References is out of order — it should appear before figure captions.", "", paras[refPos].Start); found++; }
            if (refPos >= 0 && tblPos >= 0 && refPos > tblPos) { Msg("HEADING", "ERROR", "References is out of order — it should appear before table captions.", "", paras[refPos].Start); found++; }
            if (figPos >= 0 && tblPos >= 0 && figPos > tblPos) { Msg("HEADING", "WARNING", "Figure captions is out of order — it should appear before table captions.", "", paras[figPos].Start); found++; }
            return found;
        }

        // =====================================================================
        // CHECK 4 — HEADING FORMAT RULES
        // =====================================================================
        private int CheckHeadingRules(Word.Document doc, List<ParaInfo> allParas, int ackIdx)
        {
            int found = 0;
            string docPath = ""; try { docPath = doc.FullName; } catch { }
            bool canOoxml = !string.IsNullOrEmpty(docPath) && File.Exists(docPath) &&
                (docPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) || docPath.EndsWith(".docm", StringComparison.OrdinalIgnoreCase));
            var fmtMap = canOoxml ? ReadHeadingFormats(docPath, allParas) : new Dictionary<int, FontInfo>();
            var backH1Starts = new HashSet<int>();
            if (ackIdx >= 0) foreach (var p in allParas.Skip(ackIdx).Where(IsH1)) backH1Starts.Add(p.Start);
            foreach (var p in allParas)
            {
                if (!IsHeading(p)) continue;
                string lv = IsH1(p) ? "heading-01" : IsH2(p) ? "heading-02" : IsH3(p) ? "heading-03" : "heading-04";
                bool isBackH1 = IsH1(p) && backH1Starts.Contains(p.Start);
                if (IsH1(p) && !isBackH1 && !RxH1Pattern.IsMatch(p.Text)) Err(lv, "Must start with a section number and period (e.g. \"1. Introduction\").", p, ref found);
                else if (IsH2(p) && !RxH2Pattern.IsMatch(p.Text)) Err(lv, "Must start with a lowercase letter and period (e.g. \"a. Title\").", p, ref found);
                else if (IsH3(p) && !RxH3Pattern.IsMatch(p.Text)) Err(lv, "Must start with a number and closing parenthesis (e.g. \"1) Title\").", p, ref found);
                else if (IsH4(p) && !RxH4Pattern.IsMatch(p.Text)) Err(lv, "Must start with a roman numeral in parentheses (e.g. \"(i) Title\").", p, ref found);
                if (isBackH1) { bool isDA = p.Text.IndexOf("data availab", StringComparison.OrdinalIgnoreCase) >= 0; if (isDA && !p.Text.TrimEnd().EndsWith(".")) Err(lv, $"\"{p.Text.Trim()}\" must end with a period (.).", p, ref found); }
                else if (!IsH4(p)) { if (RxEndPunct.IsMatch(p.Text)) Err(lv, "Heading ends with punctuation (. or :) — AMS headings must not end with punctuation.", p, ref found); }
                if (IsH1(p) && !isBackH1 && p.Text == p.Text.ToUpper() && p.Text.Length > 3 && RxAllCaps.IsMatch(p.Text)) Warn(lv, "ALL CAPS heading — AMS uses sentence case.", p, ref found);
                if (fmtMap.TryGetValue(p.Start, out var fi))
                {
                    if (IsH1(p)) { if (!fi.Bold) Warn(lv, "Should be Bold (16pt Bold).", p, ref found); if (fi.Italic) Warn(lv, "Should NOT be italic.", p, ref found); if (fi.Pt > 0 && Math.Abs(fi.Pt - 16) > 0.5) Warn(lv, $"Font size should be 16pt, found {fi.Pt}pt.", p, ref found); }
                    else if (IsH2(p)) { if (!fi.Italic) Warn(lv, "Should be Italic (15pt Italic).", p, ref found); if (fi.Bold) Warn(lv, "Should NOT be bold.", p, ref found); if (fi.Pt > 0 && Math.Abs(fi.Pt - 15) > 0.5) Warn(lv, $"Font size should be 15pt, found {fi.Pt}pt.", p, ref found); }
                    else if (IsH3(p)) { if (!fi.SmallCaps) Warn(lv, "Should be Small Caps (14pt Small Caps).", p, ref found); if (fi.Bold) Warn(lv, "Should NOT be bold.", p, ref found); if (fi.Italic) Warn(lv, "Should NOT be italic.", p, ref found); if (fi.Pt > 0 && Math.Abs(fi.Pt - 14) > 0.5) Warn(lv, $"Font size should be 14pt, found {fi.Pt}pt.", p, ref found); }
                    else if (IsH4(p)) { if (!fi.Italic) Warn(lv, "Should be Italic (14pt Italic).", p, ref found); if (fi.Bold) Warn(lv, "Should NOT be bold.", p, ref found); if (fi.Pt > 0 && Math.Abs(fi.Pt - 14) > 0.5) Warn(lv, $"Font size should be 14pt, found {fi.Pt}pt.", p, ref found); }
                }
            }
            string absSK = NK("abstract_title");
            foreach (var p in allParas.Where(p => p.SK == absSK))
                if (fmtMap.TryGetValue(p.Start, out var absFi) && !absFi.Bold)
                { Msg("HEADING", "ERROR", "Abstract title is not Bold — it must be formatted as Bold.", Truncate(p.Text, 60), p.Start); found++; }
            return found;
        }

        private void Err(string lv, string msg, ParaInfo p, ref int found)
        { Msg("HEADING", "ERROR", $"[{lv}] {msg}", Truncate(p.Text, 60), p.Start); found++; }
        private void Warn(string lv, string msg, ParaInfo p, ref int found)
        { Msg("HEADING", "WARNING", $"[{lv}] {msg}", Truncate(p.Text, 60), p.Start); found++; }

        // =====================================================================
        // CHECK 5 — HEADING LEVEL SEQUENCE
        // =====================================================================
        private int CheckHeadingSequence(List<ParaInfo> paras)
        {
            int found = 0, lastLevel = 0; bool hasH1 = false;
            foreach (var p in paras)
            {
                int level = HeadingLevel(p); if (level == 0) continue;
                if (level == 1) { hasH1 = true; lastLevel = 1; continue; }
                if (!hasH1) { Msg("HEADING", "ERROR", $"heading-0{level} appears before any heading-01 — the first body heading must be a heading-01.", Truncate(p.Text, 60), p.Start); found++; lastLevel = level; continue; }
                if (level > lastLevel + 1) { Msg("HEADING", "ERROR", $"Heading level skipped — heading-0{level} follows heading-0{lastLevel} but heading-0{lastLevel + 1} is expected next.", Truncate(p.Text, 60), p.Start); found++; }
                lastLevel = level;
            }
            return found;
        }

        // =====================================================================
        // CHECK 6 — HEADING LABEL SEQUENCE
        // =====================================================================
        private int CheckHeadingLabelSequence(List<ParaInfo> paras)
        {
            int found = 0, expectedH1Num = 1; bool h1NumStarted = false;
            char expectedH2 = 'a'; int expectedH3 = 1, expectedH4 = 1;
            foreach (var p in paras)
            {
                if (IsH1(p))
                {
                    expectedH2 = 'a'; expectedH3 = 1; expectedH4 = 1;
                    var mH1 = RxH1Pattern.Match(p.Text); if (!mH1.Success) continue;
                    string numStr = p.Text.Substring(0, p.Text.IndexOf('.'));
                    if (!int.TryParse(numStr.Trim(), out int actualNum)) continue;
                    if (!h1NumStarted) { expectedH1Num = actualNum + 1; h1NumStarted = true; continue; }
                    if (actualNum < expectedH1Num) { Msg("HEADING", "ERROR", $"Section {actualNum}. is already used — this heading should be {expectedH1Num}.", Truncate(p.Text, 60), p.Start); found++; }
                    else if (actualNum > expectedH1Num) { string missing = string.Join(", ", Enumerable.Range(expectedH1Num, actualNum - expectedH1Num).Select(n => $"{n}.")); Msg("HEADING", "ERROR", $"Section number(s) {missing} are missing — found {actualNum}. but expected {expectedH1Num}.", Truncate(p.Text, 60), p.Start); found++; expectedH1Num = actualNum + 1; }
                    else expectedH1Num = actualNum + 1; continue;
                }
                if (IsH2(p)) { expectedH3 = 1; expectedH4 = 1; var m = RxH2Label.Match(p.Text); if (!m.Success) continue; char actual = m.Groups[1].Value[0]; if (actual != expectedH2) { Msg("HEADING", "ERROR", $"heading-02 label is wrong — found \"{actual}.\" but expected \"{expectedH2}.\".", Truncate(p.Text, 60), p.Start); found++; } expectedH2 = (char)(actual + 1); continue; }
                if (IsH3(p)) { expectedH4 = 1; var m = RxH3Label.Match(p.Text); if (!m.Success) continue; int actual = int.Parse(m.Groups[1].Value); if (actual != expectedH3) { Msg("HEADING", "ERROR", $"heading-03 label is wrong — found \"{actual})\" but expected \"{expectedH3})\".", Truncate(p.Text, 60), p.Start); found++; } expectedH3 = actual + 1; continue; }
                if (IsH4(p)) { var m = RxH4Label.Match(p.Text); if (!m.Success) continue; int actual = RomanToInt(m.Groups[1].Value.ToLowerInvariant()); if (actual <= 0) continue; if (actual != expectedH4) { Msg("HEADING", "ERROR", $"heading-04 label is wrong — found \"({IntToRoman(actual)})\" but expected \"({IntToRoman(expectedH4)})\".", Truncate(p.Text, 60), p.Start); found++; } expectedH4 = actual + 1; }
            }
            return found;
        }

        // =====================================================================
        // CHECK 7 — DISPLAY EQUATION FOLLOW-UP
        // =====================================================================
        private int CheckDisplayEquationFollowUp(List<ParaInfo> paras)
        {
            int found = 0;
            string eqKey = NK("display_equation"), noIndKey = NK("para_no_indent"), noIndKeyH = NK("para_no-indent");
            bool lastWasEq = false;
            foreach (var p in paras)
            {
                if (p.SK == eqKey) { lastWasEq = true; continue; }
                if (lastWasEq) { if (p.SK != noIndKey && p.SK != noIndKeyH) { Msg("HEADING", "ERROR", "After a display equation the next paragraph must use para_no_indent style — found \"" + p.SN + "\" instead.", Truncate(p.Text, 60), p.Start); found++; } lastWasEq = false; }
            }
            return found;
        }

        // =====================================================================
        // CHECK 8 — SPECIAL TITLE STYLES
        // =====================================================================
        private int CheckSpecialTitleStyles(List<ParaInfo> paras)
        {
            int found = 0;
            string ackKey = NK("acknowledgement_title"), ackKey2 = NK("acknowledgment_title"), appKey = NK("appendix_title");
            foreach (var p in paras)
            {
                if (p.SK == ackKey || p.SK == ackKey2) { if (!p.Text.TrimEnd().EndsWith(".")) { Msg("HEADING", "ERROR", "Acknowledgments title must end with a period — add a period at the end (e.g. \"Acknowledgments.\").", Truncate(p.Text, 60), p.Start); found++; } continue; }
                if (p.SK == appKey) { var badWords = p.Text.Trim().Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Where(w => w.Length > 0 && char.IsLetter(w[0]) && char.IsLower(w[0])).ToList(); if (badWords.Any()) { Msg("HEADING", "ERROR", "Appendix title must use Title Case — word(s) not capitalised: " + string.Join(", ", badWords.Select(w => $"\"{w}\"")) + ".", Truncate(p.Text, 60), p.Start); found++; } }
            }
            return found;
        }

        // =====================================================================
        // CHECK 8b — NON-HEADING STYLE FORMAT
        // =====================================================================
        private int CheckNonHeadingStyles(List<ParaInfo> paras)
        {
            int found = 0;
            string absKey = NK("abstract_title"), refKey = NK("reference_title");
            foreach (var p in paras)
            {
                if (p.SK == absKey && RxEndPunct.IsMatch(p.Text)) { Msg("HEADING", "ERROR", "Abstract title ends with punctuation — it must not end with punctuation.", Truncate(p.Text, 60), p.Start); found++; }
                if (p.SK == refKey && RxEndPunct.IsMatch(p.Text)) { Msg("HEADING", "ERROR", "References title ends with punctuation — remove the period or colon.", Truncate(p.Text, 60), p.Start); found++; }
            }
            return found;
        }

        // =====================================================================
        // CHECK — HISTORY INFO CONTENT
        // =====================================================================
        // The history info (received/accepted dates) paragraph must not contain
        // zero-padded single-digit numbers such as 01., 02. … 09.
        // Correct form: "1 January 2024" not "01 January 2024".
        private int CheckHistoryInfoContent(List<ParaInfo> paras)
        {
            int found = 0;
            string histKey = NK("received");        // style key used in FrontMatter
            string histKey2 = NK("history_info");    // alternate style name
            foreach (var p in paras)
            {
                if (p.SK != histKey && p.SK != histKey2) continue;
                foreach (Match m in RxHistoryZeroPadded.Matches(p.Text))
                {
                    Msg("HEADING", "ERROR",
                        "History info contains a zero-padded number \"" + m.Value + "\" — " +
                        "remove the leading zero (e.g. use \"1.\" not \"01.\").",
                        Truncate(p.Text, 60), p.Start);
                    found++;
                }
            }
            return found;
        }

        // =====================================================================
        // CHECK — APPENDIX CITATIONS IN BODY TEXT
        // =====================================================================
        // Rule A: If only one appendix_title exists with no letter (plain "Appendix.")
        //         → the letter "A" should be removed from the title (flag it).
        //         AND the body text must contain at least one reference to "Appendix"
        //         (bare, no letter).
        //
        // Rule B: If appendix_title headings exist with letters (Appendix A, Appendix B …)
        //         → for each lettered appendix title, at least one "Appendix X" citation
        //         must appear somewhere in the scoped body text.
        //
        // Rule C: If an appendix_title exists (lettered or not) but ZERO appendix
        //         references appear anywhere in the body → flag it.
        //
        // All matching here is case-insensitive: a title styled "APPENDIX A.",
        // "Appendix A.", or "appendix a." is treated the same, and a citation
        // written "appendix a.", "Appendix A", or "APPENDIX A" all count.
        private int CheckAppendixCitations(List<ParaInfo> paras)
        {
            int found = 0;
            string appTitleKey = NK("appendix_title");

            // ── Collect all appendix_title paragraphs ─────────────────────────
            var appTitles = paras.Where(p => p.SK == appTitleKey).ToList();
            if (appTitles.Count == 0) return 0;   // no appendix at all — nothing to check here

            // Determine which letters are declared as appendix titles
            // e.g. "Appendix A." / "APPENDIX A." → letter 'A', plain "Appendix." → no letter
            var declaredLetters = new List<char>();   // letters found in titles (always uppercase)
            bool hasBareTitle = false;              // title with no letter

            foreach (var ap in appTitles)
            {
                var lm = RxAppendixLetter.Match(ap.Text);
                if (lm.Success)
                    declaredLetters.Add(char.ToUpperInvariant(lm.Groups[1].Value[0]));
                else
                    hasBareTitle = true;
            }

            // ── Rule A: bare "Appendix" title when no second appendix exists ──
            // If there is a bare title AND there are no lettered titles at all,
            // check whether there is a second appendix title — if there isn't,
            // the letter "A" (if present) should be dropped.
            // We re-examine the actual title text here.
            if (hasBareTitle && declaredLetters.Count == 0)
            {
                // Check each bare title for an incorrectly retained letter
                foreach (var ap in appTitles.Where(ap => !RxAppendixLetter.IsMatch(ap.Text)))
                {
                    // If the title contains " A" right before the period it should be plain "Appendix."
                    // This is already handled by CheckAppendixReferences for the lettered case,
                    // but here we ensure the bare title doesn't accidentally say "Appendix A"
                    // when no Appendix B exists — that is caught by the letter-check below.
                }
            }

            // ── Rule B: if only "Appendix A" exists (no B), title should be plain "Appendix." ──
            if (declaredLetters.Count == 1 && declaredLetters[0] == 'A' && !hasBareTitle)
            {
                // Check there is no Appendix B title
                bool hasB = declaredLetters.Contains('B');
                if (!hasB)
                {
                    var apA = appTitles.First(ap => RxAppendixLetter.IsMatch(ap.Text) &&
                                                     string.Equals(RxAppendixLetter.Match(ap.Text).Groups[1].Value, "A", StringComparison.OrdinalIgnoreCase));
                    Msg("HEADING", "ERROR",
                        "Only one appendix exists — remove the letter \"A\" from the appendix title. " +
                        "Change \"" + Truncate(apA.Text, 40) + "\" to \"Appendix.\"",
                        Truncate(apA.Text, 60), apA.Start);
                    found++;
                }
            }

            // ── Collect all body-text appendix references ─────────────────────
            var bodyParas = paras.Where(p => ScopedStyles.Contains(p.SK)).ToList();

            // Letters cited in body text e.g. "Appendix A", "appendix b" (always stored uppercase)
            var citedLetters = new HashSet<char>();
            bool citedBare = false;

            foreach (var p in bodyParas)
            {
                foreach (Match m in RxAppendixCiteWithLetter.Matches(p.Text))
                    citedLetters.Add(char.ToUpperInvariant(m.Groups[1].Value[0]));
                if (RxAppendixCiteBare.IsMatch(p.Text))
                    citedBare = true;
            }

            // ── Rule C: appendix exists but never cited ───────────────────────
            bool anyCited = citedBare || citedLetters.Count > 0;
            if (!anyCited)
            {
                Msg("HEADING", "ERROR",
                    "Appendix heading is present but there is no reference to it anywhere in the body text — " +
                    "add a citation such as \"(see Appendix)\" or \"(see Appendix A)\".",
                    "", appTitles.First().Start);
                found++;
                return found;   // no point checking individual letters if nothing is cited at all
            }

            // ── Rule D: each declared lettered appendix must be cited ─────────
            foreach (char letter in declaredLetters)
            {
                if (!citedLetters.Contains(letter))
                {
                    var apPara = appTitles.FirstOrDefault(ap =>
                    {
                        var lm = RxAppendixLetter.Match(ap.Text);
                        return lm.Success && char.ToUpperInvariant(lm.Groups[1].Value[0]) == letter;
                    });
                    Msg("HEADING", "ERROR",
                        "\"Appendix " + letter + "\" heading is present but is never cited in the body text — " +
                        "add a reference to \"Appendix " + letter + "\" in the text.",
                        apPara != null ? Truncate(apPara.Text, 60) : "",
                        apPara?.Start ?? 0);
                    found++;
                }
            }

            // ── Rule E: bare appendix title must be cited as bare "Appendix" ──
            if (hasBareTitle && !citedBare)
            {
                var apPara = appTitles.First(ap => !RxAppendixLetter.IsMatch(ap.Text));
                Msg("HEADING", "ERROR",
                    "\"Appendix\" heading is present but is never cited in the body text — " +
                    "add a reference to \"Appendix\" in the text.",
                    Truncate(apPara.Text, 60), apPara.Start);
                found++;
            }

            return found;
        }

        // =====================================================================
        // CHECK 9 — APPENDIX REFERENCES
        // =====================================================================
        private int CheckAppendixReferences(List<ParaInfo> paras)
        {
            int found = 0;
            string appTitleKey = NK("appendix_title");
            var appTitleTexts = paras.Where(p => p.SK == appTitleKey).Select(p => p.Text.Trim()).ToList();
            bool anyAppTitle = appTitleTexts.Count > 0;
            foreach (var p in paras)
            {
                if (!ScopedStyles.Contains(p.SK)) continue;
                if (!RxAppendixWord.IsMatch(p.Text)) continue;
                if (!anyAppTitle) { Msg("HEADING", "ERROR", "\"" + GetAppendixMatch(p.Text) + "\" is referenced but no appendix_title heading was found — add an appendix_title paragraph.", Truncate(p.Text, 60), p.Start); found++; continue; }
                foreach (Match m in RxAppendixLetter.Matches(p.Text))
                {
                    string letter = m.Groups[1].Value.ToUpperInvariant();
                    bool matchFound = appTitleTexts.Any(t => t.IndexOf("Appendix " + letter, StringComparison.OrdinalIgnoreCase) >= 0 || t.IndexOf("Appendix" + letter, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (!matchFound) { Msg("HEADING", "ERROR", "\"Appendix " + letter + "\" is referenced but no matching appendix_title heading \"Appendix " + letter + "\" was found.", Truncate(p.Text, 60), p.Start); found++; }
                }
            }
            return found;
        }

        private static string GetAppendixMatch(string text) { Match m = RxAppendixWord.Match(text); return m.Success ? m.Value : "appendix"; }

        // =====================================================================
        // CHECK — DISPLAY EQUATION NUMBERING
        // =====================================================================

        private enum EqKind { Simple, Sectional, Appendix, Unknown }

        // Structured representation of whatever numbering tail was found at the
        // end of a display_equation paragraph.
        private sealed class EqNumber
        {
            internal EqKind Kind;
            internal int? Section;     // sectional: the section part, e.g. "1" in (1.2)
            internal int Number;       // simple: the number; sectional: the sub-number; appendix: the number after the letter
            internal char? SubLetter;  // lowercase sub-equation letter, e.g. (1a) → 'a'
            internal char? ApxLetter;  // uppercase appendix letter, e.g. (A.1) / "A1." → 'A'
            internal bool LeadingZero; // true if any numeric part was zero-padded, e.g. "03"
            internal string RawTail;   // the raw matched numbering text, for messages
        }

        // Tries Appendix → Sectional → Simple, parenthesized form before bare
        // form. If nothing matches, returns EqKind.Unknown so the caller can
        // flag it for manual review instead of silently ignoring it.
        private static EqNumber ParseEquationNumber(string rawText)
        {
            string text = RxTEQuery.Replace(rawText ?? "", "").Trim();

            // ── Appendix form: "(A.1)", "(A1a)", "A1.", "A2 .", "B1." ──────────
            var m = RxEqApxParen.Match(text);
            if (!m.Success) m = RxEqApxBare.Match(text);
            if (m.Success)
            {
                string numStr = m.Groups[2].Value;
                return new EqNumber
                {
                    Kind = EqKind.Appendix,
                    ApxLetter = char.ToUpperInvariant(m.Groups[1].Value[0]),
                    Number = int.Parse(numStr),
                    SubLetter = (m.Groups[3].Success && m.Groups[3].Value.Length > 0) ? (char?)m.Groups[3].Value[0] : null,
                    LeadingZero = numStr.Length > 1 && numStr[0] == '0',
                    RawTail = m.Value.Trim()
                };
            }

            // ── Sectional form: "(1.2)", "(1.2a)", "1.2.", "01.2." ─────────────
            m = RxEqSectParen.Match(text);
            if (!m.Success) m = RxEqSectBare.Match(text);
            if (m.Success)
            {
                string secStr = m.Groups[1].Value, subStr = m.Groups[2].Value;
                return new EqNumber
                {
                    Kind = EqKind.Sectional,
                    Section = int.Parse(secStr),
                    Number = int.Parse(subStr),
                    SubLetter = (m.Groups[3].Success && m.Groups[3].Value.Length > 0) ? (char?)m.Groups[3].Value[0] : null,
                    LeadingZero = (secStr.Length > 1 && secStr[0] == '0') || (subStr.Length > 1 && subStr[0] == '0'),
                    RawTail = m.Value.Trim()
                };
            }

            // ── Simple form: "(1)", "(1a)", "(2b)", "03." ───────────────────────
            m = RxEqSimpleParen.Match(text);
            if (!m.Success) m = RxEqSimpleBare.Match(text);
            if (m.Success)
            {
                string numStr = m.Groups[1].Value;
                return new EqNumber
                {
                    Kind = EqKind.Simple,
                    Number = int.Parse(numStr),
                    SubLetter = (m.Groups[2].Success && m.Groups[2].Value.Length > 0) ? (char?)m.Groups[2].Value[0] : null,
                    LeadingZero = numStr.Length > 1 && numStr[0] == '0',
                    RawTail = m.Value.Trim()
                };
            }

            // ── Fail-safe: format not recognised — flag for manual review ──────
            return new EqNumber { Kind = EqKind.Unknown, RawTail = Truncate(text, 30) };
        }

        private int CheckDisplayEquationNumbering(List<ParaInfo> paras)
        {
            int found = 0;
            string eqKey = NK("display_equation");
            var eqParas = paras.Where(p => p.SK == eqKey).ToList();
            if (eqParas.Count == 0) return 0;

            var parsed = eqParas.Select(p => (Para: p, Eq: ParseEquationNumber(p.Text))).ToList();

            // ── Fail-safe: unrecognised numbering format ────────────────────────
            foreach (var item in parsed.Where(i => i.Eq.Kind == EqKind.Unknown))
            {
                Msg("HEADING", "WARNING",
                    "Display equation has an unrecognised numbering format — please check the equation number manually.",
                    Truncate(item.Para.Text, 60), item.Para.Start);
                found++;
            }

            // ── Zero-padded numbers, e.g. "(03)" or "03." ────────────────────────
            foreach (var item in parsed.Where(i => i.Eq.Kind != EqKind.Unknown && i.Eq.LeadingZero))
            {
                Msg("HEADING", "WARNING",
                    "Display equation number \"" + item.Eq.RawTail + "\" is zero-padded — remove the leading zero (e.g. use \"3\" not \"03\").",
                    Truncate(item.Para.Text, 60), item.Para.Start);
                found++;
            }

            // ── Appendix equations — grouped by letter (A, B, …) ─────────────────
            var apxGroups = parsed.Where(i => i.Eq.Kind == EqKind.Appendix)
                                   .GroupBy(i => i.Eq.ApxLetter.Value)
                                   .OrderBy(g => g.Key);
            foreach (var grp in apxGroups)
            {
                char letter = grp.Key;
                var items = grp.OrderBy(i => i.Para.Start)
                                .Select(i => (i.Eq.Number, i.Eq.SubLetter, i.Para)).ToList();
                found += CheckNumberLetterSequence(items,
                    num => $"Appendix equation ({letter}{num})",
                    (num, l) => $"Appendix equation ({letter}{num}{l})",
                    num => $"({letter}{num})",
                    (num, l) => $"({letter}{num}{l})");
            }

            // ── Sectional equations — grouped by section ─────────────────────────
            var sectGroups = parsed.Where(i => i.Eq.Kind == EqKind.Sectional)
                                    .GroupBy(i => i.Eq.Section.Value)
                                    .OrderBy(g => g.Key)
                                    .ToList();
            bool anySectional = sectGroups.Count > 0;
            foreach (var grp in sectGroups)
            {
                int section = grp.Key;
                var items = grp.OrderBy(i => i.Para.Start)
                                .Select(i => (i.Eq.Number, i.Eq.SubLetter, i.Para)).ToList();
                found += CheckNumberLetterSequence(items,
                    num => $"Display equation ({section}.{num})",
                    (num, l) => $"Display equation ({section}.{num}{l})",
                    num => $"({section}.{num})",
                    (num, l) => $"({section}.{num}{l})");
            }

            // ── Simple equations (only checked when sectional numbering is NOT in use) ──
            if (!anySectional)
            {
                var items = parsed.Where(i => i.Eq.Kind == EqKind.Simple)
                                   .OrderBy(i => i.Para.Start)
                                   .Select(i => (i.Eq.Number, i.Eq.SubLetter, i.Para)).ToList();
                found += CheckNumberLetterSequence(items,
                    num => $"Display equation ({num})",
                    (num, l) => $"Display equation ({num}{l})",
                    num => $"({num})",
                    (num, l) => $"({num}{l})");
            }
            else
            {
                // Article uses sectional numbering elsewhere — flag any equation
                // still using bare simple numbering as a possible mismatch.
                foreach (var item in parsed.Where(i => i.Eq.Kind == EqKind.Simple))
                {
                    Msg("HEADING", "WARNING",
                        "Display equation \"" + item.Eq.RawTail + "\" uses simple numbering, but this article uses sectional numbering (section.number) elsewhere — please check.",
                        Truncate(item.Para.Text, 60), item.Para.Start);
                    found++;
                }
            }

            return found;
        }

        // Generic increasing-sequence check that also understands lettered
        // sub-equations, e.g. 1, 1a, 1b, 2, 3a, 3b, 3c, 4 …
        private int CheckNumberLetterSequence(
            List<(int Number, char? SubLetter, ParaInfo Para)> items,
            Func<int, string> describePlain,
            Func<int, char, string> describeLettered,
            Func<int, string> expectPlain,
            Func<int, char, string> expectLettered)
        {
            int found = 0;
            int expectedNumber = 1;
            int? currentLetterRunNumber = null;
            char expectedNextLetter = 'a';

            foreach (var (number, subLetterRaw, para) in items)
            {
                if (subLetterRaw == null)
                {
                    if (number != expectedNumber)
                    {
                        Msg("HEADING", "ERROR",
                            describePlain(number) + " is out of sequence — expected " + expectPlain(expectedNumber) + ".",
                            Truncate(para.Text, 60), para.Start);
                        found++;
                    }
                    expectedNumber = number + 1;
                    currentLetterRunNumber = null;
                }
                else
                {
                    char letter = char.ToLowerInvariant(subLetterRaw.Value);
                    if (currentLetterRunNumber.HasValue && currentLetterRunNumber.Value == number)
                    {
                        // continuing an existing lettered run for the same number
                        if (letter != expectedNextLetter)
                        {
                            Msg("HEADING", "ERROR",
                                describeLettered(number, letter) + " is out of sequence — expected " + expectLettered(number, expectedNextLetter) + ".",
                                Truncate(para.Text, 60), para.Start);
                            found++;
                        }
                        expectedNextLetter = (char)(expectedNextLetter + 1);
                    }
                    else
                    {
                        // starting a new lettered run
                        if (number != expectedNumber)
                        {
                            Msg("HEADING", "ERROR",
                                describeLettered(number, letter) + " is out of sequence — expected " + expectLettered(expectedNumber, 'a') + ".",
                                Truncate(para.Text, 60), para.Start);
                            found++;
                        }
                        else if (letter != 'a')
                        {
                            Msg("HEADING", "ERROR",
                                describeLettered(number, letter) + " is out of sequence — expected " + expectLettered(number, 'a') + ".",
                                Truncate(para.Text, 60), para.Start);
                            found++;
                        }
                        currentLetterRunNumber = number;
                        expectedNextLetter = (char)(letter + 1);
                        expectedNumber = number + 1;
                    }
                }
            }
            return found;
        }

        // =====================================================================
        // OOXML — read Bold/Italic/SmallCaps/font-size
        // =====================================================================
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        internal sealed class FontInfo { internal bool Bold { get; set; } internal bool Italic { get; set; } internal bool SmallCaps { get; set; } internal double Pt { get; set; } }

        private Dictionary<string, double> ReadStyleFontSizes(string docPath)
        {
            var sizes = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var pkg = System.IO.Packaging.Package.Open(docPath, FileMode.Open, FileAccess.Read))
                {
                    var uri = new Uri("/word/styles.xml", UriKind.Relative);
                    if (!pkg.PartExists(uri)) return sizes;
                    XDocument stylesXml;
                    using (var stream = pkg.GetPart(uri).GetStream()) stylesXml = XDocument.Load(stream);
                    foreach (var styleEl in stylesXml.Descendants(W + "style"))
                    {
                        string name = (string)styleEl.Element(W + "name")?.Attribute(W + "val") ?? ""; if (string.IsNullOrEmpty(name)) continue;
                        var szEl = styleEl.Element(W + "rPr")?.Element(W + "sz"); if (szEl == null) continue;
                        if (double.TryParse((string)szEl.Attribute(W + "val") ?? "", out double half)) sizes[name] = half / 2.0;
                    }
                }
            }
            catch { }
            return sizes;
        }

        private Dictionary<int, FontInfo> ReadHeadingFormats(string docPath, List<ParaInfo> allParas)
        {
            var result = new Dictionary<int, FontInfo>();
            try
            {
                var styleSizes = ReadStyleFontSizes(docPath);
                XDocument bodyXml;
                using (var pkg = System.IO.Packaging.Package.Open(docPath, FileMode.Open, FileAccess.Read))
                {
                    var uri = new Uri("/word/document.xml", UriKind.Relative);
                    if (!pkg.PartExists(uri)) return result;
                    using (var stream = pkg.GetPart(uri).GetStream()) bodyXml = XDocument.Load(stream);
                }
                var lookup = allParas.Where(p => IsHeading(p) || p.SK == NK("abstract_title"))
                    .GroupBy(p => p.Text.Trim(), StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => (g.First().Start, g.First().SN), StringComparer.OrdinalIgnoreCase);
                foreach (var paraEl in bodyXml.Descendants(W + "p"))
                {
                    string oText = string.Concat(paraEl.Descendants(W + "t").Select(t => (string)t)).Trim();
                    if (!lookup.TryGetValue(oText, out var info)) continue;
                    long totalChars = 0, boldChars = 0, italicChars = 0, scChars = 0;
                    foreach (var run in paraEl.Elements(W + "r"))
                    {
                        string rt = string.Concat(run.Elements(W + "t").Select(t => (string)t));
                        int len = rt.Length; if (len == 0) continue;
                        var rpr = run.Element(W + "rPr"); totalChars += len;
                        if (IsOn(rpr?.Element(W + "b"))) boldChars += len;
                        if (IsOn(rpr?.Element(W + "i"))) italicChars += len;
                        if (IsOn(rpr?.Element(W + "smallCaps"))) scChars += len;
                    }
                    styleSizes.TryGetValue(info.SN, out double pt);
                    result[info.Start] = new FontInfo { Bold = totalChars > 0 && boldChars * 2 > totalChars, Italic = totalChars > 0 && italicChars * 2 > totalChars, SmallCaps = totalChars > 0 && scChars * 2 > totalChars, Pt = pt };
                }
            }
            catch { }
            return result;
        }

        private static bool IsOn(XElement el) { if (el == null) return false; string val = (string)el.Attribute(W + "val") ?? ""; return val != "0" && val != "false"; }

        // =====================================================================
        // ROMAN NUMERAL HELPERS
        // =====================================================================
        private static readonly Dictionary<char, int> RomanMap = new Dictionary<char, int> { { 'i', 1 }, { 'v', 5 }, { 'x', 10 }, { 'l', 50 }, { 'c', 100 }, { 'd', 500 }, { 'm', 1000 } };
        private static int RomanToInt(string s) { int total = 0, prev = 0; foreach (char ch in s.Reverse()) { if (!RomanMap.TryGetValue(ch, out int val)) return -1; total += val < prev ? -val : val; prev = val; } return total; }
        private static readonly int[] RVals = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
        private static readonly string[] RSyms = { "m", "cm", "d", "cd", "c", "xc", "l", "xl", "x", "ix", "v", "iv", "i" };
        private static string IntToRoman(int n) { var sb = new StringBuilder(); for (int i = 0; i < RVals.Length && n > 0; i++) while (n >= RVals[i]) { sb.Append(RSyms[i]); n -= RVals[i]; } return sb.ToString(); }

        // =====================================================================
        // HELPERS
        // =====================================================================
        private static List<string> DeduplicatedStyleKeys(IEnumerable<ParaInfo> paras) { var list = new List<string>(); string prev = null; foreach (var p in paras) if (p.SK != prev) { list.Add(p.SK); prev = p.SK; } return list; }
        private static int FirstPos(List<ParaInfo> paras, Func<ParaInfo, bool> pred) { for (int i = 0; i < paras.Count; i++) if (pred(paras[i])) return i; return -1; }
        private static int GetFirstStart(IEnumerable<ParaInfo> paras, string sk) { foreach (var p in paras) if (p.SK == sk) return p.Start; return 0; }
        private static string Truncate(string s, int max) => s == null ? "" : (s.Length <= max ? s : s.Substring(0, max) + "…");

        // =====================================================================
        // PARAGRAPH SNAPSHOT
        // =====================================================================
        internal sealed class ParaInfo
        {
            internal string SK { get; }
            internal string SN { get; }
            internal string Text { get; }
            internal int Start { get; }
            internal ParaInfo(string sk, string sn, string text, int start) { SK = sk; SN = sn; Text = text; Start = start; }
        }
    }
}