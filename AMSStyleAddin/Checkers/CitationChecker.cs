using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin.Checkers
{
    // =========================================================================
    // CITATION CHECKER
    //
    // Verifies that every reference in the reference list is cited at least
    // once in the body text, and that every in-text citation has a matching
    // entry in the reference list.
    //
    // ── AMS REFERENCE FORMATS ────────────────────────────────────────────────
    //   Personal name (1 author):    "Palmer, T., 2024: ..."
    //   Personal name (2 authors):   "Halem, M., and R. Jastrow, 1970: ..."
    //   Personal name (3+ / collab): "Chen, L., and Coauthors, 2024: ..."
    //                                "Bi, K., L. Xie, ..., 2023: ..."
    //   Org / collab (no comma):     "NOAA, 2025: ..."
    //                                "NOAA Heidke Skill Score Explanation, 2024: ..."
    //                                "NOAA/NWS/NCEP CPC, 2025: ..."
    //                                "U-Net, 2024: ..."
    //                                "C3S, 2018: ..."
    //
    // ── EXPECTED IN-TEXT CITATION FORMS ─────────────────────────────────────
    //   1-author ref    → Surname [YEAR]
    //   2-author ref    → Surname1 and Surname2 [YEAR]
    //   3+-author ref   → Surname et al. [YEAR]
    //   Org / collab    → Full Org Name [YEAR]          ← NO et al.
    //   Disambiguation  → J. Surname [YEAR]             (same-surname initials)
    //
    //   Hyperlink form: "[YEAR](N)" — the (N) hyperlink number is stripped first.
    //
    //   Sentence-style:     Author [YEAR]       (author before lone square bracket)
    //   Sentence-wrapped:   Author ([YEAR])     (author before bracketed year)
    //   Multi-year wrapped: Author ([Y1], [Y2]) (multiple years in round bracket)
    //   Round-bracket:      Author (YEAR)       (plain year in round bracket)
    //   Parenthetical:      (Author [YEAR]; Author [YEAR]; ...)
    //
    //   AMS "e.g./i.e./cf./see (also)" lead-ins:
    //       (e.g., Author YEAR; Author YEAR)
    //       (i.e., Author YEAR)
    //       (cf. Author YEAR)
    //       (see also Author YEAR)
    //
    //   Author-defined-abbreviation suffix:
    //       (Author1 and Author2 YEAR, ABBR)   ← trailing acronym after the year
    //
    // ── MATCH KEY ────────────────────────────────────────────────────────────
    //   surnameNormalised | year | yearSuffix
    //   Both the reference and the in-text citation are reduced to this key.
    //   "et al." in the citation is ignored — only the first-author/org
    //   surname matters for matching.
    //
    // ── CHANGELOG ────────────────────────────────────────────────────────────
    //   v1.1  Fixed three false-positive sources found when validating against
    //         a real AMS JCLI manuscript:
    //
    //         1. RxInnerCite now tolerates a leading "e.g.,", "i.e.,", "cf.",
    //            "see", or "see also" before the author name inside a
    //            parenthetical citation list. Previously the FIRST citation
    //            in "(e.g., Author1 YEAR; Author2 YEAR; ...)" was silently
    //            dropped because RxInnerCite is anchored with ^ and "e.g.,"
    //            does not match Init/First — causing spurious "reference
    //            never cited" errors for Author1. This is the single most
    //            common AMS citation pattern, so this fixed the large
    //            majority of false positives.
    //
    //         2. RxInnerCite now tolerates a trailing author-defined
    //            abbreviation after the year, e.g. "(Hung and Yanai 2004,
    //            HY04)". Previously the trailing ", HY04" broke the \s*$
    //            anchor and the whole citation was dropped, causing a
    //            spurious "reference never cited" error.
    //
    //         3. RxSentCite now excludes a stoplist of common sentence-
    //            initial discourse/transition words (e.g. "While",
    //            "Following", "Since") from the author-name capture via a
    //            negative lookahead. Previously, in sentences like
    //            "Following Sekizawa et al. (2023), we define..." or
    //            "While Back et al. (2025) focused...", the greedy
    //            multi-word First group absorbed the leading discourse word
    //            into the author name ("Following Sekizawa", "While Back"),
    //            producing a citation key that matches no reference and
    //            triggering a spurious "citation has no matching reference"
    //            error.
    // =========================================================================
    public class CitationChecker
    {
        // ── Issue counter (reset each Run) ────────────────────────────────────
        private int _issueNum;

        private void Msg(string cat, string level, string msg,
                         string snippet = "", int pos = 0)
        {
            _issueNum++;
            TaskPaneWinForms.AddMessage(cat, level, $"#{_issueNum} {msg}", snippet, pos);
        }

        // ── Paragraph styles ──────────────────────────────────────────────────
        // RefStyles: every paragraph style that marks a reference-list entry.
        // "collab_ref" / "collabref" covers org/dataset/software/web references
        // that AMS templates give their own style distinct from journal_ref etc.
        private static readonly HashSet<string> RefStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "journal_ref",  "book_ref",     "other_ref",
            "proceeding_ref","report_ref",  "thesis_ref",
            "collab_ref",   "dataset_ref",  "software_ref",
            "web_ref",      "tech_ref",
            // same names without underscores (NK-normalised forms)
            "journalref",   "bookref",      "otherref",
            "proceedingref","reportref",    "thesisref",
            "collabref",    "datasetref",   "softwareref",
            "webref",       "techref"
        };

        private static readonly HashSet<string> BodyStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "figurecaption","tablecaption","tablebody","tablehead",
            "paratext","acknowledgementtext","abstracttext","synopsis",
            "paranoindent","numberedlistitem","bulletedlistitem",
            "blockquot","formalarg","formalargend"
        };

        private static string NK(string s) =>
            (s ?? "").Replace(" ", "").Replace("-", "")
                     .Replace("_", "").ToLowerInvariant();

        // ── Regex safety timeout ──────────────────────────────────────────────
        private static readonly TimeSpan RxTimeout = TimeSpan.FromSeconds(2);

        // =====================================================================
        // SHARED REGEX FRAGMENTS
        // =====================================================================

        // Optional lead-in before an author name inside a parenthetical
        // citation segment: "e.g., ", "i.e., ", "cf. ", "see ", "see also ".
        // Consumed (and discarded) before Init/First so that the FIRST
        // citation after one of these lead-ins is no longer dropped.
        //   "(e.g., Ferranti et al. 1990; Higgins and Mo 1997)"
        //        → first segment "e.g., Ferranti et al. 1990" now matches
        private const string CiteLeadIn =
            @"(?:(?:e\.g\.|i\.e\.|cf\.|see(?:\s+also)?)\s*,?\s*)?";

        // Sentence-initial discourse/transition words that must NEVER be
        // absorbed into the author-name (First) capture of RxSentCite.
        // Without this, "Following Sekizawa et al. (2023)" or "While Back
        // et al. (2025)" would capture "Following Sekizawa" / "While Back"
        // as the author phrase, producing a citation key that matches no
        // reference.
        private const string SentenceLeadInStopWords =
            @"(?:While|Following|Since|Despite|Although|Because|Given|Before|After|" +
            @"Once|When|As|Unlike|Like|Until|Whereas|Though|However|Moreover|Thus|" +
            @"Hence|Therefore|Consequently|Recently|Previously|Here|There|Then|" +
            @"Notably|Importantly|Additionally|Furthermore|Conversely|Meanwhile)";

        // =====================================================================
        // REFERENCE-LIST REGEXES
        // =====================================================================

        // AMS year pattern in a reference paragraph: ", YYYY[a-z]? :"
        private static readonly Regex RxRefYear =
            new Regex(@",\s*(\d{4})([a-z]?)\s*:",
                RegexOptions.Compiled, RxTimeout);

        // "and Coauthors" or "et al." marks a 3+-author reference
        private static readonly Regex RxEtal =
            new Regex(@"\band\s+Coauthors\b|\bet\s+al\.",
                RegexOptions.Compiled | RegexOptions.IgnoreCase, RxTimeout);

        // Extracts co-author surnames to determine 2-author vs 3+-author count
        private static readonly Regex RxCoAuth =
            new Regex(
                @"(?:,\s*(?:and\s+)?)" +
                @"([A-Z][a-zA-Z\u00C0-\u024F\-]{1,40})" +
                @"(?=,|\s*\d{4}|\s*$|\.\s*(?:Eds|Jr|Sr))",
                RegexOptions.Compiled, RxTimeout);

        // =====================================================================
        // BODY-TEXT PRE-PROCESSING
        // =====================================================================

        // Word hyperlinks on "[2023]" leave a trailing "(N)" (1–3 digits) in the
        // text after field stripping: "Bi et al. [2023](3)" → "Bi et al. [2023]"
        // We strip all 1–3-digit parenthetical numbers before any citation parsing.
        private static readonly Regex RxHyperlinkNum =
            new Regex(@"\(\d{1,3}\)", RegexOptions.Compiled, RxTimeout);

        // =====================================================================
        // SENTENCE-STYLE CITATION REGEX
        //
        // Detects citations where the author/org name sits BEFORE the year bracket.
        //
        // ALL AMS sentence-citation forms (after hyperlink-number stripping):
        //
        //   Form A — lone square bracket:
        //       "Palmer [2024]"
        //       "Halem and Jastrow [1970]"
        //       "Bi et al. [2023]"
        //       "J. Smith [2020]"                              (disambiguation)
        //       "NOAA Heidke Skill Score Explanation [2024]"   (org name)
        //       "NOAA/NWS/NCEP CPC [2025]"                    (org with slashes)
        //       "C3S [2018]"                                   (org with digit)
        //
        //   Form B — round bracket wrapping one or more square-bracketed years:
        //       "Larraondo et al. ([2019])"
        //       "Taylor and Feng ([2022])"
        //       "Weyn et al. ([2019], [2020], [2021])"
        //
        //   Form C — round bracket with plain year:
        //       "Smith (2020)"
        //       "Smith et al. (2020)"
        //
        // ── DESIGN NOTES ─────────────────────────────────────────────────────
        //
        // Init   — optional initial prefix "J. " or "J.K. "
        // Stop   — negative lookahead excluding sentence-initial discourse
        //          words (see SentenceLeadInStopWords) from First. Without
        //          this, "Following Sekizawa et al. (2023)" would capture
        //          "Following Sekizawa" as the author phrase instead of just
        //          "Sekizawa". The lookahead causes the match attempt at
        //          "Following" to fail outright, so Regex.Matches retries at
        //          the next position and correctly starts the match at
        //          "Sekizawa" instead.
        // First  — first-author surname OR complete org/collab phrase.
        //          Each word in the phrase must start with an uppercase letter
        //          or digit. This prevents the regex from consuming ordinary
        //          prose words after a capitalised sentence subject.
        //          "/" is allowed inside a word (for "NOAA/NWS/NCEP").
        // Second — second author surname from " and Surname2"
        // Etal   — " et al." (OPTIONAL — collab/org names do NOT need et al.)
        // SqYear — year inside a lone [YYYY]
        // MqYears— content inside (  ) that holds one or more [YYYY] entries
        // RdYear — plain year(s) inside (  )
        // =====================================================================
        private static readonly Regex RxSentCite =
            new Regex(
                // Optional initial prefix: "J. " or "J.K. "
                @"(?<Init>(?:[A-Z]\.\s*){1,3})?" +

                // Reject sentence-initial discourse/transition words so they
                // are never absorbed into the author-name capture below.
                @"(?!" + SentenceLeadInStopWords + @"\s)" +

                // First author/org phrase.
                // ── Why uppercase-only additional words? ──────────────────────
                // Requiring each additional word to start uppercase (or digit)
                // stops the regex from greedily consuming ordinary prose words
                // that follow an uppercase sentence subject.  Without this,
                // "NOAA found that the model [2024]" would capture
                // "NOAA found that the model" as the author phrase.
                // With this restriction, only proper nouns / org acronyms
                // (e.g. "NOAA Heidke Skill Score Explanation") are captured.
                @"(?<First>[A-Z][A-Za-z0-9\u00C0-\u024F\-'/]*" +
                    @"(?:\s+[A-Z0-9][A-Za-z0-9\u00C0-\u024F\-'/]*)*)" +

                // Optional " and SecondSurname"  (two-author form)
                @"(?:\s+and\s+(?<Second>[A-Z][A-Za-z\u00C0-\u024F\-']+))?" +

                // Optional " et al."
                // IMPORTANT: this is OPTIONAL — org/collab names (NOAA, C3S, U-Net)
                // appear WITHOUT et al. in the body text and are matched correctly
                // because the First group already captured the full org name.
                @"(?<Etal>\s+et\s+al\.)?" +

                // Year bracket — three alternatives tried in order:
                @"\s+(?:" +
                    // Form A: lone [YYYY] or [YYYYa]
                    @"\[(?<SqYear>\d{4}[a-z]?)\]" +
                    @"|" +
                    // Form B: round bracket containing one or more [YYYY] entries
                    //   "([2019])"  or  "([2019], [2020], [2021])"
                    @"\((?<MqYears>(?:\[\d{4}[a-z]?\](?:,\s*)?)+)\)" +
                    @"|" +
                    // Form C: plain round bracket year list  "(2022)" or "(2019, 2020a)"
                    @"\((?<RdYear>\d{4}[a-z]?(?:,\s*\d{4}[a-z]?)*)\)" +
                @")",
                RegexOptions.Compiled, RxTimeout);

        // =====================================================================
        // PARENTHETICAL BLOCK REGEX  — CRITICAL DESIGN NOTE
        //
        // Uses \((.*?)\) — round brackets only, closing on the first ')'.
        //
        // WHY NOT [\(\[](.*?)[\)\]]?
        // That mixed-bracket form causes the non-greedy .*? to close at the
        // FIRST ']' or ')' it encounters.  Inside "(Palmer [2024])" the first
        // such character is the ']' from "[2024]", so inner = "Palmer [2024"
        // (missing the closing ']').  RxInnerCite then fails because it needs
        // the full "[2024]" to match \[\d{4}\], and the citation is silently
        // missed — producing false "reference never cited" errors.
        //
        // With \((.*?)\) on "(Palmer [2024])":
        //   \( matches '('
        //   .*? expands until the FIRST ')' — which is the outer closing paren
        //   inner = "Palmer [2024]"   ← correct, includes the full [2024] token
        //
        // The sentence-style parser (RxSentCite) ALSO detects these citations,
        // so both parsers together provide robust coverage.
        // =====================================================================
        private static readonly Regex RxParenBlock =
            new Regex(@"\((.*?)\)",
                RegexOptions.Compiled | RegexOptions.Singleline, RxTimeout);

        // =====================================================================
        // INNER-CITATION REGEX  (used for each ";" segment inside a paren block)
        //
        // Accepts both:
        //   "Palmer [2024]"       ← square bracket (standard inside a block)
        //   "Palmer 2024"         ← plain year (fallback)
        //
        // Groups mirror RxSentCite.  The \s*$ anchor ensures the year (plus an
        // optional trailing author-defined abbreviation, see below) is the
        // last token in the segment, preventing partial matches.
        //
        // ── LEAD-IN TOLERANCE ────────────────────────────────────────────────
        // CiteLeadIn optionally consumes "e.g., ", "i.e., ", "cf. ", "see ",
        // or "see also " before Init/First. Without this, splitting
        // "(e.g., Ferranti et al. 1990; Higgins and Mo 1997)" on ";" yields a
        // first segment "e.g., Ferranti et al. 1990" whose leading "e.g.,"
        // does not match Init/First, so the ^ anchor fails and the whole
        // segment — including the citation — is silently dropped.
        //
        // ── TRAILING-ABBREVIATION TOLERANCE ─────────────────────────────────
        // The optional "(?:,\s*[A-Z][A-Za-z0-9]*)?" before \s*$ allows a
        // trailing author-defined abbreviation after the year, e.g.
        // "(Hung and Yanai 2004, HY04)" → year "2004" followed by ", HY04".
        // Without this, the trailing ", HY04" prevents \s*$ from matching and
        // the whole citation is dropped.
        // =====================================================================
        private static readonly Regex RxInnerCite =
            new Regex(
                @"^\s*" +
                CiteLeadIn +
                @"(?<Init>(?:[A-Z]\.\s*){1,3})?" +
                @"(?<First>[A-Z][A-Za-z0-9\u00C0-\u024F\-'/]*" +
                    @"(?:\s+[A-Z0-9][A-Za-z0-9\u00C0-\u024F\-'/]*)*)" +
                @"(?:\s+and\s+(?<Second>[A-Z][A-Za-z\u00C0-\u024F\-']+))?" +
                @"(?:\s+et\s+al\.)?" +
                @"\s+(?:\[(?<SqYear>\d{4}[a-z]?)\]|(?<PlYear>\d{4}[a-z]?))" +
                @"(?:,\s*[A-Z][A-Za-z0-9]*)?" +
                @"\s*$",
                RegexOptions.Compiled, RxTimeout);

        // ── Year token extractor ──────────────────────────────────────────────
        private static readonly Regex RxYearToken =
            new Regex(@"(\d{4})([a-z])?", RegexOptions.Compiled, RxTimeout);

        // =====================================================================
        // DATA STRUCTURES
        // =====================================================================
        private class RefEntry
        {
            public string SurnameRaw;   // Display surname / org name as in document
            public string SurnameKey;   // Unicode-normalised ASCII sort key
            public string InitialRaw;   // "J." or "J.A." or "" (org refs = "")
            public string InitialKey;   // Normalised initial letter(s)
            public int AuthorCount;  // 1 = single / org, 2 = two named, 3 = et al.
            public int Year;
            public string YearSuffix;   // "a","b",… or ""
            public string RefKey;       // surnameKey|year|suffix
            public int ParaStart;
            public string Excerpt;
        }

        private class CiteEntry
        {
            public string SurnameRaw;
            public string SurnameKey;
            public string InitialKey;   // "" when no initial given in the citation
            public int Year;
            public string YearSuffix;
            public string CiteKey;      // surnameKey|year|suffix
            public int CharPos;
            public string Excerpt;
        }

        // =====================================================================
        // UNICODE / DIACRITIC NORMALISATION  (identical to ReferenceReorderingChecker)
        // =====================================================================
        private static readonly Dictionary<char, string> ManualMap =
            new Dictionary<char, string>
        {
            {'Ø',"o"},{'ø',"o"},{'Æ',"ae"},{'æ',"ae"},
            {'Þ',"th"},{'þ',"th"},{'Ð',"d"},{'ð',"d"},
            {'Ŋ',"n"},{'ŋ',"n"},{'ß',"ss"},{'ẞ',"ss"},
            {'Ł',"l"},{'ł',"l"},{'Œ',"oe"},{'œ',"oe"},
            {'Ĳ',"ij"},{'ĳ',"ij"},
        };

        private static string ToSortKey(string raw)
        {
            if (string.IsNullOrEmpty(raw)) return "";
            var sb = new StringBuilder(raw.Length + 4);
            foreach (char c in raw)
            {
                if (ManualMap.TryGetValue(c, out string mapped)) { sb.Append(mapped); continue; }
                string nfd = c.ToString().Normalize(NormalizationForm.FormD);
                foreach (char nc in nfd)
                {
                    if (CharUnicodeInfo.GetUnicodeCategory(nc) ==
                        UnicodeCategory.NonSpacingMark) continue;
                    if (nc >= 'a' && nc <= 'z') { sb.Append(nc); continue; }
                    if (nc >= 'A' && nc <= 'Z') { sb.Append(char.ToLowerInvariant(nc)); continue; }
                    if (nc >= '0' && nc <= '9') { sb.Append(nc); continue; }
                }
            }
            return sb.ToString();
        }

        private static string MakeKey(string surnameKey, int year, string suffix) =>
            $"{surnameKey}|{year}|{suffix ?? ""}";

        // =====================================================================
        // RUN
        // =====================================================================
        public void Run()
        {
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch { return; }

            _issueNum = 0;
            int totalParas = 0;
            try { totalParas = doc.Paragraphs.Count; } catch { }

            var refEntries = new List<RefEntry>();
            var bodyParas = new List<(string Text, int Start)>();

            bool screenWasOn = true;
            int paraIndex = 0;

            // ── Pass 1: collect reference and body paragraphs ─────────────────
            try
            {
                screenWasOn = doc.Application.ScreenUpdating;
                doc.Application.ScreenUpdating = false;

                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    paraIndex++;
                    if (paraIndex % 50 == 0)
                    {
                        TaskPaneWinForms.SetProgress(totalParas > 0
                            ? $"Scanning paragraph {paraIndex} of {totalParas}…"
                            : $"Scanning paragraph {paraIndex}…");
                        doc.Application.ScreenUpdating = screenWasOn;
                        Thread.Sleep(0);
                        doc.Application.ScreenUpdating = false;
                    }

                    string styleKey;
                    try { styleKey = NK(para.get_Style().NameLocal); }
                    catch { continue; }

                    string rawText;
                    int paraStart;
                    try
                    {
                        rawText = (para.Range.Text ?? "").TrimEnd('\r', '\n');
                        paraStart = para.Range.Start;
                    }
                    catch { continue; }

                    // Strip Word field codes
                    rawText = Regex.Replace(rawText, @"\x13[^\x14]*\x14", "");
                    rawText = rawText.Replace("\x15", "");
                    rawText = Regex.Replace(rawText, @"  +", " ").Trim();

                    if (string.IsNullOrWhiteSpace(rawText)) continue;

                    if (RefStyles.Contains(styleKey) || styleKey.EndsWith("ref"))
                    {
                        RefEntry re = ParseReference(rawText, paraStart);
                        if (re != null) refEntries.Add(re);
                    }
                    else if (BodyStyles.Contains(styleKey))
                    {
                        // Strip hyperlink reference numbers before parsing
                        string bodyText = rawText;
                        try { bodyText = RxHyperlinkNum.Replace(rawText, ""); }
                        catch (RegexMatchTimeoutException) { }
                        bodyParas.Add((bodyText, paraStart));
                    }
                }
            }
            finally { try { doc.Application.ScreenUpdating = screenWasOn; } catch { } }

            if (refEntries.Count == 0 && bodyParas.Count == 0)
            {
                Msg("CITATION", "WARNING",
                    "No reference or body paragraphs found. " +
                    "Make sure paragraphs use AMS styles (journal_ref, paratext, etc.).");
                return;
            }

            TaskPaneWinForms.SetProgress("Matching citations…");

            // ── Build reference index ─────────────────────────────────────────
            var refByKey = new Dictionary<string, RefEntry>(StringComparer.OrdinalIgnoreCase);
            var surnameInitials = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            foreach (var re in refEntries)
            {
                if (!refByKey.ContainsKey(re.RefKey))
                    refByKey[re.RefKey] = re;

                if (!surnameInitials.ContainsKey(re.SurnameKey))
                    surnameInitials[re.SurnameKey] =
                        new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                surnameInitials[re.SurnameKey].Add(re.InitialKey);
            }

            // ── Pass 2: extract in-text citations ─────────────────────────────
            var citeEntries = new List<CiteEntry>();
            foreach (var (text, start) in bodyParas)
            {
                citeEntries.AddRange(ParseSentenceCitations(text, start));
                citeEntries.AddRange(ParseParentheticalCitations(text, start));
            }

            // Group by key (duplicates from both parsers are harmless)
            var citeByKey = new Dictionary<string, List<CiteEntry>>(StringComparer.OrdinalIgnoreCase);
            foreach (var ce in citeEntries)
            {
                if (!citeByKey.ContainsKey(ce.CiteKey))
                    citeByKey[ce.CiteKey] = new List<CiteEntry>();
                citeByKey[ce.CiteKey].Add(ce);
            }

            int errorCount = 0;

            // ── Rule 1: Reference present but never cited ─────────────────────
            foreach (var re in refEntries)
            {
                if (citeByKey.ContainsKey(re.RefKey)) continue;
                Msg("CITATION", "ERROR",
                    $"Reference \"{re.SurnameRaw} {re.Year}{re.YearSuffix}\" " +
                    $"is never cited in the body text.",
                    re.Excerpt, re.ParaStart);
                errorCount++;
            }

            // ── Rule 2: Citation has no matching reference ─────────────────────
            foreach (var kvp in citeByKey)
            {
                if (refByKey.ContainsKey(kvp.Key)) continue;
                var ce = kvp.Value[0];
                Msg("CITATION", "ERROR",
                    $"Citation \"{ce.SurnameRaw} {ce.Year}{ce.YearSuffix}\" " +
                    $"has no matching reference in the reference list.",
                    ce.Excerpt, ce.CharPos);
                errorCount++;
            }

            // ── Rule 3: Ambiguous surname (initial required but not given) ─────
            foreach (var kvp in citeByKey)
            {
                var ce = kvp.Value[0];
                if (!string.IsNullOrEmpty(ce.InitialKey)) continue;
                if (!surnameInitials.TryGetValue(ce.SurnameKey, out var initSet)) continue;
                if (initSet.Count < 2) continue;
                Msg("CITATION", "ERROR",
                    $"Citation \"{ce.SurnameRaw} {ce.Year}{ce.YearSuffix}\" is ambiguous — " +
                    $"multiple authors named \"{ce.SurnameRaw}\" exist in the reference list. " +
                    $"Add an initial to disambiguate " +
                    $"(e.g. \"J. {ce.SurnameRaw} {ce.Year}{ce.YearSuffix}\").",
                    ce.Excerpt, ce.CharPos);
                errorCount++;
            }

            if (errorCount == 0)
                TaskPaneWinForms.AddMessage("CITATION", "INFO",
                    $"Citation check passed — {refEntries.Count} reference(s) all cited " +
                    "and all in-text citations have matching references.");
        }

        // =====================================================================
        // PARSE ONE REFERENCE PARAGRAPH
        //
        // Personal-name refs:
        //   "Bi, K., L. Xie, ..., 2023: ..."       → surname="Bi", initial="K", count=3
        //   "Halem, M., and R. Jastrow, 1970: ..."  → surname="Halem", initial="M", count=2
        //   "Chen, L., and Coauthors, 2024: ..."    → surname="Chen", initial="L", count=3
        //   "Palmer, T., 2024: ..."                 → surname="Palmer", initial="T", count=1
        //
        // Org / collab refs (no comma before the year-comma, so fc = -1):
        //   "NOAA, 2025: ..."                       → surname="NOAA", initial="", count=1
        //   "NOAA Heidke Skill Score Explanation, 2024: ..."  → surname=full org name
        //   "NOAA/NWS/NCEP CPC, 2025: ..."          → surname="NOAA/NWS/NCEP CPC"
        //   "U-Net, 2024: ..."                      → surname="U-Net"
        //   "C3S, 2018: ..."                        → surname="C3S"
        //
        // For org refs the author block has no comma, so fc = -1 and the entire
        // author block is treated as the surname.  No initial is extracted.
        // =====================================================================
        private static RefEntry ParseReference(string text, int paraStart)
        {
            Match mYr;
            try { mYr = RxRefYear.Match(text); }
            catch (RegexMatchTimeoutException) { return null; }
            if (!mYr.Success) return null;

            int year = int.Parse(mYr.Groups[1].Value);
            string yearSuffix = mYr.Groups[2].Value;
            string authorBlock = text.Substring(0, mYr.Index).Trim();

            // First comma separates surname from initials in personal-name refs.
            // Org refs have no comma in the author block, so fc = -1.
            int fc = authorBlock.IndexOf(',');
            string surnameRaw = fc < 0 ? authorBlock.Trim()
                                       : authorBlock.Substring(0, fc).Trim();

            string surnameKey = ToSortKey(surnameRaw);
            if (string.IsNullOrEmpty(surnameKey)) return null;

            // ── Initial (personal-name refs only) ─────────────────────────────
            string initialRaw = "";
            string initialKey = "";

            if (fc >= 0)
            {
                string afterComma = authorBlock.Substring(fc + 1).TrimStart();
                // Initial: uppercase letter immediately followed by a dot
                if (afterComma.Length >= 2 &&
                    (char.IsUpper(afterComma[0]) ||
                     CharUnicodeInfo.GetUnicodeCategory(afterComma[0]) ==
                         UnicodeCategory.UppercaseLetter) &&
                    afterComma[1] == '.')
                {
                    var rb = new StringBuilder();
                    var kb = new StringBuilder();
                    int ci = 0;
                    while (ci < afterComma.Length)
                    {
                        char ch = afterComma[ci];
                        if (ch == ' ' || ch == '\u00A0') { ci++; continue; }
                        if ((char.IsUpper(ch) ||
                             CharUnicodeInfo.GetUnicodeCategory(ch) ==
                                 UnicodeCategory.UppercaseLetter) &&
                            ci + 1 < afterComma.Length && afterComma[ci + 1] == '.')
                        {
                            kb.Append(ToSortKey(ch.ToString()));
                            rb.Append(ch);
                            ci++;
                            while (ci < afterComma.Length &&
                                   (afterComma[ci] == '.' || afterComma[ci] == '-' ||
                                    afterComma[ci] == '\u2013'))
                            { rb.Append(afterComma[ci]); ci++; }
                        }
                        else break;
                    }
                    initialRaw = rb.ToString().Trim();
                    initialKey = kb.ToString();
                }
            }

            // ── Author count ──────────────────────────────────────────────────
            // 3 = et al. or "and Coauthors" (or 3+ named authors)
            // 2 = exactly two named authors
            // 1 = single author or any org/collab name
            bool hasEtal;
            try { hasEtal = RxEtal.IsMatch(authorBlock); }
            catch (RegexMatchTimeoutException) { hasEtal = false; }

            int authorCount = 1;

            if (hasEtal)
            {
                authorCount = 3;
            }
            else if (fc >= 0)
            {
                try
                {
                    var coSurnames = new List<string>();
                    foreach (Match cm in RxCoAuth.Matches(authorBlock))
                    {
                        string cs = cm.Groups[1].Value.Trim();
                        string csL = cs.ToLowerInvariant().TrimEnd('.');
                        if (csL != "jr" && csL != "sr" && csL != "ed" &&
                            csL != "eds" && csL != "coauthors" && cs.Length > 1)
                            coSurnames.Add(cs);
                    }
                    authorCount = Math.Min(1 + coSurnames.Count, 3);
                }
                catch (RegexMatchTimeoutException) { }
            }
            // If fc < 0 (org/collab): authorCount stays 1 — no et al. for org names.

            return new RefEntry
            {
                SurnameRaw = surnameRaw,
                SurnameKey = surnameKey,
                InitialRaw = initialRaw,
                InitialKey = initialKey,
                AuthorCount = authorCount,
                Year = year,
                YearSuffix = yearSuffix,
                RefKey = MakeKey(surnameKey, year, yearSuffix),
                ParaStart = paraStart,
                Excerpt = Truncate(text, 70)
            };
        }

        // =====================================================================
        // PARSE SENTENCE-STYLE CITATIONS
        //
        // Scans body text for citations where the author/org name is BEFORE
        // the year bracket (all forms described in RxSentCite above).
        //
        // The First group captures either:
        //   • A personal surname ("Palmer", "Bi", "Grönquist")
        //   • A complete org/collab phrase ("NOAA Heidke Skill Score Explanation",
        //     "NOAA/NWS/NCEP CPC", "C3S", "U-Net")
        //
        // For ORG NAMES: there is NO "et al." in the citation.  The regex
        // handles this correctly because ?<Etal> is optional — the match
        // succeeds with or without et al. present.
        //
        // A leading sentence-initial discourse word (e.g. "Following",
        // "While") is excluded from First by RxSentCite's negative lookahead,
        // so "Following Sekizawa et al. (2023)" correctly yields the author
        // phrase "Sekizawa", not "Following Sekizawa".
        //
        // The match key is built from ToSortKey(First), year, yearSuffix.
        // This equals the RefKey built from the reference paragraph, so
        // matching just uses string comparison on the key.
        // =====================================================================
        private static List<CiteEntry> ParseSentenceCitations(string text, int paraStart)
        {
            var result = new List<CiteEntry>();

            IEnumerable<Match> matches;
            try { matches = RxSentCite.Matches(text).Cast<Match>(); }
            catch (RegexMatchTimeoutException) { return result; }

            foreach (Match m in matches)
            {
                string authorPhrase = m.Groups["First"].Value.Trim();
                string initialRaw = m.Groups["Init"].Value.Trim();

                if (string.IsNullOrEmpty(authorPhrase)) continue;

                string surnameKey = ToSortKey(authorPhrase);
                if (string.IsNullOrEmpty(surnameKey)) continue;

                string initialKey = string.IsNullOrEmpty(initialRaw)
                    ? ""
                    : ToSortKey(initialRaw.Replace(".", "").Replace(" ", ""));
                if (initialKey.Length > 1) initialKey = initialKey.Substring(0, 1);

                string yearStr =
                    m.Groups["SqYear"].Success ? m.Groups["SqYear"].Value :
                    m.Groups["MqYears"].Success ? m.Groups["MqYears"].Value :
                    m.Groups["RdYear"].Value;

                foreach (var (year, suffix) in ExtractYears(yearStr))
                {
                    result.Add(new CiteEntry
                    {
                        SurnameRaw = authorPhrase,
                        SurnameKey = surnameKey,
                        InitialKey = initialKey,
                        Year = year,
                        YearSuffix = suffix,
                        CiteKey = MakeKey(surnameKey, year, suffix),
                        CharPos = paraStart + m.Index,
                        Excerpt = Truncate(m.Value, 60)
                    });
                }
            }

            return result;
        }

        // =====================================================================
        // PARSE PARENTHETICAL CITATIONS
        //
        // Finds (…) blocks and splits on ";" to get individual citations.
        //
        // CRITICAL: uses \((.*?)\) so the first ')' properly closes the outer
        // block.  This ensures inner content for "(Palmer [2024])" is the
        // full "Palmer [2024]" string (not "Palmer [2024" with a missing ']').
        //
        // Each segment is matched by RxInnerCite, which accepts:
        //   "Author [YEAR]"                — square bracket (AMS standard inside a block)
        //   "Author YEAR"                  — plain year (fallback for older styles)
        //   "e.g., Author YEAR"             — common AMS lead-in (see CiteLeadIn)
        //   "i.e., Author YEAR"             — likewise
        //   "cf. Author YEAR"               — likewise
        //   "see [also] Author YEAR"        — likewise
        //   "Author1 and Author2 YEAR, ABBR" — trailing author-defined abbreviation
        //
        // Non-citation content (figure labels, equations, year ranges) is
        // filtered by the 4-digit-year pre-check and by RxInnerCite failing
        // (which requires an uppercase-initial author name before the year).
        // =====================================================================
        private static List<CiteEntry> ParseParentheticalCitations(
            string text, int paraStart)
        {
            var result = new List<CiteEntry>();

            IEnumerable<Match> blocks;
            try { blocks = RxParenBlock.Matches(text).Cast<Match>(); }
            catch (RegexMatchTimeoutException) { return result; }

            foreach (Match block in blocks)
            {
                string inner = block.Groups[1].Value;
                int blkPos = paraStart + block.Index;

                // Pre-filter: skip blocks with no 4-digit year
                if (!Regex.IsMatch(inner, @"\d{4}")) continue;

                string blockExcerpt = Truncate(block.Value, 60);

                foreach (string segment in inner.Split(';'))
                {
                    string seg = segment.Trim();
                    if (string.IsNullOrEmpty(seg)) continue;

                    Match m;
                    try { m = RxInnerCite.Match(seg); }
                    catch (RegexMatchTimeoutException) { continue; }
                    if (!m.Success) continue;

                    string authorPhrase = m.Groups["First"].Value.Trim();
                    string initialRaw = m.Groups["Init"].Value.Trim();

                    if (string.IsNullOrEmpty(authorPhrase)) continue;

                    string surnameKey = ToSortKey(authorPhrase);
                    if (string.IsNullOrEmpty(surnameKey)) continue;

                    string initialKey = string.IsNullOrEmpty(initialRaw)
                        ? ""
                        : ToSortKey(initialRaw.Replace(".", "").Replace(" ", ""));
                    if (initialKey.Length > 1) initialKey = initialKey.Substring(0, 1);

                    string yearStr = m.Groups["SqYear"].Success
                        ? m.Groups["SqYear"].Value
                        : m.Groups["PlYear"].Value;

                    foreach (var (year, suffix) in ExtractYears(yearStr))
                    {
                        result.Add(new CiteEntry
                        {
                            SurnameRaw = authorPhrase,
                            SurnameKey = surnameKey,
                            InitialKey = initialKey,
                            Year = year,
                            YearSuffix = suffix,
                            CiteKey = MakeKey(surnameKey, year, suffix),
                            CharPos = blkPos,
                            Excerpt = blockExcerpt
                        });
                    }
                }
            }

            return result;
        }

        // =====================================================================
        // EXTRACT YEAR TOKENS FROM A RAW YEAR STRING
        //
        // Handles all forms that appear in the year groups above:
        //   "2022"              → [(2022, "")]
        //   "2022a"             → [(2022, "a")]
        //   "2019, 2020"        → [(2019, ""), (2020, "")]
        //   "[2019], [2020]"    → [(2019, ""), (2020, "")]  (Form B MqYears)
        // =====================================================================
        private static List<(int Year, string Suffix)> ExtractYears(string raw)
        {
            var result = new List<(int, string)>();
            if (string.IsNullOrWhiteSpace(raw)) return result;
            IEnumerable<Match> matches;
            try { matches = RxYearToken.Matches(raw).Cast<Match>(); }
            catch (RegexMatchTimeoutException) { return result; }
            foreach (Match m in matches)
            {
                result.Add((
                    int.Parse(m.Groups[1].Value),
                    m.Groups[2].Success ? m.Groups[2].Value : ""
                ));
            }
            return result;
        }

        private static string Truncate(string s, int max) =>
            s == null ? "" : s.Length <= max ? s : s.Substring(0, max) + "…";
    }
}