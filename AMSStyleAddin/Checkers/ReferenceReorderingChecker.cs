using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin.Checkers
{
    public class ReferenceReorderingChecker
    {
        private int _issueNum;

        private void Msg(string cat, string level, string msg, string snippet = "", int start = 0)
        {
            _issueNum++;
            TaskPaneWinForms.AddMessage(cat, level, $"#{_issueNum} {msg}", snippet, start);
        }

        // ── Reference paragraph styles ────────────────────────────────────────
        private static readonly HashSet<string> RefStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "journal_ref","book_ref","other_ref",
            "proceeding_ref","report_ref","thesis_ref",
            "journalref","bookref","otherref",
            "proceedingref","reportref","thesisref"
        };

        private static string NK(string s) =>
            (s ?? "").Replace(" ", "").Replace("-", "")
                     .Replace("_", "").ToLowerInvariant();

        private static readonly TimeSpan RxTimeout = TimeSpan.FromSeconds(2);

        private static readonly Regex RxYear =
            new Regex(@",\s*(\d{4})([a-z]?)\s*:",
                RegexOptions.Compiled, RxTimeout);

        private static readonly Regex RxEtal =
            new Regex(@"\band\s+Coauthors\b|\bet\s+al\.",
                RegexOptions.Compiled | RegexOptions.IgnoreCase, RxTimeout);

        // ── Surname detector in co-author block ───────────────────────────────
        private static readonly Regex RxSurnameWord =
            new Regex(@"\b([A-Z\u00C0-\u024F][a-zA-Z\u00C0-\u024F\-]{1,40})\b",
                RegexOptions.Compiled, RxTimeout);

        private class RefEntry
        {
            public string Surname;
            public string SurnameRaw;
            public string Initial;
            public string InitialRaw;
            public int AuthorCount;
            public int Year;
            public string YearSuffix;
            public string Label;
            public int ParaStart;
            public string Excerpt;
        }

        // =====================================================================
        // MANUAL CHARACTER MAP — chars that don't NFD-decompose
        // =====================================================================
        private static readonly Dictionary<char, string> ManualMap =
            new Dictionary<char, string>
            {
                { 'Ø', "o"  }, { 'ø', "o"  },
                { 'Æ', "ae" }, { 'æ', "ae" },
                { 'Þ', "th" }, { 'þ', "th" },
                { 'Ð', "d"  }, { 'ð', "d"  },
                { 'Ŋ', "n"  }, { 'ŋ', "n"  },
                { 'ß', "ss" }, { 'ẞ', "ss" },
                { 'Ł', "l"  }, { 'ł', "l"  },
                { 'Œ', "oe" }, { 'œ', "oe" },
                { 'Ĳ', "ij" }, { 'ĳ', "ij" },
            };

        // =====================================================================
        // TO SORT KEY — Unicode → plain ASCII for comparison
        // =====================================================================
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
                    if (CharUnicodeInfo.GetUnicodeCategory(nc) == UnicodeCategory.NonSpacingMark) continue;
                    if (nc >= 'a' && nc <= 'z') { sb.Append(nc); continue; }
                    if (nc >= 'A' && nc <= 'Z') { sb.Append(char.ToLowerInvariant(nc)); continue; }
                }
            }
            return sb.ToString();
        }

        // =====================================================================
        // COUNT AUTHORS
        //
        // FIX 1: "et al." / "and Coauthors" is now checked TWICE:
        //   (a) via the pre-computed hasEtal flag (passed in from Run)
        //   (b) directly on the raw authorBlock as a secondary guard,
        //       in case hasEtal was evaluated on a differently-trimmed slice.
        //
        // If either check fires → return 3 immediately, no further parsing.
        //
        // FIX 2: the co-author counting fallback is unchanged but now only
        // reached when etal is genuinely absent.
        //
        // Handles both AMS author formats:
        //   Format A: "Smith, A. B., C. D. Jones, and E. F. Brown"  → 3
        //   Format B: "Smith, A. B., Jones, C. D., Brown, E. F."    → 3
        //   Single  : "Smith, A."                                    → 1
        //   Pair    : "Smith, A., and B. Jones"                      → 2
        // =====================================================================
        private static int CountAuthors(string authorBlock, bool hasEtal)
        {
            // Primary check: caller already ran RxEtal on the author slice
            if (hasEtal) return 3;

            // Secondary guard: re-check directly on the raw block.
            // Catches cases where the pre-computed flag was based on a
            // differently-trimmed string (e.g. trailing punctuation included).
            try
            {
                if (RxEtal.IsMatch(authorBlock)) return 3;
            }
            catch (RegexMatchTimeoutException)
            {
                // If the regex times out we conservatively treat it as etal
                return 3;
            }

            // --- No etal: count authors by parsing the block manually ----------
            int fc = authorBlock.IndexOf(',');
            if (fc < 0) return 1;          // only one token → single author

            string rest = authorBlock.Substring(fc + 1).Trim();
            if (string.IsNullOrWhiteSpace(rest)) return 1;

            // Split by comma and " and "
            string[] parts = Regex.Split(rest, @",|\band\b", RegexOptions.IgnoreCase);
            int coAuthors = 0;

            foreach (string part in parts)
            {
                string p = part.Trim();
                if (string.IsNullOrWhiteSpace(p)) continue;

                // A surname token: uppercase first letter, >1 alpha char,
                // not a bare initial like "A."
                string[] words = p.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string w in words)
                {
                    string clean = Regex.Replace(w, @"[^a-zA-Z\u00C0-\u024F]", "");
                    if (clean.Length > 1 && char.IsUpper(clean[0]))
                    {
                        coAuthors++;
                        break; // one surname per comma-delimited part
                    }
                }
            }

            return Math.Min(1 + coAuthors, 3);
        }

        // =====================================================================
        // RUN
        // =====================================================================
        public void Run()
        {
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch { return; }

            _issueNum = 0;

            var refs = new List<RefEntry>();

            bool screenWasOn = true;
            int paraCount = 0;
            try
            {
                screenWasOn = doc.Application.ScreenUpdating;
                doc.Application.ScreenUpdating = false;

                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    paraCount++;
                    if (paraCount % 50 == 0)
                    {
                        doc.Application.ScreenUpdating = screenWasOn;
                        Thread.Sleep(0);
                        doc.Application.ScreenUpdating = false;
                    }

                    string styleKey = "";
                    try { styleKey = NK(para.get_Style().NameLocal); }
                    catch { continue; }

                    bool isRef = false;
                    foreach (var rs in RefStyles)
                        if (styleKey == NK(rs)) { isRef = true; break; }
                    if (!isRef) continue;

                    string text = "";
                    try { text = para.Range.Text ?? ""; }
                    catch { continue; }
                    text = text.TrimEnd('\r', '\n').Trim();
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    int paraStart = 0;
                    try { paraStart = para.Range.Start; } catch { }

                    Match mYr;
                    try { mYr = RxYear.Match(text); }
                    catch (RegexMatchTimeoutException) { continue; }
                    if (!mYr.Success) continue;

                    int year = int.Parse(mYr.Groups[1].Value);
                    string yearSuffix = mYr.Groups[2].Value;
                    string authorBlock = text.Substring(0, mYr.Index).Trim();

                    bool hasEtal;
                    try { hasEtal = RxEtal.IsMatch(authorBlock); }
                    catch (RegexMatchTimeoutException) { hasEtal = false; }

                    int fc = authorBlock.IndexOf(',');
                    if (fc < 0) continue;

                    string surnameRaw = authorBlock.Substring(0, fc).Trim();
                    string surname = ToSortKey(surnameRaw);
                    if (string.IsNullOrEmpty(surname)) continue;

                    // ── Extract initial ───────────────────────────────────────
                    string initialRaw = "";
                    string initial = "";
                    {
                        string afterComma = authorBlock.Substring(fc + 1).TrimStart();
                        var rawBuilder = new StringBuilder();
                        var keyBuilder = new StringBuilder();
                        int ci = 0;
                        while (ci < afterComma.Length)
                        {
                            char ch = afterComma[ci];
                            if (ch == ' ' || ch == '\u00A0') { ci++; continue; }
                            if (char.IsUpper(ch) ||
                                CharUnicodeInfo.GetUnicodeCategory(ch) == UnicodeCategory.UppercaseLetter)
                            {
                                keyBuilder.Append(ToSortKey(ch.ToString()));
                                rawBuilder.Append(ch);
                                ci++;
                                while (ci < afterComma.Length &&
                                       (afterComma[ci] == '.' || afterComma[ci] == '-' ||
                                        afterComma[ci] == '\u2013' || afterComma[ci] == '\u2014'))
                                { rawBuilder.Append(afterComma[ci]); ci++; }
                            }
                            else if (ch == ',') break;
                            else break;
                        }
                        initialRaw = rawBuilder.ToString().Trim();
                        initial = keyBuilder.ToString();
                    }

                    // ── Author count ──────────────────────────────────────────
                    // Note: hasEtal is also re-checked inside CountAuthors as a
                    // secondary guard, so "et al." / "and Coauthors" is never
                    // missed regardless of how this flag is computed above.
                    int authorCount = CountAuthors(authorBlock, hasEtal);

                    // ── Label ─────────────────────────────────────────────────
                    string label;
                    if (authorCount == 1)
                        label = surnameRaw + " " + year + yearSuffix;
                    else if (authorCount == 2)
                        label = surnameRaw + " and ... " + year + yearSuffix;
                    else
                        label = surnameRaw + " et al. " + year + yearSuffix;

                    refs.Add(new RefEntry
                    {
                        Surname = surname,
                        SurnameRaw = surnameRaw,
                        Initial = initial,
                        InitialRaw = initialRaw,
                        AuthorCount = authorCount,
                        Year = year,
                        YearSuffix = yearSuffix,
                        Label = label,
                        ParaStart = paraStart,
                        Excerpt = Truncate(text, 70)
                    });
                }
            }
            finally
            {
                try { doc.Application.ScreenUpdating = screenWasOn; } catch { }
            }

            if (refs.Count == 0)
            {
                Msg("REF-ORD", "WARNING",
                    "No references found. Make sure reference paragraphs use " +
                    "styles: journal_ref, book_ref, other_ref, " +
                    "proceeding_ref, report_ref, thesis_ref.");
                return;
            }

            // =====================================================================
            // PASS 1 — Consecutive-pair checks
            // Covers: wrong order, same-surname orange warnings,
            //         same-surname+same-year+diff-initial orange warning.
            // These only make sense on adjacent entries because the list is
            // sorted; a disorder between non-adjacent entries is always caught
            // by a closer pair.
            // =====================================================================
            int issueCount = 0;

            for (int i = 0; i < refs.Count - 1; i++)
            {
                var a = refs[i];
                var b = refs[i + 1];

                // AMS sort: Surname → Initial → AuthorCount → Year → YearSuffix
                int cmp = string.Compare(a.Surname, b.Surname, StringComparison.OrdinalIgnoreCase);
                if (cmp == 0) cmp = string.Compare(a.Initial, b.Initial, StringComparison.OrdinalIgnoreCase);
                if (cmp == 0) cmp = a.AuthorCount.CompareTo(b.AuthorCount);
                if (cmp == 0) cmp = a.Year.CompareTo(b.Year);
                if (cmp == 0) cmp = string.Compare(a.YearSuffix, b.YearSuffix, StringComparison.OrdinalIgnoreCase);

                // ── Wrong order ───────────────────────────────────────────────
                if (cmp > 0)
                {
                    Msg("REF-ORD", "ERROR",
                        $"Wrong order: \"{a.Label}\" should come AFTER \"{b.Label}\".",
                        a.Excerpt, a.ParaStart);
                    issueCount++;
                }

                bool sameSurname = string.Equals(a.Surname, b.Surname,
                    StringComparison.OrdinalIgnoreCase);

                // ── Same surname warning (orange) ─────────────────────────────
                if (sameSurname)
                {
                    bool sameInitial = string.Equals(a.Initial, b.Initial,
                        StringComparison.OrdinalIgnoreCase);

                    if (sameInitial)
                        Msg("REF-WARN", "WARNING",
                            $"Same surname \"{a.SurnameRaw}\" and same initial " +
                            $"\"{a.InitialRaw}\" detected — please rearrange manually.",
                            a.Excerpt, a.ParaStart);
                    else
                        Msg("REF-WARN", "WARNING",
                            $"Same surname \"{a.SurnameRaw}\" detected " +
                            $"(\"{a.InitialRaw}\" and \"{b.InitialRaw}\") " +
                            $"— please verify order manually to be fully sure.",
                            a.Excerpt, a.ParaStart);
                    issueCount++;
                }

                // ── Same surname + same year + different initial (orange) ─────
                bool sameYear = a.Year == b.Year;
                bool diffInitial = !string.Equals(a.Initial, b.Initial,
                    StringComparison.OrdinalIgnoreCase);

                if (sameSurname && sameYear && diffInitial)
                {
                    Msg("REF-WARN", "WARNING",
                        $"Same surname \"{a.SurnameRaw}\" and same year \"{a.Year}\" " +
                        $"but different initials (\"{a.InitialRaw}\" vs \"{b.InitialRaw}\") — " +
                        $"please verify these are two distinct authors and citations are correct.",
                        a.Excerpt, a.ParaStart);
                    issueCount++;
                }
            }

            // =====================================================================
            // PASS 2 — Blue rule: all pairs within the same surname group
            //
            // WHY a second pass?
            // The blue rule checks: same surname + same author count + same year
            // (ignoring a/b/c suffix) + different first-author initial.
            //
            // Example that requires cross-group scanning:
            //   Zhang, A., and M. Dong,  2004:   ← refs[i]
            //   Zhang, C.,               2005:   ← refs[i+1]  ← breaks pair adjacency
            //   Zhang, C., and J. Ling,  2004a:  ← refs[i+2]
            //
            // "Zhang A. 2004" and "Zhang C. 2004a" share surname + author count (2)
            // + year (2004) + different initial — but they are NOT consecutive, so
            // Pass 1 would never compare them. This pass groups all refs by surname
            // and then checks every pair within that group.
            //
            // A HashSet<string> deduplicates so the same (a.Label, b.Label) pair
            // is never reported twice even if both orderings are encountered.
            //
            // Category "REF-BLUE" → blue badge in TaskPaneWinForms.
            // =====================================================================
            var blueReported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Group refs by normalised surname
            var bySurname = new Dictionary<string, List<RefEntry>>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in refs)
            {
                if (!bySurname.TryGetValue(r.Surname, out var bucket))
                {
                    bucket = new List<RefEntry>();
                    bySurname[r.Surname] = bucket;
                }
                bucket.Add(r);
            }

            foreach (var bucket in bySurname.Values)
            {
                // Only worth checking when two or more refs share the surname
                if (bucket.Count < 2) continue;

                for (int i = 0; i < bucket.Count - 1; i++)
                {
                    for (int j = i + 1; j < bucket.Count; j++)
                    {
                        var a = bucket[i];
                        var b = bucket[j];

                        bool sameAuthorCount = a.AuthorCount == b.AuthorCount;
                        bool sameYearNum = a.Year == b.Year;
                        bool sameInit = string.Equals(a.Initial, b.Initial,
                                                  StringComparison.OrdinalIgnoreCase);
                        bool diffInit = !sameInit;

                        // Skip entirely if year numbers differ — nothing to flag
                        if (!sameYearNum) continue;

                        // ── BLUE RULE A ───────────────────────────────────────
                        // Same surname + same author count + same year number
                        // + DIFFERENT first-author initial.
                        // e.g. Zhang A. 2004  vs  Zhang C. 2004a
                        if (sameAuthorCount && diffInit)
                        {
                            string keyAB = a.Label + "|||" + b.Label;
                            string keyBA = b.Label + "|||" + a.Label;
                            if (!blueReported.Contains(keyAB) && !blueReported.Contains(keyBA))
                            {
                                blueReported.Add(keyAB);
                                Msg("REF-BLUE", "INFO",
                                    $"Same surname \"{a.SurnameRaw}\", same author count " +
                                    $"({a.AuthorCount}), and same year \"{a.Year}\" " +
                                    $"(ignoring a/b/c suffix) but different first-author initials " +
                                    $"(\"{a.InitialRaw}\" vs \"{b.InitialRaw}\") — " +
                                    $"please confirm these are distinct authors and that " +
                                    $"in-text citations are unambiguous.",
                                    a.Excerpt, a.ParaStart);
                                issueCount++;
                            }
                        }

                        // ── BLUE RULE B ───────────────────────────────────────
                        // Same surname + same initial + same author count
                        // + same year number + one entry has a year suffix (a/b/c)
                        // and the other has NO suffix.
                        //
                        // Example caught:
                        //   Wheeler, M. C., and H. H. Hendon, 2004a  → suffix "a"
                        //   Wheeler, M. C., and A. Donald,    2004   → suffix ""
                        //
                        // These are the same first author, same year — the unsuffixed
                        // entry needs a suffix (typically "b") to disambiguate.
                        // Reports against whichever entry is missing the suffix.
                        if (sameAuthorCount && sameInit)
                        {
                            bool aHasSuffix = !string.IsNullOrEmpty(a.YearSuffix);
                            bool bHasSuffix = !string.IsNullOrEmpty(b.YearSuffix);

                            // Only fire when exactly one of the two is missing a suffix
                            if (aHasSuffix != bHasSuffix)
                            {
                                string keyAB = a.Label + "|||" + b.Label;
                                string keyBA = b.Label + "|||" + a.Label;
                                if (!blueReported.Contains(keyAB) && !blueReported.Contains(keyBA))
                                {
                                    blueReported.Add(keyAB);

                                    // Point the message at the entry missing the suffix
                                    var missing = aHasSuffix ? b : a;
                                    var hasOne = aHasSuffix ? a : b;

                                    Msg("REF-BLUE", "INFO",
                                        $"Same surname \"{missing.SurnameRaw}\", same initial " +
                                        $"\"{missing.InitialRaw}\", same author count " +
                                        $"({missing.AuthorCount}), and same year \"{missing.Year}\" — " +
                                        $"\"{missing.Label}\" has no year suffix but " +
                                        $"\"{hasOne.Label}\" uses \"{hasOne.Year}{hasOne.YearSuffix}\". " +
                                        $"Please add a \"b\" suffix to \"{missing.Label}\" " +
                                        $"(or reassign suffixes) to disambiguate in-text citations.",
                                        missing.Excerpt, missing.ParaStart);
                                    issueCount++;
                                }
                            }
                        }
                    }
                }
            }

            if (issueCount == 0)
                TaskPaneWinForms.AddMessage(
                    "REF-ORD", "INFO",
                    $"Reference order check passed — {refs.Count} reference(s) " +
                    "are in correct AMS alphabetical order.");
        }

        private string Truncate(string s, int max) =>
            s == null ? "" : s.Length <= max ? s : s.Substring(0, max) + "…";
    }
}