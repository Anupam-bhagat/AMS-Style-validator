using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin.Checkers
{
    public class ReferenceFormatChecker
    {
        // =====================================================================
        // OOXML NAMESPACE
        // =====================================================================
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static string NK(string s) =>
            (s ?? "").Replace(" ", "").Replace("-", "").Replace("_", "")
                     .ToLowerInvariant();

        // =====================================================================
        // EXCEPTION JOURNALS
        // =====================================================================
        private static readonly string[] ExceptionJournalsNK = new[]
        {
            NK("Geofys. Publ."), NK("Natl. Wea. Dig."),
            NK("Electron. J. Severe Storms Meteor."), NK("Tellus"),
            NK("Geography"), NK("Oceanography"),
            NK("Deep-Sea Res. Oceanogr. Abstr."),
        };

        private static readonly string[] ExceptionJournalsDisplay = new[]
        {
            "Geofys. Publ.", "Natl. Wea. Dig.",
            "Electron. J. Severe Storms Meteor.", "Tellus",
            "Geography", "Oceanography",
            "Deep-Sea Res. Oceanogr. Abstr.",
        };

        // =====================================================================
        // JOURNAL NAME RULES — JAS and JAMC
        // =====================================================================
        private struct JournalEra
        {
            public int FromYear, ToYear;
            public string AbbrevNK, AbbrevDisplay;
        }

        private static readonly HashSet<string> JasFamilyNK =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { NK("J. Meteor."), NK("J. Atmos. Sci.") };

        private static readonly JournalEra[] JasEras = new[]
        {
            new JournalEra { FromYear=1944, ToYear=1961,         AbbrevNK=NK("J. Meteor."),     AbbrevDisplay="J. Meteor."    },
            new JournalEra { FromYear=1962, ToYear=int.MaxValue, AbbrevNK=NK("J. Atmos. Sci."), AbbrevDisplay="J. Atmos. Sci." },
        };

        private static readonly HashSet<string> JamcFamilyNK =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { NK("J. Appl. Meteor."), NK("J. Climate Appl. Meteor."), NK("J. Appl. Meteor. Climatol.") };

        private static readonly JournalEra[] JamcEras = new[]
        {
            new JournalEra { FromYear=1962, ToYear=1982,         AbbrevNK=NK("J. Appl. Meteor."),           AbbrevDisplay="J. Appl. Meteor."           },
            new JournalEra { FromYear=1983, ToYear=1987,         AbbrevNK=NK("J. Climate Appl. Meteor."),   AbbrevDisplay="J. Climate Appl. Meteor."   },
            new JournalEra { FromYear=1988, ToYear=2005,         AbbrevNK=NK("J. Appl. Meteor."),           AbbrevDisplay="J. Appl. Meteor."           },
            new JournalEra { FromYear=2006, ToYear=int.MaxValue, AbbrevNK=NK("J. Appl. Meteor. Climatol."), AbbrevDisplay="J. Appl. Meteor. Climatol." },
        };

        // =====================================================================
        // JGR RULE
        // =====================================================================
        private const string JgrBareNK = "jgeophysres";

        private static readonly HashSet<string> JgrValidNK =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                NK("J. Geophys. Res. Atmos."), NK("J. Geophys. Res. Oceans"),
                NK("J. Geophys. Res. Space Physics"), NK("J. Geophys. Res. Solid Earth"),
                NK("J. Geophys. Res. Biogeosci."), NK("J. Geophys. Res. Planets"),
                NK("J. Geophys. Res. Earth Surface"),
            };

        private static readonly HashSet<string> JgrFamilyNK =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                NK("J. Geophys. Res."), NK("J. Geophys. Res. Atmos."),
                NK("J. Geophys. Res. Oceans"), NK("J. Geophys. Res. Space Physics"),
                NK("J. Geophys. Res. Solid Earth"), NK("J. Geophys. Res. Biogeosci."),
                NK("J. Geophys. Res. Planets"), NK("J. Geophys. Res. Earth Surface"),
            };

        private static readonly string[] JgrFamilyPlainText = new[]
        {
            "J. Geophys. Res. Atmos.", "J. Geophys. Res. Oceans",
            "J. Geophys. Res. Space Physics", "J. Geophys. Res. Solid Earth",
            "J. Geophys. Res. Biogeosci.", "J. Geophys. Res. Planets",
            "J. Geophys. Res. Earth Surface", "J. Geophys. Res.",
        };

        // =====================================================================
        // DEEP-SEA RESEARCH
        // =====================================================================
        private static readonly string _dsrBareNK = NK("Deep-Sea Res.");
        private static readonly string _dsrOceanAbstrNK = NK("Deep-Sea Res. Oceanogr. Abstr.");
        private static readonly string _dsrSeriesINK = NK("Deep-Sea Res. I");
        private static readonly string _dsrSeriesIINK = NK("Deep-Sea Res. II");

        private static readonly HashSet<string> DsrValidNK =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                NK("Deep-Sea Res."), NK("Deep-Sea Res. Oceanogr. Abstr."),
                NK("Deep-Sea Res. I"), NK("Deep-Sea Res. II"),
            };

        private static readonly Regex RxDsrFuzzy = new Regex(
            @"Deep[\s\-]*Sea[\s\-]*Res(?:earch|\.?)(?:[\s\.]*Oceanogr\.?[\s\.]*Abstr\.?|[\s\.]*(?:I{1,2})(?=[\s,;]|$))?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex RxDsrVolAB = new Regex(
            @"\b(\d+)([AB])\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex RxDsrSeriesSuffix = new Regex(
            @"Deep[\s\-]*Sea[\s\-]*Res(?:earch|\.?)[\s\.]*(?<series>I{1,2})(?=[\s,;]|$)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // =====================================================================
        // PROCEEDING ORDINALS — 1st–10th should be first–tenth
        // =====================================================================
        private static readonly Regex RxOrdinal1to10 = new Regex(
            @"\b(1st|2nd|3rd|4th|5th|6th|7th|8th|9th|10th)\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Dictionary<string, string> OrdinalWordMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "1st",  "first"   }, { "2nd",  "second"  }, { "3rd",  "third"   },
                { "4th",  "fourth"  }, { "5th",  "fifth"   }, { "6th",  "sixth"   },
                { "7th",  "seventh" }, { "8th",  "eighth"  }, { "9th",  "ninth"   },
                { "10th", "tenth"   },
            };

        // =====================================================================
        // YEAR REGEX
        // =====================================================================
        private static readonly Regex RxPubYear =
            new Regex(@"\b(\d{4})\s*:", RegexOptions.Compiled);

        // =====================================================================
        // PARAGRAPH STYLES
        // =====================================================================
        private static readonly HashSet<string> JournalParaStylesNK =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "journalref", "datalref" };

        private static readonly HashSet<string> ProceedingParaStylesNK =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "proceedingref" };

        // =====================================================================
        // FILTERS
        // =====================================================================
        private static readonly Regex RxHasYear =
            new Regex(@"\b\d{4}\b", RegexOptions.Compiled);

        private static readonly Regex RxIssnumPlainText =
            new Regex(@",\s*\d+\s*\((\d+)\)\s*,", RegexOptions.Compiled);

        private static readonly Regex RxFirstPagePlainText =
            new Regex(@",\s*\d+\s*(?:\(\d+\)\s*)?,\s*(\d+)\s*[,\u2013\-]",
                RegexOptions.Compiled);

        // =====================================================================
        // DATA MODEL
        // =====================================================================
        private class RefPara
        {
            public string FullText;
            public int WordStart;
            public bool IsProceeding;
            public string JnrlTitleText = null;
            public string IssNumText = null;
            public string FirstPageText = null;
        }

        // =====================================================================
        // RUN
        // =====================================================================
        public void Run()
        {
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch { return; }

            var ui = ThisAddIn.TaskPaneUI;

            try
            {
                string docPath = "";
                try { docPath = doc.FullName; } catch { }

                bool canUseOoxml =
                    !string.IsNullOrEmpty(docPath) &&
                    File.Exists(docPath) &&
                    !docPath.StartsWith(@"\\", StringComparison.Ordinal) &&
                    (docPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                     docPath.EndsWith(".docm", StringComparison.OrdinalIgnoreCase));

                // ── COM pass — collect reference + proceeding paragraphs ───────
                var refParas = new List<RefPara>();
                bool screenWasOn = true;

                try
                {
                    screenWasOn = doc.Application.ScreenUpdating;
                    doc.Application.ScreenUpdating = false;

                    foreach (Word.Paragraph para in doc.Paragraphs)
                    {
                        string sk;
                        try { sk = NK(para.get_Style().NameLocal); } catch { continue; }

                        bool isJournal = JournalParaStylesNK.Contains(sk);
                        bool isProceeding = ProceedingParaStylesNK.Contains(sk);

                        if (!isJournal && !isProceeding) continue;

                        string txt; int ws;
                        try { txt = (para.Range.Text ?? "").TrimEnd('\r', '\n', '\a'); ws = para.Range.Start; }
                        catch { continue; }

                        if (txt.Length < 10) continue;
                        if (isJournal && !RxHasYear.IsMatch(txt)) continue;

                        refParas.Add(new RefPara
                        {
                            FullText = txt,
                            WordStart = ws,
                            IsProceeding = isProceeding
                        });
                    }
                }
                finally
                {
                    try { doc.Application.ScreenUpdating = screenWasOn; } catch { }
                }

                if (refParas.Count == 0)
                {
                    TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                        "No journal_ref or proceeding_ref paragraphs found. " +
                        "Make sure references use the correct paragraph style.");
                    ui?.SetStatus();
                    return;
                }

                // ── Background thread ─────────────────────────────────────────
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        var journalParas = refParas.Where(p => !p.IsProceeding).ToList();
                        if (canUseOoxml && journalParas.Count > 0)
                            FillFromOoxml(docPath, journalParas);

                        int issues = 0;

                        foreach (var p in refParas)
                        {
                            if (p.IsProceeding)
                            {
                                issues += CheckProceedingOrdinals(p);
                            }
                            else
                            {
                                issues += CheckIssNum(p);
                                issues += CheckJournalName(p);
                                issues += CheckJgrSubtitle(p);
                                issues += CheckJamesFormat(p);
                                issues += CheckDeepSeaRes(p);
                            }
                        }

                        if (issues == 0)
                            TaskPaneWinForms.AddMessage("REF-FMT", "INFO",
                                $"Reference format check passed — {refParas.Count} " +
                                "reference(s) checked, no issues found.");
                    }
                    catch (Exception ex)
                    {
                        TaskPaneWinForms.AddMessage("REF-FMT", "ERROR",
                            "Reference format checker error: " + ex.Message);
                    }
                    finally { ui?.SetStatus(); }
                });
            }
            catch (Exception ex)
            {
                TaskPaneWinForms.AddMessage("REF-FMT", "ERROR",
                    "Reference format checker error: " + ex.Message);
                ui?.SetStatus();
            }
        }

        // =====================================================================
        // OOXML FILL
        // =====================================================================
        private static void FillFromOoxml(string docPath, List<RefPara> refParas)
        {
            try
            {
                XDocument bodyXml;
                using (var fs = new FileStream(docPath, FileMode.Open,
                           FileAccess.Read, FileShare.ReadWrite))
                using (var pkg = System.IO.Packaging.Package.Open(fs))
                {
                    var uri = new Uri("/word/document.xml", UriKind.Relative);
                    if (!pkg.PartExists(uri)) return;
                    using (var stream = pkg.GetPart(uri).GetStream())
                        bodyXml = XDocument.Load(stream);
                }

                int idx = 0;
                foreach (var paraElem in bodyXml.Descendants(W + "p"))
                {
                    if (idx >= refParas.Count) break;

                    var pPr = paraElem.Element(W + "pPr");
                    var pStyle = pPr?.Element(W + "pStyle");
                    if (pStyle == null) continue;

                    string sid = NK((string)pStyle.Attribute(W + "val") ?? "");
                    if (!JournalParaStylesNK.Contains(sid)) continue;

                    string paraText = string.Concat(
                        paraElem.Descendants(W + "t").Select(t => (string)t ?? ""))
                        .TrimEnd('\r', '\n', '\a');

                    if (paraText.Length < 30 || !RxHasYear.IsMatch(paraText)) continue;

                    var jnrlParts = new List<string>();
                    var issnParts = new List<string>();
                    var fpParts = new List<string>();

                    foreach (var runElem in paraElem.Descendants(W + "r"))
                    {
                        string runText = string.Concat(
                            runElem.Elements(W + "t").Select(t => (string)t ?? ""));
                        if (string.IsNullOrEmpty(runText)) continue;

                        string csId = "";
                        var rpr = runElem.Element(W + "rPr");
                        if (rpr != null)
                        {
                            var rs = rpr.Element(W + "rStyle");
                            if (rs != null)
                                csId = NK((string)rs.Attribute(W + "val") ?? "");
                        }

                        if (csId == "jnrltitle") jnrlParts.Add(runText);
                        else if (csId == "issnum") issnParts.Add(runText);
                        else if (csId == "firstpage") fpParts.Add(runText);
                    }

                    RefPara p = refParas[idx];
                    if (jnrlParts.Count > 0) p.JnrlTitleText = string.Concat(jnrlParts).Trim();
                    if (issnParts.Count > 0) p.IssNumText = string.Concat(issnParts).Trim();
                    if (fpParts.Count > 0) p.FirstPageText = string.Concat(fpParts).Trim();

                    idx++;
                }
            }
            catch { }
        }

        // =====================================================================
        // CHECK — proceeding ordinal numbers (1st–10th → first–tenth)
        // =====================================================================
        private static int CheckProceedingOrdinals(RefPara p)
        {
            string text = p.FullText;
            int pos = p.WordStart;
            int issues = 0;

            foreach (Match m in RxOrdinal1to10.Matches(text))
            {
                string ordinal = m.Value.ToLowerInvariant();
                string word = OrdinalWordMap.ContainsKey(ordinal)
                                 ? OrdinalWordMap[ordinal] : ordinal;

                TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                    $"Ordinal \"{m.Value}\" found in conference title — " +
                    $"AMS style requires the word form: use \"{word}\" instead of \"{m.Value}\".",
                    Snip(text, m.Index, m.Length),
                    pos + m.Index);
                issues++;
            }

            return issues;
        }

        // =====================================================================
        // CHECK — issue number rule
        // =====================================================================
        private static int CheckIssNum(RefPara p)
        {
            int issues = 0;
            string text = p.FullText;
            string excerpt = Snip(text, 0, Math.Min(80, text.Length));
            int pos = p.WordStart;

            string issnum = p.IssNumText ?? "";
            if (string.IsNullOrEmpty(issnum))
            {
                Match mi = RxIssnumPlainText.Match(text);
                if (mi.Success) issnum = mi.Groups[1].Value.Trim();
            }

            string firstPage = p.FirstPageText ?? "";
            if (string.IsNullOrEmpty(firstPage))
            {
                Match mf = RxFirstPagePlainText.Match(text);
                if (mf.Success)
                {
                    int ci = mf.Groups[1].Index, ce = ci + mf.Groups[1].Length;
                    bool prevLetter = ci > 0 && char.IsLetter(text[ci - 1]);
                    bool nextLetter = ce < text.Length && char.IsLetter(text[ce]);
                    if (!prevLetter && !nextLetter)
                        firstPage = mf.Groups[1].Value.Trim();
                }
            }

            string jnrlTitle = p.JnrlTitleText ?? "";
            if (string.IsNullOrEmpty(jnrlTitle))
                foreach (string jname in ExceptionJournalsDisplay)
                    if (text.Contains(jname)) { jnrlTitle = jname; break; }

            bool fpIsOne = string.Equals(firstPage.Trim(), "1", StringComparison.Ordinal);
            string jnrlNK = NK(jnrlTitle);
            bool isEJ = false;
            string matchedJournal = null;
            foreach (int i in Enumerable.Range(0, ExceptionJournalsNK.Length))
                if (jnrlNK.Contains(ExceptionJournalsNK[i]))
                { isEJ = true; matchedJournal = ExceptionJournalsDisplay[i]; break; }

            bool hasIssue = !string.IsNullOrEmpty(issnum);
            bool needsIssue = fpIsOne || isEJ;

            if (hasIssue && !needsIssue)
            {
                TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                    $"Issue number \"({issnum.Trim()})\" should be deleted. " +
                    $"AMS style omits issue numbers unless the first page is 1" +
                    (string.IsNullOrEmpty(firstPage) ? "" : $" (first page here: \"{firstPage}\")") +
                    ". Query the author if unsure.", excerpt, pos);
                issues++;
            }

            if (!hasIssue && needsIssue)
            {
                string reason = fpIsOne
                    ? "first page is 1 — issue number needed to locate the article"
                    : $"\"{matchedJournal}\" always requires an issue number";
                TaskPaneWinForms.AddMessage("REF-FMT", "ERROR",
                    $"Issue number is missing from this reference — {reason}. " +
                    "Please add the issue number in parentheses after the volume.",
                    excerpt, pos);
                issues++;
            }

            return issues;
        }

        // =====================================================================
        // CHECK — journal name correct for publication year (JAS / JAMC)
        // =====================================================================
        private static int CheckJournalName(RefPara p)
        {
            int issues = 0;
            string text = p.FullText;
            string excerpt = Snip(text, 0, Math.Min(80, text.Length));
            int pos = p.WordStart;

            string jnrlTitle = p.JnrlTitleText ?? "";
            if (string.IsNullOrEmpty(jnrlTitle))
                foreach (string name in new[]
                {
                    "J. Atmos. Sci.", "J. Meteor.",
                    "J. Appl. Meteor. Climatol.", "J. Climate Appl. Meteor.", "J. Appl. Meteor."
                })
                    if (text.Contains(name)) { jnrlTitle = name; break; }

            if (string.IsNullOrEmpty(jnrlTitle)) return 0;

            string jnrlNK = NK(jnrlTitle);
            bool isJas = JasFamilyNK.Contains(jnrlNK);
            bool isJamc = !isJas && JamcFamilyNK.Contains(jnrlNK);
            if (!isJas && !isJamc) return 0;

            Match ym = RxPubYear.Match(text);
            if (!ym.Success) return 0;
            int pubYear = int.Parse(ym.Groups[1].Value);

            JournalEra[] eras = isJas ? JasEras : JamcEras;
            string correctDisplay = null, correctNK = null;
            foreach (var era in eras)
                if (pubYear >= era.FromYear && pubYear <= era.ToYear)
                { correctDisplay = era.AbbrevDisplay; correctNK = era.AbbrevNK; break; }

            if (correctDisplay == null) return 0;
            if (string.Equals(jnrlNK, correctNK, StringComparison.OrdinalIgnoreCase)) return 0;

            string familyName = isJas ? "JAS" : "JAMC";
            TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                $"Wrong journal name detected — \"{jnrlTitle}\" is used for a {pubYear} " +
                $"{familyName} reference but \"{correctDisplay}\" should be used " +
                $"({familyName} was renamed across years). Please check and confirm.",
                excerpt, pos);
            issues++;
            return issues;
        }

        // =====================================================================
        // CHECK — JGR subtitle required from 2013 onwards
        // =====================================================================
        private static int CheckJgrSubtitle(RefPara p)
        {
            string text = p.FullText;
            string excerpt = Snip(text, 0, Math.Min(80, text.Length));
            int pos = p.WordStart;

            string jnrlTitle = p.JnrlTitleText ?? "";
            if (string.IsNullOrEmpty(jnrlTitle))
                foreach (string name in JgrFamilyPlainText)
                    if (text.Contains(name)) { jnrlTitle = name; break; }

            if (string.IsNullOrEmpty(jnrlTitle)) return 0;

            string jnrlNK = NK(jnrlTitle);
            if (!JgrFamilyNK.Contains(jnrlNK)) return 0;

            if (JgrValidNK.Contains(jnrlNK))
            {
                Match ym2 = RxPubYear.Match(text);
                if (ym2.Success && int.Parse(ym2.Groups[1].Value) < 2013)
                {
                    TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                        $"Incorrect journal name detected — \"{jnrlTitle}\" is used for a " +
                        $"{ym2.Groups[1].Value} reference but subtitles were not used before 2013. " +
                        $"Use \"J. Geophys. Res.\" instead. Please check and confirm.",
                        excerpt, pos);
                    return 1;
                }
                return 0;
            }

            Match ym = RxPubYear.Match(text);
            if (!ym.Success) return 0;
            int pubYear = int.Parse(ym.Groups[1].Value);
            if (pubYear < 2013) return 0;

            TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                $"Incomplete journal name detected — \"J. Geophys. Res.\" is used for a " +
                $"{pubYear} reference but a subtitle must be added from 2013 onwards " +
                $"(e.g. \"J. Geophys. Res. Atmos.\", \"J. Geophys. Res. Oceans\", " +
                $"\"J. Geophys. Res. Space Physics\"). Please check and confirm.",
                excerpt, pos);
            return 1;
        }

        // =====================================================================
        // CHECK — JAMES format by year
        // =====================================================================
        private static readonly HashSet<string> JamesFamilyNK =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { NK("J. Adv. Model. Earth Syst."), NK("JAMES") };

        private static readonly string[] JamesFamilyPlainText = new[]
            { "J. Adv. Model. Earth Syst.", "JAMES" };

        private static int CheckJamesFormat(RefPara p)
        {
            string text = p.FullText;
            string excerpt = Snip(text, 0, Math.Min(80, text.Length));
            int pos = p.WordStart;

            string jnrlTitle = p.JnrlTitleText ?? "";
            if (string.IsNullOrEmpty(jnrlTitle))
                foreach (string name in JamesFamilyPlainText)
                    if (text.Contains(name)) { jnrlTitle = name; break; }

            if (string.IsNullOrEmpty(jnrlTitle)) return 0;
            if (!JamesFamilyNK.Contains(NK(jnrlTitle))) return 0;

            Match ym = RxPubYear.Match(text);
            if (!ym.Success) return 0;
            int pubYear = int.Parse(ym.Groups[1].Value);

            string issnum = p.IssNumText ?? "";
            if (string.IsNullOrEmpty(issnum))
            {
                Match mi = RxIssnumPlainText.Match(text);
                if (mi.Success) issnum = mi.Groups[1].Value.Trim();
            }
            bool hasIssue = !string.IsNullOrEmpty(issnum);

            if (pubYear >= 2009 && pubYear <= 2010)
            {
                if (!hasIssue)
                {
                    TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                        $"JAMES ({pubYear}): Issue number is required for articles " +
                        $"from 2009–2010. Please verify and add the issue number " +
                        $"in parentheses after the volume.", excerpt, pos);
                    return 1;
                }
            }
            else if (pubYear >= 2011 && pubYear <= 2019)
            {
                TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                    $"JAMES ({pubYear}): Articles from 2011–2019 should include a " +
                    $"page range (not a citation number). Please check and confirm " +
                    $"the page range is present.", excerpt, pos);
                return 1;
            }
            else if (pubYear >= 2020)
            {
                TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                    $"JAMES ({pubYear}): Articles from 2020 onward should include a " +
                    $"citation number (not a page range). Please check and confirm " +
                    $"the citation number is present.", excerpt, pos);
                return 1;
            }

            return 0;
        }

        // =====================================================================
        // CHECK — Deep-Sea Research journal name and era rules
        // =====================================================================
        private static int CheckDeepSeaRes(RefPara p)
        {
            string text = p.FullText;
            string excerpt = Snip(text, 0, Math.Min(80, text.Length));
            int pos = p.WordStart;

            string rawName = p.JnrlTitleText ?? "";
            bool fromTag = !string.IsNullOrEmpty(rawName);

            if (!fromTag)
            {
                Match mf = RxDsrFuzzy.Match(text);
                if (!mf.Success) return 0;
                rawName = mf.Value.Trim();
            }
            else
            {
                if (!RxDsrFuzzy.IsMatch(rawName)) return 0;
            }

            Match ym = RxPubYear.Match(text);
            if (!ym.Success) return 0;
            int pubYear = int.Parse(ym.Groups[1].Value);

            int issues = 0;
            string rawNK = NK(rawName);
            bool spellingOk = DsrValidNK.Contains(rawNK);

            if (!spellingOk)
            {
                string suggestion = SuggestDsrCorrection(rawName, pubYear);
                TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                    $"Incorrect Deep-Sea Research journal name detected — \"{rawName}\" " +
                    $"has incorrect spelling or punctuation. " +
                    (suggestion != null ? $"Suggested correct form: \"{suggestion}\". " : "") +
                    "Please correct the journal name.", excerpt, pos);
                issues++;
            }

            if (pubYear < 1979)
            {
                if (spellingOk &&
                    !string.Equals(rawNK, _dsrBareNK, StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(rawNK, _dsrOceanAbstrNK, StringComparison.OrdinalIgnoreCase))
                {
                    TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                        $"Incorrect Deep-Sea Research journal name for {pubYear} — " +
                        $"\"{rawName}\" was used but pre-1979 references must use " +
                        $"\"Deep-Sea Res.\" or \"Deep-Sea Res. Oceanogr. Abstr.\". " +
                        $"Please check and correct.", excerpt, pos);
                    issues++;
                }
            }
            else if (pubYear <= 1992)
            {
                if (spellingOk &&
                    !string.Equals(rawNK, _dsrBareNK, StringComparison.OrdinalIgnoreCase))
                {
                    TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                        $"Incorrect Deep-Sea Research journal name for {pubYear} — " +
                        $"\"{rawName}\" was used but 1979–1992 references must use " +
                        $"\"Deep-Sea Res.\" (series A/B shown in volume number, not title). " +
                        $"Please check and correct.", excerpt, pos);
                    issues++;
                }

                if (!RxDsrVolAB.IsMatch(text))
                {
                    TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                        $"Missing series letter in Deep-Sea Research volume ({pubYear}) — " +
                        $"The volume number must include an A or B suffix (e.g. \"15A\" or \"27B\"). " +
                        $"Please check and add the series letter.", excerpt, pos);
                    issues++;
                }
            }
            else
            {
                if (spellingOk)
                {
                    bool hasI = string.Equals(rawNK, _dsrSeriesINK, StringComparison.OrdinalIgnoreCase);
                    bool hasII = string.Equals(rawNK, _dsrSeriesIINK, StringComparison.OrdinalIgnoreCase);
                    if (!hasI && !hasII && !RxDsrSeriesSuffix.IsMatch(text))
                    {
                        TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                            $"Incomplete Deep-Sea Research journal name for {pubYear} — " +
                            $"\"{rawName}\" is missing the series suffix. " +
                            $"From 1993 onwards use \"Deep-Sea Res. I\" or \"Deep-Sea Res. II\". " +
                            $"Please check and confirm.", excerpt, pos);
                        issues++;
                    }
                }
                else if (!RxDsrSeriesSuffix.IsMatch(text))
                {
                    TaskPaneWinForms.AddMessage("REF-FMT", "WARNING",
                        $"Deep-Sea Research reference ({pubYear}) also appears to be " +
                        $"missing the series suffix (I or II). From 1993 onwards " +
                        $"use \"Deep-Sea Res. I\" or \"Deep-Sea Res. II\". " +
                        $"Please check and confirm.", excerpt, pos);
                    issues++;
                }
            }

            return issues;
        }

        private static string SuggestDsrCorrection(string rawName, int pubYear)
        {
            string rNK = NK(rawName);
            bool hasOceanogr = rNK.Contains("oceanogr") || rNK.Contains("abstr");
            bool hasSeriesII = Regex.IsMatch(rawName, @"\bII\b", RegexOptions.IgnoreCase);
            bool hasSeriesI = !hasSeriesII && Regex.IsMatch(rawName, @"\bI\b", RegexOptions.IgnoreCase);

            if (pubYear < 1979) return hasOceanogr ? "Deep-Sea Res. Oceanogr. Abstr." : "Deep-Sea Res.";
            if (pubYear <= 1992) return "Deep-Sea Res.";
            if (hasSeriesII) return "Deep-Sea Res. II";
            if (hasSeriesI) return "Deep-Sea Res. I";
            return null;
        }

        // =====================================================================
        // HELPERS
        // =====================================================================
        private static string NormWs(string s) =>
            string.IsNullOrEmpty(s) ? "" : Regex.Replace(s.Trim(), @"\s+", " ");

        private static string Snip(string text, int index, int length)
        {
            if (string.IsNullOrEmpty(text)) return "";
            index = Math.Max(0, Math.Min(index, text.Length - 1));
            length = Math.Max(0, Math.Min(length, text.Length - index));
            int s = Math.Max(0, index - 15);
            int e = Math.Min(text.Length, index + length + 15);
            string snip = text.Substring(s, e - s).Trim();
            return snip.Length > 70 ? snip.Substring(0, 70) + "…" : snip;
        }
    }
}