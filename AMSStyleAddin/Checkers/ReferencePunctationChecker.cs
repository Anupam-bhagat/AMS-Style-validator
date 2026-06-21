using System;
using System.Collections.Generic;
using System.Threading;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin.Checkers
{
    public class ReferencePunctuationChecker
    {
        // ── Issue counter ─────────────────────────────────────────────────────
        private int _issueNum;

        private void Msg(string cat, string level, string msg,
                         string snippet = "", int start = 0)
        {
            _issueNum++;
            TaskPaneWinForms.AddMessage(cat, level,
                $"#{_issueNum} {msg}", snippet, start);
        }

        // ── Reference paragraph style names (normalised) ──────────────────────
        private static readonly HashSet<string> RefStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "journalref","bookref","otherref",
            "proceedingref","reportref","thesisref"
        };

        // Normalise: lowercase, strip spaces / hyphens / underscores
        private static string NK(string s) =>
            (s ?? "").Replace(" ", "").Replace("-", "")
                     .Replace("_", "").ToLowerInvariant();

        // ── Word XML namespaces ───────────────────────────────────────────────
        private const string WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // ── Char-style keys for formatting rules ─────────────────────────────
        // Author
        private const string CS_SURNAME = "authorsurname";
        private const string CS_FNAME = "authorfname";
        private const string CS_ETAL = "etal";

        // Title styles
        private const string CS_ARTICLETITLE = "articletitle";
        private const string CS_BOOKTITLE = "booktitle";
        private const string CS_CHAPTERTITLE = "chaptertitle";
        private const string CS_CONFTITLE = "conferencetitle";
        private const string CS_JNRLTITLE = "jnrltitle";
        private const string CS_THESISTITLE = "thesistitle";

        // Publication info
        private const string CS_VOLNUM = "volnum";
        private const string CS_FIRSTPAGE = "firstpage";
        private const string CS_LASTPAGE = "lastpage";
        private const string CS_PUBNAME = "pubname";
        private const string CS_PUBLOC = "publoc";
        private const string CS_EDITION = "edition";
        private const string CS_DOI = "doi";
        private const string CS_URL = "url";
        private const string CS_SUPPL = "suppl";
        private const string CS_YEAR = "year";

        // Editor
        private const string CS_EDITORSURNAME = "editorsurname";
        private const string CS_EDITORFNAME = "editorfname";

        // Other
        private const string CS_THESIS = "thesis";
        private const string CS_COMMENT = "comment";
        private const string CS_HYPERLINK = "hyperlink";

        // Some manuscripts use "0"-suffixed variants (e.g. booktitle0, articletitle0)
        private static bool IsBookTitle(string cs) =>
            cs == CS_BOOKTITLE || cs == "booktitle0";

        private static bool IsArticleTitle(string cs) =>
            cs == CS_ARTICLETITLE || cs == "articletitle0";

        // ── Token (one XML <w:r> run, collapsed) ──────────────────────────────
        private struct Token
        {
            public string Text;    // concatenated w:t text
            public string CS;      // normalised char-style, "" = plain
            public int Pos;     // approximate document position
            public bool Bold;    // true if w:b is set in rPr
            public bool Italic;  // true if w:i is set in rPr
        }

        private static string Truncate(string s, int max) =>
            s == null ? "" : s.Length <= max ? s : s.Substring(0, max) + "…";

        // =====================================================================
        // RUN  — entry point called from the ribbon
        // =====================================================================
        public void Run()
        {
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch { return; }

            _issueNum = 0;
            int refCount = 0;
            int issueCount = 0;

            bool screenWasOn = true;
            int paraIdx = 0;

            try
            {
                screenWasOn = doc.Application.ScreenUpdating;
                doc.Application.ScreenUpdating = false;

                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    // Yield to Word every 100 paragraphs
                    paraIdx++;
                    if (paraIdx % 100 == 0)
                    {
                        doc.Application.ScreenUpdating = screenWasOn;
                        Thread.Sleep(0);
                        doc.Application.ScreenUpdating = false;
                    }

                    // ── Check paragraph style ─────────────────────────────────
                    string pStyle = "";
                    try { pStyle = NK(para.get_Style().NameLocal); }
                    catch { continue; }

                    bool isRef = false;
                    foreach (var rs in RefStyles)
                        if (pStyle == rs) { isRef = true; break; }
                    if (!isRef) continue;

                    // ── Get paragraph text for excerpts ───────────────────────
                    string paraText = "";
                    int paraStart = 0;
                    try
                    {
                        paraText = (para.Range.Text ?? "").TrimEnd('\r', '\n').Trim();
                        paraStart = para.Range.Start;
                    }
                    catch { continue; }
                    if (string.IsNullOrWhiteSpace(paraText)) continue;

                    refCount++;

                    // ── Build token list from raw XML (fast — no COM loop) ────
                    List<Token> tokens = BuildTokensFromXml(para, paraStart);
                    if (tokens == null || tokens.Count == 0) continue;

                    // ── Find end of author block (first 'year' token) ─────────
                    int authorEnd = tokens.Count;
                    for (int i = 0; i < tokens.Count; i++)
                    {
                        if (tokens[i].CS == CS_YEAR)
                        { authorEnd = i; break; }
                    }
                    if (authorEnd == 0) continue;

                    // ── Apply author block rules ──────────────────────────────
                    issueCount += CheckAuthorBlock(
                        tokens, authorEnd, paraText, paraStart);

                    // ── Apply formatting rules (whole paragraph) ──────────────
                    issueCount += CheckFormattingRules(
                        tokens, paraText, paraStart);
                }
            }
            finally
            {
                try { doc.Application.ScreenUpdating = screenWasOn; } catch { }
            }

            if (refCount == 0)
            {
                Msg("REF-PUNC", "WARNING",
                    "No reference paragraphs found. Ensure paragraphs use " +
                    "styles: journal_ref, book_ref, other_ref, " +
                    "proceeding_ref, report_ref, thesis_ref.");
                return;
            }

            if (issueCount == 0)
                TaskPaneWinForms.AddMessage("REF-PUNC", "INFO",
                    $"Author punctuation style check passed — " +
                    $"{refCount} reference(s) checked, no issues found.");
        }

        // =====================================================================
        // BUILD TOKENS FROM XML
        //
        // Reads the paragraph's raw Office Open XML directly — no COM
        // character-by-character iteration.  Each <w:r> element becomes one
        // token.  Adjacent runs that share the same char style are merged.
        //
        // paraStart is the Word document position of the first character of
        // the paragraph; we accumulate a character offset to give approximate
        // positions for jump-to-location.
        // =====================================================================
        private static List<Token> BuildTokensFromXml(
            Word.Paragraph para, int paraStart)
        {
            var tokens = new List<Token>();
            try
            {
                // Get the paragraph's XML
                string xml = para.Range.WordOpenXML;
                if (string.IsNullOrEmpty(xml)) return tokens;

                var xd = new XmlDocument();
                xd.LoadXml(xml);

                var nsMgr = new XmlNamespaceManager(xd.NameTable);
                nsMgr.AddNamespace("w", WNS);

                // Walk every <w:r> in document order
                XmlNodeList runs = xd.SelectNodes("//w:r", nsMgr);
                if (runs == null) return tokens;

                string curText = "";
                string curCS = null;   // null = not started yet
                bool curBold = false;
                bool curItalic = false;
                int curOffset = 0;
                int charOffset = 0;

                foreach (XmlNode r in runs)
                {
                    // ── Char style ────────────────────────────────────────────
                    string cs = "";
                    XmlNode rStyle = r.SelectSingleNode("w:rPr/w:rStyle/@w:val", nsMgr);
                    if (rStyle != null)
                        cs = NK(rStyle.Value ?? "");

                    // ── Bold — present and not explicitly off ─────────────────
                    bool bold = false;
                    XmlNode bNode = r.SelectSingleNode("w:rPr/w:b", nsMgr);
                    if (bNode != null)
                    {
                        // w:val is in the w: namespace — use GetNamedItem
                        XmlNode bVal = bNode.Attributes?.GetNamedItem("val", WNS);
                        // Also try unprefixed fallback
                        if (bVal == null) bVal = bNode.Attributes?["w:val"];
                        bold = (bVal == null || (bVal.Value != "0" && bVal.Value != "false"));
                    }

                    // ── Italic — present and not explicitly off ───────────────
                    bool italic = false;
                    XmlNode iNode = r.SelectSingleNode("w:rPr/w:i", nsMgr);
                    if (iNode != null)
                    {
                        XmlNode iVal = iNode.Attributes?.GetNamedItem("val", WNS);
                        if (iVal == null) iVal = iNode.Attributes?["w:val"];
                        italic = (iVal == null || (iVal.Value != "0" && iVal.Value != "false"));
                    }

                    // ── Concatenate all <w:t> text in this run ────────────────
                    string text = "";
                    foreach (XmlNode t in r.SelectNodes("w:t", nsMgr))
                        text += t.InnerText;

                    if (text.Length == 0) continue;

                    // ── Merge with previous token if same style AND formatting ─
                    if (curCS == null)
                    {
                        curText = text;
                        curCS = cs;
                        curBold = bold;
                        curItalic = italic;
                        curOffset = paraStart + charOffset;
                    }
                    else if (cs == curCS && bold == curBold && italic == curItalic)
                    {
                        curText += text;
                    }
                    else
                    {
                        tokens.Add(new Token
                        {
                            Text = curText,
                            CS = curCS,
                            Pos = curOffset,
                            Bold = curBold,
                            Italic = curItalic
                        });
                        curText = text;
                        curCS = cs;
                        curBold = bold;
                        curItalic = italic;
                        curOffset = paraStart + charOffset;
                    }

                    charOffset += text.Length;
                }

                // Flush last token
                if (curCS != null && curText.Length > 0)
                    tokens.Add(new Token
                    {
                        Text = curText,
                        CS = curCS,
                        Pos = curOffset,
                        Bold = curBold,
                        Italic = curItalic
                    });
            }
            catch { /* return whatever was built */ }

            return tokens;
        }

        // =====================================================================
        // CHECK AUTHOR BLOCK
        //
        // Iterates tokens[0 .. authorEnd-1] and applies three rules.
        //
        // ── RULE 1  authorsurname ─────────────────────────────────────────────
        //   The token immediately after authorsurname must be:
        //     • plain (CS == "")
        //     • and must start with ", "  (comma + space)
        //   Errors:
        //     1a  next token is styled (comma missing entirely)
        //     1b  next token starts with "," but no space  →  Augustine,J. case
        //     1c  next token has no comma at all
        //
        // ── RULE 2  authorfname ───────────────────────────────────────────────
        //   (a) The authorfname run text must NOT end with '.'
        //       Periods must live in the following plain token.
        //       Multi-initial runs like "J. M", "T.-C", "D. J" are fine —
        //       internal dots are allowed; only trailing dot is wrong.
        //       Error 2a: run ends with '.'  →  Seko "H." case
        //
        //   (b) The token immediately after authorfname must be:
        //       • plain (CS == "")
        //       • and must start with '.'
        //       Error 2b: next token is styled (period missing)
        //       Error 2c: next plain token does not start with '.'
        //
        // ── RULE 3  plain separator between initials ──────────────────────────
        //   When a plain token sits between two authorfname tokens
        //   (i.e. prev=authorfname, next=authorfname) it is an inter-initial
        //   separator.  It must be exactly ". " (period + space).
        //   Error 3:  inter-initial plain separator is not ". "
        //             →  catches Browning K.·A. where ". " has wrong style
        //             →  catches missing space
        //
        // =====================================================================
        private int CheckAuthorBlock(
            List<Token> tokens, int authorEnd,
            string paraText, int paraStart)
        {
            int issues = 0;
            string excerpt = Truncate(paraText, 70);

            for (int i = 0; i < authorEnd; i++)
            {
                Token tok = tokens[i];
                Token next = i + 1 < authorEnd ? tokens[i + 1] : default;
                Token prev = i > 0 ? tokens[i - 1] : default;

                bool hasNext = i + 1 < authorEnd;
                bool hasPrev = i > 0;

                // ── RULE 1 — authorsurname ────────────────────────────────────
                if (tok.CS == CS_SURNAME && hasNext)
                {
                    if (next.CS != "")
                    {
                        // 1a — styled run follows directly
                        // Sub-case: styled run starts with ", " — comma exists but has wrong style
                        // Sub-case: no comma — separator missing entirely
                        if (next.Text.StartsWith(","))
                            Msg("REF-PUNC", "ERROR",
                                $"After authorsurname \"{tok.Text}\": " +
                                $"the \", \" separator has \"{next.CS}\" style applied. " +
                                $"Remove the character style from the comma+space — " +
                                $"it must be plain text.",
                                excerpt, next.Pos);
                        else
                            Msg("REF-PUNC", "ERROR",
                                $"After authorsurname \"{tok.Text}\": " +
                                $"expected plain \", \" separator but got " +
                                $"styled run \"{Truncate(next.Text, 15)}\" " +
                                $"[{next.CS}]. Comma + space missing.",
                                excerpt, tok.Pos);
                        issues++;
                    }
                    else if (!next.Text.StartsWith(", "))
                    {
                        if (next.Text.StartsWith(","))
                            // 1b — comma present, space missing
                            Msg("REF-PUNC", "ERROR",
                                $"After authorsurname \"{tok.Text}\": " +
                                $"comma found but no space after it " +
                                $"(\"{Truncate(next.Text, 10)}\"). " +
                                $"Should be \", \".",
                                excerpt, next.Pos);
                        else
                            // 1c — no comma at all
                            Msg("REF-PUNC", "ERROR",
                                $"After authorsurname \"{tok.Text}\": " +
                                $"expected \", \" but found " +
                                $"\"{Truncate(next.Text, 10)}\". " +
                                $"Comma + space required after surname.",
                                excerpt, next.Pos);
                        issues++;
                    }
                }

                // ── RULE 2 — authorfname ──────────────────────────────────────
                else if (tok.CS == CS_FNAME)
                {
                    // 2a — trailing period inside the styled run
                    if (tok.Text.EndsWith("."))
                    {
                        Msg("REF-PUNC", "ERROR",
                            $"authorfname \"{tok.Text}\" ends with '.' " +
                            $"inside the styled run. " +
                            $"Strip the trailing '.' from the style — " +
                            $"it must sit in the following plain run.",
                            excerpt, tok.Pos);
                        issues++;
                    }

                    if (hasNext)
                    {
                        if (next.CS != "")
                        {
                            // 2b — styled run follows
                            // Sub-case: the styled run IS a period — style is wrong
                            // Sub-case: the styled run is NOT a period — period missing
                            if (next.Text.StartsWith("."))
                                Msg("REF-PUNC", "ERROR",
                                    $"After authorfname \"{tok.Text}\": " +
                                    $"the '.' has \"{next.CS}\" style applied. " +
                                    $"Remove the character style from the period — " +
                                    $"it must be plain text.",
                                    excerpt, next.Pos);
                            else
                                Msg("REF-PUNC", "ERROR",
                                    $"After authorfname \"{tok.Text}\": " +
                                    $"expected plain '.' but found styled run " +
                                    $"\"{Truncate(next.Text, 15)}\" [{next.CS}]. " +
                                    $"Period missing after initial.",
                                    excerpt, tok.Pos);
                            issues++;
                        }
                        else if (!next.Text.StartsWith("."))
                        {
                            // 2c — plain run but no period
                            Msg("REF-PUNC", "ERROR",
                                $"After authorfname \"{tok.Text}\": " +
                                $"plain run \"{Truncate(next.Text, 10)}\" " +
                                $"does not start with '.'. " +
                                $"Period required after initial.",
                                excerpt, next.Pos);
                            issues++;
                        }
                    }
                }

                // ── RULE 3 — plain separator between two authorfname tokens ───
                // Detects: Browning K [authorfname] "." [plain] A [authorfname]
                // where the "." between K and A should have authorfname style.
                //
                // IMPORTANT — only flag for SAME-AUTHOR separators.
                // Cross-author separators (between last initial of one author
                // and first initial of the next) always contain a comma or
                // the word "and", e.g.:
                //   "., "   "., and "   ", and "   ". "  followed by a surname
                // Same-author separators are purely ". " or "." — no comma,
                // no "and", and the NEXT token after the following authorfname
                // is NOT an authorsurname (i.e. no surname boundary crossed).
                //
                // Simple discriminator: if the plain sep contains a comma OR
                // the word "and" it is a cross-author boundary → skip.
                else if (tok.CS == ""
                         && hasPrev && prev.CS == CS_FNAME
                         && hasNext && next.CS == CS_FNAME)
                {
                    string sep = tok.Text;

                    // Cross-author separator — skip entirely, not an error
                    bool isCrossAuthor =
                        sep.Contains(",") ||
                        sep.IndexOf("and", StringComparison.OrdinalIgnoreCase) >= 0;

                    if (!isCrossAuthor)
                    {
                        // Same-author inter-initial separator.
                        // It is plain but should carry authorfname style.
                        if (sep == ". " || sep == "." || sep == " ")
                        {
                            Msg("REF-PUNC", "WARNING",
                                $"Plain run \"{EscapeForMsg(sep)}\" between " +
                                $"authorfname \"{prev.Text}\" and " +
                                $"authorfname \"{next.Text}\" should have " +
                                $"authorfname style applied " +
                                $"(period+space between initials of same author).",
                                excerpt, tok.Pos);
                            issues++;
                        }
                        else
                        {
                            Msg("REF-PUNC", "ERROR",
                                $"Between authorfname \"{prev.Text}\" and " +
                                $"authorfname \"{next.Text}\": " +
                                $"unexpected separator \"{Truncate(sep, 10)}\" — " +
                                $"expected \". \" with authorfname style.",
                                excerpt, tok.Pos);
                            issues++;
                        }
                    }
                }
            }

            return issues;
        }

        // =====================================================================
        // CHECK FORMATTING RULES  (whole paragraph token list)
        //
        // RULE 4  volnum comma
        //   After a volnum token, the next token must be a plain comma that:
        //   4a — is NOT bold
        //   4b — does NOT have volnum style
        //
        // RULE 5  jnrltitle must be italic
        //
        // RULE 6  conferencetitle must be italic
        //
        // RULE 7  booktitle
        //   7a — must be italic
        //   7b — must be title case
        //
        // RULE 8  pubname must be title case
        //
        // RULE 9  articletitle must be sentence case
        //   Collect all consecutive articletitle tokens, concatenate text,
        //   scan word by word.  On the FIRST violation emit one warning for
        //   the whole paragraph and stop — do not report every bad word.
        //   Sentence-case rules:
        //     • First word                   → must be capitalised
        //     • First word after ':' or '—'  → must be capitalised
        //     • All other words              → must start lowercase
        //       (proper nouns / acronyms are false-positive risk — we
        //        therefore emit a WARNING not an ERROR so the editor
        //        can confirm rather than blindly fix)
        //
        // =====================================================================
        private int CheckFormattingRules(
            List<Token> tokens, string paraText, int paraStart)
        {
            int issues = 0;
            string excerpt = Truncate(paraText, 70);

            // ── Pre-pass: collect full article title text for Rule 9 ──────────
            // Concatenate all articletitle/articletitle0 tokens in this para.
            // We only need to do this once per paragraph, not per token.
            bool articleTitleChecked = false;
            var articleTitleBuf = new System.Text.StringBuilder();
            int articleTitlePos = 0;
            bool articleTitlePosSet = false;

            for (int i = 0; i < tokens.Count; i++)
            {
                Token tok = tokens[i];
                if (IsArticleTitle(tok.CS))
                {
                    if (!articleTitlePosSet)
                    { articleTitlePos = tok.Pos; articleTitlePosSet = true; }
                    articleTitleBuf.Append(tok.Text);
                }
            }

            string fullArticleTitle = articleTitleBuf.ToString().Trim();

            for (int i = 0; i < tokens.Count; i++)
            {
                Token tok = tokens[i];

                // ── RULE 4 — comma after volnum ───────────────────────────────
                if (tok.CS == CS_VOLNUM && i + 1 < tokens.Count)
                {
                    Token next = tokens[i + 1];

                    if (next.Text.StartsWith(","))
                    {
                        // 4a — must not be bold
                        if (next.Bold)
                        {
                            Msg("REF-PUNC", "ERROR",
                                $"Comma after volnum \"{Truncate(tok.Text, 10)}\" " +
                                $"is bold. The comma separator must not be bold.",
                                excerpt, next.Pos);
                            issues++;
                        }

                        // 4b — must not carry volnum style
                        if (next.CS == CS_VOLNUM)
                        {
                            Msg("REF-PUNC", "ERROR",
                                $"Comma after volnum \"{Truncate(tok.Text, 10)}\" " +
                                $"has volnum style applied. " +
                                $"The comma must be plain (no character style).",
                                excerpt, next.Pos);
                            issues++;
                        }
                    }
                }

                // ── RULE 5 — jnrltitle must be italic ────────────────────────
                if (tok.CS == CS_JNRLTITLE && !tok.Italic && HasAlpha(tok.Text))
                {
                    Msg("REF-PUNC", "ERROR",
                        $"Journal title \"{Truncate(tok.Text, 25)}\" has " +
                        $"jnrltitle style but is NOT italic.",
                        excerpt, tok.Pos);
                    issues++;
                }

                // ── RULE 6 — conferencetitle must be italic ───────────────────
                if (tok.CS == CS_CONFTITLE && !tok.Italic && HasAlpha(tok.Text))
                {
                    Msg("REF-PUNC", "ERROR",
                        $"Conference title \"{Truncate(tok.Text, 25)}\" has " +
                        $"conferencetitle style but is NOT italic.",
                        excerpt, tok.Pos);
                    issues++;
                }

                // ── RULE 7 — booktitle / booktitle0: italic + title case ──────
                if (IsBookTitle(tok.CS))
                {
                    // 7a — must be italic (skip punct-only runs)
                    if (!tok.Italic && HasAlpha(tok.Text))
                    {
                        Msg("REF-PUNC", "ERROR",
                            $"Book title \"{Truncate(tok.Text, 25)}\" has " +
                            $"booktitle style but is NOT italic.",
                            excerpt, tok.Pos);
                        issues++;
                    }

                    // 7b — must be title case (skip punct-only runs like em-dash)
                    string badWord7;
                    if (HasAlpha(tok.Text) && !IsTitleCase(tok.Text, out badWord7))
                    {
                        Msg("REF-PUNC", "WARNING",
                            $"Book title \"{Truncate(tok.Text, 30)}\" " +
                            $"may not be title case — check \"{badWord7}\".",
                            excerpt, tok.Pos);
                        issues++;
                    }
                }

                // ── RULE 8 — pubname must be title case ───────────────────────
                if (tok.CS == CS_PUBNAME)
                {
                    string badWord8;
                    if (HasAlpha(tok.Text) && !IsTitleCase(tok.Text, out badWord8))
                    {
                        Msg("REF-PUNC", "WARNING",
                            $"Publisher name \"{Truncate(tok.Text, 30)}\" " +
                            $"may not be title case — check \"{badWord8}\".",
                            excerpt, tok.Pos);
                        issues++;
                    }
                }

                // ── RULE 9 — articletitle must be sentence case ───────────────
                // Fire once per paragraph at the first articletitle token.
                // Uses the full concatenated title built in the pre-pass above.
                if (IsArticleTitle(tok.CS) && !articleTitleChecked
                    && HasAlpha(fullArticleTitle))
                {
                    articleTitleChecked = true; // only check once per paragraph

                    string offender;
                    if (!IsSentenceCase(fullArticleTitle, out offender))
                    {
                        Msg("REF-PUNC", "WARNING",
                            $"Article title \"{Truncate(fullArticleTitle, 35)}\" " +
                            $"may not be sentence case — " +
                            $"please check and confirm.",
                            excerpt, articleTitlePos);
                        issues++;
                    }
                }
            }

            return issues;
        }

        // =====================================================================
        // IS TITLE CASE
        //
        // Returns true if every word in text follows title-case rules:
        //   • First word of the string                → must be capitalised
        //   • First word after ':' or em/en dash      → must be capitalised
        //   • Words in LOWERCASE_WORDS (articles etc) → must be lowercase
        //   • All other words                         → must be capitalised
        //
        // Skips pure-digit tokens and all-caps abbreviations (II, GPM).
        // Outputs the first offending word via out parameter.
        // =====================================================================
        private static readonly System.Collections.Generic.HashSet<string>
            LowercaseWords =
            new System.Collections.Generic.HashSet<string>(
                StringComparer.OrdinalIgnoreCase)
        {
            "a","an","the","and","but","or","for","nor",
            "on","at","to","by","in","of","up","as",
            "is","it","from","with","into","onto","over",
            "than","that","this"
        };

        private static bool IsTitleCase(string text, out string badWord)
        {
            badWord = null;
            if (string.IsNullOrWhiteSpace(text)) return true;

            // Split on whitespace, keeping each token
            var words = System.Text.RegularExpressions.Regex.Split(
                text.Trim(), @"\s+");

            bool firstOfPhrase = true;

            foreach (string word in words)
            {
                if (string.IsNullOrEmpty(word)) continue;

                // Strip surrounding punctuation to get the core
                string core = word.Trim(
                    '.', ',', ';', ':', '!', '?', '(', ')',
                    '[', ']', '"', '\'', '\u2014', '\u2013', ' ');

                // Check phrase-reset: word ends with colon or dash
                bool willReset =
                    word.TrimEnd().EndsWith(":") ||
                    word.TrimEnd().EndsWith("\u2014") ||
                    word.TrimEnd().EndsWith("\u2013");

                if (string.IsNullOrEmpty(core))
                { firstOfPhrase = true; continue; }

                // Pure digits — skip (but respect colon-reset)
                if (IsAllDigits(core))
                { firstOfPhrase = willReset; continue; }

                // All-caps abbreviation (II, GPM, Vol.) — skip
                if (core.ToUpperInvariant() == core && core.Length >= 2)
                { firstOfPhrase = willReset; continue; }

                // No alphabetic characters — skip
                if (!HasAlpha(core))
                { firstOfPhrase = willReset; continue; }

                string coreLower = core.ToLowerInvariant();

                if (firstOfPhrase)
                {
                    // Must be capitalised
                    if (char.IsLower(core[0]))
                    { badWord = word; return false; }
                }
                else
                {
                    if (LowercaseWords.Contains(coreLower))
                    {
                        // Must stay lowercase
                        if (char.IsUpper(core[0]))
                        { badWord = word; return false; }
                    }
                    else
                    {
                        // Must be capitalised
                        if (char.IsLower(core[0]))
                        { badWord = word; return false; }
                    }
                }

                firstOfPhrase = willReset;
            }

            return true;
        }

        // =====================================================================
        // IS SENTENCE CASE
        //
        // Returns true if the text APPEARS to be sentence case.
        // Sentence-case rules:
        //   • First word of string         → must start with uppercase
        //   • First word after ':' or '—'  → must start with uppercase
        //   • All other words              → must start with lowercase
        //     EXCEPTIONS (not flagged):
        //       - All-caps tokens (ENSO, GPM, NATO, USA …)
        //       - Single-letter tokens (abbreviations)
        //       - Tokens containing digits
        //       - Tokens starting with a non-letter (punctuation, numbers)
        //
        // On first violation sets offender and returns false.
        // Caller emits ONE message and stops checking — does not enumerate
        // every bad word in the title.
        // =====================================================================
        private static bool IsSentenceCase(string text, out string offender)
        {
            offender = null;
            if (string.IsNullOrWhiteSpace(text)) return true;

            var words = System.Text.RegularExpressions.Regex
                .Split(text.Trim(), @"\s+");

            bool firstOfClause = true;

            foreach (string word in words)
            {
                if (string.IsNullOrEmpty(word)) continue;

                // Strip surrounding punctuation
                string core = word.Trim(
                    '.', ',', ';', ':', '!', '?', '(', ')',
                    '[', ']', '"', '\'', '\u2014', '\u2013', ' ');

                // Detect phrase reset (word ends with colon or em/en dash)
                bool willReset =
                    word.TrimEnd().EndsWith(":") ||
                    word.TrimEnd().EndsWith("\u2014") ||
                    word.TrimEnd().EndsWith("\u2013");

                if (string.IsNullOrEmpty(core))
                { firstOfClause = true; continue; }

                // Skip tokens with no alpha chars
                if (!HasAlpha(core))
                { firstOfClause = willReset; continue; }

                // Skip single-letter tokens (initials, abbreviations)
                if (core.Length == 1)
                { firstOfClause = willReset; continue; }

                // Skip all-caps tokens and acronym-like tokens
                // (NASA, GPM → all-caps; RNNs, GPMs → starts with 2+ uppercase)
                bool looksLikeAcronym = false;
                if (core == core.ToUpperInvariant())
                {
                    looksLikeAcronym = true;
                }
                else
                {
                    int upperPrefix = 0;
                    foreach (char c in core)
                    {
                        if (char.IsUpper(c)) upperPrefix++;
                        else break;
                    }
                    if (upperPrefix >= 2) looksLikeAcronym = true;
                }
                if (looksLikeAcronym)
                { firstOfClause = willReset; continue; }

                // Skip tokens containing digits (e.g. CO2, H2O, 3D)
                bool hasDigit = false;
                foreach (char c in core)
                    if (char.IsDigit(c)) { hasDigit = true; break; }
                if (hasDigit)
                { firstOfClause = willReset; continue; }

                if (firstOfClause)
                {
                    // Must start uppercase
                    if (char.IsLower(core[0]))
                    { offender = word; return false; }
                }
                else
                {
                    // Must start lowercase (proper nouns are a known false-positive
                    // risk — caller emits WARNING not ERROR)
                    if (char.IsUpper(core[0]))
                    { offender = word; return false; }
                }

                firstOfClause = willReset;
            }

            return true;
        }

        private static bool HasAlpha(string s)
        {
            foreach (char c in s)
                if (char.IsLetter(c)) return true;
            return false;
        }

        private static bool IsAllDigits(string s)
        {
            foreach (char c in s)
                if (!char.IsDigit(c)) return false;
            return true;
        }

        // Escape whitespace chars for display in messages
        private static string EscapeForMsg(string s) =>
            s.Replace(" ", "·").Replace("\t", "→");
    }
}