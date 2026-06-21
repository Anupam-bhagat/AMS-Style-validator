using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin.Checkers
{
    public class CasingErrorChecker
    {
        private const int ChunkSize = 30;
        private const int SleepMs = 20;

        // =====================================================================
        // SCOPED STYLES
        // =====================================================================
        private static readonly HashSet<string> ScopedStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "figurecaption","tablecaption","tablebody","tablehead",
            "paratext","acknowledgementtext","abstracttext","synopsis",
            "paranoindent","numberedlistitem","bulletedlistitem",
            "blockquot","formalarg","formalargend"
        };

        private static string NK(string s) =>
            (s ?? "").Replace(" ", "").Replace("-", "").Replace("_", "").ToLowerInvariant();

        // =====================================================================
        // MODELS
        // =====================================================================
        private class ParaData
        {
            public string Text;
            public int WordStart;
            public bool[] ItalicMap;   // length == Text.Length; null = unavailable
        }

        private struct Msg
        {
            public string Category, Severity, Text, Excerpt;
            public int Position;
        }

        // =====================================================================
        // SKIP SET — roman by convention, never flagged
        // =====================================================================
        private static readonly HashSet<string> Skip =
            new HashSet<string>(StringComparer.Ordinal)
        {
            "a","A","I","an","as","at","be","by","do","go","he","if","in","is","it",
            "me","my","no","of","on","or","so","to","up","us","we","are","and","but",
            "can","did","for","get","got","had","has","her","him","his","how","its",
            "let","may","not","now","off","old","one","our","out","own","put","say",
            "see","she","the","too","two","was","way","who","why","yet","you","also",
            "been","come","each","even","from","have","here","into","just","keep",
            "know","like","made","make","more","most","much","must","only","over",
            "same","some","such","than","that","them","then","they","this","thus",
            "time","used","very","well","were","what","when","with","your",
            // Standard math functions (roman by AMS convention)
            "sin","cos","tan","cot","sec","csc",
            "sinh","cosh","tanh","coth",
            "arcsin","arccos","arctan","arccot","arcsec","arccsc",
            "log","ln","exp","lg",
            "max","min","sup","inf","lim","limsup","liminf",
            "abs","det","div","curl","grad","nabla",
            "arg","ker","dim","deg","rank","trace","tr",
            "Re","Im","sgn","var","cov","corr","std",
            "obs","ref","tot","net","avg","rms",
            "mod","gcd","lcm","diag","span","col","row","null",
        };

        // =====================================================================
        // PRE-FILTER
        // =====================================================================
        private static readonly bool[] TriggerCharMap;   // bool[128], ASCII only

        // =====================================================================
        // REGEX BUILDING BLOCKS
        // =====================================================================
        private const string G = "[A-Za-z\u03b1-\u03c9\u0391-\u03a9]";       // one letter
        private const string NB = "(?<![A-Za-z\u03b1-\u03c9\u0391-\u03a9])";  // not preceded by letter
        private const RegexOptions CI = RegexOptions.Compiled | RegexOptions.IgnoreCase;

        // =====================================================================
        // REGEX A — hyphen rule only:
        // Flag a single letter immediately before a hyphen followed by a space
        // or a word (e.g. "x-axis", "y-direction", "z- component").
        // =====================================================================
        private static readonly Regex RxA = new Regex(
            NB + @"(?<hyp>" + G + @")-(?:\s|[A-Za-z]\w*)",
            CI);

        private static readonly string[] GNames = { "hyp" };
        private static readonly string[] GTemplates = { "Variable \"{0}\" before hyphen is not italic — variable missed to make italic." };

        // =====================================================================
        // STATIC CONSTRUCTOR
        // =====================================================================
        static CasingErrorChecker()
        {
            TriggerCharMap = new bool[128];
            TriggerCharMap['-'] = true;
        }

        // =====================================================================
        // PUBLIC ENTRY POINT
        // =====================================================================
        public void Run()
        {
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch { return; }

            string docPath = "";
            try { docPath = doc.FullName; } catch { }

            bool canOoxml = !string.IsNullOrEmpty(docPath) &&
                            File.Exists(docPath) &&
                            (docPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                             docPath.EndsWith(".docm", StringComparison.OrdinalIgnoreCase));

            TaskPaneWinForms.SetProgress("Collecting scoped paragraphs…");

            // ── Phase 1 (COM): collect text + WordStart only ──────────────────
            var paras = new List<ParaData>(256);
            bool wasOn = true;
            try
            {
                wasOn = doc.Application.ScreenUpdating;
                doc.Application.ScreenUpdating = false;

                int pi = 0, total = 0;
                try { total = doc.Paragraphs.Count; } catch { }

                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    pi++;
                    if (pi % ChunkSize == 0)
                    {
                        TaskPaneWinForms.SetProgress(total > 0
                            ? $"Scanning paragraph {pi} of {total}…"
                            : $"Scanning paragraph {pi}…");
                        doc.Application.ScreenUpdating = wasOn;
                        Thread.Sleep(SleepMs);
                        doc.Application.ScreenUpdating = false;
                    }

                    string sk;
                    try { sk = NK(para.get_Style().NameLocal); } catch { continue; }
                    if (!ScopedStyles.Contains(sk)) continue;

                    Word.Range rng;
                    try { rng = para.Range; } catch { continue; }

                    string txt; int ws;
                    try { txt = (rng.Text ?? "").TrimEnd('\r', '\n'); ws = rng.Start; }
                    catch { continue; }

                    if (string.IsNullOrWhiteSpace(txt) || !HasTrigger(txt)) continue;

                    paras.Add(new ParaData { Text = txt, WordStart = ws });
                }
            }
            finally { try { doc.Application.ScreenUpdating = wasOn; } catch { } }

            if (paras.Count == 0)
            {
                TaskPaneWinForms.AddMessage("CASINGCHECK", "WARNING",
                    "No scoped paragraphs found with variable-like content.");
                return;
            }

            TaskPaneWinForms.SetProgress($"Analysing {paras.Count} paragraphs…");

            // ── Phase 2 (background): OOXML italic fill + parallel regex ─────
            Task.Run(() =>
            {
                try
                {
                    if (canOoxml) FillItalicMapsFromOoxml(docPath, paras);

                    int cores = Math.Max(1, Environment.ProcessorCount - 1);
                    int chunkSize = Math.Max(1, (paras.Count + cores - 1) / cores);
                    var partitions = new List<List<ParaData>>();
                    for (int i = 0; i < paras.Count; i += chunkSize)
                    {
                        int end = Math.Min(i + chunkSize, paras.Count);
                        partitions.Add(paras.GetRange(i, end - i));
                    }

                    var allBatches = new List<Msg>[partitions.Count];
                    Parallel.For(0, partitions.Count, pi =>
                    {
                        var batch = new List<Msg>(32);
                        var rep = new HashSet<int>();
                        foreach (var p in partitions[pi])
                        {
                            rep.Clear();
                            CheckPara(p, batch, rep);
                        }
                        allBatches[pi] = batch;
                    });

                    bool any = false;
                    foreach (var b in allBatches)
                        if (b != null)
                            foreach (var m in b)
                            {
                                TaskPaneWinForms.AddMessage(m.Category, m.Severity,
                                                            m.Text, m.Excerpt, m.Position);
                                any = true;
                            }

                    if (!any)
                        TaskPaneWinForms.AddMessage("CASINGCHECK", "INFO",
                            "Casing check passed — no un-italicised variables found.");
                }
                catch (Exception ex)
                {
                    TaskPaneWinForms.AddMessage("CASINGCHECK", "ERROR",
                        "Casing checker error: " + ex.Message);
                }
            });
        }

        // =====================================================================
        // OOXML — fill ItalicMap for all paragraphs in ONE streaming XML read.
        // =====================================================================
        private static void FillItalicMapsFromOoxml(string docPath, List<ParaData> paras)
        {
            try
            {
                var lookup = new Dictionary<string, int>(paras.Count,
                                 StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < paras.Count; i++)
                {
                    string key = CollapseSpaces(paras[i].Text);
                    if (!lookup.ContainsKey(key)) lookup[key] = i;
                }

                int remaining = lookup.Count;

                using (var pkg = System.IO.Packaging.Package.Open(
                           docPath, FileMode.Open, FileAccess.Read))
                {
                    var uri = new Uri("/word/document.xml", UriKind.Relative);
                    if (!pkg.PartExists(uri)) return;

                    using (var stream = pkg.GetPart(uri).GetStream())
                    using (var reader = XmlReader.Create(stream,
                               new XmlReaderSettings { IgnoreWhitespace = false }))
                    {
                        const string WNS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                        var paraText = new StringBuilder(512);
                        var runs = new List<(int start, int len, bool italic)>(32);
                        bool inPara = false;
                        bool inRun = false;
                        bool curItalic = false;
                        int runStart = 0;
                        var runText = new StringBuilder(128);

                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                string local = reader.LocalName;
                                string ns = reader.NamespaceURI;
                                if (ns != WNS) continue;

                                switch (local)
                                {
                                    case "p":
                                        inPara = true;
                                        paraText.Clear();
                                        runs.Clear();
                                        break;

                                    case "r":
                                        if (!inPara) break;
                                        inRun = true;
                                        curItalic = false;
                                        runText.Clear();
                                        runStart = paraText.Length;
                                        break;

                                    case "i":
                                        if (inRun)
                                        {
                                            string val = reader.GetAttribute("val", WNS);
                                            curItalic = val == null ||
                                                         val == "true" ||
                                                         val == "1";
                                        }
                                        break;

                                    case "t":
                                        if (inRun && reader.Read() &&
                                            reader.NodeType == XmlNodeType.Text)
                                        {
                                            runText.Append(reader.Value);
                                        }
                                        break;
                                }
                            }
                            else if (reader.NodeType == XmlNodeType.EndElement)
                            {
                                string local = reader.LocalName;
                                string ns = reader.NamespaceURI;
                                if (ns != WNS) continue;

                                if (local == "r" && inRun)
                                {
                                    inRun = false;
                                    if (runText.Length > 0)
                                    {
                                        runs.Add((runStart, runText.Length, curItalic));
                                        paraText.Append(runText);
                                    }
                                }
                                else if (local == "p" && inPara)
                                {
                                    inPara = false;
                                    string key = CollapseSpaces(paraText.ToString().TrimEnd('\r', '\n'));
                                    if (lookup.TryGetValue(key, out int idx))
                                    {
                                        var pd = paras[idx];
                                        int len = pd.Text.Length;
                                        var map = new bool[len];
                                        foreach (var (rs, rl, italic) in runs)
                                        {
                                            if (!italic) continue;
                                            int end = Math.Min(rs + rl, len);
                                            for (int i = rs; i < end; i++) map[i] = true;
                                        }
                                        pd.ItalicMap = map;
                                        lookup.Remove(key);
                                        if (--remaining == 0) return;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch { /* fall back: all ItalicMap stay null → conservative flagging */ }
        }

        private static string CollapseSpaces(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            for (int i = 0; i < s.Length - 1; i++)
                if (char.IsWhiteSpace(s[i]) && char.IsWhiteSpace(s[i + 1]))
                    goto slow;
            return s.Trim();
        slow:
            var sb = new StringBuilder(s.Length);
            bool sp = false;
            foreach (char c in s)
            {
                if (char.IsWhiteSpace(c)) { if (!sp) { sb.Append(' '); sp = true; } }
                else { sb.Append(c); sp = false; }
            }
            return sb.ToString().Trim();
        }

        // =====================================================================
        // PRE-FILTER — only checks for hyphen character
        // =====================================================================
        private static bool HasTrigger(string t)
        {
            foreach (char c in t)
                if (c < 128 && TriggerCharMap[c]) return true;
            return false;
        }

        // =====================================================================
        // CHECK ONE PARAGRAPH — hyphen rule only
        // =====================================================================
        private static void CheckPara(ParaData p, List<Msg> batch, HashSet<int> rep)
        {
            string t = p.Text;

            foreach (Match m in RxA.Matches(t))
            {
                Group g = m.Groups["hyp"];
                if (!g.Success) continue;

                string val = g.Value;
                int idx = g.Index;
                if (Skip.Contains(val) || !rep.Add(idx) || IsItalic(p, idx)) continue;

                string msg = GTemplates[0].Replace("{0}", val);
                AddMsg(batch, p, m.Index, m.Length, msg);
            }
        }

        // =====================================================================
        // ITALIC CHECK
        // =====================================================================
        [System.Runtime.CompilerServices.MethodImpl(
            System.Runtime.CompilerServices.MethodImplOptions.AggressiveInlining)]
        private static bool IsItalic(ParaData p, int idx)
        {
            bool[] m = p.ItalicMap;
            return m != null && (uint)idx < (uint)m.Length && m[idx];
        }

        // =====================================================================
        // HELPERS
        // =====================================================================
        private static void AddMsg(List<Msg> batch, ParaData p, int idx, int len, string text)
        {
            batch.Add(new Msg
            {
                Category = "CASINGCHECK",
                Severity = "WARNING",
                Text = text,
                Excerpt = Excerpt(p.Text, idx, len),
                Position = p.WordStart + idx,
            });
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
    }
}