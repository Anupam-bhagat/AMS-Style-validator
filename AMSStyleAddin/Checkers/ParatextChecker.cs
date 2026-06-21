using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace AMSStyleAddin.Checkers
{
    public class ParatextChecker
    {
        private const int ChunkSize = 30;
        private const int SleepMs = 20;

        // =====================================================================
        // STYLE SETS
        // =====================================================================

        /// <summary>All paragraph styles that this checker inspects.</summary>
        private static readonly HashSet<string> ScopedStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "figurecaption","tablecaption","tablebody","tablehead",
            "paratext","acknowledgementtext","abstracttext","synopsis",
            "paranoindent","numberedlistitem","bulletedlistitem",
            "blockquot","formalarg","formalargend"
        };

        /// <summary>
        /// Styles where months MUST be abbreviated (AMS captions / table cells).
        /// Rule: figure captions, table captions, table body, table head → use Jan/Feb/…
        /// </summary>
        private static readonly HashSet<string> AbbrevMonthStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "figurecaption", "tablecaption", "tablebody", "tablehead" };

        /// <summary>
        /// Styles where months MUST be spelled out (running text).
        /// Rule: paratext and similar body styles → use January/February/…
        /// </summary>
        private static readonly HashSet<string> SpellOutMonthStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "paratext","acknowledgementtext","abstracttext","synopsis",
            "paranoindent","numberedlistitem","bulletedlistitem",
            "blockquot","formalarg","formalargend"
        };

        /// <summary>
        /// Styles where date-format rules (wrong order, ordinal suffixes, etc.) apply.
        /// Date rules only apply where AMS date formatting is required:
        ///   figure captions, table captions, table body, table head, and body paratext.
        /// </summary>
        private static readonly HashSet<string> DateRuleStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "figurecaption","tablecaption","tablebody","tablehead",
            "paratext","acknowledgementtext","abstracttext","synopsis",
            "paranoindent","numberedlistitem","bulletedlistitem",
            "blockquot","formalarg","formalargend"
        };

        private static readonly HashSet<string> TableStyles =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "tablebody", "tablehead" };

        // =====================================================================
        // HELPERS
        // =====================================================================

        private static string NK(string s) =>
            (s ?? "").Replace(" ", "").Replace("-", "")
                     .Replace("_", "").ToLowerInvariant();

        private static readonly HashSet<string> SkipStandaloneVars =
            new HashSet<string>(StringComparer.Ordinal) { "a", "I" };

        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // =====================================================================
        // DATA CLASSES
        // =====================================================================

        private class OoxmlRun
        {
            public string Text;
            public bool Bold, Italic, Superscript, Subscript;
            public string FontName;
        }

        private class OoxmlPara
        {
            public string StyleKey, Text;
            public List<OoxmlRun> Runs;
            public int WordStart;
        }

        private class AllPara
        {
            public string SK;
            public bool IsEmpty;
            public int WordStart;
            public string Text;
        }

        // =====================================================================
        // FONT CATEGORIES
        // =====================================================================
        private static readonly HashSet<string> SymbolFonts =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "Symbol","Wingdings","Wingdings 2","Wingdings 3","Webdings","Marlett","MT Extra" };

        private static readonly HashSet<string> MonospaceFonts =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "Courier","Courier New","Courier Final Draft","Lucida Console","Consolas","Monaco","Menlo","Letter Gothic","OCR A Extended" };

        private static readonly HashSet<string> SansSerifFonts =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "Arial","Arial Narrow","Arial Black","Helvetica","Helvetica Neue","Calibri","Candara","Corbel","Gill Sans","Tahoma","Verdana","Trebuchet MS","Franklin Gothic","Century Gothic","Segoe UI","Myriad Pro","Futura" };

        private static readonly HashSet<string> NonLatinFonts =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "Dengxian","DengXian","SimSun","SimHei","KaiTi","FangSong","MingLiU","PMingLiU","NSimSun","Microsoft YaHei","Microsoft JhengHei","Yu Gothic","Yu Mincho","MS Gothic","MS Mincho","Malgun Gothic","Batang","Dotum","Gulim","Arial Unicode MS" };

        private static readonly HashSet<string> LatexFonts =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "CMR10","CMR12","CMMI10","CMMI12","CMSY10","CMEX10","Computer Modern","CM Roman","Latin Modern Roman","LM Roman","LMRoman10-Regular","Latin Modern Math","STIX","STIX Two","XITS","Asana Math","TeX Gyre","TeX Gyre Termes","TeX Gyre Pagella","Symbola" };

        private static readonly HashSet<string> WordDefaultFonts =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "Cambria","Cambria Math","Calibri Light","Times","Times Roman","NimbusRomNo9L" };

        private static string GetFontMessage(string fontName)
        {
            if (SymbolFonts.Contains(fontName))
                return "Font \"" + fontName + "\" found — use Unicode characters instead of Symbol/Wingdings fonts (e.g. use α U+03B1 not Symbol font). Change font to Times New Roman.";
            if (MonospaceFonts.Contains(fontName))
                return "Font \"" + fontName + "\" found — monospace fonts are not permitted in AMS body text. Change to Times New Roman.";
            if (SansSerifFonts.Contains(fontName))
                return "Font \"" + fontName + "\" found — sans-serif fonts are not permitted in AMS body text. Change to Times New Roman.";
            if (NonLatinFonts.Contains(fontName))
                return "Font \"" + fontName + "\" found — non-Latin/CJK font detected in body text. Change to Times New Roman.";
            if (LatexFonts.Contains(fontName))
                return "Font \"" + fontName + "\" found — possible LaTeX conversion artifact. Change to Times New Roman.";
            if (WordDefaultFonts.Contains(fontName))
                return "Font \"" + fontName + "\" found — Word default font detected. AMS body text must use Times New Roman.";
            return "Font \"" + fontName + "\" found — all AMS body text must use Times New Roman. Change font to Times New Roman.";
        }

        // =====================================================================
        // REGEXES
        // =====================================================================
        private const string NumOrWord =
            @"(?:\d[\d.,]{0,20}|zero|one|two|three|four|five|six|seven|eight|nine|ten|" +
            @"eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|" +
            @"twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|hundred|thousand)";
        private const string OptUnit =
            @"(?:\s*[°%µ][A-Za-z]{0,5}|\s+[A-Za-z][A-Za-z0-9/°µ]{0,10}){0,2}";
        private const string AnyDash = @"[-\u2013\u2014]";
        private const string MonthsAll =
            @"January|February|March|April|May|June|July|August|September|" +
            @"October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec";

        private static readonly HashSet<string> MonthAbbrevSet =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { "Jan","Feb","Mar","Apr","Jun","Jul","Aug","Sep","Sept","Oct","Nov","Dec" };

        private static readonly Regex RxBetweenDash = new Regex(
            @"between\s+" + NumOrWord + OptUnit + @"\s*" + AnyDash + @"\s*" + NumOrWord + OptUnit,
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxBetweenTo = new Regex(
            @"between\s+" + NumOrWord + @"\s*\w{0,10}\s+to\s+" + NumOrWord,
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxFromDash = new Regex(
            @"from\s+" + NumOrWord + OptUnit + @"\s*" + AnyDash + @"\s*" + NumOrWord + OptUnit,
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxSymAsym = new Regex(
            @"(?<!\d)(\d[\d.,]{0,20})(?!\s*(?:°[CFcf]?|%|\u2030|(?:st|nd|rd|th)\b))" +
            @"\s*[±+\-\u2212\u00b7\u2013\u2014]\s*(\d[\d.,]{0,20})\s*(°[CFcf]?|%|\u2030|(?:st|nd|rd|th)\b)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxUtcOn = new Regex(@"\bUTC\s+on\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxHttp = new Regex(@"\bhttp://", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxSciNotation = new Regex(@"\b\d+(\.\d+)?[eE][+\-]?\d+\b", RegexOptions.Compiled);
        private static readonly Regex RxLeadingZero = new Regex(@"(?<![a-zA-Z\d])\.(\d)", RegexOptions.Compiled);
        private static readonly Regex RxFourDigitComma = new Regex(@"\b(\d),(\d{3})\b", RegexOptions.Compiled);
        private static readonly Regex RxPercentile = new Regex(@"\b(\d+(?:\.\d+)?)\s+percentile\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxMyrKyr = new Regex(@"\b(MYR|mYR|MYr|KYR|kYR|KYr|KA|mA|MA|Ka)\b", RegexOptions.Compiled);

        // ---- New rule: no space allowed between a number and "%" ------------
        // AMS style requires the percent sign to be glued directly to its number,
        // e.g. "36.0%" not "36.0 %". Captures the full numeric run (including any
        // decimal point / thousands separators) so the suggested fix shows the
        // complete number, not just the last digit before the space.
        private static readonly Regex RxSpaceBeforePercent = new Regex(
            @"(\d[\d.,]*)\s+%", RegexOptions.Compiled);
        // ---- End new rule -----------------------------------------------------

        // ---- New rule: missing space after a list-separator comma between
        // numbers, e.g. "Ef = 0.2,0.5, 0.8" — the comma after "0.2" has no
        // following space while the comma after "0.5" does. AMS requires a
        // space after every comma in a numeric list/series.
        // The leading group "(\d[\d.,]*)" captures the full preceding number
        // (so the suggested fix can echo it back). The negative lookahead
        // "(?!\d{3}\b)" excludes the unrelated four-digit thousands-separator
        // case (e.g. "1,234") which RxFourDigitComma already reports as a
        // different problem (delete the comma entirely) — without this
        // exclusion the same comma would generate two conflicting warnings.
        // The trailing "(?=\d)" requires the very next character after the
        // comma to be a digit with literally no space, which is the actual
        // defect being flagged; a comma already followed by a space never
        // matches.
        private static readonly Regex RxMissingSpaceAfterComma = new Regex(
            @"(\d[\d.,]*),(?!\d{3}\b)(?=\d)", RegexOptions.Compiled);
        // ---- End new rule -----------------------------------------------------

        // ---- Date regexes ---------------------------------------------------
        // Wrong order: "January 21, 2017" or "January 21 2017"
        private static readonly Regex RxWrongDate = new Regex(
            @"\b(" + MonthsAll + @")\s+\d{1,2}[,\s]+\d{4}\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        // Wrong order with day first but abbreviated month or commas: "21, January 2017"
        private static readonly Regex RxWrongDateDay = new Regex(
            @"\b(\d{1,2})(,?)\s+(" + MonthsAll + @")(,?)\s+(\d{4})\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        // Ordinal suffix on day number: "21st January" or "January 21st"
        private static readonly Regex RxOrdinalDate = new Regex(
            @"\b(\d{1,2})\s*(st|nd|rd|th)\b(?=[,\s]*(?:" + MonthsAll + @"))" +
            @"|\b(?:" + MonthsAll + @")\s+(\d{1,2})\s*(st|nd|rd|th)\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        // ---- New rule: redundant date range — "14 Sep 2024 to 24 Sep 2024"
        // repeats the month and year on both ends of the range. AMS style
        // consolidates this to "14 to 24 Sep 2024" (month/year stated once,
        // at the end). Groups: 1=day1, 2=month, 3=year, 4=day2. Backreferences
        // \2 and \3 require the second occurrence to repeat the EXACT same
        // month token and year — if the month or year differ (e.g. a range
        // spanning two months), this rule correctly does not fire, since there
        // is nothing redundant to consolidate.
        private static readonly Regex RxRedundantDateRange = new Regex(
            @"\b(\d{1,2})\s+(" + MonthsAll + @")\s+(\d{4})\s+to\s+(\d{1,2})\s+\2\s+\3\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        // ---- End date regexes -----------------------------------------------

        private static readonly Regex RxCycles = new Regex(@"\bcycles?\s*(per|/)\s*(second|sec|s|min|minute|hour|hr)\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxMicron = new Regex(@"\bmicrons?\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxMuUnit = new Regex(@"\bmu\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxArcSec = new Regex(@"\b(\d+(?:[.,]\d+)?)\s+arc[\s\-]seconds?\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxTempBothUnits = new Regex(@"\b(\d+(?:[.,]\d+)?)\s*(°[CF])\s+and\s+(\d+(?:[.,]\d+)?)\s*(°[CF])\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex RxPlusMinusAsym = new Regex(@"\b(\d+(?:[.,]\d+)?)\s*[±\u00b1]\s*(\d+(?:[.,]\d+)?)\s*(°[CF]?|%)", RegexOptions.Compiled);
        private static readonly Regex RxSeriesSymbol = new Regex(@"\b(\d+(?:[.,]\d+)?)\s*,\s*(?:\d+(?:[.,]\d+)?\s*,\s*)*(?:and\s+)?\d+(?:[.,]\d+)?\s*(°|%)", RegexOptions.Compiled);
        // FALSE-POSITIVE FIX: the old pattern @"\b[Pp]anel\b" fired on every generic
        // use of the word "panel" (e.g. "...are listed below the panel."), not just
        // the AMS issue of writing "Panel (a)" instead of "(a)". Now requires the
        // word to be immediately followed by "(<letter>)" (optionally without the
        // parentheses), e.g. "Panel (a)", "panel a)", "Panels (a)".
        private static readonly Regex RxPanel = new Regex(@"\b[Pp]anels?\b\s*\(?[a-zA-Z]\)", RegexOptions.Compiled);
        private static readonly Regex RxCaptionDirections = new Regex(@"\(\s*(top|bottom|left|right|Top|Bottom|Left|Right|TOP|BOTTOM|LEFT|RIGHT)\s*\)", RegexOptions.Compiled);

        // Month matching — groups: [1] = all months except May, [2] = "May" followed by digit
        // (May alone is a valid English word; only flag it as a month when followed by a number)
        private static readonly Regex RxAnyMonth = new Regex(
            @"\b(January|February|March|April|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)\b|\b(May)\b(?=\s+\d)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex RxUnitTypos;
        private static readonly Regex RxNumUnit;
        private static readonly Regex RxRepeatedUnit;
        private static readonly Regex RxUnitUnit;
        private static readonly Regex RxSolidusAll;

        // =====================================================================
        // UNIT DATA
        // =====================================================================
        private static readonly Dictionary<string, string> UnitTypos =
            new Dictionary<string, string>(StringComparer.Ordinal)
            {
                {"Kg","kg"},{"KG","kg"},{"hz","Hz"},{"HZ","Hz"},
                {"Sec","s"},{"sec","s"},{"SEC","s"},
                {"Mhz","MHz"},{"MHZ","MHz"},{"mhz","MHz"},
                {"Ghz","GHz"},{"GHZ","GHz"},{"ghz","GHz"},
                {"Khz","kHz"},{"KHz","kHz"},{"khz","kHz"},
                {"Km","km"},{"KM","km"},{"Cm","cm"},{"CM","cm"},
                {"Mm","mm"},{"MM","mm"},{"Mg","mg"},{"MG","mg"},
                {"Kj","kJ"},{"KJ","kJ"},{"Kw","kW"},{"KW","kW"},
                {"Kv","kV"},{"KV","kV"},
                {"Kpa","kPa"},{"KPA","kPa"},{"kpa","kPa"},
                {"Hpa","hPa"},{"HPA","hPa"},{"hpa","hPa"},
                {"Dbz","dBZ"},{"DBZ","dBZ"},{"dbz","dBZ"},
                {"Ppm","ppm"},{"PPM","ppm"},{"Ppb","ppb"},{"PPB","ppb"},
                {"Ppt","ppt"},{"PPT","ppt"},
                {"hr","h"},{"HR","h"},
                {"Inches","in."},{"inches","in."},{"Inch","in."},{"inch","in."},
            };

        private static readonly HashSet<string> KnownUnits =
            new HashSet<string>(StringComparer.Ordinal)
        {
            "m","km","cm","mm","nm","g","kg","mg","s","ms","min","h","hr","yr","d",
            "K","Pa","hPa","kPa","MPa","W","kW","MW","J","kJ","MJ",
            "Hz","kHz","MHz","GHz","N","mol","L","mL","dB","dBZ","kt","Sv","PVU","au","BTU"
        };

        // FALSE-POSITIVE FIX (Rule R below): single-letter entries in KnownUnits
        // (K, g, N, h, W, L, m, s, t, d, J, ...) are also extremely common italic
        // scalar-variable names in AMS physical-science prose — e.g. italic "K"
        // (stationary wavenumber Ks), "g" (gravitational acceleration), "N"
        // (Brunt–Väisälä frequency), "h"/"L" (moist static energy / latent heat
        // terms, e.g. <h>B), "W" (vertical velocity). Rule R previously flagged all
        // of these correctly-italicized variables as "unit should not be italic".
        // Restrict Rule R to multi-character unit symbols (km, hPa, kPa, MHz, dBZ,
        // etc.), which are unambiguous and never used as bare variable names.
        private static readonly HashSet<string> MultiCharKnownUnits =
            new HashSet<string>(KnownUnits.Where(u => u.Length > 1), StringComparer.Ordinal);

        private static readonly Dictionary<string, string> UnitAbbrevMap =
            new Dictionary<string, string>(StringComparer.Ordinal)
        {
            {"Sv","sverdrup"},{"kt","knot"},{"PVU","potential vorticity unit"},
            {"n mi","nautical mile"},{"BTU","British thermal unit"},{"au","astronomical unit"},
        };
        private static readonly Dictionary<string, Regex> UnitPatterns =
            new Dictionary<string, Regex>(StringComparer.Ordinal)
        {
            {"Sv",  new Regex(@"\bSv\b",    RegexOptions.Compiled)},
            {"kt",  new Regex(@"\bkt\b",    RegexOptions.Compiled)},
            {"PVU", new Regex(@"\bPVU\b",   RegexOptions.Compiled)},
            {"n mi",new Regex(@"\bn\s+mi\b",RegexOptions.Compiled)},
            {"BTU", new Regex(@"\bBTU\b",   RegexOptions.Compiled)},
            // FALSE-POSITIVE FIX: "\bau\b" matched the Australian top-level domain in
            // email addresses / URLs such as "linyuan.sun@unsw.edu.au" or
            // "www.bom.gov.au" (the preceding "." is a non-word character, so \b
            // matched right before "au"). Add a negative lookbehind so "au" preceded
            // by "." (".edu.au", ".gov.au", ".com.au", etc.) is not treated as the
            // "astronomical unit" abbreviation.
            {"au",  new Regex(@"(?<!\.)\bau\b", RegexOptions.Compiled)},
        };
        private static readonly Dictionary<string, Regex> UnitFullNamePatterns =
            new Dictionary<string, Regex>(StringComparer.Ordinal)
        {
            {"Sv",  new Regex(@"\bsverdrup",                    RegexOptions.Compiled|RegexOptions.IgnoreCase)},
            {"kt",  new Regex(@"\bknots?\b",                    RegexOptions.Compiled|RegexOptions.IgnoreCase)},
            {"PVU", new Regex(@"\bpotential\s+vorticity\s+unit",RegexOptions.Compiled|RegexOptions.IgnoreCase)},
            {"n mi",new Regex(@"\bnautical\s+mil",              RegexOptions.Compiled|RegexOptions.IgnoreCase)},
            {"BTU", new Regex(@"\bbritish\s+thermal\s+unit",    RegexOptions.Compiled|RegexOptions.IgnoreCase)},
            {"au",  new Regex(@"\bastronomical\s+unit",         RegexOptions.Compiled|RegexOptions.IgnoreCase)},
        };

        // Canonical month lookups
        private static readonly Dictionary<string, string> MonthAbbrev =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            {"January","Jan"},{"February","Feb"},{"March","Mar"},{"April","Apr"},
            {"June","Jun"},{"July","Jul"},{"August","Aug"},{"September","Sep"},
            {"October","Oct"},{"November","Nov"},{"December","Dec"}
        };
        private static readonly Dictionary<string, string> MonthCanonical =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            {"January","January"},{"Jan","January"},{"February","February"},{"Feb","February"},
            {"March","March"},{"Mar","March"},{"April","April"},{"Apr","April"},{"May","May"},
            {"June","June"},{"Jun","June"},{"July","July"},{"Jul","July"},
            {"August","August"},{"Aug","August"},{"September","September"},
            {"Sep","September"},{"Sept","September"},{"October","October"},{"Oct","October"},
            {"November","November"},{"Nov","November"},{"December","December"},{"Dec","December"},
        };

        // =====================================================================
        // STATIC CONSTRUCTOR
        // =====================================================================
        static ParatextChecker()
        {
            var keys = new List<string>(UnitTypos.Keys);
            keys.Sort((a, b) => b.Length.CompareTo(a.Length));
            RxUnitTypos = new Regex(
                @"\b(" + string.Join("|", keys.ConvertAll(Regex.Escape)) + @")\b",
                RegexOptions.Compiled);

            var units = new List<string>
            {
                "GtCO2","ppmv","GtC","kVA","kHz","MHz","GHz","mbar","dBZ","hPa","kPa",
                "gpm","kDa","keV","dbar","ppb","ppm","ppt","km","cm","mm","kg","mg",
                "kJ","kW","kV","mb","ha","ft","Hz","Pa","Sv","kt","cfs","yr","min",
                "bar","W","J","K","L","N","s","h","t","m","g","%"
            };
            units.Sort((a, b) => b.Length.CompareTo(a.Length));
            var eu = units.ConvertAll(Regex.Escape);
            string ua = string.Join("|", eu);
            RxNumUnit = new Regex(@"\d(" + ua + @")\b", RegexOptions.Compiled);
            RxRepeatedUnit = new Regex(
                @"\b(\d[\d.,]*)\s+(" + ua + @")(?![\u2212\u207b\-\^])\b.{1,25}\b\d[\d.,]*\s+\2(?![\u2212\u207b\-\^])\b",
                RegexOptions.Compiled);

            var ku = new List<string>(KnownUnits);
            ku.Sort((a, b) => b.Length.CompareTo(a.Length));
            string kuAlt = string.Join("|", ku.ConvertAll(Regex.Escape));
            // FALSE-POSITIVE FIX: the old pattern @"(?<![A-Za-z\d])(?:KU)(?:KU)(exp)?(?![A-Za-z])"
            // matched ANY two adjacent KnownUnits tokens with NO numeric context and an
            // OPTIONAL exponent suffix. In real text this fired on:
            //   - legitimate single compound units that happen to be decomposable into
            //     two shorter KnownUnits, e.g. "200 hPa" -> "h"+"Pa", "5 mm" -> "m"+"m"
            //   - variable names that coincidentally look like two unit symbols, e.g.
            //     "(Ks)" -> "K"+"s", "<LW>" -> "L"+"W", "uB" "hL" -> "h"+"L",
            //     and even author names like "Lau" -> "L"+"au"
            // The genuine AMS issue this rule targets is a value glued to a compound
            // unit immediately before a negative exponent, e.g. "10ms⁻¹" which should
            // be "10 m s⁻¹". Requiring (a) an immediately preceding digit (optionally
            // with one space) and (b) a MANDATORY exponent suffix eliminates all of the
            // false positives above while still catching the "10ms⁻¹" pattern.
            RxUnitUnit = new Regex(
                @"(?<=\d)\s?(?:" + kuAlt + @")(?:" + kuAlt + @")" +
                @"(?:[-\u2212][\u00b9\u00b2\u00b3\d]|[\u207b][\u00b9\u00b2\u00b3\d]|[\u00b9\u00b2\u00b3])(?![A-Za-z])",
                RegexOptions.Compiled);
            RxSolidusAll = new Regex(
                @"(?<![A-Za-z\d])(?:" + kuAlt + @")\s*/\s*(?:" + kuAlt + @")\d?(?![A-Za-z])",
                RegexOptions.Compiled);
        }

        // =====================================================================
        // RUN
        // =====================================================================
        public void Run()
        {
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch { return; }

            string docPath = "";
            try { docPath = doc.FullName; } catch { }

            bool canUseOoxml = !string.IsNullOrEmpty(docPath) &&
                               File.Exists(docPath) &&
                               (docPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase) ||
                                docPath.EndsWith(".docm", StringComparison.OrdinalIgnoreCase));

            int totalParas = 0;
            try { totalParas = doc.Paragraphs.Count; } catch { }

            var scopedParas = new List<OoxmlPara>();
            var allTexts = new List<(string Text, int Start)>();
            var allParas = new List<AllPara>();

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
                    string txt = ""; int wordStart = 0;
                    try { txt = rng.Text ?? ""; wordStart = rng.Start; } catch { continue; }

                    string sk = "";
                    try { sk = NK(para.get_Style().NameLocal); } catch { continue; }

                    string trimmed = txt.TrimEnd('\r', '\n');
                    bool isEmpty = string.IsNullOrWhiteSpace(trimmed);

                    allParas.Add(new AllPara { SK = sk, IsEmpty = isEmpty, WordStart = wordStart, Text = trimmed });

                    if (isEmpty) continue;
                    allTexts.Add((trimmed, wordStart));
                    if (!ScopedStyles.Contains(sk)) continue;

                    scopedParas.Add(new OoxmlPara
                    {
                        StyleKey = sk,
                        Text = trimmed,
                        WordStart = wordStart,
                        Runs = new List<OoxmlRun>()
                    });
                }
            }
            finally { try { doc.Application.ScreenUpdating = screenWasOn; } catch { } }

            TaskPaneWinForms.SetProgress("Analysing text…");

            if (scopedParas.Count == 0)
            {
                TaskPaneWinForms.AddMessage("PARATEXT", "WARNING",
                    "No scoped paragraphs found. Make sure the document uses AMS styles.");
                return;
            }

            Task.Run(() =>
            {
                try
                {
                    if (canUseOoxml) FillRunsFromOoxml(docPath, scopedParas);
                    int found = 0;
                    found += RunTextRules(scopedParas);
                    found += RunFormattingRules(scopedParas);
                    found += RunDocumentRules(scopedParas, allTexts);
                    found += RunTableBreakRule(allParas);
                    if (found == 0)
                        TaskPaneWinForms.AddMessage("PARATEXT", "INFO",
                            "Paratext check passed — no issues found in scoped paragraphs.");
                }
                catch (Exception ex)
                {
                    TaskPaneWinForms.AddMessage("PARATEXT", "ERROR",
                        "Paratext checker error: " + ex.Message);
                }
            });
        }

        // =====================================================================
        // RULE TB — PARAGRAPH BREAK INSIDE TABLE
        // =====================================================================
        private int RunTableBreakRule(List<AllPara> allParas)
        {
            int found = 0;
            for (int i = 1; i < allParas.Count - 1; i++)
            {
                var prev = allParas[i - 1];
                var curr = allParas[i];
                var next = allParas[i + 1];
                if (curr.IsEmpty && TableStyles.Contains(prev.SK) && TableStyles.Contains(next.SK))
                {
                    TaskPaneWinForms.AddMessage("PARATEXT", "WARNING",
                        "Paragraph break detected inside table — remove the empty paragraph between " +
                        "table rows (between \"" + Truncate(prev.Text, 30) + "\" and \"" + Truncate(next.Text, 30) + "\").",
                        "", curr.WordStart);
                    found++;
                }
            }
            return found;
        }

        // =====================================================================
        // FILL RUNS FROM OOXML
        // =====================================================================
        private static void FillRunsFromOoxml(string docPath, List<OoxmlPara> scopedParas)
        {
            try
            {
                XDocument bodyXml;
                using (var pkg = System.IO.Packaging.Package.Open(docPath, FileMode.Open, FileAccess.Read))
                {
                    var uri = new Uri("/word/document.xml", UriKind.Relative);
                    if (!pkg.PartExists(uri)) return;
                    using (var stream = pkg.GetPart(uri).GetStream())
                        bodyXml = XDocument.Load(stream);
                }

                var allParaElems = bodyXml.Descendants(W + "p").ToList();
                var scopedQueue = new Queue<OoxmlPara>(scopedParas);
                if (scopedQueue.Count == 0) return;
                OoxmlPara current = scopedQueue.Dequeue();

                foreach (var paraElem in allParaElems)
                {
                    string ooxmlText = string.Concat(
                        paraElem.Descendants(W + "t").Select(t => (string)t)).TrimEnd('\r', '\n');
                    if (!string.Equals(Normalize(ooxmlText), Normalize(current.Text),
                        StringComparison.OrdinalIgnoreCase)) continue;

                    // BUG FIX (navigation offset): this previously used
                    // paraElem.Elements(W + "r"), which only returns DIRECT child
                    // <w:r> elements of the paragraph. Any run nested one level
                    // deeper — most importantly every citation, which Word wraps as
                    // <w:hyperlink><w:r>...</w:r></w:hyperlink> — was silently
                    // skipped. Since RunFormattingRules computes each run's click
                    // position as a running offset over exactly the runs collected
                    // here, every skipped hyperlink run caused all subsequent
                    // formatting-rule positions in that paragraph to drift earlier
                    // than the real location, with the error compounding for each
                    // additional citation in the paragraph. Descendants() walks into
                    // <w:hyperlink> (and any other future wrapper, e.g. <w:ins>,
                    // <w:sdt>) while preserving document order, so the run list and
                    // its cumulative offsets now match the real document exactly.
                    foreach (var runElem in paraElem.Descendants(W + "r"))
                    {
                        string runText = string.Concat(runElem.Elements(W + "t").Select(t => (string)t));
                        if (string.IsNullOrEmpty(runText)) continue;
                        var rpr = runElem.Element(W + "rPr");
                        string fontName = null;
                        var rFonts = rpr?.Element(W + "rFonts");
                        if (rFonts != null)
                            fontName = (string)rFonts.Attribute(W + "ascii") ??
                                       (string)rFonts.Attribute(W + "hAnsi") ??
                                       (string)rFonts.Attribute(W + "cs");
                        current.Runs.Add(new OoxmlRun
                        {
                            Text = runText,
                            Bold = rpr?.Element(W + "b") != null,
                            Italic = rpr?.Element(W + "i") != null,
                            Superscript = GetValAttr(rpr, W + "vertAlign", "superscript"),
                            Subscript = GetValAttr(rpr, W + "vertAlign", "subscript"),
                            FontName = fontName
                        });
                    }
                    if (scopedQueue.Count == 0) break;
                    current = scopedQueue.Dequeue();
                }
            }
            catch { }
        }

        private static bool GetValAttr(XElement rpr, XName elem, string val)
        {
            if (rpr == null) return false;
            var e = rpr.Element(elem); if (e == null) return false;
            var attr = e.Attribute(W + "val");
            return attr != null && string.Equals((string)attr, val, StringComparison.OrdinalIgnoreCase);
        }

        private static string Normalize(string s) => Regex.Replace(s ?? "", @"\s+", " ").Trim();

        // =====================================================================
        // TEXT RULES
        // =====================================================================
        private int RunTextRules(List<OoxmlPara> paras)
        {
            int found = 0;
            foreach (var p in paras)
            {
                string text = p.Text;
                int start = p.WordStart;
                string sk = p.StyleKey;

                // Derived style-category booleans
                bool isAbbrevMonthStyle = AbbrevMonthStyles.Contains(sk);   // captions + table cells
                bool isSpellOutStyle = SpellOutMonthStyles.Contains(sk); // body text
                bool applyDateRules = DateRuleStyles.Contains(sk);       // all AMS prose + captions

                // ---- Range / numeric rules (all scoped styles) ----
                foreach (Match m in RxBetweenDash.Matches(text))
                { Add("WARNING", "Use \"between X and Y\" not a dash/hyphen: \"" + m.Value + "\".", text, m, start); found++; }
                foreach (Match m in RxBetweenTo.Matches(text))
                { Add("WARNING", "Use \"between X and Y\" not \"to\": \"" + m.Value + "\".", text, m, start); found++; }
                foreach (Match m in RxFromDash.Matches(text))
                { Add("WARNING", "Use \"from X to Y\" not a dash/hyphen: \"" + m.Value + "\".", text, m, start); found++; }
                foreach (Match m in RxUtcOn.Matches(text))
                { Add("WARNING", "\"UTC on\" — use a number instead of \"on\".", text, m, start); found++; }
                foreach (Match m in RxNumUnit.Matches(text))
                {
                    Add("WARNING", "Missing space between number and unit: \"" + m.Value + "\".", text, m, start); found++;
                }
                foreach (Match m in RxUnitUnit.Matches(text))
                { Add("WARNING", "Missing space between units: \"" + m.Value + "\" — e.g. \"ms\u207b\u00b9\" should be \"m s\u207b\u00b9\".", text, m, start); found++; }
                foreach (Match m in RxRepeatedUnit.Matches(text))
                {
                    // FALSE-POSITIVE FIX: the bare gap check ".{1,25}" matched across
                    // separate, independent parenthetical quantities such as
                    // "(0–1 mm day⁻¹), heavy (10–20 mm day⁻¹)" or
                    // "0.8 mm in (a) and 3 mm in (b)" — each unit repetition there is
                    // for a DIFFERENT value/panel, not a single series sharing one
                    // unit. A ")" between the two numbers indicates the two units
                    // belong to separate parenthetical phrases, so skip those.
                    if (m.Value.Contains(")")) continue;
                    string u = m.Groups[2].Value;
                    string snip = m.Value.Length > 50 ? m.Value.Substring(0, 50) + "…" : m.Value;
                    Add("WARNING", "Unit \"" + u + "\" repeated — use only after last number: \"" + snip + "\".", text, m, start); found++;
                }
                foreach (Match m in RxHttp.Matches(text))
                { Add("ERROR", "Use \"https://\" instead of \"http://\".", text, m, start); found++; }
                foreach (Match m in RxSciNotation.Matches(text))
                { Add("WARNING", "Scientific notation \"" + m.Value + "\" → use × 10ⁿ format.", text, m, start); found++; }
                foreach (Match m in RxLeadingZero.Matches(text))
                {
                    if (m.Index > 0 && char.IsDigit(text[m.Index - 1])) continue;
                    Add("WARNING", "Missing leading zero: \"" + m.Value + "\" → \"0" + m.Value + "\".", text, m, start); found++;
                }
                foreach (Match m in RxFourDigitComma.Matches(text))
                { Add("WARNING", "Four-digit number with comma: \"" + m.Value + "\" → \"" + m.Groups[1].Value + m.Groups[2].Value + "\".", text, m, start); found++; }
                foreach (Match m in RxSolidusAll.Matches(text))
                { Add("WARNING", "Use negative exponent not solidus: \"" + m.Value + "\" — e.g. \"W/m\" → \"W m\u207b\u00b9\".", text, m, start); found++; }
                foreach (Match m in RxPercentile.Matches(text))
                { Add("WARNING", "Percentile needs ordinal suffix: \"" + m.Value + "\" → \"" + m.Groups[1].Value + "th percentile\".", text, m, start); found++; }
                foreach (Match m in RxSpaceBeforePercent.Matches(text))
                {
                    // New rule: no space between number and "%" — e.g. "36.0 %" → "36.0%"
                    Add("WARNING",
                        "No space allowed between number and \"%\": \"" + m.Value + "\" → \"" + m.Groups[1].Value + "%\".",
                        text, m, start);
                    found++;
                }
                foreach (Match m in RxMissingSpaceAfterComma.Matches(text))
                {
                    // New rule: missing space after a list-separator comma between
                    // numbers — e.g. "0.2,0.5, 0.8" → "0.2, 0.5, 0.8".
                    Add("WARNING",
                        "Missing space after comma in number list: \"" + m.Value + "\" → \"" + m.Groups[1].Value + ", \".",
                        text, m, start);
                    found++;
                }
                foreach (Match m in RxMyrKyr.Matches(text))
                { Add("WARNING", "Wrong geological time unit: \"" + m.Value + "\" — use Myr, kyr, Ma, or ka.", text, m, start); found++; }

                // ---- Month rules --------------------------------------------
                // BUG FIX: month abbreviation rules now correctly scoped per AMS style guide:
                //
                //   AbbrevMonthStyles  (figurecaption, tablecaption, tablebody, tablehead):
                //     • Full month names MUST be abbreviated → flag full names, suggest abbrev.
                //
                //   SpellOutMonthStyles (paratext and other body styles):
                //     • Month abbreviations MUST be spelled out → flag abbreviations, suggest full name.
                //
                //   "May" is never flagged for abbreviation/expansion because it is the same in
                //   both forms.  It is only matched by RxAnyMonth when followed by a digit
                //   (the month-as-date context), and only to support the date-format rules below.
                //
                foreach (Match m in RxAnyMonth.Matches(text))
                {
                    // Resolve which capture group fired
                    string mv = (!string.IsNullOrEmpty(m.Groups[1].Value)
                                    ? m.Groups[1].Value
                                    : m.Groups[2].Value).Trim();
                    if (string.IsNullOrEmpty(mv)) continue;
                    if (!MonthCanonical.TryGetValue(mv, out string canon)) continue;

                    bool isMay = string.Equals(canon, "May", StringComparison.OrdinalIgnoreCase);
                    if (isMay) continue; // May never needs expanding or abbreviating

                    // Is the token currently a full name or an abbreviation?
                    bool tokenIsFullName = MonthAbbrev.ContainsKey(canon) &&
                                          string.Equals(mv, canon, StringComparison.OrdinalIgnoreCase);
                    bool tokenIsAbbrev = !tokenIsFullName; // abbreviated (Jan, Feb, …)

                    string fullName = canon;
                    string abbr = MonthAbbrev.ContainsKey(canon) ? MonthAbbrev[canon] : canon;

                    if (isAbbrevMonthStyle && tokenIsFullName)
                    {
                        // BUG FIX: was previously checking tokenIsAbbrev here (wrong)
                        Add("WARNING",
                            "Month \"" + mv + "\" in caption/table — use AMS abbreviation \"" + abbr + "\".",
                            text, m, start);
                        found++;
                    }
                    else if (isSpellOutStyle && tokenIsAbbrev)
                    {
                        // BUG FIX: was previously checking tokenIsFullName here (wrong)
                        Add("WARNING",
                            "Month abbreviation \"" + mv + "\" in body text — spell out in full: \"" + fullName + "\".",
                            text, m, start);
                        found++;
                    }
                }

                // ---- Date format rules (scoped to DateRuleStyles only) ------
                // BUG FIX: original code applied date rules to ALL styles; they are now
                // restricted to the styles listed in DateRuleStyles (captions, table cells,
                // and AMS body text).  This prevents spurious warnings in styles such as
                // acknowledgementtext or formalarg where date conventions may differ.
                if (applyDateRules)
                {
                    // Wrong order: "January 21, 2017"
                    foreach (Match m in RxWrongDate.Matches(text))
                    {
                        Add("WARNING",
                            "Wrong date format: \"" + m.Value + "\" — AMS uses day month year (e.g. \"21 January 2017\").",
                            text, m, start);
                        found++;
                    }

                    // Wrong order / unwanted comma: "21, January 2017" or "21 Jan 2017"
                    foreach (Match m in RxWrongDateDay.Matches(text))
                    {
                        string month = m.Groups[3].Value;
                        bool isAbbrev = MonthAbbrevSet.Contains(month);
                        bool hasComma = m.Groups[2].Value == "," || m.Groups[4].Value == ",";
                        // BUG FIX (false positive): an abbreviated month (e.g. "Feb") is
                        // only WRONG in styles that require full month names
                        // (isSpellOutStyle). In AbbrevMonthStyles — figure/table
                        // captions and table cells — AMS style REQUIRES the
                        // abbreviation, so "7 Feb 2017" there is already correct and
                        // must not be flagged. Previously this check ignored style
                        // context entirely and flagged every abbreviated month in any
                        // DateRuleStyles paragraph, which incorrectly fired on
                        // correctly-formatted caption/table dates.
                        bool abbrevIsWrongHere = isAbbrev && isSpellOutStyle;
                        if (!abbrevIsWrongHere && !hasComma) continue;
                        string reason = abbrevIsWrongHere
                            ? "abbreviated month \"" + month + "\" — spell out in full"
                            : "unwanted comma in date";
                        Add("WARNING",
                            "Wrong date format: \"" + m.Value + "\" — " + reason +
                            ". AMS format is day month year with no commas (e.g. \"21 January 2017\").",
                            text, m, start);
                        found++;
                    }

                    // Redundant date range: "14 Sep 2024 to 24 Sep 2024" →
                    // consolidate to "14 to 24 Sep 2024" (month/year stated once).
                    foreach (Match m in RxRedundantDateRange.Matches(text))
                    {
                        string day1 = m.Groups[1].Value;
                        string month = m.Groups[2].Value;
                        string year = m.Groups[3].Value;
                        string day2 = m.Groups[4].Value;
                        Add("WARNING",
                            "Redundant date range: \"" + m.Value + "\" — month/year repeated unnecessarily. " +
                            "AMS consolidates this to \"" + day1 + " to " + day2 + " " + month + " " + year + "\".",
                            text, m, start);
                        found++;
                    }

                    // Ordinal suffix on day: "21st January" or "January 21st"
                    foreach (Match m in RxOrdinalDate.Matches(text))
                    {
                        string suffix = !string.IsNullOrEmpty(m.Groups[2].Value)
                                            ? m.Groups[2].Value
                                            : m.Groups[4].Value;
                        Add("WARNING",
                            "Ordinal suffix \"" + suffix + "\" on day number in date: \"" + m.Value +
                            "\" — delete the ordinal suffix and use a plain number (e.g. \"21 January 2017\").",
                            text, m, start);
                        found++;
                    }
                }

                // ---- Remaining rules (all scoped styles) ----
                foreach (Match m in RxSymAsym.Matches(text))
                {
                    string sym = m.Groups[3].Value;
                    string leftNum = m.Groups[1].Value;
                    Add("WARNING",
                        "Symbol asymmetry: \"" + m.Value + "\" — \"" + sym +
                        "\" must appear after both numbers (e.g. \"" + leftNum + sym + " ± ...\").",
                        text, m, start);
                    found++;
                }
                foreach (Match m in RxCycles.Matches(text))
                { Add("WARNING", "\"cycles per ...\" — use Hz.", text, m, start); found++; }
                foreach (Match m in RxMicron.Matches(text))
                { Add("WARNING", "\"" + m.Value + "\" is obsolete — use \"micrometer(s)\" or \"\u00b5m\".", text, m, start); found++; }
                foreach (Match m in RxMuUnit.Matches(text))
                { Add("WARNING", "\"mu\" as unit — use \"\u00b5\" (e.g. \"\u00b5m\" not \"mu m\").", text, m, start); found++; }
                foreach (Match m in RxArcSec.Matches(text))
                { Add("WARNING", "\"" + m.Value + "\" → \"" + m.Groups[1].Value + " arc s\".", text, m, start); found++; }
                foreach (Match m in RxTempBothUnits.Matches(text))
                {
                    string u1 = m.Groups[2].Value, u2 = m.Groups[4].Value;
                    if (u1.Length > 1 && u2.Length > 1 &&
                        string.Equals(u1, u2, StringComparison.OrdinalIgnoreCase))
                    {
                        Add("WARNING",
                            "Temperature series: \"" + m.Value + "\" → \"" +
                            m.Groups[1].Value + "° and " + m.Groups[3].Value + u2 + "\".",
                            text, m, start);
                        found++;
                    }
                }
                foreach (Match m in RxPlusMinusAsym.Matches(text))
                {
                    string u = m.Groups[3].Value;
                    int fe = m.Index + m.Groups[1].Length;
                    string gap = text.Substring(fe, Math.Min(u.Length + 2, text.Length - fe));
                    if (!gap.TrimStart().StartsWith(u))
                    { Add("WARNING", "Both sides of ± must carry the unit: \"" + m.Value + "\".", text, m, start); found++; }
                }
                foreach (Match m in RxSeriesSymbol.Matches(text))
                {
                    string u = m.Groups[2].Value;
                    int pos = m.Index + m.Groups[1].Length;
                    string follow = text.Substring(pos, Math.Min(3, text.Length - pos)).TrimStart();
                    if (!follow.StartsWith(u))
                    { Add("WARNING", "\"" + u + "\" must appear after EACH number in series.", text, m, start); found++; }
                }
                foreach (Match m in RxUnitTypos.Matches(text))
                {
                    string wrong = m.Value;
                    string correct = UnitTypos.ContainsKey(wrong) ? UnitTypos[wrong] : "?";
                    Add("WARNING", "Unit typo: \"" + wrong + "\" should be \"" + correct + "\".", text, m, start); found++;
                }

                // ---- Figure-caption-only rules ----
                if (sk == "figurecaption")
                {
                    foreach (Match m in RxPanel.Matches(text))
                    { Add("WARNING", "\"" + m.Value + "\" in figure caption — use \"(a)\" not \"Panel (a)\".", text, m, start); found++; }
                    foreach (Match m in RxCaptionDirections.Matches(text))
                    {
                        string inner = m.Groups[1].Value;
                        if (!string.Equals(inner, inner.ToLowerInvariant(), StringComparison.Ordinal))
                        { Add("WARNING", "Use lowercase in caption: \"(" + inner + ")\" → \"(" + inner.ToLowerInvariant() + ")\".", text, m, start); found++; }
                    }
                }

                // ---- Whitespace character scan ----
                for (int ci = 0; ci < text.Length; ci++)
                {
                    char ch = text[ci];
                    if (ch == '\u2009' || ch == '\u202F' || ch == '\u200A' || ch == '\u00A0')
                    {
                        bool pd = ci > 0 && char.IsDigit(text[ci - 1]);
                        bool na = ci < text.Length - 1 && char.IsLetter(text[ci + 1]);
                        if (pd && na) continue; // correct thin-space between number and unit
                        // SEVERITY CHANGE: thin/non-breaking-space artifacts are reported
                        // as ERROR (not WARNING) — these are invisible characters that
                        // silently corrupt copyedited text and must always be resolved
                        // before submission, so they are treated as a hard error rather
                        // than an advisory warning.
                        TaskPaneWinForms.AddMessage("PARATEXT", "ERROR",
                            "Thin/non-breaking space (U+" + ((int)ch).ToString("X4") + ") — remove or replace.",
                            Excerpt(text, ci, 1), start + ci);
                        found++;
                    }
                }
            }
            return found;
        }

        // =====================================================================
        // FORMATTING RULES
        // =====================================================================
        private int RunFormattingRules(List<OoxmlPara> paras)
        {
            int found = 0;
            foreach (var p in paras)
            {
                if (p.Runs.Count == 0) continue;
                int offset = 0;

                for (int ri0 = 0; ri0 < p.Runs.Count; ri0++)
                {
                    var run = p.Runs[ri0];
                    string t = run.Text ?? "";
                    string tw = t.Trim();
                    bool italic = run.Italic, bold = run.Bold;
                    bool sup = run.Superscript, sub = run.Subscript;

                    // FIX: character position = paragraph start + cumulative run offset
                    // This ensures clicking the error navigates to the exact run in the document.
                    int pos = p.WordStart + offset;

                    // Rule 04 — dBZ: the Z must be italic
                    if (tw == "dBZ" && !italic)
                    { TaskPaneWinForms.AddMessage("PARATEXT", "WARNING", "In \"dBZ\" the Z must be italic.", tw, pos); found++; }

                    // Rule R — known unit must not be italic
                    // FALSE-POSITIVE FIX: now checks MultiCharKnownUnits instead of
                    // KnownUnits, so single-letter italic variable names (K, g, N, h,
                    // L, W, ...) that collide with single-letter unit symbols are no
                    // longer misflagged. Multi-character unit symbols (km, hPa, kPa,
                    // MHz, dBZ, ...) are unambiguous and still checked.
                    if (italic && !sup && !sub && MultiCharKnownUnits.Contains(tw))
                    { TaskPaneWinForms.AddMessage("PARATEXT", "WARNING", "Unit \"" + tw + "\" should not be italic — units are always roman.", tw, pos); found++; }

                    // Rule V — standalone "v" must be changed to upsilon υ (U+03C5)
                    if (tw == "v")
                    {
                        TaskPaneWinForms.AddMessage("PARATEXT", "WARNING",
                            "Standalone \"v\" detected — change to upsilon \"\u03c5\" (U+03C5). " +
                            "In AMS text \"v\" as a standalone variable should be the Greek letter υ.",
                            tw, pos);
                        found++;
                    }

                    // Rule VAR — standalone single letter not italic = variable needs italic
                    if (tw.Length == 1 && char.IsLetter(tw[0]) &&
                        !SkipStandaloneVars.Contains(tw) && tw != "v" && !italic)
                    {
                        // FALSE-POSITIVE FIX 1: compass-direction letters N, S, E, W
                        // immediately preceded by a degree-sign run (e.g. "30°S",
                        // "5°N", "160°E") are roman directional letters, not italic
                        // scalar variables — these are extremely common in
                        // latitude/longitude coordinates and were previously
                        // misflagged on every occurrence.
                        bool isCompassDirection =
                            (tw == "N" || tw == "S" || tw == "E" || tw == "W") &&
                            ri0 > 0 &&
                            (p.Runs[ri0 - 1].Text ?? "").TrimEnd().EndsWith("°");

                        // FALSE-POSITIVE FIX 2: a single letter that is really just
                        // the start of an ordinary word split across runs (e.g. the
                        // run "w" followed immediately by the run "here " forming
                        // "where") is not a standalone variable. A genuine standalone
                        // variable is followed by a space, punctuation, or a
                        // subscript/superscript digit — not directly by another
                        // lowercase letter with no separator.
                        bool isWordFragment =
                            ri0 < p.Runs.Count - 1 &&
                            !string.IsNullOrEmpty(p.Runs[ri0 + 1].Text) &&
                            char.IsLower(p.Runs[ri0 + 1].Text[0]);

                        // FALSE-POSITIVE FIX 3: figure/table panel-letter ranges such
                        // as "Figs. 4a–c" or "Fig. 5d–j" are frequently split into
                        // separate runs around the en dash (the dash often carries
                        // its own formatting, e.g. from a copyedit/QC pass), which
                        // isolates the trailing panel letter ("c", "j", ...) into its
                        // own single-character run. That letter is correctly roman —
                        // panel labels are never italic — and must not be mistaken
                        // for an un-italicized variable. A genuine standalone
                        // variable is never immediately preceded by a bare dash with
                        // no separating space; a lettered range/list always is.
                        bool isRangeLetter =
                            ri0 > 0 &&
                            Regex.IsMatch((p.Runs[ri0 - 1].Text ?? "").Trim(), @"^[-\u2013\u2014]$");

                        // FALSE-POSITIVE FIX 4: single-letter unit symbols (m, h, s,
                        // g, K, L, N, W, J, ...) are frequently isolated into their
                        // own run immediately before a superscript negative exponent
                        // run — e.g. "kg m" + "−"(superscript) + "2"(superscript) +
                        // " h" + "−"(superscript) + "1"(superscript) for the unit
                        // "kg m⁻² h⁻¹". These unit letters are correctly roman (Rule R
                        // already enforces this for multi-character units) and must
                        // not be flagged here as needing italics. This is restricted
                        // to the case where the very next run is ITSELF a superscript
                        // that looks like a bare exponent (a digit or minus sign) —
                        // the unit-immediately-before-its-own-exponent pattern — so it
                        // does not suppress a genuine non-italic variable that
                        // happens to carry a subscripted/superscripted label (e.g. an
                        // italic-missing "h" followed by a subscript "B").
                        bool isUnitExponent =
                            KnownUnits.Contains(tw) &&
                            ri0 < p.Runs.Count - 1 &&
                            p.Runs[ri0 + 1].Superscript &&
                            Regex.IsMatch((p.Runs[ri0 + 1].Text ?? "").TrimStart(), @"^[-\u2212\u207b\d]");

                        if (!isCompassDirection && !isWordFragment && !isRangeLetter && !isUnitExponent)
                        {
                            TaskPaneWinForms.AddMessage("PARATEXT", "WARNING",
                                "Standalone variable \"" + tw + "\" is not italic — variables must be italic.",
                                tw, pos);
                            found++;
                        }
                    }

                    // Rule Y — Greek scalar must not be bold
                    if (bold && (tw == "\u03c3" || tw == "\u03c0" || tw == "\u03bc" ||
                                 tw == "\u03b1" || tw == "\u03b2" || tw == "\u03b3" ||
                                 tw == "\u03bb" || tw == "\u03c9"))
                    { TaskPaneWinForms.AddMessage("PARATEXT", "WARNING", "Greek symbol \"" + tw + "\" is bold — scalar variables should be italic.", tw, pos); found++; }

                    // Rule MU2 — U+00B5 (MICRO SIGN) must be replaced with U+03BC (GREEK SMALL MU)
                    if (tw.Contains('\u00B5'))
                    {
                        TaskPaneWinForms.AddMessage("PARATEXT", "WARNING",
                            "Character \"\u00b5\" (U+00B5 MICRO SIGN) found — replace with " +
                            "\"\u03bc\" (U+03BC GREEK SMALL LETTER MU). " +
                            "These look identical but U+00B5 is the wrong code point for AMS text.",
                            Excerpt(t, 0, Math.Min(t.Length, 30)), pos);
                        found++;
                    }

                    // Rule T — font must be Times New Roman
                    if (!string.IsNullOrEmpty(run.FontName) &&
                        !string.Equals(run.FontName, "Times New Roman", StringComparison.OrdinalIgnoreCase))
                    {
                        TaskPaneWinForms.AddMessage("PARATEXT", "WARNING",
                            GetFontMessage(run.FontName),
                            Excerpt(t, 0, Math.Min(t.Length, 30)), pos);
                        found++;
                    }

                    offset += t.Length; // advance AFTER using pos so pos points to run start
                }

                // Rule SS — subscript immediately followed by superscript (or vice versa)
                for (int ri = 0; ri < p.Runs.Count - 1; ri++)
                {
                    var curr = p.Runs[ri];
                    var next = p.Runs[ri + 1];
                    bool currSub = curr.Subscript && !string.IsNullOrWhiteSpace(curr.Text);
                    bool currSup = curr.Superscript && !string.IsNullOrWhiteSpace(curr.Text);
                    bool nextSub = next.Subscript && !string.IsNullOrWhiteSpace(next.Text);
                    bool nextSup = next.Superscript && !string.IsNullOrWhiteSpace(next.Text);

                    if ((currSub && nextSup) || (currSup && nextSub))
                    {
                        string kind = (currSub && nextSup)
                            ? "Subscript \"" + curr.Text.Trim() + "\" immediately followed by superscript \"" + next.Text.Trim() + "\""
                            : "Superscript \"" + curr.Text.Trim() + "\" immediately followed by subscript \"" + next.Text.Trim() + "\"";

                        // FIX: compute run offset precisely so clicking navigates correctly
                        int runOffset = 0;
                        for (int k = 0; k < ri; k++) runOffset += p.Runs[k].Text?.Length ?? 0;

                        TaskPaneWinForms.AddMessage("PARATEXT", "WARNING",
                            kind + " — this combination should be set in MathType.",
                            Truncate(curr.Text.Trim() + next.Text.Trim(), 30),
                            p.WordStart + runOffset);
                        found++;
                    }
                }
            }
            return found;
        }

        // =====================================================================
        // DOCUMENT-LEVEL RULES
        // =====================================================================
        private int RunDocumentRules(List<OoxmlPara> paras, List<(string Text, int Start)> allTexts)
        {
            int found = 0;
            foreach (string unit in UnitAbbrevMap.Keys)
            {
                string fullName = UnitAbbrevMap[unit];
                Regex rxAbbr = UnitPatterns[unit], rxFull = UnitFullNamePatterns[unit];
                int abbrStart = 0, fullStart = 0;
                string abbrEx = "", fullEx = "";
                bool abbrFound = false, fullFound = false;

                foreach (var (at, astart) in allTexts)
                {
                    if (!abbrFound)
                    {
                        Match am = rxAbbr.Match(at);
                        if (am.Success) { abbrFound = true; abbrStart = astart + am.Index; abbrEx = Excerpt(at, am.Index, am.Length); }
                    }
                    if (!fullFound)
                    {
                        Match fm = rxFull.Match(at);
                        if (fm.Success) { fullFound = true; fullStart = astart + fm.Index; fullEx = Excerpt(at, fm.Index, fm.Length); }
                    }
                    if (abbrFound && fullFound) break;
                }

                if (abbrFound && !fullFound)
                { TaskPaneWinForms.AddMessage("PARATEXT", "WARNING", "\"" + unit + "\" used but \"" + fullName + "\" never defined — add \"" + fullName + " (" + unit + ")\" at first use.", abbrEx, abbrStart); found++; }
                if (fullFound && !abbrFound)
                { TaskPaneWinForms.AddMessage("PARATEXT", "WARNING", "\"" + fullName + "\" appears but \"" + unit + "\" never introduced — add \"(" + unit + ")\" after first use.", fullEx, fullStart); found++; }
            }
            return found;
        }

        // =====================================================================
        // HELPERS
        // =====================================================================
        private void Add(string sev, string msg, string text, Match m, int paraStart)
        {
            // FIX: pass paraStart + m.Index so the character offset is exact.
            // The task pane's click handler must call:
            //   doc.Application.ActiveDocument.Range(charPos, charPos).Select();
            // with this value to navigate reliably.
            TaskPaneWinForms.AddMessage("PARATEXT", sev, msg,
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