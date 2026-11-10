using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

public static class AlifboCore
{
    const string APOS = "'\u2018\u2019\u201B\u02B9\u02BB\u02BC\u02C8\u02CA\u2032`\u00B4\uFF07";
    static bool IsL(char c) { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= 0xC0 && c <= 0x24F); }
    static bool IsA(char c) { return c != '\0' && APOS.IndexOf(c) >= 0; }
    static bool IsW(char c) { return IsL(c) || IsA(c); }
    static bool IsSp(char c) { return c == ' ' || c == '\u00A0'; }
    static bool IsOpenCtx(char p) { return p == '\0' || " \t\r\n\u000B([{-/\u00A0\u2013\u2014\u201C".IndexOf(p) >= 0; }

    static void WordFrag(string P, int st, int len, bool toNew, string[] fr)
    {
        int end = st + len, i = st;
        if (toNew)
        {
            while (i < end)
            {
                char c = P[i], c1 = (i + 1 < end) ? P[i + 1] : '\0', c2 = (i + 2 < end) ? P[i + 2] : '\0';
                if ("oOgG".IndexOf(c) >= 0 && c1 != '\0' && IsA(c1) && (c2 == '\0' || IsL(c2)))
                {
                    fr[i] = c == 'o' ? "\u00F6" : c == 'O' ? "\u00D6" : c == 'g' ? "\u011F" : "\u011E";
                    fr[i + 1] = ""; i += 2;
                }
                else if (IsA(c) && i > st && i < end - 1 && IsL(P[i - 1]) && IsL(c1))
                {
                    fr[i] = "\u02BC"; i++;
                }
                else if ((c == 's' || c == 'S') && (c1 == 'h' || c1 == 'H'))
                {
                    fr[i] = c == 'S' ? "\u015E" : "\u015F"; fr[i + 1] = ""; i += 2;
                }
                else if ((c == 'c' || c == 'C') && (c1 == 'h' || c1 == 'H'))
                {
                    fr[i] = c == 'C' ? "\u00C7" : "\u00E7"; fr[i + 1] = ""; i += 2;
                }
                else i++;
            }
        }
        else
        {
            string t = P.Substring(st, len);
            bool up = len > 1 && t == t.ToUpperInvariant() && t != t.ToLowerInvariant();
            for (i = st; i < end; i++)
            {
                switch (P[i])
                {
                    case '\u00F6': fr[i] = "o\u02BB"; break;
                    case '\u00D6': fr[i] = "O\u02BB"; break;
                    case '\u011F': fr[i] = "g\u02BB"; break;
                    case '\u011E': fr[i] = "G\u02BB"; break;
                    case '\u015F': fr[i] = "sh"; break;
                    case '\u015E': fr[i] = up ? "SH" : "Sh"; break;
                    case '\u00E7': fr[i] = "ch"; break;
                    case '\u00C7': fr[i] = up ? "CH" : "Ch"; break;
                }
            }
        }
    }

    static void QuoteFrag(string P, string[] fr, bool spaces)
    {
        int n = P.Length;
        char last = '\0';
        bool lowOpen = false, skipSp = false, runLS = true;
        for (int i = 0; i < n; i++)
        {
            char c = P[i], c1 = (i + 1 < n) ? P[i + 1] : '\0';
            if (c == '\u200B' || c == '\u200C' || c == '\u200D' || c == '\uFEFF' || c == '\u00AD') { fr[i] = ""; continue; }
            if (skipSp && IsSp(c)) { fr[i] = ""; continue; }
            skipSp = false;
            int act = 0, extra = 0;
            if (c == '<' && c1 == '<') { act = 1; extra = 1; }
            else if (c == '>' && c1 == '>') { act = 2; extra = 1; }
            else if (c == '\u00AB' || c == '\u201F') act = 1;
            else if (c == '\u201E') { act = 1; lowOpen = true; }
            else if (c == '\u00BB') act = 2;
            else if (c == '\u201C' && lowOpen) { act = 2; lowOpen = false; }
            else if (c == '\u2018' && c1 == '\u2018') { act = 1; extra = 1; }
            else if (c == '\u2019' && c1 == '\u2019') { act = 2; extra = 1; }
            else if (c == '"') act = IsOpenCtx(last) ? 1 : 2;
            if (act == 1)
            {
                fr[i] = "\u201C"; if (extra == 1) fr[i + 1] = "";
                last = '\u201C'; skipSp = true; i += extra; continue;
            }
            if (act == 2)
            {
                for (int j = i - 1; j >= 0 && (fr[j] == " " || fr[j] == "\u00A0" || fr[j] == ""); j--) fr[j] = "";
                fr[i] = "\u201D"; if (extra == 1) fr[i + 1] = "";
                last = '\u201D'; i += extra; continue;
            }
            if (spaces && c == ' ')
            {
                if (last == ' ') { if (!runLS) { fr[i] = ""; continue; } }
                else runLS = (last == '\0' || last == '\n' || last == '\r');
            }
            if (fr[i].Length > 0) last = fr[i][fr[i].Length - 1];
        }
    }

    public static string[] Frag(string P, bool toNew, bool spaces, HashSet<string> exc)
    {
        int n = P.Length;
        string[] fr = new string[n];
        for (int i = 0; i < n; i++) fr[i] = P[i].ToString();
        int k = 0;
        while (k < n)
        {
            if (IsW(P[k]))
            {
                int st = k;
                while (k < n && IsW(P[k])) k++;
                int len = k - st;
                if (len >= 2 && !(exc != null && exc.Contains(P.Substring(st, len).ToLowerInvariant())))
                    WordFrag(P, st, len, toNew, fr);
            }
            else k++;
        }
        QuoteFrag(P, fr, spaces);
        return fr;
    }

    public static string ConvertString(string s, bool toNew, bool spaces, HashSet<string> exc)
    {
        string[] fr = Frag(s, toNew, spaces, exc);
        StringBuilder sb = new StringBuilder(s.Length);
        for (int i = 0; i < fr.Length; i++) sb.Append(fr[i]);
        return sb.ToString();
    }

    public static HashSet<string> LoadExc(string path)
    {
        HashSet<string> h = new HashSet<string>();
        try
        {
            if (File.Exists(path))
                foreach (string ln in File.ReadAllLines(path, Encoding.Default))
                {
                    string t = ln.Trim().ToLowerInvariant();
                    if (t.Length > 0 && t[0] != '#') h.Add(t);
                }
        }
        catch { }
        return h;
    }

    static bool HasAny(string gap, string[] marks)
    {
        for (int i = 0; i < marks.Length; i++) if (gap.IndexOf(marks[i], StringComparison.Ordinal) >= 0) return true;
        return false;
    }

    static string Enc(string s) { return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;"); }

    static string ConvertXml(string xml, string tag, string[] bnd, string[] sep, bool toNew, HashSet<string> exc, bool spaces, out int changed)
    {
        Regex rx = new Regex("<" + tag + @"(?:\s[^>]*)?>([^<]*)</" + tag + ">", RegexOptions.Singleline);
        MatchCollection ms = rx.Matches(xml);
        StringBuilder sb = new StringBuilder(xml.Length + 64);
        int pos = 0, gi = 0;
        changed = 0;
        while (gi < ms.Count)
        {
            int gj = gi;
            while (gj + 1 < ms.Count)
            {
                int e0 = ms[gj].Index + ms[gj].Length;
                if (HasAny(xml.Substring(e0, ms[gj + 1].Index - e0), bnd)) break;
                gj++;
            }
            StringBuilder P = new StringBuilder();
            int cnt = gj - gi + 1;
            int[] starts = new int[cnt], lens = new int[cnt];
            string[] texts = new string[cnt];
            for (int k = gi; k <= gj; k++)
            {
                if (k > gi)
                {
                    int e0 = ms[k - 1].Index + ms[k - 1].Length;
                    if (HasAny(xml.Substring(e0, ms[k].Index - e0), sep)) P.Append('\t');
                }
                string txt = WebUtility.HtmlDecode(ms[k].Groups[1].Value);
                texts[k - gi] = txt;
                starts[k - gi] = P.Length;
                lens[k - gi] = txt.Length;
                P.Append(txt);
            }
            string[] fr = Frag(P.ToString(), toNew, spaces, exc);
            for (int k = gi; k <= gj; k++)
            {
                Match m = ms[k];
                sb.Append(xml, pos, m.Index - pos);
                StringBuilder nt = new StringBuilder();
                for (int q = starts[k - gi]; q < starts[k - gi] + lens[k - gi]; q++) nt.Append(fr[q]);
                string nts = nt.ToString();
                if (nts == texts[k - gi]) sb.Append(m.Value);
                else
                {
                    string open = m.Value.Substring(0, m.Value.IndexOf('>') + 1);
                    sb.Append(open).Append(Enc(nts)).Append("</").Append(tag).Append(">");
                    changed++;
                }
                pos = m.Index + m.Length;
            }
            gi = gj + 1;
        }
        sb.Append(xml, pos, xml.Length - pos);
        return sb.ToString();
    }

    static int ProcessEntry(ZipArchive zip, string name, string tag, string[] bnd, string[] sep, bool toNew, HashSet<string> exc)
    {
        ZipArchiveEntry e = zip.GetEntry(name);
        if (e == null) return 0;
        string xml;
        using (StreamReader sr = new StreamReader(e.Open(), new UTF8Encoding(false))) xml = sr.ReadToEnd();
        int changed;
        string outXml = ConvertXml(xml, tag, bnd, sep, toNew, exc, false, out changed);
        if (changed > 0)
        {
            byte[] bytes = new UTF8Encoding(false).GetBytes(outXml);
            e.Delete();
            ZipArchiveEntry newEntry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream st = newEntry.Open())
            {
                st.Write(bytes, 0, bytes.Length);
            }
        }
        return changed;
    }

    static readonly string[] DocxB = { "</w:p>", "<w:p>", "<w:p " };
    static readonly string[] DocxS = { "<w:tab", "<w:br", "<w:cr", "<w:noBreakHyphen" };
    static readonly string[] PptxB = { "</a:p>", "<a:p>", "<a:p " };
    static readonly string[] PptxS = { "<a:br", "<a:fld" };
    static readonly string[] XlsxB = { "</si>", "<si>", "<si ", "</is>", "<is>", "</c>", "<c ", "<c>", "</text>", "<text>" };
    static readonly string[] NoSep = { };

    // natija: o'zgargan matn bo'laklari soni; -1 = qo'llab-quvvatlanmaydigan format
    public static int ConvertFile(string src, string dst, bool toNew, HashSet<string> exc)
    {
        string ext = Path.GetExtension(src).ToLowerInvariant();
        if (ext == ".txt" || ext == ".md" || ext == ".csv" || ext == ".srt")
        {
            byte[] raw = File.ReadAllBytes(src);
            bool bom = raw.Length >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF;
            string text;
            try { text = new UTF8Encoding(false, true).GetString(raw, bom ? 3 : 0, raw.Length - (bom ? 3 : 0)); }
            catch { text = Encoding.Default.GetString(raw); bom = false; }
            string res = ConvertString(text, toNew, false, exc);
            File.WriteAllBytes(dst, new UTF8Encoding(bom).GetPreamble().Length > 0 ? Cat(new UTF8Encoding(true).GetPreamble(), Encoding.UTF8.GetBytes(res)) : Encoding.UTF8.GetBytes(res));
            return res == text ? 0 : 1;
        }
        if (ext != ".docx" && ext != ".xlsx" && ext != ".pptx") return -1;
        if (!string.Equals(Path.GetFullPath(src), Path.GetFullPath(dst), StringComparison.OrdinalIgnoreCase)) File.Copy(src, dst, true);
        int total = 0;
        using (FileStream fs = new FileStream(dst, FileMode.Open, FileAccess.ReadWrite))
        using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Update))
        {
            List<string> names = new List<string>();
            foreach (ZipArchiveEntry e in zip.Entries) names.Add(e.FullName);
            foreach (string n in names)
            {
                if (ext == ".docx" && Regex.IsMatch(n, @"^word/(document|footnotes|endnotes|comments|header\d*|footer\d*)\.xml$"))
                    total += ProcessEntry(zip, n, "w:t", DocxB, DocxS, toNew, exc);
                else if (ext == ".xlsx" && Regex.IsMatch(n, @"^xl/(sharedStrings|comments\d*|worksheets/sheet\d+)\.xml$"))
                    total += ProcessEntry(zip, n, "t", XlsxB, NoSep, toNew, exc);
                else if (ext == ".pptx" && Regex.IsMatch(n, @"^ppt/(slides/slide|notesSlides/notesSlide)\d+\.xml$"))
                    total += ProcessEntry(zip, n, "a:t", PptxB, PptxS, toNew, exc);
            }
        }
        return total;
    }

    static byte[] Cat(byte[] a, byte[] b)
    {
        byte[] r = new byte[a.Length + b.Length];
        Buffer.BlockCopy(a, 0, r, 0, a.Length);
        Buffer.BlockCopy(b, 0, r, a.Length, b.Length);
        return r;
    }
}
