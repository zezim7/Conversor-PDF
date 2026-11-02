using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

// Conversão básica de .docx/.xlsx/.pptx para HTML, usada quando não há
// Microsoft Office nem LibreOffice no computador. O Edge depois imprime em PDF.
namespace ConversorPDF
{
    class Rel { public string Tipo, Alvo; public bool Externo; }

    static class Ooxml
    {
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static readonly XNamespace MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        public static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public static ZipArchiveEntry Entrada(ZipArchive z, string c)
        {
            if (c == null) return null;
            c = c.TrimStart('/');
            return z.GetEntry(c) ?? z.Entries.FirstOrDefault(x => string.Equals(x.FullName, c, StringComparison.OrdinalIgnoreCase));
        }

        public static XDocument Xml(ZipArchive z, string c)
        {
            var e = Entrada(z, c);
            if (e == null) return null;
            using (var s = e.Open()) return XDocument.Load(s);
        }

        public static byte[] Bytes(ZipArchive z, string c)
        {
            var e = Entrada(z, c);
            if (e == null) return null;
            using (var s = e.Open())
            using (var ms = new MemoryStream())
            {
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }

        public static Dictionary<string, Rel> Rels(ZipArchive z, string parte)
        {
            int i = parte.LastIndexOf('/');
            string pasta = i >= 0 ? parte.Substring(0, i) : "";
            string arq = parte.Substring(i + 1);
            string caminho = (pasta.Length > 0 ? pasta + "/" : "") + "_rels/" + arq + ".rels";
            var d = new Dictionary<string, Rel>();
            var x = Xml(z, caminho);
            if (x == null) return d;
            foreach (var r in x.Root.Elements())
            {
                string id = (string)r.Attribute("Id"), alvo = (string)r.Attribute("Target");
                if (id == null || alvo == null) continue;
                bool ext = (string)r.Attribute("TargetMode") == "External";
                d[id] = new Rel { Tipo = (string)r.Attribute("Type") ?? "", Externo = ext, Alvo = ext ? alvo : Combinar(pasta, alvo) };
            }
            return d;
        }

        public static string Relacionado(Dictionary<string, Rel> rels, string tipo)
        {
            foreach (var r in rels.Values) if (!r.Externo && r.Tipo.EndsWith("/" + tipo)) return r.Alvo;
            return null;
        }

        public static string DocumentoPrincipal(ZipArchive z)
        {
            return Relacionado(Rels(z, ""), "officeDocument");
        }

        public static string Combinar(string pasta, string alvo)
        {
            alvo = Uri.UnescapeDataString(alvo);
            if (alvo.StartsWith("/")) return alvo.Substring(1);
            var partes = new List<string>(pasta.Length > 0 ? pasta.Split('/') : new string[0]);
            foreach (var p in alvo.Split('/'))
            {
                if (p == "..") { if (partes.Count > 0) partes.RemoveAt(partes.Count - 1); }
                else if (p != "." && p != "") partes.Add(p);
            }
            return string.Join("/", partes);
        }

        public static string DataUri(ZipArchive z, string caminho, Dictionary<string, string> cache)
        {
            if (caminho == null) return null;
            string r;
            if (cache.TryGetValue(caminho, out r)) return r;
            r = null;
            var b = Bytes(z, caminho);
            if (b != null)
            {
                string mime = null;
                switch (Path.GetExtension(caminho).ToLowerInvariant())
                {
                    case ".png": mime = "image/png"; break;
                    case ".jpg": case ".jpeg": case ".jpe": mime = "image/jpeg"; break;
                    case ".gif": mime = "image/gif"; break;
                    case ".bmp": mime = "image/bmp"; break;
                    case ".svg": mime = "image/svg+xml"; break;
                    case ".webp": mime = "image/webp"; break;
                }
                if (mime == null)
                {
                    // EMF, WMF, TIFF... viram PNG
                    try
                    {
                        using (var ms = new MemoryStream(b))
                        using (var img = Image.FromStream(ms))
                        {
                            double s = Math.Min(1.0, 2000.0 / Math.Max(img.Width, img.Height));
                            int w = Math.Max(1, (int)(img.Width * s)), h = Math.Max(1, (int)(img.Height * s));
                            using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                            using (var o = new MemoryStream())
                            {
                                using (var g = Graphics.FromImage(bmp))
                                {
                                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                    g.DrawImage(img, 0, 0, w, h);
                                }
                                bmp.Save(o, ImageFormat.Png);
                                b = o.ToArray();
                                mime = "image/png";
                            }
                        }
                    }
                    catch { b = null; }
                }
                if (b != null) r = "data:" + mime + ";base64," + Convert.ToBase64String(b);
            }
            cache[caminho] = r;
            return r;
        }

        public static double Num(string s, double padrao)
        {
            double d;
            return s != null && double.TryParse(s, NumberStyles.Float, Inv, out d) ? d : padrao;
        }

        public static string Pt(double v) { return v.ToString("0.##", Inv) + "pt"; }
        public static string Px(double v) { return v.ToString("0.##", Inv) + "px"; }

        public static string Fonte(string nome)
        {
            if (string.IsNullOrEmpty(nome)) return "Calibri,Carlito,Arial,sans-serif";
            string n = nome.ToLowerInvariant();
            bool serifa = n.Contains("times") || n.Contains("georgia") || n.Contains("cambria") || n.Contains("garamond") || n.Contains("book") || n.Contains("serif") && !n.Contains("sans");
            return "'" + nome.Replace("'", "") + "'," + (serifa ? "'Times New Roman',serif" : "Calibri,Arial,sans-serif");
        }

        public static XElement Escolha(XElement ac, bool preferirFallback)
        {
            var ch = ac.Element(MC + "Choice");
            var fb = ac.Element(MC + "Fallback");
            return preferirFallback ? (fb ?? ch) : (ch ?? fb);
        }
    }

    static class Cores
    {
        public static int[] Rgb(string hex)
        {
            hex = (hex ?? "000000").Trim().TrimStart('#');
            if (hex.Length == 8) hex = hex.Substring(2);
            if (hex.Length != 6) return new[] { 0, 0, 0 };
            int v;
            if (!int.TryParse(hex, NumberStyles.HexNumber, null, out v)) return new[] { 0, 0, 0 };
            return new[] { (v >> 16) & 255, (v >> 8) & 255, v & 255 };
        }

        public static string Hex(double r, double g, double b)
        {
            Func<double, int> c = x => Math.Max(0, Math.Min(255, (int)Math.Round(x * 255)));
            return "#" + c(r).ToString("x2") + c(g).ToString("x2") + c(b).ToString("x2");
        }

        public static void ParaHsl(double r, double g, double b, out double h, out double s, out double l)
        {
            double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
            l = (max + min) / 2;
            if (max == min) { h = 0; s = 0; return; }
            double d = max - min;
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            if (max == r) h = (g - b) / d + (g < b ? 6 : 0);
            else if (max == g) h = (b - r) / d + 2;
            else h = (r - g) / d + 4;
            h /= 6;
        }

        public static void DeHsl(double h, double s, double l, out double r, out double g, out double b)
        {
            if (s == 0) { r = g = b = l; return; }
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
            r = Canal(p, q, h + 1.0 / 3); g = Canal(p, q, h); b = Canal(p, q, h - 1.0 / 3);
        }

        static double Canal(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6) return p + (q - p) * 6 * t;
            if (t < 0.5) return q;
            if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
            return p;
        }

        static double Lim(double v) { return Math.Max(0, Math.Min(1, v)); }

        // Modificadores de cor do DrawingML (lumMod, lumOff, tint, shade, alpha...)
        public static string Aplicar(string hex, IEnumerable<XElement> mods)
        {
            var c = Rgb(hex);
            double r = c[0] / 255.0, g = c[1] / 255.0, b = c[2] / 255.0, alfa = 1;
            foreach (var m in mods)
            {
                double v = Ooxml.Num((string)m.Attribute("val"), 100000) / 100000;
                double h, s, l;
                switch (m.Name.LocalName)
                {
                    case "lumMod": ParaHsl(r, g, b, out h, out s, out l); DeHsl(h, s, Lim(l * v), out r, out g, out b); break;
                    case "lumOff": ParaHsl(r, g, b, out h, out s, out l); DeHsl(h, s, Lim(l + v), out r, out g, out b); break;
                    case "satMod": ParaHsl(r, g, b, out h, out s, out l); DeHsl(h, Lim(s * v), l, out r, out g, out b); break;
                    case "tint": r += (1 - r) * (1 - v); g += (1 - g) * (1 - v); b += (1 - b) * (1 - v); break;
                    case "shade": r *= v; g *= v; b *= v; break;
                    case "alpha": alfa = v; break;
                }
            }
            if (alfa < 0.999)
                return "rgba(" + (int)(r * 255) + "," + (int)(g * 255) + "," + (int)(b * 255) + "," + alfa.ToString("0.##", Ooxml.Inv) + ")";
            return Hex(r, g, b);
        }

        public static string TintExcel(string hex, double tint)
        {
            var c = Rgb(hex);
            double h, s, l, r, g, b;
            ParaHsl(c[0] / 255.0, c[1] / 255.0, c[2] / 255.0, out h, out s, out l);
            l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
            DeHsl(h, s, Lim(l), out r, out g, out b);
            return Hex(r, g, b);
        }
    }

    // ===================================================================== DOCX
    class DocxConv
    {
        static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        static readonly XNamespace WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
        static readonly XNamespace R = Ooxml.R, A = Ooxml.A;

        ZipArchive z;
        Dictionary<string, Rel> rels;
        readonly Dictionary<string, XElement> estilos = new Dictionary<string, XElement>();
        readonly Dictionary<string, string> cacheImg = new Dictionary<string, string>();
        readonly Dictionary<string, string> numParaAbs = new Dictionary<string, string>();
        readonly Dictionary<string, XElement> abstratos = new Dictionary<string, XElement>();
        readonly Dictionary<string, Dictionary<int, int>> inicios = new Dictionary<string, Dictionary<int, int>>();
        readonly Dictionary<string, int[]> contadores = new Dictionary<string, int[]>();
        XElement rPrPadrao, pPrPadrao;
        string estiloParPadrao, estiloTabPadrao;
        string fonteMaior = "Calibri Light", fonteMenor = "Calibri";

        class RP
        {
            public bool? B, I, U, S, Caps, Oculto;
            public double? Tam;
            public string Cor, Fonte, Realce, Vert, Fundo;
            public RP Copia() { return (RP)MemberwiseClone(); }
        }

        class PP
        {
            public string Jc, NumId, Fundo, BordaInf, BordaSup;
            public double? Antes, Depois, Esq, Dir, Primeira, Linha, LinhaPt;
            public int Nivel;
            public bool QuebraAntes, Contextual;
        }

        class Cel { public XElement Tc; public int Col, Span = 1, RowSpan = 1; public string VMerge; public bool Pular; }

        static string Val(XElement e) { return e == null ? null : (string)e.Attribute(W + "val"); }
        static XElement Filho(XElement e, string nome) { return e == null ? null : e.Element(W + nome); }
        static double Twips(string s, double padrao) { return Ooxml.Num(s, padrao * 20) / 20; }
        static string Pt(double v) { return Ooxml.Pt(v); }
        static bool Liga(string v) { return v == null || (v != "0" && v != "false" && v != "off" && v != "none"); }
        static int Int(string s) { int i; return int.TryParse(s, out i) ? i : 0; }

        public string Converter(string caminho)
        {
            using (z = ZipFile.OpenRead(caminho))
            {
                string principal = Ooxml.DocumentoPrincipal(z) ?? "word/document.xml";
                var doc = Ooxml.Xml(z, principal);
                if (doc == null) throw new Exception("Documento Word inválido.");
                rels = Ooxml.Rels(z, principal);
                CarregarTema();
                CarregarEstilos();
                CarregarNumeracao();
                var body = doc.Root.Element(W + "body");
                var sb = new StringBuilder();
                if (body != null) Bloco(body, sb);
                var rp = RpBase();
                string css = CssPagina(body == null ? null : body.Element(W + "sectPr")) +
                    "body{margin:0;font-family:" + Ooxml.Fonte(rp.Fonte) + ";font-size:" + Pt(rp.Tam ?? 10) + "}" +
                    "div.p{white-space:pre-wrap;overflow-wrap:break-word;orphans:2;widows:2}" +
                    ".tab{display:inline-block;width:36pt}.qp{display:block;break-before:page}" +
                    "table{border-collapse:collapse}td{vertical-align:top;padding:0 5.4pt}" +
                    ".caixa{display:inline-block;vertical-align:top}img{vertical-align:text-bottom}a{color:inherit;text-decoration:none}";
                return Html.Pagina(Path.GetFileName(caminho), css, sb.ToString());
            }
        }

        void CarregarTema()
        {
            var t = Ooxml.Xml(z, Ooxml.Relacionado(rels, "theme"));
            if (t == null) return;
            var fs = t.Descendants(A + "fontScheme").FirstOrDefault();
            if (fs == null) return;
            var mj = fs.Element(A + "majorFont"); var mn = fs.Element(A + "minorFont");
            if (mj != null && mj.Element(A + "latin") != null) fonteMaior = (string)mj.Element(A + "latin").Attribute("typeface") ?? fonteMaior;
            if (mn != null && mn.Element(A + "latin") != null) fonteMenor = (string)mn.Element(A + "latin").Attribute("typeface") ?? fonteMenor;
        }

        void CarregarEstilos()
        {
            var x = Ooxml.Xml(z, Ooxml.Relacionado(rels, "styles"));
            if (x == null) return;
            var dd = x.Root.Element(W + "docDefaults");
            if (dd != null)
            {
                var rd = dd.Element(W + "rPrDefault"); if (rd != null) rPrPadrao = rd.Element(W + "rPr");
                var pd = dd.Element(W + "pPrDefault"); if (pd != null) pPrPadrao = pd.Element(W + "pPr");
            }
            foreach (var s in x.Root.Elements(W + "style"))
            {
                string id = (string)s.Attribute(W + "styleId");
                if (id == null) continue;
                estilos[id] = s;
                string tipo = (string)s.Attribute(W + "type");
                if (Liga((string)s.Attribute(W + "default") ?? "0"))
                {
                    if (tipo == "paragraph") estiloParPadrao = id;
                    if (tipo == "table") estiloTabPadrao = id;
                }
            }
        }

        void CarregarNumeracao()
        {
            var x = Ooxml.Xml(z, Ooxml.Relacionado(rels, "numbering"));
            if (x == null) return;
            foreach (var a in x.Root.Elements(W + "abstractNum"))
            {
                string id = (string)a.Attribute(W + "abstractNumId");
                if (id != null) abstratos[id] = a;
            }
            foreach (var n in x.Root.Elements(W + "num"))
            {
                string id = (string)n.Attribute(W + "numId");
                if (id == null) continue;
                numParaAbs[id] = Val(n.Element(W + "abstractNumId"));
                var ini = new Dictionary<int, int>();
                foreach (var o in n.Elements(W + "lvlOverride"))
                {
                    var so = o.Element(W + "startOverride");
                    if (so != null) ini[Int((string)o.Attribute(W + "ilvl"))] = Int(Val(so));
                }
                inicios[id] = ini;
            }
        }

        RP RpBase()
        {
            var rp = new RP();
            AplicarRP(rp, rPrPadrao);
            if (rp.Fonte == null) rp.Fonte = fonteMenor;
            if (rp.Tam == null) rp.Tam = 10;
            return rp;
        }

        void AplicarRP(RP rp, XElement rPr)
        {
            if (rPr == null) return;
            foreach (var e in rPr.Elements())
            {
                string val = Val(e);
                switch (e.Name.LocalName)
                {
                    case "b": rp.B = Liga(val); break;
                    case "i": rp.I = Liga(val); break;
                    case "u": rp.U = val != null && val != "none"; break;
                    case "strike": case "dstrike": rp.S = Liga(val); break;
                    case "caps": rp.Caps = Liga(val); break;
                    case "vanish": rp.Oculto = Liga(val); break;
                    case "sz": rp.Tam = Ooxml.Num(val, 20) / 2; break;
                    case "color": if (val != null && val != "auto") rp.Cor = "#" + val; break;
                    case "highlight": rp.Realce = CorRealce(val); break;
                    case "shd":
                        string f = (string)e.Attribute(W + "fill");
                        if (f != null && f != "auto") rp.Fundo = "#" + f;
                        break;
                    case "vertAlign": rp.Vert = val; break;
                    case "rFonts":
                        string tema = (string)e.Attribute(W + "asciiTheme") ?? (string)e.Attribute(W + "hAnsiTheme");
                        string nome = (string)e.Attribute(W + "ascii") ?? (string)e.Attribute(W + "hAnsi");
                        if (tema != null) rp.Fonte = tema.StartsWith("major") ? fonteMaior : fonteMenor;
                        else if (nome != null) rp.Fonte = nome;
                        break;
                }
            }
        }

        static string CorRealce(string v)
        {
            switch (v)
            {
                case "yellow": return "#ffff00";
                case "green": return "#00ff00";
                case "cyan": return "#00ffff";
                case "magenta": return "#ff00ff";
                case "blue": return "#0000ff";
                case "red": return "#ff0000";
                case "darkBlue": return "#000080";
                case "darkCyan": return "#008080";
                case "darkGreen": return "#008000";
                case "darkMagenta": return "#800080";
                case "darkRed": return "#800000";
                case "darkYellow": return "#808000";
                case "darkGray": return "#808080";
                case "lightGray": return "#c0c0c0";
                case "black": return "#000000";
                default: return null;
            }
        }

        void AplicarPP(PP pp, XElement pPr)
        {
            if (pPr == null) return;
            foreach (var e in pPr.Elements())
            {
                string val = Val(e);
                switch (e.Name.LocalName)
                {
                    case "jc": pp.Jc = val; break;
                    case "spacing":
                        if (e.Attribute(W + "before") != null) pp.Antes = Twips((string)e.Attribute(W + "before"), 0);
                        if (e.Attribute(W + "after") != null) pp.Depois = Twips((string)e.Attribute(W + "after"), 0);
                        if (e.Attribute(W + "beforeAutospacing") != null && Liga((string)e.Attribute(W + "beforeAutospacing"))) pp.Antes = 14;
                        if (e.Attribute(W + "afterAutospacing") != null && Liga((string)e.Attribute(W + "afterAutospacing"))) pp.Depois = 14;
                        if (e.Attribute(W + "line") != null)
                        {
                            double l = Ooxml.Num((string)e.Attribute(W + "line"), 240);
                            string regra = (string)e.Attribute(W + "lineRule") ?? "auto";
                            if (regra == "auto") { pp.Linha = l / 240; pp.LinhaPt = null; }
                            else { pp.LinhaPt = l / 20; pp.Linha = null; }
                        }
                        break;
                    case "ind":
                        string esq = (string)e.Attribute(W + "left") ?? (string)e.Attribute(W + "start");
                        string dir = (string)e.Attribute(W + "right") ?? (string)e.Attribute(W + "end");
                        if (esq != null) pp.Esq = Twips(esq, 0);
                        if (dir != null) pp.Dir = Twips(dir, 0);
                        if (e.Attribute(W + "firstLine") != null) pp.Primeira = Twips((string)e.Attribute(W + "firstLine"), 0);
                        if (e.Attribute(W + "hanging") != null) pp.Primeira = -Twips((string)e.Attribute(W + "hanging"), 0);
                        break;
                    case "numPr":
                        string nid = Val(e.Element(W + "numId"));
                        if (nid != null) pp.NumId = nid;
                        var il = e.Element(W + "ilvl");
                        if (il != null) pp.Nivel = Int(Val(il));
                        break;
                    case "pageBreakBefore": pp.QuebraAntes = Liga(val); break;
                    case "contextualSpacing": pp.Contextual = Liga(val); break;
                    case "shd":
                        string f = (string)e.Attribute(W + "fill");
                        if (f != null && f != "auto") pp.Fundo = "#" + f;
                        break;
                    case "pBdr":
                        pp.BordaInf = Borda(e.Element(W + "bottom"));
                        pp.BordaSup = Borda(e.Element(W + "top"));
                        break;
                }
            }
        }

        static string Borda(XElement b)
        {
            if (b == null) return null;
            string v = Val(b);
            if (v == null || v == "nil" || v == "none") return null;
            string cor = (string)b.Attribute(W + "color");
            double sz = Ooxml.Num((string)b.Attribute(W + "sz"), 4) / 8;
            return Pt(Math.Max(0.5, sz)) + (v == "dashed" ? " dashed " : v == "dotted" ? " dotted " : v == "double" ? " double " : " solid ") +
                   (cor == null || cor == "auto" ? "#000" : "#" + cor);
        }

        void EstiloEm(string id, RP rp, PP pp, int prof)
        {
            XElement s;
            if (id == null || prof > 20 || !estilos.TryGetValue(id, out s)) return;
            EstiloEm(Val(s.Element(W + "basedOn")), rp, pp, prof + 1);
            if (pp != null) AplicarPP(pp, s.Element(W + "pPr"));
            if (rp != null) AplicarRP(rp, s.Element(W + "rPr"));
        }

        string EstiloDe(XElement p)
        {
            if (p == null || p.Name != W + "p") return null;
            var pPr = p.Element(W + "pPr");
            return Val(Filho(pPr, "pStyle")) ?? estiloParPadrao;
        }

        void Bloco(XElement c, StringBuilder sb)
        {
            foreach (var e in c.Elements())
            {
                switch (e.Name.LocalName)
                {
                    case "p": Paragrafo(e, sb); break;
                    case "tbl": Tabela(e, sb); break;
                    case "sdt":
                        var sc = e.Element(W + "sdtContent");
                        if (sc != null) Bloco(sc, sb);
                        break;
                    case "customXml": case "ins": case "smartTag": Bloco(e, sb); break;
                    case "AlternateContent":
                        var ch = Ooxml.Escolha(e, false);
                        if (ch != null) Bloco(ch, sb);
                        break;
                }
            }
        }

        void Paragrafo(XElement p, StringBuilder sb)
        {
            var pPr = p.Element(W + "pPr");
            string sid = EstiloDe(p);
            var rp = RpBase();
            var pp = new PP();
            AplicarPP(pp, pPrPadrao);
            EstiloEm(sid, rp, pp, 0);
            bool indDireto = pPr != null && pPr.Element(W + "ind") != null;
            AplicarPP(pp, pPr);
            var marca = rp.Copia();
            if (pPr != null) AplicarRP(marca, pPr.Element(W + "rPr"));

            if (pp.Contextual)
            {
                var ant = p.ElementsBeforeSelf().LastOrDefault();
                var prox = p.ElementsAfterSelf().FirstOrDefault();
                if (EstiloDe(ant) == sid) pp.Antes = 0;
                if (EstiloDe(prox) == sid) pp.Depois = 0;
            }

            string prefixo = "";
            if (pp.NumId != null && pp.NumId != "0")
            {
                double esq, pend;
                string m = Marcador(pp.NumId, pp.Nivel, out esq, out pend);
                if (m != null)
                {
                    if (!indDireto && esq >= 0) { pp.Esq = esq; pp.Primeira = -pend; }
                    double larg = pp.Primeira.HasValue && pp.Primeira < 0 ? -pp.Primeira.Value : 0;
                    if (m.Length > 0)
                        prefixo = "<span style=\"display:inline-block;min-width:" + Pt(larg) + ";text-indent:0;font-family:" +
                                  Ooxml.Fonte(rp.Fonte) + ";font-size:" + Pt(rp.Tam ?? 10) + ";" + (rp.B == true ? "font-weight:bold;" : "") +
                                  (rp.Cor != null ? "color:" + rp.Cor : "") + "\">" + m + "</span>";
                }
            }

            var conteudo = new StringBuilder();
            Inline(p, rp, conteudo);
            string html = conteudo.ToString();
            if (html.Length == 0 && prefixo.Length == 0)
                html = "<span style=\"font-size:" + Pt(marca.Tam ?? 10) + "\">&nbsp;</span>";

            var css = new StringBuilder();
            css.Append("margin:" + Pt(pp.Antes ?? 0) + " " + Pt(Math.Max(0, pp.Dir ?? 0)) + " " + Pt(pp.Depois ?? 0) + " " + Pt(Math.Max(0, pp.Esq ?? 0)) + ";");
            if (pp.Primeira.HasValue && pp.Primeira.Value != 0) css.Append("text-indent:" + Pt(pp.Primeira.Value) + ";");
            if (pp.Jc != null) css.Append("text-align:" + Alinhamento(pp.Jc) + ";");
            css.Append("font-size:" + Pt(rp.Tam ?? 10) + ";");
            if (pp.LinhaPt.HasValue) css.Append("line-height:" + Pt(pp.LinhaPt.Value) + ";");
            else css.Append("line-height:" + ((pp.Linha ?? 1.0) * 1.17).ToString("0.###", Ooxml.Inv) + ";");
            if (pp.Fundo != null) css.Append("background:" + pp.Fundo + ";");
            if (pp.BordaInf != null) css.Append("border-bottom:" + pp.BordaInf + ";padding-bottom:1pt;");
            if (pp.BordaSup != null) css.Append("border-top:" + pp.BordaSup + ";padding-top:1pt;");
            if (pp.QuebraAntes) css.Append("break-before:page;");
            sb.Append("<div class=\"p\" style=\"" + css + "\">" + prefixo + html + "</div>");
            if (pPr != null && pPr.Element(W + "sectPr") != null && p.ElementsAfterSelf().Any())
                sb.Append("<div style=\"break-after:page\"></div>");
        }

        static string Alinhamento(string jc)
        {
            switch (jc)
            {
                case "center": return "center";
                case "right": case "end": return "right";
                case "both": case "distribute": return "justify";
                default: return "left";
            }
        }

        void Inline(XElement pai, RP rpPar, StringBuilder sb)
        {
            foreach (var e in pai.Elements())
            {
                switch (e.Name.LocalName)
                {
                    case "r": Run(e, rpPar, sb); break;
                    case "hyperlink":
                        string href = null;
                        string rid = (string)e.Attribute(R + "id");
                        Rel rl;
                        if (rid != null && rels.TryGetValue(rid, out rl)) href = rl.Alvo;
                        if (href != null) sb.Append("<a href=\"" + Html.Esc(href) + "\">");
                        Inline(e, rpPar, sb);
                        if (href != null) sb.Append("</a>");
                        break;
                    case "ins": case "smartTag": case "fldSimple": case "customXml": case "bdo": case "dir": case "moveTo":
                        Inline(e, rpPar, sb);
                        break;
                    case "sdt":
                        var sc = e.Element(W + "sdtContent");
                        if (sc != null) Inline(sc, rpPar, sb);
                        break;
                    case "AlternateContent":
                        var ch = Ooxml.Escolha(e, false);
                        if (ch != null) Inline(ch, rpPar, sb);
                        break;
                    case "oMath": case "oMathPara":
                        sb.Append(Html.Esc(string.Concat(e.Descendants().Where(x => x.Name.LocalName == "t").Select(x => x.Value))));
                        break;
                }
            }
        }

        void Run(XElement r, RP rpPar, StringBuilder sb)
        {
            var rp = rpPar.Copia();
            var rPr = r.Element(W + "rPr");
            if (rPr != null)
            {
                EstiloEm(Val(rPr.Element(W + "rStyle")), rp, null, 0);
                AplicarRP(rp, rPr);
            }
            if (rp.Oculto == true) return;
            var c = new StringBuilder();
            ConteudoRun(r.Elements(), rp, c);
            if (c.Length == 0) return;
            sb.Append("<span style=\"" + CssRun(rp) + "\">" + c + "</span>");
        }

        void ConteudoRun(IEnumerable<XElement> els, RP rp, StringBuilder c)
        {
            foreach (var e in els)
            {
                switch (e.Name.LocalName)
                {
                    case "t":
                        string t = e.Value;
                        if (rp.Caps == true) t = t.ToUpper();
                        c.Append(Html.Esc(t));
                        break;
                    case "tab": c.Append("<span class=\"tab\"></span>"); break;
                    case "br": c.Append((string)e.Attribute(W + "type") == "page" ? "<span class=\"qp\"></span>" : "<br>"); break;
                    case "cr": c.Append("<br>"); break;
                    case "noBreakHyphen": c.Append("-"); break;
                    case "softHyphen": c.Append("&shy;"); break;
                    case "sym": c.Append(Simbolo(e)); break;
                    case "drawing": c.Append(Desenho(e)); break;
                    case "pict": case "object": c.Append(Vml(e)); break;
                    case "AlternateContent":
                        var ch = Ooxml.Escolha(e, false);
                        if (ch != null) ConteudoRun(ch.Elements(), rp, c);
                        break;
                }
            }
        }

        static string Simbolo(XElement e)
        {
            int code;
            if (!int.TryParse((string)e.Attribute(W + "char") ?? "", NumberStyles.HexNumber, null, out code)) return "";
            if (code < 0x100) code += 0xF000;
            string fonte = (string)e.Attribute(W + "font") ?? "Symbol";
            return "<span style=\"font-family:'" + fonte + "'\">&#x" + code.ToString("X") + ";</span>";
        }

        static string CssRun(RP rp)
        {
            var s = new StringBuilder();
            s.Append("font-family:" + Ooxml.Fonte(rp.Fonte) + ";");
            double tam = rp.Tam ?? 10;
            if (rp.Vert == "superscript" || rp.Vert == "subscript")
            {
                s.Append("vertical-align:" + (rp.Vert == "superscript" ? "super" : "sub") + ";");
                tam *= 0.65;
            }
            s.Append("font-size:" + Pt(tam) + ";");
            if (rp.B == true) s.Append("font-weight:bold;");
            if (rp.I == true) s.Append("font-style:italic;");
            var dec = new List<string>();
            if (rp.U == true) dec.Add("underline");
            if (rp.S == true) dec.Add("line-through");
            if (dec.Count > 0) s.Append("text-decoration:" + string.Join(" ", dec) + ";");
            if (rp.Cor != null) s.Append("color:" + rp.Cor + ";");
            if (rp.Realce != null) s.Append("background:" + rp.Realce + ";");
            else if (rp.Fundo != null) s.Append("background:" + rp.Fundo + ";");
            return s.ToString();
        }

        string Imagem(string id)
        {
            Rel rl;
            if (id == null || !rels.TryGetValue(id, out rl) || rl.Externo) return null;
            return Ooxml.DataUri(z, rl.Alvo, cacheImg);
        }

        static double Emu(XElement e, string at) { return e == null ? 0 : Ooxml.Num((string)e.Attribute(at), 0) / 12700; }

        string Desenho(XElement d)
        {
            var sb = new StringBuilder();
            var ext = d.Descendants(WP + "extent").FirstOrDefault();
            double w = Emu(ext, "cx"), h = Emu(ext, "cy");
            var blips = d.Descendants(A + "blip").ToList();
            foreach (var blip in blips)
            {
                string uri = Imagem((string)blip.Attribute(R + "embed"));
                if (uri == null) continue;
                double bw = w, bh = h;
                if (blips.Count > 1)
                {
                    var pic = blip.Ancestors().FirstOrDefault(x => x.Name.LocalName == "pic");
                    var ax = pic == null ? null : pic.Descendants(A + "ext").FirstOrDefault();
                    if (ax != null) { bw = Emu(ax, "cx"); bh = Emu(ax, "cy"); }
                }
                sb.Append("<img src=\"" + uri + "\" style=\"width:" + Pt(bw) + ";height:" + Pt(bh) + "\">");
            }
            foreach (var tx in d.Descendants(W + "txbxContent"))
            {
                var inner = new StringBuilder();
                Bloco(tx, inner);
                sb.Append("<span class=\"caixa\" style=\"width:" + Pt(w) + "\">" + inner + "</span>");
            }
            return sb.ToString();
        }

        string Vml(XElement e)
        {
            var sb = new StringBuilder();
            foreach (var im in e.Descendants(V + "imagedata"))
            {
                string uri = Imagem((string)im.Attribute(R + "id"));
                if (uri == null) continue;
                string st = im.Parent == null ? "" : ((string)im.Parent.Attribute("style") ?? "");
                double w = Dim(st, "width"), h = Dim(st, "height");
                sb.Append("<img src=\"" + uri + "\" style=\"" + (w > 0 ? "width:" + Pt(w) + ";" : "") + (h > 0 ? "height:" + Pt(h) : "") + "\">");
            }
            foreach (var tx in e.Descendants(W + "txbxContent"))
            {
                var inner = new StringBuilder();
                Bloco(tx, inner);
                sb.Append("<span class=\"caixa\">" + inner + "</span>");
            }
            return sb.ToString();
        }

        static double Dim(string estilo, string nome)
        {
            var m = Regex.Match(estilo, "(?:^|;)\\s*" + nome + "\\s*:\\s*([\\d.]+)\\s*(pt|in|px|cm|mm)?", RegexOptions.IgnoreCase);
            if (!m.Success) return 0;
            double v = Ooxml.Num(m.Groups[1].Value, 0);
            switch (m.Groups[2].Value.ToLowerInvariant())
            {
                case "in": return v * 72;
                case "px": return v * 0.75;
                case "cm": return v * 72 / 2.54;
                case "mm": return v * 72 / 25.4;
                default: return v;
            }
        }

        IEnumerable<XElement> Celulas(XElement tr)
        {
            foreach (var e in tr.Elements())
            {
                if (e.Name == W + "tc") yield return e;
                else if (e.Name == W + "sdt" || e.Name == W + "customXml")
                {
                    var c = e.Name == W + "sdt" ? e.Element(W + "sdtContent") : e;
                    if (c != null) foreach (var tc in c.Elements(W + "tc")) yield return tc;
                }
            }
        }

        string BordasTabela(XElement tblPr, string est)
        {
            var b = Filho(tblPr, "tblBorders");
            int prof = 0;
            while (b == null && est != null && prof++ < 10)
            {
                XElement s;
                if (!estilos.TryGetValue(est, out s)) break;
                b = Filho(s.Element(W + "tblPr"), "tblBorders");
                est = Val(s.Element(W + "basedOn"));
            }
            if (b == null) return null;
            foreach (var x in b.Elements())
            {
                string css = Borda(x);
                if (css != null) return css;
            }
            return null;
        }

        static string BordasCelula(XElement tcPr, string padrao)
        {
            var lados = new Dictionary<string, string> { { "top", padrao }, { "right", padrao }, { "bottom", padrao }, { "left", padrao } };
            var tb = Filho(tcPr, "tcBorders");
            if (tb != null)
                foreach (var x in tb.Elements())
                {
                    string lado = x.Name.LocalName == "start" ? "left" : (x.Name.LocalName == "end" ? "right" : x.Name.LocalName);
                    if (lados.ContainsKey(lado)) lados[lado] = Borda(x);
                }
            var s = new StringBuilder();
            foreach (var kv in lados) s.Append("border-" + kv.Key + ":" + (kv.Value ?? "none") + ";");
            return s.ToString();
        }

        void Tabela(XElement tbl, StringBuilder sb)
        {
            var tblPr = tbl.Element(W + "tblPr");
            string est = Val(Filho(tblPr, "tblStyle")) ?? estiloTabPadrao;
            string borda = BordasTabela(tblPr, est);
            var grade = tbl.Element(W + "tblGrid");
            var cols = grade == null ? new List<double>() : grade.Elements(W + "gridCol").Select(g => Twips((string)g.Attribute(W + "w"), 0)).ToList();
            var linhas = tbl.Elements(W + "tr").ToList();
            var mapa = new List<List<Cel>>();
            foreach (var tr in linhas)
            {
                var lista = new List<Cel>();
                int col = Int(Val(Filho(tr.Element(W + "trPr"), "gridBefore")));
                foreach (var tc in Celulas(tr))
                {
                    var tcPr = tc.Element(W + "tcPr");
                    int span = Math.Max(1, Int(Val(Filho(tcPr, "gridSpan")) ?? "1"));
                    var vm = Filho(tcPr, "vMerge");
                    lista.Add(new Cel { Tc = tc, Col = col, Span = span, VMerge = vm == null ? null : (Val(vm) ?? "continue") });
                    col += span;
                }
                mapa.Add(lista);
            }
            for (int i = 0; i < mapa.Count; i++)
                foreach (var c in mapa[i])
                {
                    if (c.VMerge == "continue") { c.Pular = true; continue; }
                    if (c.VMerge != "restart") continue;
                    for (int j = i + 1; j < mapa.Count; j++)
                    {
                        var baixo = mapa[j].FirstOrDefault(x => x.Col == c.Col);
                        if (baixo == null || baixo.VMerge != "continue") break;
                        c.RowSpan++;
                    }
                }
            double total = cols.Sum();
            string jc = Val(Filho(tblPr, "jc"));
            var ind = Filho(tblPr, "tblInd");
            sb.Append("<table style=\"" + (total > 0 ? "width:" + Pt(total) + ";table-layout:fixed;" : "") +
                      (jc == "center" ? "margin-left:auto;margin-right:auto;" : "") +
                      (ind != null ? "margin-left:" + Pt(Twips((string)ind.Attribute(W + "w"), 0)) + ";" : "") + "\">");
            if (cols.Count > 0)
            {
                sb.Append("<colgroup>");
                foreach (var w in cols) sb.Append("<col style=\"width:" + Pt(w) + "\">");
                sb.Append("</colgroup>");
            }
            for (int i = 0; i < mapa.Count; i++)
            {
                var trH = Filho(linhas[i].Element(W + "trPr"), "trHeight");
                sb.Append("<tr" + (trH != null ? " style=\"height:" + Pt(Twips(Val(trH), 0)) + "\"" : "") + ">");
                foreach (var c in mapa[i])
                {
                    if (c.Pular) continue;
                    var tcPr = c.Tc.Element(W + "tcPr");
                    var css = new StringBuilder(BordasCelula(tcPr, borda));
                    var shd = Filho(tcPr, "shd");
                    string fill = shd == null ? null : (string)shd.Attribute(W + "fill");
                    if (fill != null && fill != "auto") css.Append("background:#" + fill + ";");
                    string va = Val(Filho(tcPr, "vAlign"));
                    if (va == "center") css.Append("vertical-align:middle;");
                    else if (va == "bottom") css.Append("vertical-align:bottom;");
                    var inner = new StringBuilder();
                    Bloco(c.Tc, inner);
                    sb.Append("<td" + (c.Span > 1 ? " colspan=\"" + c.Span + "\"" : "") + (c.RowSpan > 1 ? " rowspan=\"" + c.RowSpan + "\"" : "") +
                              " style=\"" + css + "\">" + inner + "</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
        }

        int Inicio(string numId, int nivel, XElement lvl)
        {
            Dictionary<int, int> ini;
            int v;
            if (inicios.TryGetValue(numId, out ini) && ini.TryGetValue(nivel, out v)) return v;
            return lvl == null ? 1 : Int(Val(lvl.Element(W + "start")) ?? "1");
        }

        string Marcador(string numId, int nivel, out double esq, out double pend)
        {
            esq = -1; pend = 0;
            string absId;
            XElement abs;
            if (!numParaAbs.TryGetValue(numId, out absId) || absId == null || !abstratos.TryGetValue(absId, out abs)) return null;
            var lvls = abs.Elements(W + "lvl").ToList();
            Func<int, XElement> lvl = n => lvls.FirstOrDefault(l => Int((string)l.Attribute(W + "ilvl")) == n);
            var L = lvl(nivel);
            if (L == null) return null;
            int[] cont;
            if (!contadores.TryGetValue(numId, out cont))
            {
                cont = Enumerable.Repeat(int.MinValue, 9).ToArray();
                contadores[numId] = cont;
            }
            nivel = Math.Max(0, Math.Min(8, nivel));
            cont[nivel] = cont[nivel] == int.MinValue ? Inicio(numId, nivel, L) : cont[nivel] + 1;
            for (int k = nivel + 1; k < 9; k++) cont[k] = int.MinValue;

            var ind = Filho(L.Element(W + "pPr"), "ind");
            if (ind != null)
            {
                esq = Twips((string)ind.Attribute(W + "left") ?? (string)ind.Attribute(W + "start"), 0);
                pend = Twips((string)ind.Attribute(W + "hanging"), 0);
            }
            string fmt = Val(L.Element(W + "numFmt")) ?? "decimal";
            string txt = Val(L.Element(W + "lvlText")) ?? "";
            if (fmt == "none") return "";
            if (fmt == "bullet") return Html.Esc(Bullet(txt));
            string r = Regex.Replace(txt, "%([1-9])", m =>
            {
                int k = m.Groups[1].Value[0] - '1';
                var Lk = lvl(k);
                int v = cont[k] == int.MinValue ? Inicio(numId, k, Lk) : cont[k];
                return FormatarNum(v, Val(Lk == null ? null : Lk.Element(W + "numFmt")) ?? "decimal");
            });
            return Html.Esc(r);
        }

        static string Bullet(string t)
        {
            if (string.IsNullOrEmpty(t)) return "•";
            char c = t[0];
            switch (c)
            {
                case '': case '·': return "•";
                case '': case '': return "▪";
                case 'o': return "◦";
                case '': return "➢";
                case '': return "✓";
                case '': return "❖";
                case '': return "◆";
            }
            if (c >= '' && c <= '') return "•";
            return t;
        }

        static string FormatarNum(int n, string fmt)
        {
            switch (fmt)
            {
                case "decimalZero": return n.ToString("00");
                case "lowerLetter": return Letras(n).ToLowerInvariant();
                case "upperLetter": return Letras(n);
                case "lowerRoman": return Romano(n).ToLowerInvariant();
                case "upperRoman": return Romano(n);
                case "ordinal": return n + "º";
                default: return n.ToString();
            }
        }

        static string Letras(int n)
        {
            if (n < 1) return n.ToString();
            int rep = (n - 1) / 26 + 1;
            return new string((char)('A' + (n - 1) % 26), rep);
        }

        public static string Romano(int n)
        {
            if (n < 1 || n > 3999) return n.ToString();
            int[] v = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] s = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            var sb = new StringBuilder();
            for (int i = 0; i < v.Length; i++) while (n >= v[i]) { sb.Append(s[i]); n -= v[i]; }
            return sb.ToString();
        }

        static string CssPagina(XElement sect)
        {
            double w = 595.3, h = 841.9, mt = 72, mb = 72, ml = 72, mr = 72;
            if (sect != null)
            {
                var sz = sect.Element(W + "pgSz");
                if (sz != null)
                {
                    w = Twips((string)sz.Attribute(W + "w"), w);
                    h = Twips((string)sz.Attribute(W + "h"), h);
                }
                var mg = sect.Element(W + "pgMar");
                if (mg != null)
                {
                    mt = Math.Abs(Twips((string)mg.Attribute(W + "top"), mt));
                    mb = Math.Abs(Twips((string)mg.Attribute(W + "bottom"), mb));
                    ml = Twips((string)mg.Attribute(W + "left"), ml);
                    mr = Twips((string)mg.Attribute(W + "right"), mr);
                }
            }
            return "@page { size: " + Pt(w) + " " + Pt(h) + "; margin: " + Pt(mt) + " " + Pt(mr) + " " + Pt(mb) + " " + Pt(ml) + "; }";
        }
    }

    // ===================================================================== XLSX
    class XlsxConv
    {
        static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        static readonly XNamespace R = Ooxml.R, A = Ooxml.A;
        const int MaxLinhas = 5000, MaxColunas = 200;

        ZipArchive z;
        readonly List<string> textos = new List<string>();
        readonly Dictionary<int, string> formatos = new Dictionary<int, string>();
        readonly List<Xf> xfs = new List<Xf>();
        readonly List<string> cssFonte = new List<string>(), cssFundo = new List<string>(), cssBorda = new List<string>();
        readonly List<string> tema = new List<string>();
        bool d1904;
        readonly CultureInfo cult = CultureInfo.CurrentCulture;

        class Xf { public int Fmt, Fonte, Fundo, Borda; public string H, V; public bool Quebra; }

        static int Int(string s) { int i; return int.TryParse(s, out i) ? i : 0; }

        public string Converter(string caminho, Opcoes o)
        {
            using (z = ZipFile.OpenRead(caminho))
            {
                string wbPath = Ooxml.DocumentoPrincipal(z) ?? "xl/workbook.xml";
                var wb = Ooxml.Xml(z, wbPath);
                if (wb == null) throw new Exception("Planilha inválida.");
                var rels = Ooxml.Rels(z, wbPath);
                var pr = wb.Root.Element(S + "workbookPr");
                string d19 = pr == null ? null : (string)pr.Attribute("date1904");
                d1904 = d19 == "1" || d19 == "true";
                CarregarTema(rels);
                CarregarTextos(rels);
                CarregarEstilos(rels);

                var sheets = wb.Root.Element(S + "sheets");
                var folhas = sheets == null ? new List<XElement>() : sheets.Elements(S + "sheet")
                    .Where(s => (string)s.Attribute("state") != "hidden" && (string)s.Attribute("state") != "veryHidden").ToList();
                var sb = new StringBuilder();
                int n = 0;
                foreach (var f in folhas)
                {
                    Rel rl;
                    string rid = (string)f.Attribute(R + "id");
                    if (rid == null || !rels.TryGetValue(rid, out rl)) continue;
                    string html = Folha(rl.Alvo, (string)f.Attribute("name"), folhas.Count > 1, o);
                    if (html != null) { sb.Append(html); n++; }
                }
                if (n == 0) sb.Append("<p>(planilha vazia)</p>");
                string papel = o.Papel == 1 ? "letter" : (o.Papel == 2 ? "legal" : "A4");
                string css = "@page { size: " + papel + "; margin: 10mm } @page paisagem { size: " + papel + " landscape; margin: 10mm }" +
                    ".paisagem{page:paisagem} .folha + .folha{break-before:page}" +
                    "body{margin:0;font-family:Calibri,Carlito,Arial,sans-serif;font-size:11pt}" +
                    "h2{font:bold 12pt 'Segoe UI',Arial,sans-serif;margin:0 0 6px;color:#333}.aviso{font:9pt Arial;color:#888;margin-top:4px}" +
                    "table{border-collapse:collapse;table-layout:fixed}tr{break-inside:avoid}" +
                    "td{border:1px solid #dadada;padding:0 3px;white-space:nowrap;overflow:hidden;vertical-align:bottom}";
                return Html.Pagina(Path.GetFileName(caminho), css, sb.ToString());
            }
        }

        void CarregarTema(Dictionary<string, Rel> rels)
        {
            var t = Ooxml.Xml(z, Ooxml.Relacionado(rels, "theme"));
            var cs = t == null ? null : t.Descendants(A + "clrScheme").FirstOrDefault();
            var cores = new List<string>();
            if (cs != null)
                foreach (var c in cs.Elements())
                {
                    var s = c.Element(A + "srgbClr");
                    var y = c.Element(A + "sysClr");
                    cores.Add(s != null ? (string)s.Attribute("val") : (y != null ? (string)y.Attribute("lastClr") ?? "000000" : "000000"));
                }
            while (cores.Count < 12) cores.Add("000000");
            // No Excel os índices 0/1 e 2/3 vêm trocados em relação ao XML (lt1, dk1, lt2, dk2)
            tema.AddRange(new[] { cores[1], cores[0], cores[3], cores[2] });
            tema.AddRange(cores.Skip(4));
        }

        void CarregarTextos(Dictionary<string, Rel> rels)
        {
            var x = Ooxml.Xml(z, Ooxml.Relacionado(rels, "sharedStrings"));
            if (x == null) return;
            foreach (var si in x.Root.Elements(S + "si")) textos.Add(TextoRico(si));
        }

        static string TextoRico(XElement e)
        {
            var sb = new StringBuilder();
            foreach (var t in e.Elements(S + "t")) sb.Append(t.Value);
            foreach (var r in e.Elements(S + "r")) foreach (var t in r.Elements(S + "t")) sb.Append(t.Value);
            return sb.ToString();
        }

        static readonly string[] Indexadas = {
            "000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF","000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF",
            "800000","008000","000080","808000","800080","008080","C0C0C0","808080","9999FF","993366","FFFFCC","CCFFFF","660066","FF8080","0066CC","CCCCFF",
            "000080","FF00FF","FFFF00","00FFFF","800080","800000","008080","0000FF","00CCFF","CCFFFF","CCFFCC","FFFF99","99CCFF","FF99CC","CC99FF","FFCC99",
            "3366FF","33CCCC","99CC00","FFCC00","FF9900","FF6600","666699","969696","003366","339966","003300","333300","993300","993366","333399","333333" };

        string Cor(XElement c)
        {
            if (c == null) return null;
            string hex = null;
            if (c.Attribute("rgb") != null) hex = (string)c.Attribute("rgb");
            else if (c.Attribute("theme") != null) { int i = Int((string)c.Attribute("theme")); if (i >= 0 && i < tema.Count) hex = tema[i]; }
            else if (c.Attribute("indexed") != null)
            {
                int i = Int((string)c.Attribute("indexed"));
                if (i == 64) hex = "000000";
                else if (i >= 0 && i < Indexadas.Length) hex = Indexadas[i];
            }
            if (hex == null) return null;
            double tint = Ooxml.Num((string)c.Attribute("tint"), 0);
            if (tint != 0) return Cores.TintExcel(hex, tint);
            var rgb = Cores.Rgb(hex);
            return Cores.Hex(rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0);
        }

        void CarregarEstilos(Dictionary<string, Rel> rels)
        {
            var x = Ooxml.Xml(z, Ooxml.Relacionado(rels, "styles"));
            if (x == null) { cssFonte.Add(""); return; }
            var root = x.Root;
            var nf = root.Element(S + "numFmts");
            if (nf != null)
                foreach (var f in nf.Elements(S + "numFmt"))
                    formatos[Int((string)f.Attribute("numFmtId"))] = (string)f.Attribute("formatCode");

            var fonts = root.Element(S + "fonts");
            if (fonts != null)
                foreach (var f in fonts.Elements(S + "font"))
                {
                    var c = new StringBuilder();
                    if (f.Element(S + "b") != null && (string)f.Element(S + "b").Attribute("val") != "0") c.Append("font-weight:bold;");
                    if (f.Element(S + "i") != null && (string)f.Element(S + "i").Attribute("val") != "0") c.Append("font-style:italic;");
                    if (f.Element(S + "u") != null) c.Append("text-decoration:underline;");
                    if (f.Element(S + "strike") != null) c.Append("text-decoration:line-through;");
                    var sz = f.Element(S + "sz");
                    if (sz != null) c.Append("font-size:" + Ooxml.Pt(Ooxml.Num((string)sz.Attribute("val"), 11)) + ";");
                    string cor = Cor(f.Element(S + "color"));
                    if (cor != null) c.Append("color:" + cor + ";");
                    var nm = f.Element(S + "name");
                    if (nm != null) c.Append("font-family:" + Ooxml.Fonte((string)nm.Attribute("val")) + ";");
                    cssFonte.Add(c.ToString());
                }
            if (cssFonte.Count == 0) cssFonte.Add("");

            var fills = root.Element(S + "fills");
            if (fills != null)
                foreach (var f in fills.Elements(S + "fill"))
                {
                    string css = null;
                    var pf = f.Element(S + "patternFill");
                    if (pf != null)
                    {
                        string tipo = (string)pf.Attribute("patternType");
                        if (tipo != null && tipo != "none" && tipo != "gray125")
                            css = Cor(pf.Element(S + "fgColor")) ?? Cor(pf.Element(S + "bgColor"));
                    }
                    var gf = f.Element(S + "gradientFill");
                    if (gf != null)
                    {
                        var st = gf.Elements(S + "stop").FirstOrDefault();
                        if (st != null) css = Cor(st.Element(S + "color"));
                    }
                    cssFundo.Add(css);
                }

            var borders = root.Element(S + "borders");
            if (borders != null)
                foreach (var b in borders.Elements(S + "border"))
                {
                    var c = new StringBuilder();
                    foreach (var lado in new[] { "left", "right", "top", "bottom" })
                    {
                        var e = b.Element(S + lado);
                        string est = e == null ? null : (string)e.Attribute("style");
                        if (est == null || est == "none") continue;
                        string larg = est == "medium" || est == "mediumDashed" ? "2px" : (est == "thick" || est == "double" ? "3px" : "1px");
                        string tipo = est.Contains("ashed") ? "dashed" : (est.Contains("otted") || est == "hair" ? "dotted" : (est == "double" ? "double" : "solid"));
                        c.Append("border-" + lado + ":" + larg + " " + tipo + " " + (Cor(e.Element(S + "color")) ?? "#000") + ";");
                    }
                    cssBorda.Add(c.ToString());
                }

            var cellXfs = root.Element(S + "cellXfs");
            if (cellXfs != null)
                foreach (var e in cellXfs.Elements(S + "xf"))
                {
                    var xf = new Xf
                    {
                        Fmt = Int((string)e.Attribute("numFmtId")),
                        Fonte = Int((string)e.Attribute("fontId")),
                        Fundo = Int((string)e.Attribute("fillId")),
                        Borda = Int((string)e.Attribute("borderId"))
                    };
                    var al = e.Element(S + "alignment");
                    if (al != null)
                    {
                        xf.H = (string)al.Attribute("horizontal");
                        xf.V = (string)al.Attribute("vertical");
                        string wt = (string)al.Attribute("wrapText");
                        xf.Quebra = wt == "1" || wt == "true";
                    }
                    xfs.Add(xf);
                }
        }

        static int ColDeRef(string r)
        {
            int c = 0;
            foreach (char ch in r)
            {
                if (ch >= 'A' && ch <= 'Z') c = c * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') c = c * 26 + (ch - 'a' + 1);
                else break;
            }
            return c;
        }

        static int LinDeRef(string r) { return Int(new string(r.Where(char.IsDigit).ToArray())); }
        static long Chave(int r, int c) { return ((long)r << 16) | (uint)c; }

        bool Estilizado(int s)
        {
            if (s <= 0 || s >= xfs.Count) return false;
            var x = xfs[s];
            return (x.Fundo < cssFundo.Count && cssFundo[x.Fundo] != null) || (x.Borda < cssBorda.Count && cssBorda[x.Borda].Length > 0);
        }

        string Folha(string caminho, string nome, bool titulo, Opcoes o)
        {
            var x = Ooxml.Xml(z, caminho);
            if (x == null) return null;
            var root = x.Root;

            double larguraPadrao = 64, alturaPadrao = 20;
            var sfp = root.Element(S + "sheetFormatPr");
            if (sfp != null)
            {
                if (sfp.Attribute("defaultColWidth") != null) larguraPadrao = Ooxml.Num((string)sfp.Attribute("defaultColWidth"), 8.43) * 7 + 5;
                else if (sfp.Attribute("baseColWidth") != null) larguraPadrao = Ooxml.Num((string)sfp.Attribute("baseColWidth"), 8) * 7 + 10;
                if (sfp.Attribute("defaultRowHeight") != null) alturaPadrao = Ooxml.Num((string)sfp.Attribute("defaultRowHeight"), 15) * 4 / 3;
            }
            var larguras = new Dictionary<int, double>();
            var colOcultas = new HashSet<int>();
            var colsEl = root.Element(S + "cols");
            if (colsEl != null)
                foreach (var c in colsEl.Elements(S + "col"))
                {
                    int mi = Int((string)c.Attribute("min")), ma = Math.Min(Int((string)c.Attribute("max")), MaxColunas);
                    double w = c.Attribute("width") != null ? Ooxml.Num((string)c.Attribute("width"), 8.43) * 7 + 5 : larguraPadrao;
                    bool oc = (string)c.Attribute("hidden") == "1" || (string)c.Attribute("hidden") == "true";
                    for (int k = mi; k <= ma; k++) { larguras[k] = w; if (oc) colOcultas.Add(k); }
                }

            var cels = new Dictionary<long, XElement>();
            var alturas = new Dictionary<int, double>();
            var linOcultas = new HashSet<int>();
            int minR = int.MaxValue, maxR = 0, minC = int.MaxValue, maxC = 0;
            var sd = root.Element(S + "sheetData");
            bool cortado = false;
            if (sd != null)
            {
                int rAuto = 0;
                foreach (var row in sd.Elements(S + "row"))
                {
                    int r = row.Attribute("r") != null ? Int((string)row.Attribute("r")) : rAuto + 1;
                    rAuto = r;
                    if ((string)row.Attribute("hidden") == "1" || (string)row.Attribute("hidden") == "true") linOcultas.Add(r);
                    if (row.Attribute("ht") != null && (string)row.Attribute("customHeight") == "1")
                        alturas[r] = Ooxml.Num((string)row.Attribute("ht"), 15) * 4 / 3;
                    int cAuto = 0;
                    foreach (var c in row.Elements(S + "c"))
                    {
                        string rf = (string)c.Attribute("r");
                        int col = rf != null ? ColDeRef(rf) : cAuto + 1;
                        cAuto = col;
                        var v = c.Element(S + "v");
                        bool temValor = (v != null && v.Value.Length > 0) || c.Element(S + "is") != null;
                        if (!temValor && !Estilizado(Int((string)c.Attribute("s")))) continue;
                        if (col > MaxColunas) continue;
                        if (minR != int.MaxValue && r - minR >= MaxLinhas) { cortado = true; continue; }
                        cels[Chave(r, col)] = c;
                        minR = Math.Min(minR, r); maxR = Math.Max(maxR, r);
                        minC = Math.Min(minC, col); maxC = Math.Max(maxC, col);
                    }
                }
            }
            if (cels.Count == 0) return null;

            var spans = new Dictionary<long, int[]>();
            var cobertas = new HashSet<long>();
            var mcs = root.Element(S + "mergeCells");
            if (mcs != null)
                foreach (var mc in mcs.Elements(S + "mergeCell"))
                {
                    var partes = ((string)mc.Attribute("ref") ?? "").Split(':');
                    if (partes.Length != 2) continue;
                    int r1 = LinDeRef(partes[0]), c1 = ColDeRef(partes[0]), r2 = LinDeRef(partes[1]), c2 = ColDeRef(partes[1]);
                    if (r1 > maxR || c1 > maxC || r1 < minR || c1 < minC) continue;
                    r2 = Math.Min(r2, maxR); c2 = Math.Min(c2, maxC);
                    int rs = 0, cs = 0;
                    for (int r = r1; r <= r2; r++) if (!linOcultas.Contains(r)) rs++;
                    for (int c = c1; c <= c2; c++) if (!colOcultas.Contains(c)) cs++;
                    spans[Chave(r1, c1)] = new[] { Math.Max(1, rs), Math.Max(1, cs) };
                    for (int r = r1; r <= r2; r++)
                        for (int c = c1; c <= c2; c++)
                            if (r != r1 || c != c1) cobertas.Add(Chave(r, c));
                }

            var colunas = new List<int>();
            for (int c = minC; c <= maxC; c++) if (!colOcultas.Contains(c)) colunas.Add(c);
            Func<int, double> larg = c => { double w; return larguras.TryGetValue(c, out w) ? w : larguraPadrao; };
            double total = colunas.Sum(c => larg(c));

            var ps = root.Element(S + "pageSetup");
            bool paisagem = ps != null && (string)ps.Attribute("orientation") == "landscape";
            double[] papel = Layout.TamanhoPapel(o.Papel == 3 ? 0 : o.Papel);
            double utilRetrato = (papel[0] - 2 * 28.35) * 96 / 72;
            double utilPaisagem = (papel[1] - 2 * 28.35) * 96 / 72;
            if (o.Orientacao == 2 || (o.Orientacao == 0 && total > utilRetrato)) paisagem = true;
            if (o.Orientacao == 1) paisagem = false;
            double util = paisagem ? utilPaisagem : utilRetrato;
            double zoom = total > util ? util / total : 1;

            var sb = new StringBuilder("<div class=\"folha" + (paisagem ? " paisagem" : "") + "\">");
            if (titulo) sb.Append("<h2>" + Html.Esc(nome) + "</h2>");
            sb.Append("<table style=\"width:" + Ooxml.Px(total) + (zoom < 0.999 ? ";zoom:" + zoom.ToString("0.###", Ooxml.Inv) : "") + "\"><colgroup>");
            foreach (var c in colunas) sb.Append("<col style=\"width:" + Ooxml.Px(larg(c)) + "\">");
            sb.Append("</colgroup>");
            for (int r = minR; r <= maxR; r++)
            {
                if (linOcultas.Contains(r)) continue;
                double h;
                if (!alturas.TryGetValue(r, out h)) h = alturaPadrao;
                sb.Append("<tr style=\"height:" + Ooxml.Px(h) + "\">");
                for (int i = 0; i < colunas.Count; i++)
                {
                    int c = colunas[i];
                    long k = Chave(r, c);
                    if (cobertas.Contains(k)) continue;
                    XElement ce;
                    cels.TryGetValue(k, out ce);
                    int[] sp;
                    spans.TryGetValue(k, out sp);
                    var css = new StringBuilder();
                    string conteudo = "";
                    if (ce != null)
                    {
                        int s = Int((string)ce.Attribute("s"));
                        Xf xf = s >= 0 && s < xfs.Count ? xfs[s] : null;
                        bool numero;
                        conteudo = Valor(ce, xf, out numero);
                        if (xf != null)
                        {
                            if (xf.Fonte < cssFonte.Count) css.Append(cssFonte[xf.Fonte]);
                            if (xf.Fundo < cssFundo.Count && cssFundo[xf.Fundo] != null) css.Append("background:" + cssFundo[xf.Fundo] + ";");
                            if (xf.Borda < cssBorda.Count) css.Append(cssBorda[xf.Borda]);
                            if (xf.Quebra) css.Append("white-space:pre-wrap;");
                        }
                        else css.Append(cssFonte[0]);
                        string h2 = xf == null ? null : xf.H;
                        string al = h2 == "center" || h2 == "centerContinuous" ? "center" : h2 == "right" ? "right" : h2 == "justify" ? "justify" :
                                    h2 == "left" || h2 == "fill" ? "left" : (numero ? "right" : "left");
                        if (al != "left") css.Append("text-align:" + al + ";");
                        string v = xf == null ? null : xf.V;
                        if (v == "center") css.Append("vertical-align:middle;");
                        else if (v == "top") css.Append("vertical-align:top;");
                        // Texto longo "vaza" para a célula vizinha vazia, como no Excel
                        bool vizinhaVazia = i + 1 < colunas.Count && !cels.ContainsKey(Chave(r, colunas[i + 1])) && !cobertas.Contains(Chave(r, colunas[i + 1]));
                        if (!numero && vizinhaVazia && (xf == null || !xf.Quebra) && sp == null) css.Append("overflow:visible;");
                    }
                    sb.Append("<td" + (sp != null && sp[1] > 1 ? " colspan=\"" + sp[1] + "\"" : "") + (sp != null && sp[0] > 1 ? " rowspan=\"" + sp[0] + "\"" : "") +
                              (css.Length > 0 ? " style=\"" + css + "\"" : "") + ">" + Html.Esc(conteudo) + "</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
            if (cortado) sb.Append("<div class=\"aviso\">(mostrando as primeiras " + MaxLinhas + " linhas)</div>");
            sb.Append("</div>");
            return sb.ToString();
        }

        string Valor(XElement c, Xf xf, out bool numero)
        {
            numero = false;
            string t = (string)c.Attribute("t");
            if (t == "inlineStr") { var isx = c.Element(S + "is"); return isx == null ? "" : TextoRico(isx); }
            var v = c.Element(S + "v");
            if (v == null) return "";
            string val = v.Value;
            switch (t)
            {
                case "s":
                    int i;
                    return int.TryParse(val, out i) && i >= 0 && i < textos.Count ? textos[i] : "";
                case "str": return val;
                case "b": return val == "1" ? "VERDADEIRO" : "FALSO";
                case "e": return val;
            }
            double d;
            if (!double.TryParse(val, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return val;
            numero = true;
            int fmtId = xf == null ? 0 : xf.Fmt;
            string cod;
            if (!formatos.TryGetValue(fmtId, out cod)) cod = Embutido(fmtId);
            try { return FormatarNum(d, cod); } catch { return Geral(d); }
        }

        string Embutido(int id)
        {
            string moeda = "\"" + cult.NumberFormat.CurrencySymbol + " \"";
            switch (id)
            {
                case 1: return "0";
                case 2: return "0.00";
                case 3: case 41: return "#,##0";
                case 4: case 43: return "#,##0.00";
                case 5: case 6: case 42: return moeda + "#,##0";
                case 7: case 8: case 44: return moeda + "#,##0.00";
                case 9: return "0%";
                case 10: return "0.00%";
                case 11: return "0.00E+00";
                case 14: return "dd/mm/yyyy";
                case 15: return "d-mmm-yy";
                case 16: return "d-mmm";
                case 17: return "mmm-yy";
                case 18: return "h:mm AM/PM";
                case 19: return "h:mm:ss AM/PM";
                case 20: return "hh:mm";
                case 21: return "hh:mm:ss";
                case 22: return "dd/mm/yyyy hh:mm";
                case 37: case 38: return "#,##0;-#,##0";
                case 39: case 40: return "#,##0.00;-#,##0.00";
                case 45: return "mm:ss";
                case 46: return "[h]:mm:ss";
                case 47: return "mm:ss.0";
                case 48: return "##0.0E+0";
                case 49: return "@";
                default: return "General";
            }
        }

        string Geral(double v)
        {
            if (v == Math.Floor(v) && Math.Abs(v) < 1e15) return v.ToString("0", cult);
            double r = double.Parse(v.ToString("G10", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            if (Math.Abs(r) >= 1e11 || Math.Abs(r) < 1e-9) return r.ToString("0.#####E+00", cult);
            return r.ToString("0.##########", cult);
        }

        static List<string> Secoes(string cod)
        {
            var r = new List<string>();
            var sb = new StringBuilder();
            bool aspas = false;
            for (int i = 0; i < cod.Length; i++)
            {
                char c = cod[i];
                if (c == '"') aspas = !aspas;
                if (c == '\\' && i + 1 < cod.Length) { sb.Append(c).Append(cod[++i]); continue; }
                if (c == ';' && !aspas) { r.Add(sb.ToString()); sb.Clear(); continue; }
                sb.Append(c);
            }
            r.Add(sb.ToString());
            return r;
        }

        // Remove textos literais (entre aspas, escapados, _x e *x)
        static string SemLiterais(string s)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"') { int j = s.IndexOf('"', i + 1); i = j < 0 ? s.Length : j; continue; }
                if (c == '\\' || c == '_' || c == '*') { i++; continue; }
                sb.Append(c);
            }
            return sb.ToString();
        }

        static string Literal(string s)
        {
            var sb = new StringBuilder();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"') { int j = s.IndexOf('"', i + 1); if (j < 0) j = s.Length; sb.Append(s.Substring(i + 1, j - i - 1)); i = j; continue; }
                if (c == '\\') { if (i + 1 < s.Length) sb.Append(s[++i]); continue; }
                if (c == '_') { i++; sb.Append(' '); continue; }
                if (c == '*') { i++; continue; }
                if (c == '@') continue;
                sb.Append(c);
            }
            return sb.ToString();
        }

        static int Placeholder(string s, bool primeiro)
        {
            int achado = -1;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"') { int j = s.IndexOf('"', i + 1); i = j < 0 ? s.Length : j; continue; }
                if (c == '\\' || c == '_' || c == '*') { i++; continue; }
                if (c == '0' || c == '#' || c == '?')
                {
                    if (primeiro) return i;
                    achado = i;
                }
            }
            return achado;
        }

        string FormatarNum(double v, string cod)
        {
            if (cod == null || cod.Equals("General", StringComparison.OrdinalIgnoreCase) || cod == "@") return Geral(v);
            var secs = Secoes(cod);
            string sec = secs[0];
            double val = v;
            bool sinalProprio = false;
            if (v < 0 && secs.Count > 1 && secs[1].Length > 0) { sec = secs[1]; val = -v; sinalProprio = true; }
            else if (v == 0 && secs.Count > 2 && secs[2].Length > 0) sec = secs[2];

            if (Regex.IsMatch(sec, @"\[[hH]+\]"))
            {
                double tot = Math.Abs(val) * 24;
                int hh = (int)Math.Floor(tot);
                double resto = (tot - hh) * 60;
                int mm = (int)Math.Floor(resto), ss = (int)Math.Round((resto - mm) * 60);
                if (ss == 60) { ss = 0; mm++; }
                return (val < 0 ? "-" : "") + hh + ":" + mm.ToString("00") + (Regex.IsMatch(sec, @":ss") ? ":" + ss.ToString("00") : "");
            }
            sec = Regex.Replace(sec, @"\[\$([^\]\-]*)(-[^\]]*)?\]", m => m.Groups[1].Value.Length > 0 ? "\"" + m.Groups[1].Value + "\"" : "");
            sec = Regex.Replace(sec, @"\[[^\]]*\]", "");
            string limpo = SemLiterais(sec);
            if (limpo.Trim().Length == 0) return Literal(sec);
            if (Regex.IsMatch(limpo, "[dDmMyYhHsS]") && !limpo.Contains("E+") && !limpo.Contains("e+")) return FormatarData(val, sec);
            if (limpo.Contains("E+") || limpo.Contains("e+")) return val.ToString("0.00E+00", cult);

            if (limpo.Contains("%")) val *= 100;
            int ip = limpo.IndexOf('.');
            int zeros = 0, hashes = 0;
            if (ip >= 0)
                for (int i = ip + 1; i < limpo.Length; i++)
                {
                    if (limpo[i] == '0') zeros++;
                    else if (limpo[i] == '#' || limpo[i] == '?') hashes++;
                    else break;
                }
            bool milhar = Regex.IsMatch(limpo, "[#0?],[#0?]");
            string parteInt = ip >= 0 ? limpo.Substring(0, ip) : limpo;
            int zerosInt = parteInt.Count(ch => ch == '0');
            string f = milhar ? "#,##0" : (zerosInt > 1 ? new string('0', zerosInt) : "0");
            if (zeros + hashes > 0) f += "." + new string('0', zeros) + new string('#', hashes);
            bool negativo = val < 0 && !sinalProprio;
            string num = Math.Abs(val).ToString(f, cult);
            int a = Placeholder(sec, true), b = Placeholder(sec, false);
            string pre = a > 0 ? Literal(sec.Substring(0, a)) : "";
            string suf = b >= 0 && b < sec.Length - 1 ? Literal(sec.Substring(b + 1)) : "";
            suf = Regex.Replace(suf, @"^[.,]+", "");
            return (negativo ? "-" : "") + pre + num + suf;
        }

        static string Escapar(char c)
        {
            if (c == '/' || c == ':' || c == ' ' || c == '-' || c == ',' || c == '(' || c == ')') return c.ToString();
            return "\\" + c;
        }

        static bool ProximoEhSegundo(string s, int pos)
        {
            for (int i = pos; i < s.Length; i++)
            {
                char c = char.ToLowerInvariant(s[i]);
                if (c == 's') return true;
                if (char.IsLetter(c)) return false;
            }
            return false;
        }

        string FormatarData(double v, string sec)
        {
            DateTime dt;
            try { dt = DateTime.FromOADate(d1904 ? v + 1462 : v); } catch { return Geral(v); }
            var f = new StringBuilder();
            bool ampm = Regex.IsMatch(sec, "AM/PM|A/P", RegexOptions.IgnoreCase);
            string ultimo = "";
            int i = 0;
            while (i < sec.Length)
            {
                char c = sec[i];
                if (c == '"')
                {
                    int j = sec.IndexOf('"', i + 1);
                    if (j < 0) j = sec.Length;
                    foreach (char ch in sec.Substring(i + 1, j - i - 1)) f.Append(Escapar(ch));
                    i = j + 1;
                    continue;
                }
                if (c == '\\') { if (i + 1 < sec.Length) f.Append(Escapar(sec[i + 1])); i += 2; continue; }
                if (c == '_' || c == '*') { if (c == '_') f.Append(' '); i += 2; continue; }
                string resto = sec.Substring(i);
                if (resto.StartsWith("AM/PM", StringComparison.OrdinalIgnoreCase)) { f.Append("tt"); i += 5; continue; }
                if (resto.StartsWith("A/P", StringComparison.OrdinalIgnoreCase)) { f.Append("t"); i += 3; continue; }
                char lc = char.ToLowerInvariant(c);
                if ("ymdhs".IndexOf(lc) >= 0)
                {
                    int n = 1;
                    while (i + n < sec.Length && char.ToLowerInvariant(sec[i + n]) == lc) n++;
                    switch (lc)
                    {
                        case 'y': f.Append(n <= 2 ? "yy" : "yyyy"); break;
                        case 'd': f.Append(n == 1 ? "d" : n == 2 ? "dd" : n == 3 ? "ddd" : "dddd"); break;
                        case 'h': f.Append(ampm ? (n == 1 ? "h" : "hh") : (n == 1 ? "H" : "HH")); break;
                        case 's': f.Append(n == 1 ? "s" : "ss"); break;
                        case 'm':
                            bool minuto = n <= 2 && (ultimo == "h" || ProximoEhSegundo(sec, i + n));
                            if (minuto) f.Append(n == 1 ? "m" : "mm");
                            else f.Append(n == 1 ? "M" : n == 2 ? "MM" : n == 3 ? "MMM" : "MMMM");
                            break;
                    }
                    ultimo = lc.ToString();
                    i += n;
                    continue;
                }
                if (c == '.' && ultimo == "s")
                {
                    int n = 0;
                    while (i + 1 + n < sec.Length && sec[i + 1 + n] == '0') n++;
                    f.Append("." + new string('f', Math.Max(1, n)));
                    i += 1 + n;
                    continue;
                }
                f.Append(Escapar(c));
                i++;
            }
            string fs = f.ToString();
            if (fs.Length == 1) fs = "%" + fs;
            try { return dt.ToString(fs, cult); } catch { return dt.ToString(cult); }
        }
    }

    // ===================================================================== PPTX
    class PptxConv
    {
        static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
        static readonly XNamespace A = Ooxml.A, R = Ooxml.R, MC = Ooxml.MC;

        ZipArchive z;
        readonly Dictionary<string, Parte> partes = new Dictionary<string, Parte>();
        readonly Dictionary<string, string> cacheImg = new Dictionary<string, string>();

        class Parte { public string Caminho; public XElement Raiz; public Dictionary<string, Rel> Rels; }

        class Ctx
        {
            public Parte Slide, Layout, Master;
            public Dictionary<string, string> Tema = new Dictionary<string, string>();
            public Dictionary<string, string> Mapa = new Dictionary<string, string>();
            public string FonteMaior = "Calibri Light", FonteMenor = "Calibri";
        }

        struct Tf { public double A, B, C, D; }

        class Traco { public string Cor; public double Larg; public bool Tracejado; }

        static string Pt(double v) { return Ooxml.Pt(v); }
        static double Emu(string s, double padrao) { return Ooxml.Num(s, padrao); }
        static string At(XElement e, string nome) { return e == null ? null : (string)e.Attribute(nome); }

        Parte Carregar(string caminho)
        {
            if (caminho == null) return null;
            Parte p;
            if (partes.TryGetValue(caminho, out p)) return p;
            var x = Ooxml.Xml(z, caminho);
            p = x == null ? null : new Parte { Caminho = caminho, Raiz = x.Root, Rels = Ooxml.Rels(z, caminho) };
            partes[caminho] = p;
            return p;
        }

        Parte Relacionada(Parte p, string tipo) { return p == null ? null : Carregar(Ooxml.Relacionado(p.Rels, tipo)); }

        public string Converter(string caminho)
        {
            using (z = ZipFile.OpenRead(caminho))
            {
                string presPath = Ooxml.DocumentoPrincipal(z) ?? "ppt/presentation.xml";
                var pres = Carregar(presPath);
                if (pres == null) throw new Exception("Apresentação inválida.");
                var sz = pres.Raiz.Element(P + "sldSz");
                double w = Emu(At(sz, "cx"), 9144000) / 12700, h = Emu(At(sz, "cy"), 6858000) / 12700;
                var sb = new StringBuilder();
                var lista = pres.Raiz.Element(P + "sldIdLst");
                if (lista != null)
                    foreach (var sid in lista.Elements(P + "sldId"))
                    {
                        Rel rl;
                        string rid = (string)sid.Attribute(R + "id");
                        if (rid == null || !pres.Rels.TryGetValue(rid, out rl)) continue;
                        var slide = Carregar(rl.Alvo);
                        if (slide == null || At(slide.Raiz, "show") == "0") continue;
                        sb.Append(Slide(slide));
                    }
                string css = "@page { size: " + Pt(w) + " " + Pt(h) + "; margin: 0 }" +
                    "body{margin:0;font-family:Calibri,Arial,sans-serif}" +
                    ".slide{position:relative;overflow:hidden;width:" + Pt(w) + ";height:" + Pt(h) + "}.slide + .slide{break-before:page}" +
                    ".forma{position:absolute;box-sizing:border-box}" +
                    ".texto{position:absolute;inset:0;box-sizing:border-box;display:flex;flex-direction:column;overflow:visible}" +
                    ".texto p{margin:0;white-space:pre-wrap;overflow-wrap:break-word}" +
                    "table.forma{border-collapse:collapse;table-layout:fixed}table.forma td{overflow:hidden;vertical-align:top}";
                return Html.Pagina(Path.GetFileName(caminho), css, sb.ToString());
            }
        }

        string Slide(Parte slide)
        {
            var ctx = new Ctx { Slide = slide };
            ctx.Layout = Relacionada(slide, "slideLayout");
            ctx.Master = Relacionada(ctx.Layout, "slideMaster");
            if (ctx.Master != null)
            {
                CarregarTema(Relacionada(ctx.Master, "theme"), ctx);
                var cm = ctx.Master.Raiz.Element(P + "clrMap");
                if (cm != null) foreach (var a in cm.Attributes()) ctx.Mapa[a.Name.LocalName] = a.Value;
            }
            var sb = new StringBuilder();
            sb.Append("<div class=\"slide\" style=\"background:" + Fundo(ctx) + "\">");
            var id = new Tf { A = 1, C = 1 };
            bool mostrarMestre = At(slide.Raiz, "showMasterSp") != "0" && (ctx.Layout == null || At(ctx.Layout.Raiz, "showMasterSp") != "0");
            if (ctx.Master != null && mostrarMestre) Arvore(ArvoreDe(ctx.Master), ctx.Master, ctx, id, true, sb);
            if (ctx.Layout != null) Arvore(ArvoreDe(ctx.Layout), ctx.Layout, ctx, id, true, sb);
            Arvore(ArvoreDe(slide), slide, ctx, id, false, sb);
            sb.Append("</div>");
            return sb.ToString();
        }

        static XElement ArvoreDe(Parte p)
        {
            var c = p.Raiz.Element(P + "cSld");
            return c == null ? null : c.Element(P + "spTree");
        }

        void CarregarTema(Parte tema, Ctx ctx)
        {
            if (tema == null) return;
            var cs = tema.Raiz.Descendants(A + "clrScheme").FirstOrDefault();
            if (cs != null)
                foreach (var c in cs.Elements())
                {
                    var s = c.Element(A + "srgbClr");
                    var y = c.Element(A + "sysClr");
                    ctx.Tema[c.Name.LocalName] = s != null ? (string)s.Attribute("val") : (y != null ? ((string)y.Attribute("lastClr") ?? "000000") : "000000");
                }
            var fs = tema.Raiz.Descendants(A + "fontScheme").FirstOrDefault();
            if (fs != null)
            {
                var mj = fs.Element(A + "majorFont"); var mn = fs.Element(A + "minorFont");
                if (mj != null && mj.Element(A + "latin") != null) ctx.FonteMaior = At(mj.Element(A + "latin"), "typeface") ?? ctx.FonteMaior;
                if (mn != null && mn.Element(A + "latin") != null) ctx.FonteMenor = At(mn.Element(A + "latin"), "typeface") ?? ctx.FonteMenor;
            }
        }

        string Cor(XElement pai, Ctx ctx)
        {
            if (pai == null) return null;
            var c = pai.Elements().FirstOrDefault(x => x.Name.Namespace == A && x.Name.LocalName.EndsWith("Clr"));
            if (c == null) return null;
            string hex = null;
            switch (c.Name.LocalName)
            {
                case "srgbClr": hex = At(c, "val"); break;
                case "schemeClr": hex = Esquema(At(c, "val"), ctx); break;
                case "sysClr": hex = At(c, "lastClr") ?? (At(c, "val") == "window" ? "FFFFFF" : "000000"); break;
                case "prstClr": hex = Predefinida(At(c, "val")); break;
                case "scrgbClr":
                    hex = Cores.Hex(Emu(At(c, "r"), 0) / 100000, Emu(At(c, "g"), 0) / 100000, Emu(At(c, "b"), 0) / 100000).Substring(1);
                    break;
            }
            if (hex == null) return null;
            return Cores.Aplicar(hex, c.Elements());
        }

        static string Esquema(string val, Ctx ctx)
        {
            if (val == null || val == "phClr") return null;
            string m;
            if (ctx.Mapa.TryGetValue(val, out m)) val = m;
            else if (val == "bg1") val = "lt1";
            else if (val == "tx1") val = "dk1";
            else if (val == "bg2") val = "lt2";
            else if (val == "tx2") val = "dk2";
            string hex;
            return ctx.Tema.TryGetValue(val, out hex) ? hex : null;
        }

        static string Predefinida(string v)
        {
            switch (v)
            {
                case "black": return "000000";
                case "white": return "FFFFFF";
                case "red": return "FF0000";
                case "green": return "008000";
                case "blue": return "0000FF";
                case "yellow": return "FFFF00";
                case "gray": case "grey": return "808080";
                case "orange": return "FFA500";
                default: return "000000";
            }
        }

        string Fundo(Ctx ctx)
        {
            foreach (var p in new[] { ctx.Slide, ctx.Layout, ctx.Master })
            {
                if (p == null) continue;
                var cs = p.Raiz.Element(P + "cSld");
                var bg = cs == null ? null : cs.Element(P + "bg");
                if (bg == null) continue;
                var bgPr = bg.Element(P + "bgPr");
                if (bgPr != null)
                {
                    var f = Preenchimento(new[] { bgPr }, null, p, ctx);
                    if (f != null) return f;
                }
                var c = Cor(bg.Element(P + "bgRef"), ctx);
                if (c != null) return c;
            }
            return "#fff";
        }

        string ImagemRel(Parte parte, string id)
        {
            Rel rl;
            if (id == null || !parte.Rels.TryGetValue(id, out rl) || rl.Externo) return null;
            return Ooxml.DataUri(z, rl.Alvo, cacheImg);
        }

        string Preenchimento(IEnumerable<XElement> spPrs, XElement estilo, Parte parte, Ctx ctx)
        {
            foreach (var sp in spPrs)
            {
                if (sp == null) continue;
                if (sp.Element(A + "noFill") != null) return null;
                var sf = sp.Element(A + "solidFill");
                if (sf != null) return Cor(sf, ctx);
                var gf = sp.Element(A + "gradFill");
                if (gf != null)
                {
                    var stops = gf.Descendants(A + "gs").Select(g => new { Pos = Emu(At(g, "pos"), 0) / 1000, Cor = Cor(g, ctx) }).Where(g => g.Cor != null).ToList();
                    if (stops.Count == 0) return null;
                    if (stops.Count == 1) return stops[0].Cor;
                    var lin = gf.Element(A + "lin");
                    double ang = lin == null ? 90 : Emu(At(lin, "ang"), 5400000) / 60000 + 90;
                    return "linear-gradient(" + ang.ToString("0", Ooxml.Inv) + "deg," +
                           string.Join(",", stops.Select(s => s.Cor + " " + s.Pos.ToString("0.#", Ooxml.Inv) + "%")) + ")";
                }
                var bf = sp.Element(A + "blipFill");
                if (bf != null)
                {
                    var blip = bf.Element(A + "blip");
                    string uri = blip == null ? null : ImagemRel(parte, (string)blip.Attribute(R + "embed"));
                    return uri == null ? null : "url(" + uri + ") center/100% 100% no-repeat";
                }
                var pf = sp.Element(A + "pattFill");
                if (pf != null) return Cor(pf.Element(A + "fgClr"), ctx);
                if (sp.Element(A + "grpFill") != null) return null;
            }
            if (estilo != null)
            {
                var fr = estilo.Element(A + "fillRef");
                if (fr != null && At(fr, "idx") != "0") return Cor(fr, ctx);
            }
            return null;
        }

        Traco Linha(IEnumerable<XElement> spPrs, XElement estilo, Ctx ctx)
        {
            double? larg = null;
            foreach (var sp in spPrs)
            {
                if (sp == null) continue;
                var ln = sp.Element(A + "ln");
                if (ln == null) continue;
                if (ln.Element(A + "noFill") != null) return null;
                if (ln.Attribute("w") != null) larg = Emu(At(ln, "w"), 12700) / 12700;
                var dash = ln.Element(A + "prstDash");
                bool tr = dash != null && At(dash, "val") != "solid";
                var sf = ln.Element(A + "solidFill");
                if (sf != null) return new Traco { Cor = Cor(sf, ctx) ?? "#000", Larg = larg ?? 0.75, Tracejado = tr };
                break;
            }
            if (estilo != null)
            {
                var lr = estilo.Element(A + "lnRef");
                if (lr != null && At(lr, "idx") != "0")
                {
                    string c = Cor(lr, ctx);
                    if (c != null) return new Traco { Cor = c, Larg = larg ?? 0.75 };
                }
            }
            return null;
        }

        static XElement Ph(XElement e)
        {
            var nv = e.Elements().FirstOrDefault(x => x.Name.LocalName.StartsWith("nv"));
            var nvPr = nv == null ? null : nv.Element(P + "nvPr");
            return nvPr == null ? null : nvPr.Element(P + "ph");
        }

        static string TipoPh(XElement ph)
        {
            string t = At(ph, "type") ?? "obj";
            if (t == "ctrTitle") return "title";
            if (t == "subTitle" || t == "obj" || t == "pic" || t == "chart" || t == "tbl" || t == "media" || t == "clipArt" || t == "dgm") return "body";
            return t;
        }

        static XElement Correspondente(XElement ph, Parte alvo)
        {
            if (ph == null || alvo == null) return null;
            var arv = ArvoreDe(alvo);
            if (arv == null) return null;
            var cands = arv.Descendants().Where(x => (x.Name == P + "sp" || x.Name == P + "pic") && Ph(x) != null).ToList();
            string idx = At(ph, "idx");
            if (idx != null)
            {
                var m = cands.FirstOrDefault(x => At(Ph(x), "idx") == idx);
                if (m != null) return m;
            }
            string tipo = TipoPh(ph);
            return cands.FirstOrDefault(x => TipoPh(Ph(x)) == tipo);
        }

        static XElement SpPr(XElement e)
        {
            if (e == null) return null;
            return e.Element(P + "spPr") ?? e.Element(P + "grpSpPr");
        }

        void Arvore(XElement arvore, Parte parte, Ctx ctx, Tf t, bool soDecoracao, StringBuilder sb)
        {
            if (arvore == null) return;
            foreach (var e in arvore.Elements())
            {
                switch (e.Name.LocalName)
                {
                    case "sp":
                    case "cxnSp":
                        if (soDecoracao && Ph(e) != null) break;
                        Forma(e, parte, ctx, t, sb);
                        break;
                    case "pic":
                        if (soDecoracao && Ph(e) != null) break;
                        Figura(e, parte, ctx, t, sb);
                        break;
                    case "grpSp": Grupo(e, parte, ctx, t, soDecoracao, sb); break;
                    case "graphicFrame": Quadro(e, parte, ctx, t, sb); break;
                    case "AlternateContent":
                        var esc = Ooxml.Escolha(e, true);
                        if (esc != null) Arvore(esc, parte, ctx, t, soDecoracao, sb);
                        break;
                }
            }
        }

        void Herdados(XElement e, Parte parte, Ctx ctx, out XElement lph, out XElement mph)
        {
            lph = null; mph = null;
            var ph = Ph(e);
            if (ph == null) return;
            if (parte == ctx.Slide)
            {
                lph = Correspondente(ph, ctx.Layout);
                mph = Correspondente(lph != null ? Ph(lph) : ph, ctx.Master);
            }
            else if (parte == ctx.Layout) mph = Correspondente(ph, ctx.Master);
        }

        static XElement Xfrm(IEnumerable<XElement> heranca)
        {
            foreach (var x in heranca)
            {
                var sp = SpPr(x);
                var xf = sp == null ? null : sp.Element(A + "xfrm");
                if (xf != null && xf.Element(A + "off") != null) return xf;
            }
            return null;
        }

        static void Caixa(XElement xfrm, Tf t, out double x, out double y, out double w, out double h)
        {
            var off = xfrm.Element(A + "off");
            var ext = xfrm.Element(A + "ext");
            x = (t.A * Emu(At(off, "x"), 0) + t.B) / 12700;
            y = (t.C * Emu(At(off, "y"), 0) + t.D) / 12700;
            w = t.A * Emu(At(ext, "cx"), 0) / 12700;
            h = t.C * Emu(At(ext, "cy"), 0) / 12700;
        }

        static string Transformacao(XElement xfrm, bool espelhar)
        {
            double rot = Emu(At(xfrm, "rot"), 0) / 60000;
            var partes = new List<string>();
            if (rot != 0) partes.Add("rotate(" + rot.ToString("0.##", Ooxml.Inv) + "deg)");
            if (espelhar && At(xfrm, "flipH") == "1") partes.Add("scaleX(-1)");
            if (espelhar && At(xfrm, "flipV") == "1") partes.Add("scaleY(-1)");
            return partes.Count == 0 ? "" : "transform:" + string.Join(" ", partes) + ";";
        }

        static string FormaCss(string geo, double w, double h)
        {
            switch (geo)
            {
                case "ellipse": case "flowChartConnector": return "border-radius:50%;";
                case "roundRect": case "flowChartAlternateProcess": case "round2SameRect": return "border-radius:" + Pt(Math.Min(w, h) * 0.1667) + ";";
                case "triangle": return Clip("50% 0,100% 100%,0 100%");
                case "rtTriangle": return Clip("0 0,100% 100%,0 100%");
                case "diamond": case "flowChartDecision": return Clip("50% 0,100% 50%,50% 100%,0 50%");
                case "parallelogram": return Clip("25% 0,100% 0,75% 100%,0 100%");
                case "trapezoid": return Clip("25% 0,75% 0,100% 100%,0 100%");
                case "pentagon": case "homePlate": return Clip("0 0,80% 0,100% 50%,80% 100%,0 100%");
                case "chevron": return Clip("0 0,80% 0,100% 50%,80% 100%,0 100%,20% 50%");
                case "hexagon": return Clip("25% 0,75% 0,100% 50%,75% 100%,25% 100%,0 50%");
                case "octagon": return Clip("30% 0,70% 0,100% 30%,100% 70%,70% 100%,30% 100%,0 70%,0 30%");
                case "rightArrow": return Clip("0 30%,65% 30%,65% 0,100% 50%,65% 100%,65% 70%,0 70%");
                case "leftArrow": return Clip("100% 30%,35% 30%,35% 0,0 50%,35% 100%,35% 70%,100% 70%");
                case "upArrow": return Clip("30% 100%,30% 35%,0 35%,50% 0,100% 35%,70% 35%,70% 100%");
                case "downArrow": return Clip("30% 0,70% 0,70% 65%,100% 65%,50% 100%,0 65%,30% 65%");
                case "star5": return Clip("50% 0,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%");
                case "plus": case "mathPlus": return Clip("35% 0,65% 0,65% 35%,100% 35%,100% 65%,65% 65%,65% 100%,35% 100%,35% 65%,0 65%,0 35%,35% 35%");
                default: return "";
            }
        }

        static string Clip(string pontos) { return "clip-path:polygon(" + pontos + ");"; }

        void Forma(XElement e, Parte parte, Ctx ctx, Tf t, StringBuilder sb)
        {
            XElement lph, mph;
            Herdados(e, parte, ctx, out lph, out mph);
            var heranca = new[] { e, lph, mph }.Where(q => q != null).ToList();
            var xfrm = Xfrm(heranca);
            if (xfrm == null) return;
            double x, y, w, h;
            Caixa(xfrm, t, out x, out y, out w, out h);
            var spPrs = heranca.Select(SpPr).ToList();
            var geoEl = spPrs.Where(s => s != null).Select(s => s.Element(A + "prstGeom")).FirstOrDefault(g => g != null);
            string geo = At(geoEl, "prst") ?? "rect";
            var estilo = e.Element(P + "style");
            var traco = Linha(spPrs, estilo, ctx);

            if (geo == "line" || geo.Contains("Connector") || e.Name.LocalName == "cxnSp")
            {
                if (traco == null) return;
                bool fh = At(xfrm, "flipH") == "1", fv = At(xfrm, "flipV") == "1";
                double ww = Math.Max(w, 0.01), hh = Math.Max(h, 0.01);
                sb.Append("<svg class=\"forma\" style=\"left:" + Pt(x) + ";top:" + Pt(y) + ";width:" + Pt(ww) + ";height:" + Pt(hh) + ";overflow:visible\" viewBox=\"0 0 " +
                          N(ww) + " " + N(hh) + "\" preserveAspectRatio=\"none\"><line x1=\"" + N(fh ? ww : 0) + "\" y1=\"" + N(fv ? hh : 0) + "\" x2=\"" + N(fh ? 0 : ww) +
                          "\" y2=\"" + N(fv ? 0 : hh) + "\" stroke=\"" + traco.Cor + "\" stroke-width=\"" + N(traco.Larg) + "\"" + (traco.Tracejado ? " stroke-dasharray=\"4 3\"" : "") + "/></svg>");
                return;
            }

            string fundo = Preenchimento(spPrs, estilo, parte, ctx);
            var css = new StringBuilder("left:" + Pt(x) + ";top:" + Pt(y) + ";width:" + Pt(w) + ";height:" + Pt(h) + ";");
            if (fundo != null) css.Append("background:" + fundo + ";");
            if (traco != null) css.Append("border:" + Pt(traco.Larg) + (traco.Tracejado ? " dashed " : " solid ") + traco.Cor + ";");
            css.Append(FormaCss(geo, w, h));
            css.Append(Transformacao(xfrm, false));
            sb.Append("<div class=\"forma\" style=\"" + css + "\">");
            var txBody = e.Element(P + "txBody");
            if (txBody != null) sb.Append(Texto(txBody, e, lph, mph, parte, ctx));
            sb.Append("</div>");
        }

        static string N(double v) { return v.ToString("0.##", Ooxml.Inv); }

        XElement EstiloMestre(XElement shape, Parte parte, Ctx ctx)
        {
            if (ctx.Master == null) return null;
            var ts = ctx.Master.Raiz.Element(P + "txStyles");
            if (ts == null) return null;
            var ph = Ph(shape);
            if (ph == null) return ts.Element(P + "otherStyle");
            return TipoPh(ph) == "title" ? ts.Element(P + "titleStyle") : ts.Element(P + "bodyStyle");
        }

        string Texto(XElement txBody, XElement shape, XElement lph, XElement mph, Parte parte, Ctx ctx)
        {
            var listas = new List<XElement> { txBody.Element(A + "lstStyle") };
            foreach (var x in new[] { lph, mph })
            {
                var tb = x == null ? null : x.Element(P + "txBody");
                if (tb != null) listas.Add(tb.Element(A + "lstStyle"));
            }
            listas.Add(EstiloMestre(shape, parte, ctx));
            listas = listas.Where(l => l != null).ToList();

            var bps = new List<XElement> { txBody.Element(A + "bodyPr") };
            foreach (var x in new[] { lph, mph })
            {
                var tb = x == null ? null : x.Element(P + "txBody");
                if (tb != null) bps.Add(tb.Element(A + "bodyPr"));
            }
            bps = bps.Where(b => b != null).ToList();
            Func<string, string> bp = n => bps.Select(b => At(b, n)).FirstOrDefault(v => v != null);

            string anchor = bp("anchor") ?? "t";
            double li = Emu(bp("lIns"), 91440) / 12700, ti = Emu(bp("tIns"), 45720) / 12700, ri = Emu(bp("rIns"), 91440) / 12700, bi = Emu(bp("bIns"), 45720) / 12700;
            double escala = 1, reducao = 0;
            var na = bps.Count > 0 ? bps[0].Element(A + "normAutofit") : null;
            if (na != null)
            {
                escala = Emu(At(na, "fontScale"), 100000) / 100000;
                reducao = Emu(At(na, "lnSpcReduction"), 0) / 100000;
            }
            string vert = bp("vert");
            string just = anchor == "ctr" ? "center" : (anchor == "b" ? "flex-end" : "flex-start");
            var sb = new StringBuilder("<div class=\"texto\" style=\"padding:" + Pt(ti) + " " + Pt(ri) + " " + Pt(bi) + " " + Pt(li) + ";justify-content:" + just +
                (bp("wrap") == "none" ? ";white-space:nowrap" : "") +
                (vert == "vert" || vert == "eaVert" ? ";writing-mode:vertical-rl" : (vert == "vert270" ? ";writing-mode:vertical-rl;transform:rotate(180deg)" : "")) + "\">");
            var cont = new int[9];
            foreach (var p in txBody.Elements(A + "p")) sb.Append(Paragrafo(p, listas, ctx, escala, reducao, cont));
            sb.Append("</div>");
            return sb.ToString();
        }

        string Paragrafo(XElement p, List<XElement> listas, Ctx ctx, double escala, double reducao, int[] cont)
        {
            var pPr = p.Element(A + "pPr");
            int lvl = Math.Max(0, Math.Min(8, (int)Emu(At(pPr, "lvl"), 0)));
            var niveis = new List<XElement>();
            if (pPr != null) niveis.Add(pPr);
            foreach (var l in listas)
            {
                var n = l.Element(A + ("lvl" + (lvl + 1) + "pPr"));
                if (n != null) niveis.Add(n);
            }
            Func<string, string> at = nome => niveis.Select(n => At(n, nome)).FirstOrDefault(v => v != null);
            Func<string, XElement> el = nome => niveis.Select(n => n.Element(A + nome)).FirstOrDefault(v => v != null);
            var defs = niveis.Select(n => n.Element(A + "defRPr")).Where(d => d != null).ToList();

            string algn = at("algn") ?? "l";
            double marL = Emu(at("marL"), 0) / 12700, indent = Emu(at("indent"), 0) / 12700;
            var runs = new StringBuilder();
            double primeiroTam = 0;
            foreach (var c in p.Elements())
            {
                if (c.Name == A + "r" || c.Name == A + "fld")
                {
                    double tam;
                    runs.Append(Run(c, defs, ctx, escala, out tam));
                    if (primeiroTam == 0) primeiroTam = tam;
                }
                else if (c.Name == A + "br") runs.Append("<br>");
            }
            if (primeiroTam == 0)
            {
                var end = p.Element(A + "endParaRPr");
                var cadeia = new List<XElement>();
                if (end != null) cadeia.Add(end);
                cadeia.AddRange(defs);
                string sz = cadeia.Select(x => At(x, "sz")).FirstOrDefault(v => v != null);
                primeiroTam = Emu(sz, 1800) / 100 * escala;
            }
            bool vazio = runs.Length == 0;

            double linha = 1.0;
            var lnSpc = el("lnSpc");
            if (lnSpc != null)
            {
                var pct = lnSpc.Element(A + "spcPct");
                var pts = lnSpc.Element(A + "spcPts");
                if (pct != null) linha = Emu(At(pct, "val"), 100000) / 100000;
                else if (pts != null) linha = Emu(At(pts, "val"), 1200) / 100 / Math.Max(1, primeiroTam) / 1.2;
            }
            linha *= (1 - reducao);
            Func<XElement, double> espaco = e =>
            {
                if (e == null) return 0;
                var pts = e.Element(A + "spcPts");
                if (pts != null) return Emu(At(pts, "val"), 0) / 100;
                var pct = e.Element(A + "spcPct");
                if (pct != null) return Emu(At(pct, "val"), 0) / 100000 * primeiroTam * 1.2;
                return 0;
            };

            string marcador = null;
            if (!vazio)
                foreach (var n in niveis)
                {
                    if (n.Element(A + "buNone") != null) break;
                    var bc = n.Element(A + "buChar");
                    if (bc != null)
                    {
                        var bf = niveis.Select(x => x.Element(A + "buFont")).FirstOrDefault(x => x != null);
                        marcador = Html.Esc(MapearBullet(At(bc, "char") ?? "•", At(bf, "typeface")));
                        break;
                    }
                    var ba = n.Element(A + "buAutoNum");
                    if (ba != null)
                    {
                        cont[lvl]++;
                        for (int k = lvl + 1; k < 9; k++) cont[k] = 0;
                        marcador = AutoNum(At(ba, "type") ?? "arabicPeriod", cont[lvl] + (int)Emu(At(ba, "startAt"), 1) - 1);
                        break;
                    }
                }

            var css = new StringBuilder();
            css.Append("margin:" + Pt(espaco(el("spcBef"))) + " 0 " + Pt(espaco(el("spcAft"))) + " 0;");
            if (marL != 0) css.Append("padding-left:" + Pt(marL) + ";");
            if (indent != 0) css.Append("text-indent:" + Pt(indent) + ";");
            css.Append("text-align:" + (algn == "ctr" ? "center" : algn == "r" ? "right" : (algn == "just" || algn == "dist") ? "justify" : "left") + ";");
            css.Append("line-height:" + (linha * 1.2).ToString("0.###", Ooxml.Inv) + ";font-size:" + Pt(primeiroTam) + ";");
            string bullet = "";
            if (marcador != null)
            {
                string corB = Cor(niveis.Select(x => x.Element(A + "buClr")).FirstOrDefault(x => x != null), ctx);
                bullet = "<span style=\"display:inline-block;text-indent:0;min-width:" + Pt(indent < 0 ? -indent : primeiroTam * 0.8) +
                         (corB != null ? ";color:" + corB : "") + "\">" + marcador + "</span>";
            }
            return "<p style=\"" + css + "\">" + bullet + (vazio ? "&nbsp;" : runs.ToString()) + "</p>";
        }

        string Run(XElement r, List<XElement> defs, Ctx ctx, double escala, out double tamanho)
        {
            var rPr = r.Element(A + "rPr");
            var cadeia = new List<XElement>();
            if (rPr != null) cadeia.Add(rPr);
            cadeia.AddRange(defs);
            Func<string, string> at = n => cadeia.Select(x => At(x, n)).FirstOrDefault(v => v != null);
            tamanho = Emu(at("sz"), 1800) / 100 * escala;
            string cor = null;
            foreach (var x in cadeia)
            {
                if (x.Element(A + "noFill") != null) { cor = "transparent"; break; }
                var sf = x.Element(A + "solidFill");
                if (sf != null) { cor = Cor(sf, ctx); if (cor != null) break; }
            }
            if (cor == null) cor = "#" + (Esquema("tx1", ctx) ?? "000000");
            string fonte = null;
            foreach (var x in cadeia)
            {
                var lat = x.Element(A + "latin");
                if (lat == null) continue;
                fonte = At(lat, "typeface");
                if (fonte == "+mj-lt") fonte = ctx.FonteMaior;
                else if (fonte == "+mn-lt") fonte = ctx.FonteMenor;
                if (fonte != null) break;
            }
            if (fonte == null) fonte = ctx.FonteMenor;
            var css = new StringBuilder("font-size:" + Pt(tamanho) + ";color:" + cor + ";font-family:" + Ooxml.Fonte(fonte) + ";");
            string b = at("b"), i = at("i"), u = at("u"), st = at("strike"), bl = at("baseline"), cap = at("cap");
            if (b == "1" || b == "true") css.Append("font-weight:bold;");
            if (i == "1" || i == "true") css.Append("font-style:italic;");
            var dec = new List<string>();
            if (u != null && u != "none") dec.Add("underline");
            if (st != null && st != "noStrike") dec.Add("line-through");
            if (dec.Count > 0) css.Append("text-decoration:" + string.Join(" ", dec) + ";");
            if (bl != null && bl != "0") css.Append("vertical-align:" + (Emu(bl, 0) > 0 ? "super" : "sub") + ";font-size:" + Pt(tamanho * 0.65) + ";");
            if (cap == "all") css.Append("text-transform:uppercase;");
            else if (cap == "small") css.Append("font-variant:small-caps;");
            var t = r.Element(A + "t");
            return "<span style=\"" + css + "\">" + Html.Esc(t == null ? "" : t.Value) + "</span>";
        }

        static string MapearBullet(string c, string fonte)
        {
            if (fonte != null && fonte.StartsWith("Wingdings", StringComparison.OrdinalIgnoreCase))
            {
                switch (c)
                {
                    case "§": return "▪";
                    case "Ø": return "➢";
                    case "ü": return "✓";
                    case "q": return "❑";
                    case "v": return "❖";
                    case "n": return "■";
                    case "l": return "●";
                    case "Ü": return "➤";
                    default: return "•";
                }
            }
            if (c.Length > 0 && c[0] >= '' && c[0] <= '') return "•";
            return c;
        }

        static string AutoNum(string tipo, int n)
        {
            string v;
            if (tipo.StartsWith("alphaLc")) v = ((char)('a' + (n - 1) % 26)).ToString();
            else if (tipo.StartsWith("alphaUc")) v = ((char)('A' + (n - 1) % 26)).ToString();
            else if (tipo.StartsWith("romanLc")) v = DocxConv.Romano(n).ToLowerInvariant();
            else if (tipo.StartsWith("romanUc")) v = DocxConv.Romano(n);
            else v = n.ToString();
            if (tipo.EndsWith("ParenBoth")) return "(" + v + ")";
            if (tipo.EndsWith("ParenR")) return v + ")";
            if (tipo.EndsWith("Plain")) return v;
            return v + ".";
        }

        void Figura(XElement e, Parte parte, Ctx ctx, Tf t, StringBuilder sb)
        {
            XElement lph, mph;
            Herdados(e, parte, ctx, out lph, out mph);
            var heranca = new[] { e, lph, mph }.Where(q => q != null).ToList();
            var xfrm = Xfrm(heranca);
            if (xfrm == null) return;
            var bf = e.Element(P + "blipFill");
            var blip = bf == null ? null : bf.Element(A + "blip");
            string uri = blip == null ? null : ImagemRel(parte, (string)blip.Attribute(R + "embed"));
            if (uri == null) return;
            double x, y, w, h;
            Caixa(xfrm, t, out x, out y, out w, out h);
            var traco = Linha(new[] { SpPr(e) }, null, ctx);
            string borda = traco == null ? "" : "border:" + Pt(traco.Larg) + " solid " + traco.Cor + ";";
            string geo = At(SpPr(e) == null ? null : SpPr(e).Element(A + "prstGeom"), "prst") ?? "rect";
            string forma = FormaCss(geo, w, h);
            var src = bf.Element(A + "srcRect");
            double l = Emu(At(src, "l"), 0) / 100000, tp = Emu(At(src, "t"), 0) / 100000, rr = Emu(At(src, "r"), 0) / 100000, bb = Emu(At(src, "b"), 0) / 100000;
            string caixa = "left:" + Pt(x) + ";top:" + Pt(y) + ";width:" + Pt(w) + ";height:" + Pt(h) + ";";
            if (l == 0 && tp == 0 && rr == 0 && bb == 0)
            {
                sb.Append("<img class=\"forma\" src=\"" + uri + "\" style=\"" + caixa + borda + forma + Transformacao(xfrm, true) + "\">");
                return;
            }
            double iw = w / Math.Max(0.01, 1 - l - rr), ih = h / Math.Max(0.01, 1 - tp - bb);
            sb.Append("<div class=\"forma\" style=\"" + caixa + "overflow:hidden;" + borda + forma + Transformacao(xfrm, true) + "\"><img src=\"" + uri +
                      "\" style=\"position:absolute;left:" + Pt(-l * iw) + ";top:" + Pt(-tp * ih) + ";width:" + Pt(iw) + ";height:" + Pt(ih) + "\"></div>");
        }

        void Grupo(XElement e, Parte parte, Ctx ctx, Tf t, bool soDecoracao, StringBuilder sb)
        {
            var gp = e.Element(P + "grpSpPr");
            var xf = gp == null ? null : gp.Element(A + "xfrm");
            var nt = t;
            if (xf != null && xf.Element(A + "off") != null && xf.Element(A + "chOff") != null)
            {
                double ox = Emu(At(xf.Element(A + "off"), "x"), 0), oy = Emu(At(xf.Element(A + "off"), "y"), 0);
                double ex = Emu(At(xf.Element(A + "ext"), "cx"), 0), ey = Emu(At(xf.Element(A + "ext"), "cy"), 0);
                double cx = Emu(At(xf.Element(A + "chOff"), "x"), 0), cy = Emu(At(xf.Element(A + "chOff"), "y"), 0);
                double cex = Emu(At(xf.Element(A + "chExt"), "cx"), 0), cey = Emu(At(xf.Element(A + "chExt"), "cy"), 0);
                double kx = cex != 0 ? ex / cex : 1, ky = cey != 0 ? ey / cey : 1;
                nt.A = t.A * kx; nt.B = t.A * (ox - cx * kx) + t.B;
                nt.C = t.C * ky; nt.D = t.C * (oy - cy * ky) + t.D;
            }
            Arvore(e, parte, ctx, nt, soDecoracao, sb);
        }

        void Quadro(XElement e, Parte parte, Ctx ctx, Tf t, StringBuilder sb)
        {
            var xfrm = e.Element(P + "xfrm");
            var tbl = e.Descendants(A + "tbl").FirstOrDefault();
            if (xfrm == null || tbl == null || xfrm.Element(A + "off") == null) return;
            double x, y, w, h;
            Caixa(xfrm, t, out x, out y, out w, out h);
            var grid = tbl.Element(A + "tblGrid");
            var cols = grid == null ? new List<double>() : grid.Elements(A + "gridCol").Select(g => t.A * Emu(At(g, "w"), 0) / 12700).ToList();
            var tblPr = tbl.Element(A + "tblPr");
            bool cabecalho = At(tblPr, "firstRow") == "1", faixas = At(tblPr, "bandRow") == "1";
            bool estiloPadrao = tblPr != null && tblPr.Element(A + "tableStyleId") != null;
            string acento = "#" + (Esquema("accent1", ctx) ?? "4472C4");
            string claro1 = Cores.Aplicar(acento, new[] { new XElement(A + "tint", new XAttribute("val", "20000")) });
            string claro2 = Cores.Aplicar(acento, new[] { new XElement(A + "tint", new XAttribute("val", "40000")) });

            sb.Append("<table class=\"forma\" style=\"left:" + Pt(x) + ";top:" + Pt(y) + ";width:" + Pt(cols.Count > 0 ? cols.Sum() : w) + "\"><colgroup>");
            foreach (var c in cols) sb.Append("<col style=\"width:" + Pt(c) + "\">");
            sb.Append("</colgroup>");
            int li = 0;
            var listasVazias = new List<XElement>();
            if (ctx.Master != null)
            {
                var ts = ctx.Master.Raiz.Element(P + "txStyles");
                if (ts != null && ts.Element(P + "otherStyle") != null) listasVazias.Add(ts.Element(P + "otherStyle"));
            }
            foreach (var tr in tbl.Elements(A + "tr"))
            {
                bool ehCab = cabecalho && li == 0;
                sb.Append("<tr style=\"height:" + Pt(t.C * Emu(At(tr, "h"), 0) / 12700) + "\">");
                foreach (var tc in tr.Elements(A + "tc"))
                {
                    if (At(tc, "hMerge") == "1" || At(tc, "vMerge") == "1") continue;
                    int span = (int)Emu(At(tc, "gridSpan"), 1), rspan = (int)Emu(At(tc, "rowSpan"), 1);
                    var tcPr = tc.Element(A + "tcPr");
                    string fundo = tcPr == null ? null : Preenchimento(new[] { tcPr }, null, parte, ctx);
                    if (fundo == null && (tcPr == null || tcPr.Element(A + "noFill") == null) && estiloPadrao)
                        fundo = ehCab ? acento : (faixas ? ((li - (cabecalho ? 1 : 0)) % 2 == 0 ? claro2 : claro1) : claro1);
                    var css = new StringBuilder();
                    if (fundo != null) css.Append("background:" + fundo + ";");
                    foreach (var lado in new[] { new[] { "lnL", "left" }, new[] { "lnR", "right" }, new[] { "lnT", "top" }, new[] { "lnB", "bottom" } })
                    {
                        var ln = tcPr == null ? null : tcPr.Element(A + lado[0]);
                        string borda = estiloPadrao ? "1px solid #fff" : "1px solid #bbb";
                        if (ln != null)
                        {
                            if (ln.Element(A + "noFill") != null) borda = "none";
                            else
                            {
                                string c = Cor(ln.Element(A + "solidFill"), ctx);
                                if (c != null) borda = Pt(Emu(At(ln, "w"), 12700) / 12700) + " solid " + c;
                            }
                        }
                        css.Append("border-" + lado[1] + ":" + borda + ";");
                    }
                    double ml = Emu(At(tcPr, "marL"), 91440) / 12700, mr = Emu(At(tcPr, "marR"), 91440) / 12700;
                    double mt = Emu(At(tcPr, "marT"), 45720) / 12700, mb = Emu(At(tcPr, "marB"), 45720) / 12700;
                    css.Append("padding:" + Pt(mt) + " " + Pt(mr) + " " + Pt(mb) + " " + Pt(ml) + ";");
                    string anc = At(tcPr, "anchor");
                    if (anc == "ctr") css.Append("vertical-align:middle;");
                    else if (anc == "b") css.Append("vertical-align:bottom;");
                    var conteudo = new StringBuilder();
                    var tx = tc.Element(A + "txBody");
                    if (tx != null)
                    {
                        var cont = new int[9];
                        foreach (var p in tx.Elements(A + "p")) conteudo.Append(Paragrafo(p, listasVazias, ctx, 1, 0, cont));
                    }
                    string html = conteudo.ToString();
                    if (ehCab && estiloPadrao) html = "<div style=\"font-weight:bold\">" + Regex.Replace(html, @"color:#000000", "color:#ffffff") + "</div>";
                    sb.Append("<td" + (span > 1 ? " colspan=\"" + span + "\"" : "") + (rspan > 1 ? " rowspan=\"" + rspan + "\"" : "") + " style=\"" + css + "\">" + html + "</td>");
                }
                sb.Append("</tr>");
                li++;
            }
            sb.Append("</table>");
        }
    }
}
