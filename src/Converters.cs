using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace ConversorPDF
{
    enum Tipo { Imagem, Pdf, Word, Excel, PowerPoint, Web, Texto, Markdown, Csv, Email, Site, Desconhecido }

    static class Conversor
    {
        static readonly Dictionary<string, Tipo> Extensoes = new Dictionary<string, Tipo>();

        static Conversor()
        {
            Reg(Tipo.Imagem, ".jpg .jpeg .jpe .jfif .png .bmp .gif .tif .tiff .webp .heic .heif .avif .ico .jxl .dng .cr2 .cr3 .nef .arw .orf .rw2");
            Reg(Tipo.Pdf, ".pdf");
            Reg(Tipo.Word, ".doc .docx .docm .dot .dotx .dotm .rtf .odt .wpd .wps");
            Reg(Tipo.Excel, ".xls .xlsx .xlsm .xlsb .xltx .ods");
            Reg(Tipo.PowerPoint, ".ppt .pptx .pptm .pps .ppsx .potx .odp");
            Reg(Tipo.Web, ".html .htm .xhtml .mht .mhtml .svg");
            Reg(Tipo.Markdown, ".md .markdown");
            Reg(Tipo.Csv, ".csv .tsv");
            Reg(Tipo.Email, ".eml");
            Reg(Tipo.Texto, ".txt .log .json .xml .ini .cfg .conf .yml .yaml .toml .cs .js .ts .jsx .tsx .py .java .c .cpp .h .hpp .css .scss " +
                ".sql .bat .cmd .ps1 .sh .php .rb .go .rs .kt .swift .vb .lua .r .tex .srt .vtt .nfo .properties .gradle .env");
        }

        static void Reg(Tipo t, string exts) { foreach (var e in exts.Split(' ')) if (e.Length > 0) Extensoes[e] = t; }

        public static Tipo Detectar(string caminho)
        {
            Tipo t;
            return Extensoes.TryGetValue(Path.GetExtension(caminho).ToLowerInvariant(), out t) ? t : Tipo.Desconhecido;
        }

        public static string Descricao(Tipo t)
        {
            switch (t)
            {
                case Tipo.Imagem: return "Imagem";
                case Tipo.Pdf: return "PDF";
                case Tipo.Word: return "Documento de texto (Word)";
                case Tipo.Excel: return "Planilha (Excel)";
                case Tipo.PowerPoint: return "Apresentação (PowerPoint)";
                case Tipo.Web: return "Página web / SVG";
                case Tipo.Texto: return "Arquivo de texto";
                case Tipo.Markdown: return "Markdown";
                case Tipo.Csv: return "Tabela CSV";
                case Tipo.Email: return "E-mail (.eml)";
                case Tipo.Site: return "Site (link)";
                default: return "Desconhecido";
            }
        }

        static string pastaTemp;
        public static string PastaTemp
        {
            get
            {
                if (pastaTemp == null)
                {
                    pastaTemp = Path.Combine(Path.GetTempPath(), "ConversorPDF", Process.GetCurrentProcess().Id.ToString());
                    Directory.CreateDirectory(pastaTemp);
                }
                return pastaTemp;
            }
        }

        public static string NovoTemp(string ext)
        {
            return Path.Combine(PastaTemp, Guid.NewGuid().ToString("N").Substring(0, 12) + ext);
        }

        public static void LimparTemp()
        {
            try { if (pastaTemp != null) Directory.Delete(pastaTemp, true); } catch { }
        }

        // Converte qualquer item não-imagem em um PDF temporário e devolve o caminho
        public static string ParaPdf(string caminho, string url, Tipo tipo, Opcoes o)
        {
            string saida = NovoTemp(".pdf");
            switch (tipo)
            {
                case Tipo.Word:
                case Tipo.Excel:
                case Tipo.PowerPoint:
                    ConverterOffice(caminho, tipo, saida, o);
                    break;
                case Tipo.Web:
                    Edge(new Uri(caminho).AbsoluteUri, saida, 5000);
                    break;
                case Tipo.Site:
                    Edge(url, saida, 20000);
                    break;
                case Tipo.Texto:
                    HtmlParaPdf(Html.DeTexto(caminho, o), saida, 2000);
                    break;
                case Tipo.Markdown:
                    HtmlParaPdf(Html.DeMarkdown(caminho, o), saida, 3000);
                    break;
                case Tipo.Csv:
                    HtmlParaPdf(Html.DeCsv(caminho, o), saida, 3000);
                    break;
                case Tipo.Email:
                    HtmlParaPdf(Html.DeEmail(caminho, o), saida, 8000);
                    break;
                default:
                    throw new Exception("Formato não suportado.");
            }
            return saida;
        }

        public static void HtmlParaPdf(string html, string saida, int esperaMs)
        {
            string arq = NovoTemp(".html");
            File.WriteAllText(arq, html, new UTF8Encoding(false));
            Edge(new Uri(arq).AbsoluteUri, saida, esperaMs);
        }

        // ---------- Microsoft Edge (já vem no Windows) ----------

        public static string AcharEdge()
        {
            foreach (var pf in new[] { Environment.GetEnvironmentVariable("ProgramFiles(x86)"), Environment.GetEnvironmentVariable("ProgramFiles"),
                                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) })
            {
                if (string.IsNullOrEmpty(pf)) continue;
                string c = Path.Combine(pf, @"Microsoft\Edge\Application\msedge.exe");
                if (File.Exists(c)) return c;
            }
            return null;
        }

        static void Edge(string url, string saida, int esperaMs)
        {
            string edge = AcharEdge();
            if (edge == null) throw new Exception("O Microsoft Edge não foi encontrado (ele é usado para converter este tipo de arquivo).");
            string perfil = Path.Combine(PastaTemp, "edge");
            string args = "--headless --disable-gpu --no-first-run --no-default-browser-check --disable-extensions --disable-sync " +
                          "--hide-scrollbars --no-pdf-header-footer --print-to-pdf-no-header --run-all-compositor-stages-before-draw " +
                          "--virtual-time-budget=" + esperaMs + " --user-data-dir=\"" + perfil + "\" --print-to-pdf=\"" + saida + "\" \"" + url + "\"";
            if (!Executar(edge, args, 180000) || !File.Exists(saida) || new FileInfo(saida).Length == 0)
                throw new Exception("Não foi possível converter (o Edge não gerou o PDF).");
        }

        static bool Executar(string exe, string args, int timeoutMs)
        {
            var psi = new ProcessStartInfo(exe, args) { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden };
            using (var p = Process.Start(psi))
            {
                if (p.WaitForExit(timeoutMs)) return true;
                try { p.Kill(); } catch { }
                return false;
            }
        }

        // ---------- Word / Excel / PowerPoint ----------

        public static bool TemMsOffice(Tipo t)
        {
            string prog = t == Tipo.Word ? "Word.Application" : (t == Tipo.Excel ? "Excel.Application" : "PowerPoint.Application");
            return Type.GetTypeFromProgID(prog) != null;
        }

        public static string AcharLibreOffice()
        {
            foreach (var pf in new[] { Environment.GetEnvironmentVariable("ProgramFiles"), Environment.GetEnvironmentVariable("ProgramFiles(x86)") })
            {
                if (string.IsNullOrEmpty(pf)) continue;
                string c = Path.Combine(pf, @"LibreOffice\program\soffice.exe");
                if (File.Exists(c)) return c;
            }
            return null;
        }

        static void ConverterOffice(string caminho, Tipo tipo, string saida, Opcoes o)
        {
            string ext = Path.GetExtension(caminho).ToLowerInvariant();
            string erro = null;
            if (TemMsOffice(tipo))
            {
                try { ViaMsOffice(caminho, tipo, saida); return; }
                catch (Exception ex) { erro = ex.Message; }
            }
            string so = AcharLibreOffice();
            if (so != null)
            {
                ViaLibreOffice(so, caminho, saida);
                return;
            }
            if (ext == ".docx" || ext == ".docm" || ext == ".dotx" || ext == ".dotm")
            { HtmlParaPdf(new DocxConv().Converter(caminho), saida, 3000); return; }
            if (ext == ".xlsx" || ext == ".xlsm" || ext == ".xltx")
            { HtmlParaPdf(new XlsxConv().Converter(caminho, o), saida, 3000); return; }
            if (ext == ".pptx" || ext == ".ppsx" || ext == ".pptm" || ext == ".potx")
            { HtmlParaPdf(new PptxConv().Converter(caminho), saida, 3000); return; }
            if (erro != null) throw new Exception("O Office não conseguiu converter: " + erro);
            throw new Exception("Arquivos " + ext + " precisam do Microsoft Office ou do LibreOffice (gratuito: libreoffice.org) instalado.");
        }

        static void ViaMsOffice(string caminho, Tipo tipo, string saida)
        {
            if (tipo == Tipo.Word)
            {
                dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
                try
                {
                    app.Visible = false;
                    app.DisplayAlerts = 0;
                    dynamic doc = app.Documents.Open(caminho, false, true, false);
                    try { doc.ExportAsFixedFormat(saida, 17); }
                    finally { doc.Close(0); }
                }
                finally { app.Quit(0); Marshal.FinalReleaseComObject(app); }
            }
            else if (tipo == Tipo.Excel)
            {
                dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
                try
                {
                    app.Visible = false;
                    app.DisplayAlerts = false;
                    dynamic wb = app.Workbooks.Open(caminho, 0, true);
                    try { wb.ExportAsFixedFormat(0, saida); }
                    finally { wb.Close(false); }
                }
                finally { app.Quit(); Marshal.FinalReleaseComObject(app); }
            }
            else
            {
                dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
                try
                {
                    dynamic pres = app.Presentations.Open(caminho, -1, 0, 0);
                    try { pres.SaveAs(saida, 32); }
                    finally { pres.Close(); }
                }
                finally { app.Quit(); Marshal.FinalReleaseComObject(app); }
            }
            if (!File.Exists(saida)) throw new Exception("o arquivo não foi gerado");
        }

        static void ViaLibreOffice(string so, string caminho, string saida)
        {
            string pasta = NovoTemp("");
            Directory.CreateDirectory(pasta);
            string perfil = new Uri(Path.Combine(PastaTemp, "lo-perfil")).AbsoluteUri;
            string args = "-env:UserInstallation=" + perfil + " --headless --norestore --nologo --convert-to pdf --outdir \"" + pasta + "\" \"" + caminho + "\"";
            Executar(so, args, 300000);
            string gerado = Path.Combine(pasta, Path.GetFileNameWithoutExtension(caminho) + ".pdf");
            if (!File.Exists(gerado)) throw new Exception("O LibreOffice não conseguiu converter este arquivo.");
            File.Move(gerado, saida);
        }
    }

    // ---------- Geração de HTML para texto, Markdown, CSV e e-mail ----------

    static class Html
    {
        public static string Esc(string s) { return WebUtility.HtmlEncode(s ?? ""); }

        public static string Pagina(string titulo, string css, string corpo, string extraHead = "")
        {
            return "<!doctype html><html><head><meta charset=\"utf-8\"><title>" + Esc(titulo) + "</title>" + extraHead +
                   "<style>" + css + "</style></head><body>" + corpo + "</body></html>";
        }

        public static string LerTexto(string caminho)
        {
            return DecodificarTexto(File.ReadAllBytes(caminho));
        }

        public static string DecodificarTexto(byte[] b)
        {
            if (b.Length >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) return Encoding.UTF8.GetString(b, 3, b.Length - 3);
            if (b.Length >= 2 && b[0] == 0xFF && b[1] == 0xFE) return Encoding.Unicode.GetString(b, 2, b.Length - 2);
            if (b.Length >= 2 && b[0] == 0xFE && b[1] == 0xFF) return Encoding.BigEndianUnicode.GetString(b, 2, b.Length - 2);
            try { return new UTF8Encoding(false, true).GetString(b); }
            catch { return Encoding.GetEncoding(1252).GetString(b); }
        }

        public static string DeTexto(string caminho, Opcoes o)
        {
            string css = o.CssPagina(false) +
                "body{margin:0} pre{white-space:pre-wrap;word-wrap:break-word;overflow-wrap:anywhere;font-family:Consolas,'Courier New',monospace;" +
                "font-size:9.5pt;line-height:1.4;tab-size:4;margin:0}";
            return Pagina(Path.GetFileName(caminho), css, "<pre>" + Esc(LerTexto(caminho)) + "</pre>");
        }

        // ----- Markdown -----

        public static string DeMarkdown(string caminho, Opcoes o)
        {
            string pasta = Path.GetDirectoryName(Path.GetFullPath(caminho));
            string css = o.CssPagina(false) +
                "body{font-family:'Segoe UI',Arial,sans-serif;font-size:11pt;line-height:1.55;color:#1f2328;margin:0}" +
                "h1,h2{border-bottom:1px solid #d8dee4;padding-bottom:.25em}h1{font-size:2em}h2{font-size:1.5em}h3{font-size:1.25em}" +
                "h1,h2,h3,h4,h5,h6{margin:1.2em 0 .5em;line-height:1.25;break-after:avoid}" +
                "code{font-family:Consolas,monospace;background:#f3f4f6;padding:.1em .35em;border-radius:4px;font-size:.9em}" +
                "pre{background:#f6f8fa;padding:12px;border-radius:6px;white-space:pre-wrap;word-wrap:break-word}pre code{background:none;padding:0}" +
                "blockquote{margin:0 0 1em;padding:0 1em;color:#59636e;border-left:.25em solid #d1d9e0}" +
                "table{border-collapse:collapse;margin:1em 0}th,td{border:1px solid #d1d9e0;padding:5px 12px}th{background:#f6f8fa}" +
                "img{max-width:100%}hr{border:0;border-top:2px solid #d1d9e0;margin:1.5em 0}a{color:#0969da;text-decoration:none}" +
                "ul,ol{padding-left:2em}li{margin:.2em 0}";
            return Pagina(Path.GetFileName(caminho), css, Markdown.Converter(LerTexto(caminho), pasta));
        }

        // ----- CSV -----

        public static string DeCsv(string caminho, Opcoes o)
        {
            string texto = LerTexto(caminho);
            char sep = Path.GetExtension(caminho).ToLowerInvariant() == ".tsv" ? '\t' : DetectarSeparador(texto);
            var linhas = LerCsv(texto, sep);
            int cols = linhas.Count == 0 ? 0 : linhas.Max(l => l.Count);
            bool paisagem = cols > 6;
            double larg = o.LarguraUtilPx(paisagem);
            var sb = new StringBuilder("<table class=\"ajustar\">");
            for (int i = 0; i < linhas.Count; i++)
            {
                if (i == 0) sb.Append("<thead>");
                sb.Append("<tr>");
                for (int c = 0; c < cols; c++)
                {
                    string v = c < linhas[i].Count ? linhas[i][c] : "";
                    bool num = Regex.IsMatch(v.Trim(), @"^[-+]?(R\$\s*)?[\d.,]+%?$");
                    string tag = i == 0 ? "th" : "td";
                    sb.Append("<" + tag + (num && i > 0 ? " class=\"n\"" : "") + ">" + Esc(v) + "</" + tag + ">");
                }
                sb.Append("</tr>");
                if (i == 0) sb.Append("</thead><tbody>");
            }
            sb.Append("</tbody></table>");
            string css = o.CssPagina(paisagem) +
                "body{margin:0;font-family:'Segoe UI',Arial,sans-serif;font-size:9pt}" +
                "table{border-collapse:collapse}th,td{border:1px solid #b9b9b9;padding:2px 6px;vertical-align:top;max-width:320px;overflow-wrap:anywhere}" +
                "th{background:#e8edf3;text-align:left}td.n{text-align:right;white-space:nowrap}tr{break-inside:avoid}";
            string script = "<script>window.addEventListener('load',function(){var W=" + larg.ToString("0", CultureInfo.InvariantCulture) +
                ";document.body.style.width=W+'px';document.querySelectorAll('table.ajustar').forEach(function(t){" +
                "if(t.scrollWidth>W)t.style.zoom=(W/t.scrollWidth);});});</script>";
            return Pagina(Path.GetFileName(caminho), css, sb + script);
        }

        static char DetectarSeparador(string t)
        {
            int fim = t.IndexOf('\n');
            string l = fim < 0 ? t : t.Substring(0, fim);
            int pv = l.Count(c => c == ';'), vg = l.Count(c => c == ','), tb = l.Count(c => c == '\t');
            if (tb > pv && tb > vg) return '\t';
            return pv >= vg && pv > 0 ? ';' : ',';
        }

        static List<List<string>> LerCsv(string t, char sep)
        {
            var linhas = new List<List<string>>();
            var linha = new List<string>();
            var campo = new StringBuilder();
            bool aspas = false;
            for (int i = 0; i < t.Length; i++)
            {
                char c = t[i];
                if (aspas)
                {
                    if (c == '"')
                    {
                        if (i + 1 < t.Length && t[i + 1] == '"') { campo.Append('"'); i++; }
                        else aspas = false;
                    }
                    else campo.Append(c);
                }
                else if (c == '"' && campo.Length == 0) aspas = true;
                else if (c == sep) { linha.Add(campo.ToString()); campo.Clear(); }
                else if (c == '\n' || c == '\r')
                {
                    if (c == '\r' && i + 1 < t.Length && t[i + 1] == '\n') i++;
                    linha.Add(campo.ToString()); campo.Clear();
                    linhas.Add(linha); linha = new List<string>();
                }
                else campo.Append(c);
            }
            if (campo.Length > 0 || linha.Count > 0) { linha.Add(campo.ToString()); linhas.Add(linha); }
            while (linhas.Count > 0 && linhas[linhas.Count - 1].All(x => x.Length == 0)) linhas.RemoveAt(linhas.Count - 1);
            return linhas;
        }

        // ----- E-mail (.eml) -----

        class Email
        {
            public string Html, Texto;
            public Dictionary<string, string> Cids = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public List<string> Anexos = new List<string>();
        }

        class Parte
        {
            public Dictionary<string, string> Cab = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public string Corpo = "";
            public string Get(string k) { string v; return Cab.TryGetValue(k, out v) ? v : null; }
        }

        static Parte LerParte(string raw)
        {
            var p = new Parte();
            int sep = raw.IndexOf("\r\n\r\n"), sl = 4;
            int s2 = raw.IndexOf("\n\n");
            if (sep < 0 || (s2 >= 0 && s2 < sep)) { sep = s2; sl = 2; }
            string cab = sep < 0 ? raw : raw.Substring(0, sep);
            p.Corpo = sep < 0 ? "" : raw.Substring(sep + sl);
            cab = Regex.Replace(cab, @"\r?\n[ \t]+", " ");
            foreach (var linha in cab.Split('\n'))
            {
                int c = linha.IndexOf(':');
                if (c <= 0) continue;
                string k = linha.Substring(0, c).Trim();
                if (!p.Cab.ContainsKey(k)) p.Cab[k] = linha.Substring(c + 1).Trim();
            }
            return p;
        }

        static string Param(string valor, string nome)
        {
            if (valor == null) return null;
            var m = Regex.Match(valor, nome + @"\*=([^;]+)", RegexOptions.IgnoreCase);
            if (m.Success)
            {
                string v = m.Groups[1].Value.Trim().Trim('"');
                int a = v.IndexOf("''");
                if (a >= 0)
                {
                    string cs = v.Substring(0, a);
                    try { return Charset(cs).GetString(PdfWriter.Latin1.GetBytes(Uri.UnescapeDataString(v.Substring(a + 2)))); }
                    catch { return Uri.UnescapeDataString(v.Substring(a + 2)); }
                }
            }
            m = Regex.Match(valor, "(?:^|;)\\s*" + nome + "\\s*=\\s*(\"([^\"]*)\"|([^;]+))", RegexOptions.IgnoreCase);
            if (!m.Success) return null;
            return (m.Groups[2].Success ? m.Groups[2].Value : m.Groups[3].Value).Trim();
        }

        static Encoding Charset(string cs)
        {
            try { return string.IsNullOrEmpty(cs) ? Encoding.GetEncoding(1252) : Encoding.GetEncoding(cs.Trim().Trim('"')); }
            catch { return Encoding.UTF8; }
        }

        static byte[] DecodificarCorpo(Parte p)
        {
            string cte = (p.Get("Content-Transfer-Encoding") ?? "").Trim().ToLowerInvariant();
            if (cte == "base64")
            {
                string s = Regex.Replace(p.Corpo, @"[^A-Za-z0-9+/=]", "");
                s = s.TrimEnd('=');
                s += new string('=', (4 - s.Length % 4) % 4);
                try { return Convert.FromBase64String(s); } catch { return new byte[0]; }
            }
            if (cte == "quoted-printable") return QP(p.Corpo, false);
            return PdfWriter.Latin1.GetBytes(p.Corpo);
        }

        static byte[] QP(string s, bool cabecalho)
        {
            var o = new List<byte>();
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (cabecalho && c == '_') { o.Add(32); continue; }
                if (c == '=')
                {
                    if (i + 1 < s.Length && s[i + 1] == '\n') { i++; continue; }
                    if (i + 2 < s.Length && s[i + 1] == '\r' && s[i + 2] == '\n') { i += 2; continue; }
                    int v;
                    if (i + 2 < s.Length && int.TryParse(s.Substring(i + 1, 2), NumberStyles.HexNumber, null, out v)) { o.Add((byte)v); i += 2; continue; }
                }
                o.Add((byte)c);
            }
            return o.ToArray();
        }

        static string DecodificarCabecalho(string s)
        {
            if (s == null) return "";
            s = Regex.Replace(s, @"(\?=)\s+(=\?)", "$1$2");
            s = Regex.Replace(s, @"=\?([^?]+)\?([BbQq])\?([^?]*)\?=", m =>
            {
                try
                {
                    byte[] b = m.Groups[2].Value.ToUpperInvariant() == "B" ? Convert.FromBase64String(m.Groups[3].Value) : QP(m.Groups[3].Value, true);
                    return Charset(m.Groups[1].Value.Split('*')[0]).GetString(b);
                }
                catch { return m.Value; }
            });
            // cabeçalhos com UTF-8 "cru"
            if (s.Any(ch => ch > 127))
            {
                try { return new UTF8Encoding(false, true).GetString(PdfWriter.Latin1.GetBytes(s)); } catch { }
            }
            return s;
        }

        static void Coletar(Parte p, Email e, int prof)
        {
            if (prof > 20) return;
            string ct = p.Get("Content-Type") ?? "text/plain";
            string tipo = ct.Split(';')[0].Trim().ToLowerInvariant();
            if (tipo.StartsWith("multipart/"))
            {
                string bnd = Param(ct, "boundary");
                if (bnd == null) return;
                var pedacos = p.Corpo.Split(new[] { "--" + bnd }, StringSplitOptions.None);
                for (int i = 1; i < pedacos.Length; i++)
                {
                    string ch = pedacos[i];
                    if (ch.StartsWith("--")) break;
                    ch = ch.TrimStart(' ', '\t');
                    if (ch.StartsWith("\r\n")) ch = ch.Substring(2); else if (ch.StartsWith("\n")) ch = ch.Substring(1);
                    Coletar(LerParte(ch), e, prof + 1);
                }
                return;
            }
            string disp = p.Get("Content-Disposition") ?? "";
            bool anexo = disp.Trim().ToLowerInvariant().StartsWith("attachment");
            byte[] dados = DecodificarCorpo(p);
            string nome = Param(disp, "filename") ?? Param(ct, "name");
            if (!anexo && tipo == "text/html" && e.Html == null) { e.Html = Charset(Param(ct, "charset")).GetString(dados); return; }
            if (!anexo && tipo == "text/plain" && e.Texto == null) { e.Texto = Charset(Param(ct, "charset")).GetString(dados); return; }
            string cid = p.Get("Content-ID");
            if (cid != null && tipo.StartsWith("image/"))
                e.Cids[cid.Trim().Trim('<', '>')] = "data:" + tipo + ";base64," + Convert.ToBase64String(dados);
            if (nome != null || anexo)
                e.Anexos.Add(DecodificarCabecalho(nome ?? "anexo") + " (" + Tamanho(dados.Length) + ")");
        }

        public static string Tamanho(long b)
        {
            if (b < 1024) return b + " B";
            if (b < 1024 * 1024) return (b / 1024.0).ToString("0") + " KB";
            return (b / 1048576.0).ToString("0.0") + " MB";
        }

        public static string DeEmail(string caminho, Opcoes o)
        {
            string raw = PdfWriter.Latin1.GetString(File.ReadAllBytes(caminho));
            var raiz = LerParte(raw);
            var e = new Email();
            Coletar(raiz, e, 0);

            var cab = new StringBuilder("<div class=\"cab\"><h1>" + Esc(DecodificarCabecalho(raiz.Get("Subject") ?? "(sem assunto)")) + "</h1><table>");
            foreach (var par in new[] { new[] { "From", "De" }, new[] { "To", "Para" }, new[] { "Cc", "Cc" }, new[] { "Date", "Data" } })
            {
                string v = raiz.Get(par[0]);
                if (v == null) continue;
                if (par[0] == "Date")
                {
                    DateTimeOffset dt;
                    if (DateTimeOffset.TryParse(Regex.Replace(v, @"\s*\(.*\)$", ""), CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                        v = dt.ToLocalTime().ToString("dd/MM/yyyy HH:mm");
                }
                cab.Append("<tr><td class=\"k\">" + par[1] + ":</td><td>" + Esc(DecodificarCabecalho(v)) + "</td></tr>");
            }
            if (e.Anexos.Count > 0) cab.Append("<tr><td class=\"k\">Anexos:</td><td>" + Esc(string.Join(", ", e.Anexos)) + "</td></tr>");
            cab.Append("</table></div>");

            string corpo, estilos = "";
            if (e.Html != null)
            {
                string h = e.Html;
                foreach (Match m in Regex.Matches(h, @"<style[^>]*>[\s\S]*?</style>", RegexOptions.IgnoreCase)) estilos += m.Value;
                var mb = Regex.Match(h, @"<body[^>]*>([\s\S]*)</body>", RegexOptions.IgnoreCase);
                if (mb.Success) h = mb.Groups[1].Value;
                h = Regex.Replace(h, @"cid:([^""'\s>)]+)", m =>
                {
                    string uri;
                    return e.Cids.TryGetValue(m.Groups[1].Value, out uri) ? uri : m.Value;
                });
                corpo = h;
            }
            else corpo = "<pre class=\"txt\">" + Esc(e.Texto ?? "") + "</pre>";

            string css = o.CssPagina(false) +
                "body{margin:0;font-family:'Segoe UI',Arial,sans-serif;font-size:10.5pt}" +
                ".cab{border-bottom:2px solid #444;margin-bottom:14px;padding-bottom:8px}.cab h1{font-size:15pt;margin:0 0 8px}" +
                ".cab table{border-collapse:collapse;font-size:9.5pt}.cab td{padding:1px 8px 1px 0;vertical-align:top}.cab td.k{color:#555;font-weight:600}" +
                "pre.txt{white-space:pre-wrap;word-wrap:break-word;font-family:'Segoe UI',Arial,sans-serif}img{max-width:100%;height:auto}";
            string head = "<meta http-equiv=\"Content-Security-Policy\" content=\"script-src 'none'\">" + estilos;
            return Pagina(DecodificarCabecalho(raiz.Get("Subject") ?? "E-mail"), css, cab + "<div class=\"corpo\">" + corpo + "</div>", head);
        }
    }

    // Conversor de Markdown simples (títulos, listas, tabelas, código, citações, links e imagens)
    static class Markdown
    {
        public static string Converter(string md, string pasta)
        {
            var linhas = md.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            var sb = new StringBuilder();
            var par = new List<string>();
            var listas = new Stack<Tuple<string, int>>();  // tag, indentação
            bool emCitacao = false;

            Action fecharPar = () =>
            {
                if (par.Count == 0) return;
                var partes = new StringBuilder();
                foreach (var x in par) partes.Append(Inline(x.Trim(), pasta) + (x.EndsWith("  ") ? "<br>" : " "));
                sb.Append("<p>" + partes.ToString().Trim() + "</p>");
                par.Clear();
            };
            Action fecharListas = () => { while (listas.Count > 0) sb.Append("</li></" + listas.Pop().Item1 + ">"); };
            Action fecharCitacao = () => { if (emCitacao) { fecharPar(); sb.Append("</blockquote>"); emCitacao = false; } };

            for (int i = 0; i < linhas.Length; i++)
            {
                string l = linhas[i];
                string t = l.Trim();

                if (t.StartsWith("```") || t.StartsWith("~~~"))
                {
                    fecharPar(); fecharListas(); fecharCitacao();
                    string cerca = t.Substring(0, 3);
                    var cod = new StringBuilder();
                    for (i++; i < linhas.Length && !linhas[i].Trim().StartsWith(cerca); i++) cod.Append(linhas[i]).Append('\n');
                    sb.Append("<pre><code>" + Html.Esc(cod.ToString().TrimEnd('\n')) + "</code></pre>");
                    continue;
                }
                if (t.Length == 0) { fecharPar(); fecharCitacao(); if (i + 1 < linhas.Length && !EhItem(linhas[i + 1])) fecharListas(); continue; }

                var h = Regex.Match(t, @"^(#{1,6})\s+(.*?)\s*#*$");
                if (h.Success)
                {
                    fecharPar(); fecharListas(); fecharCitacao();
                    int n = h.Groups[1].Value.Length;
                    sb.Append("<h" + n + ">" + Inline(h.Groups[2].Value, pasta) + "</h" + n + ">");
                    continue;
                }
                if (Regex.IsMatch(t, @"^([-*_])(\s*\1){2,}$"))
                {
                    fecharPar(); fecharListas(); fecharCitacao();
                    sb.Append("<hr>");
                    continue;
                }
                if (t.StartsWith(">"))
                {
                    fecharListas();
                    if (!emCitacao) { fecharPar(); sb.Append("<blockquote>"); emCitacao = true; }
                    par.Add(t.Substring(1).TrimStart());
                    continue;
                }
                if (t.Contains("|") && i + 1 < linhas.Length && Regex.IsMatch(linhas[i + 1].Trim(), @"^\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)*\|?$"))
                {
                    fecharPar(); fecharListas(); fecharCitacao();
                    var alinh = Celulas(linhas[i + 1]).Select(c => c.StartsWith(":") && c.EndsWith(":") ? "center" : (c.EndsWith(":") ? "right" : "left")).ToList();
                    sb.Append("<table><thead><tr>");
                    var cab = Celulas(l);
                    for (int c = 0; c < cab.Count; c++) sb.Append("<th style=\"text-align:" + (c < alinh.Count ? alinh[c] : "left") + "\">" + Inline(cab[c], pasta) + "</th>");
                    sb.Append("</tr></thead><tbody>");
                    for (i += 2; i < linhas.Length && linhas[i].Contains("|") && linhas[i].Trim().Length > 0; i++)
                    {
                        sb.Append("<tr>");
                        var cs = Celulas(linhas[i]);
                        for (int c = 0; c < cab.Count; c++)
                            sb.Append("<td style=\"text-align:" + (c < alinh.Count ? alinh[c] : "left") + "\">" + (c < cs.Count ? Inline(cs[c], pasta) : "") + "</td>");
                        sb.Append("</tr>");
                    }
                    i--;
                    sb.Append("</tbody></table>");
                    continue;
                }
                var li = Regex.Match(l, @"^(\s*)([-*+]|\d+[.)])\s+(.*)$");
                if (li.Success)
                {
                    fecharPar(); fecharCitacao();
                    int ind = li.Groups[1].Value.Replace("\t", "    ").Length;
                    string tag = char.IsDigit(li.Groups[2].Value[0]) ? "ol" : "ul";
                    while (listas.Count > 0 && listas.Peek().Item2 > ind) sb.Append("</li></" + listas.Pop().Item1 + ">");
                    if (listas.Count == 0 || listas.Peek().Item2 < ind)
                        { sb.Append("<" + tag + ">"); listas.Push(Tuple.Create(tag, ind)); }
                    else if (listas.Peek().Item1 != tag)
                        { sb.Append("</li></" + listas.Pop().Item1 + "><" + tag + ">"); listas.Push(Tuple.Create(tag, ind)); }
                    else sb.Append("</li>");
                    string item = li.Groups[3].Value;
                    var cb = Regex.Match(item, @"^\[([ xX])\]\s+(.*)$");
                    if (cb.Success) item = (cb.Groups[1].Value == " " ? "☐ " : "☑ ") + cb.Groups[2].Value;
                    sb.Append("<li>" + Inline(item, pasta));
                    continue;
                }
                if (listas.Count > 0 && l.StartsWith("  ")) { sb.Append(" " + Inline(t, pasta)); continue; }
                fecharListas();
                par.Add(l.TrimStart());
            }
            fecharPar(); fecharListas(); fecharCitacao();
            return sb.ToString();
        }

        static bool EhItem(string l) { return Regex.IsMatch(l, @"^\s*([-*+]|\d+[.)])\s+"); }

        static List<string> Celulas(string l)
        {
            string t = l.Trim();
            if (t.StartsWith("|")) t = t.Substring(1);
            if (t.EndsWith("|")) t = t.Substring(0, t.Length - 1);
            return t.Split('|').Select(c => c.Trim()).ToList();
        }

        static string Inline(string s, string pasta)
        {
            var codigos = new List<string>();
            s = Regex.Replace(s, @"`([^`]+)`", m => { codigos.Add(m.Groups[1].Value); return "\u0001" + (codigos.Count - 1) + "\u0002"; });
            s = Html.Esc(s);
            s = Regex.Replace(s, @"!\[([^\]]*)\]\(([^)\s]+)(?:\s+&quot;[^&]*&quot;)?\)", m => "<img alt=\"" + m.Groups[1].Value + "\" src=\"" + Url(m.Groups[2].Value, pasta) + "\">");
            s = Regex.Replace(s, @"\[([^\]]+)\]\(([^)\s]+)(?:\s+&quot;[^&]*&quot;)?\)", m => "<a href=\"" + m.Groups[2].Value + "\">" + m.Groups[1].Value + "</a>");
            s = Regex.Replace(s, @"(?<![""=])\b(https?://[^\s<]+)", "<a href=\"$1\">$1</a>");
            s = Regex.Replace(s, @"(\*\*|__)(?=\S)(.+?)(?<=\S)\1", "<strong>$2</strong>");
            s = Regex.Replace(s, @"(?<![\w*])(\*|_)(?=\S)(.+?)(?<=\S)\1(?![\w*])", "<em>$2</em>");
            s = Regex.Replace(s, @"~~(.+?)~~", "<del>$1</del>");
            s = Regex.Replace(s, "\u0001(\\d+)\u0002", m => "<code>" + Html.Esc(codigos[int.Parse(m.Groups[1].Value)]) + "</code>");
            return s;
        }

        static string Url(string u, string pasta)
        {
            if (Regex.IsMatch(u, @"^[a-zA-Z][a-zA-Z0-9+.-]*:")) return u;
            try { return new Uri(Path.Combine(pasta, Uri.UnescapeDataString(u).Replace('/', '\\'))).AbsoluteUri; } catch { return u; }
        }
    }
}
