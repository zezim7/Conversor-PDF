using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace ConversorPDF
{
    class Opcoes
    {
        public int PorPagina = 1;      // 1, 2, 4 ou 6 imagens por página
        public int Papel = 0;          // 0 A4, 1 Carta, 2 Ofício, 3 tamanho da imagem
        public int Orientacao = 0;     // 0 automática, 1 retrato, 2 paisagem
        public int Margem = 0;         // 0 normal, 1 pequena, 2 sem margem
        public bool Preencher;         // true = preenche a célula cortando as bordas
        public int Qualidade = 1;      // 0 alta, 1 média, 2 pequena
        public bool Documento;         // modo scanner

        public Opcoes Clone() { return (Opcoes)MemberwiseClone(); }

        // Tamanho de página para conteúdos convertidos via HTML (texto, planilhas CSV etc.)
        public string CssPagina(bool paisagem)
        {
            string papel = Papel == 1 ? "letter" : (Papel == 2 ? "legal" : "A4");
            if (Orientacao == 2) paisagem = true;
            if (Orientacao == 1) paisagem = false;
            string margem = Margem == 0 ? "18mm" : (Margem == 1 ? "10mm" : "5mm");
            return "@page { size: " + papel + (paisagem ? " landscape" : "") + "; margin: " + margem + "; }";
        }

        // Largura útil (px CSS) da página usada em CssPagina
        public double LarguraUtilPx(bool paisagem)
        {
            if (Orientacao == 2) paisagem = true;
            if (Orientacao == 1) paisagem = false;
            double[] p = Layout.TamanhoPapel(Papel == 3 ? 0 : Papel);
            double wPt = paisagem ? p[1] : p[0];
            double mm = Margem == 0 ? 18 : (Margem == 1 ? 10 : 5);
            return (wPt - 2 * mm * 72 / 25.4) * 96 / 72;
        }
    }

    class FonteImagem
    {
        public string Caminho;
        public int Quadro;
        public int Rotacao;
    }

    struct Ret
    {
        public double X, Y, W, H;
        public Ret(double x, double y, double w, double h) { X = x; Y = y; W = w; H = h; }
    }

    class LayoutPagina
    {
        public double W, H;
        public List<Ret> Celulas = new List<Ret>();
    }

    // Coordenadas em pontos, origem no canto inferior esquerdo (padrão PDF)
    static class Layout
    {
        public static double[] TamanhoPapel(int papel)
        {
            switch (papel)
            {
                case 1: return new[] { 612.0, 792.0 };
                case 2: return new[] { 612.0, 1008.0 };
                default: return new[] { 595.28, 841.89 };
            }
        }

        public static double MargemPt(int m) { return m == 0 ? 28 : (m == 1 ? 12 : 0); }

        // tamanhos: largura/altura (só a proporção importa) de cada imagem da página
        public static LayoutPagina Calcular(Opcoes o, IList<double[]> tamanhos)
        {
            int n = Math.Max(1, o.PorPagina);
            double m = MargemPt(o.Margem);
            double esp = o.Margem == 2 ? 0 : (o.Margem == 0 ? 12 : 6);

            if (o.Papel == 3 && n == 1 && tamanhos.Count > 0)
            {
                double iw = tamanhos[0][0], ih = tamanhos[0][1];
                double s = (841.89 - 2 * m) / Math.Max(iw, ih);
                var lp = new LayoutPagina { W = iw * s + 2 * m, H = ih * s + 2 * m };
                lp.Celulas.Add(new Ret(m, m, iw * s, ih * s));
                return lp;
            }

            double[] papel = TamanhoPapel(o.Papel == 3 ? 0 : o.Papel);
            var retrato = Montar(papel[0], papel[1], n, m, esp, false);
            var paisagem = Montar(papel[1], papel[0], n, m, esp, true);
            if (o.Orientacao == 1) return retrato;
            if (o.Orientacao == 2) return paisagem;
            return Pontuar(paisagem, tamanhos, o.Preencher) > Pontuar(retrato, tamanhos, o.Preencher) * 1.02 ? paisagem : retrato;
        }

        static LayoutPagina Montar(double pw, double ph, int n, double m, double esp, bool paisagem)
        {
            int cols, rows;
            switch (n)
            {
                case 1: cols = 1; rows = 1; break;
                case 2: cols = paisagem ? 2 : 1; rows = paisagem ? 1 : 2; break;
                case 4: cols = 2; rows = 2; break;
                case 6: cols = paisagem ? 3 : 2; rows = paisagem ? 2 : 3; break;
                default:
                    cols = (int)Math.Ceiling(Math.Sqrt(n));
                    rows = (int)Math.Ceiling((double)n / cols);
                    if (!paisagem) { int t = cols; cols = rows; rows = t; }
                    break;
            }
            double cw = (pw - 2 * m - (cols - 1) * esp) / cols;
            double ch = (ph - 2 * m - (rows - 1) * esp) / rows;
            var lp = new LayoutPagina { W = pw, H = ph };
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                    lp.Celulas.Add(new Ret(m + c * (cw + esp), ph - m - (r + 1) * ch - r * esp, cw, ch));
            return lp;
        }

        // Encaixar: área ocupada pelas imagens. Preencher: quanto da imagem continua visível.
        static double Pontuar(LayoutPagina lp, IList<double[]> tamanhos, bool preencher)
        {
            double total = 0;
            for (int i = 0; i < tamanhos.Count && i < lp.Celulas.Count; i++)
            {
                var c = lp.Celulas[i];
                double iw = tamanhos[i][0], ih = tamanhos[i][1];
                if (preencher)
                {
                    double s = Math.Max(c.W / iw, c.H / ih);
                    total += (c.W * c.H) / (iw * ih * s * s);
                }
                else
                {
                    double s = Math.Min(c.W / iw, c.H / ih);
                    total += iw * ih * s * s;
                }
            }
            return total;
        }

        public static Ret Encaixe(Ret c, double iw, double ih, bool preencher)
        {
            double s = preencher ? Math.Max(c.W / iw, c.H / ih) : Math.Min(c.W / iw, c.H / ih);
            double w = iw * s, h = ih * s;
            return new Ret(c.X + (c.W - w) / 2, c.Y + (c.H - h) / 2, w, h);
        }
    }

    // Monta o arquivo PDF final
    class PdfWriter
    {
        public static readonly Encoding Latin1 = Encoding.GetEncoding(28591);
        readonly List<byte[]> objs = new List<byte[]>();
        public readonly List<int> Paginas = new List<int>();
        public readonly int RaizPaginas;

        public PdfWriter() { RaizPaginas = Reservar(); }

        public int Reservar() { objs.Add(null); return objs.Count; }
        public void Definir(int n, byte[] dados) { objs[n - 1] = dados; }
        public int Adicionar(string s) { objs.Add(Bytes(s)); return objs.Count; }

        public int AdicionarStream(string dict, byte[] dados)
        {
            var cab = Bytes("<< " + dict + " /Length " + dados.Length + " >>\nstream\n");
            var fim = Bytes("\nendstream");
            var b = new byte[cab.Length + dados.Length + fim.Length];
            Buffer.BlockCopy(cab, 0, b, 0, cab.Length);
            Buffer.BlockCopy(dados, 0, b, cab.Length, dados.Length);
            Buffer.BlockCopy(fim, 0, b, cab.Length + dados.Length, fim.Length);
            objs.Add(b);
            return objs.Count;
        }

        public static byte[] Bytes(string s) { return Latin1.GetBytes(s); }
        public static string N(double v) { return v.ToString("0.###", CultureInfo.InvariantCulture); }

        public void AdicionarPagina(double w, double h, string recursos, string conteudo)
        {
            int c = AdicionarStream("", Bytes(conteudo));
            Paginas.Add(Adicionar("<< /Type /Page /Parent " + RaizPaginas + " 0 R /MediaBox [0 0 " + N(w) + " " + N(h) +
                "] /Resources << " + recursos + " >> /Contents " + c + " 0 R >>"));
        }

        public void Salvar(string caminho)
        {
            Definir(RaizPaginas, Bytes("<< /Type /Pages /Count " + Paginas.Count + " /Kids [" +
                string.Join(" ", Paginas.Select(p => p + " 0 R")) + "] >>"));
            int cat = Adicionar("<< /Type /Catalog /Pages " + RaizPaginas + " 0 R >>");
            int info = Adicionar("<< /Producer (Conversor PDF) /CreationDate (D:" + DateTime.Now.ToString("yyyyMMddHHmmss") + ") >>");

            using (var fs = new FileStream(caminho, FileMode.Create, FileAccess.Write))
            using (var bs = new BufferedStream(fs, 1 << 16))
            {
                long pos = 0;
                Action<byte[]> W = b => { bs.Write(b, 0, b.Length); pos += b.Length; };
                W(Bytes("%PDF-1.7\n%âãÏÓ\n"));
                var offsets = new long[objs.Count];
                for (int i = 0; i < objs.Count; i++)
                {
                    offsets[i] = pos;
                    W(Bytes((i + 1) + " 0 obj\n"));
                    W(objs[i] ?? Bytes("null"));
                    W(Bytes("\nendobj\n"));
                }
                long xref = pos;
                var sb = new StringBuilder();
                sb.Append("xref\n0 " + (objs.Count + 1) + "\n0000000000 65535 f \n");
                foreach (var o in offsets) sb.Append(o.ToString("D10") + " 00000 n \n");
                sb.Append("trailer\n<< /Size " + (objs.Count + 1) + " /Root " + cat + " 0 R /Info " + info + " 0 R >>\nstartxref\n" + xref + "\n%%EOF\n");
                W(Bytes(sb.ToString()));
            }
        }
    }

    static class GeradorImagens
    {
        // Coloca as imagens nas páginas conforme as opções
        public static void Adicionar(PdfWriter pdf, List<FonteImagem> fontes, Opcoes o, Action aoProcessar, Action<FonteImagem, Exception> aoErrar)
        {
            int n = Math.Max(1, o.PorPagina);
            for (int i = 0; i < fontes.Count; i += n)
            {
                var imgs = new List<ImagemPdf>();
                foreach (var f in fontes.Skip(i).Take(n))
                {
                    try { imgs.Add(Imaging.Preparar(f, o)); }
                    catch (Exception ex) { if (aoErrar == null) throw; aoErrar(f, ex); }
                    if (aoProcessar != null) aoProcessar();
                }
                if (imgs.Count == 0) continue;
                var lp = Layout.Calcular(o, imgs.Select(x => new double[] { x.W, x.H }).ToList());
                var conteudo = new StringBuilder();
                var rec = new StringBuilder("/XObject << ");
                for (int k = 0; k < imgs.Count; k++)
                {
                    var im = imgs[k];
                    int obj = pdf.AdicionarStream("/Type /XObject /Subtype /Image /Width " + im.W + " /Height " + im.H +
                        " /ColorSpace " + im.Cor + " /BitsPerComponent 8 /Filter " + im.Filtro, im.Dados);
                    rec.Append("/Im" + k + " " + obj + " 0 R ");
                    var cel = lp.Celulas[k];
                    var r = Layout.Encaixe(cel, im.W, im.H, o.Preencher);
                    conteudo.Append("q ");
                    if (o.Preencher)
                        conteudo.Append(PdfWriter.N(cel.X) + " " + PdfWriter.N(cel.Y) + " " + PdfWriter.N(cel.W) + " " + PdfWriter.N(cel.H) + " re W n ");
                    conteudo.Append(PdfWriter.N(r.W) + " 0 0 " + PdfWriter.N(r.H) + " " + PdfWriter.N(r.X) + " " + PdfWriter.N(r.Y) + " cm /Im" + k + " Do Q\n");
                }
                rec.Append(">>");
                pdf.AdicionarPagina(lp.W, lp.H, rec.ToString(), conteudo.ToString());
            }
        }
    }
}
