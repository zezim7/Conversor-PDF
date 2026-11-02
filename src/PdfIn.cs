using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace ConversorPDF
{
    class PdfErro : Exception { public PdfErro(string m) : base(m) { } }

    // Modelo mínimo de objetos PDF
    abstract class PObj { }
    class PNum : PObj
    {
        public string V;
        public PNum(string v) { V = v; }
        public double D { get { double d; double.TryParse(V, NumberStyles.Float, CultureInfo.InvariantCulture, out d); return d; } }
        public int I { get { return (int)D; } }
    }
    class PName : PObj { public string V; public PName(string v) { V = v; } }
    class PStr : PObj { public byte[] Raw; public PStr(byte[] r) { Raw = r; } }
    class PKw : PObj
    {
        public string V;
        public PKw(string v) { V = v; }
        public static readonly PKw Nulo = new PKw("null");
    }
    class PRef : PObj { public int N, G; public PRef(int n, int g) { N = n; G = g; } }
    class PArr : PObj { public List<PObj> L = new List<PObj>(); }
    class PStream : PObj { public PDict D; public byte[] Dados; }
    class PDict : PObj
    {
        public List<KeyValuePair<string, PObj>> L = new List<KeyValuePair<string, PObj>>();
        public PObj this[string k]
        {
            get { foreach (var kv in L) if (kv.Key == k) return kv.Value; return null; }
            set
            {
                for (int i = 0; i < L.Count; i++)
                    if (L[i].Key == k)
                    {
                        if (value == null) L.RemoveAt(i); else L[i] = new KeyValuePair<string, PObj>(k, value);
                        return;
                    }
                if (value != null) L.Add(new KeyValuePair<string, PObj>(k, value));
            }
        }
        public PDict Clone() { var d = new PDict(); d.L.AddRange(L); return d; }
    }

    class Lexer
    {
        readonly byte[] b;
        public int P;
        public Lexer(byte[] b, int p) { this.b = b; P = p; }

        public static bool Branco(byte c) { return c == 0 || c == 9 || c == 10 || c == 12 || c == 13 || c == 32; }
        public static bool Delim(byte c) { return c == '(' || c == ')' || c == '<' || c == '>' || c == '[' || c == ']' || c == '{' || c == '}' || c == '/' || c == '%'; }

        public void Pular()
        {
            while (P < b.Length)
            {
                if (Branco(b[P])) P++;
                else if (b[P] == '%') { while (P < b.Length && b[P] != 10 && b[P] != 13) P++; }
                else break;
            }
        }

        string Regular()
        {
            int s = P;
            while (P < b.Length && !Branco(b[P]) && !Delim(b[P])) P++;
            return PdfWriter.Latin1.GetString(b, s, P - s);
        }

        byte[] Fatia(int s, int e)
        {
            e = Math.Min(e, b.Length);
            var r = new byte[Math.Max(0, e - s)];
            Buffer.BlockCopy(b, s, r, 0, r.Length);
            return r;
        }

        static bool Inteiro(string s)
        {
            if (s.Length == 0 || s.Length > 9) return false;
            foreach (char c in s) if (c < '0' || c > '9') return false;
            return true;
        }

        public PObj Ler()
        {
            Pular();
            if (P >= b.Length) throw new PdfErro("fim inesperado do arquivo");
            byte c = b[P];
            if (c == '/') { P++; return new PName("/" + Regular()); }
            if (c == '(') return LerString();
            if (c == '<')
            {
                if (P + 1 < b.Length && b[P + 1] == '<') { P += 2; return LerDict(); }
                int s = P;
                while (P < b.Length && b[P] != '>') P++;
                P++;
                return new PStr(Fatia(s, P));
            }
            if (c == '[')
            {
                P++;
                var a = new PArr();
                while (true)
                {
                    Pular();
                    if (P >= b.Length) throw new PdfErro("lista incompleta");
                    if (b[P] == ']') { P++; return a; }
                    a.L.Add(Ler());
                }
            }
            if ((c >= '0' && c <= '9') || c == '-' || c == '+' || c == '.')
            {
                string n = Regular();
                if (n.Length == 0) { P++; return PKw.Nulo; }
                if (Inteiro(n))
                {
                    int salvo = P;
                    Pular();
                    if (P < b.Length && b[P] >= '0' && b[P] <= '9')
                    {
                        string g = Regular();
                        Pular();
                        if (Inteiro(g) && P < b.Length && b[P] == 'R' && (P + 1 >= b.Length || Branco(b[P + 1]) || Delim(b[P + 1])))
                        {
                            P++;
                            return new PRef(int.Parse(n), int.Parse(g));
                        }
                    }
                    P = salvo;
                }
                return new PNum(n);
            }
            if (Delim(c)) { P++; return PKw.Nulo; }
            string k = Regular();
            if (k.Length == 0) { P++; return PKw.Nulo; }
            return new PKw(k);
        }

        PStr LerString()
        {
            int s = P, prof = 0;
            while (P < b.Length)
            {
                byte c = b[P];
                if (c == '\\') { P += 2; continue; }
                if (c == '(') prof++;
                else if (c == ')') { prof--; if (prof == 0) { P++; break; } }
                P++;
            }
            return new PStr(Fatia(s, P));
        }

        PDict LerDict()
        {
            var d = new PDict();
            while (true)
            {
                Pular();
                if (P + 1 >= b.Length) throw new PdfErro("dicionário incompleto");
                if (b[P] == '>' && b[P + 1] == '>') { P += 2; return d; }
                var nome = Ler() as PName;
                if (nome == null) continue;
                Pular();
                if (P + 1 < b.Length && b[P] == '>' && b[P + 1] == '>') { d[nome.V] = PKw.Nulo; continue; }
                d[nome.V] = Ler();
            }
        }
    }

    // Leitor de PDF (para juntar PDFs existentes)
    class PdfIn
    {
        readonly byte[] B;
        readonly Dictionary<int, long[]> xref = new Dictionary<int, long[]>();   // {tipo, a, b}
        readonly Dictionary<int, PObj> cache = new Dictionary<int, PObj>();
        readonly Dictionary<int, Tuple<byte[], int[], int[]>> objStms = new Dictionary<int, Tuple<byte[], int[], int[]>>();
        Dictionary<int, int> varredura;
        PDict trailer;

        public PdfIn(string caminho) : this(File.ReadAllBytes(caminho)) { }

        public PdfIn(byte[] b)
        {
            B = b;
            bool ok;
            try { ok = LerXrefs() && Raiz() != null; } catch { ok = false; }
            if (!ok) Reconstruir();
            if (trailer["/Encrypt"] != null)
                throw new PdfErro("Este PDF é protegido (criptografado). Abra-o no navegador, use Imprimir > Salvar como PDF e adicione o novo arquivo.");
        }

        public PObj Resolver(PObj o)
        {
            int n = 0;
            while (o is PRef && n++ < 32) o = Obter(((PRef)o).N);
            return o;
        }

        public static string Nome(PObj o) { var n = o as PName; return n == null ? null : n.V; }
        int Int(PObj o) { var n = Resolver(o) as PNum; return n == null ? 0 : n.I; }

        public PDict Raiz() { return Resolver(trailer["/Root"]) as PDict; }

        bool Comeca(int p, string s)
        {
            if (p < 0 || p + s.Length > B.Length) return false;
            for (int i = 0; i < s.Length; i++) if (B[p + i] != s[i]) return false;
            return true;
        }

        int UltimoIndice(string s, int antes)
        {
            for (int p = Math.Min(antes, B.Length - s.Length); p >= 0; p--)
                if (B[p] == s[0] && Comeca(p, s)) return p;
            return -1;
        }

        int Indice(string s, int desde)
        {
            for (int p = desde; p <= B.Length - s.Length; p++)
                if (B[p] == s[0] && Comeca(p, s)) return p;
            return -1;
        }

        bool LerXrefs()
        {
            int sx = UltimoIndice("startxref", B.Length);
            if (sx < 0) return false;
            var n = new Lexer(B, sx + 9).Ler() as PNum;
            if (n == null) return false;
            long off = (long)n.D;
            var vistos = new HashSet<long>();
            while (off > 0 && off < B.Length && vistos.Add(off))
            {
                PDict t = LerSecaoXref((int)off);
                if (t == null) return false;
                if (trailer == null) trailer = t;
                var xs = t["/XRefStm"] as PNum;
                if (xs != null && vistos.Add((long)xs.D)) LerSecaoXref((int)xs.D);
                var prev = t["/Prev"] as PNum;
                off = prev == null ? -1 : (long)prev.D;
            }
            return trailer != null && trailer["/Root"] != null;
        }

        PDict LerSecaoXref(int off)
        {
            var lx = new Lexer(B, off);
            lx.Pular();
            if (Comeca(lx.P, "xref"))
            {
                lx.P += 4;
                while (true)
                {
                    lx.Pular();
                    if (Comeca(lx.P, "trailer")) { lx.P += 7; return lx.Ler() as PDict; }
                    var a = lx.Ler() as PNum;
                    var c = lx.Ler() as PNum;
                    if (a == null || c == null) return null;
                    int ini = a.I, cnt = c.I;
                    for (int i = 0; i < cnt; i++)
                    {
                        var o = lx.Ler() as PNum;
                        lx.Ler();
                        var t = lx.Ler() as PKw;
                        if (o == null || t == null) return null;
                        int num = ini + i;
                        if (!xref.ContainsKey(num))
                            xref[num] = t.V == "n" ? new long[] { 1, (long)o.D, 0 } : new long[] { 0, 0, 0 };
                    }
                }
            }
            var s = LerObjetoEm(lx.P, -1) as PStream;
            if (s == null || Nome(s.D["/Type"]) != "/XRef") return null;
            byte[] dados = Decodificar(s);
            var w = Resolver(s.D["/W"]) as PArr;
            if (w == null || w.L.Count < 3) return null;
            int w0 = Int(w.L[0]), w1 = Int(w.L[1]), w2 = Int(w.L[2]), larg = w0 + w1 + w2;
            var faixas = new List<int>();
            var idx = Resolver(s.D["/Index"]) as PArr;
            if (idx != null) foreach (var x in idx.L) faixas.Add(Int(x));
            else { faixas.Add(0); faixas.Add(Int(s.D["/Size"])); }
            int p = 0;
            for (int f = 0; f + 1 < faixas.Count; f += 2)
                for (int i = 0; i < faixas[f + 1] && p + larg <= dados.Length; i++, p += larg)
                {
                    long t = w0 == 0 ? 1 : Campo(dados, p, w0);
                    long a = Campo(dados, p + w0, w1), c = Campo(dados, p + w0 + w1, w2);
                    int num = faixas[f] + i;
                    if (!xref.ContainsKey(num)) xref[num] = new long[] { t, a, c };
                }
            return s.D;
        }

        static long Campo(byte[] d, int p, int n)
        {
            long v = 0;
            for (int i = 0; i < n; i++) v = (v << 8) | d[p + i];
            return v;
        }

        PObj LerObjetoEm(int off, int esperado)
        {
            var lx = new Lexer(B, off);
            var n = lx.Ler() as PNum;
            lx.Ler();
            var k = lx.Ler() as PKw;
            if (n == null || k == null || k.V != "obj") throw new PdfErro("objeto inválido");
            if (esperado >= 0 && n.I != esperado) throw new PdfErro("tabela xref incorreta");
            PObj o = lx.Ler();
            var d = o as PDict;
            if (d == null) return o;
            lx.Pular();
            if (!Comeca(lx.P, "stream")) return o;
            int p = lx.P + 6;
            if (p < B.Length && B[p] == 13) p++;
            if (p < B.Length && B[p] == 10) p++;
            int len = -1;
            try { var lo = Resolver(d["/Length"]) as PNum; if (lo != null) len = lo.I; } catch { }
            if (len < 0 || p + len > B.Length || !EndstreamEm(p + len))
            {
                int e = Indice("endstream", p);
                if (e < 0) e = B.Length;
                if (e > p && B[e - 1] == 10) e--;
                if (e > p && B[e - 1] == 13) e--;
                len = e - p;
            }
            var dados = new byte[len];
            Buffer.BlockCopy(B, p, dados, 0, len);
            return new PStream { D = d, Dados = dados };
        }

        bool EndstreamEm(int p)
        {
            while (p < B.Length && Lexer.Branco(B[p])) p++;
            return Comeca(p, "endstream");
        }

        public bool Existe(int n) { return !ReferenceEquals(Obter(n), PKw.Nulo); }

        public PObj Obter(int n)
        {
            PObj o;
            if (cache.TryGetValue(n, out o)) return o;
            cache[n] = PKw.Nulo;
            o = PKw.Nulo;
            try
            {
                long[] e;
                if (xref.TryGetValue(n, out e))
                {
                    if (e[0] == 1) o = LerObjetoEm((int)e[1], n);
                    else if (e[0] == 2) o = DeObjStm((int)e[1], (int)e[2], n);
                }
                else o = PorVarredura(n);
            }
            catch
            {
                try { o = PorVarredura(n); } catch { o = PKw.Nulo; }
            }
            cache[n] = o;
            return o;
        }

        PObj PorVarredura(int n)
        {
            if (varredura == null) varredura = Varrer();
            int off;
            return varredura.TryGetValue(n, out off) ? LerObjetoEm(off, n) : PKw.Nulo;
        }

        PObj DeObjStm(int stm, int idx, int esperado)
        {
            Tuple<byte[], int[], int[]> t;
            if (!objStms.TryGetValue(stm, out t))
            {
                var s = Obter(stm) as PStream;
                if (s == null) throw new PdfErro("object stream ausente");
                byte[] d = Decodificar(s);
                int n = Int(s.D["/N"]), first = Int(s.D["/First"]);
                var lx = new Lexer(d, 0);
                var nums = new int[n];
                var offs = new int[n];
                for (int i = 0; i < n; i++)
                {
                    nums[i] = ((PNum)lx.Ler()).I;
                    offs[i] = first + ((PNum)lx.Ler()).I;
                }
                t = Tuple.Create(d, nums, offs);
                objStms[stm] = t;
            }
            int k = idx;
            if (k < 0 || k >= t.Item2.Length || t.Item2[k] != esperado) k = Array.IndexOf(t.Item2, esperado);
            if (k < 0) throw new PdfErro("objeto não encontrado");
            return new Lexer(t.Item1, t.Item3[k]).Ler();
        }

        // Procura "N G obj" no arquivo inteiro (para PDFs com tabela xref quebrada)
        Dictionary<int, int> Varrer()
        {
            var r = new Dictionary<int, int>();
            for (int p = 0; p + 3 <= B.Length; p++)
            {
                if (B[p] != 'o' || B[p + 1] != 'b' || B[p + 2] != 'j') continue;
                if (p + 3 < B.Length && !Lexer.Branco(B[p + 3]) && !Lexer.Delim(B[p + 3])) continue;
                int q = p - 1;
                while (q >= 0 && Lexer.Branco(B[q])) q--;
                int fimG = q;
                while (q >= 0 && B[q] >= '0' && B[q] <= '9') q--;
                if (q == fimG) continue;
                int ws = q;
                while (q >= 0 && Lexer.Branco(B[q])) q--;
                if (q == ws) continue;
                int fimN = q;
                while (q >= 0 && B[q] >= '0' && B[q] <= '9') q--;
                if (q == fimN || fimN - q > 9) continue;
                if (q >= 0 && !Lexer.Branco(B[q]) && !Lexer.Delim(B[q])) continue;
                int num = int.Parse(PdfWriter.Latin1.GetString(B, q + 1, fimN - q));
                r[num] = q + 1;
            }
            return r;
        }

        void Reconstruir()
        {
            xref.Clear(); cache.Clear(); objStms.Clear();
            varredura = Varrer();
            foreach (var kv in varredura) xref[kv.Key] = new long[] { 1, kv.Value, 0 };
            foreach (var num in varredura.Keys.ToList())
            {
                var s = Obter(num) as PStream;
                if (s == null || Nome(s.D["/Type"]) != "/ObjStm") continue;
                try
                {
                    byte[] d = Decodificar(s);
                    var lx = new Lexer(d, 0);
                    int n = Int(s.D["/N"]);
                    for (int i = 0; i < n; i++)
                    {
                        int on = ((PNum)lx.Ler()).I;
                        lx.Ler();
                        if (!xref.ContainsKey(on)) xref[on] = new long[] { 2, num, i };
                    }
                }
                catch { }
            }
            trailer = null;
            int pos = B.Length;
            while ((pos = UltimoIndice("trailer", pos - 1)) >= 0)
            {
                try
                {
                    var t = new Lexer(B, pos + 7).Ler() as PDict;
                    if (t != null && t["/Root"] != null && Resolver(t["/Root"]) is PDict) { trailer = t; break; }
                }
                catch { }
            }
            if (trailer == null)
            {
                foreach (var num in xref.Keys.OrderByDescending(x => x).ToList())
                {
                    var o = Obter(num);
                    var d = o is PStream ? ((PStream)o).D : o as PDict;
                    if (d == null) continue;
                    if (Nome(d["/Type"]) == "/XRef" && d["/Root"] != null) { trailer = d; break; }
                    if (Nome(d["/Type"]) == "/Catalog")
                    {
                        trailer = new PDict();
                        trailer["/Root"] = new PRef(num, 0);
                        break;
                    }
                }
            }
            if (trailer == null) throw new PdfErro("Não foi possível ler este PDF (o arquivo parece danificado).");
        }

        public byte[] Decodificar(PStream s)
        {
            var f = Resolver(s.D["/Filter"]);
            var nomes = new List<string>();
            if (f is PName) nomes.Add(((PName)f).V);
            else if (f is PArr) foreach (var x in ((PArr)f).L) nomes.Add(Nome(Resolver(x)));
            var dp = Resolver(s.D["/DecodeParms"]);
            byte[] dados = s.Dados;
            for (int i = 0; i < nomes.Count; i++)
            {
                PDict parm = dp as PDict;
                if (dp is PArr) parm = i < ((PArr)dp).L.Count ? Resolver(((PArr)dp).L[i]) as PDict : null;
                if (nomes[i] == "/FlateDecode" || nomes[i] == "/Fl")
                {
                    dados = Inflar(dados);
                    if (parm != null) dados = Preditor(dados, parm);
                }
                else throw new PdfErro("filtro não suportado: " + nomes[i]);
            }
            return dados;
        }

        static byte[] Inflar(byte[] d)
        {
            int ini = (d.Length > 2 && (d[0] & 0x0F) == 8) ? 2 : 0;
            using (var ms = new MemoryStream(d, ini, d.Length - ini))
            using (var ds = new DeflateStream(ms, CompressionMode.Decompress))
            using (var o = new MemoryStream())
            {
                var buf = new byte[65536];
                try
                {
                    int n;
                    while ((n = ds.Read(buf, 0, buf.Length)) > 0) o.Write(buf, 0, n);
                }
                catch (InvalidDataException) { }
                return o.ToArray();
            }
        }

        byte[] Preditor(byte[] d, PDict p)
        {
            int pred = Int(p["/Predictor"]);
            if (pred < 10) return d;
            int cols = p["/Columns"] != null ? Int(p["/Columns"]) : 1;
            int cores = p["/Colors"] != null ? Int(p["/Colors"]) : 1;
            int bpc = p["/BitsPerComponent"] != null ? Int(p["/BitsPerComponent"]) : 8;
            int bpp = Math.Max(1, cores * bpc / 8);
            int rl = (cols * cores * bpc + 7) / 8;
            var outp = new MemoryStream();
            byte[] ant = new byte[rl], cur = new byte[rl];
            for (int pos = 0; pos + 1 + rl <= d.Length; pos += rl + 1)
            {
                int ft = d[pos];
                for (int i = 0; i < rl; i++)
                {
                    int raw = d[pos + 1 + i];
                    int a = i >= bpp ? cur[i - bpp] : 0, b = ant[i], c = i >= bpp ? ant[i - bpp] : 0;
                    int v;
                    switch (ft)
                    {
                        case 1: v = raw + a; break;
                        case 2: v = raw + b; break;
                        case 3: v = raw + ((a + b) >> 1); break;
                        case 4:
                            int pp = a + b - c, pa = Math.Abs(pp - a), pb = Math.Abs(pp - b), pc = Math.Abs(pp - c);
                            v = raw + ((pa <= pb && pa <= pc) ? a : (pb <= pc ? b : c));
                            break;
                        default: v = raw; break;
                    }
                    cur[i] = (byte)v;
                }
                outp.Write(cur, 0, rl);
                var tmp = ant; ant = cur; cur = tmp;
            }
            return outp.ToArray();
        }

        // Páginas na ordem, com atributos herdados já aplicados. Key = número do objeto (-1 se direto)
        public List<KeyValuePair<int, PDict>> Paginas()
        {
            var res = new List<KeyValuePair<int, PDict>>();
            var raiz = Raiz();
            if (raiz == null) throw new PdfErro("PDF sem catálogo de páginas.");
            Andar(raiz["/Pages"], new PDict(), new HashSet<int>(), res, 0);
            return res;
        }

        void Andar(PObj no, PDict herdado, HashSet<int> vistos, List<KeyValuePair<int, PDict>> res, int prof)
        {
            int num = -1;
            var r = no as PRef;
            if (r != null) { if (!vistos.Add(r.N)) return; num = r.N; }
            var d = Resolver(no) as PDict;
            if (d == null || prof > 64) return;
            var kids = Resolver(d["/Kids"]) as PArr;
            if (kids != null && Nome(d["/Type"]) != "/Page")
            {
                var h = herdado.Clone();
                foreach (var k in new[] { "/Resources", "/MediaBox", "/CropBox", "/Rotate" })
                    if (d[k] != null) h[k] = d[k];
                foreach (var kid in kids.L) Andar(kid, h, vistos, res, prof + 1);
            }
            else
            {
                var p = d.Clone();
                foreach (var kv in herdado.L) if (p[kv.Key] == null) p[kv.Key] = kv.Value;
                if (p["/MediaBox"] == null)
                {
                    var mb = new PArr();
                    foreach (var v in new[] { "0", "0", "612", "792" }) mb.L.Add(new PNum(v));
                    p["/MediaBox"] = mb;
                }
                res.Add(new KeyValuePair<int, PDict>(num, p));
            }
        }

        public int ContarPaginas() { return Paginas().Count; }
    }

    static class Importador
    {
        // Copia todas as páginas de um PDF para o documento de saída
        public static int Importar(PdfWriter w, PdfIn src, int rotacao)
        {
            var paginas = src.Paginas();
            if (paginas.Count == 0) throw new PdfErro("O PDF não tem páginas.");
            var mapa = new Dictionary<int, int>();
            var fila = new Queue<int>();
            foreach (var p in paginas)
                if (p.Key >= 0 && !mapa.ContainsKey(p.Key)) mapa[p.Key] = w.Reservar();

            Func<PObj, PObj> conv = null;
            conv = o =>
            {
                var r = o as PRef;
                if (r != null)
                {
                    int novo;
                    if (!mapa.TryGetValue(r.N, out novo))
                    {
                        if (!src.Existe(r.N)) return PKw.Nulo;
                        novo = w.Reservar();
                        mapa[r.N] = novo;
                        fila.Enqueue(r.N);
                    }
                    return new PRef(novo, 0);
                }
                var a = o as PArr;
                if (a != null)
                {
                    var na = new PArr();
                    foreach (var x in a.L) na.L.Add(conv(x));
                    return na;
                }
                var d = o as PDict;
                if (d != null)
                {
                    var nd = new PDict();
                    foreach (var kv in d.L) nd.L.Add(new KeyValuePair<string, PObj>(kv.Key, conv(kv.Value)));
                    return nd;
                }
                var s = o as PStream;
                if (s != null)
                {
                    var sd0 = s.D.Clone();
                    sd0["/Length"] = null;
                    var sd = (PDict)conv(sd0);
                    sd["/Length"] = new PNum(s.Dados.Length.ToString(CultureInfo.InvariantCulture));
                    return new PStream { D = sd, Dados = s.Dados };
                }
                return o;
            };

            var usados = new HashSet<int>();
            foreach (var p in paginas)
            {
                var d = p.Value.Clone();
                foreach (var k in new[] { "/Parent", "/B", "/StructParents", "/Thumb", "/PieceInfo", "/Metadata" }) d[k] = null;
                if (rotacao % 360 != 0)
                {
                    var ro = src.Resolver(d["/Rotate"]) as PNum;
                    int r0 = ro == null ? 0 : ro.I;
                    d["/Rotate"] = new PNum(((((r0 + rotacao) % 360) + 360) % 360).ToString(CultureInfo.InvariantCulture));
                }
                var nd = (PDict)conv(d);
                nd["/Parent"] = new PRef(w.RaizPaginas, 0);
                int num = (p.Key >= 0 && usados.Add(p.Key)) ? mapa[p.Key] : w.Reservar();
                w.Definir(num, Serializar(nd));
                w.Paginas.Add(num);
            }
            while (fila.Count > 0)
            {
                int n = fila.Dequeue();
                w.Definir(mapa[n], Serializar(conv(src.Obter(n))));
            }
            return paginas.Count;
        }

        public static byte[] Serializar(PObj o)
        {
            using (var ms = new MemoryStream())
            {
                Escrever(ms, o);
                return ms.ToArray();
            }
        }

        static void T(Stream s, string t) { var b = PdfWriter.Latin1.GetBytes(t); s.Write(b, 0, b.Length); }

        static void Escrever(Stream s, PObj o)
        {
            if (o is PNum) T(s, ((PNum)o).V);
            else if (o is PName) T(s, ((PName)o).V);
            else if (o is PStr) { var r = ((PStr)o).Raw; s.Write(r, 0, r.Length); }
            else if (o is PKw) T(s, ((PKw)o).V);
            else if (o is PRef) T(s, ((PRef)o).N + " 0 R");
            else if (o is PArr)
            {
                T(s, "[");
                bool primeiro = true;
                foreach (var x in ((PArr)o).L)
                {
                    if (!primeiro) T(s, " ");
                    Escrever(s, x);
                    primeiro = false;
                }
                T(s, "]");
            }
            else if (o is PDict)
            {
                T(s, "<<");
                foreach (var kv in ((PDict)o).L)
                {
                    T(s, kv.Key + " ");
                    Escrever(s, kv.Value);
                    T(s, " ");
                }
                T(s, ">>");
            }
            else if (o is PStream)
            {
                var st = (PStream)o;
                Escrever(s, st.D);
                T(s, "\nstream\n");
                s.Write(st.Dados, 0, st.Dados.Length);
                T(s, "\nendstream");
            }
            else T(s, "null");
        }
    }
}
