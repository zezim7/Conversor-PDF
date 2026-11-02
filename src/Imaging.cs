using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using WI = System.Windows.Media.Imaging;
using WM = System.Windows.Media;

namespace ConversorPDF
{
    // Imagem pronta para entrar no PDF
    class ImagemPdf
    {
        public byte[] Dados;
        public int W, H;
        public string Filtro;   // /DCTDecode (JPEG) ou /FlateDecode (sem perda)
        public string Cor;      // /DeviceRGB ou /DeviceGray
    }

    static class Imaging
    {
        static readonly string[] SemPerda = { ".png", ".bmp", ".gif", ".tif", ".tiff", ".ico" };
        static readonly string[] Jpegs = { ".jpg", ".jpeg", ".jpe", ".jfif" };

        static FileStream Abrir(string c)
        {
            return new FileStream(c, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        }

        public static int ContarQuadros(string caminho)
        {
            string ext = Path.GetExtension(caminho).ToLowerInvariant();
            if (ext != ".tif" && ext != ".tiff") return 1;
            try
            {
                using (var fs = Abrir(caminho))
                {
                    var dec = WI.BitmapDecoder.Create(fs, WI.BitmapCreateOptions.DelayCreation, WI.BitmapCacheOption.None);
                    return Math.Max(1, dec.Frames.Count);
                }
            }
            catch { return 1; }
        }

        public static int LerOrientacao(string caminho)
        {
            try
            {
                using (var fs = Abrir(caminho))
                {
                    var dec = WI.BitmapDecoder.Create(fs, WI.BitmapCreateOptions.DelayCreation, WI.BitmapCacheOption.None);
                    return Orientacao(dec.Frames[0]);
                }
            }
            catch { return 1; }
        }

        static int Orientacao(WI.BitmapFrame fr)
        {
            try
            {
                var meta = fr.Metadata as WI.BitmapMetadata;
                if (meta == null) return 1;
                object o = meta.GetQuery("System.Photo.Orientation");
                int v = o == null ? 1 : Convert.ToInt32(o);
                return v >= 1 && v <= 8 ? v : 1;
            }
            catch { return 1; }
        }

        // Carrega a imagem na orientação correta, sobre fundo branco, em 24 bits.
        // maxLadoRapido > 0 reduz já na decodificação (rápido, usado em miniaturas).
        // largOrig/altOrig: tamanho real da imagem (já considerando a orientação EXIF).
        public static Bitmap Carregar(string caminho, int quadro, int maxLadoRapido, out int largOrig, out int altOrig)
        {
            Bitmap bmp;
            int orient;
            try
            {
                bmp = CarregarWic(caminho, quadro, maxLadoRapido, out orient, out largOrig, out altOrig);
            }
            catch
            {
                bmp = CarregarGdi(caminho, quadro, out orient);
                largOrig = bmp.Width; altOrig = bmp.Height;
                if (maxLadoRapido > 0)
                {
                    var r = Redimensionar(bmp, maxLadoRapido);
                    if (r != bmp) { bmp.Dispose(); bmp = r; }
                }
            }
            AplicarOrientacao(bmp, orient);
            if (orient >= 5) { int t = largOrig; largOrig = altOrig; altOrig = t; }
            return bmp;
        }

        static Bitmap CarregarWic(string caminho, int quadro, int maxLado, out int orient, out int w0, out int h0)
        {
            using (var fs = Abrir(caminho))
            {
                var dec = WI.BitmapDecoder.Create(fs, WI.BitmapCreateOptions.None, WI.BitmapCacheOption.OnLoad);
                WI.BitmapFrame fr = dec.Frames[Math.Min(quadro, dec.Frames.Count - 1)];
                orient = Orientacao(fr);
                w0 = fr.PixelWidth; h0 = fr.PixelHeight;
                WI.BitmapSource src = fr;
                if (maxLado > 0 && Math.Max(w0, h0) > maxLado)
                {
                    double s = (double)maxLado / Math.Max(w0, h0);
                    src = new WI.TransformedBitmap(src, new WM.ScaleTransform(s, s));
                }
                var conv = new WI.FormatConvertedBitmap(src, WM.PixelFormats.Bgra32, null, 0);
                int w = conv.PixelWidth, h = conv.PixelHeight;
                using (var argb = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                {
                    var d = argb.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    try { conv.CopyPixels(System.Windows.Int32Rect.Empty, d.Scan0, d.Stride * h, d.Stride); }
                    finally { argb.UnlockBits(d); }
                    return SobreBranco(argb);
                }
            }
        }

        static Bitmap CarregarGdi(string caminho, int quadro, out int orient)
        {
            using (var fs = Abrir(caminho))
            using (var img = Image.FromStream(fs))
            {
                orient = 1;
                try
                {
                    if (quadro > 0 && img.GetFrameCount(FrameDimension.Page) > quadro)
                        img.SelectActiveFrame(FrameDimension.Page, quadro);
                }
                catch { }
                if (img.PropertyIdList.Contains(0x0112))
                {
                    var v = img.GetPropertyItem(0x0112).Value;
                    if (v != null && v.Length > 0 && v[0] >= 1 && v[0] <= 8) orient = v[0];
                }
                return SobreBranco(img);
            }
        }

        static Bitmap SobreBranco(Image img)
        {
            var bmp = new Bitmap(img.Width, img.Height, PixelFormat.Format24bppRgb);
            bmp.SetResolution(96, 96);
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.White);
                g.DrawImage(img, new Rectangle(0, 0, img.Width, img.Height));
            }
            return bmp;
        }

        static void AplicarOrientacao(Bitmap bmp, int orient)
        {
            switch (orient)
            {
                case 2: bmp.RotateFlip(RotateFlipType.RotateNoneFlipX); break;
                case 3: bmp.RotateFlip(RotateFlipType.Rotate180FlipNone); break;
                case 4: bmp.RotateFlip(RotateFlipType.Rotate180FlipX); break;
                case 5: bmp.RotateFlip(RotateFlipType.Rotate90FlipX); break;
                case 6: bmp.RotateFlip(RotateFlipType.Rotate90FlipNone); break;
                case 7: bmp.RotateFlip(RotateFlipType.Rotate270FlipX); break;
                case 8: bmp.RotateFlip(RotateFlipType.Rotate270FlipNone); break;
            }
        }

        public static void Rotacionar(Bitmap bmp, int graus)
        {
            switch (((graus % 360) + 360) % 360)
            {
                case 90: bmp.RotateFlip(RotateFlipType.Rotate90FlipNone); break;
                case 180: bmp.RotateFlip(RotateFlipType.Rotate180FlipNone); break;
                case 270: bmp.RotateFlip(RotateFlipType.Rotate270FlipNone); break;
            }
        }

        // Devolve a mesma imagem se já for pequena o bastante
        public static Bitmap Redimensionar(Bitmap bmp, int maxLado)
        {
            if (maxLado <= 0 || Math.Max(bmp.Width, bmp.Height) <= maxLado) return bmp;
            double s = (double)maxLado / Math.Max(bmp.Width, bmp.Height);
            int w = Math.Max(1, (int)Math.Round(bmp.Width * s)), h = Math.Max(1, (int)Math.Round(bmp.Height * s));
            var r = new Bitmap(w, h, PixelFormat.Format24bppRgb);
            r.SetResolution(96, 96);
            using (var g = Graphics.FromImage(r))
            using (var at = new ImageAttributes())
            {
                at.SetWrapMode(WrapMode.TileFlipXY);
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(bmp, new Rectangle(0, 0, w, h), 0, 0, bmp.Width, bmp.Height, GraphicsUnit.Pixel, at);
            }
            return r;
        }

        // "Modo scanner": remove sombras e deixa o fundo do papel branco
        public static void MelhorarDocumento(Bitmap bmp)
        {
            int w = bmp.Width, h = bmp.Height;
            var d = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            int stride = d.Stride;
            byte[] px = new byte[stride * h];
            Marshal.Copy(d.Scan0, px, 0, px.Length);

            byte[] g = new byte[w * h];
            for (int y = 0; y < h; y++)
            {
                int o = y * stride;
                for (int x = 0; x < w; x++, o += 3)
                    g[y * w + x] = (byte)((px[o] * 114 + px[o + 1] * 587 + px[o + 2] * 299) / 1000);
            }

            // Estima o brilho do papel em blocos
            int B = Math.Max(16, Math.Min(w, h) / 24);
            int gw = (w + B - 1) / B, gh = (h + B - 1) / B;
            float[] fundo = new float[gw * gh];
            int[] hist = new int[256];
            for (int by = 0; by < gh; by++)
                for (int bx = 0; bx < gw; bx++)
                {
                    Array.Clear(hist, 0, 256);
                    int n = 0;
                    int x1 = Math.Min(w, (bx + 1) * B), y1 = Math.Min(h, (by + 1) * B);
                    for (int y = by * B; y < y1; y += 2)
                        for (int x = bx * B; x < x1; x += 2) { hist[g[y * w + x]]++; n++; }
                    int alvo = (int)(n * 0.9), acc = 0, v = 255;
                    for (int i = 0; i < 256; i++) { acc += hist[i]; if (acc >= alvo) { v = i; break; } }
                    fundo[by * gw + bx] = v;
                }
            fundo = Filtro3x3(fundo, gw, gh, true);
            fundo = Filtro3x3(fundo, gw, gh, false);

            // Curva: papel vira branco, texto fica mais escuro
            byte[] curva = new byte[1025];
            for (int i = 0; i <= 1024; i++)
            {
                double v = i / 1024.0 * 1.25;
                curva[i] = v >= 0.88 ? (byte)255 : (byte)(Math.Pow(v / 0.88, 1.7) * 255);
            }

            for (int y = 0; y < h; y++)
            {
                float fy = (y + 0.5f) / B - 0.5f;
                int y0 = (int)Math.Floor(fy);
                float ty = fy - y0;
                int ya = Limitar(y0, 0, gh - 1), yb = Limitar(y0 + 1, 0, gh - 1);
                int o = y * stride;
                for (int x = 0; x < w; x++, o += 3)
                {
                    float fx = (x + 0.5f) / B - 0.5f;
                    int x0 = (int)Math.Floor(fx);
                    float tx = fx - x0;
                    int xa = Limitar(x0, 0, gw - 1), xb = Limitar(x0 + 1, 0, gw - 1);
                    float f = (fundo[ya * gw + xa] * (1 - tx) + fundo[ya * gw + xb] * tx) * (1 - ty)
                            + (fundo[yb * gw + xa] * (1 - tx) + fundo[yb * gw + xb] * tx) * ty;
                    if (f < 60) f = 60;
                    float v = g[y * w + x] / f;
                    int idx = (int)(v / 1.25f * 1024);
                    byte r = curva[Limitar(idx, 0, 1024)];
                    px[o] = px[o + 1] = px[o + 2] = r;
                }
            }
            Marshal.Copy(px, 0, d.Scan0, px.Length);
            bmp.UnlockBits(d);
        }

        static int Limitar(int v, int a, int b) { return v < a ? a : (v > b ? b : v); }

        static float[] Filtro3x3(float[] a, int w, int h, bool maximo)
        {
            var r = new float[a.Length];
            for (int y = 0; y < h; y++)
                for (int x = 0; x < w; x++)
                {
                    float acc = maximo ? 0 : 0; int n = 0;
                    for (int dy = -1; dy <= 1; dy++)
                        for (int dx = -1; dx <= 1; dx++)
                        {
                            int xx = x + dx, yy = y + dy;
                            if (xx < 0 || yy < 0 || xx >= w || yy >= h) continue;
                            float v = a[yy * w + xx];
                            if (maximo) acc = Math.Max(acc, v); else acc += v;
                            n++;
                        }
                    r[y * w + x] = maximo ? acc : acc / n;
                }
            return r;
        }

        // Converte uma imagem da lista no formato que vai dentro do PDF
        public static ImagemPdf Preparar(FonteImagem f, Opcoes o)
        {
            string ext = Path.GetExtension(f.Caminho).ToLowerInvariant();
            bool jpeg = Jpegs.Contains(ext);
            int maxLado = o.Qualidade == 0 ? 0 : (o.Qualidade == 1 ? 2400 : 1600);
            int q = o.Qualidade == 0 ? 90 : (o.Qualidade == 1 ? 82 : 70);

            // JPEG sem alterações: usa o arquivo original (sem perder qualidade)
            if (jpeg && f.Quadro == 0 && f.Rotacao % 360 == 0 && !o.Documento && LerOrientacao(f.Caminho) == 1)
            {
                var orig = JpegOriginal(f.Caminho);
                if (orig != null && (maxLado == 0 || Math.Max(orig.W, orig.H) <= maxLado * 1.15)) return orig;
            }

            int w0, h0;
            Bitmap bmp = Carregar(f.Caminho, f.Quadro, maxLado > 0 ? maxLado * 2 : 0, out w0, out h0);
            try
            {
                Rotacionar(bmp, f.Rotacao);
                var menor = Redimensionar(bmp, maxLado);
                if (menor != bmp) { bmp.Dispose(); bmp = menor; }
                if (o.Documento) MelhorarDocumento(bmp);
                bool semPerda = o.Qualidade == 0 && !o.Documento && SemPerda.Contains(ext);
                return semPerda ? CodificarFlate(bmp) : CodificarJpeg(bmp, q);
            }
            finally { bmp.Dispose(); }
        }

        static ImagemPdf JpegOriginal(string caminho)
        {
            byte[] b;
            try { b = File.ReadAllBytes(caminho); } catch { return null; }
            if (b.Length < 4 || b[0] != 0xFF || b[1] != 0xD8) return null;
            int p = 2;
            while (p + 9 < b.Length)
            {
                if (b[p] != 0xFF) return null;
                byte m = b[p + 1];
                if (m == 0xFF) { p++; continue; }
                if (m == 0xD8 || m == 0x01 || (m >= 0xD0 && m <= 0xD7)) { p += 2; continue; }
                int len = (b[p + 2] << 8) | b[p + 3];
                if (m >= 0xC0 && m <= 0xCF && m != 0xC4 && m != 0xC8 && m != 0xCC)
                {
                    if (m != 0xC0 && m != 0xC1 && m != 0xC2) return null;
                    int h = (b[p + 5] << 8) | b[p + 6], w = (b[p + 7] << 8) | b[p + 8], nc = b[p + 9];
                    if ((nc != 1 && nc != 3) || w == 0 || h == 0) return null;
                    return new ImagemPdf { Dados = b, W = w, H = h, Filtro = "/DCTDecode", Cor = nc == 1 ? "/DeviceGray" : "/DeviceRGB" };
                }
                p += 2 + len;
            }
            return null;
        }

        static ImagemPdf CodificarJpeg(Bitmap bmp, int qualidade)
        {
            using (var ms = new MemoryStream())
            {
                var codec = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
                var ps = new EncoderParameters(1);
                ps.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)qualidade);
                bmp.Save(ms, codec, ps);
                return new ImagemPdf { Dados = ms.ToArray(), W = bmp.Width, H = bmp.Height, Filtro = "/DCTDecode", Cor = "/DeviceRGB" };
            }
        }

        static ImagemPdf CodificarFlate(Bitmap bmp)
        {
            int w = bmp.Width, h = bmp.Height;
            var d = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
            byte[] linha = new byte[d.Stride];
            byte[] raw = new byte[w * 3 * h];
            for (int y = 0; y < h; y++)
            {
                Marshal.Copy(d.Scan0 + y * d.Stride, linha, 0, d.Stride);
                int o = y * w * 3;
                for (int x = 0; x < w; x++)
                {
                    raw[o + x * 3] = linha[x * 3 + 2];
                    raw[o + x * 3 + 1] = linha[x * 3 + 1];
                    raw[o + x * 3 + 2] = linha[x * 3];
                }
            }
            bmp.UnlockBits(d);
            return new ImagemPdf { Dados = Zlib(raw), W = w, H = h, Filtro = "/FlateDecode", Cor = "/DeviceRGB" };
        }

        public static byte[] Zlib(byte[] raw)
        {
            using (var ms = new MemoryStream())
            {
                ms.WriteByte(0x78); ms.WriteByte(0xDA);
                using (var ds = new DeflateStream(ms, CompressionLevel.Optimal, true))
                    ds.Write(raw, 0, raw.Length);
                uint a = 1, b = 0;
                int i = 0;
                while (i < raw.Length)
                {
                    int fim = Math.Min(raw.Length, i + 5552);
                    for (; i < fim; i++) { a += raw[i]; b += a; }
                    a %= 65521; b %= 65521;
                }
                uint ad = (b << 16) | a;
                ms.WriteByte((byte)(ad >> 24)); ms.WriteByte((byte)(ad >> 16));
                ms.WriteByte((byte)(ad >> 8)); ms.WriteByte((byte)ad);
                return ms.ToArray();
            }
        }
    }
}
