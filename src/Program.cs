using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ConversorPDF
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Principal(args));
            Conversor.LimparTemp();
        }
    }

    class Item
    {
        public string Caminho, Url, Erro;
        public Tipo Tipo;
        public int Rotacao, Paginas = 1, LargOrig, AltOrig, IndiceImagem;
        public long Tamanho;
        public Bitmap Mini, MiniDoc;
        public bool Carregado;

        public string Nome { get { return Tipo == Tipo.Site ? Url : Path.GetFileName(Caminho); } }
        public bool EhImagem { get { return Tipo == Tipo.Imagem; } }
        public int Quadros { get { return EhImagem ? Math.Max(1, Paginas) : 1; } }
    }

    class PainelDuplo : Panel
    {
        public PainelDuplo() { DoubleBuffered = true; ResizeRedraw = true; }
    }

    static class Config
    {
        static string Arquivo
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ConversorPDF", "config.ini"); }
        }

        public static Dictionary<string, string> Ler()
        {
            var d = new Dictionary<string, string>();
            try
            {
                foreach (var l in File.ReadAllLines(Arquivo))
                {
                    int i = l.IndexOf('=');
                    if (i > 0) d[l.Substring(0, i)] = l.Substring(i + 1);
                }
            }
            catch { }
            return d;
        }

        public static void Salvar(Dictionary<string, string> d)
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(Arquivo));
                File.WriteAllLines(Arquivo, d.Select(kv => kv.Key + "=" + kv.Value));
            }
            catch { }
        }
    }

    class Principal : Form
    {
        [DllImport("shlwapi.dll", CharSet = CharSet.Unicode)]
        static extern int StrCmpLogicalW(string a, string b);
        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);

        const string FormatoInterno = "ConversorPDF.Itens";
        static readonly Color Vermelho = Color.FromArgb(217, 48, 37);
        static readonly Color Fundo = Color.FromArgb(243, 244, 247);

        readonly List<Item> itens = new List<Item>();
        readonly ListView lista = new ListView();
        readonly ImageList imagens = new ImageList();
        readonly PainelDuplo previa = new PainelDuplo();
        readonly Label lblInfo = new Label(), lblVazio = new Label(), lblStatus = new Label(), lblResumo = new Label();
        readonly LinkLabel lnkResultado = new LinkLabel();
        readonly ProgressBar barra = new ProgressBar();
        readonly Button btnGerar = new Button();
        readonly BlockingCollection<Item> fila = new BlockingCollection<Item>();
        readonly Dictionary<string, string> cfg = Config.Ler();
        readonly string[] argumentos;
        ComboBox cbPorPagina, cbPapel, cbOrient, cbMargem, cbAjuste, cbQualidade;
        CheckBox chkScanner, chkAbrir;
        RadioButton rbJuntar, rbSeparados;
        ToolStrip ferramentas;
        Panel painelOpcoes;
        string ultimoResultado;
        bool ocupado, carregandoCfg;

        public Principal(string[] args)
        {
            argumentos = args ?? new string[0];
            Text = "Conversor PDF";
            Font = new Font("Segoe UI", 9f);
            Size = new Size(1140, 760);
            MinimumSize = new Size(920, 620);
            StartPosition = FormStartPosition.CenterScreen;
            KeyPreview = true;
            BackColor = Fundo;
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }

            MontarFerramentas();
            MontarLista();
            var direita = MontarDireita();
            MontarOpcoes();
            var rodape = MontarRodape();

            Controls.Add(lista);
            Controls.Add(direita);
            Controls.Add(painelOpcoes);
            Controls.Add(rodape);
            Controls.Add(ferramentas);

            CarregarConfig();
            KeyDown += TeclaPressionada;
            IniciarTrabalhador();
            AtualizarTudo();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (argumentos.Length > 0) AdicionarArquivos(argumentos, -1);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (ocupado && MessageBox.Show(this, "Um PDF ainda está sendo gerado. Deseja sair mesmo assim?", Text,
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) { e.Cancel = true; return; }
            SalvarConfig();
            fila.CompleteAdding();
            base.OnFormClosing(e);
        }

        // ------------------------------------------------------------------ montagem da tela

        void MontarFerramentas()
        {
            ferramentas = new ToolStrip
            {
                GripStyle = ToolStripGripStyle.Hidden,
                Padding = new Padding(8, 6, 8, 6),
                RenderMode = ToolStripRenderMode.System,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9.5f)
            };
            ferramentas.Items.Add(Botao("+ Adicionar arquivos", "Escolher arquivos (Ctrl+O)", (s, e) => EscolherArquivos(), true));
            ferramentas.Items.Add(Botao("Adicionar site", "Transformar uma página da internet em PDF", (s, e) => AdicionarSite(null), false));
            ferramentas.Items.Add(Botao("Colar", "Colar um print, imagem ou arquivos copiados (Ctrl+V)", (s, e) => Colar(), false));
            ferramentas.Items.Add(new ToolStripSeparator());
            ferramentas.Items.Add(Botao("⟲ Girar", "Girar para a esquerda", (s, e) => Girar(-90), false));
            ferramentas.Items.Add(Botao("Girar ⟳", "Girar para a direita", (s, e) => Girar(90), false));
            ferramentas.Items.Add(new ToolStripSeparator());
            ferramentas.Items.Add(Botao("▲", "Mover para antes", (s, e) => Mover(-1), false));
            ferramentas.Items.Add(Botao("▼", "Mover para depois", (s, e) => Mover(1), false));
            var ordenar = new ToolStripDropDownButton("Ordenar") { ToolTipText = "Mudar a ordem de todos os arquivos" };
            ordenar.DropDownItems.Add("Por nome (A → Z)", null, (s, e) => Ordenar((a, b) => StrCmpLogicalW(a.Nome, b.Nome)));
            ordenar.DropDownItems.Add("Por nome (Z → A)", null, (s, e) => Ordenar((a, b) => StrCmpLogicalW(b.Nome, a.Nome)));
            ordenar.DropDownItems.Add("Por data (mais antigo primeiro)", null, (s, e) => Ordenar((a, b) => DataDe(a).CompareTo(DataDe(b))));
            ordenar.DropDownItems.Add("Por data (mais recente primeiro)", null, (s, e) => Ordenar((a, b) => DataDe(b).CompareTo(DataDe(a))));
            ordenar.DropDownItems.Add(new ToolStripSeparator());
            ordenar.DropDownItems.Add("Inverter a ordem", null, (s, e) => { itens.Reverse(); Reconstruir(null); });
            ferramentas.Items.Add(ordenar);
            ferramentas.Items.Add(new ToolStripSeparator());
            ferramentas.Items.Add(Botao("Remover", "Remover os selecionados (Delete)", (s, e) => Remover(), false));
            ferramentas.Items.Add(Botao("Limpar", "Remover todos", (s, e) => { itens.Clear(); Reconstruir(null); }, false));

            var ajuda = new ToolStripDropDownButton("Ajuda") { Alignment = ToolStripItemAlignment.Right };
            ajuda.DropDownItems.Add("O que dá para converter?", null, (s, e) => MostrarFormatos());
            ajuda.DropDownItems.Add("Colocar no botão direito do Windows (\"Enviar para\")", null, (s, e) => CriarAtalhoEnviarPara());
            ajuda.DropDownItems.Add("Atalhos de teclado", null, (s, e) => MessageBox.Show(this,
                "Ctrl+O  adicionar arquivos\nCtrl+V  colar print, imagem, arquivos ou link\nCtrl+A  selecionar tudo\nDelete  remover selecionados\n" +
                "Ctrl+←/→  girar\nAlt+↑/↓  mudar a ordem\nCtrl+Enter  gerar o PDF\n\nTambém dá para arrastar as miniaturas para mudar a ordem.", Text));
            ferramentas.Items.Add(ajuda);
        }

        static ToolStripButton Botao(string texto, string dica, EventHandler h, bool destaque)
        {
            var b = new ToolStripButton(texto) { ToolTipText = dica, Padding = new Padding(4, 2, 4, 2) };
            if (destaque) b.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            b.Click += h;
            return b;
        }

        void MontarLista()
        {
            imagens.ColorDepth = ColorDepth.Depth32Bit;
            imagens.ImageSize = new Size(128, 128);
            lista.View = View.LargeIcon;
            lista.LargeImageList = imagens;
            lista.Dock = DockStyle.Fill;
            lista.AllowDrop = true;
            lista.HideSelection = false;
            lista.MultiSelect = true;
            lista.BorderStyle = BorderStyle.None;
            lista.BackColor = Color.White;
            lista.SelectedIndexChanged += (s, e) => { previa.Invalidate(); AtualizarInfo(); };
            lista.ItemDrag += (s, e) => lista.DoDragDrop(new DataObject(FormatoInterno, true), DragDropEffects.Move);
            lista.DragEnter += ArrastarSobre;
            lista.DragOver += ArrastarSobre;
            lista.DragLeave += (s, e) => lista.InsertionMark.Index = -1;
            lista.DragDrop += Soltar;
            lista.DoubleClick += (s, e) => AbrirSelecionado();
            lista.HandleCreated += (s, e) =>
                SendMessage(lista.Handle, 0x1000 + 53, IntPtr.Zero, (IntPtr)((170 << 16) | 156)); // LVM_SETICONSPACING

            var menu = new ContextMenuStrip();
            menu.Items.Add("Abrir arquivo", null, (s, e) => AbrirSelecionado());
            menu.Items.Add("Girar para a esquerda", null, (s, e) => Girar(-90));
            menu.Items.Add("Girar para a direita", null, (s, e) => Girar(90));
            menu.Items.Add("Mover para o início", null, (s, e) => MoverPara(true));
            menu.Items.Add("Mover para o fim", null, (s, e) => MoverPara(false));
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Remover", null, (s, e) => Remover());
            lista.ContextMenuStrip = menu;

            lblVazio.Text = "Arraste arquivos para cá\n\nFotos, prints, PDFs, Word, Excel, PowerPoint, textos, páginas web, e-mails...\n\n" +
                            "ou clique aqui para escolher  •  Ctrl+V cola um print";
            lblVazio.Font = new Font("Segoe UI", 12f);
            lblVazio.ForeColor = Color.FromArgb(120, 125, 135);
            lblVazio.BackColor = Color.White;
            lblVazio.TextAlign = ContentAlignment.MiddleCenter;
            lblVazio.Dock = DockStyle.Fill;
            lblVazio.Cursor = Cursors.Hand;
            lblVazio.AllowDrop = true;
            lblVazio.Click += (s, e) => EscolherArquivos();
            lblVazio.DragEnter += (s, e) => e.Effect = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            lblVazio.DragDrop += (s, e) => { var f = e.Data.GetData(DataFormats.FileDrop) as string[]; if (f != null) AdicionarArquivos(f, -1); };
            lista.Controls.Add(lblVazio);
        }

        Panel MontarDireita()
        {
            var direita = new Panel { Dock = DockStyle.Right, Width = 330, Padding = new Padding(14, 10, 14, 10), BackColor = Fundo };
            var titulo = new Label { Text = "Prévia da página", Dock = DockStyle.Top, Height = 26, Font = new Font("Segoe UI", 10f, FontStyle.Bold), ForeColor = Color.FromArgb(50, 55, 65) };
            previa.Dock = DockStyle.Fill;
            previa.BackColor = Fundo;
            previa.Paint += (s, e) => PintarPrevia(e.Graphics, previa.ClientRectangle);
            lblInfo.Dock = DockStyle.Bottom;
            lblInfo.Height = 128;
            lblInfo.ForeColor = Color.FromArgb(70, 75, 85);
            lblInfo.Padding = new Padding(0, 8, 0, 0);
            direita.Controls.Add(previa);
            direita.Controls.Add(lblInfo);
            direita.Controls.Add(titulo);
            return direita;
        }

        ComboBox Combo(params string[] opcoes)
        {
            var c = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, Margin = new Padding(0, 2, 14, 4) };
            c.Items.AddRange(opcoes);
            c.SelectedIndex = 0;
            c.SelectedIndexChanged += (s, e) => OpcaoMudou();
            return c;
        }

        static Label Rotulo(string t)
        {
            return new Label { Text = t, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 4), ForeColor = Color.FromArgb(60, 65, 75) };
        }

        void MontarOpcoes()
        {
            painelOpcoes = new Panel { Dock = DockStyle.Bottom, Height = 150, Padding = new Padding(12, 6, 12, 4), BackColor = Fundo };

            var gImg = new GroupBox { Text = "Como colocar as imagens nas páginas", Dock = DockStyle.Fill, Padding = new Padding(10, 6, 10, 6) };
            var t = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 6, RowCount = 3 };
            for (int i = 0; i < 3; i++)
            {
                t.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                t.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.3f));
            }
            cbPorPagina = Combo("1 por página", "2 por página", "4 por página", "6 por página");
            cbPapel = Combo("A4", "Carta", "Ofício", "Do tamanho da foto");
            cbOrient = Combo("Automática", "Em pé (retrato)", "Deitada (paisagem)");
            cbMargem = Combo("Normal", "Pequena", "Sem margem");
            cbAjuste = Combo("Imagem inteira", "Preencher (corta bordas)");
            cbQualidade = Combo("Alta (original)", "Média (recomendada)", "Pequena (WhatsApp / e-mail)");
            t.Controls.Add(Rotulo("Imagens:"), 0, 0); t.Controls.Add(cbPorPagina, 1, 0);
            t.Controls.Add(Rotulo("Papel:"), 2, 0); t.Controls.Add(cbPapel, 3, 0);
            t.Controls.Add(Rotulo("Orientação:"), 4, 0); t.Controls.Add(cbOrient, 5, 0);
            t.Controls.Add(Rotulo("Margem:"), 0, 1); t.Controls.Add(cbMargem, 1, 1);
            t.Controls.Add(Rotulo("Ajuste:"), 2, 1); t.Controls.Add(cbAjuste, 3, 1);
            t.Controls.Add(Rotulo("Qualidade:"), 4, 1); t.Controls.Add(cbQualidade, 5, 1);
            chkScanner = new CheckBox { Text = "Modo scanner — para fotos de documentos: deixa o fundo branco e tira sombras", AutoSize = true, Margin = new Padding(0, 4, 0, 0) };
            chkScanner.CheckedChanged += (s, e) => OpcaoMudou();
            t.Controls.Add(chkScanner, 0, 2);
            t.SetColumnSpan(chkScanner, 6);
            gImg.Controls.Add(t);

            var gSaida = new GroupBox { Text = "Resultado", Dock = DockStyle.Right, Width = 300, Padding = new Padding(10, 6, 10, 6) };
            var f = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false };
            rbJuntar = new RadioButton { Text = "Juntar tudo em um único PDF", AutoSize = true, Checked = true, Margin = new Padding(0, 4, 0, 2) };
            rbSeparados = new RadioButton { Text = "Um PDF separado para cada arquivo", AutoSize = true, Margin = new Padding(0, 2, 0, 2) };
            chkAbrir = new CheckBox { Text = "Abrir o PDF quando terminar", AutoSize = true, Checked = true, Margin = new Padding(0, 8, 0, 0) };
            rbJuntar.CheckedChanged += (s, e) => OpcaoMudou();
            chkAbrir.CheckedChanged += (s, e) => OpcaoMudou();
            f.Controls.AddRange(new Control[] { rbJuntar, rbSeparados, chkAbrir });
            gSaida.Controls.Add(f);

            var espaco = new Panel { Dock = DockStyle.Right, Width = 12 };
            painelOpcoes.Controls.Add(gImg);
            painelOpcoes.Controls.Add(espaco);
            painelOpcoes.Controls.Add(gSaida);
        }

        Panel MontarRodape()
        {
            var rodape = new Panel { Dock = DockStyle.Bottom, Height = 66, Padding = new Padding(12, 8, 12, 12), BackColor = Fundo };
            btnGerar.Text = "Gerar PDF";
            btnGerar.Dock = DockStyle.Right;
            btnGerar.Width = 230;
            btnGerar.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
            btnGerar.BackColor = Vermelho;
            btnGerar.ForeColor = Color.White;
            btnGerar.FlatStyle = FlatStyle.Flat;
            btnGerar.FlatAppearance.BorderSize = 0;
            btnGerar.Cursor = Cursors.Hand;
            btnGerar.Click += (s, e) => Gerar();

            var meio = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0, 0, 16, 0) };
            lblResumo.Dock = DockStyle.Top;
            lblResumo.Height = 22;
            lblResumo.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            lblResumo.ForeColor = Color.FromArgb(50, 55, 65);
            barra.Dock = DockStyle.Bottom;
            barra.Height = 8;
            barra.Maximum = 1000;
            barra.Visible = false;
            lblStatus.Dock = DockStyle.Fill;
            lblStatus.ForeColor = Color.FromArgb(90, 95, 105);
            lblStatus.AutoEllipsis = true;
            lnkResultado.Dock = DockStyle.Right;
            lnkResultado.AutoSize = true;
            lnkResultado.Visible = false;
            lnkResultado.Text = "Mostrar na pasta";
            lnkResultado.Padding = new Padding(0, 3, 0, 0);
            lnkResultado.LinkClicked += (s, e) => MostrarNaPasta(ultimoResultado);
            meio.Controls.Add(lblStatus);
            meio.Controls.Add(lnkResultado);
            meio.Controls.Add(barra);
            meio.Controls.Add(lblResumo);
            rodape.Controls.Add(meio);
            rodape.Controls.Add(btnGerar);
            return rodape;
        }

        // ------------------------------------------------------------------ configuração

        void CarregarConfig()
        {
            carregandoCfg = true;
            Func<string, int, int> I = (k, p) => { string v; int r; return cfg.TryGetValue(k, out v) && int.TryParse(v, out r) ? r : p; };
            Action<ComboBox, string, int> C = (c, k, p) => { int v = I(k, p); c.SelectedIndex = v >= 0 && v < c.Items.Count ? v : p; };
            C(cbPorPagina, "porPagina", 0);
            C(cbPapel, "papel", 0);
            C(cbOrient, "orientacao", 0);
            C(cbMargem, "margem", 0);
            C(cbAjuste, "ajuste", 0);
            C(cbQualidade, "qualidade", 1);
            chkScanner.Checked = I("scanner", 0) == 1;
            rbSeparados.Checked = I("separados", 0) == 1;
            rbJuntar.Checked = !rbSeparados.Checked;
            chkAbrir.Checked = I("abrir", 1) == 1;
            carregandoCfg = false;
        }

        void SalvarConfig()
        {
            cfg["porPagina"] = cbPorPagina.SelectedIndex.ToString();
            cfg["papel"] = cbPapel.SelectedIndex.ToString();
            cfg["orientacao"] = cbOrient.SelectedIndex.ToString();
            cfg["margem"] = cbMargem.SelectedIndex.ToString();
            cfg["ajuste"] = cbAjuste.SelectedIndex.ToString();
            cfg["qualidade"] = cbQualidade.SelectedIndex.ToString();
            cfg["scanner"] = chkScanner.Checked ? "1" : "0";
            cfg["separados"] = rbSeparados.Checked ? "1" : "0";
            cfg["abrir"] = chkAbrir.Checked ? "1" : "0";
            Config.Salvar(cfg);
        }

        void OpcaoMudou()
        {
            if (carregandoCfg) return;
            previa.Invalidate();
            AtualizarResumo();
        }

        Opcoes LerOpcoes()
        {
            return new Opcoes
            {
                PorPagina = new[] { 1, 2, 4, 6 }[Math.Max(0, cbPorPagina.SelectedIndex)],
                Papel = cbPapel.SelectedIndex,
                Orientacao = cbOrient.SelectedIndex,
                Margem = cbMargem.SelectedIndex,
                Preencher = cbAjuste.SelectedIndex == 1,
                Qualidade = cbQualidade.SelectedIndex,
                Documento = chkScanner.Checked
            };
        }

        // ------------------------------------------------------------------ adicionar itens

        void EscolherArquivos()
        {
            using (var d = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Escolha os arquivos para transformar em PDF",
                Filter = "Todos os formatos aceitos|*.jpg;*.jpeg;*.jfif;*.png;*.bmp;*.gif;*.tif;*.tiff;*.webp;*.heic;*.heif;*.avif;*.ico;" +
                         "*.pdf;*.doc;*.docx;*.rtf;*.odt;*.xls;*.xlsx;*.ods;*.csv;*.ppt;*.pptx;*.odp;*.txt;*.md;*.json;*.xml;*.log;*.html;*.htm;*.mht;*.svg;*.eml|" +
                         "Imagens|*.jpg;*.jpeg;*.jfif;*.png;*.bmp;*.gif;*.tif;*.tiff;*.webp;*.heic;*.heif;*.avif;*.ico|" +
                         "PDF|*.pdf|Documentos do Office|*.doc;*.docx;*.rtf;*.odt;*.xls;*.xlsx;*.ods;*.csv;*.ppt;*.pptx;*.odp|" +
                         "Textos e páginas web|*.txt;*.md;*.json;*.xml;*.log;*.html;*.htm;*.mht;*.svg|E-mails|*.eml|Todos os arquivos|*.*"
            })
            {
                string pasta;
                if (cfg.TryGetValue("pastaEntrada", out pasta) && Directory.Exists(pasta)) d.InitialDirectory = pasta;
                if (d.ShowDialog(this) != DialogResult.OK) return;
                cfg["pastaEntrada"] = Path.GetDirectoryName(d.FileNames[0]);
                AdicionarArquivos(d.FileNames, -1);
            }
        }

        void AdicionarArquivos(IEnumerable<string> caminhos, int posicao)
        {
            var novos = new List<Item>();
            var ignorados = new List<string>();
            foreach (var c0 in caminhos)
            {
                string c = c0.Trim('"');
                if (Directory.Exists(c))
                {
                    var arqs = Directory.GetFiles(c).Where(a => Conversor.Detectar(a) != Tipo.Desconhecido).ToList();
                    arqs.Sort(StrCmpLogicalW);
                    foreach (var a in arqs) novos.Add(NovoItem(a));
                    continue;
                }
                if (!File.Exists(c)) continue;
                if (Conversor.Detectar(c) == Tipo.Desconhecido) { ignorados.Add(Path.GetFileName(c)); continue; }
                novos.Add(NovoItem(c));
            }
            if (novos.Count > 0) Inserir(novos, posicao);
            if (ignorados.Count > 0)
                MessageBox.Show(this, "Estes arquivos não são de um formato aceito:\n\n" + string.Join("\n", ignorados.Take(15)) +
                    (ignorados.Count > 15 ? "\n..." : "") + "\n\nVeja em Ajuda > O que dá para converter?", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        Item NovoItem(string c)
        {
            var it = new Item { Caminho = Path.GetFullPath(c), Tipo = Conversor.Detectar(c) };
            try { it.Tamanho = new FileInfo(c).Length; } catch { }
            return it;
        }

        void Inserir(List<Item> novos, int posicao)
        {
            if (posicao < 0 || posicao > itens.Count) posicao = itens.Count;
            itens.InsertRange(posicao, novos);
            foreach (var it in novos)
            {
                it.IndiceImagem = imagens.Images.Count;
                imagens.Images.Add(ComporTile(it));
                fila.Add(it);
            }
            Reconstruir(null);
            var ultimo = lista.Items.Cast<ListViewItem>().FirstOrDefault(x => x.Tag == novos[novos.Count - 1]);
            if (ultimo != null) ultimo.EnsureVisible();
            lnkResultado.Visible = false;
        }

        void AdicionarSite(string url)
        {
            if (url == null)
            {
                using (var f = new Form
                {
                    Text = "Adicionar site",
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    StartPosition = FormStartPosition.CenterParent,
                    ClientSize = new Size(520, 130),
                    MinimizeBox = false,
                    MaximizeBox = false,
                    Font = Font
                })
                {
                    var l = new Label { Text = "Endereço da página (ex.: www.exemplo.com.br/noticia):", Location = new Point(14, 14), AutoSize = true };
                    var tb = new TextBox { Location = new Point(14, 38), Width = 490 };
                    try { if (Clipboard.ContainsText() && Clipboard.GetText().Trim().StartsWith("http")) tb.Text = Clipboard.GetText().Trim(); } catch { }
                    var ok = new Button { Text = "Adicionar", DialogResult = DialogResult.OK, Location = new Point(318, 82), Width = 90 };
                    var cancel = new Button { Text = "Cancelar", DialogResult = DialogResult.Cancel, Location = new Point(414, 82), Width = 90 };
                    f.Controls.AddRange(new Control[] { l, tb, ok, cancel });
                    f.AcceptButton = ok;
                    f.CancelButton = cancel;
                    if (f.ShowDialog(this) != DialogResult.OK) return;
                    url = tb.Text.Trim();
                }
            }
            if (url.Length == 0) return;
            if (!url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) && !url.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) url = "https://" + url;
            Uri u;
            if (!Uri.TryCreate(url, UriKind.Absolute, out u) || u.Host.IndexOf('.') < 0)
            {
                MessageBox.Show(this, "Esse endereço não parece válido.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var it = new Item { Url = u.AbsoluteUri, Tipo = Tipo.Site, Carregado = true, Paginas = 0 };
            Inserir(new List<Item> { it }, -1);
        }

        void Colar()
        {
            try
            {
                if (Clipboard.ContainsFileDropList())
                {
                    AdicionarArquivos(Clipboard.GetFileDropList().Cast<string>(), -1);
                    return;
                }
                if (Clipboard.ContainsImage())
                {
                    using (var img = Clipboard.GetImage())
                    {
                        string pasta = Path.Combine(Conversor.PastaTemp, "colados");
                        Directory.CreateDirectory(pasta);
                        string c = Path.Combine(pasta, "Print " + DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".png");
                        img.Save(c, ImageFormat.Png);
                        AdicionarArquivos(new[] { c }, -1);
                    }
                    return;
                }
                if (Clipboard.ContainsText())
                {
                    string t = Clipboard.GetText().Trim();
                    if (Regex(t)) { AdicionarSite(t); return; }
                    if (File.Exists(t) || Directory.Exists(t)) { AdicionarArquivos(new[] { t }, -1); return; }
                }
                MessageBox.Show(this, "Não há imagem, arquivo ou link copiado.\n\nDica: tire um print com Win+Shift+S e depois aperte Ctrl+V aqui.", Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Não foi possível colar: " + ex.Message, Text);
            }
        }

        static bool Regex(string t)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(t, @"^(https?://|www\.)\S+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        // ------------------------------------------------------------------ miniaturas (em segundo plano)

        void IniciarTrabalhador()
        {
            var t = new Thread(() =>
            {
                foreach (var it in fila.GetConsumingEnumerable())
                {
                    CarregarItem(it);
                    try { BeginInvoke((Action)(() => ItemCarregado(it))); } catch { }
                }
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        static void CarregarItem(Item it)
        {
            try
            {
                if (it.Tipo == Tipo.Imagem)
                {
                    it.Paginas = Imaging.ContarQuadros(it.Caminho);
                    int w, h;
                    it.Mini = Imaging.Carregar(it.Caminho, 0, 360, out w, out h);
                    it.LargOrig = w; it.AltOrig = h;
                }
                else if (it.Tipo == Tipo.Pdf)
                {
                    it.Paginas = new PdfIn(it.Caminho).ContarPaginas();
                }
                else it.Paginas = 0;
            }
            catch (Exception ex)
            {
                it.Erro = it.Tipo == Tipo.Imagem
                    ? "Não foi possível abrir a imagem. " + (Path.GetExtension(it.Caminho).ToLowerInvariant().StartsWith(".he")
                        ? "Para fotos HEIC do iPhone, instale a \"Extensão de Imagem HEIF\" na Microsoft Store." : ex.Message)
                    : ex.Message;
            }
            it.Carregado = true;
        }

        void ItemCarregado(Item it)
        {
            if (!itens.Contains(it)) return;
            it.MiniDoc = null;
            imagens.Images[it.IndiceImagem] = ComporTile(it);
            lista.Invalidate();
            previa.Invalidate();
            AtualizarInfo();
            AtualizarResumo();
        }

        static Color CorTipo(Tipo t)
        {
            switch (t)
            {
                case Tipo.Pdf: return Color.FromArgb(217, 48, 37);
                case Tipo.Word: return Color.FromArgb(43, 87, 154);
                case Tipo.Excel: case Tipo.Csv: return Color.FromArgb(33, 115, 70);
                case Tipo.PowerPoint: return Color.FromArgb(210, 71, 38);
                case Tipo.Web: return Color.FromArgb(107, 72, 200);
                case Tipo.Email: return Color.FromArgb(0, 137, 123);
                case Tipo.Site: return Color.FromArgb(26, 115, 232);
                default: return Color.FromArgb(95, 99, 104);
            }
        }

        static Bitmap Girada(Bitmap b, int rot)
        {
            var r = (Bitmap)b.Clone();
            Imaging.Rotacionar(r, rot);
            return r;
        }

        static string Etiqueta(Item it)
        {
            if (it.Tipo == Tipo.Site) return "WWW";
            string e = Path.GetExtension(it.Caminho).TrimStart('.').ToUpperInvariant();
            return e.Length > 5 ? e.Substring(0, 5) : e;
        }

        Bitmap ComporTile(Item it)
        {
            var bmp = new Bitmap(128, 128, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                if (it.Mini != null)
                {
                    using (var m = Girada(it.Mini, it.Rotacao))
                    {
                        double s = Math.Min(118.0 / m.Width, 118.0 / m.Height);
                        int w = Math.Max(1, (int)(m.Width * s)), h = Math.Max(1, (int)(m.Height * s));
                        var r = new Rectangle((128 - w) / 2, (128 - h) / 2, w, h);
                        using (var sombra = new SolidBrush(Color.FromArgb(40, 0, 0, 0))) g.FillRectangle(sombra, r.X + 2, r.Y + 3, r.Width, r.Height);
                        g.DrawImage(m, r);
                        using (var p = new Pen(Color.FromArgb(60, 0, 0, 0))) g.DrawRectangle(p, r.X, r.Y, r.Width - 1, r.Height - 1);
                    }
                }
                else
                {
                    Color cor = CorTipo(it.Tipo);
                    var folha = new Rectangle(26, 6, 76, 104);
                    using (var path = new GraphicsPath())
                    {
                        path.AddLines(new[] { new Point(folha.Left, folha.Top), new Point(folha.Right - 20, folha.Top), new Point(folha.Right, folha.Top + 20),
                                              new Point(folha.Right, folha.Bottom), new Point(folha.Left, folha.Bottom) });
                        path.CloseFigure();
                        using (var sombra = new SolidBrush(Color.FromArgb(35, 0, 0, 0))) { g.TranslateTransform(2, 3); g.FillPath(sombra, path); g.ResetTransform(); }
                        g.FillPath(Brushes.White, path);
                        using (var p = new Pen(Color.FromArgb(190, 195, 205))) g.DrawPath(p, path);
                    }
                    using (var dobra = new SolidBrush(Color.FromArgb(225, 228, 234)))
                        g.FillPolygon(dobra, new[] { new Point(folha.Right - 20, folha.Top), new Point(folha.Right - 20, folha.Top + 20), new Point(folha.Right, folha.Top + 20) });
                    using (var linhas = new Pen(Color.FromArgb(215, 218, 225), 3))
                        for (int y = 34; y <= 58; y += 8) g.DrawLine(linhas, folha.Left + 10, y, folha.Right - 12, y);
                    var faixa = new Rectangle(14, 70, 100, 30);
                    using (var b = new SolidBrush(cor)) g.FillRectangle(b, faixa);
                    using (var f = new Font("Segoe UI", 13f, FontStyle.Bold))
                    using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        g.DrawString(Etiqueta(it), f, Brushes.White, faixa, sf);
                }
                // selos
                string selo = null;
                Color corSelo = Color.FromArgb(60, 64, 72);
                if (it.Erro != null) { selo = "erro"; corSelo = Vermelho; }
                else if (!it.Carregado) selo = "...";
                else if (it.Paginas > 1) selo = it.Paginas + " págs";
                else if (it.Rotacao != 0 && it.Mini == null) selo = "↻ " + it.Rotacao + "°";
                if (selo != null)
                    using (var f = new Font("Segoe UI", 8f, FontStyle.Bold))
                    {
                        var tam = g.MeasureString(selo, f);
                        var r = new RectangleF(128 - tam.Width - 10, 128 - tam.Height - 4, tam.Width + 8, tam.Height + 2);
                        using (var b = new SolidBrush(corSelo)) g.FillRectangle(b, r);
                        g.DrawString(selo, f, Brushes.White, r.X + 4, r.Y + 1);
                    }
            }
            return bmp;
        }

        // ------------------------------------------------------------------ lista: ordem, seleção, arrastar

        void Reconstruir(IList<Item> selecionar)
        {
            var sel = selecionar ?? Selecionados();
            lista.BeginUpdate();
            lista.Items.Clear();
            for (int i = 0; i < itens.Count; i++)
            {
                var it = itens[i];
                var lvi = new ListViewItem((i + 1) + ". " + it.Nome) { ImageIndex = it.IndiceImagem, Tag = it, ToolTipText = it.Caminho ?? it.Url };
                lista.Items.Add(lvi);
                if (sel.Contains(it)) lvi.Selected = true;
            }
            lista.EndUpdate();
            if (sel.Count > 0)
            {
                var primeiro = lista.Items.Cast<ListViewItem>().FirstOrDefault(x => x.Selected);
                if (primeiro != null) { primeiro.Focused = true; primeiro.EnsureVisible(); }
            }
            AtualizarTudo();
        }

        void AtualizarTudo()
        {
            lblVazio.Visible = itens.Count == 0;
            previa.Invalidate();
            AtualizarInfo();
            AtualizarResumo();
        }

        List<Item> Selecionados()
        {
            return lista.SelectedItems.Cast<ListViewItem>().Select(x => (Item)x.Tag).ToList();
        }

        void Remover()
        {
            var sel = Selecionados();
            if (sel.Count == 0) return;
            int primeiro = itens.IndexOf(sel[0]);
            itens.RemoveAll(sel.Contains);
            var proximo = itens.Count == 0 ? new List<Item>() : new List<Item> { itens[Math.Min(primeiro, itens.Count - 1)] };
            Reconstruir(proximo);
        }

        void Girar(int graus)
        {
            var sel = Selecionados();
            if (sel.Count == 0)
            {
                if (itens.Count == 0) return;
                MessageBox.Show(this, "Selecione primeiro o que deseja girar.", Text);
                return;
            }
            foreach (var it in sel)
            {
                it.Rotacao = ((it.Rotacao + graus) % 360 + 360) % 360;
                it.MiniDoc = null;
                imagens.Images[it.IndiceImagem] = ComporTile(it);
            }
            lista.Invalidate();
            previa.Invalidate();
            AtualizarInfo();
        }

        void Mover(int dir)
        {
            var sel = Selecionados();
            if (sel.Count == 0) return;
            var idx = sel.Select(s => itens.IndexOf(s)).OrderBy(i => i).ToList();
            if (dir < 0 && idx[0] == 0) return;
            if (dir > 0 && idx[idx.Count - 1] == itens.Count - 1) return;
            if (dir > 0) idx.Reverse();
            foreach (int i in idx)
            {
                var t = itens[i];
                itens[i] = itens[i + dir];
                itens[i + dir] = t;
            }
            Reconstruir(sel);
        }

        void MoverPara(bool inicio)
        {
            var sel = Selecionados();
            if (sel.Count == 0) return;
            var ordem = sel.OrderBy(itens.IndexOf).ToList();
            itens.RemoveAll(sel.Contains);
            itens.InsertRange(inicio ? 0 : itens.Count, ordem);
            Reconstruir(sel);
        }

        void Ordenar(Comparison<Item> comp)
        {
            var ordenada = itens.ToList();
            ordenada.Sort(comp);
            itens.Clear();
            itens.AddRange(ordenada);
            Reconstruir(null);
        }

        static DateTime DataDe(Item it)
        {
            try { return it.Caminho == null ? DateTime.MaxValue : File.GetLastWriteTime(it.Caminho); } catch { return DateTime.MaxValue; }
        }

        int PosicaoSoltura()
        {
            int i = lista.InsertionMark.Index;
            if (i < 0) return -1;
            return lista.InsertionMark.AppearsAfterItem ? i + 1 : i;
        }

        void ArrastarSobre(object sender, DragEventArgs e)
        {
            bool interno = e.Data.GetDataPresent(FormatoInterno);
            bool arquivos = e.Data.GetDataPresent(DataFormats.FileDrop);
            if (!interno && !arquivos) { e.Effect = DragDropEffects.None; return; }
            e.Effect = interno ? DragDropEffects.Move : DragDropEffects.Copy;
            var pt = lista.PointToClient(new Point(e.X, e.Y));
            int i = lista.InsertionMark.NearestIndex(pt);
            if (i >= 0)
            {
                var r = lista.GetItemRect(i);
                lista.InsertionMark.AppearsAfterItem = pt.X > r.Left + r.Width / 2;
            }
            lista.InsertionMark.Index = i;
        }

        void Soltar(object sender, DragEventArgs e)
        {
            int pos = PosicaoSoltura();
            lista.InsertionMark.Index = -1;
            if (e.Data.GetDataPresent(FormatoInterno))
            {
                var sel = Selecionados();
                if (sel.Count == 0) return;
                if (pos < 0) pos = itens.Count;
                int antes = sel.Count(s => itens.IndexOf(s) < pos);
                var ordem = sel.OrderBy(itens.IndexOf).ToList();
                itens.RemoveAll(sel.Contains);
                itens.InsertRange(Math.Max(0, Math.Min(itens.Count, pos - antes)), ordem);
                Reconstruir(sel);
            }
            else
            {
                var f = e.Data.GetData(DataFormats.FileDrop) as string[];
                if (f != null) AdicionarArquivos(f, pos);
            }
        }

        void AbrirSelecionado()
        {
            var sel = Selecionados().FirstOrDefault();
            if (sel == null) return;
            try { Process.Start(sel.Tipo == Tipo.Site ? sel.Url : sel.Caminho); } catch { }
        }

        void TeclaPressionada(object sender, KeyEventArgs e)
        {
            if (ocupado) return;
            if (e.KeyCode == Keys.Delete && lista.Focused) { Remover(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.V) { Colar(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.O) { EscolherArquivos(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.A) { foreach (ListViewItem i in lista.Items) i.Selected = true; e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.Left) { Girar(-90); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.Right) { Girar(90); e.Handled = true; }
            else if (e.Alt && e.KeyCode == Keys.Up) { Mover(-1); e.Handled = true; }
            else if (e.Alt && e.KeyCode == Keys.Down) { Mover(1); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.Enter) { Gerar(); e.Handled = true; }
        }

        // ------------------------------------------------------------------ prévia e informações

        Bitmap MiniParaPrevia(Item it, bool documento)
        {
            if (it.Mini == null) return null;
            if (!documento) return Girada(it.Mini, it.Rotacao);
            if (it.MiniDoc == null)
            {
                var c = (Bitmap)it.Mini.Clone();
                try { Imaging.MelhorarDocumento(c); } catch { }
                it.MiniDoc = c;
            }
            return Girada(it.MiniDoc, it.Rotacao);
        }

        void PintarPrevia(Graphics g, Rectangle area)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            var sel = Selecionados().FirstOrDefault() ?? itens.FirstOrDefault();
            if (sel == null)
            {
                TextRenderer.DrawText(g, "Adicione arquivos para ver\ncomo vai ficar o PDF.", Font, area, Color.Gray,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
                return;
            }
            if (!sel.EhImagem || sel.Mini == null)
            {
                using (var tile = ComporTile(sel))
                {
                    int s = Math.Min(200, Math.Min(area.Width, area.Height) - 20);
                    g.DrawImage(tile, area.X + (area.Width - s) / 2, area.Y + (area.Height - s) / 2 - 10, s, s);
                }
                return;
            }

            var o = LerOpcoes();
            // grupo de imagens consecutivas (cada quadro de TIFF conta como uma imagem)
            int idx = itens.IndexOf(sel);
            int ini = idx, fim = idx;
            if (!rbSeparados.Checked)
            {
                while (ini > 0 && itens[ini - 1].EhImagem) ini--;
                while (fim < itens.Count - 1 && itens[fim + 1].EhImagem) fim++;
            }
            var quadros = new List<Item>();
            int posSel = -1;
            for (int i = ini; i <= fim; i++)
            {
                if (itens[i] == sel) posSel = quadros.Count;
                for (int q = 0; q < itens[i].Quadros; q++) quadros.Add(itens[i]);
            }
            int n = o.PorPagina;
            int inicio = posSel - posSel % n;
            var grupo = quadros.Skip(inicio).Take(n).ToList();
            var tamanhos = grupo.Select(it =>
            {
                double w = Math.Max(1, it.Mini == null ? 1 : it.Mini.Width), h = Math.Max(1, it.Mini == null ? 1 : it.Mini.Height);
                return it.Rotacao % 180 != 0 ? new[] { h, w } : new[] { w, h };
            }).ToList();
            var lp = ConversorPDF.Layout.Calcular(o, tamanhos);

            double esc = Math.Min((area.Width - 24) / lp.W, (area.Height - 40) / lp.H);
            float pw = (float)(lp.W * esc), ph = (float)(lp.H * esc);
            float px = area.X + (area.Width - pw) / 2, py = area.Y + 6;
            using (var sombra = new SolidBrush(Color.FromArgb(45, 0, 0, 0))) g.FillRectangle(sombra, px + 3, py + 4, pw, ph);
            g.FillRectangle(Brushes.White, px, py, pw, ph);
            using (var p = new Pen(Color.FromArgb(200, 204, 212))) g.DrawRectangle(p, px, py, pw, ph);

            for (int k = 0; k < grupo.Count && k < lp.Celulas.Count; k++)
            {
                var c = lp.Celulas[k];
                var r = ConversorPDF.Layout.Encaixe(c, tamanhos[k][0], tamanhos[k][1], o.Preencher);
                var celTela = new RectangleF(px + (float)(c.X * esc), py + (float)((lp.H - c.Y - c.H) * esc), (float)(c.W * esc), (float)(c.H * esc));
                var imgTela = new RectangleF(px + (float)(r.X * esc), py + (float)((lp.H - r.Y - r.H) * esc), (float)(r.W * esc), (float)(r.H * esc));
                using (var m = MiniParaPrevia(grupo[k], o.Documento))
                {
                    var estado = g.Save();
                    g.SetClip(celTela);
                    if (m != null) g.DrawImage(m, imgTela);
                    g.Restore(estado);
                }
                if (grupo[k] == sel && grupo.Count > 1)
                    using (var p = new Pen(Color.FromArgb(26, 115, 232), 2)) g.DrawRectangle(p, imgTela.X, imgTela.Y, imgTela.Width, imgTela.Height);
            }
            string legenda = (lp.W > lp.H ? "Página deitada" : "Página em pé") + " • " + grupo.Count + (grupo.Count == 1 ? " imagem" : " imagens");
            TextRenderer.DrawText(g, legenda, Font, new Rectangle(area.X, (int)(py + ph + 8), area.Width, 20), Color.FromArgb(100, 105, 115), TextFormatFlags.HorizontalCenter);
        }

        string MotorOffice(Item it)
        {
            if (Conversor.TemMsOffice(it.Tipo)) return "Microsoft Office";
            if (Conversor.AcharLibreOffice() != null) return "LibreOffice";
            string e = Path.GetExtension(it.Caminho).ToLowerInvariant();
            if (new[] { ".docx", ".docm", ".dotx", ".xlsx", ".xlsm", ".xltx", ".pptx", ".ppsx", ".pptm", ".potx" }.Contains(e))
                return "conversor interno (layout aproximado; para ficar idêntico, instale o LibreOffice grátis)";
            return "precisa do Microsoft Office ou do LibreOffice (grátis) instalado";
        }

        void AtualizarInfo()
        {
            var sel = Selecionados();
            if (sel.Count == 0) { lblInfo.Text = itens.Count == 0 ? "" : "Clique em um arquivo para ver os detalhes."; return; }
            if (sel.Count > 1) { lblInfo.Text = sel.Count + " itens selecionados.\nUse os botões de girar e mover, ou arraste para reordenar."; return; }
            var it = sel[0];
            var l = new List<string> { it.Nome, Conversor.Descricao(it.Tipo) + (it.Tamanho > 0 ? " • " + Html.Tamanho(it.Tamanho) : "") };
            if (it.EhImagem && it.LargOrig > 0) l.Add(it.LargOrig + " × " + it.AltOrig + " pixels");
            if (it.Paginas > 1 || it.Tipo == Tipo.Pdf) l.Add(it.Paginas + (it.Paginas == 1 ? " página" : " páginas"));
            if (it.Rotacao != 0) l.Add("Girado " + it.Rotacao + "°");
            if (it.Tipo == Tipo.Word || it.Tipo == Tipo.Excel || it.Tipo == Tipo.PowerPoint) l.Add("Conversão: " + MotorOffice(it));
            if (it.Tipo == Tipo.Site) l.Add("A página será baixada e convertida na hora de gerar.");
            if (it.Erro != null) l.Add("⚠ " + it.Erro);
            lblInfo.Text = string.Join("\n", l);
        }

        void AtualizarResumo()
        {
            if (itens.Count == 0) { lblResumo.Text = "Nenhum arquivo adicionado"; btnGerar.Enabled = false; return; }
            btnGerar.Enabled = !ocupado;
            var o = LerOpcoes();
            int paginas = 0;
            bool incerto = false;
            int seq = 0;
            Action fechar = () => { paginas += (seq + o.PorPagina - 1) / o.PorPagina; seq = 0; };
            foreach (var it in itens)
            {
                if (it.EhImagem) { if (rbSeparados.Checked) paginas += (it.Quadros + o.PorPagina - 1) / o.PorPagina; else seq += it.Quadros; continue; }
                fechar();
                if (it.Tipo == Tipo.Pdf && it.Carregado && it.Erro == null) paginas += it.Paginas;
                else incerto = true;
            }
            fechar();
            string pags = incerto ? "pelo menos " + paginas : paginas.ToString();
            lblResumo.Text = itens.Count + (itens.Count == 1 ? " arquivo" : " arquivos") + " • " +
                             (rbSeparados.Checked ? itens.Count + " PDFs separados" : "1 PDF com " + pags + (paginas == 1 && !incerto ? " página" : " páginas"));
        }

        // ------------------------------------------------------------------ geração

        static string NomeBase(Item it)
        {
            string n;
            if (it.Tipo == Tipo.Site)
            {
                Uri u;
                n = Uri.TryCreate(it.Url, UriKind.Absolute, out u) ? (u.Host + u.AbsolutePath).Replace("www.", "") : "site";
            }
            else n = Path.GetFileNameWithoutExtension(it.Caminho);
            foreach (char c in Path.GetInvalidFileNameChars()) n = n.Replace(c, '-');
            n = n.Trim('-', ' ', '.');
            if (n.Length > 80) n = n.Substring(0, 80);
            return n.Length == 0 ? "documento" : n;
        }

        static string NomeUnico(string pasta, string nome)
        {
            string c = Path.Combine(pasta, nome);
            string b = Path.GetFileNameWithoutExtension(nome), e = Path.GetExtension(nome);
            for (int i = 2; File.Exists(c); i++) c = Path.Combine(pasta, b + " (" + i + ")" + e);
            return c;
        }

        void Gerar()
        {
            if (ocupado || itens.Count == 0) return;
            var pendentes = itens.Where(i => !i.Carregado).ToList();
            if (pendentes.Count > 0) { lblStatus.Text = "Aguarde, ainda carregando " + pendentes.Count + " arquivo(s)..."; return; }
            var o = LerOpcoes();
            bool separados = rbSeparados.Checked;
            string destino;
            string pastaSaida;
            cfg.TryGetValue("pastaSaida", out pastaSaida);
            if (pastaSaida == null || !Directory.Exists(pastaSaida))
                pastaSaida = itens[0].Caminho != null ? Path.GetDirectoryName(itens[0].Caminho) : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (pastaSaida.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase))
                pastaSaida = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            if (separados)
            {
                using (var d = new FolderBrowserDialog { Description = "Escolha a pasta onde salvar os PDFs", SelectedPath = pastaSaida })
                {
                    if (d.ShowDialog(this) != DialogResult.OK) return;
                    destino = d.SelectedPath;
                    cfg["pastaSaida"] = destino;
                }
            }
            else
            {
                string sugestao = itens.Count == 1 ? NomeBase(itens[0]) : NomeBase(itens[0]) + (itens.Count > 1 ? " e mais " + (itens.Count - 1) : "");
                using (var d = new SaveFileDialog { Filter = "PDF|*.pdf", FileName = sugestao + ".pdf", Title = "Salvar PDF", InitialDirectory = pastaSaida, OverwritePrompt = true })
                {
                    if (d.ShowDialog(this) != DialogResult.OK) return;
                    destino = d.FileName;
                    cfg["pastaSaida"] = Path.GetDirectoryName(destino);
                }
            }
            SalvarConfig();

            var copia = itens.ToList();
            bool abrir = chkAbrir.Checked;
            Ocupado(true);
            var t = new Thread(() => Processar(copia, o, separados, destino, abrir));
            t.SetApartmentState(ApartmentState.STA);
            t.IsBackground = true;
            t.Start();
        }

        void Ocupado(bool sim)
        {
            ocupado = sim;
            btnGerar.Enabled = !sim && itens.Count > 0;
            btnGerar.Text = sim ? "Gerando..." : "Gerar PDF";
            btnGerar.BackColor = sim ? Color.Gray : Vermelho;
            ferramentas.Enabled = !sim;
            painelOpcoes.Enabled = !sim;
            lista.AllowDrop = !sim;
            barra.Visible = sim;
            barra.Value = 0;
            if (sim) lnkResultado.Visible = false;
            Cursor = sim ? Cursors.AppStarting : Cursors.Default;
        }

        void Processar(List<Item> copia, Opcoes o, bool separados, string destino, bool abrir)
        {
            var erros = new List<string>();
            var gerados = new List<string>();
            int total = copia.Sum(i => i.Quadros), passo = 0;
            int paginasGeradas = 0;
            Action<string> status = s =>
            {
                int v = (int)(1000.0 * passo / Math.Max(1, total));
                try { BeginInvoke((Action)(() => { lblStatus.Text = s; barra.Value = Math.Max(0, Math.Min(1000, v)); })); } catch { }
            };
            Action<FonteImagem, Exception> erroImg = (f, ex) => erros.Add(Path.GetFileName(f.Caminho) + ": " + ex.Message);
            try
            {
                if (!separados)
                {
                    var pdf = new PdfWriter();
                    var pendentes = new List<FonteImagem>();
                    Action descarregar = () =>
                    {
                        if (pendentes.Count == 0) return;
                        GeradorImagens.Adicionar(pdf, pendentes, o, () => { passo++; status("Preparando imagens (" + passo + " de " + total + ")..."); }, erroImg);
                        pendentes.Clear();
                    };
                    foreach (var it in copia)
                    {
                        if (it.EhImagem)
                        {
                            if (it.Erro != null) { erros.Add(it.Nome + ": " + it.Erro); passo += it.Quadros; continue; }
                            for (int q = 0; q < it.Quadros; q++) pendentes.Add(new FonteImagem { Caminho = it.Caminho, Quadro = q, Rotacao = it.Rotacao });
                            continue;
                        }
                        descarregar();
                        status("Convertendo " + it.Nome + "...");
                        try { AdicionarDocumento(pdf, it, o); }
                        catch (Exception ex) { erros.Add(it.Nome + ": " + ex.Message); }
                        passo++;
                    }
                    descarregar();
                    if (pdf.Paginas.Count == 0) throw new Exception("Nenhuma página pôde ser gerada.");
                    status("Salvando...");
                    Salvar(pdf, destino);
                    paginasGeradas = pdf.Paginas.Count;
                    gerados.Add(destino);
                }
                else
                {
                    foreach (var it in copia)
                    {
                        status("Convertendo " + it.Nome + "...");
                        try
                        {
                            if (it.Erro != null && it.EhImagem) throw new Exception(it.Erro);
                            var pdf = new PdfWriter();
                            if (it.EhImagem)
                            {
                                var fs = Enumerable.Range(0, it.Quadros).Select(q => new FonteImagem { Caminho = it.Caminho, Quadro = q, Rotacao = it.Rotacao }).ToList();
                                GeradorImagens.Adicionar(pdf, fs, o, () => passo++, erroImg);
                            }
                            else { AdicionarDocumento(pdf, it, o); passo++; }
                            if (pdf.Paginas.Count == 0) continue;
                            string arq = NomeUnico(destino, NomeBase(it) + ".pdf");
                            Salvar(pdf, arq);
                            paginasGeradas += pdf.Paginas.Count;
                            gerados.Add(arq);
                        }
                        catch (Exception ex) { erros.Add(it.Nome + ": " + ex.Message); }
                    }
                    if (gerados.Count == 0) throw new Exception("Nenhum PDF pôde ser gerado.");
                }
                BeginInvoke((Action)(() => Concluido(gerados, erros, separados, destino, abrir, paginasGeradas)));
            }
            catch (Exception ex)
            {
                BeginInvoke((Action)(() => Falhou(ex.Message, erros)));
            }
        }

        static void AdicionarDocumento(PdfWriter pdf, Item it, Opcoes o)
        {
            string arq = it.Tipo == Tipo.Pdf ? it.Caminho : Conversor.ParaPdf(it.Caminho, it.Url, it.Tipo, o);
            Importador.Importar(pdf, new PdfIn(arq), it.Rotacao);
        }

        static void Salvar(PdfWriter pdf, string destino)
        {
            string temp = destino + ".tmp";
            try
            {
                pdf.Salvar(temp);
                File.Copy(temp, destino, true);
            }
            catch (IOException)
            {
                throw new Exception("Não foi possível salvar \"" + Path.GetFileName(destino) + "\". Se ele estiver aberto em outro programa, feche-o e tente de novo.");
            }
            catch (UnauthorizedAccessException)
            {
                throw new Exception("Sem permissão para salvar nessa pasta. Escolha outra pasta.");
            }
            finally
            {
                try { File.Delete(temp); } catch { }
            }
        }

        void Concluido(List<string> gerados, List<string> erros, bool separados, string destino, bool abrir, int paginas)
        {
            Ocupado(false);
            long bytes = gerados.Sum(g => { try { return new FileInfo(g).Length; } catch { return 0L; } });
            ultimoResultado = separados ? destino : gerados[0];
            lblStatus.Text = "✔ " + (separados ? gerados.Count + " PDFs criados" : "PDF criado: " + Path.GetFileName(gerados[0])) +
                             " (" + paginas + (paginas == 1 ? " página" : " páginas") + ", " + Html.Tamanho(bytes) + ")";
            lnkResultado.Visible = true;
            if (erros.Count > 0)
                MessageBox.Show(this, "O PDF foi criado, mas alguns arquivos tiveram problema e ficaram de fora:\n\n" +
                    string.Join("\n\n", erros.Take(10)) + (erros.Count > 10 ? "\n\n..." : ""), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            if (abrir)
            {
                try { Process.Start(separados ? destino : gerados[0]); } catch { }
            }
        }

        void Falhou(string msg, List<string> erros)
        {
            Ocupado(false);
            lblStatus.Text = "Não foi possível gerar o PDF.";
            MessageBox.Show(this, msg + (erros.Count > 0 ? "\n\nDetalhes:\n" + string.Join("\n\n", erros.Take(10)) : ""), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        static void MostrarNaPasta(string caminho)
        {
            if (caminho == null) return;
            try
            {
                if (Directory.Exists(caminho)) Process.Start("explorer.exe", "\"" + caminho + "\"");
                else Process.Start("explorer.exe", "/select,\"" + caminho + "\"");
            }
            catch { }
        }

        // ------------------------------------------------------------------ ajuda

        void MostrarFormatos()
        {
            string office = Conversor.TemMsOffice(Tipo.Word) ? "usando o Microsoft Office instalado" :
                            Conversor.AcharLibreOffice() != null ? "usando o LibreOffice instalado" :
                            ".docx, .xlsx e .pptx com o conversor interno (layout aproximado). Para .doc/.xls/.ppt antigos e .odt, instale o LibreOffice (grátis)";
            MessageBox.Show(this,
                "IMAGENS: JPG, PNG, WEBP, HEIC (iPhone), BMP, GIF, TIFF (várias páginas), ICO, AVIF, prints colados (Ctrl+V)\n" +
                "   → 1, 2, 4 ou 6 por página, modo scanner para documentos, girar, cortar ou encaixar\n\n" +
                "PDF: junta vários PDFs em um só (e pode girar páginas)\n\n" +
                "WORD, EXCEL, POWERPOINT: " + office + "\n\n" +
                "TEXTOS: TXT, Markdown (.md), CSV (vira tabela), JSON, XML, códigos e logs\n\n" +
                "WEB: arquivos HTML, MHT e SVG, ou o endereço de um site (botão \"Adicionar site\")\n\n" +
                "E-MAIL: arquivos .eml (salvos do Outlook, Gmail, Thunderbird...)\n\n" +
                "Dica: você pode misturar tudo — por exemplo, um contrato .docx + fotos do RG + um comprovante em PDF — e sai um único PDF na ordem da lista.",
                "O que dá para converter?", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        void CriarAtalhoEnviarPara()
        {
            try
            {
                string pasta = Environment.GetFolderPath(Environment.SpecialFolder.SendTo);
                string lnk = Path.Combine(pasta, "Conversor PDF.lnk");
                dynamic sh = Activator.CreateInstance(Type.GetTypeFromProgID("WScript.Shell"));
                dynamic atalho = sh.CreateShortcut(lnk);
                atalho.TargetPath = Application.ExecutablePath;
                atalho.Description = "Transformar em PDF";
                atalho.IconLocation = Application.ExecutablePath + ",0";
                atalho.Save();
                MessageBox.Show(this, "Pronto!\n\nAgora, no Explorador de Arquivos, selecione os arquivos, clique com o botão direito e escolha:\n\n" +
                    "Mostrar mais opções > Enviar para > Conversor PDF\n\nO programa abre com os arquivos já na lista.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Não foi possível criar o atalho: " + ex.Message, Text);
            }
        }
    }
}
