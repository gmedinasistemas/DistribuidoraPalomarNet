namespace DistribuidoraPalomarNet
{
    partial class frmPrincipal
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        // Controles públicos
        public System.Windows.Forms.CheckBox ckbOfertas;
        public System.Windows.Forms.TextBox txtCodigo;
        public System.Windows.Forms.TextBox txtDescripcion;
        public System.Windows.Forms.TextBox txtEquivalencia;
        public System.Windows.Forms.TextBox txtSugerido;
        public System.Windows.Forms.TextBox txtStock;
        public System.Windows.Forms.TextBox txtDescuento1;
        public System.Windows.Forms.TextBox txtDescuento2;
        public System.Windows.Forms.TextBox txtDescuentoFamilia;
        public System.Windows.Forms.TextBox txtDescuento3;
        public System.Windows.Forms.TextBox txtPreciolista;
        public System.Windows.Forms.TextBox txtPreciopublico;
        public System.Windows.Forms.TextBox txtBeneficio;
        public System.Windows.Forms.PictureBox ptbImagen;
        public System.Windows.Forms.PictureBox imgBanner;
        public System.Windows.Forms.Label lblSelecion;
        public System.Windows.Forms.DataGridView DataGridView1;
        public System.Windows.Forms.Label lblusuario;
        public System.Windows.Forms.Label lblFechaAct;
        public System.Windows.Forms.ComboBox cmbArmadora;
        public System.Windows.Forms.Label lblD3;
        public System.Windows.Forms.Button btnGenerarTxt;
        public System.Windows.Forms.Button btnVerCarrito;
        public System.Windows.Forms.Button btnLimpiar;
        public System.Windows.Forms.Label Label1;
        public System.Windows.Forms.Label Label2;
        public System.Windows.Forms.Label Label4;
        public System.Windows.Forms.Label lblDescripcion;
        public System.Windows.Forms.PictureBox PictureBox1;

        // Controles privados
        private System.Windows.Forms.Button btncerrar;
        private System.Windows.Forms.Button btnActualizar;
        private System.Windows.Forms.GroupBox GroupBox1;
        private System.Windows.Forms.Label lblPP;
        private System.Windows.Forms.Label lblB;
        private System.Windows.Forms.Label lblPC;
        private System.Windows.Forms.Label lblD2;
        private System.Windows.Forms.Label lblD1;
        private System.Windows.Forms.Label lblD3Fam;
        private System.Windows.Forms.Label lblPL;
        private System.Windows.Forms.ToolTip ToolTip1;
        private System.Windows.Forms.Panel PanelFiltros;
        private System.Windows.Forms.CheckBox CheckBox1;
        private System.Windows.Forms.Button btnMinimizar;
        private System.Windows.Forms.Button btnMaximizar;
        private System.Windows.Forms.Button btnCerrar;

        // Controles del menú
        private System.Windows.Forms.MenuStrip MenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem ArchivoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ActualizarBasesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ActualizarDatosLocalToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ConfiguracionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ClientesToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem DescPorMarcaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem VerToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem PedidosEnviadosToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem AcercaDeToolStripMenuItem;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmPrincipal));
            btncerrar = new Button();
            btnActualizar = new Button();
            GroupBox1 = new GroupBox();
            txtIvaFlia = new TextBox();
            lblIvaFlia = new Label();
            btnConfUsu = new Button();
            Label3 = new Label();
            txtPreciolista = new TextBox();
            txtDescuento1 = new TextBox();
            txtDescuento2 = new TextBox();
            txtDescuentoFamilia = new TextBox();
            txtPreciopublico = new TextBox();
            txtBeneficio = new TextBox();
            txtSugerido = new TextBox();
            lblD1 = new Label();
            lblD2 = new Label();
            lblD3Fam = new Label();
            lblPC = new Label();
            lblB = new Label();
            lblPP = new Label();
            lblPL = new Label();
            CheckBox1 = new CheckBox();
            txtStock = new TextBox();
            ptbImagen = new PictureBox();
            lblD3 = new Label();
            txtDescuento3 = new TextBox();
            PanelFiltros = new Panel();
            cmbRubros = new ComboBox();
            label9 = new Label();
            cmbMarcas = new ComboBox();
            label12 = new Label();
            cmbArmadora = new ComboBox();
            txtModelo = new TextBox();
            label7 = new Label();
            cmbLineas = new ComboBox();
            label5 = new Label();
            txtCodigo = new TextBox();
            Label1 = new Label();
            Label2 = new Label();
            txtDescripcion = new TextBox();
            btnLimpiar = new Button();
            label6 = new Label();
            lblTotalRegistros = new Label();
            MenuStrip1 = new MenuStrip();
            ArchivoToolStripMenuItem = new ToolStripMenuItem();
            ActualizarBasesToolStripMenuItem = new ToolStripMenuItem();
            ActualizarDatosLocalToolStripMenuItem = new ToolStripMenuItem();
            actualizarAplicaciónToolStripMenuItem = new ToolStripMenuItem();
            ConfiguracionToolStripMenuItem = new ToolStripMenuItem();
            ClientesToolStripMenuItem = new ToolStripMenuItem();
            DescPorMarcaToolStripMenuItem = new ToolStripMenuItem();
            confPorFamiliaToolStripMenuItem = new ToolStripMenuItem();
            VerToolStripMenuItem = new ToolStripMenuItem();
            PedidosEnviadosToolStripMenuItem = new ToolStripMenuItem();
            AcercaDeToolStripMenuItem = new ToolStripMenuItem();
            ToolTip1 = new ToolTip(components);
            btnGenerarTxt = new Button();
            btnVerCarrito = new Button();
            btnWhatsapp = new Button();
            Timer1 = new System.Windows.Forms.Timer(components);
            DataGridView1 = new DataGridView();
            txtEquivalencia = new TextBox();
            ckbOfertas = new CheckBox();
            btnEquivalencia = new Button();
            PictureBox1 = new PictureBox();
            lblSelecion = new Label();
            lblusuario = new Label();
            lblFechaAct = new Label();
            Label4 = new Label();
            lblDescripcion = new Label();
            lblemail = new Label();
            imgBanner = new PictureBox();
            DataGridView2 = new DataGridView();
            label8 = new Label();
            groupBox2 = new GroupBox();
            lblOrigen = new Label();
            label10 = new Label();
            label11 = new Label();
            lblObservaciones = new Label();
            lblRubro = new Label();
            btnnovedades = new Button();
            btnOfertas = new Button();
            TimerVerOfertas = new System.Windows.Forms.Timer(components);
            TimerVerOfertasEquivalencias = new System.Windows.Forms.Timer(components);
            PictureBox3 = new PictureBox();
            lblAlertaUpdate = new Label();
            timer2 = new System.Windows.Forms.Timer(components);
            backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            webView21 = new Microsoft.Web.WebView2.WinForms.WebView2();
            toolStripProgressBar1 = new ProgressBar();
            webView22 = new Microsoft.Web.WebView2.WinForms.WebView2();
            GroupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)ptbImagen).BeginInit();
            PanelFiltros.SuspendLayout();
            MenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)DataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)imgBanner).BeginInit();
            ((System.ComponentModel.ISupportInitialize)DataGridView2).BeginInit();
            groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox3).BeginInit();
            ((System.ComponentModel.ISupportInitialize)webView21).BeginInit();
            ((System.ComponentModel.ISupportInitialize)webView22).BeginInit();
            SuspendLayout();
            // 
            // btncerrar
            // 
            btncerrar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btncerrar.Location = new Point(621, 593);
            btncerrar.Margin = new Padding(4, 3, 4, 3);
            btncerrar.Name = "btncerrar";
            btncerrar.Size = new Size(52, 52);
            btncerrar.TabIndex = 5;
            btncerrar.Text = "Cerrar";
            btncerrar.UseVisualStyleBackColor = true;
            btncerrar.Visible = false;
            // 
            // btnActualizar
            // 
            btnActualizar.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnActualizar.Image = (Image)resources.GetObject("btnActualizar.Image");
            btnActualizar.Location = new Point(1083, 27);
            btnActualizar.Name = "btnActualizar";
            btnActualizar.Size = new Size(52, 52);
            btnActualizar.TabIndex = 15;
            ToolTip1.SetToolTip(btnActualizar, "ACTUALIZAR BASE");
            btnActualizar.UseVisualStyleBackColor = false;
            btnActualizar.Click += BtnActualizar_Click;
            // 
            // GroupBox1
            // 
            GroupBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            GroupBox1.BackColor = Color.FromArgb(32, 59, 175);
            GroupBox1.Controls.Add(txtIvaFlia);
            GroupBox1.Controls.Add(lblIvaFlia);
            GroupBox1.Controls.Add(btnConfUsu);
            GroupBox1.Controls.Add(Label3);
            GroupBox1.Controls.Add(txtPreciolista);
            GroupBox1.Controls.Add(txtDescuento1);
            GroupBox1.Controls.Add(txtDescuento2);
            GroupBox1.Controls.Add(txtDescuentoFamilia);
            GroupBox1.Controls.Add(txtPreciopublico);
            GroupBox1.Controls.Add(txtBeneficio);
            GroupBox1.Controls.Add(txtSugerido);
            GroupBox1.Controls.Add(lblD1);
            GroupBox1.Controls.Add(lblD2);
            GroupBox1.Controls.Add(lblD3Fam);
            GroupBox1.Controls.Add(lblPC);
            GroupBox1.Controls.Add(lblB);
            GroupBox1.Controls.Add(lblPP);
            GroupBox1.Controls.Add(lblPL);
            GroupBox1.Controls.Add(CheckBox1);
            GroupBox1.Location = new Point(12, 582);
            GroupBox1.Margin = new Padding(4, 3, 4, 3);
            GroupBox1.Name = "GroupBox1";
            GroupBox1.Padding = new Padding(4, 3, 4, 3);
            GroupBox1.Size = new Size(581, 67);
            GroupBox1.TabIndex = 4;
            GroupBox1.TabStop = false;
            // 
            // txtIvaFlia
            // 
            txtIvaFlia.Location = new Point(286, 38);
            txtIvaFlia.Name = "txtIvaFlia";
            txtIvaFlia.Size = new Size(37, 23);
            txtIvaFlia.TabIndex = 21;
            txtIvaFlia.TextAlign = HorizontalAlignment.Right;
            // 
            // lblIvaFlia
            // 
            lblIvaFlia.ForeColor = Color.White;
            lblIvaFlia.Location = new Point(286, 20);
            lblIvaFlia.Margin = new Padding(4, 0, 4, 0);
            lblIvaFlia.Name = "lblIvaFlia";
            lblIvaFlia.Size = new Size(29, 27);
            lblIvaFlia.TabIndex = 20;
            lblIvaFlia.Text = "IVA";
            // 
            // btnConfUsu
            // 
            btnConfUsu.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnConfUsu.BackColor = Color.FromArgb(32, 59, 175);
            btnConfUsu.Image = (Image)resources.GetObject("btnConfUsu.Image");
            btnConfUsu.Location = new Point(522, 9);
            btnConfUsu.Name = "btnConfUsu";
            btnConfUsu.Size = new Size(52, 52);
            btnConfUsu.TabIndex = 19;
            ToolTip1.SetToolTip(btnConfUsu, "MIS DATOS");
            btnConfUsu.UseVisualStyleBackColor = false;
            btnConfUsu.Click += BtnConfUsu_Click;
            // 
            // Label3
            // 
            Label3.Location = new Point(22, 89);
            Label3.Margin = new Padding(4, 0, 4, 0);
            Label3.Name = "Label3";
            Label3.Size = new Size(68, 23);
            Label3.TabIndex = 15;
            Label3.Text = "Stock:";
            Label3.TextAlign = ContentAlignment.MiddleLeft;
            Label3.Visible = false;
            // 
            // txtPreciolista
            // 
            txtPreciolista.Location = new Point(26, 38);
            txtPreciolista.Margin = new Padding(4, 3, 4, 3);
            txtPreciolista.Name = "txtPreciolista";
            txtPreciolista.Size = new Size(65, 23);
            txtPreciolista.TabIndex = 2;
            txtPreciolista.TextAlign = HorizontalAlignment.Right;
            // 
            // txtDescuento1
            // 
            txtDescuento1.Location = new Point(91, 38);
            txtDescuento1.Margin = new Padding(4, 3, 4, 3);
            txtDescuento1.Name = "txtDescuento1";
            txtDescuento1.Size = new Size(65, 23);
            txtDescuento1.TabIndex = 4;
            txtDescuento1.TextAlign = HorizontalAlignment.Right;
            // 
            // txtDescuento2
            // 
            txtDescuento2.Location = new Point(156, 38);
            txtDescuento2.Margin = new Padding(4, 3, 4, 3);
            txtDescuento2.Name = "txtDescuento2";
            txtDescuento2.Size = new Size(65, 23);
            txtDescuento2.TabIndex = 6;
            txtDescuento2.TextAlign = HorizontalAlignment.Right;
            // 
            // txtDescuentoFamilia
            // 
            txtDescuentoFamilia.Location = new Point(221, 38);
            txtDescuentoFamilia.Margin = new Padding(4, 3, 4, 3);
            txtDescuentoFamilia.Name = "txtDescuentoFamilia";
            txtDescuentoFamilia.ReadOnly = true;
            txtDescuentoFamilia.Size = new Size(65, 23);
            txtDescuentoFamilia.TabIndex = 7;
            txtDescuentoFamilia.TextAlign = HorizontalAlignment.Right;
            // 
            // txtPreciopublico
            // 
            txtPreciopublico.Location = new Point(324, 38);
            txtPreciopublico.Margin = new Padding(4, 3, 4, 3);
            txtPreciopublico.Name = "txtPreciopublico";
            txtPreciopublico.Size = new Size(65, 23);
            txtPreciopublico.TabIndex = 8;
            txtPreciopublico.TextAlign = HorizontalAlignment.Right;
            // 
            // txtBeneficio
            // 
            txtBeneficio.Location = new Point(389, 38);
            txtBeneficio.Name = "txtBeneficio";
            txtBeneficio.Size = new Size(65, 23);
            txtBeneficio.TabIndex = 10;
            txtBeneficio.TextAlign = HorizontalAlignment.Right;
            // 
            // txtSugerido
            // 
            txtSugerido.Location = new Point(455, 38);
            txtSugerido.Name = "txtSugerido";
            txtSugerido.Size = new Size(65, 23);
            txtSugerido.TabIndex = 12;
            txtSugerido.TextAlign = HorizontalAlignment.Right;
            // 
            // lblD1
            // 
            lblD1.ForeColor = Color.White;
            lblD1.Location = new Point(91, 20);
            lblD1.Margin = new Padding(4, 0, 4, 0);
            lblD1.Name = "lblD1";
            lblD1.Size = new Size(66, 27);
            lblD1.TabIndex = 3;
            lblD1.Text = "DTO 1";
            // 
            // lblD2
            // 
            lblD2.ForeColor = Color.White;
            lblD2.Location = new Point(156, 20);
            lblD2.Margin = new Padding(4, 0, 4, 0);
            lblD2.Name = "lblD2";
            lblD2.Size = new Size(66, 27);
            lblD2.TabIndex = 5;
            lblD2.Text = "DTO 2";
            // 
            // lblD3Fam
            // 
            lblD3Fam.ForeColor = Color.White;
            lblD3Fam.Location = new Point(221, 20);
            lblD3Fam.Margin = new Padding(4, 0, 4, 0);
            lblD3Fam.Name = "lblD3Fam";
            lblD3Fam.Size = new Size(66, 27);
            lblD3Fam.TabIndex = 6;
            lblD3Fam.Text = "DTO FLIA";
            // 
            // lblPC
            // 
            lblPC.ForeColor = Color.White;
            lblPC.Location = new Point(324, 20);
            lblPC.Name = "lblPC";
            lblPC.Size = new Size(66, 27);
            lblPC.TabIndex = 7;
            lblPC.Text = "COSTO";
            // 
            // lblB
            // 
            lblB.ForeColor = Color.White;
            lblB.Location = new Point(389, 20);
            lblB.Name = "lblB";
            lblB.Size = new Size(66, 27);
            lblB.TabIndex = 9;
            lblB.Text = "BENEFICIO";
            // 
            // lblPP
            // 
            lblPP.ForeColor = Color.White;
            lblPP.Location = new Point(455, 20);
            lblPP.Margin = new Padding(4, 0, 4, 0);
            lblPP.Name = "lblPP";
            lblPP.Size = new Size(66, 27);
            lblPP.TabIndex = 11;
            lblPP.Text = "VENTA";
            // 
            // lblPL
            // 
            lblPL.ForeColor = Color.White;
            lblPL.Location = new Point(26, 20);
            lblPL.Margin = new Padding(4, 0, 4, 0);
            lblPL.Name = "lblPL";
            lblPL.Size = new Size(66, 27);
            lblPL.TabIndex = 1;
            lblPL.Text = "LISTA";
            // 
            // CheckBox1
            // 
            CheckBox1.Checked = true;
            CheckBox1.CheckState = CheckState.Checked;
            CheckBox1.Location = new Point(7, 17);
            CheckBox1.Name = "CheckBox1";
            CheckBox1.Size = new Size(67, 20);
            CheckBox1.TabIndex = 0;
            CheckBox1.Click += CheckBox1_Click;
            // 
            // txtStock
            // 
            txtStock.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            txtStock.Location = new Point(539, 557);
            txtStock.Name = "txtStock";
            txtStock.Size = new Size(65, 23);
            txtStock.TabIndex = 14;
            txtStock.TextAlign = HorizontalAlignment.Right;
            txtStock.Visible = false;
            // 
            // ptbImagen
            // 
            ptbImagen.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            ptbImagen.Location = new Point(939, 433);
            ptbImagen.Margin = new Padding(4, 3, 4, 3);
            ptbImagen.Name = "ptbImagen";
            ptbImagen.Size = new Size(248, 212);
            ptbImagen.SizeMode = PictureBoxSizeMode.Zoom;
            ptbImagen.TabIndex = 0;
            ptbImagen.TabStop = false;
            ptbImagen.Click += ptbImagen_Click;
            // 
            // lblD3
            // 
            lblD3.Location = new Point(12, 369);
            lblD3.Margin = new Padding(4, 0, 4, 0);
            lblD3.Name = "lblD3";
            lblD3.Size = new Size(117, 27);
            lblD3.TabIndex = 7;
            lblD3.Text = "Descuento 3:";
            // 
            // txtDescuento3
            // 
            txtDescuento3.Location = new Point(105, 369);
            txtDescuento3.Margin = new Padding(4, 3, 4, 3);
            txtDescuento3.Name = "txtDescuento3";
            txtDescuento3.Size = new Size(116, 23);
            txtDescuento3.TabIndex = 8;
            // 
            // PanelFiltros
            // 
            PanelFiltros.BackColor = Color.FromArgb(32, 59, 175);
            PanelFiltros.Controls.Add(cmbRubros);
            PanelFiltros.Controls.Add(label9);
            PanelFiltros.Controls.Add(cmbMarcas);
            PanelFiltros.Controls.Add(label12);
            PanelFiltros.Controls.Add(cmbArmadora);
            PanelFiltros.Controls.Add(txtModelo);
            PanelFiltros.Controls.Add(label7);
            PanelFiltros.Controls.Add(cmbLineas);
            PanelFiltros.Controls.Add(label5);
            PanelFiltros.Controls.Add(txtCodigo);
            PanelFiltros.Controls.Add(Label1);
            PanelFiltros.Controls.Add(Label2);
            PanelFiltros.Controls.Add(txtDescripcion);
            PanelFiltros.Controls.Add(btnLimpiar);
            PanelFiltros.Controls.Add(label6);
            PanelFiltros.Controls.Add(lblTotalRegistros);
            PanelFiltros.Location = new Point(212, 75);
            PanelFiltros.Margin = new Padding(4, 3, 4, 3);
            PanelFiltros.Name = "PanelFiltros";
            PanelFiltros.Size = new Size(869, 87);
            PanelFiltros.TabIndex = 2;
            // 
            // cmbRubros
            // 
            cmbRubros.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbRubros.Location = new Point(76, 32);
            cmbRubros.Margin = new Padding(4, 3, 4, 3);
            cmbRubros.Name = "cmbRubros";
            cmbRubros.Size = new Size(133, 23);
            cmbRubros.TabIndex = 13;
            cmbRubros.SelectedIndexChanged += cmbRubros_SelectedIndexChanged;
            // 
            // label9
            // 
            label9.ForeColor = Color.White;
            label9.Location = new Point(1, 32);
            label9.Margin = new Padding(4, 0, 4, 0);
            label9.Name = "label9";
            label9.Size = new Size(52, 23);
            label9.TabIndex = 16;
            label9.Text = "RUBRO";
            // 
            // cmbMarcas
            // 
            cmbMarcas.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbMarcas.Location = new Point(671, 32);
            cmbMarcas.Margin = new Padding(4, 3, 4, 3);
            cmbMarcas.Name = "cmbMarcas";
            cmbMarcas.Size = new Size(133, 23);
            cmbMarcas.TabIndex = 14;
            cmbMarcas.SelectedIndexChanged += cmbMarcas_SelectedIndexChanged;
            // 
            // label12
            // 
            label12.ForeColor = Color.White;
            label12.Location = new Point(620, 34);
            label12.Margin = new Padding(4, 0, 4, 0);
            label12.Name = "label12";
            label12.Size = new Size(64, 23);
            label12.TabIndex = 15;
            label12.Text = "MARCA";
            label12.Click += label12_Click;
            // 
            // cmbArmadora
            // 
            cmbArmadora.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbArmadora.Location = new Point(294, 32);
            cmbArmadora.Margin = new Padding(4, 3, 4, 3);
            cmbArmadora.Name = "cmbArmadora";
            cmbArmadora.Size = new Size(133, 23);
            cmbArmadora.TabIndex = 5;
            cmbArmadora.SelectedIndexChanged += cmbArmadora_SelectedIndexChanged;
            // 
            // txtModelo
            // 
            txtModelo.Location = new Point(649, 3);
            txtModelo.Margin = new Padding(4, 3, 4, 3);
            txtModelo.Name = "txtModelo";
            txtModelo.Size = new Size(156, 23);
            txtModelo.TabIndex = 8;
            txtModelo.TextChanged += txtModelo_TextChanged;
            // 
            // label7
            // 
            label7.ForeColor = Color.White;
            label7.Location = new Point(219, 34);
            label7.Margin = new Padding(4, 0, 4, 0);
            label7.Name = "label7";
            label7.Size = new Size(78, 23);
            label7.TabIndex = 11;
            label7.Text = "ARMADORA";
            // 
            // cmbLineas
            // 
            cmbLineas.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbLineas.Location = new Point(480, 32);
            cmbLineas.Margin = new Padding(4, 3, 4, 3);
            cmbLineas.Name = "cmbLineas";
            cmbLineas.Size = new Size(133, 23);
            cmbLineas.TabIndex = 9;
            cmbLineas.SelectedIndexChanged += cmbLineas_SelectedIndexChanged;
            // 
            // label5
            // 
            label5.ForeColor = Color.White;
            label5.Location = new Point(520, 3);
            label5.Margin = new Padding(4, 0, 4, 0);
            label5.Name = "label5";
            label5.Size = new Size(129, 32);
            label5.TabIndex = 7;
            label5.Text = "OBSERVACION / CODIGO UNIVERSAL";
            // 
            // txtCodigo
            // 
            txtCodigo.Location = new Point(75, 3);
            txtCodigo.Margin = new Padding(4, 3, 4, 3);
            txtCodigo.Name = "txtCodigo";
            txtCodigo.Size = new Size(163, 23);
            txtCodigo.TabIndex = 1;
            txtCodigo.KeyPress += txtCodigo_KeyPress;
            // 
            // Label1
            // 
            Label1.ForeColor = Color.White;
            Label1.Location = new Point(0, 3);
            Label1.Margin = new Padding(4, 0, 4, 0);
            Label1.Name = "Label1";
            Label1.Size = new Size(93, 23);
            Label1.TabIndex = 0;
            Label1.Text = "CÓDIGO";
            // 
            // Label2
            // 
            Label2.ForeColor = Color.White;
            Label2.Location = new Point(238, 3);
            Label2.Margin = new Padding(4, 0, 4, 0);
            Label2.Name = "Label2";
            Label2.Size = new Size(93, 23);
            Label2.TabIndex = 2;
            Label2.Text = "DESCRIPCIÓN";
            // 
            // txtDescripcion
            // 
            txtDescripcion.Location = new Point(331, 3);
            txtDescripcion.Margin = new Padding(4, 3, 4, 3);
            txtDescripcion.Name = "txtDescripcion";
            txtDescripcion.Size = new Size(181, 23);
            txtDescripcion.TabIndex = 3;
            txtDescripcion.TextChanged += txtDescripcion_TextChanged;
            // 
            // btnLimpiar
            // 
            btnLimpiar.Image = (Image)resources.GetObject("btnLimpiar.Image");
            btnLimpiar.Location = new Point(812, 3);
            btnLimpiar.Margin = new Padding(4, 3, 4, 3);
            btnLimpiar.Name = "btnLimpiar";
            btnLimpiar.Size = new Size(52, 52);
            btnLimpiar.TabIndex = 6;
            ToolTip1.SetToolTip(btnLimpiar, "LIMPIAR");
            btnLimpiar.UseVisualStyleBackColor = true;
            btnLimpiar.Click += BtnLimpiar_Click;
            // 
            // label6
            // 
            label6.ForeColor = Color.White;
            label6.Location = new Point(432, 34);
            label6.Margin = new Padding(4, 0, 4, 0);
            label6.Name = "label6";
            label6.Size = new Size(45, 23);
            label6.TabIndex = 10;
            label6.Text = "LINEAS";
            // 
            // lblTotalRegistros
            // 
            lblTotalRegistros.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            lblTotalRegistros.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point, 0);
            lblTotalRegistros.ForeColor = Color.White;
            lblTotalRegistros.Location = new Point(12, 65);
            lblTotalRegistros.Margin = new Padding(4, 0, 4, 0);
            lblTotalRegistros.Name = "lblTotalRegistros";
            lblTotalRegistros.Size = new Size(361, 27);
            lblTotalRegistros.TabIndex = 12;
            // 
            // MenuStrip1
            // 
            MenuStrip1.Items.AddRange(new ToolStripItem[] { ArchivoToolStripMenuItem, ConfiguracionToolStripMenuItem, VerToolStripMenuItem, AcercaDeToolStripMenuItem });
            MenuStrip1.Location = new Point(0, 0);
            MenuStrip1.Name = "MenuStrip1";
            MenuStrip1.Padding = new Padding(7, 2, 0, 2);
            MenuStrip1.Size = new Size(1195, 24);
            MenuStrip1.TabIndex = 0;
            // 
            // ArchivoToolStripMenuItem
            // 
            ArchivoToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { ActualizarBasesToolStripMenuItem, ActualizarDatosLocalToolStripMenuItem, actualizarAplicaciónToolStripMenuItem });
            ArchivoToolStripMenuItem.Name = "ArchivoToolStripMenuItem";
            ArchivoToolStripMenuItem.Size = new Size(60, 20);
            ArchivoToolStripMenuItem.Text = "Archivo";
            // 
            // ActualizarBasesToolStripMenuItem
            // 
            ActualizarBasesToolStripMenuItem.Name = "ActualizarBasesToolStripMenuItem";
            ActualizarBasesToolStripMenuItem.Size = new Size(190, 22);
            ActualizarBasesToolStripMenuItem.Text = "Actualizar Bases";
            // 
            // ActualizarDatosLocalToolStripMenuItem
            // 
            ActualizarDatosLocalToolStripMenuItem.Name = "ActualizarDatosLocalToolStripMenuItem";
            ActualizarDatosLocalToolStripMenuItem.Size = new Size(190, 22);
            ActualizarDatosLocalToolStripMenuItem.Text = "Actualizar Datos Local";
            // 
            // actualizarAplicaciónToolStripMenuItem
            // 
            actualizarAplicaciónToolStripMenuItem.Name = "actualizarAplicaciónToolStripMenuItem";
            actualizarAplicaciónToolStripMenuItem.Size = new Size(190, 22);
            actualizarAplicaciónToolStripMenuItem.Text = "Actualizar Aplicación";
            actualizarAplicaciónToolStripMenuItem.Click += actualizarAplicaciónToolStripMenuItem_Click;
            // 
            // ConfiguracionToolStripMenuItem
            // 
            ConfiguracionToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { ClientesToolStripMenuItem, DescPorMarcaToolStripMenuItem, confPorFamiliaToolStripMenuItem });
            ConfiguracionToolStripMenuItem.Name = "ConfiguracionToolStripMenuItem";
            ConfiguracionToolStripMenuItem.Size = new Size(95, 20);
            ConfiguracionToolStripMenuItem.Text = "Configuración";
            // 
            // ClientesToolStripMenuItem
            // 
            ClientesToolStripMenuItem.Name = "ClientesToolStripMenuItem";
            ClientesToolStripMenuItem.Size = new Size(165, 22);
            ClientesToolStripMenuItem.Text = "Clientes";
            ClientesToolStripMenuItem.Click += ClientesToolStripMenuItem_Click;
            // 
            // DescPorMarcaToolStripMenuItem
            // 
            DescPorMarcaToolStripMenuItem.Name = "DescPorMarcaToolStripMenuItem";
            DescPorMarcaToolStripMenuItem.Size = new Size(165, 22);
            DescPorMarcaToolStripMenuItem.Text = "Desc. por Marca";
            DescPorMarcaToolStripMenuItem.Visible = false;
            DescPorMarcaToolStripMenuItem.Click += DescPorMarcaToolStripMenuItem_Click;
            // 
            // confPorFamiliaToolStripMenuItem
            // 
            confPorFamiliaToolStripMenuItem.Name = "confPorFamiliaToolStripMenuItem";
            confPorFamiliaToolStripMenuItem.Size = new Size(165, 22);
            confPorFamiliaToolStripMenuItem.Text = "Conf. por Familia";
            confPorFamiliaToolStripMenuItem.Click += confPorFamiliaToolStripMenuItem_Click;
            // 
            // VerToolStripMenuItem
            // 
            VerToolStripMenuItem.DropDownItems.AddRange(new ToolStripItem[] { PedidosEnviadosToolStripMenuItem });
            VerToolStripMenuItem.Name = "VerToolStripMenuItem";
            VerToolStripMenuItem.Size = new Size(35, 20);
            VerToolStripMenuItem.Text = "Ver";
            // 
            // PedidosEnviadosToolStripMenuItem
            // 
            PedidosEnviadosToolStripMenuItem.Name = "PedidosEnviadosToolStripMenuItem";
            PedidosEnviadosToolStripMenuItem.Size = new Size(166, 22);
            PedidosEnviadosToolStripMenuItem.Text = "Pedidos Enviados";
            // 
            // AcercaDeToolStripMenuItem
            // 
            AcercaDeToolStripMenuItem.Name = "AcercaDeToolStripMenuItem";
            AcercaDeToolStripMenuItem.Size = new Size(71, 20);
            AcercaDeToolStripMenuItem.Text = "Acerca de";
            AcercaDeToolStripMenuItem.Click += AcercaDeToolStripMenuItem_Click;
            // 
            // btnGenerarTxt
            // 
            btnGenerarTxt.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnGenerarTxt.Image = (Image)resources.GetObject("btnGenerarTxt.Image");
            btnGenerarTxt.Location = new Point(1083, 81);
            btnGenerarTxt.Name = "btnGenerarTxt";
            btnGenerarTxt.Size = new Size(52, 52);
            btnGenerarTxt.TabIndex = 16;
            ToolTip1.SetToolTip(btnGenerarTxt, "EXPORTAR LISTA");
            btnGenerarTxt.UseVisualStyleBackColor = true;
            btnGenerarTxt.Click += BtnGenerarTxt_Click;
            // 
            // btnVerCarrito
            // 
            btnVerCarrito.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnVerCarrito.Image = (Image)resources.GetObject("btnVerCarrito.Image");
            btnVerCarrito.Location = new Point(1135, 27);
            btnVerCarrito.Margin = new Padding(4, 3, 4, 3);
            btnVerCarrito.Name = "btnVerCarrito";
            btnVerCarrito.Size = new Size(52, 52);
            btnVerCarrito.TabIndex = 6;
            ToolTip1.SetToolTip(btnVerCarrito, "VER CARRITO");
            btnVerCarrito.UseVisualStyleBackColor = false;
            // 
            // btnWhatsapp
            // 
            btnWhatsapp.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnWhatsapp.Image = (Image)resources.GetObject("btnWhatsapp.Image");
            btnWhatsapp.Location = new Point(1135, 81);
            btnWhatsapp.Margin = new Padding(4, 3, 4, 3);
            btnWhatsapp.Name = "btnWhatsapp";
            btnWhatsapp.Size = new Size(52, 52);
            btnWhatsapp.TabIndex = 184;
            ToolTip1.SetToolTip(btnWhatsapp, "CONTACTENOS");
            btnWhatsapp.UseVisualStyleBackColor = false;
            btnWhatsapp.Click += btnWhatsapp_Click;
            // 
            // DataGridView1
            // 
            DataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically;
            DataGridView1.Location = new Point(12, 168);
            DataGridView1.Margin = new Padding(4, 3, 4, 3);
            DataGridView1.MultiSelect = false;
            DataGridView1.Name = "DataGridView1";
            DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            DataGridView1.Size = new Size(1175, 262);
            DataGridView1.TabIndex = 3;
            DataGridView1.ColumnHeaderMouseClick += DataGridView1_ColumnHeaderMouseClick;
            DataGridView1.Click += DataGridView1_Click;
            DataGridView1.DoubleClick += DataGridView1_DoubleClick;
            // 
            // txtEquivalencia
            // 
            txtEquivalencia.Location = new Point(0, 0);
            txtEquivalencia.Name = "txtEquivalencia";
            txtEquivalencia.Size = new Size(100, 23);
            txtEquivalencia.TabIndex = 0;
            // 
            // ckbOfertas
            // 
            ckbOfertas.Location = new Point(0, 0);
            ckbOfertas.Name = "ckbOfertas";
            ckbOfertas.Size = new Size(104, 24);
            ckbOfertas.TabIndex = 0;
            // 
            // btnEquivalencia
            // 
            btnEquivalencia.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnEquivalencia.Location = new Point(586, 593);
            btnEquivalencia.Name = "btnEquivalencia";
            btnEquivalencia.Size = new Size(52, 52);
            btnEquivalencia.TabIndex = 17;
            btnEquivalencia.Text = "Equivalencias";
            btnEquivalencia.UseVisualStyleBackColor = true;
            btnEquivalencia.Visible = false;
            btnEquivalencia.Click += BtnEquivalencia_Click;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = (Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new Point(12, 27);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new Size(200, 152);
            PictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            PictureBox1.TabIndex = 144;
            PictureBox1.TabStop = false;
            // 
            // lblSelecion
            // 
            lblSelecion.Location = new Point(0, 0);
            lblSelecion.Name = "lblSelecion";
            lblSelecion.Size = new Size(100, 23);
            lblSelecion.TabIndex = 0;
            // 
            // lblusuario
            // 
            lblusuario.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            lblusuario.Location = new Point(599, 535);
            lblusuario.Margin = new Padding(4, 0, 4, 0);
            lblusuario.Name = "lblusuario";
            lblusuario.Size = new Size(117, 27);
            lblusuario.TabIndex = 8;
            lblusuario.Visible = false;
            lblusuario.Click += lblusuario_Click;
            // 
            // lblFechaAct
            // 
            lblFechaAct.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            lblFechaAct.Location = new Point(939, 647);
            lblFechaAct.Margin = new Padding(4, 0, 4, 0);
            lblFechaAct.Name = "lblFechaAct";
            lblFechaAct.Size = new Size(117, 18);
            lblFechaAct.TabIndex = 9;
            lblFechaAct.Visible = false;
            // 
            // Label4
            // 
            Label4.Location = new Point(0, 0);
            Label4.Name = "Label4";
            Label4.Size = new Size(100, 23);
            Label4.TabIndex = 0;
            // 
            // lblDescripcion
            // 
            lblDescripcion.Location = new Point(0, 0);
            lblDescripcion.Name = "lblDescripcion";
            lblDescripcion.Size = new Size(100, 23);
            lblDescripcion.TabIndex = 0;
            // 
            // lblemail
            // 
            lblemail.Location = new Point(0, 0);
            lblemail.Name = "lblemail";
            lblemail.Size = new Size(100, 23);
            lblemail.TabIndex = 0;
            // 
            // imgBanner
            // 
            imgBanner.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            imgBanner.Location = new Point(211, 24);
            imgBanner.Name = "imgBanner";
            imgBanner.Size = new Size(870, 50);
            imgBanner.TabIndex = 11;
            imgBanner.TabStop = false;
            // 
            // DataGridView2
            // 
            DataGridView2.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            DataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            DataGridView2.EditMode = DataGridViewEditMode.EditProgrammatically;
            DataGridView2.Location = new Point(12, 447);
            DataGridView2.Name = "DataGridView2";
            DataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            DataGridView2.Size = new Size(519, 133);
            DataGridView2.TabIndex = 176;
            DataGridView2.ColumnHeaderMouseClick += DataGridView2_ColumnHeaderMouseClick;
            DataGridView2.Click += DataGridView2_Click;
            DataGridView2.DoubleClick += DataGridView2_DoubleClick;
            // 
            // label8
            // 
            label8.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            label8.BackColor = Color.FromArgb(179, 179, 179);
            label8.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label8.ForeColor = Color.Black;
            label8.Location = new Point(12, 432);
            label8.Margin = new Padding(4, 0, 4, 0);
            label8.Name = "label8";
            label8.Size = new Size(519, 23);
            label8.TabIndex = 177;
            label8.Text = "EQUIVALENCIAS";
            // 
            // groupBox2
            // 
            groupBox2.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            groupBox2.BackColor = Color.FromArgb(204, 204, 204);
            groupBox2.Controls.Add(lblOrigen);
            groupBox2.Controls.Add(label10);
            groupBox2.Controls.Add(label11);
            groupBox2.Controls.Add(lblObservaciones);
            groupBox2.Controls.Add(lblRubro);
            groupBox2.Location = new Point(538, 432);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(394, 92);
            groupBox2.TabIndex = 178;
            groupBox2.TabStop = false;
            groupBox2.Text = "OBSERVACIONES";
            // 
            // lblOrigen
            // 
            lblOrigen.BackColor = Color.FromArgb(255, 255, 191);
            lblOrigen.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblOrigen.ForeColor = Color.Black;
            lblOrigen.Location = new Point(134, 20);
            lblOrigen.Margin = new Padding(4, 0, 4, 0);
            lblOrigen.Name = "lblOrigen";
            lblOrigen.Size = new Size(124, 27);
            lblOrigen.TabIndex = 11;
            // 
            // label10
            // 
            label10.BackColor = Color.FromArgb(255, 255, 191);
            label10.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            label10.ForeColor = Color.Black;
            label10.Location = new Point(257, 20);
            label10.Margin = new Padding(4, 0, 4, 0);
            label10.Name = "label10";
            label10.Size = new Size(66, 27);
            label10.TabIndex = 12;
            // 
            // label11
            // 
            label11.BackColor = Color.FromArgb(255, 255, 191);
            label11.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            label11.ForeColor = Color.Black;
            label11.Location = new Point(322, 20);
            label11.Name = "label11";
            label11.Size = new Size(66, 27);
            label11.TabIndex = 13;
            // 
            // lblObservaciones
            // 
            lblObservaciones.BackColor = Color.FromArgb(255, 255, 191);
            lblObservaciones.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblObservaciones.ForeColor = Color.Black;
            lblObservaciones.Location = new Point(13, 47);
            lblObservaciones.Name = "lblObservaciones";
            lblObservaciones.Size = new Size(375, 39);
            lblObservaciones.TabIndex = 14;
            // 
            // lblRubro
            // 
            lblRubro.BackColor = Color.FromArgb(255, 255, 191);
            lblRubro.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            lblRubro.ForeColor = Color.Black;
            lblRubro.Location = new Point(13, 20);
            lblRubro.Margin = new Padding(4, 0, 4, 0);
            lblRubro.Name = "lblRubro";
            lblRubro.Size = new Size(124, 27);
            lblRubro.TabIndex = 10;
            // 
            // btnnovedades
            // 
            btnnovedades.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnnovedades.BackColor = Color.FromArgb(87, 151, 161);
            btnnovedades.ForeColor = Color.White;
            btnnovedades.Location = new Point(598, 620);
            btnnovedades.Name = "btnnovedades";
            btnnovedades.Size = new Size(98, 24);
            btnnovedades.TabIndex = 179;
            btnnovedades.Text = "NOVEDADES";
            btnnovedades.UseVisualStyleBackColor = false;
            btnnovedades.Click += Btnnovedades_Click;
            // 
            // btnOfertas
            // 
            btnOfertas.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnOfertas.BackColor = Color.FromArgb(87, 151, 161);
            btnOfertas.ForeColor = Color.White;
            btnOfertas.Location = new Point(598, 590);
            btnOfertas.Name = "btnOfertas";
            btnOfertas.Size = new Size(98, 24);
            btnOfertas.TabIndex = 180;
            btnOfertas.Text = "OFERTAS";
            btnOfertas.UseVisualStyleBackColor = false;
            btnOfertas.Click += btnOfertas_Click;
            // 
            // TimerVerOfertas
            // 
            TimerVerOfertas.Interval = 10;
            TimerVerOfertas.Tick += TimerVerOfertas_Tick;
            // 
            // TimerVerOfertasEquivalencias
            // 
            TimerVerOfertasEquivalencias.Interval = 10;
            TimerVerOfertasEquivalencias.Tick += TimerVerOfertasEquivalencias_Tick;
            // 
            // PictureBox3
            // 
            PictureBox3.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            PictureBox3.BackColor = Color.Transparent;
            PictureBox3.Image = (Image)resources.GetObject("PictureBox3.Image");
            PictureBox3.Location = new Point(738, 535);
            PictureBox3.Name = "PictureBox3";
            PictureBox3.Size = new Size(194, 110);
            PictureBox3.SizeMode = PictureBoxSizeMode.Zoom;
            PictureBox3.TabIndex = 183;
            PictureBox3.TabStop = false;
            // 
            // lblAlertaUpdate
            // 
            lblAlertaUpdate.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            lblAlertaUpdate.AutoSize = true;
            lblAlertaUpdate.Location = new Point(738, 650);
            lblAlertaUpdate.Name = "lblAlertaUpdate";
            lblAlertaUpdate.Size = new Size(0, 15);
            lblAlertaUpdate.TabIndex = 186;
            lblAlertaUpdate.Visible = false;
            // 
            // timer2
            // 
            timer2.Interval = 1400;
            timer2.Tick += timer2_Tick;
            // 
            // webView21
            // 
            webView21.AllowExternalDrop = true;
            webView21.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            webView21.CreationProperties = null;
            webView21.DefaultBackgroundColor = Color.White;
            webView21.Location = new Point(0, 677);
            webView21.Name = "webView21";
            webView21.Size = new Size(1195, 23);
            webView21.TabIndex = 187;
            webView21.ZoomFactor = 1D;
            // 
            // toolStripProgressBar1
            // 
            toolStripProgressBar1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            toolStripProgressBar1.Location = new Point(538, 522);
            toolStripProgressBar1.Name = "toolStripProgressBar1";
            toolStripProgressBar1.Size = new Size(394, 11);
            toolStripProgressBar1.TabIndex = 188;
            // 
            // webView22
            // 
            webView22.AllowExternalDrop = true;
            webView22.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            webView22.CreationProperties = null;
            webView22.DefaultBackgroundColor = Color.White;
            webView22.Location = new Point(0, 647);
            webView22.Name = "webView22";
            webView22.Size = new Size(696, 31);
            webView22.TabIndex = 189;
            webView22.ZoomFactor = 1D;
            // 
            // frmPrincipal
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(232, 234, 233);
            ClientSize = new Size(1195, 701);
            Controls.Add(toolStripProgressBar1);
            Controls.Add(webView21);
            Controls.Add(lblAlertaUpdate);
            Controls.Add(btnWhatsapp);
            Controls.Add(btnOfertas);
            Controls.Add(btnnovedades);
            Controls.Add(groupBox2);
            Controls.Add(DataGridView2);
            Controls.Add(GroupBox1);
            Controls.Add(ptbImagen);
            Controls.Add(imgBanner);
            Controls.Add(MenuStrip1);
            Controls.Add(PanelFiltros);
            Controls.Add(txtStock);
            Controls.Add(btncerrar);
            Controls.Add(btnVerCarrito);
            Controls.Add(lblusuario);
            Controls.Add(lblFechaAct);
            Controls.Add(btnActualizar);
            Controls.Add(btnGenerarTxt);
            Controls.Add(btnEquivalencia);
            Controls.Add(DataGridView1);
            Controls.Add(label8);
            Controls.Add(PictureBox1);
            Controls.Add(PictureBox3);
            Controls.Add(webView22);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MainMenuStrip = MenuStrip1;
            Margin = new Padding(4, 3, 4, 3);
            Name = "frmPrincipal";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Distribuidora Palomar";
            WindowState = FormWindowState.Maximized;
            Load += frmPrincipal_Load;
            GroupBox1.ResumeLayout(false);
            GroupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)ptbImagen).EndInit();
            PanelFiltros.ResumeLayout(false);
            PanelFiltros.PerformLayout();
            MenuStrip1.ResumeLayout(false);
            MenuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)DataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)imgBanner).EndInit();
            ((System.ComponentModel.ISupportInitialize)DataGridView2).EndInit();
            groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)PictureBox3).EndInit();
            ((System.ComponentModel.ISupportInitialize)webView21).EndInit();
            ((System.ComponentModel.ISupportInitialize)webView22).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public Label Label3;
        private Button btnConfUsu;
        public ComboBox cmbLineas;
        public Label label5;
        public TextBox txtModelo;
        public Label label7;
        public Label label6;
        private DataGridView DataGridView2;
        public Label label8;
        private GroupBox groupBox2;
        private Button btnnovedades;
        private Button btnOfertas;
        public System.Windows.Forms.Timer TimerVerOfertas;
        public Button btnEquivalencia;
        public System.Windows.Forms.Timer TimerVerOfertasEquivalencias;
        public System.Windows.Forms.Timer Timer1;
        private Label label10;
        private Label label11;
        public Label lblRubro;
        public Label lblOrigen;
        public Label lblObservaciones;
        public PictureBox PictureBox3;
        public Button btnWhatsapp;
        public Label lblemail;
        public Label lblTotalRegistros;
        private ToolStripMenuItem confPorFamiliaToolStripMenuItem;
        public Label lblAlertaUpdate;
        public System.Windows.Forms.Timer timer2;
        private ToolStripMenuItem actualizarAplicaciónToolStripMenuItem;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        public Microsoft.Web.WebView2.WinForms.WebView2 webView21;
        public ProgressBar toolStripProgressBar1;
        public TextBox txtIvaFlia;
        private Label lblIvaFlia;
        public ComboBox cmbRubros;
        public Label label9;
        public ComboBox cmbMarcas;
        public Label label12;
        public Microsoft.Web.WebView2.WinForms.WebView2 webView22;
    }
}
