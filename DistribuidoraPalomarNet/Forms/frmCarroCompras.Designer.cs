namespace DistribuidoraPalomarNet
{
    partial class frmCarroCompras
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        // Controles públicos
        public System.Windows.Forms.DataGridView gridPedidos;
        public System.Windows.Forms.Label lblTotal;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmCarroCompras));
            gridPedidos = new DataGridView();
            lblTotal = new Label();
            cboxNoEnviar = new CheckBox();
            panel1 = new Panel();
            label3 = new Label();
            txtNotaPedido = new TextBox();
            PictureBox1 = new PictureBox();
            Label2 = new Label();
            label1 = new Label();
            btnEliminarItem = new Button();
            btnExportarCarro = new Button();
            btnPedirCarro = new Button();
            btnmulti = new Button();
            btnVaciarCarro = new Button();
            btnSalir = new Button();
            toolTip1 = new ToolTip(components);
            ((System.ComponentModel.ISupportInitialize)gridPedidos).BeginInit();
            panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).BeginInit();
            SuspendLayout();
            // 
            // gridPedidos
            // 
            gridPedidos.AllowUserToAddRows = false;
            gridPedidos.AllowUserToDeleteRows = false;
            gridPedidos.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            gridPedidos.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            gridPedidos.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            gridPedidos.EditMode = DataGridViewEditMode.EditProgrammatically;
            gridPedidos.Location = new Point(4, 74);
            gridPedidos.Margin = new Padding(4, 3, 4, 3);
            gridPedidos.Name = "gridPedidos";
            gridPedidos.ReadOnly = true;
            gridPedidos.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gridPedidos.Size = new Size(946, 440);
            gridPedidos.TabIndex = 0;
            gridPedidos.Click += gridPedidos_Click;
            gridPedidos.DoubleClick += gridPedidos_DoubleClick;
            // 
            // lblTotal
            // 
            lblTotal.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            lblTotal.AutoSize = true;
            lblTotal.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblTotal.ForeColor = Color.White;
            lblTotal.Location = new Point(655, 525);
            lblTotal.Margin = new Padding(4, 0, 4, 0);
            lblTotal.Name = "lblTotal";
            lblTotal.Size = new Size(0, 16);
            lblTotal.TabIndex = 7;
            // 
            // cboxNoEnviar
            // 
            cboxNoEnviar.AutoSize = true;
            cboxNoEnviar.ForeColor = Color.White;
            cboxNoEnviar.Location = new Point(367, 9);
            cboxNoEnviar.Name = "cboxNoEnviar";
            cboxNoEnviar.Size = new Size(109, 19);
            cboxNoEnviar.TabIndex = 10;
            cboxNoEnviar.Text = "No Enviar EMail";
            cboxNoEnviar.UseVisualStyleBackColor = true;
            cboxNoEnviar.Visible = false;
            // 
            // panel1
            // 
            panel1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panel1.BackColor = Color.FromArgb(32, 59, 175);
            panel1.Controls.Add(label3);
            panel1.Controls.Add(txtNotaPedido);
            panel1.Controls.Add(PictureBox1);
            panel1.Controls.Add(Label2);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(btnEliminarItem);
            panel1.Controls.Add(btnExportarCarro);
            panel1.Controls.Add(btnPedirCarro);
            panel1.Controls.Add(cboxNoEnviar);
            panel1.Controls.Add(btnmulti);
            panel1.Controls.Add(btnVaciarCarro);
            panel1.Controls.Add(btnSalir);
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(954, 68);
            panel1.TabIndex = 11;
            // 
            // label3
            // 
            label3.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 11.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label3.ForeColor = Color.White;
            label3.Location = new Point(653, 15);
            label3.Name = "label3";
            label3.Size = new Size(94, 20);
            label3.TabIndex = 16;
            label3.Text = "Escribir nota:";
            // 
            // txtNotaPedido
            // 
            txtNotaPedido.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            txtNotaPedido.Location = new Point(654, 35);
            txtNotaPedido.Name = "txtNotaPedido";
            txtNotaPedido.Size = new Size(185, 23);
            txtNotaPedido.TabIndex = 15;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = (Image)resources.GetObject("PictureBox1.Image");
            PictureBox1.Location = new Point(595, 6);
            PictureBox1.Margin = new Padding(4, 3, 4, 3);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new Size(52, 52);
            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            PictureBox1.TabIndex = 14;
            PictureBox1.TabStop = false;
            PictureBox1.Visible = false;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold);
            Label2.ForeColor = Color.White;
            Label2.Location = new Point(368, 24);
            Label2.Margin = new Padding(4, 0, 4, 0);
            Label2.Name = "Label2";
            Label2.Size = new Size(226, 16);
            Label2.TabIndex = 13;
            Label2.Text = "Procesando, por favor espere...";
            Label2.Visible = false;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Georgia", 9.75F, FontStyle.Bold);
            label1.ForeColor = Color.White;
            label1.Location = new Point(14, 25);
            label1.Name = "label1";
            label1.Size = new Size(125, 16);
            label1.TabIndex = 12;
            label1.Text = "Carro de Pedido:";
            // 
            // btnEliminarItem
            // 
            btnEliminarItem.BackColor = Color.White;
            btnEliminarItem.Enabled = false;
            btnEliminarItem.Image = (Image)resources.GetObject("btnEliminarItem.Image");
            btnEliminarItem.Location = new Point(152, 7);
            btnEliminarItem.Margin = new Padding(4, 3, 4, 3);
            btnEliminarItem.Name = "btnEliminarItem";
            btnEliminarItem.Size = new Size(52, 52);
            btnEliminarItem.TabIndex = 12;
            toolTip1.SetToolTip(btnEliminarItem, "BORRAR RENGLON");
            btnEliminarItem.UseVisualStyleBackColor = false;
            btnEliminarItem.Click += btnEliminarItem_Click;
            // 
            // btnExportarCarro
            // 
            btnExportarCarro.BackColor = Color.White;
            btnExportarCarro.Image = (Image)resources.GetObject("btnExportarCarro.Image");
            btnExportarCarro.Location = new Point(204, 7);
            btnExportarCarro.Margin = new Padding(4, 3, 4, 3);
            btnExportarCarro.Name = "btnExportarCarro";
            btnExportarCarro.Size = new Size(52, 52);
            btnExportarCarro.TabIndex = 11;
            toolTip1.SetToolTip(btnExportarCarro, "VER PEDIDO");
            btnExportarCarro.UseVisualStyleBackColor = false;
            btnExportarCarro.Click += btnExportarCarro_Click;
            // 
            // btnPedirCarro
            // 
            btnPedirCarro.BackColor = Color.White;
            btnPedirCarro.Image = (Image)resources.GetObject("btnPedirCarro.Image");
            btnPedirCarro.Location = new Point(256, 7);
            btnPedirCarro.Margin = new Padding(4, 3, 4, 3);
            btnPedirCarro.Name = "btnPedirCarro";
            btnPedirCarro.Size = new Size(52, 52);
            btnPedirCarro.TabIndex = 10;
            toolTip1.SetToolTip(btnPedirCarro, "ENVIAR PEDIDO");
            btnPedirCarro.UseVisualStyleBackColor = false;
            btnPedirCarro.Click += btnPedirCarro_Click;
            // 
            // btnmulti
            // 
            btnmulti.BackColor = Color.White;
            btnmulti.Image = (Image)resources.GetObject("btnmulti.Image");
            btnmulti.Location = new Point(308, 7);
            btnmulti.Margin = new Padding(4, 3, 4, 3);
            btnmulti.Name = "btnmulti";
            btnmulti.Size = new Size(52, 52);
            btnmulti.TabIndex = 9;
            toolTip1.SetToolTip(btnmulti, "ACTUALIZAR");
            btnmulti.UseVisualStyleBackColor = false;
            btnmulti.Click += btnmulti_Click;
            // 
            // btnVaciarCarro
            // 
            btnVaciarCarro.BackColor = Color.White;
            btnVaciarCarro.Image = (Image)resources.GetObject("btnVaciarCarro.Image");
            btnVaciarCarro.Location = new Point(846, 6);
            btnVaciarCarro.Margin = new Padding(4, 3, 4, 3);
            btnVaciarCarro.Name = "btnVaciarCarro";
            btnVaciarCarro.Size = new Size(52, 52);
            btnVaciarCarro.TabIndex = 8;
            toolTip1.SetToolTip(btnVaciarCarro, "ANULAR TODO");
            btnVaciarCarro.UseVisualStyleBackColor = false;
            btnVaciarCarro.Click += btnVaciarCarro_Click;
            // 
            // btnSalir
            // 
            btnSalir.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnSalir.ForeColor = Color.White;
            btnSalir.Location = new Point(898, 6);
            btnSalir.Margin = new Padding(4, 3, 4, 3);
            btnSalir.Name = "btnSalir";
            btnSalir.Size = new Size(52, 52);
            btnSalir.TabIndex = 7;
            btnSalir.Text = "X";
            toolTip1.SetToolTip(btnSalir, "CERRAR CARRO");
            btnSalir.Click += btnSalir_Click;
            // 
            // frmCarroCompras
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(32, 59, 175);
            ClientSize = new Size(956, 548);
            Controls.Add(panel1);
            Controls.Add(lblTotal);
            Controls.Add(gridPedidos);
            FormBorderStyle = FormBorderStyle.None;
            Margin = new Padding(4, 3, 4, 3);
            Name = "frmCarroCompras";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Carro de Compras";
            Load += frmCarroCompras_Load;
            ((System.ComponentModel.ISupportInitialize)gridPedidos).EndInit();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public CheckBox cboxNoEnviar;
        private Panel panel1;
        public Button btnSalir;
        public Button btnVaciarCarro;
        public Button btnmulti;
        public Button btnPedirCarro;
        public Button btnEliminarItem;
        public Button btnExportarCarro;
        private Label label1;
        public PictureBox PictureBox1;
        public Label Label2;
        private ToolTip toolTip1;
        private Label label3;
        public TextBox txtNotaPedido;
    }
}
