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
        public System.Windows.Forms.Button btnEliminarItem;
        public System.Windows.Forms.Button btnVaciarCarro;
        public System.Windows.Forms.Button btnExportarCarro;
        public System.Windows.Forms.Button btnPedirCarro;
        public System.Windows.Forms.Button btnmulti;
        public System.Windows.Forms.Button btnSalir;
        public System.Windows.Forms.Label Label2;
        public System.Windows.Forms.PictureBox PictureBox1;

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
            gridPedidos = new DataGridView();
            btnEliminarItem = new Button();
            btnVaciarCarro = new Button();
            btnExportarCarro = new Button();
            btnPedirCarro = new Button();
            btnmulti = new Button();
            btnSalir = new Button();
            lblTotal = new Label();
            Label2 = new Label();
            PictureBox1 = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)gridPedidos).BeginInit();
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
            gridPedidos.Location = new Point(14, 14);
            gridPedidos.Margin = new Padding(4, 3, 4, 3);
            gridPedidos.Name = "gridPedidos";
            gridPedidos.ReadOnly = true;
            gridPedidos.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gridPedidos.Size = new Size(905, 447);
            gridPedidos.TabIndex = 0;
            gridPedidos.Click += gridPedidos_Click;
            gridPedidos.DoubleClick += gridPedidos_DoubleClick;
            // 
            // btnEliminarItem
            // 
            btnEliminarItem.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnEliminarItem.Enabled = false;
            btnEliminarItem.Location = new Point(14, 467);
            btnEliminarItem.Margin = new Padding(4, 3, 4, 3);
            btnEliminarItem.Name = "btnEliminarItem";
            btnEliminarItem.Size = new Size(117, 27);
            btnEliminarItem.TabIndex = 1;
            btnEliminarItem.Text = "Eliminar Item";
            btnEliminarItem.Click += btnEliminarItem_Click;
            // 
            // btnVaciarCarro
            // 
            btnVaciarCarro.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnVaciarCarro.Location = new Point(138, 467);
            btnVaciarCarro.Margin = new Padding(4, 3, 4, 3);
            btnVaciarCarro.Name = "btnVaciarCarro";
            btnVaciarCarro.Size = new Size(117, 27);
            btnVaciarCarro.TabIndex = 2;
            btnVaciarCarro.Text = "Vaciar Carro";
            btnVaciarCarro.Click += btnVaciarCarro_Click;
            // 
            // btnExportarCarro
            // 
            btnExportarCarro.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnExportarCarro.Location = new Point(261, 467);
            btnExportarCarro.Margin = new Padding(4, 3, 4, 3);
            btnExportarCarro.Name = "btnExportarCarro";
            btnExportarCarro.Size = new Size(117, 27);
            btnExportarCarro.TabIndex = 3;
            btnExportarCarro.Text = "Exportar PDF";
            btnExportarCarro.Click += btnExportarCarro_Click;
            // 
            // btnPedirCarro
            // 
            btnPedirCarro.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnPedirCarro.Location = new Point(385, 467);
            btnPedirCarro.Margin = new Padding(4, 3, 4, 3);
            btnPedirCarro.Name = "btnPedirCarro";
            btnPedirCarro.Size = new Size(117, 27);
            btnPedirCarro.TabIndex = 4;
            btnPedirCarro.Text = "Pedir Carro";
            btnPedirCarro.Click += btnPedirCarro_Click;
            // 
            // btnmulti
            // 
            btnmulti.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnmulti.Location = new Point(509, 467);
            btnmulti.Margin = new Padding(4, 3, 4, 3);
            btnmulti.Name = "btnmulti";
            btnmulti.Size = new Size(117, 27);
            btnmulti.TabIndex = 5;
            btnmulti.Text = "Actualizar";
            btnmulti.Click += btnmulti_Click;
            // 
            // btnSalir
            // 
            btnSalir.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnSalir.Location = new Point(803, 467);
            btnSalir.Margin = new Padding(4, 3, 4, 3);
            btnSalir.Name = "btnSalir";
            btnSalir.Size = new Size(117, 27);
            btnSalir.TabIndex = 6;
            btnSalir.Text = "Salir";
            btnSalir.Click += btnSalir_Click;
            // 
            // lblTotal
            // 
            lblTotal.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            lblTotal.AutoSize = true;
            lblTotal.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblTotal.Location = new Point(632, 471);
            lblTotal.Margin = new Padding(4, 0, 4, 0);
            lblTotal.Name = "lblTotal";
            lblTotal.Size = new Size(0, 16);
            lblTotal.TabIndex = 7;
            // 
            // Label2
            // 
            Label2.AutoSize = true;
            Label2.Font = new Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold);
            Label2.Location = new Point(14, 497);
            Label2.Margin = new Padding(4, 0, 4, 0);
            Label2.Name = "Label2";
            Label2.Size = new Size(226, 16);
            Label2.TabIndex = 8;
            Label2.Text = "Procesando, por favor espere...";
            Label2.Visible = false;
            // 
            // PictureBox1
            // 
            PictureBox1.Image = Properties.Resources.loading;
            PictureBox1.Location = new Point(262, 497);
            PictureBox1.Margin = new Padding(4, 3, 4, 3);
            PictureBox1.Name = "PictureBox1";
            PictureBox1.Size = new Size(16, 16);
            PictureBox1.SizeMode = PictureBoxSizeMode.AutoSize;
            PictureBox1.TabIndex = 9;
            PictureBox1.TabStop = false;
            PictureBox1.Visible = false;
            // 
            // frmCarroCompras
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(933, 519);
            Controls.Add(PictureBox1);
            Controls.Add(Label2);
            Controls.Add(lblTotal);
            Controls.Add(btnSalir);
            Controls.Add(btnmulti);
            Controls.Add(btnPedirCarro);
            Controls.Add(btnExportarCarro);
            Controls.Add(btnVaciarCarro);
            Controls.Add(btnEliminarItem);
            Controls.Add(gridPedidos);
            Margin = new Padding(4, 3, 4, 3);
            Name = "frmCarroCompras";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Carro de Compras";
            Load += frmCarroCompras_Load;
            ((System.ComponentModel.ISupportInitialize)gridPedidos).EndInit();
            ((System.ComponentModel.ISupportInitialize)PictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
