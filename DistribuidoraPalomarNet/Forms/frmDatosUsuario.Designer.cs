namespace DistribuidoraPalomarNet
{
    partial class frmDatosUsuario
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmDatosUsuario));
            txtCodCliente = new TextBox();
            txtCuitCliente = new TextBox();
            txtEmail = new TextBox();
            txtDesc1 = new TextBox();
            txtDesc2 = new TextBox();
            txtBene = new TextBox();
            btnGrabar = new Button();
            btnSalir = new Button();
            ptbImgSi = new PictureBox();
            ptbImgNo = new PictureBox();
            lblCodCliente = new Label();
            lblCuitCliente = new Label();
            lblEmail = new Label();
            lblDesc1 = new Label();
            lblDesc2 = new Label();
            lblBene = new Label();
            lblUsuario = new Label();
            label1 = new Label();
            label2 = new Label();
            ((System.ComponentModel.ISupportInitialize)ptbImgSi).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgNo).BeginInit();
            SuspendLayout();
            // 
            // txtCodCliente
            // 
            txtCodCliente.Location = new Point(166, 67);
            txtCodCliente.Margin = new Padding(4, 3, 4, 3);
            txtCodCliente.Name = "txtCodCliente";
            txtCodCliente.Size = new Size(256, 23);
            txtCodCliente.TabIndex = 68;
            // 
            // txtCuitCliente
            // 
            txtCuitCliente.Location = new Point(166, 115);
            txtCuitCliente.Margin = new Padding(4, 3, 4, 3);
            txtCuitCliente.Name = "txtCuitCliente";
            txtCuitCliente.Size = new Size(256, 23);
            txtCuitCliente.TabIndex = 69;
            // 
            // txtEmail
            // 
            txtEmail.Location = new Point(166, 145);
            txtEmail.Margin = new Padding(4, 3, 4, 3);
            txtEmail.Name = "txtEmail";
            txtEmail.Size = new Size(256, 23);
            txtEmail.TabIndex = 70;
            // 
            // txtDesc1
            // 
            txtDesc1.Location = new Point(166, 189);
            txtDesc1.Margin = new Padding(4, 3, 4, 3);
            txtDesc1.Name = "txtDesc1";
            txtDesc1.Size = new Size(116, 23);
            txtDesc1.TabIndex = 74;
            // 
            // txtDesc2
            // 
            txtDesc2.Location = new Point(166, 219);
            txtDesc2.Margin = new Padding(4, 3, 4, 3);
            txtDesc2.Name = "txtDesc2";
            txtDesc2.Size = new Size(116, 23);
            txtDesc2.TabIndex = 75;
            // 
            // txtBene
            // 
            txtBene.Location = new Point(166, 248);
            txtBene.Margin = new Padding(4, 3, 4, 3);
            txtBene.Name = "txtBene";
            txtBene.Size = new Size(116, 23);
            txtBene.TabIndex = 76;
            txtBene.Text = "0";
            txtBene.Visible = false;
            // 
            // btnGrabar
            // 
            btnGrabar.Image = (Image)resources.GetObject("btnGrabar.Image");
            btnGrabar.Location = new Point(375, 3);
            btnGrabar.Margin = new Padding(4, 3, 4, 3);
            btnGrabar.Name = "btnGrabar";
            btnGrabar.Size = new Size(42, 42);
            btnGrabar.TabIndex = 77;
            btnGrabar.Tag = "Grabar";
            btnGrabar.UseVisualStyleBackColor = true;
            btnGrabar.Click += btnGrabar_Click;
            // 
            // btnSalir
            // 
            btnSalir.DialogResult = DialogResult.Cancel;
            btnSalir.Location = new Point(425, 3);
            btnSalir.Margin = new Padding(4, 3, 4, 3);
            btnSalir.Name = "btnSalir";
            btnSalir.Size = new Size(42, 42);
            btnSalir.TabIndex = 78;
            btnSalir.Tag = "Salir";
            btnSalir.Text = "X";
            btnSalir.UseVisualStyleBackColor = true;
            btnSalir.Click += btnSalir_Click;
            // 
            // ptbImgSi
            // 
            ptbImgSi.Image = (Image)resources.GetObject("ptbImgSi.Image");
            ptbImgSi.Location = new Point(318, 5);
            ptbImgSi.Margin = new Padding(4, 3, 4, 3);
            ptbImgSi.Name = "ptbImgSi";
            ptbImgSi.Size = new Size(37, 37);
            ptbImgSi.SizeMode = PictureBoxSizeMode.StretchImage;
            ptbImgSi.TabIndex = 85;
            ptbImgSi.TabStop = false;
            ptbImgSi.Visible = false;
            // 
            // ptbImgNo
            // 
            ptbImgNo.Image = (Image)resources.GetObject("ptbImgNo.Image");
            ptbImgNo.Location = new Point(318, 5);
            ptbImgNo.Margin = new Padding(4, 3, 4, 3);
            ptbImgNo.Name = "ptbImgNo";
            ptbImgNo.Size = new Size(37, 37);
            ptbImgNo.SizeMode = PictureBoxSizeMode.StretchImage;
            ptbImgNo.TabIndex = 86;
            ptbImgNo.TabStop = false;
            ptbImgNo.Visible = false;
            // 
            // lblCodCliente
            // 
            lblCodCliente.AutoSize = true;
            lblCodCliente.ForeColor = Color.White;
            lblCodCliente.Location = new Point(12, 67);
            lblCodCliente.Margin = new Padding(4, 0, 4, 0);
            lblCodCliente.Name = "lblCodCliente";
            lblCodCliente.Size = new Size(49, 15);
            lblCodCliente.TabIndex = 65;
            lblCodCliente.Text = "Código:";
            // 
            // lblCuitCliente
            // 
            lblCuitCliente.AutoSize = true;
            lblCuitCliente.ForeColor = Color.White;
            lblCuitCliente.Location = new Point(12, 115);
            lblCuitCliente.Margin = new Padding(4, 0, 4, 0);
            lblCuitCliente.Name = "lblCuitCliente";
            lblCuitCliente.Size = new Size(35, 15);
            lblCuitCliente.TabIndex = 66;
            lblCuitCliente.Text = "CUIT:";
            // 
            // lblEmail
            // 
            lblEmail.AutoSize = true;
            lblEmail.ForeColor = Color.White;
            lblEmail.Location = new Point(12, 145);
            lblEmail.Margin = new Padding(4, 0, 4, 0);
            lblEmail.Name = "lblEmail";
            lblEmail.Size = new Size(39, 15);
            lblEmail.TabIndex = 67;
            lblEmail.Text = "Email:";
            // 
            // lblDesc1
            // 
            lblDesc1.AutoSize = true;
            lblDesc1.ForeColor = Color.White;
            lblDesc1.Location = new Point(12, 189);
            lblDesc1.Margin = new Padding(4, 0, 4, 0);
            lblDesc1.Name = "lblDesc1";
            lblDesc1.Size = new Size(109, 15);
            lblDesc1.TabIndex = 71;
            lblDesc1.Text = "Descuento General:";
            // 
            // lblDesc2
            // 
            lblDesc2.AutoSize = true;
            lblDesc2.ForeColor = Color.White;
            lblDesc2.Location = new Point(12, 219);
            lblDesc2.Margin = new Padding(4, 0, 4, 0);
            lblDesc2.Name = "lblDesc2";
            lblDesc2.Size = new Size(111, 15);
            lblDesc2.TabIndex = 73;
            lblDesc2.Text = "Descuento Especial:";
            // 
            // lblBene
            // 
            lblBene.AutoSize = true;
            lblBene.ForeColor = Color.White;
            lblBene.Location = new Point(12, 248);
            lblBene.Margin = new Padding(4, 0, 4, 0);
            lblBene.Name = "lblBene";
            lblBene.Size = new Size(59, 15);
            lblBene.TabIndex = 87;
            lblBene.Text = "Beneficio:";
            lblBene.Visible = false;
            // 
            // lblUsuario
            // 
            lblUsuario.AutoSize = true;
            lblUsuario.Location = new Point(167, 48);
            lblUsuario.Name = "lblUsuario";
            lblUsuario.Size = new Size(0, 15);
            lblUsuario.TabIndex = 88;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.ForeColor = Color.White;
            label1.Location = new Point(166, 93);
            label1.Name = "label1";
            label1.Size = new Size(147, 15);
            label1.TabIndex = 0;
            label1.Text = "(puede verlo en la Factura)";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Georgia", 9.75F, FontStyle.Bold);
            label2.ForeColor = Color.White;
            label2.Location = new Point(12, 17);
            label2.Name = "label2";
            label2.Size = new Size(220, 16);
            label2.TabIndex = 89;
            label2.Text = "Ingrese Sus Datos Personales:";
            // 
            // frmDatosUsuario
            // 
            AcceptButton = btnGrabar;
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(32, 59, 175);
            CancelButton = btnSalir;
            ClientSize = new Size(474, 301);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(lblUsuario);
            Controls.Add(txtBene);
            Controls.Add(lblBene);
            Controls.Add(btnGrabar);
            Controls.Add(ptbImgNo);
            Controls.Add(ptbImgSi);
            Controls.Add(txtDesc2);
            Controls.Add(txtDesc1);
            Controls.Add(lblDesc2);
            Controls.Add(lblDesc1);
            Controls.Add(txtEmail);
            Controls.Add(txtCuitCliente);
            Controls.Add(txtCodCliente);
            Controls.Add(lblEmail);
            Controls.Add(lblCuitCliente);
            Controls.Add(lblCodCliente);
            Controls.Add(btnSalir);
            FormBorderStyle = FormBorderStyle.None;
            Margin = new Padding(4, 3, 4, 3);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "frmDatosUsuario";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Datos de Usuario";
            Load += frmDatosUsuario_Load;
            ((System.ComponentModel.ISupportInitialize)ptbImgSi).EndInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgNo).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        // Controles públicos
        public System.Windows.Forms.TextBox txtCodCliente;
        public System.Windows.Forms.TextBox txtCuitCliente;
        public System.Windows.Forms.TextBox txtEmail;
        public System.Windows.Forms.TextBox txtDesc1;
        public System.Windows.Forms.TextBox txtDesc2;
        public System.Windows.Forms.TextBox txtBene;
        public System.Windows.Forms.Button btnGrabar;
        public System.Windows.Forms.Button btnSalir;
        public System.Windows.Forms.PictureBox ptbImgSi;
        public System.Windows.Forms.PictureBox ptbImgNo;

        // Controles privados
        private System.Windows.Forms.Label lblCodCliente;
        private System.Windows.Forms.Label lblCuitCliente;
        private System.Windows.Forms.Label lblEmail;
        private System.Windows.Forms.Label lblDesc1;
        private System.Windows.Forms.Label lblDesc2;
        private System.Windows.Forms.Label lblBene;
        public Label lblUsuario;
        private Label label1;
        private Label label2;
    }
}
