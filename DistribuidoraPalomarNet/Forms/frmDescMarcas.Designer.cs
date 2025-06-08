namespace DistribuidoraPalomarNet
{
    partial class frmDescMarcas
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            Panel1 = new Panel();
            CheckBox1 = new CheckBox();
            btnSalir = new Button();
            lblbgral = new Label();
            btnGrabar = new Button();
            lbld3gral = new Label();
            Label6 = new Label();
            txtbgral = new TextBox();
            ptbImgNo = new PictureBox();
            txtdesc3g = new TextBox();
            ptbImgSi = new PictureBox();
            DataGridView1 = new DataGridView();

            Panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)ptbImgNo).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgSi).BeginInit();
            ((System.ComponentModel.ISupportInitialize)DataGridView1).BeginInit();
            SuspendLayout();

            // Panel1
            Panel1.BackColor = Color.FromArgb(0, 82, 94);
            Panel1.Controls.AddRange(new Control[] {
                CheckBox1,
                btnSalir,
                lblbgral,
                btnGrabar,
                lbld3gral,
                Label6,
                txtbgral,
                ptbImgNo,
                txtdesc3g,
                ptbImgSi
            });
            Panel1.Dock = DockStyle.Top;
            Panel1.Location = new Point(0, 0);
            Panel1.Margin = new Padding(2);
            Panel1.Name = "Panel1";
            Panel1.Size = new Size(473, 79);
            Panel1.TabIndex = 89;

            // CheckBox1
            CheckBox1.AutoSize = true;
            CheckBox1.ForeColor = Color.White;
            CheckBox1.Location = new Point(16, 35);
            CheckBox1.Name = "CheckBox1";
            CheckBox1.Size = new Size(135, 17);
            CheckBox1.TabIndex = 95;
            CheckBox1.Text = "Habilitar Carga General";
            CheckBox1.UseVisualStyleBackColor = true;
            CheckBox1.CheckedChanged += new EventHandler(CheckBox1_CheckedChanged);

            // btnSalir
            btnSalir.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSalir.BackColor = Color.FromArgb(48, 98, 126);
            btnSalir.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold);
            btnSalir.ForeColor = Color.White;
            btnSalir.Location = new Point(434, 35);
            btnSalir.Name = "btnSalir";
            btnSalir.Size = new Size(36, 36);
            btnSalir.TabIndex = 5;
            btnSalir.Text = "X";
            btnSalir.UseVisualStyleBackColor = false;
            btnSalir.Click += new EventHandler(btnSalir_Click);

            // lblbgral
            lblbgral.AutoSize = true;
            lblbgral.Enabled = false;
            lblbgral.Font = new Font("MS Reference Sans Serif", 8.25F);
            lblbgral.ForeColor = Color.White;
            lblbgral.Location = new Point(149, 56);
            lblbgral.Name = "lblbgral";
            lblbgral.Size = new Size(94, 15);
            lblbgral.TabIndex = 94;
            lblbgral.Text = "Beneficio gral.:";

            // btnGrabar
            btnGrabar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnGrabar.BackColor = Color.FromArgb(48, 98, 126);
            btnGrabar.Font = new Font("Microsoft Sans Serif", 9F, FontStyle.Bold);
            btnGrabar.ForeColor = Color.White;
            btnGrabar.Location = new Point(392, 35);
            btnGrabar.Name = "btnGrabar";
            btnGrabar.Size = new Size(36, 36);
            btnGrabar.TabIndex = 4;
            btnGrabar.Text = "✓";
            btnGrabar.UseVisualStyleBackColor = false;
            btnGrabar.Click += new EventHandler(btnGrabar_Click);

            // lbld3gral
            lbld3gral.AutoSize = true;
            lbld3gral.Enabled = false;
            lbld3gral.Font = new Font("MS Reference Sans Serif", 8.25F);
            lbld3gral.ForeColor = Color.White;
            lbld3gral.Location = new Point(149, 38);
            lbld3gral.Name = "lbld3gral";
            lbld3gral.Size = new Size(94, 15);
            lbld3gral.TabIndex = 92;
            lbld3gral.Text = "Descuento 3 gral.:";

            // Label6
            Label6.AutoSize = true;
            Label6.Font = new Font("MS Reference Sans Serif", 12F, FontStyle.Bold);
            Label6.ForeColor = Color.White;
            Label6.Location = new Point(12, 9);
            Label6.Name = "Label6";
            Label6.Size = new Size(200, 20);
            Label6.TabIndex = 91;
            Label6.Text = "Descuentos por Marca";

            // txtbgral
            txtbgral.Enabled = false;
            txtbgral.Location = new Point(249, 53);
            txtbgral.Name = "txtbgral";
            txtbgral.Size = new Size(100, 20);
            txtbgral.TabIndex = 3;

            // ptbImgNo
            ptbImgNo.Location = new Point(355, 35);
            ptbImgNo.Name = "ptbImgNo";
            ptbImgNo.Size = new Size(31, 36);
            ptbImgNo.SizeMode = PictureBoxSizeMode.StretchImage;
            ptbImgNo.TabIndex = 88;
            ptbImgNo.TabStop = false;
            ptbImgNo.Visible = false;

            // txtdesc3g
            txtdesc3g.Enabled = false;
            txtdesc3g.Location = new Point(249, 35);
            txtdesc3g.Name = "txtdesc3g";
            txtdesc3g.Size = new Size(100, 20);
            txtdesc3g.TabIndex = 2;

            // ptbImgSi
            ptbImgSi.Location = new Point(355, 35);
            ptbImgSi.Name = "ptbImgSi";
            ptbImgSi.Size = new Size(31, 36);
            ptbImgSi.SizeMode = PictureBoxSizeMode.StretchImage;
            ptbImgSi.TabIndex = 87;
            ptbImgSi.TabStop = false;
            ptbImgSi.Visible = false;

            // DataGridView1
            DataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            DataGridView1.BackgroundColor = Color.White;
            DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            DataGridView1.Location = new Point(12, 95);
            DataGridView1.Name = "DataGridView1";
            DataGridView1.Size = new Size(449, 244);
            DataGridView1.TabIndex = 90;

            // frmDescMarcas
            AutoScaleDimensions = new SizeF(6F, 13F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.White;
            ClientSize = new Size(473, 351);
            Controls.Add(DataGridView1);
            Controls.Add(Panel1);
            FormBorderStyle = FormBorderStyle.None;
            Name = "frmDescMarcas";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Descuentos por Marca";

            Panel1.ResumeLayout(false);
            Panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)ptbImgNo).EndInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgSi).EndInit();
            ((System.ComponentModel.ISupportInitialize)DataGridView1).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private Panel Panel1;
        private CheckBox CheckBox1;
        private Button btnSalir;
        private Label lblbgral;
        private Button btnGrabar;
        private Label lbld3gral;
        private Label Label6;
        private TextBox txtbgral;
        private PictureBox ptbImgNo;
        private TextBox txtdesc3g;
        private PictureBox ptbImgSi;
        private DataGridView DataGridView1;
    }
}
