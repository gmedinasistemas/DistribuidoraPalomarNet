using System.Drawing;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    partial class frmConfFamilia
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
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmConfFamilia));
            DataGridView1 = new DataGridView();
            lblbgral = new Label();
            txtbgral = new TextBox();
            lbliva = new Label();
            txtiva = new TextBox();
            CheckBox1 = new CheckBox();
            btnGrabar = new Button();
            btnSalir = new Button();
            btnRestaurar = new Button();
            ptbImgSi = new PictureBox();
            ptbImgNo = new PictureBox();
            ToolTip1 = new ToolTip(components);
            ((System.ComponentModel.ISupportInitialize)DataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgSi).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgNo).BeginInit();
            SuspendLayout();
            // 
            // DataGridView1
            // 
            DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            DataGridView1.Location = new Point(16, 7);
            DataGridView1.Margin = new Padding(4, 3, 4, 3);
            DataGridView1.Name = "DataGridView1";
            DataGridView1.Size = new Size(589, 677);
            DataGridView1.TabIndex = 0;
            // 
            // lblbgral
            // 
            lblbgral.AutoSize = true;
            lblbgral.Enabled = false;
            lblbgral.Font = new Font("MS Reference Sans Serif", 8.25F);
            lblbgral.ForeColor = Color.White;
            lblbgral.Location = new Point(196, 688);
            lblbgral.Margin = new Padding(4, 0, 4, 0);
            lblbgral.Name = "lblbgral";
            lblbgral.Size = new Size(94, 15);
            lblbgral.TabIndex = 1;
            lblbgral.Text = "Beneficio gral.:";
            // 
            // txtbgral
            // 
            txtbgral.Enabled = false;
            txtbgral.Font = new Font("MS Reference Sans Serif", 8.25F);
            txtbgral.Location = new Point(296, 688);
            txtbgral.Margin = new Padding(4, 3, 4, 3);
            txtbgral.Name = "txtbgral";
            txtbgral.Size = new Size(102, 21);
            txtbgral.TabIndex = 2;
            txtbgral.TextAlign = HorizontalAlignment.Center;
            // 
            // lbliva
            // 
            lbliva.AutoSize = true;
            lbliva.Enabled = false;
            lbliva.Font = new Font("MS Reference Sans Serif", 8.25F);
            lbliva.ForeColor = Color.White;
            lbliva.Location = new Point(196, 715);
            lbliva.Margin = new Padding(4, 0, 4, 0);
            lbliva.Name = "lbliva";
            lbliva.Size = new Size(33, 15);
            lbliva.TabIndex = 3;
            lbliva.Text = "IVA:";
            // 
            // txtiva
            // 
            txtiva.Enabled = false;
            txtiva.Font = new Font("MS Reference Sans Serif", 8.25F);
            txtiva.Location = new Point(296, 715);
            txtiva.Margin = new Padding(4, 3, 4, 3);
            txtiva.Name = "txtiva";
            txtiva.Size = new Size(102, 21);
            txtiva.TabIndex = 4;
            txtiva.TextAlign = HorizontalAlignment.Center;
            // 
            // CheckBox1
            // 
            CheckBox1.AutoSize = true;
            CheckBox1.ForeColor = Color.White;
            CheckBox1.Location = new Point(16, 688);
            CheckBox1.Margin = new Padding(4, 3, 4, 3);
            CheckBox1.Name = "CheckBox1";
            CheckBox1.Size = new Size(148, 19);
            CheckBox1.TabIndex = 5;
            CheckBox1.Text = "Habilitar Carga General";
            CheckBox1.UseVisualStyleBackColor = true;
            // 
            // btnGrabar
            // 
            btnGrabar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnGrabar.Font = new Font("Calibri", 11.25F, FontStyle.Bold);
            btnGrabar.ForeColor = Color.Purple;
            btnGrabar.Location = new Point(478, 688);
            btnGrabar.Margin = new Padding(4, 3, 4, 3);
            btnGrabar.Name = "btnGrabar";
            btnGrabar.Size = new Size(58, 60);
            btnGrabar.TabIndex = 6;
            btnGrabar.Text = "✓";
            ToolTip1.SetToolTip(btnGrabar, "Grabar");
            btnGrabar.UseVisualStyleBackColor = true;
            btnGrabar.Click += btnGrabar_Click;
            // 
            // btnSalir
            // 
            btnSalir.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnSalir.Font = new Font("Calibri", 11.25F, FontStyle.Bold);
            btnSalir.ForeColor = Color.Purple;
            btnSalir.Location = new Point(545, 688);
            btnSalir.Margin = new Padding(4, 3, 4, 3);
            btnSalir.Name = "btnSalir";
            btnSalir.Size = new Size(58, 60);
            btnSalir.TabIndex = 7;
            btnSalir.Text = "X";
            ToolTip1.SetToolTip(btnSalir, "Salir");
            btnSalir.UseVisualStyleBackColor = true;
            btnSalir.Click += btnSalir_Click;
            // 
            // btnRestaurar
            // 
            btnRestaurar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnRestaurar.Font = new Font("Calibri", 11.25F, FontStyle.Bold);
            btnRestaurar.ForeColor = Color.Purple;
            btnRestaurar.Location = new Point(16, 715);
            btnRestaurar.Margin = new Padding(4, 3, 4, 3);
            btnRestaurar.MaximumSize = new Size(58, 32);
            btnRestaurar.MinimumSize = new Size(58, 32);
            btnRestaurar.Name = "btnRestaurar";
            btnRestaurar.Size = new Size(58, 32);
            btnRestaurar.TabIndex = 8;
            btnRestaurar.Text = "↺";
            ToolTip1.SetToolTip(btnRestaurar, "Restaurar Valores");
            btnRestaurar.UseVisualStyleBackColor = true;
            // 
            // ptbImgSi
            // 
            ptbImgSi.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            ptbImgSi.Image = (Image)resources.GetObject("ptbImgSi.Image");
            ptbImgSi.Location = new Point(408, 688);
            ptbImgSi.Margin = new Padding(4, 3, 4, 3);
            ptbImgSi.Name = "ptbImgSi";
            ptbImgSi.Size = new Size(61, 57);
            ptbImgSi.SizeMode = PictureBoxSizeMode.StretchImage;
            ptbImgSi.TabIndex = 9;
            ptbImgSi.TabStop = false;
            ptbImgSi.Visible = false;
            // 
            // ptbImgNo
            // 
            ptbImgNo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            ptbImgNo.Image = (Image)resources.GetObject("ptbImgNo.Image");
            ptbImgNo.Location = new Point(408, 688);
            ptbImgNo.Margin = new Padding(4, 3, 4, 3);
            ptbImgNo.Name = "ptbImgNo";
            ptbImgNo.Size = new Size(61, 57);
            ptbImgNo.SizeMode = PictureBoxSizeMode.StretchImage;
            ptbImgNo.TabIndex = 10;
            ptbImgNo.TabStop = false;
            ptbImgNo.Visible = false;
            // 
            // frmConfFamilia
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(29, 29, 27);
            ClientSize = new Size(621, 768);
            Controls.Add(lbliva);
            Controls.Add(txtiva);
            Controls.Add(CheckBox1);
            Controls.Add(lblbgral);
            Controls.Add(txtbgral);
            Controls.Add(DataGridView1);
            Controls.Add(btnRestaurar);
            Controls.Add(ptbImgNo);
            Controls.Add(ptbImgSi);
            Controls.Add(btnSalir);
            Controls.Add(btnGrabar);
            FormBorderStyle = FormBorderStyle.None;
            Margin = new Padding(4, 3, 4, 3);
            Name = "frmConfFamilia";
            StartPosition = FormStartPosition.CenterParent;
            Text = "Configuración x Familia";
            Load += frmConfFamilia_Load;
            ((System.ComponentModel.ISupportInitialize)DataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgSi).EndInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgNo).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private DataGridView DataGridView1;
        private Label lblbgral;
        private TextBox txtbgral;
        private Label lbliva;
        private TextBox txtiva;
        private CheckBox CheckBox1;
        private Button btnGrabar;
        private Button btnSalir;
        private Button btnRestaurar;
        private PictureBox ptbImgSi;
        private PictureBox ptbImgNo;
        private ToolTip ToolTip1;
    }
}
