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
            tableLayoutPanel1 = new TableLayoutPanel();
            DataGridView1 = new DataGridView();
            pnlBottom = new Panel();
            lbliva = new Label();
            txtiva = new TextBox();
            lblDescFlia = new Label();
            txtDescFlia = new TextBox();
            CheckBox1 = new CheckBox();
            lblbgral = new Label();
            txtbgral = new TextBox();
            btnRestaurar = new Button();
            ptbImgNo = new PictureBox();
            ptbImgSi = new PictureBox();
            btnSalir = new Button();
            btnGrabar = new Button();
            ToolTip1 = new ToolTip(components);
            tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)DataGridView1).BeginInit();
            pnlBottom.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)ptbImgNo).BeginInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgSi).BeginInit();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 1;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.Controls.Add(DataGridView1, 0, 0);
            tableLayoutPanel1.Controls.Add(pnlBottom, 0, 1);
            tableLayoutPanel1.Dock = DockStyle.Fill;
            tableLayoutPanel1.Location = new Point(0, 0);
            tableLayoutPanel1.Margin = new Padding(0);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 2;
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            tableLayoutPanel1.RowStyles.Add(new RowStyle(SizeType.Absolute, 110F));
            tableLayoutPanel1.Size = new Size(621, 768);
            tableLayoutPanel1.TabIndex = 11;
            // 
            // DataGridView1
            // 
            DataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            DataGridView1.Dock = DockStyle.Fill;
            DataGridView1.Location = new Point(4, 3);
            DataGridView1.Margin = new Padding(4, 3, 4, 3);
            DataGridView1.Name = "DataGridView1";
            DataGridView1.Size = new Size(613, 678);
            DataGridView1.TabIndex = 0;
            // 
            // pnlBottom
            // 
            pnlBottom.Controls.Add(lbliva);
            pnlBottom.Controls.Add(txtiva);
            pnlBottom.Controls.Add(lblDescFlia);
            pnlBottom.Controls.Add(txtDescFlia);
            pnlBottom.Controls.Add(CheckBox1);
            pnlBottom.Controls.Add(lblbgral);
            pnlBottom.Controls.Add(txtbgral);
            pnlBottom.Controls.Add(btnRestaurar);
            pnlBottom.Controls.Add(ptbImgNo);
            pnlBottom.Controls.Add(ptbImgSi);
            pnlBottom.Controls.Add(btnSalir);
            pnlBottom.Controls.Add(btnGrabar);
            pnlBottom.Dock = DockStyle.Fill;
            pnlBottom.Location = new Point(0, 684);
            pnlBottom.Margin = new Padding(0);
            pnlBottom.Name = "pnlBottom";
            pnlBottom.Size = new Size(621, 110);
            pnlBottom.TabIndex = 12;
            // 
            // lbliva
            // 
            lbliva.AutoSize = true;
            lbliva.Enabled = false;
            lbliva.Font = new Font("MS Reference Sans Serif", 8.25F);
            lbliva.ForeColor = Color.White;
            lbliva.Location = new Point(184, 45);
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
            txtiva.Location = new Point(284, 42);
            txtiva.Margin = new Padding(4, 3, 4, 3);
            txtiva.Name = "txtiva";
            txtiva.Size = new Size(102, 21);
            txtiva.TabIndex = 4;
            txtiva.TextAlign = HorizontalAlignment.Center;
            // 
            // lblDescFlia
            // 
            lblDescFlia.AutoSize = true;
            lblDescFlia.Enabled = false;
            lblDescFlia.Font = new Font("MS Reference Sans Serif", 8.25F);
            lblDescFlia.ForeColor = Color.White;
            lblDescFlia.Location = new Point(184, 72);
            lblDescFlia.Margin = new Padding(4, 0, 4, 0);
            lblDescFlia.Name = "lblDescFlia";
            lblDescFlia.Size = new Size(74, 15);
            lblDescFlia.TabIndex = 11;
            lblDescFlia.Text = "Desc. Flia:";
            // 
            // txtDescFlia
            // 
            txtDescFlia.Enabled = false;
            txtDescFlia.Font = new Font("MS Reference Sans Serif", 8.25F);
            txtDescFlia.Location = new Point(284, 69);
            txtDescFlia.Margin = new Padding(4, 3, 4, 3);
            txtDescFlia.Name = "txtDescFlia";
            txtDescFlia.Size = new Size(102, 21);
            txtDescFlia.TabIndex = 12;
            txtDescFlia.TextAlign = HorizontalAlignment.Center;
            // 
            // CheckBox1
            // 
            CheckBox1.AutoSize = true;
            CheckBox1.ForeColor = Color.White;
            CheckBox1.Location = new Point(16, 8);
            CheckBox1.Margin = new Padding(4, 3, 4, 3);
            CheckBox1.Name = "CheckBox1";
            CheckBox1.Size = new Size(148, 19);
            CheckBox1.TabIndex = 5;
            CheckBox1.Text = "Habilitar Carga General";
            CheckBox1.UseVisualStyleBackColor = true;
            // 
            // lblbgral
            // 
            lblbgral.AutoSize = true;
            lblbgral.Enabled = false;
            lblbgral.Font = new Font("MS Reference Sans Serif", 8.25F);
            lblbgral.ForeColor = Color.White;
            lblbgral.Location = new Point(184, 18);
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
            txtbgral.Location = new Point(284, 15);
            txtbgral.Margin = new Padding(4, 3, 4, 3);
            txtbgral.Name = "txtbgral";
            txtbgral.Size = new Size(102, 21);
            txtbgral.TabIndex = 2;
            txtbgral.TextAlign = HorizontalAlignment.Center;
            // 
            // btnRestaurar
            // 
            btnRestaurar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnRestaurar.BackColor = Color.Red;
            btnRestaurar.Font = new Font("Calibri", 20.25F, FontStyle.Bold, GraphicsUnit.Point, 0);
            btnRestaurar.ForeColor = Color.White;
            btnRestaurar.Location = new Point(16, 28);
            btnRestaurar.Margin = new Padding(4, 3, 4, 3);
            btnRestaurar.MinimumSize = new Size(58, 32);
            btnRestaurar.Name = "btnRestaurar";
            btnRestaurar.Size = new Size(58, 49);
            btnRestaurar.TabIndex = 8;
            btnRestaurar.Text = "↺";
            ToolTip1.SetToolTip(btnRestaurar, "Restaurar Valores a 0");
            btnRestaurar.UseVisualStyleBackColor = false;
            btnRestaurar.Click += btnRestaurar_Click;
            // 
            // ptbImgNo
            // 
            ptbImgNo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            ptbImgNo.Image = (Image)resources.GetObject("ptbImgNo.Image");
            ptbImgNo.Location = new Point(410, 12);
            ptbImgNo.Margin = new Padding(4, 3, 4, 3);
            ptbImgNo.Name = "ptbImgNo";
            ptbImgNo.Size = new Size(61, 57);
            ptbImgNo.SizeMode = PictureBoxSizeMode.StretchImage;
            ptbImgNo.TabIndex = 10;
            ptbImgNo.TabStop = false;
            ptbImgNo.Visible = false;
            // 
            // ptbImgSi
            // 
            ptbImgSi.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            ptbImgSi.Image = (Image)resources.GetObject("ptbImgSi.Image");
            ptbImgSi.Location = new Point(410, 12);
            ptbImgSi.Margin = new Padding(4, 3, 4, 3);
            ptbImgSi.Name = "ptbImgSi";
            ptbImgSi.Size = new Size(61, 57);
            ptbImgSi.SizeMode = PictureBoxSizeMode.StretchImage;
            ptbImgSi.TabIndex = 9;
            ptbImgSi.TabStop = false;
            ptbImgSi.Visible = false;
            // 
            // btnSalir
            // 
            btnSalir.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnSalir.Font = new Font("Calibri", 11.25F, FontStyle.Bold);
            btnSalir.ForeColor = Color.Purple;
            btnSalir.Location = new Point(547, 12);
            btnSalir.Margin = new Padding(4, 3, 4, 3);
            btnSalir.Name = "btnSalir";
            btnSalir.Size = new Size(58, 60);
            btnSalir.TabIndex = 7;
            btnSalir.Text = "X";
            ToolTip1.SetToolTip(btnSalir, "Salir");
            btnSalir.UseVisualStyleBackColor = true;
            btnSalir.Click += btnSalir_Click;
            // 
            // btnGrabar
            // 
            btnGrabar.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnGrabar.Font = new Font("Calibri", 11.25F, FontStyle.Bold);
            btnGrabar.ForeColor = Color.Purple;
            btnGrabar.Image = (Image)resources.GetObject("btnGrabar.Image");
            btnGrabar.Location = new Point(480, 12);
            btnGrabar.Margin = new Padding(4, 3, 4, 3);
            btnGrabar.Name = "btnGrabar";
            btnGrabar.Size = new Size(58, 60);
            btnGrabar.TabIndex = 6;
            ToolTip1.SetToolTip(btnGrabar, "Grabar");
            btnGrabar.UseVisualStyleBackColor = true;
            btnGrabar.Click += btnGrabar_Click;
            // 
            // frmConfFamilia
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(29, 29, 27);
            ClientSize = new Size(621, 768);
            Controls.Add(tableLayoutPanel1);
            FormBorderStyle = FormBorderStyle.None;
            Margin = new Padding(4, 3, 4, 3);
            Name = "frmConfFamilia";
            StartPosition = FormStartPosition.Manual;
            Text = "Configuración x Familia";
            Load += frmConfFamilia_Load;
            tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)DataGridView1).EndInit();
            pnlBottom.ResumeLayout(false);
            pnlBottom.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)ptbImgNo).EndInit();
            ((System.ComponentModel.ISupportInitialize)ptbImgSi).EndInit();
            ResumeLayout(false);
        }

        #endregion

        private DataGridView DataGridView1;
        private Label lblbgral;
        private TextBox txtbgral;
        private Label lbliva;
        private TextBox txtiva;
        private Label lblDescFlia;
        private TextBox txtDescFlia;
        private CheckBox CheckBox1;
        private Button btnGrabar;
        private Button btnSalir;
        private Button btnRestaurar;
        private PictureBox ptbImgSi;
        private PictureBox ptbImgNo;
        private ToolTip ToolTip1;
        private TableLayoutPanel tableLayoutPanel1;
        private Panel pnlBottom;
    }
}
