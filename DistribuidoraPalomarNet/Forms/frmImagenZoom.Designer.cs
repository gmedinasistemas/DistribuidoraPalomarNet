namespace ImageViewerApp
{
    partial class frmImagenZoom
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            pictureBox1 = new PictureBox();
            btnPrevious = new Button();
            btnNext = new Button();
            txtCode = new TextBox();
            btnSearch = new Button();
            label1 = new Label();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            // pictureBox1
            // 
            pictureBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            pictureBox1.Location = new Point(12, 41);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(760, 458);
            pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;
            pictureBox1.TabIndex = 0;
            pictureBox1.TabStop = false;
            // 
            // btnPrevious
            // 
            btnPrevious.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            btnPrevious.Location = new Point(12, 505);
            btnPrevious.Name = "btnPrevious";
            btnPrevious.Size = new Size(94, 29);
            btnPrevious.TabIndex = 1;
            btnPrevious.Text = "Anterior";
            btnPrevious.Click += btnPrevious_Click;
            // 
            // btnNext
            // 
            btnNext.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnNext.Location = new Point(678, 505);
            btnNext.Name = "btnNext";
            btnNext.Size = new Size(94, 29);
            btnNext.TabIndex = 2;
            btnNext.Text = "Siguiente";
            btnNext.Click += btnNext_Click;
            // 
            // txtCode
            // 
            txtCode.Location = new Point(89, 12);
            txtCode.Name = "txtCode";
            txtCode.Size = new Size(125, 23);
            txtCode.TabIndex = 3;
            txtCode.Visible = false;
            // 
            // btnSearch
            // 
            btnSearch.Location = new Point(220, 11);
            btnSearch.Name = "btnSearch";
            btnSearch.Size = new Size(94, 23);
            btnSearch.TabIndex = 4;
            btnSearch.Text = "Buscar";
            btnSearch.Visible = false;
            btnSearch.Click += btnSearch_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(12, 15);
            label1.Name = "label1";
            label1.Size = new Size(49, 15);
            label1.TabIndex = 5;
            label1.Text = "Código:";
            label1.Visible = false;
            // 
            // frmImagenZoom
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(784, 546);
            Controls.Add(label1);
            Controls.Add(btnSearch);
            Controls.Add(txtCode);
            Controls.Add(btnNext);
            Controls.Add(btnPrevious);
            Controls.Add(pictureBox1);
            Name = "frmImagenZoom";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Visor de Imágenes";
            Load += frmImagenZoom_Load;
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnPrevious;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Label label1;

        #endregion

        public TextBox txtCode;
        public Button btnSearch;
    }
}
