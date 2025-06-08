namespace DistribuidoraPalomarNet
{
    partial class frmImagen
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
            ptbImagen = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)ptbImagen).BeginInit();
            SuspendLayout();
            
            // ptbImagen
            ptbImagen.Dock = DockStyle.Fill;
            ptbImagen.Location = new Point(0, 0);
            ptbImagen.Name = "ptbImagen";
            ptbImagen.Size = new Size(800, 450);
            ptbImagen.SizeMode = PictureBoxSizeMode.Zoom;
            ptbImagen.TabIndex = 0;
            ptbImagen.TabStop = false;
            
            // frmImagen
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(ptbImagen);
            Name = "frmImagen";
            Text = "Imagen";
            Load += new System.EventHandler(frmImagen_Load);
            ((System.ComponentModel.ISupportInitialize)ptbImagen).EndInit();
            ResumeLayout(false);
        }

        #endregion

        public PictureBox ptbImagen;
    }
}
