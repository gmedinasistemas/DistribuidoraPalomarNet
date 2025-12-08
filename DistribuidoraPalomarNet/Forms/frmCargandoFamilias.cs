using System;
using System.Drawing;
using System.Windows.Forms;

namespace DistribuidoraPalomarNet
{
    public class frmCargandoFamilias : Form
    {
        private Label lblMensaje;
        private ProgressBar progressBar;

        public frmCargandoFamilias()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.lblMensaje = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // lblMensaje
            // 
            this.lblMensaje.AutoSize = true;
            this.lblMensaje.Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMensaje.ForeColor = System.Drawing.Color.Navy;
            this.lblMensaje.Location = new System.Drawing.Point(25, 20);
            this.lblMensaje.Name = "lblMensaje";
            this.lblMensaje.Size = new System.Drawing.Size(180, 20);
            this.lblMensaje.TabIndex = 0;
            this.lblMensaje.Text = "...Cargando Familias...";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(20, 55);
            this.progressBar.MarqueeAnimationSpeed = 30;
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(230, 18);
            this.progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar.TabIndex = 1;
            // 
            // frmCargandoFamilias
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(270, 95);
            this.ControlBox = false;
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblMensaje);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmCargandoFamilias";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Por favor espere";
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
