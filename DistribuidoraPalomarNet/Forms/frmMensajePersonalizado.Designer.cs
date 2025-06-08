namespace DistribuidoraPalomarNet.Forms
{
    partial class frmMensajePersonalizado
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
            lblcaption = new Label();
            LblMensaje = new Label();
            panel1 = new Panel();
            BtnOk = new Button();
            BtnCancelar = new Button();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // lblcaption
            // 
            lblcaption.Font = new Font("Microsoft Sans Serif", 9F);
            lblcaption.Location = new Point(0, 0);
            lblcaption.Name = "lblcaption";
            lblcaption.Size = new Size(376, 63);
            lblcaption.TabIndex = 0;
            // 
            // LblMensaje
            // 
            LblMensaje.Font = new Font("Microsoft Sans Serif", 9.75F);
            LblMensaje.Location = new Point(0, 62);
            LblMensaje.Name = "LblMensaje";
            LblMensaje.Size = new Size(376, 39);
            LblMensaje.TabIndex = 1;
            // 
            // panel1
            // 
            panel1.Controls.Add(BtnCancelar);
            panel1.Controls.Add(BtnOk);
            panel1.Location = new Point(0, 102);
            panel1.Name = "panel1";
            panel1.Size = new Size(384, 33);
            panel1.TabIndex = 2;
            // 
            // BtnOk
            // 
            BtnOk.Location = new Point(215, 5);
            BtnOk.Name = "BtnOk";
            BtnOk.Size = new Size(75, 23);
            BtnOk.TabIndex = 3;
            BtnOk.Text = "Si";
            BtnOk.UseVisualStyleBackColor = true;
            BtnOk.Click += BtnOk_Click;
            // 
            // BtnCancelar
            // 
            BtnCancelar.Location = new Point(301, 5);
            BtnCancelar.Name = "BtnCancelar";
            BtnCancelar.Size = new Size(75, 23);
            BtnCancelar.TabIndex = 4;
            BtnCancelar.Text = "No";
            BtnCancelar.UseVisualStyleBackColor = true;
            // 
            // frmMensajePersonalizado
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(384, 137);
            Controls.Add(panel1);
            Controls.Add(LblMensaje);
            Controls.Add(lblcaption);
            FormBorderStyle = FormBorderStyle.None;
            Name = "frmMensajePersonalizado";
            StartPosition = FormStartPosition.CenterParent;
            Text = "frmMensajePersonalizado";
            panel1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Label lblcaption;
        private Label LblMensaje;
        private Panel panel1;
        private Button BtnCancelar;
        private Button BtnOk;
    }
}