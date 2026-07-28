using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    partial class P_P_TextBox
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.textBox = new ShivaNegar.CustomControls.P_TextBox();
            this.panelCaption = new ShivaNegar.CustomControls.P_Panel();
            this.lblCaption = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panelCaption.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.textBox);
            this.panel1.Controls.Add(this.panelCaption);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(38, 6);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(545, 53);
            this.panel1.TabIndex = 77;
            // 
            // textBox
            // 
            this.textBox.AutoSize = true;
            this.textBox.BackColor = System.Drawing.SystemColors.Window;
            this.textBox.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.textBox.BorderFocusColor = System.Drawing.Color.Black;
            this.textBox.BorderRadius = 15;
            this.textBox.BorderSize = 1;
            this.textBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox.Font = new System.Drawing.Font("Segoe UI", 9.75F);
            this.textBox.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.textBox.Location = new System.Drawing.Point(0, 0);
            this.textBox.Margin = new System.Windows.Forms.Padding(5);
            this.textBox.Multiline = true;
            this.textBox.Name = "textBox";
            this.textBox.Padding = new System.Windows.Forms.Padding(6);
            this.textBox.PaletteDrawBorder = ShivaNegar.CustomControls.P_TextBox.PaletteDrawBorders.TopBottomLeft;
            this.textBox.PasswordChar = false;
            this.textBox.PlaceholderColor = System.Drawing.Color.DarkGray;
            this.textBox.PlaceholderText = "";
            this.textBox.Size = new System.Drawing.Size(395, 53);
            this.textBox.TabIndex = 0;
            this.textBox.UnderlinedStyle = false;
            this.textBox._TextChanged += new System.EventHandler(this.textBox__TextChanged);
            // 
            // panelCaption
            // 
            this.panelCaption.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(6)))), ((int)(((byte)(174)))), ((int)(((byte)(244)))));
            this.panelCaption.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
            this.panelCaption.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.panelCaption.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.panelCaption.BorderRadius = 15;
            this.panelCaption.BorderSize = 1;
            this.panelCaption.Controls.Add(this.lblCaption);
            this.panelCaption.Dock = System.Windows.Forms.DockStyle.Right;
            this.panelCaption.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.panelCaption.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(6)))), ((int)(((byte)(174)))), ((int)(((byte)(244)))));
            this.panelCaption.Location = new System.Drawing.Point(395, 0);
            this.panelCaption.Margin = new System.Windows.Forms.Padding(4);
            this.panelCaption.Name = "panelCaption";
            this.panelCaption.Padding = new System.Windows.Forms.Padding(6);
            this.panelCaption.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.TopBottomRight;
            this.panelCaption.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.panelCaption.Size = new System.Drawing.Size(150, 53);
            this.panelCaption.TabIndex = 78;
            // 
            // lblCaption
            // 
            this.lblCaption.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblCaption.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblCaption.ForeColor = System.Drawing.Color.White;
            this.lblCaption.Location = new System.Drawing.Point(6, 6);
            this.lblCaption.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCaption.Name = "lblCaption";
            this.lblCaption.Size = new System.Drawing.Size(138, 41);
            this.lblCaption.TabIndex = 0;
            this.lblCaption.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // P_P_TextBox
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.Name = "P_P_TextBox";
            this.Padding = new System.Windows.Forms.Padding(38, 6, 38, 6);
            this.Size = new System.Drawing.Size(621, 65);
            this.Enter += new System.EventHandler(this.textBox_Enter);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panelCaption.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private P_TextBox textBox;
        private P_Panel panelCaption;
        private Label lblCaption;
    }
}
