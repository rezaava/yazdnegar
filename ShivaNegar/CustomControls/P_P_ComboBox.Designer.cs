using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    partial class P_P_ComboBox
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
            this.panelMain = new System.Windows.Forms.Panel();
            this.p_Panel1 = new ShivaNegar.CustomControls.P_Panel();
            this.comboBox = new ShivaNegar.CustomControls.P_ComboBox();
            this.panelCaption = new ShivaNegar.CustomControls.P_Panel();
            this.lblCaption = new System.Windows.Forms.Label();
            this.panelMain.SuspendLayout();
            this.p_Panel1.SuspendLayout();
            this.panelCaption.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelMain
            // 
            this.panelMain.BackColor = System.Drawing.Color.White;
            this.panelMain.Controls.Add(this.p_Panel1);
            this.panelMain.Controls.Add(this.panelCaption);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMain.Location = new System.Drawing.Point(38, 6);
            this.panelMain.Margin = new System.Windows.Forms.Padding(0);
            this.panelMain.Name = "panelMain";
            this.panelMain.Size = new System.Drawing.Size(544, 53);
            this.panelMain.TabIndex = 85;
            this.panelMain.TabStop = true;
            // 
            // p_Panel1
            // 
            this.p_Panel1.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
            this.p_Panel1.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.p_Panel1.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.p_Panel1.BorderRadius = 15;
            this.p_Panel1.BorderSize = 1;
            this.p_Panel1.Controls.Add(this.comboBox);
            this.p_Panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.p_Panel1.Location = new System.Drawing.Point(0, 0);
            this.p_Panel1.Name = "p_Panel1";
            this.p_Panel1.Padding = new System.Windows.Forms.Padding(8);
            this.p_Panel1.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.TopBottomLeft;
            this.p_Panel1.Size = new System.Drawing.Size(394, 53);
            this.p_Panel1.TabIndex = 70;
            this.p_Panel1.TabStop = true;
            // 
            // comboBox
            // 
            this.comboBox.BorderColor = System.Drawing.Color.Transparent;
            this.comboBox.BorderRadius = 0;
            this.comboBox.BorderSize = 1;
            this.comboBox.ButtonArrowSize = new System.Drawing.Size(10, 4);
            this.comboBox.ButtonArrowThickness = ((byte)(2));
            this.comboBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.comboBox.ForeColor = System.Drawing.SystemColors.WindowText;
            this.comboBox.IconColor = System.Drawing.Color.DimGray;
            this.comboBox.ListBackColor = System.Drawing.Color.White;
            this.comboBox.ListTextColor = System.Drawing.SystemColors.WindowText;
            this.comboBox.Location = new System.Drawing.Point(8, 8);
            this.comboBox.Margin = new System.Windows.Forms.Padding(4);
            this.comboBox.MinimumSize = new System.Drawing.Size(200, 20);
            this.comboBox.Name = "comboBox";
            this.comboBox.Padding = new System.Windows.Forms.Padding(1);
            this.comboBox.PaletteDrawBorder = ShivaNegar.CustomControls.P_ComboBox.PaletteDrawBorders.All;
            this.comboBox.PlaceholderColor = System.Drawing.Color.DarkGray;
            this.comboBox.PlaceholderText = "";
            this.comboBox.Size = new System.Drawing.Size(378, 37);
            this.comboBox.TabIndex = 0;
            this.comboBox.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.comboBox.OnSelectedIndexChanged += new System.EventHandler(this.comboSelectUniversity_OnSelectedIndexChanged);
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
            this.panelCaption.ForeColor = System.Drawing.Color.White;
            this.panelCaption.Location = new System.Drawing.Point(394, 0);
            this.panelCaption.Name = "panelCaption";
            this.panelCaption.Padding = new System.Windows.Forms.Padding(6);
            this.panelCaption.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.TopBottomRight;
            this.panelCaption.Size = new System.Drawing.Size(150, 53);
            this.panelCaption.TabIndex = 71;
            // 
            // lblCaption
            // 
            this.lblCaption.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblCaption.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.lblCaption.Location = new System.Drawing.Point(6, 6);
            this.lblCaption.Name = "lblCaption";
            this.lblCaption.Size = new System.Drawing.Size(138, 41);
            this.lblCaption.TabIndex = 0;
            this.lblCaption.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // P_P_ComboBox
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.panelMain);
            this.Name = "P_P_ComboBox";
            this.Padding = new System.Windows.Forms.Padding(38, 6, 38, 6);
            this.Size = new System.Drawing.Size(620, 65);
            this.panelMain.ResumeLayout(false);
            this.p_Panel1.ResumeLayout(false);
            this.panelCaption.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelMain;
        private P_Panel p_Panel1;
        private P_ComboBox comboBox;
        private P_Panel panelCaption;
        private Label lblCaption;
    }
}
