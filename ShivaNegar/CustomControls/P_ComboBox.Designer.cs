
namespace ShivaNegar.CustomControls
{
    partial class P_ComboBox
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
        /// 
        private void InitializeComponent()
        {
            this.mainPanel = new System.Windows.Forms.Panel();
            this.lblText = new System.Windows.Forms.Label();
            this.txtText = new ShivaNegar.CustomControls.P_TextBox();
            this.btnIcon = new System.Windows.Forms.Button();
            this.cmbList = new System.Windows.Forms.ComboBox();
            this.mainPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainPanel
            // 
            this.mainPanel.BackColor = System.Drawing.SystemColors.Window;
            this.mainPanel.Controls.Add(this.lblText);
            this.mainPanel.Controls.Add(this.txtText);
            this.mainPanel.Controls.Add(this.btnIcon);
            this.mainPanel.Controls.Add(this.cmbList);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.ForeColor = System.Drawing.SystemColors.WindowText;
            this.mainPanel.Location = new System.Drawing.Point(10, 7);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Size = new System.Drawing.Size(230, 24);
            this.mainPanel.TabIndex = 0;
            this.mainPanel.TabStop = true;
            // 
            // lblText
            // 
            this.lblText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.lblText.Location = new System.Drawing.Point(30, 0);
            this.lblText.Name = "lblText";
            this.lblText.Padding = new System.Windows.Forms.Padding(8, 0, 0, 0);
            this.lblText.Size = new System.Drawing.Size(200, 24);
            this.lblText.TabIndex = 0;
            this.lblText.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtText
            // 
            this.txtText.BackColor = System.Drawing.SystemColors.Window;
            this.txtText.BorderColor = System.Drawing.SystemColors.ActiveBorder;
            this.txtText.BorderFocusColor = System.Drawing.Color.Black;
            this.txtText.BorderRadius = 0;
            this.txtText.BorderSize = 0;
            this.txtText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtText.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtText.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.txtText.Location = new System.Drawing.Point(30, 0);
            this.txtText.Margin = new System.Windows.Forms.Padding(4);
            this.txtText.Multiline = true;
            this.txtText.Name = "txtText";
            this.txtText.Padding = new System.Windows.Forms.Padding(10, 7, 10, 7);
            this.txtText.PaletteDrawBorder = ShivaNegar.CustomControls.P_TextBox.PaletteDrawBorders.All;
            this.txtText.PasswordChar = false;
            this.txtText.PlaceholderColor = System.Drawing.Color.DarkGray;
            this.txtText.PlaceholderText = "";
            this.txtText.Size = new System.Drawing.Size(200, 24);
            this.txtText.TabIndex = 1;
            this.txtText.UnderlinedStyle = false;
            // 
            // btnIcon
            // 
            this.btnIcon.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnIcon.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnIcon.FlatAppearance.BorderSize = 0;
            this.btnIcon.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnIcon.Location = new System.Drawing.Point(0, 0);
            this.btnIcon.Name = "btnIcon";
            this.btnIcon.Size = new System.Drawing.Size(30, 24);
            this.btnIcon.TabIndex = 0;
            this.btnIcon.UseVisualStyleBackColor = true;
            // 
            // cmbList
            // 
            this.cmbList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.cmbList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbList.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cmbList.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.cmbList.FormattingEnabled = true;
            this.cmbList.ItemHeight = 16;
            this.cmbList.Location = new System.Drawing.Point(0, 0);
            this.cmbList.MinimumSize = new System.Drawing.Size(200, 0);
            this.cmbList.Name = "cmbList";
            this.cmbList.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.cmbList.Size = new System.Drawing.Size(230, 24);
            this.cmbList.TabIndex = 0;
            this.cmbList.TabStop = false;
            // 
            // P_ComboBox
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.Controls.Add(this.mainPanel);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.ForeColor = System.Drawing.Color.DimGray;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MinimumSize = new System.Drawing.Size(200, 30);
            this.Name = "P_ComboBox";
            this.Padding = new System.Windows.Forms.Padding(10, 7, 10, 7);
            this.Size = new System.Drawing.Size(250, 38);
            this.mainPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Label lblText;
        private System.Windows.Forms.Button btnIcon;
        private System.Windows.Forms.ComboBox cmbList;
        private P_TextBox txtText;
    }
}
