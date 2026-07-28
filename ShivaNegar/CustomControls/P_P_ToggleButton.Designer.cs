using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    partial class P_P_ToggleButton
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
            this.p_Label1 = new Label();
            this.p_ToggleButton1 = new ShivaNegar.CustomControls.P_ToggleButton();
            this.SuspendLayout();
            // 
            // p_Label1
            // 
            this.p_Label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.p_Label1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.p_Label1.Location = new System.Drawing.Point(2, 2);
            this.p_Label1.Name = "p_Label1";
            this.p_Label1.Padding = new System.Windows.Forms.Padding(0, 0, 5, 0);
            this.p_Label1.Size = new System.Drawing.Size(160, 31);
            this.p_Label1.TabIndex = 1;
            this.p_Label1.Text = "عنوان";
            this.p_Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.p_Label1.Click += new System.EventHandler(this.p_Label1_Click);
            // 
            // p_ToggleButton1
            // 
            this.p_ToggleButton1.DisableBackColor = System.Drawing.SystemColors.InactiveCaption;
            this.p_ToggleButton1.DisableToggleColor = System.Drawing.Color.WhiteSmoke;
            this.p_ToggleButton1.Dock = System.Windows.Forms.DockStyle.Right;
            this.p_ToggleButton1.Location = new System.Drawing.Point(162, 2);
            this.p_ToggleButton1.MinimumSize = new System.Drawing.Size(60, 10);
            this.p_ToggleButton1.Name = "p_ToggleButton1";
            this.p_ToggleButton1.OffBackColor = System.Drawing.Color.Gray;
            this.p_ToggleButton1.OffToggleColor = System.Drawing.Color.WhiteSmoke;
            this.p_ToggleButton1.OnBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(6)))), ((int)(((byte)(174)))), ((int)(((byte)(244)))));
            this.p_ToggleButton1.OnToggleColor = System.Drawing.Color.WhiteSmoke;
            this.p_ToggleButton1.Size = new System.Drawing.Size(60, 31);
            this.p_ToggleButton1.TabIndex = 2;
            this.p_ToggleButton1.UseVisualStyleBackColor = true;
            this.p_ToggleButton1.CheckedChanged += new System.EventHandler(this.p_ToggleButton1_CheckedChanged);
            this.p_ToggleButton1.Enter += new System.EventHandler(this.FocusBorderEnter);
            this.p_ToggleButton1.Leave += new System.EventHandler(this.FocusBorderLeave);
            // 
            // P_P_ToggleButton
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.White;
            this.Controls.Add(this.p_Label1);
            this.Controls.Add(this.p_ToggleButton1);
            this.Name = "P_P_ToggleButton";
            this.Padding = new System.Windows.Forms.Padding(2);
            this.Size = new System.Drawing.Size(224, 35);
            this.Paint += new System.Windows.Forms.PaintEventHandler(this.P_P_ToggleButton_Paint);
            this.ResumeLayout(false);

        }

        #endregion
        private Label p_Label1;
        private P_ToggleButton p_ToggleButton1;
    }
}
