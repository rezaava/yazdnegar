using System.Windows.Forms;

namespace ShivaNegar.Forms
{
    partial class SplashScreen
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
			this.components = new System.ComponentModel.Container();
			this.timerDotAnimation = new System.Windows.Forms.Timer(this.components);
			this.timerOpenAnimation = new System.Windows.Forms.Timer(this.components);
			this.timerLifeTime = new System.Windows.Forms.Timer(this.components);
			this.panelMain = new ShivaNegar.CustomControls.P_Panel();
			this.lblVersion = new System.Windows.Forms.Label();
			this.lblStatus = new System.Windows.Forms.Label();
			this.panelMain.SuspendLayout();
			this.SuspendLayout();
			// 
			// timerDotAnimation
			// 
			this.timerDotAnimation.Enabled = true;
			this.timerDotAnimation.Interval = 150;
			this.timerDotAnimation.Tick += new System.EventHandler(this.timerDotAnimation_Tick);
			// 
			// timerOpenAnimation
			// 
			this.timerOpenAnimation.Interval = 10;
			this.timerOpenAnimation.Tick += new System.EventHandler(this.timerOpenAnimation_Tick);
			// 
			// timerLifeTime
			// 
			this.timerLifeTime.Interval = 8000;
			this.timerLifeTime.Tick += new System.EventHandler(this.timerLifeTime_Tick);
			// 
			// panelMain
			// 
			this.panelMain.BackColor = System.Drawing.Color.White;
			this.panelMain.BackgroundImage = global::ShivaNegar.Properties.Resources.SplashBackground;
			this.panelMain.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
			this.panelMain.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
			this.panelMain.BorderColor = System.Drawing.Color.White;
			this.panelMain.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
			this.panelMain.BorderRadius = 0;
			this.panelMain.BorderSize = 1;
			this.panelMain.Controls.Add(this.lblVersion);
			this.panelMain.Controls.Add(this.lblStatus);
			this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
			this.panelMain.Location = new System.Drawing.Point(0, 0);
			this.panelMain.Name = "panelMain";
			this.panelMain.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.All;
			this.panelMain.Size = new System.Drawing.Size(450, 250);
			this.panelMain.TabIndex = 2;
			// 
			// lblVersion
			// 
			this.lblVersion.BackColor = System.Drawing.Color.Transparent;
			this.lblVersion.Font = new System.Drawing.Font("Vazirmatn", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblVersion.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(6)))), ((int)(((byte)(174)))), ((int)(((byte)(244)))));
			this.lblVersion.Location = new System.Drawing.Point(1, 208);
			this.lblVersion.Name = "lblVersion";
			this.lblVersion.Size = new System.Drawing.Size(98, 36);
			this.lblVersion.TabIndex = 2;
			this.lblVersion.Text = "نسخه: ۱.۰.۱.۳۰";
			this.lblVersion.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// lblStatus
			// 
			this.lblStatus.BackColor = System.Drawing.Color.Transparent;
			this.lblStatus.Font = new System.Drawing.Font("Vazirmatn", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(6)))), ((int)(((byte)(174)))), ((int)(((byte)(244)))));
			this.lblStatus.Location = new System.Drawing.Point(187, 208);
			this.lblStatus.Name = "lblStatus";
			this.lblStatus.Size = new System.Drawing.Size(254, 36);
			this.lblStatus.TabIndex = 0;
			this.lblStatus.Text = "در حال شروع برنامه";
			this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// SplashScreen
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(450, 250);
			this.ControlBox = false;
			this.Controls.Add(this.panelMain);
			this.Cursor = System.Windows.Forms.Cursors.WaitCursor;
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "SplashScreen";
			this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.ShowIcon = false;
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "ShivaNeger_SplashScreen";
			this.panelMain.ResumeLayout(false);
			this.ResumeLayout(false);

        }

        #endregion
        private ShivaNegar.CustomControls.P_Panel panelMain;
        private Label lblVersion;
        private Timer timerDotAnimation;
        private Timer timerOpenAnimation;
        private Label lblStatus;
        private Timer timerLifeTime;
    }
}