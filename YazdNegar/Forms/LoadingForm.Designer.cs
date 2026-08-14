namespace YazdNegar.Forms
{
    partial class LoadingForm
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

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LoadingForm));
            this.timerOpenAnimation = new System.Windows.Forms.Timer(this.components);
            this.timerDotAnimation = new System.Windows.Forms.Timer(this.components);
            this.panelMain = new YazdNegar.CustomControls.P_Panel();
            this.picBoxStatus = new System.Windows.Forms.PictureBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.lblCountOfChanges = new System.Windows.Forms.Label();
            this.panelMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxStatus)).BeginInit();
            this.SuspendLayout();
            // 
            // timerOpenAnimation
            // 
            this.timerOpenAnimation.Interval = 1;
            this.timerOpenAnimation.Tick += new System.EventHandler(this.timerOpenAnimation_Tick);
            // 
            // timerDotAnimation
            // 
            this.timerDotAnimation.Interval = 500;
            this.timerDotAnimation.Tick += new System.EventHandler(this.timerDotAnimation_Tick);
            // 
            // panelMain
            // 
            this.panelMain.BackColor = System.Drawing.Color.White;
            this.panelMain.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
            this.panelMain.BorderColor = System.Drawing.Color.FromArgb(0, 122, 193);
            this.panelMain.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.panelMain.BorderRadius = 20;
            this.panelMain.BorderSize = 2;
            this.panelMain.Controls.Add(this.picBoxStatus);
            this.panelMain.Controls.Add(this.lblStatus);
            this.panelMain.Controls.Add(this.lblCountOfChanges);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMain.Location = new System.Drawing.Point(0, 0);
            this.panelMain.Margin = new System.Windows.Forms.Padding(4);
            this.panelMain.Name = "panelMain";
            this.panelMain.Padding = new System.Windows.Forms.Padding(12);
            this.panelMain.PaletteDrawBorder = YazdNegar.CustomControls.PaletteDrawBorders.All;
            this.panelMain.Size = new System.Drawing.Size(260, 230);
            this.panelMain.TabIndex = 0;
            // 
            // picBoxStatus
            // 
            this.picBoxStatus.BackColor = System.Drawing.Color.Transparent;
            this.picBoxStatus.Image = ((System.Drawing.Image)(resources.GetObject("picBoxStatus.Image")));
            this.picBoxStatus.Location = new System.Drawing.Point(12, 12);
            this.picBoxStatus.Margin = new System.Windows.Forms.Padding(4);
            this.picBoxStatus.Name = "picBoxStatus";
            this.picBoxStatus.Size = new System.Drawing.Size(236, 120);
            this.picBoxStatus.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picBoxStatus.TabIndex = 2;
            this.picBoxStatus.TabStop = false;
            // 
            // lblStatus
            // 
            this.lblStatus.Font = new System.Drawing.Font("Vazirmatn", 10F, System.Drawing.FontStyle.Regular);
            this.lblStatus.ForeColor = System.Drawing.Color.FromArgb(64, 64, 64);
            this.lblStatus.Location = new System.Drawing.Point(12, 140);
            this.lblStatus.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(236, 35);
            this.lblStatus.TabIndex = 3;
            this.lblStatus.Text = "لطفا منتظر بمانید";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblCountOfChanges
            // 
            this.lblCountOfChanges.Font = new System.Drawing.Font("Vazirmatn", 9F, System.Drawing.FontStyle.Regular);
            this.lblCountOfChanges.ForeColor = System.Drawing.Color.FromArgb(0, 122, 193);
            this.lblCountOfChanges.Location = new System.Drawing.Point(12, 180);
            this.lblCountOfChanges.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblCountOfChanges.Name = "lblCountOfChanges";
            this.lblCountOfChanges.Size = new System.Drawing.Size(236, 30);
            this.lblCountOfChanges.TabIndex = 4;
            this.lblCountOfChanges.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LoadingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.FromArgb(240, 242, 245);
            this.ClientSize = new System.Drawing.Size(260, 230);
            this.ControlBox = false;
            this.Controls.Add(this.panelMain);
            this.Font = new System.Drawing.Font("Vazirmatn", 8.25F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LoadingForm";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "مقاله نگار";
            this.TopMost = true;
            this.panelMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picBoxStatus)).EndInit();
            this.ResumeLayout(false);
        }

        private System.Windows.Forms.Timer timerOpenAnimation;
        private System.Windows.Forms.Timer timerDotAnimation;
        private YazdNegar.CustomControls.P_Panel panelMain;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label lblCountOfChanges;
        private System.Windows.Forms.PictureBox picBoxStatus;
    }
}