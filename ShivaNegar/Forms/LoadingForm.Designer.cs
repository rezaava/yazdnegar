using ShivaNegar.CustomControls;

namespace ShivaNegar.Forms
{
    partial class LoadingForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(LoadingForm));
            this.timerOpenAnimation = new System.Windows.Forms.Timer(this.components);
            this.panelMain = new ShivaNegar.CustomControls.P_Panel();
            this.picBoxStatus = new System.Windows.Forms.PictureBox();
            this.lblStatus = new System.Windows.Forms.Label();
            this.timerDotAnimation = new System.Windows.Forms.Timer(this.components);
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
            // panelMain
            // 
            this.panelMain.BackColor = System.Drawing.Color.White;
            this.panelMain.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
            this.panelMain.BorderColor = System.Drawing.SystemColors.Control;
            this.panelMain.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.panelMain.BorderRadius = 0;
            this.panelMain.BorderSize = 1;
            this.panelMain.Controls.Add(this.picBoxStatus);
            this.panelMain.Controls.Add(this.lblStatus);
            this.panelMain.Controls.Add(this.lblCountOfChanges);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMain.Location = new System.Drawing.Point(0, 0);
            this.panelMain.Name = "panelMain";
            this.panelMain.Padding = new System.Windows.Forms.Padding(10);
            this.panelMain.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.All;
            this.panelMain.Size = new System.Drawing.Size(300, 300);
            this.panelMain.TabIndex = 0;
            // 
            // picBoxStatus
            // 
            this.picBoxStatus.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picBoxStatus.Image = ((System.Drawing.Image)(resources.GetObject("picBoxStatus.Image")));
            this.picBoxStatus.Location = new System.Drawing.Point(10, 10);
            this.picBoxStatus.Name = "picBoxStatus";
            this.picBoxStatus.Size = new System.Drawing.Size(280, 228);
            this.picBoxStatus.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picBoxStatus.TabIndex = 2;
            this.picBoxStatus.TabStop = false;
            // 
            // lblStatus
            // 
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblStatus.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.Location = new System.Drawing.Point(10, 238);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(280, 26);
            this.lblStatus.TabIndex = 3;
            this.lblStatus.Text = "لطفا منتظر بمانید";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // timerDotAnimation
            // 
            this.timerDotAnimation.Interval = 500;
            this.timerDotAnimation.Tick += new System.EventHandler(this.timerDotAnimation_Tick);
            // 
            // lblCountOfChanges
            // 
            this.lblCountOfChanges.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblCountOfChanges.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCountOfChanges.Location = new System.Drawing.Point(10, 264);
            this.lblCountOfChanges.Name = "lblCountOfChanges";
            this.lblCountOfChanges.Size = new System.Drawing.Size(280, 26);
            this.lblCountOfChanges.TabIndex = 4;
            this.lblCountOfChanges.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // LoadingForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(300, 300);
            this.ControlBox = false;
            this.Controls.Add(this.panelMain);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "LoadingForm";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "شیوانگار";
            this.panelMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picBoxStatus)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Timer timerOpenAnimation;
        private P_Panel panelMain;
        private System.Windows.Forms.PictureBox picBoxStatus;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Timer timerDotAnimation;
        private System.Windows.Forms.Label lblCountOfChanges;
    }
}