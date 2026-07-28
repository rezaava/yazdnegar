namespace ShivaNegar.Forms.AddBibliography
{
    partial class AddBibliographyForm
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
            this.timerOpenAnimation = new System.Windows.Forms.Timer(this.components);
            this.panelMain = new ShivaNegar.CustomControls.P_Panel();
            this.panelFilePadding = new System.Windows.Forms.Panel();
            this.panelFiles = new System.Windows.Forms.Panel();
            this.btnAlwaysOnTop = new ShivaNegar.CustomControls.P_Button();
            this.panelDragContrainer = new System.Windows.Forms.Panel();
            this.panelImportDragFiles = new ShivaNegar.CustomControls.P_Panel();
            this.panelPictureContainer = new System.Windows.Forms.Panel();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.lblDragStatus = new System.Windows.Forms.Label();
            this.panelTop = new System.Windows.Forms.Panel();
            this.lblWindowName = new System.Windows.Forms.Label();
            this.panelBottom = new ShivaNegar.CustomControls.P_Panel();
            this.btnCancel = new ShivaNegar.CustomControls.P_Button();
            this.btnSubmit = new ShivaNegar.CustomControls.P_Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.panelMain.SuspendLayout();
            this.panelFilePadding.SuspendLayout();
            this.panelDragContrainer.SuspendLayout();
            this.panelImportDragFiles.SuspendLayout();
            this.panelPictureContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.panelTop.SuspendLayout();
            this.panelBottom.SuspendLayout();
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
            this.panelMain.BorderColor = System.Drawing.Color.White;
            this.panelMain.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.panelMain.BorderRadius = 10;
            this.panelMain.BorderSize = 1;
            this.panelMain.Controls.Add(this.panelFilePadding);
            this.panelMain.Controls.Add(this.btnAlwaysOnTop);
            this.panelMain.Controls.Add(this.panelDragContrainer);
            this.panelMain.Controls.Add(this.panelTop);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMain.Location = new System.Drawing.Point(0, 0);
            this.panelMain.Name = "panelMain";
            this.panelMain.Padding = new System.Windows.Forms.Padding(30, 5, 30, 10);
            this.panelMain.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.TopLeftRight;
            this.panelMain.Size = new System.Drawing.Size(500, 350);
            this.panelMain.TabIndex = 31;
            // 
            // panelFilePadding
            // 
            this.panelFilePadding.Controls.Add(this.panelFiles);
            this.panelFilePadding.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelFilePadding.Location = new System.Drawing.Point(30, 204);
            this.panelFilePadding.Name = "panelFilePadding";
            this.panelFilePadding.Padding = new System.Windows.Forms.Padding(0, 5, 0, 5);
            this.panelFilePadding.Size = new System.Drawing.Size(440, 136);
            this.panelFilePadding.TabIndex = 28;
            // 
            // panelFiles
            // 
            this.panelFiles.AutoScroll = true;
            this.panelFiles.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelFiles.Location = new System.Drawing.Point(0, 5);
            this.panelFiles.Name = "panelFiles";
            this.panelFiles.Size = new System.Drawing.Size(440, 126);
            this.panelFiles.TabIndex = 28;
            // 
            // btnAlwaysOnTop
            // 
            this.btnAlwaysOnTop.BackColor = System.Drawing.Color.Transparent;
            this.btnAlwaysOnTop.BackgroundColor = System.Drawing.Color.Transparent;
            this.btnAlwaysOnTop.BackgroundImage = global::ShivaNegar.Properties.Resources.office_push_unpin;
            this.btnAlwaysOnTop.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.btnAlwaysOnTop.BorderColor = System.Drawing.Color.Firebrick;
            this.btnAlwaysOnTop.BorderRadius = 0;
            this.btnAlwaysOnTop.BorderSize = 0;
            this.btnAlwaysOnTop.FlatAppearance.BorderSize = 0;
            this.btnAlwaysOnTop.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnAlwaysOnTop.ForeColor = System.Drawing.Color.Black;
            this.btnAlwaysOnTop.Location = new System.Drawing.Point(5, 5);
            this.btnAlwaysOnTop.Name = "btnAlwaysOnTop";
            this.btnAlwaysOnTop.Size = new System.Drawing.Size(20, 20);
            this.btnAlwaysOnTop.TabIndex = 5;
            this.btnAlwaysOnTop.TextColor = System.Drawing.Color.Black;
            this.btnAlwaysOnTop.UseVisualStyleBackColor = false;
            this.btnAlwaysOnTop.Click += new System.EventHandler(this.btnAlwaysOnTop_Click);
            // 
            // panelDragContrainer
            // 
            this.panelDragContrainer.Controls.Add(this.panelImportDragFiles);
            this.panelDragContrainer.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelDragContrainer.Location = new System.Drawing.Point(30, 39);
            this.panelDragContrainer.Name = "panelDragContrainer";
            this.panelDragContrainer.Padding = new System.Windows.Forms.Padding(0, 5, 0, 5);
            this.panelDragContrainer.Size = new System.Drawing.Size(440, 165);
            this.panelDragContrainer.TabIndex = 30;
            // 
            // panelImportDragFiles
            // 
            this.panelImportDragFiles.BackColor = System.Drawing.Color.White;
            this.panelImportDragFiles.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
            this.panelImportDragFiles.BorderColor = System.Drawing.Color.FromArgb(((int)(((byte)(8)))), ((int)(((byte)(142)))), ((int)(((byte)(254)))));
            this.panelImportDragFiles.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Dash;
            this.panelImportDragFiles.BorderRadius = 15;
            this.panelImportDragFiles.BorderSize = 2;
            this.panelImportDragFiles.Controls.Add(this.panelPictureContainer);
            this.panelImportDragFiles.Controls.Add(this.lblDragStatus);
            this.panelImportDragFiles.Cursor = System.Windows.Forms.Cursors.Hand;
            this.panelImportDragFiles.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelImportDragFiles.Location = new System.Drawing.Point(0, 5);
            this.panelImportDragFiles.Name = "panelImportDragFiles";
            this.panelImportDragFiles.Padding = new System.Windows.Forms.Padding(10);
            this.panelImportDragFiles.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.All;
            this.panelImportDragFiles.Size = new System.Drawing.Size(440, 155);
            this.panelImportDragFiles.TabIndex = 24;
            // 
            // panelPictureContainer
            // 
            this.panelPictureContainer.Controls.Add(this.pictureBox1);
            this.panelPictureContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelPictureContainer.Location = new System.Drawing.Point(10, 10);
            this.panelPictureContainer.Name = "panelPictureContainer";
            this.panelPictureContainer.Padding = new System.Windows.Forms.Padding(30, 20, 30, 10);
            this.panelPictureContainer.Size = new System.Drawing.Size(420, 82);
            this.panelPictureContainer.TabIndex = 2;
            this.panelPictureContainer.Click += new System.EventHandler(this.importDragFiles_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Image = global::ShivaNegar.Properties.Resources.downloadIcon;
            this.pictureBox1.Location = new System.Drawing.Point(30, 20);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(360, 52);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.importDragFiles_Click);
            // 
            // lblDragStatus
            // 
            this.lblDragStatus.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblDragStatus.Font = new System.Drawing.Font("Vazirmatn", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDragStatus.Location = new System.Drawing.Point(10, 92);
            this.lblDragStatus.Name = "lblDragStatus";
            this.lblDragStatus.Size = new System.Drawing.Size(420, 53);
            this.lblDragStatus.TabIndex = 0;
            this.lblDragStatus.Text = "فایل های خود را انتخاب نمایید\r\nیا در اینجا رها کنید";
            this.lblDragStatus.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.lblDragStatus.Click += new System.EventHandler(this.importDragFiles_Click);
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.lblWindowName);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(30, 5);
            this.panelTop.Name = "panelTop";
            this.panelTop.Padding = new System.Windows.Forms.Padding(0, 5, 0, 5);
            this.panelTop.Size = new System.Drawing.Size(440, 34);
            this.panelTop.TabIndex = 29;
            // 
            // lblWindowName
            // 
            this.lblWindowName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblWindowName.Font = new System.Drawing.Font("Vazirmatn", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblWindowName.Location = new System.Drawing.Point(0, 5);
            this.lblWindowName.Name = "lblWindowName";
            this.lblWindowName.Size = new System.Drawing.Size(440, 24);
            this.lblWindowName.TabIndex = 0;
            this.lblWindowName.Text = "وارد کردن منابع";
            this.lblWindowName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panelBottom
            // 
            this.panelBottom.BackColor = System.Drawing.Color.White;
            this.panelBottom.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
            this.panelBottom.BorderColor = System.Drawing.Color.White;
            this.panelBottom.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.panelBottom.BorderRadius = 10;
            this.panelBottom.BorderSize = 1;
            this.panelBottom.Controls.Add(this.btnCancel);
            this.panelBottom.Controls.Add(this.btnSubmit);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelBottom.Location = new System.Drawing.Point(0, 350);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Padding = new System.Windows.Forms.Padding(8, 5, 8, 5);
            this.panelBottom.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.BottomLeftRight;
            this.panelBottom.Size = new System.Drawing.Size(500, 50);
            this.panelBottom.TabIndex = 32;
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(254)))), ((int)(((byte)(142)))), ((int)(((byte)(8)))));
            this.btnCancel.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(254)))), ((int)(((byte)(142)))), ((int)(((byte)(8)))));
            this.btnCancel.BorderColor = System.Drawing.SystemColors.ControlDark;
            this.btnCancel.BorderRadius = 15;
            this.btnCancel.BorderSize = 1;
            this.btnCancel.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnCancel.FlatAppearance.BorderSize = 0;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.Font = new System.Drawing.Font("Vazirmatn", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.ForeColor = System.Drawing.Color.White;
            this.btnCancel.Location = new System.Drawing.Point(8, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(150, 40);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "لغو";
            this.btnCancel.TextColor = System.Drawing.Color.White;
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSubmit
            // 
            this.btnSubmit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(8)))), ((int)(((byte)(142)))), ((int)(((byte)(254)))));
            this.btnSubmit.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(8)))), ((int)(((byte)(142)))), ((int)(((byte)(254)))));
            this.btnSubmit.BorderColor = System.Drawing.SystemColors.ControlDark;
            this.btnSubmit.BorderRadius = 15;
            this.btnSubmit.BorderSize = 1;
            this.btnSubmit.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnSubmit.Enabled = false;
            this.btnSubmit.FlatAppearance.BorderSize = 0;
            this.btnSubmit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSubmit.Font = new System.Drawing.Font("Vazirmatn", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSubmit.ForeColor = System.Drawing.Color.White;
            this.btnSubmit.Location = new System.Drawing.Point(342, 5);
            this.btnSubmit.Name = "btnSubmit";
            this.btnSubmit.Size = new System.Drawing.Size(150, 40);
            this.btnSubmit.TabIndex = 0;
            this.btnSubmit.Text = "وارد کردن منابع";
            this.btnSubmit.TextColor = System.Drawing.Color.White;
            this.btnSubmit.UseVisualStyleBackColor = false;
            this.btnSubmit.Click += new System.EventHandler(this.btnSubmit_Click);
            // 
            // AddBibliography
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(500, 400);
            this.ControlBox = false;
            this.Controls.Add(this.panelMain);
            this.Controls.Add(this.panelBottom);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(350, 350);
            this.Name = "AddBibliography";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "شیوانگار";
            this.panelMain.ResumeLayout(false);
            this.panelFilePadding.ResumeLayout(false);
            this.panelDragContrainer.ResumeLayout(false);
            this.panelImportDragFiles.ResumeLayout(false);
            this.panelPictureContainer.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.panelTop.ResumeLayout(false);
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Timer timerOpenAnimation;
        private System.Windows.Forms.Panel panelFilePadding;
        private System.Windows.Forms.Panel panelFiles;
        private CustomControls.P_Button btnAlwaysOnTop;
        private System.Windows.Forms.Panel panelDragContrainer;
        private CustomControls.P_Panel panelImportDragFiles;
        private System.Windows.Forms.Panel panelPictureContainer;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label lblDragStatus;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Label lblWindowName;
        private CustomControls.P_Panel panelBottom;
        private CustomControls.P_Button btnCancel;
        private CustomControls.P_Button btnSubmit;
        private CustomControls.P_Panel panelMain;
        private System.Windows.Forms.ToolTip toolTip1;
    }
}