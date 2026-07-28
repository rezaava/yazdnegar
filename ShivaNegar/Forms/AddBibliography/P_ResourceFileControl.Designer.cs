using ShivaNegar.CustomControls;

namespace ShivaNegar.Forms.AddBibliography
{
    partial class P_ResourceFileControl
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
            this.components = new System.ComponentModel.Container();
            this.tooltipStatus = new System.Windows.Forms.ToolTip(this.components);
            this.panelMain = new ShivaNegar.CustomControls.P_Panel();
            this.lblFileName = new System.Windows.Forms.Label();
            this.lblPath = new System.Windows.Forms.Label();
            this.lblStatus = new System.Windows.Forms.Label();
            this.panelPicture = new System.Windows.Forms.Panel();
            this.picFileExtension = new System.Windows.Forms.PictureBox();
            this.panelRemoveButton = new System.Windows.Forms.Panel();
            this.panelRemoveButtonContainer = new ShivaNegar.CustomControls.P_Panel();
            this.picRemoveControl = new System.Windows.Forms.PictureBox();
            this.panelMain.SuspendLayout();
            this.panelPicture.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picFileExtension)).BeginInit();
            this.panelRemoveButton.SuspendLayout();
            this.panelRemoveButtonContainer.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picRemoveControl)).BeginInit();
            this.SuspendLayout();
            // 
            // panelMain
            // 
            this.panelMain.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
            this.panelMain.BorderColor = System.Drawing.Color.White;
            this.panelMain.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.panelMain.BorderRadius = 15;
            this.panelMain.BorderSize = 1;
            this.panelMain.Controls.Add(this.lblFileName);
            this.panelMain.Controls.Add(this.lblPath);
            this.panelMain.Controls.Add(this.lblStatus);
            this.panelMain.Controls.Add(this.panelPicture);
            this.panelMain.Controls.Add(this.panelRemoveButton);
            this.panelMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMain.Location = new System.Drawing.Point(0, 0);
            this.panelMain.Name = "panelMain";
            this.panelMain.Padding = new System.Windows.Forms.Padding(7);
            this.panelMain.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.All;
            this.panelMain.Size = new System.Drawing.Size(494, 80);
            this.panelMain.TabIndex = 0;
            // 
            // lblFileName
            // 
            this.lblFileName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblFileName.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.lblFileName.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFileName.Location = new System.Drawing.Point(73, 7);
            this.lblFileName.Name = "lblFileName";
            this.lblFileName.Size = new System.Drawing.Size(194, 41);
            this.lblFileName.TabIndex = 5;
            this.lblFileName.Text = "fileName";
            this.lblFileName.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblPath
            // 
            this.lblPath.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lblPath.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.lblPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPath.ForeColor = System.Drawing.Color.DimGray;
            this.lblPath.Location = new System.Drawing.Point(73, 48);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(194, 25);
            this.lblPath.TabIndex = 7;
            this.lblPath.Text = "filePath";
            this.lblPath.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblStatus
            // 
            this.lblStatus.Dock = System.Windows.Forms.DockStyle.Right;
            this.lblStatus.Location = new System.Drawing.Point(267, 7);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(154, 66);
            this.lblStatus.TabIndex = 10;
            this.lblStatus.Text = "Status";
            this.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.lblStatus.FontChanged += new System.EventHandler(this.lblStatus_FontChanged);
            // 
            // panelPicture
            // 
            this.panelPicture.Controls.Add(this.picFileExtension);
            this.panelPicture.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelPicture.Location = new System.Drawing.Point(7, 7);
            this.panelPicture.Name = "panelPicture";
            this.panelPicture.Padding = new System.Windows.Forms.Padding(5);
            this.panelPicture.Size = new System.Drawing.Size(66, 66);
            this.panelPicture.TabIndex = 6;
            // 
            // picFileExtension
            // 
            this.picFileExtension.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picFileExtension.Image = global::ShivaNegar.Properties.Resources.fileIconUnknown;
            this.picFileExtension.Location = new System.Drawing.Point(5, 5);
            this.picFileExtension.Name = "picFileExtension";
            this.picFileExtension.Size = new System.Drawing.Size(56, 56);
            this.picFileExtension.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picFileExtension.TabIndex = 0;
            this.picFileExtension.TabStop = false;
            // 
            // panelRemoveButton
            // 
            this.panelRemoveButton.Controls.Add(this.panelRemoveButtonContainer);
            this.panelRemoveButton.Dock = System.Windows.Forms.DockStyle.Right;
            this.panelRemoveButton.Location = new System.Drawing.Point(421, 7);
            this.panelRemoveButton.Name = "panelRemoveButton";
            this.panelRemoveButton.Padding = new System.Windows.Forms.Padding(10);
            this.panelRemoveButton.Size = new System.Drawing.Size(66, 66);
            this.panelRemoveButton.TabIndex = 9;
            // 
            // panelRemoveButtonContainer
            // 
            this.panelRemoveButtonContainer.BorderCapStyle = System.Drawing.Drawing2D.DashCap.Flat;
            this.panelRemoveButtonContainer.BorderColor = System.Drawing.Color.MediumSlateBlue;
            this.panelRemoveButtonContainer.BorderLineStyle = System.Drawing.Drawing2D.DashStyle.Solid;
            this.panelRemoveButtonContainer.BorderRadius = 0;
            this.panelRemoveButtonContainer.BorderSize = 1;
            this.panelRemoveButtonContainer.Controls.Add(this.picRemoveControl);
            this.panelRemoveButtonContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelRemoveButtonContainer.Location = new System.Drawing.Point(10, 10);
            this.panelRemoveButtonContainer.Name = "panelRemoveButtonContainer";
            this.panelRemoveButtonContainer.PaletteDrawBorder = ShivaNegar.CustomControls.PaletteDrawBorders.All;
            this.panelRemoveButtonContainer.Size = new System.Drawing.Size(46, 46);
            this.panelRemoveButtonContainer.TabIndex = 1;
            // 
            // picRemoveControl
            // 
            this.picRemoveControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.picRemoveControl.Image = global::ShivaNegar.Properties.Resources.recycleBinNormal;
            this.picRemoveControl.Location = new System.Drawing.Point(0, 0);
            this.picRemoveControl.Name = "picRemoveControl";
            this.picRemoveControl.Size = new System.Drawing.Size(46, 46);
            this.picRemoveControl.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picRemoveControl.TabIndex = 0;
            this.picRemoveControl.TabStop = false;
            this.tooltipStatus.SetToolTip(this.picRemoveControl, "حذف منبع از لیست");
            this.picRemoveControl.Click += new System.EventHandler(this.picRemoveControl_Click);
            this.picRemoveControl.MouseEnter += new System.EventHandler(this.picRemoveControl_MouseEnter);
            this.picRemoveControl.MouseLeave += new System.EventHandler(this.picRemoveControl_MouseLeave);
            // 
            // P_ResourceFileControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.Controls.Add(this.panelMain);
            this.Name = "P_ResourceFileControl";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.Size = new System.Drawing.Size(494, 80);
            this.panelMain.ResumeLayout(false);
            this.panelPicture.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picFileExtension)).EndInit();
            this.panelRemoveButton.ResumeLayout(false);
            this.panelRemoveButtonContainer.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picRemoveControl)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.ToolTip tooltipStatus;
        private P_Panel panelMain;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.Panel panelPicture;
        private System.Windows.Forms.PictureBox picFileExtension;
        private System.Windows.Forms.Panel panelRemoveButton;
        private System.Windows.Forms.Label lblStatus;
        private P_Panel panelRemoveButtonContainer;
        private System.Windows.Forms.PictureBox picRemoveControl;
    }
}
