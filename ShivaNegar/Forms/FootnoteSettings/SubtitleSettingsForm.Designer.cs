using Microsoft.Office.Interop.Word;

namespace ShivaNegar.Forms.FootnoteSettings
{
    partial class FootnoteSettingsForm
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
        private void InitializeComponent(Document doc)
        {
            this.components = new System.ComponentModel.Container();
            this.timerOpenAnimation = new System.Windows.Forms.Timer(this.components);
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.footnoteSettingsControl = new ShivaNegar.Forms.FootnoteSettings.FootnoteSettingsControl(doc);
            this.SuspendLayout();
            // 
            // timerOpenAnimation
            // 
            this.timerOpenAnimation.Interval = 1;
            this.timerOpenAnimation.Tick += new System.EventHandler(this.timerOpenAnimation_Tick);
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(800, 600);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.footnoteSettingsControl;
            // 
            // DocumentManagerForm
            // 
            this.ClientSize = new System.Drawing.Size(800, 600);
            this.ControlBox = false;
            this.Controls.Add(this.elementHost1);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DocumentManagerForm";
            this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.ShowIcon = false;
            this.Text = "DocumentManagerForm";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer timerOpenAnimation;
        private System.Windows.Forms.Integration.ElementHost elementHost1;
        public FootnoteSettingsControl footnoteSettingsControl;
    }
}