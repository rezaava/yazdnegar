using Microsoft.Office.Interop.Word;
using ShivaNegar.Forms.CaptionSettings;
using System.Windows.Controls;

namespace ShivaNegar.TaskPanes.ChatBoxNetworking
{
	[System.ComponentModel.ToolboxItemAttribute(false)]
	partial class ChatBoxNetworkingTaskPane
    {
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary> 
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose(bool disposing)
		{
			if(disposing && (components != null))
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
			this.elementHost = new System.Windows.Forms.Integration.ElementHost();
			this.chatBoxControl = new ShivaNegar.TaskPanes.ChatBoxNetworking.ChatBoxNetworkingControl();
			this.SuspendLayout();
			// 
			// elementHost
			// 
			this.elementHost.Dock = System.Windows.Forms.DockStyle.Fill;
			this.elementHost.Location = new System.Drawing.Point(0 , 0);
			this.elementHost.Name = "elementHost";
			this.elementHost.Size = new System.Drawing.Size(350 , 500);
			this.elementHost.TabIndex = 0;
			this.elementHost.Text = "elementHost1";
			this.elementHost.Child = this.chatBoxControl;
			// 
			// CrossReferenceTaskPane
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(96F , 96F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
			this.Controls.Add(this.elementHost);
			this.DoubleBuffered = true;
			this.MinimumSize = new System.Drawing.Size(250 , 100);
			this.Name = "CrossReferenceTaskPane";
			this.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
			this.Size = new System.Drawing.Size(350 , 500);
			this.ResumeLayout(false);

		}


		#endregion

		private System.Windows.Forms.Integration.ElementHost elementHost;
		internal ChatBoxNetworkingControl chatBoxControl;
	}
}
