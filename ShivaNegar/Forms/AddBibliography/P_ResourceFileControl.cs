using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using ShivaNegar.CustomControls;
using ShivaNegar.Forms.AddBibliography.Utilities;

namespace ShivaNegar.Forms.AddBibliography
{
    public partial class P_ResourceFileControl : UserControl
    {
        public enum FileType
        {
            txt,
            bib,


            ris,
            enw,
            unknown
        }
        public enum FileStatus
        {
            Success,
            Fail
        }

        private FileStatus status;
        private FileType type;


        [Category("Dedicated")]
        public Color BackgroundColor { get => panelMain.BackColor; set => panelMain.BackColor = value; }

        [Category("Dedicated")]
        public Padding PaddingPanelMain
        {
            get { return panelMain.Padding; }
            set
            {
                panelMain.Padding = value;
                this.Invalidate();
            }
        }

        [Category("Dedicated")]
        public Padding PaddingPicFileExtension { get => panelPicture.Padding; set => panelPicture.Padding = value; }
        [Category("Dedicated")]
        public Padding PaddingRemoveButton { get => panelRemoveButton.Padding; set => panelRemoveButton.Padding = value; }
        [Category("Dedicated")]
        public Size StatusBoxSize { get => lblStatus.Size; set => lblStatus.Size = value; }

        [Category("Dedicated")]
        public string FileName { get => lblFileName.Text; set => lblFileName.Text = value; }
        [Category("Dedicated")]
        public string FilePath
        {
            get => lblPath.Text;
            set
            {
                lblPath.Text = value;
                this.tooltipStatus.SetToolTip(this.lblPath, value);
                lblPath.Height = (int)lblPath.Font.GetHeight();
            }
        }

        [Category("Dedicated")]
        public FileStatus StatusColor
        {
            get { return status; }
            set
            {
                status = value;
                if (value == FileStatus.Success)
                {
                    lblStatus.ForeColor = Color.Green;
                }
                else if (value == FileStatus.Fail)
                {
                    lblStatus.ForeColor = Color.Red;
                }
            }
        }
        [Category("Dedicated")]
        public string StatusText
        {
            get { return lblStatus.Text; }
            set
            {
                lblStatus.Text = value;

                Graphics gfx = Graphics.FromImage(new Bitmap(1, 1));
                //lblStatus.Width = (int)gfx.MeasureString(lblStatus.Text, lblStatus.Font).Width;
            }
        }
        [Category("Dedicated")]
        public FileType Type
        {
            get { return type; }
            set
            {
                type = value;
                if (value == FileType.txt)
                {
                    picFileExtension.Image = Properties.Resources.fileIcontxt;
                }
                else if (value == FileType.bib)
                {
                    picFileExtension.Image = Properties.Resources.fileIconbib;
                }
                else if (value == FileType.ris || value == FileType.enw)
                {
                    picFileExtension.Image = IconHelper.GetIconOfFile(FilePath, false, false).ToBitmap();
                }
                else if (value == FileType.unknown)
                {
                    picFileExtension.Image = Properties.Resources.fileIconUnknown;
                }

            }
        }

        #region -> Panel Properties
        [Category("Dedicated(Container)")]
        public Color BorderColor
        {
            get { return panelMain.BorderColor; }
            set
            {
                panelMain.BorderColor = value;
                this.Invalidate();
            }
        }
        [Category("Dedicated(Container)")]
        public int BorderSize
        {
            get { return panelMain.BorderSize; }
            set
            {
                if (value >= 1)
                {
                    panelMain.BorderSize = value;
                    if (panelMain.BorderRadius > 0)
                    {
                        this.Padding = new Padding(panelMain.BorderSize * 2);//Border Size
                    }
                    else
                    {
                        this.Padding = new Padding(panelMain.BorderSize);//Border Size
                    }
                    this.Invalidate();
                }
            }
        }
        [Category("Dedicated(Container)")]
        public int BorderRadius
        {
            get { return panelMain.BorderRadius; }
            set
            {
                if (value >= 0)
                {
                    panelMain.BorderRadius = value;
                    this.Invalidate();//Redraw control
                }
            }
        }
        [Category("Dedicated(Container)")]
        public PaletteDrawBorders PaletteDrawBorder
        {
            get => panelMain.PaletteDrawBorder; set => panelMain.PaletteDrawBorder = value;
        }



        [Category("Dedicated(RemoveButton)")]
        public int BorderSizeRemoveButton
        {
            get { return panelRemoveButtonContainer.BorderSize; }
            set
            {
                if (value >= 1)
                {
                    panelRemoveButtonContainer.BorderSize = value;
                    if (panelRemoveButtonContainer.BorderRadius > 0)
                    {
                        this.Padding = new Padding(panelRemoveButtonContainer.BorderSize * 2);//Border Size
                    }
                    else
                    {
                        this.Padding = new Padding(panelRemoveButtonContainer.BorderSize);//Border Size
                    }
                    this.Invalidate();
                }
            }
        }
        [Category("Dedicated(RemoveButton)")]
        public Color BorderColorRemoveButton
        {
            get { return panelRemoveButtonContainer.BorderColor; }
            set
            {
                panelRemoveButtonContainer.BorderColor = value;
                this.Invalidate();
            }
        }
        [Category("Dedicated(RemoveButton)")]
        public int BorderRadiusRemoveButton
        {
            get { return panelRemoveButtonContainer.BorderRadius; }
            set
            {
                if (value >= 0)
                {
                    panelRemoveButtonContainer.BorderRadius = value;
                    this.Invalidate();//Redraw control
                }
            }
        }
        [Category("Dedicated(RemoveButton)")]
        public PaletteDrawBorders PaletteDrawBorderRemoveButton
        {
            get => panelRemoveButtonContainer.PaletteDrawBorder; set => panelRemoveButtonContainer.PaletteDrawBorder = value;
        }
        #endregion


        public P_ResourceFileControl()
        {
            InitializeComponent();

            BackgroundColor = Color.White;
            BorderColor = Color.FromArgb(255, 8, 142, 254);
            BorderSize = 2;
            Type = FileType.txt;
            StatusColor = FileStatus.Success;

        }

        public override Font Font
        {

            get => base.Font;
            set
            {
                base.Font = value;
                lblFileName.Font = new Font(value.FontFamily, lblFileName.Font.Size);
                lblPath.Font = new Font(value.FontFamily, lblPath.Font.Size);
                lblStatus.Font = new Font(value.FontFamily, lblStatus.Font.Size);
            }
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            panelPicture.Width = panelPicture.Height;
            panelRemoveButton.Width = panelRemoveButton.Height;

            int fontSize = (int)lblPath.Font.GetHeight() + (int)lblFileName.Font.GetHeight() + Padding.Top + Padding.Bottom + PaddingPanelMain.Top + PaddingPanelMain.Bottom;
            if (fontSize > this.Height)
            {
                lblPath.Height = (int)lblPath.Font.GetHeight();
                int statusFontSize = ((int)lblStatus.Font.GetHeight() * 2) + Padding.Top + Padding.Bottom + PaddingPanelMain.Top + PaddingPanelMain.Bottom;
                if (statusFontSize > fontSize)
                    this.Height = statusFontSize;
                else
                    this.Height = fontSize;
            }
        }

        [Category("Dedicated")]
        public delegate void btnRemoveControl_Click(object sender, EventArgs e);
        // The event
        public event btnRemoveControl_Click onRemoveControlClick;

        private void picRemoveControl_Click(object sender, EventArgs e)
        {
            // Check if there are any Subscribers
            if (onRemoveControlClick != null)
            {
                // Call the Event
                onRemoveControlClick(this, e);
            }
        }

        private void picRemoveControl_MouseEnter(object sender, EventArgs e)
        {
            picRemoveControl.Image = Properties.Resources.recycleBinOnHover;
            picRemoveControl.BackColor = Color.Gainsboro;
        }

        private void picRemoveControl_MouseLeave(object sender, EventArgs e)
        {
            picRemoveControl.Image = Properties.Resources.recycleBinNormal;
            picRemoveControl.BackColor = this.BackColor;


        }

        private void lblStatus_FontChanged(object sender, EventArgs e)
        {
            lblPath.Height = (int)lblPath.Font.GetHeight();
        }
    }
}
