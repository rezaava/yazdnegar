using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    [DefaultEvent("CheckedChanged")]
    public partial class P_P_ToggleButton : UserControl
    {
        private Color _BorderColor = Color.White;
        private Color borderColor = Color.Black;

        public event EventHandler CheckedChanged;


        public P_P_ToggleButton()
        {
            InitializeComponent();
            _BorderColor = this.BackColor;
            //BorderColor = Color.Black;
            this.Refresh();
        }

        #region -> Properties

        [Category("Dedicated Style")]
        public Color OnBackColor { get => p_ToggleButton1.OnBackColor; set { p_ToggleButton1.OnBackColor = value; this.Invalidate(); } }
        [Category("Dedicated Style")]
        public Color OnToggleColor { get => p_ToggleButton1.OnToggleColor; set { p_ToggleButton1.OnToggleColor = value; this.Invalidate(); } }
        [Category("Dedicated Style")]
        public Color OffBackColor { get => p_ToggleButton1.OffBackColor; set { p_ToggleButton1.OffBackColor = value; this.Invalidate(); } }
        [Category("Dedicated Style")]
        public Color OffToggleColor { get => p_ToggleButton1.OffToggleColor; set { p_ToggleButton1.OffToggleColor = value; this.Invalidate(); } }
        [Category("Dedicated Style")]
        [DefaultValue(true)]
        public bool SolidStyle { get => p_ToggleButton1.SolidStyle; set { p_ToggleButton1.SolidStyle = value; this.Invalidate(); } }
        [Category("Dedicated Style")]
        public Color DisableBackColor { get => p_ToggleButton1.DisableBackColor; set => p_ToggleButton1.DisableBackColor = value; }
        [Category("Dedicated Style")]
        public Color DisableToggleColor { get => p_ToggleButton1.DisableToggleColor; set => p_ToggleButton1.DisableToggleColor = value; }



        [Category("Dedicated")]
        public ContentAlignment TextAlign
        {
            get => p_Label1.TextAlign;
            set => p_Label1.TextAlign = value;
        }
        [EditorBrowsable(EditorBrowsableState.Always)]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        [Bindable(true)]
        [EditorAttribute(
    "System.ComponentModel.Design.MultilineStringEditor, System.Design",
    "System.Drawing.Design.UITypeEditor")]
        [Category("Dedicated")]
        public override string Text
        {
            get => p_Label1.Text;
            set => p_Label1.Text = value;
        }
        [Category("Dedicated")]
        public new bool Enabled
        {
            get => p_ToggleButton1.Enabled;
            set
            {
                p_ToggleButton1.Enabled = value;
                this.TabStop = value;
            }
        }
        [Category("Dedicated")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        [Bindable(true)]
        public override Font Font
        {
            get => p_Label1.Font;
            set => p_Label1.Font = value;
        }
        [Category("Dedicated")]
        public Size ToggleButtonSize
        {
            get => p_ToggleButton1.Size;
            set => p_ToggleButton1.Size = value;
        }
        [Category("Dedicated")]
        public Size ToggleButtonMinimumSize
        {
            get => p_ToggleButton1.MinimumSize;
            set => p_ToggleButton1.MinimumSize = value;
        }
        [Category("Dedicated")]
        public Color BorderColor { get => this.borderColor; set => this.borderColor = value; }

        [Category("Dedicated")]
        public override RightToLeft RightToLeft
        {
            get => base.RightToLeft;
            set
            {
                base.RightToLeft = value;

                if (value == RightToLeft.Yes)
                    p_ToggleButton1.Dock = DockStyle.Right;
                else if (value == RightToLeft.No)
                    p_ToggleButton1.Dock = DockStyle.Left;
                else if (value == RightToLeft.Inherit)
                    if (this.RightToLeft == RightToLeft.Yes)
                        p_ToggleButton1.Dock = DockStyle.Right;
                    else if (this.RightToLeft == RightToLeft.No)
                        p_ToggleButton1.Dock = DockStyle.Left;
            }
        }
        [Category("Dedicated")]
        public bool Checked
        {
            get => p_ToggleButton1.Checked;
            set => p_ToggleButton1.Checked = value;
        }
        #endregion

        #region -> Events
        private void p_Label1_Click(object sender, EventArgs e)
        {
            if (p_ToggleButton1.Enabled)
            {
                p_ToggleButton1.Checked = !p_ToggleButton1.Checked;
                this.Refresh();
                this.Focus();
            }
        }
        private void p_ToggleButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckedChanged != null)
                CheckedChanged.Invoke(sender, e);
        }
        private void FocusBorderEnter(object sender, EventArgs e)
        {
            _BorderColor = BorderColor;
            this.Refresh();
        }
        private void FocusBorderLeave(object sender, EventArgs e)
        {
            _BorderColor = this.BackColor;
            this.Refresh();
        }
        private void P_P_ToggleButton_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, this.ClientRectangle, _BorderColor, ButtonBorderStyle.Solid);
        }
        #endregion
    }
}
