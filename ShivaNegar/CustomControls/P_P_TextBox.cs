using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    public partial class P_P_TextBox : UserControl
    {
        public P_P_TextBox()
        {
            InitializeComponent();
        }
        [Category("Dedicated")]
        [EditorAttribute(
    "System.ComponentModel.Design.MultilineStringEditor, System.Design",
    "System.Drawing.Design.UITypeEditor")]
        public string Caption { get => lblCaption.Text; set => lblCaption.Text = value; }

        [EditorBrowsable(EditorBrowsableState.Always)]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Visible)]
        [Bindable(true)]
        [EditorAttribute(
    "System.ComponentModel.Design.MultilineStringEditor, System.Design",
    "System.Drawing.Design.UITypeEditor")]
        [Category("Dedicated")]
        public override string Text { get => textBox.Text; set => textBox.Text = value; }
        [Category("Dedicated")]
        [EditorAttribute(
    "System.ComponentModel.Design.MultilineStringEditor, System.Design",
    "System.Drawing.Design.UITypeEditor")]
        public string Hint { get => textBox.PlaceholderText; set => textBox.PlaceholderText = value; }

        [Category("Dedicated")]
        public Size panelCaptionSize { get => panelCaption.Size; set => panelCaption.Size = value; }

        [Category("Dedicated")]
        public Font captionFont { get => lblCaption.Font; set => lblCaption.Font = value; }

        [Category("Dedicated")]
        public Font txtBoxFont { get => textBox.Font; set => textBox.Font = value; }

        [Category("Dedicated")]
        public bool MultiLine { get => textBox.Multiline; set => textBox.Multiline = value; }

        public override RightToLeft RightToLeft
        {
            get => base.RightToLeft;
            set
            {
                panelCaption.RightToLeft = value;
                lblCaption.RightToLeft = value;
                textBox.RightToLeft = value;
                base.RightToLeft = value;


                if (value == RightToLeft.Yes)
                {
                    panelCaption.PaletteDrawBorder = PaletteDrawBorders.TopBottomRight;
                    textBox.PaletteDrawBorder = P_TextBox.PaletteDrawBorders.TopBottomLeft;
                    panelCaption.Dock = DockStyle.Right;
                }
                else if (value == RightToLeft.No)
                {
                    panelCaption.PaletteDrawBorder = PaletteDrawBorders.TopBottomLeft;
                    textBox.PaletteDrawBorder = P_TextBox.PaletteDrawBorders.TopBottomRight;
                    panelCaption.Dock = DockStyle.Left;
                }
                else if (value == RightToLeft.Inherit)
                {
                    if (this.RightToLeft == RightToLeft.Yes)
                    {
                        panelCaption.PaletteDrawBorder = PaletteDrawBorders.TopBottomRight;
                        textBox.PaletteDrawBorder = P_TextBox.PaletteDrawBorders.TopBottomLeft;
                        panelCaption.Dock = DockStyle.Right;

                    }
                    else if (this.RightToLeft == RightToLeft.No)
                    {
                        panelCaption.PaletteDrawBorder = PaletteDrawBorders.TopBottomLeft;
                        textBox.PaletteDrawBorder = P_TextBox.PaletteDrawBorders.TopBottomRight;
                        panelCaption.Dock = DockStyle.Left;

                    }
                }


            }

        }


        public delegate void LinkTextChangedHandler(object sender, EventArgs e);
        // The event
        public event LinkTextChangedHandler onTextChanged;

        private void textBox__TextChanged(object sender, EventArgs e)
        {
            // Check if there are any Subscribers
            if (onTextChanged != null)
            {
                // Call the Event
                onTextChanged(this, e);
            }
        }

        public delegate void LinkFocusEnterTextBoxHandler(object sender, EventArgs e);
        // The event
        public event LinkFocusEnterTextBoxHandler onFocusEnter;

        private void textBox_Enter(object sender, EventArgs e)
        {
            // Check if there are any Subscribers
            if (onFocusEnter != null)
            {
                // Call the Event
                onFocusEnter(this, e);
            }
        }
    }

}
