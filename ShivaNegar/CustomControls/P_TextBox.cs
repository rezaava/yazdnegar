using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    [DefaultEvent("_TextChanged")]
    public partial class P_TextBox : UserControl
    {

        public enum PaletteDrawBorders
        {
            TopLeft,
            TopRight,
            BottomLeft,
            BottomRight,
            TopBottomLeft,
            TopBottomRight,
            TopLeftRight,
            BottomLeftRight,
            All,
        }

        #region -> Fields
        //Fields
        private Color borderColor = SystemColors.ActiveBorder;
        private Color borderFocusColor = Color.Black;
        private int borderRadius = 0;
        private int borderSize = 2;

        private bool underlinedStyle = false;
        private bool isFocused = false;

        private Color placeholderColor = Color.DarkGray;
        private string placeholderText = "";
        private bool isPlaceholder = false;
        private bool isPasswordChar = false;

        private PaletteDrawBorders paletteDrawBorders = PaletteDrawBorders.All;

        //Events
        public event EventHandler _TextChanged;

        #endregion

        //-> Constructor
        public P_TextBox()
        {
            //Created by designer
            InitializeComponent();
        }

        #region -> Properties
        [Category("Dedicated")]
        public Color BorderColor
        {
            get { return borderColor; }
            set
            {
                borderColor = value;
                this.Invalidate();
            }
        }

        [Category("Dedicated")]
        public Color BorderFocusColor
        {
            get { return borderFocusColor; }
            set { borderFocusColor = value; }
        }

        [Category("Dedicated")]
        public int BorderSize
        {
            get { return borderSize; }
            set
            {
                if (value >= 0)
                {
                    borderSize = value;
                    this.Invalidate();
                }
            }
        }

        [Category("Dedicated")]
        public bool UnderlinedStyle
        {
            get { return underlinedStyle; }
            set
            {
                underlinedStyle = value;
                this.Invalidate();
            }
        }

        [Category("Dedicated")]
        public bool PasswordChar
        {
            get { return isPasswordChar; }
            set
            {
                isPasswordChar = value;
                if (!isPlaceholder)
                    textBox1.UseSystemPasswordChar = value;
            }
        }

        [Category("Dedicated")]
        public bool Multiline
        {
            get { return textBox1.Multiline; }
            set { textBox1.Multiline = value; }
        }

        [Category("Dedicated")]
        public override Color BackColor
        {
            get { return base.BackColor; }
            set
            {
                base.BackColor = value;
                textBox1.BackColor = value;
            }
        }

        [Category("Dedicated")]
        public override Color ForeColor
        {
            get { return base.ForeColor; }
            set
            {
                base.ForeColor = value;
                textBox1.ForeColor = value;
            }
        }

        [Category("Dedicated")]
        public override Font Font
        {
            get { return base.Font; }
            set
            {
                base.Font = value;
                textBox1.Font = value;
                if (this.DesignMode)
                    UpdateControlHeight();
            }
        }

        [Category("Dedicated")]
        public override string Text
        {
            get
            {
                if (isPlaceholder) return "";
                else return textBox1.Text;
            }
            set
            {
                textBox1.Text = value;
                if (string.IsNullOrEmpty(value))
                    SetPlaceholder();
                else
                    RemovePlaceholder();
            }
        }

        [Category("Dedicated")]
        public int BorderRadius
        {
            get { return borderRadius; }
            set
            {
                if (value >= 0)
                {
                    borderRadius = value;
                    this.Invalidate();//Redraw control
                }
            }
        }

        [Category("Dedicated")]
        public Color PlaceholderColor
        {
            get { return placeholderColor; }
            set
            {
                placeholderColor = value;
                if (isPlaceholder)
                    textBox1.ForeColor = value;
            }
        }

        [Category("Dedicated")]
        public string PlaceholderText
        {
            get { return placeholderText; }
            set
            {
                placeholderText = value;
                if (string.IsNullOrEmpty(textBox1.Text))
                    SetPlaceholder();
            }
        }

        [Category("Dedicated")]
        public PaletteDrawBorders PaletteDrawBorder
        {
            get => paletteDrawBorders;
            set => paletteDrawBorders = value;
        }
        #endregion

        #region -> Overridden methods
        protected override void OnResize(EventArgs e)
        {
            if (this.DesignMode)
                UpdateControlHeight();
            this.Refresh();
            base.OnResize(e);
        }

        protected override void OnLoad(EventArgs e)
        {
            UpdateControlHeight();
            base.OnLoad(e);
        }
        protected override void OnPaint(PaintEventArgs e)
        {

            base.OnPaint(e);
            Graphics graph = e.Graphics;

            if (borderRadius > 1)//Rounded TextBox
            {
                //-Fields
                SuspendLayout();

                var rectBorderSmooth = this.ClientRectangle;
                var rectBorder = Rectangle.Inflate(rectBorderSmooth, -borderSize, -borderSize);
                int smoothSize = borderSize > 0 ? borderSize : 1;

                using (GraphicsPath pathBorderSmooth = GetFigurePath(rectBorderSmooth, borderRadius))
                using (GraphicsPath pathBorder = GetFigurePath(rectBorder, borderRadius - borderSize))
                using (Pen penBorderSmooth = new Pen(this.Parent.BackColor, smoothSize))
                using (Pen penBorder = new Pen(borderColor, borderSize))
                {
                    //-Drawing
                    this.Region = new Region(pathBorderSmooth);//Set the rounded region of UserControl
                    if (borderRadius > 15) SetTextBoxRoundedRegion();//Set the rounded region of TextBox component

                    if (isFocused) penBorder.Color = borderFocusColor;

                    graph.SmoothingMode = SmoothingMode.AntiAlias;
                    penBorder.Alignment = System.Drawing.Drawing2D.PenAlignment.Center;

                    if (underlinedStyle) //Line Style
                    {
                        //Draw border smoothing
                        graph.DrawPath(penBorderSmooth, pathBorderSmooth);
                        //Draw border
                        graph.SmoothingMode = SmoothingMode.None;
                        graph.DrawLine(penBorder, 0, this.Height - 1, this.Width, this.Height - 1);
                    }
                    else //Normal Style
                    {
                        //Draw border smoothing
                        graph.DrawPath(penBorderSmooth, pathBorderSmooth);
                        //Draw border
                        graph.DrawPath(penBorder, pathBorder);
                    }
                }
                ResumeLayout();
            }
            else //Square/Normal TextBox
            {
                if (borderSize != 0)
                {
                    //Draw border
                    using (Pen penBorder = new Pen(borderColor, borderSize))
                    {
                        this.Region = new Region(this.ClientRectangle);
                        penBorder.Alignment = PenAlignment.Inset;
                        if (isFocused) penBorder.Color = borderFocusColor;

                        if (underlinedStyle) //Line Style
                            graph.DrawLine(penBorder, 0, this.Height - 1, this.Width, this.Height - 1);
                        else //Normal Style
                            graph.DrawRectangle(penBorder, 0, 0, this.Width - 0.5F, this.Height - 0.5F);
                    }
                }
            }

        }
        #endregion

        #region -> Private methods
        private void SetPlaceholder()
        {
            if (string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(placeholderText))
            {
                isPlaceholder = true;
                textBox1.Text = placeholderText;
                textBox1.ForeColor = placeholderColor;
                if (isPasswordChar)
                    textBox1.UseSystemPasswordChar = false;
            }
        }
        private void RemovePlaceholder()
        {
            if (isPlaceholder && !string.IsNullOrEmpty(placeholderText) && textBox1.Text.Contains(placeholderText))
            {
                isPlaceholder = false;
                textBox1.Text = "";
                textBox1.ForeColor = this.ForeColor;
                if (isPasswordChar)
                    textBox1.UseSystemPasswordChar = true;
            }
            else if (!textBox1.Text.Contains(placeholderText))
            {
                isPlaceholder = false;
                textBox1.ForeColor = this.ForeColor;
                if (isPasswordChar)
                    textBox1.UseSystemPasswordChar = true;
            }
        }
        private GraphicsPath GetFigurePath(Rectangle rect, int radius)
        {
            GraphicsPath borderPath = new GraphicsPath();
            float curveSize = radius * 2F;

            borderPath.StartFigure();
            // Reduce the width and height by 1 pixel for drawing into rectangle
            //rect.Width -= 1;
            //rect.Height -= 1;

            // We create the path using a floating point rectangle
            RectangleF rectF = new RectangleF(rect.X, rect.Y, rect.Width, rect.Height);

            /*            if (!middle)
                        {
                            rectF.X -= 0.25f;
                            rectF.Y -= 0.25f;
                            rectF.Width += 0.75f;
                            rectF.Height += 0.75f;
                        }*/

            // Add only the border for drawing
            switch (paletteDrawBorders)
            {
                case PaletteDrawBorders.TopLeft:
                    borderPath.AddLine(rectF.Left, rectF.Bottom + 1, rectF.Left, rectF.Top + curveSize);
                    borderPath.AddArc(rectF.Left, rectF.Top, curveSize, curveSize, 180f, 90f);
                    borderPath.AddLine(rectF.Left + curveSize, rectF.Top, rectF.Right + 1, rectF.Top);
                    borderPath.AddLine(rectF.Right, rectF.Top, rectF.Right, rectF.Bottom);
                    break;
                case PaletteDrawBorders.TopRight:
                    borderPath.AddLine(rectF.Left - 1, rectF.Top, rectF.Right - curveSize, rectF.Top);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Top, curveSize, curveSize, -90f, 90f);
                    borderPath.AddLine(rectF.Right, rectF.Top + curveSize, rectF.Right, rectF.Bottom + 1);
                    borderPath.AddLine(rectF.Right, rectF.Bottom, rectF.Left, rectF.Bottom);
                    break;
                case PaletteDrawBorders.BottomRight:
                    borderPath.AddLine(rectF.Right, rectF.Top - 1, rectF.Right, rectF.Bottom - curveSize);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Bottom - curveSize, curveSize, curveSize, 0f, 90f);
                    borderPath.AddLine(rectF.Right - curveSize, rectF.Bottom, rectF.Left - 1, rectF.Bottom);
                    borderPath.AddLine(rectF.Left, rectF.Bottom, rectF.Left, rectF.Top);
                    break;
                case PaletteDrawBorders.BottomLeft:
                    borderPath.AddLine(rectF.Right + 1, rectF.Bottom, rectF.Left + curveSize, rectF.Bottom);
                    borderPath.AddArc(rectF.Left, rectF.Bottom - curveSize, curveSize, curveSize, 90f, 90f);
                    borderPath.AddLine(rectF.Left, rectF.Bottom - curveSize, rectF.Left, rectF.Top - 1);
                    borderPath.AddLine(rectF.Left, rectF.Top, rectF.Right, rectF.Top);
                    break;
                case PaletteDrawBorders.TopBottomLeft:
                    borderPath.AddLine(rectF.Right + 1, rectF.Bottom, rectF.Left + curveSize, rectF.Bottom);
                    borderPath.AddArc(rectF.Left, rectF.Bottom - curveSize, curveSize, curveSize, 90f, 90f);
                    borderPath.AddArc(rectF.Left, rectF.Top, curveSize, curveSize, 180f, 90f);
                    borderPath.AddLine(rectF.Left + curveSize, rectF.Top, rectF.Right + 1, rectF.Top);
                    break;
                case PaletteDrawBorders.TopBottomRight:
                    borderPath.AddLine(rectF.Left - 1, rectF.Top, rectF.Right - curveSize, rectF.Top);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Top, curveSize, curveSize, -90f, 90f);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Bottom - curveSize, curveSize, curveSize, 0f, 90f);
                    borderPath.AddLine(rectF.Right - curveSize, rectF.Bottom, rectF.Left - 1, rectF.Bottom);
                    break;
                case PaletteDrawBorders.TopLeftRight:
                    borderPath.AddLine(rectF.Left, rectF.Bottom + 1, rectF.Left, rectF.Top + curveSize);
                    borderPath.AddArc(rectF.Left, rectF.Top, curveSize, curveSize, 180f, 90f);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Top, curveSize, curveSize, -90f, 90f);
                    borderPath.AddLine(rectF.Right, rectF.Top + curveSize, rectF.Right, rectF.Bottom + 1);
                    break;
                case PaletteDrawBorders.BottomLeftRight:
                    borderPath.AddLine(rectF.Right, rectF.Top - 1, rectF.Right, rectF.Bottom - curveSize);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Bottom - curveSize, curveSize, curveSize, 0f, 90f);
                    borderPath.AddArc(rectF.Left, rectF.Bottom - curveSize, curveSize, curveSize, 90f, 90f);
                    borderPath.AddLine(rectF.Left, rectF.Bottom - curveSize, rectF.Left, rectF.Top - 1);
                    break;
                case PaletteDrawBorders.All:
                    borderPath.AddArc(rectF.Left, rectF.Top, curveSize, curveSize, 180f, 90f);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Top, curveSize, curveSize, 270f, 90f);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Bottom - curveSize, curveSize, curveSize, 0f, 90f);
                    borderPath.AddArc(rectF.Left, rectF.Bottom - curveSize, curveSize, curveSize, 90f, 90f);

                    //borderPath.AddArc(rect.X, rect.Y, arcLength, arcLength, 180, 90);
                    //borderPath.AddArc(rect.Right - arcLength, rect.Y, arcLength, arcLength, 270, 90);
                    //borderPath.AddArc(rect.Right - arcLength, rect.Bottom - arcLength, arcLength, arcLength, 0, 90);
                    //borderPath.AddArc(rect.X, rect.Bottom - arcLength, arcLength, arcLength, 90, 90);
                    break;
            }
            borderPath.CloseFigure();

            return borderPath;
        }

        private void SetTextBoxRoundedRegion()
        {
            GraphicsPath pathTxt;
            if (Multiline)
            {
                pathTxt = GetFigurePath(textBox1.ClientRectangle, borderRadius - borderSize);
                textBox1.Region = new Region(pathTxt);
            }
            else
            {
                pathTxt = GetFigurePath(textBox1.ClientRectangle, borderSize * 2);
                textBox1.Region = new Region(pathTxt);
            }
            pathTxt.Dispose();
        }
        private void UpdateControlHeight()
        {
            if (textBox1.Multiline == false)
            {
                int txtHeight = TextRenderer.MeasureText("Text", this.Font).Height + 1;
                textBox1.Multiline = true;
                textBox1.MinimumSize = new Size(0, txtHeight);
                textBox1.Multiline = false;

                this.Height = textBox1.Height + this.Padding.Top + this.Padding.Bottom;
            }
        }
        #endregion

        #region -> TextBox events
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (_TextChanged != null)
                _TextChanged.Invoke(sender, e);
        }
        private void textBox1_Click(object sender, EventArgs e)
        {
            this.OnClick(e);
        }
        private void textBox1_MouseEnter(object sender, EventArgs e)
        {
            this.OnMouseEnter(e);
        }
        private void textBox1_MouseLeave(object sender, EventArgs e)
        {
            this.OnMouseLeave(e);
        }
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            this.OnKeyPress(e);
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {
            isFocused = true;
            this.Invalidate();
            RemovePlaceholder();
        }
        private void textBox1_Leave(object sender, EventArgs e)
        {
            isFocused = false;
            this.Invalidate();
            SetPlaceholder();
        }
        ///::::+
        #endregion
    }
}
