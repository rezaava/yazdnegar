using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    [DefaultEvent("OnSelectedIndexChanged")]
    public partial class P_ComboBox : UserControl
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

        //Fields
        private Color backColor = Color.White;
        private Color iconColor = Color.DimGray;

        private Color listBackColor = Color.White;
        private Color listTextColor = SystemColors.WindowText;

        private Color borderColor = SystemColors.ActiveBorder;
        private int borderSize = 1;
        private int borderRadius = 0;

        private Size buttonArrowSize = new Size(10, 4);
        private byte buttonArrowThickness = 2;

        ComboBoxStyle dropDownStyle;


        private PaletteDrawBorders paletteDrawBorders = PaletteDrawBorders.All;

        //Events
        public event EventHandler OnSelectedIndexChanged;//Default event

        //Constructor
        public P_ComboBox()
        {
            InitializeComponent();

            this.SuspendLayout();

            mainPanel.Dock = DockStyle.Fill;

            //ComboBox: Dropdown list
            cmbList.BackColor = backColor;
            cmbList.ForeColor = listTextColor;
            cmbList.SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);//Default event
            cmbList.TextChanged += new EventHandler(ComboBox_TextChanged);//Refresh text
            dropDownStyle = ComboBoxStyle.DropDownList;

            //Button: Icon
            btnIcon.BackColor = backColor;
            btnIcon.Click += new EventHandler(Icon_Click);//Open dropdown list
            btnIcon.Paint += new PaintEventHandler(Icon_Paint);//Draw icon

            //Label: Text
            lblText.BackColor = backColor;
            //->Attach label events to user control event
            lblText.Click += new EventHandler(Surface_Click);//Select combo box

            //User Control
            this.Padding = new Padding(borderSize);//Border Size
            BackColor = borderColor; //Border Color
            this.ResumeLayout();

            AdjustComboBoxDimensions();
        }

        //Properties

        //-> Appearance
        [Category("Dedicated")]
        public int BorderRadius
        {
            get { return borderRadius; }
            set
            {
                if (value >= 0)
                    if (value >= 0)
                    {
                        borderRadius = value;
                        this.Invalidate();//Redraw control
                    }
            }
        }
        [Category("Dedicated")]
        public PaletteDrawBorders PaletteDrawBorder
        {
            get => paletteDrawBorders; set => paletteDrawBorders = value;
        }
        [Category("Dedicated")]
        public new Color BackColor
        {
            get { return backColor; }
            set
            {
                backColor = value;
                lblText.BackColor = backColor;
                //txtText.BackColor = backColor;//transparent not supported
                btnIcon.BackColor = backColor;
            }
        }

        [Category("Dedicated")]
        public Color IconColor
        {
            get { return iconColor; }
            set
            {
                iconColor = value;
                btnIcon.Invalidate();//Redraw icon
            }
        }

        [Category("Dedicated")]
        public Color ListBackColor
        {
            get { return listBackColor; }
            set
            {
                listBackColor = value;
                cmbList.BackColor = listBackColor;
            }
        }

        [Category("Dedicated")]
        public Color ListTextColor
        {
            get { return listTextColor; }
            set
            {
                listTextColor = value;
                cmbList.ForeColor = listTextColor;
            }
        }

        [Category("Dedicated")]
        public Color BorderColor
        {
            get { return borderColor; }
            set
            {
                borderColor = value;
                this.BackColor = borderColor;
                //base.BackColor = borderColor; //Border Color
            }
        }

        [Category("Dedicated")]
        public int BorderSize
        {
            get { return borderSize; }
            set
            {
                borderSize = value;
                this.Padding = new Padding(borderSize);//Border Size
                AdjustComboBoxDimensions();
            }
        }

        [Category("Dedicated")]
        public override Color ForeColor
        {
            get { return base.ForeColor; }
            set
            {
                base.ForeColor = value;
                lblText.ForeColor = value;
                txtText.ForeColor = value;
            }
        }

        [Category("Dedicated")]
        public override Font Font
        {
            get { return base.Font; }
            set
            {
                base.Font = value;
                lblText.Font = value;
                txtText.Font = value;
                cmbList.Font = value;//Optional
            }
        }
        [Category("Dedicated")]
        public override string Text
        {
            get
            {
                if (dropDownStyle == ComboBoxStyle.DropDownList)
                {
                    return lblText.Text;
                }
                else
                {
                    return txtText.Text;
                }
            }
            set
            {
                if (dropDownStyle == ComboBoxStyle.DropDownList)
                {
                    lblText.Text = value;
                }
                else
                {
                    txtText.Text = value;
                }
            }
        }

        [Category("Dedicated")]
        public Color PlaceholderColor
        {
            get { return txtText.PlaceholderColor; }
            set
            {
                txtText.PlaceholderColor = value;
            }
        }

        [Category("Dedicated")]
        public string PlaceholderText
        {
            get { return txtText.PlaceholderText; }
            set
            {
                txtText.PlaceholderText = value;
            }
        }

        [Category("Dedicated")]
        public ComboBoxStyle DropDownStyle
        {
            get { return dropDownStyle; }
            set
            {
                //cmbList.DropDownStyle = value;
                dropDownStyle = value;

                SuspendLayout();
                if (value == ComboBoxStyle.DropDown)
                {
                    lblText.Hide();
                    txtText.Show();
                    txtText.BringToFront();
                }
                else if (value == ComboBoxStyle.DropDownList)
                {
                    txtText.Hide();
                    lblText.Show();
                    lblText.BringToFront();
                }
                ResumeLayout();
            }
        }
        [Category("Dedicated")]
        public ContentAlignment TextAlign
        {
            get { return lblText.TextAlign; }
            set
            {
                if (cmbList.DropDownStyle != ComboBoxStyle.DropDown)
                    lblText.TextAlign = value;
            }
        }

        //Properties
        //-> Data
        [Category("Special Data")]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Localizable(true)]
        [MergableProperty(false)]
        public System.Windows.Forms.ComboBox.ObjectCollection Items
        {
            get { return cmbList.Items; }
        }

        [Category("Special Data")]
        [AttributeProvider(typeof(IListSource))]
        [DefaultValue(null)]
        public object DataSource
        {
            get { return cmbList.DataSource; }
            set { cmbList.DataSource = value; }
        }

        [Category("Special Data")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [EditorBrowsable(EditorBrowsableState.Always)]
        [Localizable(true)]
        public AutoCompleteStringCollection AutoCompleteCustomSource
        {
            get { return cmbList.AutoCompleteCustomSource; }
            set { cmbList.AutoCompleteCustomSource = value; }
        }

        [Category("Special Data")]
        [Browsable(true)]
        [DefaultValue(AutoCompleteSource.None)]
        [EditorBrowsable(EditorBrowsableState.Always)]
        public AutoCompleteSource AutoCompleteSource
        {
            get { return cmbList.AutoCompleteSource; }
            set { cmbList.AutoCompleteSource = value; }
        }

        [Category("Special Data")]
        [Browsable(true)]
        [DefaultValue(AutoCompleteMode.None)]
        [EditorBrowsable(EditorBrowsableState.Always)]
        public AutoCompleteMode AutoCompleteMode
        {
            get { return cmbList.AutoCompleteMode; }
            set { cmbList.AutoCompleteMode = value; }
        }

        [Category("Special Data")]
        [Bindable(true)]
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public object SelectedItem
        {
            get { return cmbList.SelectedItem; }
            set { cmbList.SelectedItem = value; }
        }

        [Category("Special Data")]
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int SelectedIndex
        {
            get { return cmbList.SelectedIndex; }
            set { cmbList.SelectedIndex = value; }
        }

        [Category("Special Data")]
        [DefaultValue("")]
        [Editor("System.Windows.Forms.Design.DataMemberFieldEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [TypeConverter("System.Windows.Forms.Design.DataMemberFieldConverter, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
        public string DisplayMember
        {
            get { return cmbList.DisplayMember; }
            set { cmbList.DisplayMember = value; }
        }

        [Category("Special Data")]
        [DefaultValue("")]
        [Editor("System.Windows.Forms.Design.DataMemberFieldEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public string ValueMember
        {
            get { return cmbList.ValueMember; }
            set { cmbList.ValueMember = value; }
        }

        [Category("Dedicated")]
        public Size ButtonArrowSize
        {
            get => buttonArrowSize;
            set
            {
                buttonArrowSize = value;
                this.Invalidate();
            }
        }

        [Category("Dedicated")]
        public byte ButtonArrowThickness { get => buttonArrowThickness; set => buttonArrowThickness = value; }

        //Private methods
        private void AdjustComboBoxDimensions()
        {
            //cmbList.Width = lblText.Width;
            cmbList.Width = ((int)lblText.Width) + btnIcon.Width;

            cmbList.Location = new Point()
            {
                X = this.Width - this.Padding.Right - cmbList.Width,
                Y = ((int)lblText.Bottom) - cmbList.Height
            };
        }

        //Event methods

        //-> Default event

        private void P_ComboBox_Leave(object sender, EventArgs e)
        {
        }
        private void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (OnSelectedIndexChanged != null)
                OnSelectedIndexChanged.Invoke(sender, e);
            //Refresh text
            lblText.Text = cmbList.Text;
        }

        //-> Draw icon
        private void Icon_Paint(object sender, PaintEventArgs e)
        {
            //Fields
            int iconWidth = buttonArrowSize.Width;//14,6
            int iconHeight = buttonArrowSize.Height;
            var rectIcon = new Rectangle((btnIcon.Width - iconWidth) / 2, (btnIcon.Height - iconHeight) / 2, iconWidth, iconHeight);
            Graphics graph = e.Graphics;

            //Draw arrow down icon
            using (GraphicsPath path = new GraphicsPath())
            using (Pen pen = new Pen(iconColor, ButtonArrowThickness))
            {
                graph.SmoothingMode = SmoothingMode.AntiAlias;
                path.AddLine(rectIcon.X, rectIcon.Y, rectIcon.X + (iconWidth / 2), rectIcon.Bottom);
                path.AddLine(rectIcon.X + (iconWidth / 2), rectIcon.Bottom, rectIcon.Right, rectIcon.Y);
                graph.DrawPath(pen, path);
            }
        }

        //-> Items actions
        private void Icon_Click(object sender, EventArgs e)
        {
            //Open dropdown list
            cmbList.Select();
            cmbList.DroppedDown = true;
        }
        private void Surface_Click(object sender, EventArgs e)
        {
            //Attach label click to user control click
            this.OnClick(e);
            //Select combo box
            cmbList.Select();
            if (cmbList.DropDownStyle == ComboBoxStyle.DropDownList)
                cmbList.DroppedDown = true;//Open dropdown list
        }
        private void ComboBox_TextChanged(object sender, EventArgs e)
        {
            //Refresh text
            lblText.Text = cmbList.Text;
        }

        //Overridden methods
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            AdjustComboBoxDimensions();
            this.Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            SuspendLayout();

            base.OnPaint(e);
            Graphics graph = e.Graphics;

            if (borderRadius > 1)//Rounded TextBox
            {
                //-Fields
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
                    if (borderRadius > 15) SetComboBoxRoundedRegion();//Set the rounded region of TextBox component
                    graph.SmoothingMode = SmoothingMode.AntiAlias;
                    penBorder.Alignment = System.Drawing.Drawing2D.PenAlignment.Center;

                    //Draw border smoothing
                    graph.DrawPath(penBorderSmooth, pathBorderSmooth);
                    //Draw border
                    graph.DrawPath(penBorder, pathBorder);
                }
            }
            else //Square/Normal TextBox
            {
                //Draw border
                using (Pen penBorder = new Pen(borderColor, borderSize))
                {
                    this.Region = new Region(this.ClientRectangle);
                    penBorder.Alignment = PenAlignment.Inset;

                    graph.DrawRectangle(penBorder, 0, 0, this.Width - 0.5F, this.Height - 0.5F);
                }
            }
            ResumeLayout();
        }

        private void SetComboBoxRoundedRegion()
        {
            GraphicsPath pathTxt;
            //pathTxt = GetFigurePath(mainPanel.ClientRectangle, borderRadius - borderSize);
            //mainPanel.Region = new Region(pathTxt);

            pathTxt = GetFigurePath(mainPanel.ClientRectangle, borderSize * 2);
            mainPanel.Region = new Region(pathTxt);

            pathTxt.Dispose();
        }

        private GraphicsPath GetFigurePath(Rectangle rect, int radius)
        {
            GraphicsPath borderPath = new GraphicsPath();
            float curveSize = radius * 2F;

            borderPath.StartFigure();
            // Reduce the width and height by 1 pixel for drawing into rectangle
            rect.Width -= 1;
            rect.Height -= 1;

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

        public void FocusControl()
        {
            if (dropDownStyle == ComboBoxStyle.DropDownList)
            {
                btnIcon.Focus();
                btnIcon.Select();
            }
            else
            {
                txtText.Focus();
                txtText.Select();
            }
        }
    }
}