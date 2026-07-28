using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Design;
using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    public partial class P_P_ComboBox : UserControl
    {
        public P_P_ComboBox()
        {
            InitializeComponent();
        }

        [Category("Dedicated")]
        public int BorderRadius
        {
            get { return comboBox.BorderRadius; }
            set
            {
                if (value >= 0)
                    if (value >= 0)
                    {
                        comboBox.BorderRadius = value;
                        this.Invalidate();//Redraw control
                    }
            }
        }
        [Category("Dedicated")]
        public P_ComboBox.PaletteDrawBorders PaletteDrawBorder
        {
            get => comboBox.PaletteDrawBorder; set => comboBox.PaletteDrawBorder = value;
        }
        [Category("Dedicated")]
        public new Color BackColor
        {
            get { return comboBox.BackColor; }
            set
            {
                comboBox.BackColor = value;
            }
        }

        [Category("Dedicated")]
        public Color IconColor
        {
            get { return comboBox.IconColor; }
            set
            {
                comboBox.IconColor = value;

            }
        }

        [Category("Dedicated")]
        public Color ListBackColor
        {
            get { return comboBox.ListBackColor; }
            set
            {
                comboBox.ListBackColor = value;
            }
        }

        [Category("Dedicated")]
        public Color ListTextColor
        {
            get { return comboBox.ListTextColor; }
            set
            {
                comboBox.ListTextColor = value;
            }
        }

        [Category("Dedicated")]
        public Color BorderColor
        {
            get { return comboBox.BorderColor; }
            set
            {
                comboBox.BorderColor = value;

            }
        }

        [Category("Dedicated")]
        public int BorderSize
        {
            get { return comboBox.BorderSize; }
            set
            {
                comboBox.BorderSize = value;
            }
        }

        [Category("Dedicated")]
        public override Color ForeColor
        {
            get { return comboBox.ForeColor; }
            set
            {
                comboBox.ForeColor = value;
            }
        }

        [Category("Dedicated")]
        public override Font Font
        {
            get { return comboBox.Font; }
            set
            {
                comboBox.Font = value;
            }
        }
        [Category("Dedicated")]
        public override string Text
        {
            get
            {
                return comboBox.Text;
            }
            set
            {
                comboBox.Text = value;
            }
        }

        [Category("Dedicated")]
        [Description("for DropDown Mode in TextBox for Writing")]
        public Color PlaceholderColor
        {
            get { return comboBox.PlaceholderColor; }
            set
            {
                comboBox.PlaceholderColor = value;
            }
        }

        [Category("Dedicated")]
        [Description("for DropDown Mode in TextBox for Writing")]
        public string PlaceholderText
        {
            get { return comboBox.PlaceholderText; }
            set
            {
                comboBox.PlaceholderText = value;
            }
        }

        [Category("Dedicated")]
        public ComboBoxStyle DropDownStyle
        {
            get { return comboBox.DropDownStyle; }
            set
            {
                comboBox.DropDownStyle = value;
            }
        }
        [Category("Dedicated")]
        public ContentAlignment TextAlign
        {
            get { return comboBox.TextAlign; }
            set
            {
                comboBox.TextAlign = value;
            }
        }


        [Category("Dedicated")]
        [EditorAttribute(
"System.ComponentModel.Design.MultilineStringEditor, System.Design",
"System.Drawing.Design.UITypeEditor")]
        public string Caption { get => lblCaption.Text; set => lblCaption.Text = value; }

        [Category("Dedicated")]
        public Size panelCaptionSize { get => panelCaption.Size; set => panelCaption.Size = value; }

        [Category("Dedicated")]
        public Font captionFont { get => lblCaption.Font; set => lblCaption.Font = value; }

        [Category("Special Data")]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [Localizable(true)]
        [MergableProperty(false)]
        public System.Windows.Forms.ComboBox.ObjectCollection Items
        {
            get { return comboBox.Items; }
        }

        [Category("Special Data")]
        [AttributeProvider(typeof(IListSource))]
        [DefaultValue(null)]
        public object DataSource
        {
            get { return comboBox.DataSource; }
            set { comboBox.DataSource = value; }
        }

        [Category("Special Data")]
        [Browsable(true)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Content)]
        [Editor("System.Windows.Forms.Design.ListControlStringCollectionEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [EditorBrowsable(EditorBrowsableState.Always)]
        [Localizable(true)]
        public AutoCompleteStringCollection AutoCompleteCustomSource
        {
            get { return comboBox.AutoCompleteCustomSource; }
            set { comboBox.AutoCompleteCustomSource = value; }
        }

        [Category("Special Data")]
        [Browsable(true)]
        [DefaultValue(AutoCompleteSource.None)]
        [EditorBrowsable(EditorBrowsableState.Always)]
        public AutoCompleteSource AutoCompleteSource
        {
            get { return comboBox.AutoCompleteSource; }
            set { comboBox.AutoCompleteSource = value; }
        }

        [Category("Special Data")]
        [Browsable(true)]
        [DefaultValue(AutoCompleteMode.None)]
        [EditorBrowsable(EditorBrowsableState.Always)]
        public AutoCompleteMode AutoCompleteMode
        {
            get { return comboBox.AutoCompleteMode; }
            set { comboBox.AutoCompleteMode = value; }
        }

        [Category("Special Data")]
        [Bindable(true)]
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public object SelectedItem
        {
            get { return comboBox.SelectedItem; }
            set { comboBox.SelectedItem = value; }
        }

        [Category("Special Data")]
        [Browsable(false)]
        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public int SelectedIndex
        {
            get { return comboBox.SelectedIndex; }
            set { comboBox.SelectedIndex = value; }
        }

        [Category("Special Data")]
        [DefaultValue("")]
        [Editor("System.Windows.Forms.Design.DataMemberFieldEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        [TypeConverter("System.Windows.Forms.Design.DataMemberFieldConverter, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")]
        public string DisplayMember
        {
            get { return comboBox.DisplayMember; }
            set { comboBox.DisplayMember = value; }
        }

        [Category("Special Data")]
        [DefaultValue("")]
        [Editor("System.Windows.Forms.Design.DataMemberFieldEditor, System.Design, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a", typeof(UITypeEditor))]
        public string ValueMember
        {
            get { return comboBox.ValueMember; }
            set { comboBox.ValueMember = value; }
        }
        [Category("Dedicated")]
        public Size ButtonArrowSize
        {
            get => comboBox.ButtonArrowSize;
            set
            {
                comboBox.ButtonArrowSize = value;
                this.Invalidate();
            }
        }

        [Category("Dedicated")]
        public byte ButtonArrowThickness { get => comboBox.ButtonArrowThickness; set => comboBox.ButtonArrowThickness = value; }

        public override RightToLeft RightToLeft
        {
            get => base.RightToLeft;
            set
            {
                panelCaption.RightToLeft = value;
                lblCaption.RightToLeft = value;
                comboBox.RightToLeft = value;
                base.RightToLeft = value;


                if (value == RightToLeft.Yes)
                {
                    panelCaption.PaletteDrawBorder = PaletteDrawBorders.TopBottomRight;
                    p_Panel1.PaletteDrawBorder = PaletteDrawBorders.TopBottomLeft;
                    panelCaption.Dock = DockStyle.Right;
                }
                else if (value == RightToLeft.No)
                {
                    panelCaption.PaletteDrawBorder = PaletteDrawBorders.TopBottomLeft;
                    p_Panel1.PaletteDrawBorder = PaletteDrawBorders.TopBottomRight;
                    panelCaption.Dock = DockStyle.Left;
                }
                else if (value == RightToLeft.Inherit)
                {
                    if (this.RightToLeft == RightToLeft.Yes)
                    {
                        panelCaption.PaletteDrawBorder = PaletteDrawBorders.TopBottomRight;
                        p_Panel1.PaletteDrawBorder = PaletteDrawBorders.TopBottomLeft;
                        panelCaption.Dock = DockStyle.Right;

                    }
                    else if (this.RightToLeft == RightToLeft.No)
                    {
                        panelCaption.PaletteDrawBorder = PaletteDrawBorders.TopBottomLeft;
                        p_Panel1.PaletteDrawBorder = PaletteDrawBorders.TopBottomRight;
                        panelCaption.Dock = DockStyle.Left;

                    }
                }
            }

        }

        public delegate void LinkFocusEnterTextBoxHandler(object sender, EventArgs e);


        public delegate void LinkOnSelectedIndexChangedHandler(object sender, EventArgs e);
        // The event
        public event LinkOnSelectedIndexChangedHandler OnSelectedIndexChanged;
        private void comboSelectUniversity_OnSelectedIndexChanged(object sender, EventArgs e)
        {
            if (OnSelectedIndexChanged != null)
                OnSelectedIndexChanged(sender, e);
        }

        public void FocusControl()
        {
            comboBox.FocusControl();
        }
    }
}
