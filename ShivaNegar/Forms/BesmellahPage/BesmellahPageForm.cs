using System;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ShivaNegar.Forms.BesmellahPage
{
    public partial class BesmellahPageForm : Form
    {
        [Flags]
        enum AnimateWindowFlags
        {
            AW_HOR_POSITIVE = 0x0000000,
            AW_HOR_NEGATIVE = 0x00000002,
            AW_VER_POSITIVE = 0x00000004,
            AW_VER_NEGATIVE = 0x00000008,
            AW_CENTER = 0x00000010,
            AW_HIDE = 0x00010000,
            AW_ACTIVATE = 0x00020000,
            AW_SLIDE = 0x00040000,
            AW_BLEND = 0x00080000
        }

        [DllImport("user32.dll")]
        static extern bool AnimateWindow(IntPtr hWnd, int time, AnimateWindowFlags flags);


        #region for Windows form and Controls
        bool resizableForm = true;

        private int formBorderRadius = 10;
        private int formBorderSize = 5;
        private Color formBorderColor = Color.FromArgb(6, 174, 244);
        //PrivateFontCollection pfcRegular;
        //PrivateFontCollection pfcBold;
        #endregion

        public BesmellahPageForm()
        {
            InitializeComponent();

            DedicatedFunctions.setManualScale(this, new Size(700, 500), new Size(600, 400));
            this.CenterToScreen();

            //#region set Custom Resolution in Multiple DPI

            //float dpi;
            //Graphics graphics = this.CreateGraphics();
            //dpi = graphics.DpiX;

            //int screenWitdh = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Size.Width;
            //int screenHeight = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Size.Height;

            //this.MinimumSize = new System.Drawing.Size(700 , 500);
            //this.Width = 800;
            //this.Height = 600;
            //if(dpi >= 140)
            //{
            //	if(screenWitdh >= 1200)
            //		this.Width = 1200;
            //	else if(screenWitdh >= 1100)
            //		this.Width = 1100;

            //	if(screenHeight >= 1000)
            //		this.Height = 1000;
            //	else if(screenHeight >= 900)
            //		this.Height = 900;

            //	this.MinimumSize = new System.Drawing.Size(1000 , 800);
            //}
            //else if(dpi >= 110)
            //{
            //	if(screenWitdh >= 950)
            //		this.Width = 950;

            //	if(screenHeight >= 750)
            //		this.Height = 750;

            //	this.MinimumSize = new System.Drawing.Size(900 , 700);
            //}
            //#endregion

            this.FormClosing += (e, a) =>
            {
                AnimateWindow(this.Handle, 100, AnimateWindowFlags.AW_BLEND | AnimateWindowFlags.AW_HIDE);
                Globals.ThisAddIn.DocumentManagerFormVisible = false;
            };

            #region Form inital settings
            this.Padding = new Padding(formBorderSize);
            this.BackColor = formBorderColor;
            this.Opacity = 0;
            setDoubleBuffer(elementHost1, true);
            this.SetStyle(ControlStyles.ResizeRedraw, true);
            this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(1, 1, Width, Height, formBorderRadius * 2, formBorderRadius * 2));
            timerOpenAnimation.Enabled = true;
            #endregion

            besmellahPageControl.CloseFormRequest += () =>//CloseFormButton Event
            {
                closeForm();
            };
            besmellahPageControl.MaximizeStateFormRequest += () =>//CloseFormButton Event
            {
                if (this.WindowState == FormWindowState.Normal)
                    this.WindowState = FormWindowState.Maximized;
                else if (this.WindowState == FormWindowState.Maximized)
                    this.WindowState = FormWindowState.Normal;
            };
            besmellahPageControl.MinimizeStateFormRequest += () =>//CloseFormButton Event
            {
                this.WindowState = FormWindowState.Minimized;
            };

            besmellahPageControl.MouseDown += ChangeContentsControl_MouseDown;
        }


        private void closeForm()
        {
            GC.Collect();
            //this.Dispose();
            this.Close();
            Globals.ThisAddIn.ChangeContentFormVisible = false;
        }
        private void ChangeContentsControl_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            releaseCapture();
            sendMessage(this.Handle, 0x112, 0xf012, 0);
        }


        #region (UI) WindowsForm

        //#region Configure custom font from Resources(in Properties)
        //public void initCustomLabelFontRegular()
        //{
        //	//Create your private font collection object.
        //	pfcRegular = new PrivateFontCollection();

        //	//Select your font from the resources.
        //	//My font here is "Digireu.ttf"
        //	int fontLength = Properties.Resources.Vazirmatn_Regular.Length;

        //	// create a buffer to read in to
        //	byte[] fontdata = Properties.Resources.Vazirmatn_Regular;

        //	// create an unsafe memory block for the font data
        //	System.IntPtr data = Marshal.AllocCoTaskMem(fontLength);

        //	// copy the bytes to the unsafe memory block
        //	Marshal.Copy(fontdata , 0 , data , fontLength);

        //	// pass the font to the font collection
        //	pfcRegular.AddMemoryFont(data , fontLength);

        //}
        //public void initCustomLabelFontBold()
        //{
        //	//Create your private font collection object.
        //	pfcBold = new PrivateFontCollection();

        //	//Select your font from the resources.
        //	//My font here is "Digireu.ttf"
        //	int fontLength = Properties.Resources.Vazirmatn_Bold.Length;

        //	// create a buffer to read in to
        //	byte[] fontdata = Properties.Resources.Vazirmatn_Bold;

        //	// create an unsafe memory block for the font data
        //	System.IntPtr data = Marshal.AllocCoTaskMem(fontLength);

        //	// copy the bytes to the unsafe memory block
        //	Marshal.Copy(fontdata , 0 , data , fontLength);

        //	// pass the font to the font collection
        //	pfcBold.AddMemoryFont(data , fontLength);
        //}
        //public void setAllControlsFont(Control.ControlCollection ctrls)
        //{
        //	foreach(Control ctrl in ctrls)
        //	{
        //		if(ctrl.Controls != null)
        //			setAllControlsFont(ctrl.Controls);

        //		if(ctrl.Name != "lblSelectedNameOfGod" && ctrl.Name != "lstNameOfAllah")
        //		{
        //			if(ctrl.Font.Style == FontStyle.Regular)
        //				ctrl.Font = new System.Drawing.Font(pfcRegular.Families[0] , ctrl.Font.Size , FontStyle.Regular);
        //			else if(ctrl.Font.Style == FontStyle.Bold)
        //				ctrl.Font = new System.Drawing.Font(pfcBold.Families[0] , ctrl.Font.Size , FontStyle.Bold);
        //			else if(ctrl.Font.Style == FontStyle.Italic)
        //				ctrl.Font = new System.Drawing.Font(pfcRegular.Families[0] , ctrl.Font.Size , FontStyle.Italic);
        //			else if(ctrl.Font.Style == FontStyle.Underline)
        //				ctrl.Font = new System.Drawing.Font(pfcRegular.Families[0] , ctrl.Font.Size , FontStyle.Underline);
        //			else if(ctrl.Font.Style == FontStyle.Strikeout)
        //				ctrl.Font = new System.Drawing.Font(pfcRegular.Families[0] , ctrl.Font.Size , FontStyle.Strikeout);
        //		}
        //	}
        //}
        //#endregion

        #region Optimize
        static void setDoubleBuffer(Control ctrl, bool doubleBuffered)
        {
            try
            {
                typeof(Control).InvokeMember("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty, null, ctrl, new object[] { doubleBuffered });
            }
            catch (Exception)
            {
            }
        }
        #endregion

        #region Resizable
        int Thickness = 10;
        int Area = 8;
        private bool Above, right, Under, left, Right_above, Right_under, Left_under, Left_above;
        public string getMosuePosition(Point mouse, Form form)
        {
            bool above_underArea = mouse.X > Area && mouse.X < form.ClientRectangle.Width - Area; /* |\AngleArea(Left_Above)\(=======above_underArea========)/AngleArea(Right_Above)/| */ //Area===>(==)
            bool right_left_Area = mouse.Y > Area && mouse.Y < form.ClientRectangle.Height - Area;

            bool _Above = mouse.Y <= Thickness;  //Mouse in Above All Area
            bool _Right = mouse.X >= form.ClientRectangle.Width - Thickness;
            bool _Under = mouse.Y >= form.ClientRectangle.Height - Thickness;
            bool _Left = mouse.X <= Thickness;

            Above = _Above && (above_underArea); if (Above) return "a";   /*Mouse in Above All Area WithOut Angle Area */
            right = _Right && (right_left_Area); if (right) return "r";
            Under = _Under && (above_underArea); if (Under) return "u";
            left = _Left && (right_left_Area); if (left) return "l";


            Right_above =/*Right*/ (_Right && (!right_left_Area)) && /*Above*/ (_Above && (!above_underArea)); if (Right_above) return "ra";     /*if Mouse  Right_above */
            Right_under =/* Right*/((_Right) && (!right_left_Area)) && /*Under*/(_Under && (!above_underArea)); if (Right_under) return "ru";     //if Mouse  Right_under 
            Left_under = /*Left*/((_Left) && (!right_left_Area)) && /*Under*/ (_Under && (!above_underArea)); if (Left_under) return "lu";      //if Mouse  Left_under
            Left_above = /*Left*/((_Left) && (!right_left_Area)) && /*Above*/(_Above && (!above_underArea)); if (Left_above) return "la";      //if Mouse  Left_above

            return "";

        }
        protected override void WndProc(ref Message m)
        {
            if (resizableForm)
            {
                int x = (int)(m.LParam.ToInt64() & 0xFFFF);               //get x mouse position
                int y = (int)((m.LParam.ToInt64() & 0xFFFF0000) >> 16);   //get y mouse position  you can gave (x,y) it from "MouseEventArgs" too
                Point pt = PointToClient(new Point(x, y));

                if (m.Msg == 0x84)
                {
                    switch (getMosuePosition(pt, this))
                    {
                        case "l": m.Result = (IntPtr)10; return;  // the Mouse on Left Form
                        case "r": m.Result = (IntPtr)11; return;  // the Mouse on Right Form
                        case "a": m.Result = (IntPtr)12; return;
                        case "la": m.Result = (IntPtr)13; return;
                        case "ra": m.Result = (IntPtr)14; return;
                        case "u": m.Result = (IntPtr)15; return;
                        case "lu": m.Result = (IntPtr)16; return;
                        case "ru": m.Result = (IntPtr)17; return; // the Mouse on Right_Under Form
                        case "": m.Result = pt.Y < 32 /*mouse on caption Bar*/ ? (IntPtr)2 : (IntPtr)1; return;
                    }
                }
            }
            base.WndProc(ref m);

        }
        protected override void OnResizeEnd(EventArgs e)
        {
            base.OnResizeEnd(e);
            this.Refresh();
        }
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            this.Refresh();
        }
        #endregion

        #region Draggable 
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void releaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void sendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void StartFormV3_Dragging_MouseDown(object sender, EventArgs e)
        {
            releaseCapture();
            sendMessage(this.Handle, 0x112, 0xf012, 0);
        }


        #endregion

        #region Round Corner Windows Form

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (this.WindowState != FormWindowState.Minimized)
                this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(1, 1, Width, Height, formBorderRadius * 2, formBorderRadius * 2));
        }

        [DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn
        (
            int nLeftRect,     // x-coordinate of upper-left corner
            int nTopRect,      // y-coordinate of upper-left corner
            int nRightRect,    // x-coordinate of lower-right corner
            int nBottomRect,   // y-coordinate of lower-right corner
            int nWidthEllipse, // width of ellipse
            int nHeightEllipse // height of ellipse
        );
        #endregion

        #endregion

        private void timerOpenAnimation_Tick(object sender, EventArgs e)
        {
            Opacity += 0.1;

            if (Opacity == 1)
            {
                timerOpenAnimation.Enabled = false;
            }
        }


    }
}
