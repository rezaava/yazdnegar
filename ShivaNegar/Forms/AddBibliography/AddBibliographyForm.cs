using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.CustomControls;
using ShivaNegar.Forms.AddBibliography.Models;
using ShivaNegar.Properties;

namespace ShivaNegar.Forms.AddBibliography
{

    public partial class AddBibliographyForm : Form
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

        #region patternsRegex
        string patternbibTex = @"@(?<CContentType>\w+)\s*\{\s*(?<citationKey>\w+)\s*,(?<keyValue>\s*(?<key>\w+)\s*=\s*(\{(?<value>.*?)\}|\""(?<value>.*?)\"")\s*\,?)+(\,|)+\s*}";//bibtex
                                                                                                                                                                                 //string patternRis = @"(?<key>.*)\s\s-\s(?<value>.*)";//ris
                                                                                                                                                                                 //string patternEnw = @"(?<key>\%.)\s(?<value>.*)";//enw
        string patternOtherTxt = @"(?<key>..)\s(?<value>.*)";//otherTxt
        #endregion

        #region for Windows form and Controls
        private bool resizableForm = true;

        private int formBorderRadius = 10;
        private int formBorderSize = 5;
        private Color formBorderColor = ColorsApp.PrimaryColor;
        private Color formBorderColorOnDeActivated = Color.Gray;
        //private PrivateFontCollection pfcRegular;
        //private PrivateFontCollection pfcBold;

        private List<Control> controlsToDrag;

        #endregion

        #region Define variables

        string defaultTextPanelImportDrag = "فایل های خود را انتخاب نمایید\r\nیا در اینجا رها کنید";
        string dragTextPanelImportDrag = "فایل را رها کنید";

        List<BibliographyResourcesFileModel> resourcesFiles;
        int resourceCounter = 0;

        #endregion
        public AddBibliographyForm()
        {
            InitializeComponent();
            this.SuspendLayout();

            DedicatedFunctions.setManualScale(this, new Size(370, 600), new Size(300, 485));
            this.CenterToScreen();

            #region Form initial settings
            this.FormBorderStyle = FormBorderStyle.None;
            this.Padding = new Padding(formBorderSize);
            this.BackColor = formBorderColor;
            this.Opacity = 0;
            this.TopMost = true;
            toolTip1.SetToolTip(btnAlwaysOnTop, "غیرفعال کردن همیشه بالا بودن پنجره");
            this.DoubleBuffered = true;
            setDoubleBuffer(panelMain, true);
            this.SetStyle(ControlStyles.ResizeRedraw, true);

            this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(1, 1, Width, Height, formBorderRadius * 2, formBorderRadius * 2));

            #endregion

            this.FormClosing += (e, a) =>
            {
                AnimateWindow(this.Handle, 100, AnimateWindowFlags.AW_BLEND | AnimateWindowFlags.AW_HIDE);
                Globals.ThisAddIn.DocumentManagerFormVisible = false;
            };

            #region Panel initial settings
            this.panelImportDragFiles.AllowDrop = true;
            this.panelImportDragFiles.DragDrop += new System.Windows.Forms.DragEventHandler(this.panelDragFiles_DragDrop);
            this.panelImportDragFiles.DragEnter += new System.Windows.Forms.DragEventHandler(this.panelDragFiles_DragEnter);
            this.panelImportDragFiles.DragLeave += new System.EventHandler(this.panelDragFiles_DragLeave);
            #endregion

            #region other initialize
            resourcesFiles = new List<BibliographyResourcesFileModel>() { };

            timerOpenAnimation.Enabled = true;
            #endregion

            #region setup Dragging WindowsForm
            controlsToDrag = new List<Control>
                {
                    panelBottom,
                    panelMain,
                    panelTop,
                    panelFiles,
                    panelFilePadding,
                    panelDragContrainer,
                    lblWindowName
                };

            foreach (Control control in controlsToDrag)
            {
                control.MouseDown += new MouseEventHandler(this.StartFormV3_Dragging_MouseDown);
            }
            #endregion

            this.ResumeLayout(false);
        }

        #region (UI) WindowsForm

        #region Configure custom font from Resources(in Properties)
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

        //		if(ctrl.Font.Style == FontStyle.Regular)
        //			ctrl.Font = new System.Drawing.Font(pfcRegular.Families[0] , ctrl.Font.Size , FontStyle.Regular);
        //		else if(ctrl.Font.Style == FontStyle.Bold)
        //			ctrl.Font = new System.Drawing.Font(pfcBold.Families[0] , ctrl.Font.Size , FontStyle.Bold);
        //		else if(ctrl.Font.Style == FontStyle.Italic)
        //			ctrl.Font = new System.Drawing.Font(pfcRegular.Families[0] , ctrl.Font.Size , FontStyle.Italic);
        //		else if(ctrl.Font.Style == FontStyle.Underline)
        //			ctrl.Font = new System.Drawing.Font(pfcRegular.Families[0] , ctrl.Font.Size , FontStyle.Underline);
        //		else if(ctrl.Font.Style == FontStyle.Strikeout)
        //			ctrl.Font = new System.Drawing.Font(pfcRegular.Families[0] , ctrl.Font.Size , FontStyle.Strikeout);
        //	}
        //}
        #endregion

        #region Focus Interaction
        protected override void OnActivated(EventArgs e)
        {
            this.BackColor = formBorderColor;
            base.OnActivated(e);
        }
        protected override void OnDeactivate(EventArgs e)
        {
            this.BackColor = formBorderColorOnDeActivated;
            base.OnDeactivate(e);
        }
        #endregion

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
        public string getMosuePosition(System.Drawing.Point mouse, Form form)
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
                System.Drawing.Point pt = PointToClient(new System.Drawing.Point(x, y));

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

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Source replaceSources = null;
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            int oldSourcesCount = doc.Bibliography.Sources.Count;
            int newSourcesCount = 0;
            int conflictCounter = 0;
            if (resourcesFiles.Count != 0)
            {
                foreach (BibliographyResourcesFileModel item in resourcesFiles)
                {
                    bool isExist = false;
                    try
                    {
                        foreach (Microsoft.Office.Interop.Word.Source item2 in doc.Bibliography.Sources)
                        {
                            int startTagLocation = item.Content.IndexOf("<b:Tag>") + 7;
                            int endTagLocation = item.Content.IndexOf("</b:Tag>");
                            string getTagResource = item.Content.Substring(startTagLocation, endTagLocation - startTagLocation);
                            if (item2.Tag == getTagResource)
                            {
                                isExist = true;
                                DialogResult dialogResult = DedicatedFunctions.ShowMessage("در لیست منابع سند، نام تگ\n" + item2.Tag + "\nوجود دارد، آیا تمایل به جایگزینی با منبع جدید دارید؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                                if (dialogResult == DialogResult.Yes)
                                {
                                    replaceSources = item2;
                                    conflictCounter++;
                                }
                            }
                        }
                        if (isExist)
                        {
                            if (replaceSources != null)
                            {
                                replaceSources.Delete();
                                doc.Bibliography.Sources.Add(item.Content);

                            }
                        }
                        else
                        {
                            doc.Bibliography.Sources.Add(item.Content);
                        }

                    }
                    catch (Exception)
                    {
                        DedicatedFunctions.ShowMessage("خطایی در افزودن سند در آدرس زیر به وجود آمد:\n" + item.Path + "\nاین فایل به لیست منابع سند اضافه نخواهد شد", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                newSourcesCount = doc.Bibliography.Sources.Count;
                if (newSourcesCount + conflictCounter <= oldSourcesCount)
                {
                    DedicatedFunctions.ShowMessage("هیچ یک از منابع شما به لیست منابع سند اضافه نشد!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    DedicatedFunctions.ShowMessage(((newSourcesCount + conflictCounter) - oldSourcesCount).ToString() + " منبع به لیست منابع سند اضافه گردید.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            this.DialogResult = DialogResult.OK;
            doc.Activate();
            doc.ActiveWindow.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateNormal;
            Close();
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
            Close();
        }
        private void timerOpenAnimation_Tick(object sender, EventArgs e)
        {
            Opacity += 0.15;

            if (Opacity == 1)
            {
                timerOpenAnimation.Enabled = false;
            }
        }
        private void panelDragFiles_DragEnter(object sender, DragEventArgs e)
        {
            lblDragStatus.Text = dragTextPanelImportDrag;
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
            else
                e.Effect = DragDropEffects.None;
        }
        private void panelDragFiles_DragLeave(object sender, EventArgs e)
        {
            lblDragStatus.Text = defaultTextPanelImportDrag;
        }
        private void panelDragFiles_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            createFileObject(files);
            lblDragStatus.Text = defaultTextPanelImportDrag;
        }
        private void importDragFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Filter = "Resources Files|*.txt;*.bib;*.enw;*.ris";

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                createFileObject(fileDialog.FileNames);
            }
        }
        private void createFileObject(string[] filePaths)
        {
            Control[] controlsExist = getExistControlByFilePath(filePaths);
            if (controlsExist.Length != 0)
            {
                if (controlsExist.Length == 1)
                    DedicatedFunctions.ShowMessage("فایل وارد شده قبلا در لیست وارد شده است; فایل قبلی از لیست حذف شده و فایل جدید جایگزین آن خواهد شد.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                    DedicatedFunctions.ShowMessage("برخی از فایل های وارد شده قبلا در لیست وارد شده است; فایل های قبلی از لیست حذف شده و فایل های جدید جایگزین آن خواهد شد.", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                foreach (Control item in controlsExist)
                {
                    removeFileControl(item);
                }
            }
            foreach (var filePath in filePaths)
            {
                resourceCounter++;
                P_ResourceFileControl resourceFileControl = new P_ResourceFileControl();

                if (Path.GetExtension(filePath).ToLower() == ".txt" || Path.GetExtension(filePath).ToLower() == ".bib" ||
                    Path.GetExtension(filePath).ToLower() == ".ris" || Path.GetExtension(filePath).ToLower() == ".enw")
                {
                    string[] resourcesXml = { "-1" };
                    if (Path.GetExtension(filePath).ToLower() == ".bib")
                    {
                        resourcesXml = readAndCheckBibTexFile(File.ReadAllText(filePath));
                    }
                    else if (Path.GetExtension(filePath).ToLower() == ".txt")
                    {
                        resourcesXml = readAndCheckBibTexFile(File.ReadAllText(filePath));
                        if (resourcesXml[0] == "-1" || resourcesXml[0] == "-2")
                        {
                            resourcesXml = readAndCheckOtherTxtFile(File.ReadAllText(filePath));
                        }
                    }
                    //else if (Path.GetExtension(filePath).ToLower() == ".ris")
                    //{
                    //    resourcesXml = readAndCheckRisFile(File.ReadAllText(filePath));
                    //}
                    //else if (Path.GetExtension(filePath).ToLower() == ".enw")
                    //{
                    //    resourcesXml = readAndCheckEnwFile(File.ReadAllText(filePath));
                    //}
                    if (resourcesXml[0] == "-1" || resourcesXml[0] == "-2")
                    {
                        resourceFileControl.StatusColor = P_ResourceFileControl.FileStatus.Fail;
                        resourceFileControl.StatusText = "محتوای فایل نادرست است";
                    }
                    else
                    {
                        foreach (string item in resourcesXml)
                        {
                            resourcesFiles.Add(new BibliographyResourcesFileModel(
                                resourceCounter,
                                Path.GetFileName(filePath),
                                Path.GetExtension(filePath),
                                filePath,
                                item));
                        }
                        if (resourcesFiles.Count != 0)
                        {
                            btnSubmit.Enabled = true;
                        }
                        resourceFileControl.StatusColor = P_ResourceFileControl.FileStatus.Success;
                        resourceFileControl.StatusText = "محتوای فایل مورد تایید است";
                    }
                }
                else
                {
                    resourceFileControl.StatusColor = P_ResourceFileControl.FileStatus.Fail;
                    resourceFileControl.StatusText = "نوع فایل پشتیبانی نمیشود";
                }

                resourceFileControl.BackgroundColor = Color.White;
                resourceFileControl.Name = "resourceControl_" + resourceCounter.ToString();
                resourceFileControl.Font = new System.Drawing.Font(Constants.FontNames.fontVazirmatn, 8.25F, FontStyle.Regular);
                resourceFileControl.Padding = new Padding(4);
                resourceFileControl.RightToLeft = RightToLeft.Yes;
                resourceFileControl.Size = new Size(494, 50);
                resourceFileControl.TabIndex = resourceCounter + 1;
                resourceFileControl.Dock = DockStyle.Top;
                resourceFileControl.StatusBoxSize = new System.Drawing.Size(154, 65);
                resourceFileControl.PaddingPanelMain = new Padding(7);
                resourceFileControl.PaddingPicFileExtension = new Padding(3);
                resourceFileControl.PaddingRemoveButton = new Padding(2);
                resourceFileControl.PaletteDrawBorder = PaletteDrawBorders.All;
                resourceFileControl.BorderColor = Color.FromArgb((byte)(8), (byte)(142), (byte)(254));
                resourceFileControl.BorderRadius = 15;
                resourceFileControl.BorderSize = 2;
                resourceFileControl.FileName = Path.GetFileName(filePath);
                resourceFileControl.FilePath = filePath;
                resourceFileControl.onRemoveControlClick += new P_ResourceFileControl.btnRemoveControl_Click(onRemoveControlClick);

                resourceFileControl.PaletteDrawBorderRemoveButton = PaletteDrawBorders.All;
                resourceFileControl.BorderColorRemoveButton = resourceFileControl.BackColor;
                resourceFileControl.BorderRadiusRemoveButton = 10;
                resourceFileControl.BorderSizeRemoveButton = 1;

                string[] exceptionDrag = { "picRemoveControl" };
                setDraggbleWithExceptions(resourceFileControl.Controls, exceptionDrag);

                if (Path.GetExtension(filePath) == ".txt")
                    resourceFileControl.Type = P_ResourceFileControl.FileType.txt;
                else if (Path.GetExtension(filePath) == ".bib")
                    resourceFileControl.Type = P_ResourceFileControl.FileType.bib;
                else if (Path.GetExtension(filePath) == ".ris")
                    resourceFileControl.Type = P_ResourceFileControl.FileType.ris;
                else if (Path.GetExtension(filePath) == ".enw")
                    resourceFileControl.Type = P_ResourceFileControl.FileType.enw;
                else
                    resourceFileControl.Type = P_ResourceFileControl.FileType.unknown;

                //setAllControlsFont(resourceFileControl.Controls);
                panelFiles.Controls.Add(resourceFileControl);
                resourceFileControl.BringToFront();

            }
        }
        public void setDraggbleWithExceptions(Control.ControlCollection ctrls, string[] exceptionNames)
        {
            foreach (Control ctrl in ctrls)
            {
                if (ctrl.Controls != null)
                    setDraggbleWithExceptions(ctrl.Controls, exceptionNames);

                foreach (string item in exceptionNames)
                {
                    if (ctrl.Name != item)
                        ctrl.MouseDown += new MouseEventHandler(this.StartFormV3_Dragging_MouseDown);
                }
            }
        }

        public Control[] getExistControlByFilePath(string[] filePaths)
        {
            List<Control> listResult = new List<Control>();
            foreach (Control ctrl in panelFiles.Controls)
            {
                foreach (string item in filePaths)
                {
                    if (((P_ResourceFileControl)ctrl).FilePath == item)
                    {
                        listResult.Add(ctrl);
                    }
                }
            }
            return listResult.ToArray();
        }
        private string[] readAndCheckBibTexFile(string line)
        {
            //string type, citationKey, caption, author, journal, volume, number, pages, year, publisher, city, editor, edition, url, doi;
            //
            //if (line.IndexOf("@") != -1 && line.IndexOf("{") != -1 && line.IndexOf(",") != -1)
            //{
            //    type = line.Substring(line.IndexOf("@") + 1, line.IndexOf("{") - line.IndexOf("@") - 1).Trim();
            //    if (type.ToLower() == "book") type = "Book";
            //    else if (type.ToLower() == "article") type = "Journal Article";
            //    else return false;
            //    citationKey = line.Substring(line.IndexOf("{") + 1, line.IndexOf(",") - line.IndexOf("{") - 1).Trim();
            //}
            //else
            //    return false;
            //caption = getValue(line, "caption");
            //author = getValue(line, "author").Replace(" and", ";");
            //journal = getValue(line, "journal");
            //volume = getValue(line, "volume");
            //number = getValue(line, "number");
            //pages = getValue(line, "pages");
            //year = getValue(line, "year");
            //publisher = getValue(line, "publisher");
            //city = getValue(line, "city");
            //editor = getValue(line, "editor");
            //edition = getValue(line, "edition");
            //url = getValue(line, "url");
            //doi = getValue(line, "doi");
            //
            //if (type == "Journal Article" && (author == "-1" || caption == "-1" || journal == "-1" || year == "-1"))
            //{
            //    return false;
            //}
            //else if (type == "Book" && ((author == "-1" && editor == "-1") || caption == "-1" || publisher == "-1" || year == "-1"))
            //{
            //    return false;
            //}
            //string strXml = @$"<b:Source xmlns:b=""http://nazariansani.ir"">
            //                <b:Tag>{citationKey}</b:Tag> 
            //                <b:SourceType>{type}</b:SourceType> 
            //                <b:Author> 
            //                    <b:Author> 
            //                        <b:NameList> 
            //                            <b:Person> 
            //                                <b:Last>{author}</b:Last> 
            //                                <b:First></b:First> 
            //                            </b:Person> 
            //                        </b:NameList> 
            //                    </b:Author> 
            //                </b:Author> 
            //                <b:Caption>{caption}</b:Caption> 
            //                <b:Year>{year}</b:Year> 
            //                <b:City>{city}</b:City> 
            //                <b:Volume>{volume}</b:Volume> 
            //                <b:Issue>{number}</b:Issue> 
            //                <b:Pages>{pages}</b:Pages> 
            //                <b:JournalName>{journal}</b:JournalName> 
            //                <b:Publisher>{publisher}</b:Publisher> 
            //                <b:Edition>{edition}</b:Edition> 
            //                <b:URL>{url}</b:URL> 
            //                <b:DOI>{doi}</b:DOI> 
            //            </b:Source>";
            //
            ////                                <b:Guid>{6D86D06C-9022-4932-8D4C-84C2B0843381}</b:Guid> 
            //
            //if (Regex.IsMatch(line, @"^[\u0600-\u06FF]+$"))//is persian Chars
            //{
            //    DedicatedFunctions.ShowMessage("Tested Regex");
            //}
            ////Globals.ThisAddIn.Application.ActiveDocument.Bibliography.Sources.Add(strXml);
            //return true;

            MatchCollection matches = Regex.Matches(line, patternbibTex);
            if (matches.Count != 0)
            {
                string strXml = "";
                List<string> list = new List<string>();

                foreach (Match m in matches)
                {
                    strXml = "<b:Source xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\">\r\n";
                    strXml += $"<b:Tag>{m.Groups["citationKey"].Value}</b:Tag>\r\n";

                    if (m.Groups["CContentType"].Value.Trim().ToLower() == "book")
                        strXml += "<b:SourceType>Book</b:SourceType>\r\n";
                    else if (m.Groups["CContentType"].Value.Trim().ToLower() == "article")
                    {
                        strXml += "<b:SourceType>JournalArticle</b:SourceType>\r\n";
                    }
                    else
                        strXml += $"<b:SourceType>{m.Groups["CContentType"].Value}</b:SourceType>\r\n";


                    for (int i = 0; i < m.Groups["key"].Captures.Count; i++)
                    {
                        if (m.Groups["key"].Captures[i].Value.Trim().ToLower() == "author")
                        {
                            strXml += $"<b:Author><b:Author><b:NameList><b:Person><b:Last>{m.Groups["value"].Captures[i].Value.Replace(" and", ";")}</b:Last><b:First></b:First></b:Person></b:NameList></b:Author></b:Author>\r\n";
                        }
                        else if ((m.Groups["key"].Captures[i].Value.Trim().ToLower() == "journal"))
                        {
                            strXml += $"<b:JournalName>{m.Groups["value"].Captures[i].Value}</b:JournalName>\r\n";
                        }
                        else if ((m.Groups["key"].Captures[i].Value.Trim().ToLower() == "url"))
                        {
                            strXml += $"<b:URL>{m.Groups["value"].Captures[i].Value}</b:URL>\r\n";
                        }
                        else if ((m.Groups["key"].Captures[i].Value.Trim().ToLower() == "DOI"))
                        {
                            strXml += $"<b:DOI>{m.Groups["value"].Captures[i].Value}</b:DOI>\r\n";
                        }
                        else
                        {
                            string keyFirstLetterUpper = FirstLetterToUpper(m.Groups["key"].Captures[i].Value);
                            strXml += $"<b:{keyFirstLetterUpper}>{m.Groups["value"].Captures[i].Value}</b:{keyFirstLetterUpper}>\r\n";
                        }
                    }
                    strXml += "</b:Source>";
                    list.Add(strXml);
                }
                foreach (string item in list)//if was not XML Format
                {
                    if (!IsValidXhtml(item))
                    {
                        string[] result = { "-2" };
                        return result;
                    }
                }
                return list.ToArray();
            }
            else
            {
                string[] result = { "-1" };
                return result;
            }
        }
        private string[] readAndCheckOtherTxtFile(string line)
        {
            MatchCollection matches = Regex.Matches(line, patternOtherTxt);
            if (matches.Count != 0)
            {
                string strXml = "";
                List<string> list = new List<string>();

                foreach (Match m in matches)
                {
                    strXml = "<b:Source xmlns:b=\"http://schemas.openxmlformats.org/officeDocument/2006/bibliography\">\r\n";
                    strXml += $"<b:Tag>{m.Groups["citationKey"].Value}</b:Tag>\r\n";

                    if (m.Groups["CContentType"].Value.Trim().ToLower() == "book")
                        strXml += "<b:SourceType>Book</b:SourceType>\r\n";
                    else if (m.Groups["CContentType"].Value.Trim().ToLower() == "article")
                    {
                        strXml += "<b:SourceType>JournalArticle</b:SourceType>\r\n";
                    }
                    else
                        strXml += $"<b:SourceType>{m.Groups["CContentType"].Value}</b:SourceType>\r\n";


                    for (int i = 0; i < m.Groups["key"].Captures.Count; i++)
                    {
                        if (m.Groups["key"].Captures[i].Value.Trim().ToLower() == "author")
                        {
                            strXml += $"<b:Author><b:Author><b:NameList><b:Person><b:Last>{m.Groups["value"].Captures[i].Value.Replace(" and", ";")}</b:Last><b:First></b:First></b:Person></b:NameList></b:Author></b:Author>\r\n";
                        }
                        else if ((m.Groups["key"].Captures[i].Value.Trim().ToLower() == "journal"))
                        {
                            strXml += $"<b:JournalName>{m.Groups["value"].Captures[i].Value}</b:JournalName>\r\n";
                        }
                        else if ((m.Groups["key"].Captures[i].Value.Trim().ToLower() == "url"))
                        {
                            strXml += $"<b:URL>{m.Groups["value"].Captures[i].Value}</b:URL>\r\n";
                        }
                        else if ((m.Groups["key"].Captures[i].Value.Trim().ToLower() == "DOI"))
                        {
                            strXml += $"<b:DOI>{m.Groups["value"].Captures[i].Value}</b:DOI>\r\n";
                        }
                        else
                        {
                            string keyFirstLetterUpper = FirstLetterToUpper(m.Groups["key"].Captures[i].Value);
                            strXml += $"<b:{keyFirstLetterUpper}>{m.Groups["value"].Captures[i].Value}</b:{keyFirstLetterUpper}>\r\n";
                        }
                    }
                    strXml += "</b:Source>";
                    list.Add(strXml);
                }
                foreach (string item in list)//if was not XML Format
                {
                    if (!IsValidXhtml(item))
                    {
                        string[] result = { "-2" };
                        return result;
                    }
                }
                return list.ToArray();
            }
            else
            {
                string[] result = { "-1" };
                return result;
            }
        }
        public string FirstLetterToUpper(string str)
        {
            if (str == null)
                return null;

            if (str.Length > 1)
                return char.ToUpper(str[0]) + str.Substring(1);

            return str.ToUpper();
        }
        public static bool IsValidXhtml(string text)
        {
            try
            {
                if (!string.IsNullOrEmpty(text))
                {
                    System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                    xmlDoc.LoadXml(text);
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (System.Xml.XmlException)
            {
                return false;
            }
        }
        private void btnAlwaysOnTop_Click(object sender, EventArgs e)
        {
            this.TopMost = !this.TopMost;
            if (this.TopMost)
            {
                btnAlwaysOnTop.BackgroundImage = Resources.office_push_unpin;
                toolTip1.SetToolTip(btnAlwaysOnTop, "غیرفعال کردن همیشه بالا بودن پنجره");
            }
            else
            {
                btnAlwaysOnTop.BackgroundImage = Resources.office_push_pin;
                toolTip1.SetToolTip(btnAlwaysOnTop, "فعال کردن همیشه بالا بودن پنجره");
            }
        }
        private void onRemoveControlClick(object sender, EventArgs e)
        {
            removeFileControl(((P_ResourceFileControl)sender));
        }
        public void removeFileControl(Control control)
        {
            string idControl = control.Name.Substring(control.Name.IndexOf("_") + 1);

            BibliographyResourcesFileModel resouceFileToRemove = null;
            foreach (BibliographyResourcesFileModel item in resourcesFiles)
            {
                if (item.Id == int.Parse(idControl))
                {
                    resouceFileToRemove = item;
                    break;
                }
            }
            if (resouceFileToRemove != null)
            {
                resourcesFiles.Remove(resouceFileToRemove);
            }
            if (resourcesFiles.Count == 0)
            {
                btnSubmit.Enabled = false;
            }

            this.Controls.RemoveByKey(control.Name);
            control.Dispose();

        }
    }
}