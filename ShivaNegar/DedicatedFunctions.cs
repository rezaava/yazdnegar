using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Win32;
using ShivaNegar.Constants;
using ShivaNegar.Forms;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View;
using ShivaNegar.Models;
using ShivaNegar.TaskPanes.CrossReference;
using ShivaNegar.Templates;


// new using 

using Task = System.Threading.Tasks.Task;

namespace ShivaNegar
{
    internal class DedicatedFunctions
    {
        internal static void setManualScale(Form form, Size baseSize, Size minimumSize)
        {
            #region set Manual Scale
            //this.AutoScaleMode = AutoScaleMode.Dpi;
            form.PerformAutoScale();

            // دریافت DPI فعلی
            float dpiX, dpiY;
            using (Graphics g = form.CreateGraphics())
            {
                dpiX = g.DpiX;
                dpiY = g.DpiY;
            }

            // محاسبه ابعاد جدید بر اساس DPI
            float scaleX = dpiX / 96;
            float scaleY = dpiY / 96;

            int screenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int screenHeight = Screen.PrimaryScreen.WorkingArea.Height;

            int targetWidth = (int)(baseSize.Width * scaleX);
            int targetHeight = (int)(baseSize.Height * scaleY);

            int targetMinimumWidth = (int)(minimumSize.Width * scaleX);
            int targetMinimumHeight = (int)(minimumSize.Height * scaleY);

            if (screenWidth > targetWidth && screenHeight > targetHeight)
            {
                form.Width = targetWidth;
                form.Height = targetHeight;
                //form.MinimumSize = new System.Drawing.Size(targetWidth - 100 , targetHeight - 100);
                form.MinimumSize = new System.Drawing.Size(targetMinimumWidth, targetMinimumHeight);
            }
            else
            {
                form.Width = baseSize.Width;
                form.Height = baseSize.Height;
                form.MinimumSize = new System.Drawing.Size(minimumSize.Width, minimumSize.Height);

                //if(screenHeight > targetHeight)
                //	this.Height = screenHeight;
            }

            #endregion
        }

        internal static string[] getUndoList()
        {
            var undoControls = Globals.ThisAddIn.Application.CommandBars.FindControls(Id: 128);

            List<string> undoList = new List<string>();
            foreach (Microsoft.Office.Core.CommandBarControl undoControl in undoControls)
            {
                //if(undoControl.Parent.Name != "Send Mail" && undoControl.Parent.Name != "Outlook Send Mail")
                //{
                //	Debug.WriteLine(undoControl.Parent.Name);
                //	var undoComboBox = undoControl as Microsoft.Office.Core.CommandBarComboBox;
                //	if(undoComboBox != null && undoComboBox.accChildCount > 0)
                //	{
                //		undoList = new List<string>();
                //		for(int i = 1 ; i <= undoComboBox.ListCount ; i++)
                //		{
                //			undoList.Add(undoComboBox.List[i]);
                //			Debug.WriteLine("\t\t" + undoComboBox.List[i]);
                //			//MessageBox.Show(undoComboBox.List[i]);
                //		}
                //	}
                //}

                if (undoControl.Parent.Name == "Standard" && undoControl is Microsoft.Office.Core.CommandBarComboBox undoComboBox)
                {
                    //Debug.WriteLine(undoControl.Parent.Name);
                    //var undoComboBox = undoControl as Microsoft.Office.Core.CommandBarComboBox;
                    if (undoComboBox != null && undoComboBox.accChildCount > 0)
                    {
                        for (int i = 1; i <= undoComboBox.ListCount; i++)
                        {
                            undoList.Add(undoComboBox.List[i]);
                            //Debug.WriteLine("\t\t" + undoComboBox.List[i]);
                        }
                        break;
                    }
                }
                else if (undoControl.Parent.Name == "Word for Windows 2.0" && undoControl is Microsoft.Office.Core.CommandBarComboBox undoComboBox2)
                {
                    //Debug.WriteLine(undoControl.Parent.Name);
                    //var undoComboBox = undoControl as Microsoft.Office.Core.CommandBarComboBox;
                    if (undoComboBox2 != null && undoComboBox2.accChildCount > 0)
                    {
                        for (int i = 1; i <= undoComboBox2.ListCount; i++)
                        {
                            undoList.Add(undoComboBox2.List[i]);
                            //Debug.WriteLine("\t\t" + undoComboBox2.List[i]);
                        }
                        break;
                    }
                }
            }
            return undoList.ToArray();
        }

        #region dll Imports and Constants

        /*
        // For Windows Mobile, replace user32.dll with coredll.dll
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr hWndChildAfter, string className, string windowTitle);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);


        [DllImport("user32.dll")]
        static extern bool LockWindowUpdate(IntPtr hWndLock);
        [DllImport("User32")]
        private static extern int ShowWindow(int hwnd, int nCmdShow);//private const int SW_HIDE = 0;private const int SW_SHOW = 5;
        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);


        void CloseWindow(IntPtr hwnd)
        {
            SendMessage(hwnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
        }
        IntPtr Ret, ChildRet;

        private const UInt32 WM_CLOSE = 0x0010;
        private const int WM_SETTEXT = 0X000C;
        private const int WM_GETTEXT = 0x000D;
        private const int BM_CLICK = 0x00F5;

        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOCOPYBITS = 0x0100;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_HIDEWINDOW = 0x0080;
        private const uint SWP_DEFERERASE = 0x2000;
        */

        private const int KB_Escape = 0x0029;

        //Virtual Keyboard
        private const int VK_CONTROL = 0x11;
        private const int VK_ESCAPE = 0x1B;
        private const int VK_RETURN = 0x0D;//Enter Key
        private const int VK_SHIFT = 0x0D;//SHIFT Key

        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int WM_CHAR = 0x0102;

        private const int BM_CLICK = 0x00F5;
        private const int WM_GETTEXTLENGTH = 0x000E;
        private const int WM_GETTEXT = 0x000D;
        private const int WM_SETTEXT = 0x000C;

        private const int WM_CLOSE = 0x0010;
        private const int WM_DESTROY = 0x0002;
        private const int WM_QUIT = 0x0012;

        private const uint SWP_NOZORDER = 0x0004;
        private const uint SWP_NOCOPYBITS = 0x0100;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint SWP_HIDEWINDOW = 0x0080;
        private const uint SWP_DEFERERASE = 0x2000;

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString,
    int nMaxCount);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr hWndChildAfter, string className, string windowTitle);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

        [DllImport("user32.dll")]
        internal static extern int SendMessage(IntPtr hWnd, int msg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        internal static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);


        // وارد کردن توابع از user32.dll
        [DllImport("user32.dll")]
        internal static extern IntPtr SetFocus(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        internal const int SW_SHOW = 5;

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        #endregion

        #region Other (useless)
        internal void SaNa_Index()//is UseFull?
        {
            String strFile = "";
            Microsoft.Office.Core.FileDialog sda = Globals.ThisAddIn.Application.FileDialog[Microsoft.Office.Core.MsoFileDialogType.msoFileDialogFilePicker];
            sda.Filters.Clear();
            sda.Filters.Add("Text File", "*.txt", 1);
            sda.Filters.Add("Word Documents", "*.docx", 1);
            sda.Title = "Open Index AutoMark File ";
            sda.AllowMultiSelect = false;
            sda.InitialFileName = "";

            if (sda.Show() == 1)
            {
                strFile = sda.SelectedItems.Item(1);
            }
            Globals.ThisAddIn.Application.ActiveWindow.ActivePane.View.ShowAll = true;

            if (strFile != "")
            {
                Globals.ThisAddIn.Application.ActiveDocument.Indexes.AutoMarkEntries(strFile);
            }
            //.Hidden Paragraph Marks
            Globals.ThisAddIn.Application.ActiveWindow.ActivePane.View.ShowAll = !Globals.ThisAddIn.Application.ActiveWindow.ActivePane.View.ShowAll;
            //'        With ActiveDocument
            //'        .Indexes.Add Range:=Selection.Range, HeadingSeparator:= _
            //'            wdHeadingSeparatorNone, Type:=wdIndexIndent, RightAlignPageNumbers:= _
            //'            True, NumberOfColumns:=4, IndexLanguage:=wdPersian
            //'        .Indexes(1).TabLeader = wdTabLeaderDashes
        }

        #endregion

        #region Save and SaveAs Document
        internal static bool saveDocument(Document doc)
        {
            try
            {
                Globals.ThisAddIn.DisableSaveEvent = false;
                doc.Save();

                //Globals.ThisAddIn.Application.NormalTemplate.Saved = true;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        internal static void saveAsDocument(Document doc, string path, bool raiseSaveEvent = false)
        {
            if (!raiseSaveEvent)
                Globals.ThisAddIn.DisableSaveEvent = true;
            doc.SaveAs2(path, WdSaveFormat.wdFormatDocumentDefault);
            Globals.ThisAddIn.DisableSaveEvent = false;

            //Globals.ThisAddIn.Application.NormalTemplate.Saved = true;
        }
        #endregion

        #region open Document
        internal async static void checkAndInitializeOpenedDocument(Document doc)
        {
            Globals.ThisAddIn.accessedInOpen = false;

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

            doc.ContentControlOnExit += Globals.ThisAddIn.Doc_ContentControlOnExit;
            if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                //await DedicatedFunctions.checkingUpdate();

                if (!Globals.ThisAddIn.accessedInStartup)
                {
                    Globals.ThisAddIn.DisableEvents = true;
                    DedicatedFunctions.checkingFonts();
                    Ribbon.InitializeRibbon(StringConstant.NameOfProject);
                    DedicatedFunctions.checkPersianKeyboardLayout();

                    Ribbon.loadKeyboardShortcut();
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);
                    Globals.ThisAddIn.DisableEvents = false;
                }

                if (doc != null && File.Exists(doc.FullName))
                {
                    doc.UndoClear();
                    DedicatedFunctions.saveDocument(doc);
                }

                Globals.ThisAddIn.accessedInStartup = true;
                Globals.ThisAddIn.accessedInOpen = true;
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                DedicatedFunctions.showSplashScreen();

                bool checkResult = await DedicatedFunctions.checkingUpdate();

                if (checkResult)
                {
                    Globals.ThisAddIn.DisableEvents = true;
                    if (!Globals.ThisAddIn.accessedInStartup)
                    {
                        DedicatedFunctions.initialSettings();

                        Ribbon.loadKeyboardShortcut();

                        DedicatedFunctions.checkingFonts();

                        DedicatedFunctions.checkPersianKeyboardLayout();

                        Globals.ThisAddIn.accessedInStartup = true;
                    }

                    #region Scroll
                    try
                    {
                        int scrollTo = int.Parse(DedicatedFunctions.getStaticVariableValue(doc, VariableDocumentIDs._variable_document_PositionCurrentPage.ToString()));
                        DedicatedFunctions.scrollToPage(doc.ActiveWindow, doc.ActiveWindow.Selection, scrollTo);
                    }
                    catch (Exception)
                    {
                    }
                    #endregion

                    try
                    {
                        Universities university = DedicatedFunctions.getUniversity(doc);
                        TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

                        //copy Template File for createDocument based on Template
                        string templateName = TemplateAccess.getTemplateFileName(university, (TemplateTypes)templateType);

                        using (Stream stream = TemplateAccess.getTemplateFileStream(university, (TemplateTypes)templateType))
                        {
                            string shivanegarTemplatesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates", "ShivaNegarTemplates");
                            Directory.CreateDirectory(shivanegarTemplatesPath);
                            string templatePath = "";

                            try
                            {
                                templatePath = DedicatedFunctions.copyFileToFolder(stream, templateName, shivanegarTemplatesPath);
                            }
                            catch (Exception)
                            {
                                // on using Template File
                                templatePath = shivanegarTemplatesPath + "\\" + nameof(EmbeddedResourceNames.ShivaNegarShortcut) + ".dotm";
                            }
                            doc.UpdateStylesOnOpen = true;
                            doc.set_AttachedTemplate(templatePath);
                        }
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        Ribbon.InitializeRibbon(StringConstant.NameOfProject + "(" + DedicatedFunctions.getDocumentTypePersianName(doc) + ")");
                    }
                    catch (Exception)
                    {
                        Ribbon.InitializeRibbon(StringConstant.NameOfProject);
                    }
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);

                    Globals.ThisAddIn.Application.ScreenUpdating = true;

                    if (doc != null && File.Exists(doc.FullName))
                    {
                        doc.UndoClear();
                        DedicatedFunctions.saveDocument(doc);
                    }
                    Globals.ThisAddIn.DisableEvents = false;

                    Globals.ThisAddIn.accessedInOpen = true;
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("خطا در گرفتن نسخه افزونه از سرور");
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                    DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges);
                }
            }
            //else if (accessType == DedicatedFunctions.AccessType.AccessDenied_NotExistInWorksapce)
            //{
            //    DedicatedFunctions.ShowErrorMessage("سند لازم است در پوشه محیط کاری قرار داشته باشد، سند بسته میشود");
            //    DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges);
            //}
            else if (accessType == DedicatedFunctions.AccessType.AccessDenied_NoHasAccessing)
            {
                Ribbon.RibbonControlsVisibility(false);
            }
        }
        #endregion

        #region Variables

        #region TemplateType Variable

        /// <summary>
        /// get <see cref="TemplateTypes"/> from Variable
        /// </summary>
        internal static TemplateTypes getTemplateType(string templateTypeName)
        {
            if (templateTypeName == TemplateTypeValues.TemplateType_Parsa)
                return TemplateTypes.ParsaTemplate;
            else if (templateTypeName == TemplateTypeValues.TemplateType_University)
                return TemplateTypes.UniversityTemplate;
            else if (templateTypeName == TemplateTypeValues.TemplateType_ShivaNegar)
                return TemplateTypes.ShivaNegarTemplate;
            else if (templateTypeName == TemplateTypeValues.TemplateType_Nothing)
                return TemplateTypes.Nothing;
            else
                throw new Exception("not founded Template Type!");
        }
        /*		internal static TemplateTypes getTemplateType(Document doc)
				{
					string variableData = DedicatedFunctions.getStaticVariableValue(doc , VariableTypeIDs._variable_type_Template.ToString());

					if(variableData == TemplateTypeValues.TemplateType_Parsa)
						return TemplateTypes.ParsaTemplate;
					else if(variableData == TemplateTypeValues.TemplateType_University)
						return TemplateTypes.UniversityTemplate;
					else if(variableData == TemplateTypeValues.TemplateType_ShivaNegar)
						return TemplateTypes.ShivaNegarTemplate;
					else
						throw new Exception("not founded Template Type!");
				}
		*/
        internal static TemplateTypes getTemplateType(Document doc)
        {
            string variableData = DedicatedFunctions.getStaticVariableValue(doc, VariableTypeIDs._variable_type_Template.ToString());

            if (string.IsNullOrEmpty(variableData))
            {
                DedicatedFunctions.ShowErrorMessage("مقدار نوع قالب در متغیر های سند نادرست است", email: StringConstant.SupportEmail);
                throw new Exception("مقدار نوع قالب در متغیر های سند نادرست است");
            }

            if (variableData == null || variableData == SettingValues.NotExist)
                return TemplateTypes.Nothing;
            else
                return (TemplateTypes)int.Parse(variableData);
        }

        internal static string getTemplateTypePersianName(Document doc)
        {
            string docTypeVariableValue = DedicatedFunctions.getStaticVariableValue(doc, VariableTypeIDs._variable_type_Template.ToString());
            TemplateTypes templateType = (TemplateTypes)int.Parse(docTypeVariableValue);
            string value;

            if (templateType == TemplateTypes.ParsaTemplate)
            {
                value = TemplateTypeValues.TemplateType_Parsa;
            }
            else if (templateType == TemplateTypes.UniversityTemplate)
            {
                value = TemplateTypeValues.TemplateType_University;
            }
            else if (templateType == TemplateTypes.ShivaNegarTemplate)
            {
                value = TemplateTypeValues.TemplateType_ShivaNegar;
            }
            else if (templateType == TemplateTypes.Nothing)
            {
                value = TemplateTypeValues.TemplateType_Nothing;
            }
            else
                throw new Exception("not founded Template Type!");

            return value;
        }
        internal static string getTemplateTypePersianName(int type)
        {
            TemplateTypes templateType = (TemplateTypes)type;
            string value;

            if (templateType == TemplateTypes.ParsaTemplate)
            {
                value = TemplateTypeValues.TemplateType_Parsa;
            }
            else if (templateType == TemplateTypes.UniversityTemplate)
            {
                value = TemplateTypeValues.TemplateType_University;
            }
            else if (templateType == TemplateTypes.ShivaNegarTemplate)
            {
                value = TemplateTypeValues.TemplateType_ShivaNegar;
            }
            else if (templateType == TemplateTypes.Nothing)
            {
                value = TemplateTypeValues.TemplateType_Nothing;
            }
            else if (templateType == TemplateTypes.StudentTemplate)
            {
                value = TemplateTypeValues.TemplateType_Student;
            }
            else if (templateType == TemplateTypes.OfficeTemplate)
            {
                value = TemplateTypeValues.TemplateType_Office;
            }
            else
                throw new Exception("not founded Template Type!");

            return value;
        }
        public static string getTemplateTypePersianName(TemplateTypes templateType)
        {
            if (templateType == TemplateTypes.ParsaTemplate)
                return TemplateTypeValues.TemplateType_Parsa;
            else if (templateType == TemplateTypes.UniversityTemplate)
                return TemplateTypeValues.TemplateType_University;
            else if (templateType == TemplateTypes.ShivaNegarTemplate)
                return TemplateTypeValues.TemplateType_ShivaNegar;
            else if (templateType == TemplateTypes.Nothing)
                return TemplateTypeValues.TemplateType_Nothing;

            else throw new System.Exception();
        }
        public static string getTemplateTypeEnglishName(TemplateTypes templateType)
        {
            if (templateType == TemplateTypes.ParsaTemplate)
                return TemplateTypeValues.TemplateType_ParsaEn;
            else if (templateType == TemplateTypes.UniversityTemplate)
                return TemplateTypeValues.TemplateType_UniversityEn;
            else if (templateType == TemplateTypes.ShivaNegarTemplate)
                return TemplateTypeValues.TemplateType_ShivaNegarEn;
            else if (templateType == TemplateTypes.Nothing)
                return TemplateTypeValues.TemplateType_NothingEn;

            else throw new System.Exception();
        }

        public static int getTemplateID(TemplateTypes templateType, Universities university)
        {
            // اگر الگو دانشگاه بود، شماره الگو رو از لیست الگو دانشگاه برگرداند مگر نه شماره الگو بر اساس templateType تنظیم برگردانده میشود
            if (templateType == TemplateTypes.UniversityTemplate)
            {
                return ((int)university);
            }
            else
                return (int)templateType;
        }
        #endregion

        #region DocumentType Variable

        /// <summary>
        /// get <see cref="DocumentTypes"/> from Variable
        /// </summary>
        internal static DocumentTypes getDocumentType(Document doc)
        {
            string variableData = DedicatedFunctions.getStaticVariableValue(doc, VariableTypeIDs._variable_type_Document.ToString());

            if (string.IsNullOrEmpty(variableData))
            {
                DedicatedFunctions.ShowErrorMessage("مقدار نوع سند در متغیر های سند نادرست است", email: StringConstant.SupportEmail, sendErrorToServer: true);
                throw new Exception("مقدار دانشگاه در متغیر های سند نادرست است");
            }

            if (variableData == null || variableData == SettingValues.NotExist)
                return DocumentTypes.Nothing;
            else
                return (DocumentTypes)int.Parse(variableData);
        }
        internal static DocumentTypes getDocumentType(string documentTypeName)
        {
            if (string.IsNullOrEmpty(documentTypeName))
            {
                return DocumentTypes.Nothing;
            }

            if (documentTypeName == DocumentTypeValues.DocumentType_DissertationFa)
                return DocumentTypes.Dissertation;
            else if (documentTypeName == DocumentTypeValues.DocumentType_ProjectFa)
                return DocumentTypes.Project;
            else if (documentTypeName == DocumentTypeValues.DocumentType_ThesisFa)
                return DocumentTypes.Thesis;
            else if (documentTypeName.Contains(DocumentTypeValues.DocumentType_SchoolResearchFa))
                return DocumentTypes.SchoolResearch;
            else if (documentTypeName.Contains(DocumentTypeValues.DocumentType_NothingFa))
                return DocumentTypes.Nothing;
            else
                throw new Exception("not founded Document Type!");
        }

        internal static string getDocumentTypePersianName(int type)
        {
            DocumentTypes documentType = (DocumentTypes)type;
            string value;

            if (documentType == DocumentTypes.Project)
            {
                value = DocumentTypeValues.DocumentType_ProjectFa;
            }
            else if (documentType == DocumentTypes.Thesis)
            {
                value = DocumentTypeValues.DocumentType_ThesisFa;
            }
            else if (documentType == DocumentTypes.Dissertation)
            {
                value = DocumentTypeValues.DocumentType_DissertationFa;
            }
            else if (documentType == DocumentTypes.SchoolResearch)
            {
                value = DocumentTypeValues.DocumentType_SchoolResearchFa;
            }
            else if (documentType == DocumentTypes.Nothing)
            {
                value = DocumentTypeValues.DocumentType_NothingFa;
            }
            else
                throw new Exception("not founded Document Type!");

            return value;
        }
        internal static string getDocumentTypePersianName(Document doc)
        {
            string docTypeVariableValue = DedicatedFunctions.getStaticVariableValue(doc, VariableTypeIDs._variable_type_Document.ToString());
            DocumentTypes documentType = (DocumentTypes)int.Parse(docTypeVariableValue);
            string value;

            if (documentType == DocumentTypes.Project)
            {
                value = DocumentTypeValues.DocumentType_ProjectFa;
            }
            else if (documentType == DocumentTypes.Thesis)
            {
                value = DocumentTypeValues.DocumentType_ThesisFa;
            }
            else if (documentType == DocumentTypes.Dissertation)
            {
                value = DocumentTypeValues.DocumentType_DissertationFa;
            }
            else if (documentType == DocumentTypes.SchoolResearch)
            {
                value = DocumentTypeValues.DocumentType_SchoolResearchFa;
            }
            else if (documentType == DocumentTypes.Nothing)
            {
                value = DocumentTypeValues.DocumentType_NothingFa;
            }
            else
                throw new Exception("not founded Document Type!");

            return value;
        }
        public static string getDocumentTypePersianName(DocumentTypes documentType)
        {
            if (documentType == DocumentTypes.Project)
                return DocumentTypeValues.DocumentType_ProjectFa;
            else if (documentType == DocumentTypes.Thesis)
                return DocumentTypeValues.DocumentType_ThesisFa;
            else if (documentType == DocumentTypes.Dissertation)
                return DocumentTypeValues.DocumentType_DissertationFa;
            else if (documentType == DocumentTypes.SchoolResearch)
                return DocumentTypeValues.DocumentType_SchoolResearchFa;
            else if (documentType == DocumentTypes.Nothing)
                return DocumentTypeValues.DocumentType_NothingFa;
            else throw new System.Exception();
        }

        public static string getDocumentTypeEnglishName(DocumentTypes documentType)
        {
            if (documentType == DocumentTypes.Project)
                return DocumentTypeValues.DocumentType_ProjectEn;
            else if (documentType == DocumentTypes.Thesis)
                return DocumentTypeValues.DocumentType_ThesisEn;
            else if (documentType == DocumentTypes.Dissertation)
                return DocumentTypeValues.DocumentType_DissertationEn;
            else if (documentType == DocumentTypes.SchoolResearch)
                return DocumentTypeValues.DocumentType_SchoolResearchEn;
            else if (documentType == DocumentTypes.Nothing)
                return DocumentTypeValues.DocumentType_NothingEn;
            else throw new System.Exception();
        }

        internal static string getDocumentTypeEnglishName(Document doc)
        {
            string docTypeVariableValue = DedicatedFunctions.getStaticVariableValue(doc, VariableTypeIDs._variable_type_Document.ToString());
            DocumentTypes documentType = (DocumentTypes)int.Parse(docTypeVariableValue);
            string value;
            if (documentType == DocumentTypes.Project)
            {
                value = DocumentTypeValues.DocumentType_ProjectEn;
            }
            else if (documentType == DocumentTypes.Thesis)
            {
                value = DocumentTypeValues.DocumentType_ThesisEn;
            }
            else if (documentType == DocumentTypes.Dissertation)
            {
                value = DocumentTypeValues.DocumentType_DissertationEn;
            }
            else if (documentType == DocumentTypes.SchoolResearch)
            {
                value = DocumentTypeValues.DocumentType_SchoolResearchEn;
            }
            else if (documentType == DocumentTypes.Nothing)
            {
                value = DocumentTypeValues.DocumentType_NothingEn;
            }
            else
                throw new Exception("not founded Document Type!");

            return value;
        }

        internal static string getDocumentTypeEnglishName(int type)
        {
            DocumentTypes documentType = (DocumentTypes)type;
            string value;
            if (documentType == DocumentTypes.Project)
            {
                value = DocumentTypeValues.DocumentType_ProjectEn;
            }
            else if (documentType == DocumentTypes.Thesis)
            {
                value = DocumentTypeValues.DocumentType_ThesisEn;
            }
            else if (documentType == DocumentTypes.Dissertation)
            {
                value = DocumentTypeValues.DocumentType_DissertationEn;
            }
            else if (documentType == DocumentTypes.SchoolResearch)
            {
                value = DocumentTypeValues.DocumentType_SchoolResearchEn;
            }
            else if (documentType == DocumentTypes.Nothing)
            {
                value = DocumentTypeValues.DocumentType_NothingEn;
            }
            else
                throw new Exception("not founded Document Type!");

            return value;
        }

        #endregion

        #region AcademicDegree
        internal static AcademicDegrees getAcademicDegreeID(Document doc)
        {
            string academicID = DedicatedFunctions.getStaticVariableValue(doc, VariableIdentifierIDs._variable_id_AcademicDegree.ToString());
            int academicIDInteger = 0;

            if (academicID == null || academicID == SettingValues.NotExist || string.IsNullOrEmpty(academicID))
            {
                string ad = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString());

                if (ad == null || ad == SettingValues.NotExist || string.IsNullOrEmpty(ad))
                {
                    return AcademicDegrees.Nothing;
                }
                else
                {
                    AcademicDegrees value = getAcademicDegreeID(ad);
                    academicIDInteger = (int)value;

                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableIdentifierIDs._variable_id_AcademicDegree.ToString(), ((int)value).ToString());
                }
            }
            else
            {
                academicIDInteger = int.Parse(academicID);
            }

            return (AcademicDegrees)academicIDInteger;
        }

        internal static AcademicDegrees getAcademicDegreeID(string academicDegreeName)
        {
            if (academicDegreeName == AcademicDegreeValues.AcademicDegree_AssociateOfScienceFa ||
                academicDegreeName == AcademicDegreeValues.AcademicDegree_AssociateOfScienceEn)
            {
                return AcademicDegrees.AssociateOfScience;
            }
            else if (academicDegreeName == AcademicDegreeValues.AcademicDegree_BachelorOfScienceFa ||
                academicDegreeName == AcademicDegreeValues.AcademicDegree_BachelorOfScienceEn)
            {
                return AcademicDegrees.BachelorOfScience;
            }
            else if (academicDegreeName == AcademicDegreeValues.AcademicDegree_MasterOfScienceFa ||
                academicDegreeName == AcademicDegreeValues.AcademicDegree_MasterOfScienceEn)
            {
                return AcademicDegrees.MasterOfScience;
            }
            else if (academicDegreeName == AcademicDegreeValues.AcademicDegree_DoctoralFa ||
                academicDegreeName == AcademicDegreeValues.AcademicDegree_DoctoralEn)
            {
                return AcademicDegrees.Doctoral;
            }
            else
            {
                return AcademicDegrees.Nothing;
            }
        }
        #endregion

        #region Universities
        internal static Universities getUniversity(Document doc)
        {
            string universityID = DedicatedFunctions.getStaticVariableValue(doc, VariableIdentifierIDs._variable_id_University.ToString());

            if (string.IsNullOrEmpty(universityID))
            {
                DedicatedFunctions.ShowErrorMessage("مقدار دانشگاه در داده های سند نادرست است", email: StringConstant.SupportEmail);
                throw new Exception("مقدار دانشگاه در داده های سند نادرست است");
            }

            if (universityID == null || universityID == SettingValues.NotExist)
                return Universities.Nothing;
            else
            {
                int universityIDInteger = int.Parse(universityID);
                return (Universities)universityIDInteger;
            }
        }
        #endregion

        #endregion

        #region delete section and page
        internal static void deleteSpecificPage(Document doc, int startNumberOfPage, int countToDelete = 1)
        {
            for (int i = 1; i <= countToDelete; i++)
            {
                doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, Name: startNumberOfPage).Select();
                doc.ActiveWindow.Selection.Bookmarks["\\Page"].Select();
                doc.ActiveWindow.Selection.Delete();
            }

        }
        internal static void deleteSpecificSection(Document doc, int startNumberOfSection, int countToDelete = 1)
        {
            for (int i = 1; i <= countToDelete; i++)
            {
                doc.Sections[startNumberOfSection].Range.Select();
                doc.ActiveWindow.Selection.Delete();
            }
        }

        #endregion

        #region exports


        internal static void exportIdentification(Document doc)
        {
            Universities university = DedicatedFunctions.getUniversity(doc);
            DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
            TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);
            PageIDs previousPageID = getLastSectionPageID(doc, university, templateType, documentType, PageIDs._page_Chapter1, WhichIndex.Previous);
            int firstSectionIndexPageChapter1 = DedicatedFunctions.getPageIDIndex(doc, previousPageID.ToString());
            int chaptersCount = DedicatedFunctions.getCurrentChaptersCount(doc);
            int sectionNumberLastPageChapter = DedicatedFunctions.getPageIDIndex(doc, InitialVariables.initialPageChapters.ToString() + chaptersCount);
            int lastSection = doc.Sections.Count;

            string titleOfDocument = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Title_Fa.ToString());
            if (string.IsNullOrEmpty(titleOfDocument))
                titleOfDocument = "عنوانی وارد نشده است";

            string Abstract = "";
            ContentControl[] abstractContentControl = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_Fa.ToString());
            if (abstractContentControl != null && abstractContentControl.Length != 0)
            {
                Range rangeAbstract = abstractContentControl[0].Range;
                if (rangeAbstract != null)
                    Abstract = rangeAbstract.Text;
            }

            if (string.IsNullOrEmpty(Abstract.Trim()))
                Abstract = "چکیده ای وارد نشده است";

            deleteSpecificSection(doc, sectionNumberLastPageChapter + 1, lastSection - (sectionNumberLastPageChapter));
            deleteSpecificSection(doc, 1, firstSectionIndexPageChapter1 - 1);

            DedicatedFunctions.removeAllCrossReferences(doc);
            foreach (Microsoft.Office.Interop.Word.Shape item in doc.Shapes)
            {
                item.Delete();
            }
            foreach (InlineShape item in doc.InlineShapes)
            {
                item.Delete();
            }
            foreach (Table item in doc.Tables)
            {
                item.Delete();
            }

            foreach (TableOfFigures item in doc.TablesOfFigures)
            {
                item.Delete();
            }
            foreach (TableOfContents item in doc.TablesOfContents)
            {
                item.Delete();
            }
            foreach (Index item in doc.Indexes)
            {
                item.Delete();
            }
            //foreach (Field item in doc.Fields)
            //{
            //    item.Delete();
            //}

            doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, Name: 1).Select();
            doc.ActiveWindow.Selection.HomeKey();
            doc.ActiveWindow.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage);
            doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, Name: 1).Select();

            string defaultPersianFont = doc.Styles[StyleNames.styleNormal].Font.NameBi;

            DedicatedFunctions.changeTextDirection(doc.ActiveWindow.Selection, TextDirection.RTL);

            DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection);
            DedicatedFunctions.changeParagraphAlignment(doc.ActiveWindow.Selection, WdParagraphAlignment.wdAlignParagraphCenter);
            DedicatedFunctions.insertText(doc, doc.ActiveWindow.Selection, titleOfDocument, defaultPersianFont, false, TextTypes.BOTH, 20, true);

            DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection, 2);
            DedicatedFunctions.changeParagraphAlignment(doc.ActiveWindow.Selection, WdParagraphAlignment.wdAlignParagraphRight);
            DedicatedFunctions.insertText(doc, doc.ActiveWindow.Selection, "چکیده", defaultPersianFont, false, TextTypes.BOTH, 12, false);
            DedicatedFunctions.insertText(doc, doc.ActiveWindow.Selection, Abstract, defaultPersianFont, false, TextTypes.BOTH, 15, false);
            DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection, 4);
            doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, Name: 1).Select();

        }

        #region Grayscale Contents
        internal static void ChangeColorIndex(Document doc, Range rng, WdColorIndex highlightColor)
        {
            doc.Application.Options.DefaultHighlightColorIndex = highlightColor;

            Selection selection = Globals.ThisAddIn.Application.Selection;

            rng.Find.ClearFormatting();
            rng.Find.Replacement.ClearFormatting();

            rng.Find.Highlight = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng.Find.Replacement.Highlight = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng.Find.Execute(Replace: WdReplace.wdReplaceAll, Wrap: WdFindWrap.wdFindContinue);

            rng.Find.Replacement.ClearFormatting();
            rng.Find.ClearFormatting();

            //Selection selection = Globals.ThisAddIn.Application.Selection;
            //
            //selection.Find.ClearFormatting();
            //selection.Find.Replacement.ClearFormatting();
            //
            //selection.Find.Highlight = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            //selection.Find.Replacement.Highlight = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            //selection.Find.Execute(Replace: WdReplace.wdReplaceAll , Wrap: WdFindWrap.wdFindContinue);
            //
            //selection.Find.Replacement.ClearFormatting();
            //selection.Find.ClearFormatting();
        }

        #region Content
        internal static void convertToGrayScaleContent(Document doc)
        {
            //set as Grayscale Contents
            doc.Content.Font.Color = WdColor.wdColorBlack;
            doc.Content.Font.TextColor.RGB = 0;
            doc.Content.Font.ColorIndex = WdColorIndex.wdBlack;
            doc.Content.Font.ColorIndexBi = WdColorIndex.wdBlack;
            doc.Content.Font.UnderlineColor = WdColor.wdColorBlack;
            doc.Content.Font.DiacriticColor = WdColor.wdColorBlack;
            ChangeColorIndex(doc, Globals.ThisAddIn.Application.Selection.Range, WdColorIndex.wdGray25);

            //doc.Content.Font.Borders;
            //doc.Content.Font.Fill;
            //doc.Content.Font.Line;
            //doc.Content.Font.TextShadow;

            //doc.Content.Shading
            //doc.Tables[1]
            //doc.Content.
        }
        internal static void convertToGrayScaleHeaderFooterContent(Document doc)
        {
            foreach (Range rng in doc.StoryRanges)
            {
                if (rng.StoryType == WdStoryType.wdPrimaryHeaderStory ||
                    rng.StoryType == WdStoryType.wdFirstPageHeaderStory ||
                    rng.StoryType == WdStoryType.wdEvenPagesHeaderStory ||
                    rng.StoryType == WdStoryType.wdPrimaryFooterStory ||
                    rng.StoryType == WdStoryType.wdFirstPageFooterStory ||
                    rng.StoryType == WdStoryType.wdEvenPagesFooterStory)
                {
                    //set as Grayscale Contents
                    rng.Font.Color = WdColor.wdColorBlack;
                    rng.Font.TextColor.RGB = 0;
                    rng.Font.ColorIndex = WdColorIndex.wdBlack;
                    rng.Font.ColorIndexBi = WdColorIndex.wdBlack;
                    rng.Font.UnderlineColor = WdColor.wdColorBlack;
                    rng.Font.DiacriticColor = WdColor.wdColorBlack;
                    ChangeColorIndex(doc, rng, WdColorIndex.wdGray25);

                    //doc.Content.Font.Borders;
                    //doc.Content.Font.Fill;
                    //doc.Content.Font.Line;
                    //doc.Content.Font.TextShadow;

                    //doc.Content.Shading
                    //doc.Tables[1]
                    //doc.Content.
                }
            }
        }
        internal static void convertToGrayScaleOtherContent(Document doc)
        {
            foreach (Range rng in doc.StoryRanges)
            {
                if (rng.StoryType == WdStoryType.wdTextFrameStory ||
                rng.StoryType == WdStoryType.wdCommentsStory ||
                rng.StoryType == WdStoryType.wdFootnotesStory ||
                rng.StoryType == WdStoryType.wdEndnotesStory)
                {
                    //set as Grayscale Contents
                    rng.Font.Color = WdColor.wdColorBlack;
                    rng.Font.TextColor.RGB = 0;
                    rng.Font.ColorIndex = WdColorIndex.wdBlack;
                    rng.Font.ColorIndexBi = WdColorIndex.wdBlack;
                    rng.Font.UnderlineColor = WdColor.wdColorBlack;
                    rng.Font.DiacriticColor = WdColor.wdColorBlack;
                    ChangeColorIndex(doc, rng, WdColorIndex.wdGray25);

                    //doc.Content.Font.Borders;
                    //doc.Content.Font.Fill;
                    //doc.Content.Font.Line;
                    //doc.Content.Font.TextShadow;

                    //doc.Content.Shading
                    //doc.Tables[1]
                    //doc.Content.
                }
            }
        }
        #endregion

        #region Table
        internal static void convertToGrayScaleTable(Document doc)
        {
            foreach (Table table in doc.Tables)
            {
                table.Borders.InsideColor = WdColor.wdColorBlack;
                table.Borders.InsideColorIndex = WdColorIndex.wdBlack;
                table.Borders.OutsideColor = WdColor.wdColorBlack;
                table.Borders.OutsideColorIndex = WdColorIndex.wdBlack;
            }
        }

        #endregion

        #region Shape
        internal static void convertToGrayScaleHeaderFooterShape(Document doc)
        {
            foreach (Range rng in doc.StoryRanges)
            {
                if (rng.StoryType == WdStoryType.wdPrimaryHeaderStory ||
                rng.StoryType == WdStoryType.wdFirstPageHeaderStory ||
                rng.StoryType == WdStoryType.wdEvenPagesHeaderStory ||
                rng.StoryType == WdStoryType.wdPrimaryFooterStory ||
                rng.StoryType == WdStoryType.wdFirstPageFooterStory ||
                rng.StoryType == WdStoryType.wdEvenPagesFooterStory)
                {
                    foreach (InlineShape iShape in rng.InlineShapes)
                    {
                        DedicatedFunctions.convertToGrayScaleShape(iShape);
                    }
                    if (rng.ShapeRange != null)
                    {
                        DedicatedFunctions.convertToGrayScaleShape(rng.ShapeRange);
                    }
                }
            }
        }
        internal static void convertToGrayScaleOtherShape(Document doc)
        {
            foreach (Range rng in doc.StoryRanges)
            {
                if (rng.StoryType == WdStoryType.wdTextFrameStory ||
                rng.StoryType == WdStoryType.wdCommentsStory ||
                rng.StoryType == WdStoryType.wdFootnotesStory ||
                rng.StoryType == WdStoryType.wdEndnotesStory)
                {
                    foreach (InlineShape iShape in rng.InlineShapes)
                    {
                        DedicatedFunctions.convertToGrayScaleShape(iShape);
                    }
                    if (rng.ShapeRange != null)
                    {
                        DedicatedFunctions.convertToGrayScaleShape(rng.ShapeRange);
                    }
                }
            }
        }

        //shape.BackgroundStyle == Microsoft.Office.Core.MsoBackgroundStyleIndex.msoBackgroundStylePreset1
        //shape.Fill.PictureEffects.Insert(Microsoft.Office.Core.MsoPictureEffectType.msoEffectSaturation).EffectParameters[1].Value = 0;
        //shape.Type == Microsoft.Office.Core.MsoShapeType.
        //
        //shape.Fill
        //
        //shape.GroupItems
        //shape.Nodes
        //
        //
        //shape.HasChart
        //
        //shape.Child ??

        static int previousCanvasShapeId = 0;
        static int previousGroupShapeId = 0;
        internal static void convertToGrayScaleShape(object shapeObject)
        {
            //if(shape.Diagram)
            //if(shape.DiagramNode)
            //if(shape.Nodes)

            if (shapeObject is InlineShape)//InlineShape
            {
                InlineShape shape = (InlineShape)shapeObject;
                //set GrayScale
                //shape.Fill.PictureEffects.Insert(Microsoft.Office.Core.MsoPictureEffectType.msoEffectSaturation).EffectParameters[1].Value = 0;
                //set as Grayscale Shapes
                try
                {
                    if (shape.Type == WdInlineShapeType.wdInlineShapePicture || shape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture || shape.Type == WdInlineShapeType.wdInlineShapePictureBullet)
                    {
                        shape.PictureFormat.ColorType = Microsoft.Office.Core.MsoPictureColorType.msoPictureGrayscale;
                    }
                    else if (shape.Type == WdInlineShapeType.wdInlineShapeChart)
                    {
                        shape.Chart.ChartColor = 5;//1,5,9 is grayscale , 5 recommended
                    }
                    else if (shape.Type == WdInlineShapeType.wdInlineShapeSmartArt)
                    {
                        shape.SmartArt.Color = Globals.ThisAddIn.Application.SmartArtColors[1];
                        foreach (Microsoft.Office.Core.SmartArtNode node in shape.SmartArt.AllNodes)
                        {
                            Microsoft.Office.Core.ShapeRange sRange = node.Shapes;
                            DedicatedFunctions.setColorsShapeAsGrayscale(sRange, DedicatedFunctions.GrayScaleType.Luminosity);
                        }
                    }
                    else
                    {
                        if (shape.Type != WdInlineShapeType.wdInlineShapeOLEControlObject && shape.Type != WdInlineShapeType.wdInlineShapeEmbeddedOLEObject && shape.Type != WdInlineShapeType.wdInlineShapeLinkedOLEObject && shape.Type != WdInlineShapeType.wdInlineShapeOWSAnchor && shape.Type != WdInlineShapeType.wdInlineShapeHorizontalLine && shape.Type != WdInlineShapeType.wdInlineShapePictureHorizontalLine && shape.Type != WdInlineShapeType.wdInlineShapeLinkedPictureHorizontalLine)
                        {// if not OLE And OWS and HorizontalLine, run.
                            DedicatedFunctions.setColorsShapeAsGrayscale(shape, DedicatedFunctions.GrayScaleType.Luminosity);
                        }

                    }
                }
                catch (Exception e)
                {
                }
            }
            else if (shapeObject is Shape)//Shape
            {
                Shape shape = (Shape)shapeObject;
                //set GrayScale
                //shape.Fill.PictureEffects.Insert(Microsoft.Office.Core.MsoPictureEffectType.msoEffectSaturation).EffectParameters[1].Value = 0;
                //set as Grayscale Shapes
                try
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture || shape.Type == Microsoft.Office.Core.MsoShapeType.msoLinkedPicture)
                    {
                        shape.PictureFormat.ColorType = Microsoft.Office.Core.MsoPictureColorType.msoPictureGrayscale;
                    }
                    else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                    {
                        foreach (Shape cShape in shape.CanvasItems)
                        {
                            if (cShape.ID != previousCanvasShapeId)
                            {
                                if (cShape.Type != Microsoft.Office.Core.MsoShapeType.msoWebVideo &&
                                    cShape.Type != Microsoft.Office.Core.MsoShapeType.msoCanvas)
                                {
                                    previousCanvasShapeId = cShape.ID;
                                    convertToGrayScaleShape(cShape);
                                }
                            }
                        }
                    }
                    else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        foreach (Shape gShape in shape.GroupItems)
                        {
                            if (gShape.ID != previousGroupShapeId)
                            {
                                //if(gShape.Type != Microsoft.Office.Core.MsoShapeType.msoGroup)
                                previousGroupShapeId = gShape.ID;
                                convertToGrayScaleShape(gShape);
                            }
                        }
                    }
                    else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoChart)
                    {
                        shape.Chart.ChartColor = 5;//1,5,9 is grayscale , 5 recommended
                    }
                    else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoSmartArt)
                    {
                        shape.SmartArt.Color = Globals.ThisAddIn.Application.SmartArtColors[1];
                        foreach (Microsoft.Office.Core.SmartArtNode node in shape.SmartArt.AllNodes)
                        {
                            Microsoft.Office.Core.ShapeRange sRange = node.Shapes;
                            DedicatedFunctions.setColorsShapeAsGrayscale(sRange, DedicatedFunctions.GrayScaleType.Luminosity);
                        }
                    }
                    else
                    {
                        if (shape.Type != Microsoft.Office.Core.MsoShapeType.msoOLEControlObject && shape.Type != Microsoft.Office.Core.MsoShapeType.msoEmbeddedOLEObject && shape.Type != Microsoft.Office.Core.MsoShapeType.msoLinkedOLEObject && shape.Type != Microsoft.Office.Core.MsoShapeType.msoMedia)
                        {// if not OLE and Media, run.
                            DedicatedFunctions.setColorsShapeAsGrayscale(shape, DedicatedFunctions.GrayScaleType.Luminosity);
                        }
                    }
                }
                catch (Exception e)
                {
                }
            }
            else if (shapeObject is ShapeRange)
            {
                ShapeRange sRange = (ShapeRange)shapeObject;
                DedicatedFunctions.setColorsShapeAsGrayscale(sRange, DedicatedFunctions.GrayScaleType.Luminosity);
            }
        }


        internal static void setColorsShapeAsGrayscale(Microsoft.Office.Core.ShapeRange shape, GrayScaleType grayScaleType = GrayScaleType.Luminosity)
        {
            int grayScaleValue = 0;
            System.Drawing.Color grayScale = new System.Drawing.Color();

            try//Fill Shape
            {
                if (shape.Fill?.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Fill.ForeColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Fill.ForeColor.RGB = grayScale.ToArgb();

                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Fill.BackColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Fill.BackColor.RGB = grayScale.ToArgb();

                    //Gradient Shape
                    if (shape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillGradient)
                    {
                        foreach (Microsoft.Office.Core.GradientStop gs in shape.Fill.GradientStops)
                        {
                            grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(gs.Color.RGB), grayScaleType);
                            grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                            gs.Color.RGB = grayScale.ToArgb();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Fill black and white a shape range Core");
                }
            }
            try//OutLine Shape
            {
                if (shape.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Line.ForeColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Line.ForeColor.RGB = grayScale.ToArgb();

                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Line.BackColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Line.BackColor.RGB = grayScale.ToArgb();
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Line black and white a shape range Core");
                }
            }
        }
        internal static void setColorsShapeAsGrayscale(ShapeRange shape, GrayScaleType grayScaleType = GrayScaleType.Luminosity)
        {
            int grayScaleValue = 0;
            System.Drawing.Color grayScale = new System.Drawing.Color();

            try//Fill Shape
            {
                if (shape.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Fill.ForeColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Fill.ForeColor.RGB = grayScale.ToArgb();

                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Fill.BackColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Fill.BackColor.RGB = grayScale.ToArgb();

                    //Gradient Shape
                    if (shape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillGradient)
                    {
                        foreach (Microsoft.Office.Core.GradientStop gs in shape.Fill.GradientStops)
                        {
                            grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(gs.Color.RGB), grayScaleType);
                            grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                            gs.Color.RGB = grayScale.ToArgb();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Fill black and white a shape range");
                }
            }
            try//OutLine Shape
            {
                if (shape.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Line.ForeColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Line.ForeColor.RGB = grayScale.ToArgb();

                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Line.BackColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Line.BackColor.RGB = grayScale.ToArgb();
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Line black and white a shape range");
                }
            }


            try//Texts
            {
                //if(shape.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                //{
                //}

                if (shape.TextFrame.HasText == (int)Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.TextFrame.TextRange.Font.TextColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.TextFrame.TextRange.Font.TextColor.RGB = grayScale.ToArgb();

                    convertToGrayScaleShape(shape.TextFrame.TextRange.InlineShapes);
                    if (shape.TextFrame.TextRange.ShapeRange != null)
                    {
                        convertToGrayScaleShape(shape.TextFrame.TextRange.ShapeRange);
                    }
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2147467259 => Error HRESULT E_FAIL has been returned from a call to a COM component.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2147467259 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Text black and white a shape range");
                }
            }


        }
        internal static void setColorsShapeAsGrayscale(Shape shape, GrayScaleType grayScaleType = GrayScaleType.Luminosity)
        {
            int grayScaleValue = 0;
            System.Drawing.Color grayScale = new System.Drawing.Color();

            try//Fill Shape
            {
                if (shape.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Fill.ForeColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Fill.ForeColor.RGB = grayScale.ToArgb();

                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Fill.BackColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Fill.BackColor.RGB = grayScale.ToArgb();

                    //Gradient Shape
                    if (shape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillGradient)
                    {
                        foreach (Microsoft.Office.Core.GradientStop gs in shape.Fill.GradientStops)
                        {
                            grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(gs.Color.RGB), grayScaleType);
                            grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                            gs.Color.RGB = grayScale.ToArgb();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Fill black and white a shape");
                }
            }

            //if(shape.Chart)
            //if(shape.Diagram)
            //if(shape.DiagramNode)
            //if(shape.Nodes)
            try//OutLine Shape
            {

                if (shape.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Line.ForeColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Line.ForeColor.RGB = grayScale.ToArgb();

                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.Line.BackColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.Line.BackColor.RGB = grayScale.ToArgb();
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Line black and white a shape");
                }
            }


            try//Texts
            {
                //if(shape.TextFrame2.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
                //{
                //}

                if (shape.TextFrame.HasText == (int)Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(shape.TextFrame.TextRange.Font.TextColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    shape.TextFrame.TextRange.Font.TextColor.RGB = grayScale.ToArgb();

                    convertToGrayScaleShape(shape.TextFrame.TextRange.InlineShapes);
                    if (shape.TextFrame.TextRange.ShapeRange != null)
                    {
                        convertToGrayScaleShape(shape.TextFrame.TextRange.ShapeRange);
                    }
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2147467259 => Error HRESULT E_FAIL has been returned from a call to a COM component.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2147467259 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Text black and white a shape");
                }
            }
        }
        internal static void setColorsShapeAsGrayscale(InlineShape iShape, GrayScaleType grayScaleType = GrayScaleType.Luminosity)
        {
            int grayScaleValue = 0;
            System.Drawing.Color grayScale = new System.Drawing.Color();

            //if(iShape.Borders)
            //if(iShape.Chart)
            //if(iShape.Field)
            //if(iShape.GroupItems)
            //if(iShape.HasSmartArt)
            //if(iShape.PictureFormat)
            //if(iShape.SmartArt)

            try//Fill Shape
            {
                if (iShape.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(iShape.Fill.ForeColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    iShape.Fill.ForeColor.RGB = grayScale.ToArgb();

                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(iShape.Fill.BackColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    iShape.Fill.BackColor.RGB = grayScale.ToArgb();

                    //Gradient Shape
                    if (iShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillGradient)
                    {
                        foreach (Microsoft.Office.Core.GradientStop gs in iShape.Fill.GradientStops)
                        {
                            grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(gs.Color.RGB), grayScaleType);
                            grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                            gs.Color.RGB = grayScale.ToArgb();
                        }
                    }
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Fill black and white a inline shape");
                }
            }

            try//OutLine Shape
            {
                if (iShape.Line.Visible == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(iShape.Line.ForeColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    iShape.Line.ForeColor.RGB = grayScale.ToArgb();

                    grayScaleValue = convertRGBToGrayscale(System.Drawing.Color.FromArgb(iShape.Line.BackColor.RGB), grayScaleType);
                    grayScale = System.Drawing.Color.FromArgb(0, grayScaleValue, grayScaleValue, grayScaleValue);
                    iShape.Line.BackColor.RGB = grayScale.ToArgb();
                }
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Line black and white a inline shape");
                }
            }

            try//OutLine Shape
            {
                iShape.Borders.InsideColor = WdColor.wdColorBlack;
                iShape.Borders.InsideColorIndex = WdColorIndex.wdBlack;
                iShape.Borders.OutsideColor = WdColor.wdColorBlack;
                iShape.Borders.OutsideColorIndex = WdColorIndex.wdBlack;
            }
            catch (Exception e)
            {
                //-2147024809 => Value does not fall within the expected range.
                //-2146827864 => Exception from HRESULT: 0x800A01A8
                if (e.HResult != -2147024809 && e.HResult != -2146827864)
                {
                    throw new Exception("Bug in Borders black and white a inline shape");
                }
            }

        }

        internal enum GrayScaleType
        {
            Lightness,//(max(R, G, B) + min(R, G, B)) / 2
            Average,//(R + G + B) / 3
            Luminosity//0.21 R + 0.72 G + 0.07 B
        }

        internal static int convertRGBToGrayscale(System.Drawing.Color color, GrayScaleType grayScaleType)
        {
            if (grayScaleType == GrayScaleType.Lightness)
                return (new[] { color.R, color.G, color.B }.Max() + new[] { color.R, color.G, color.B }.Min()) / 2;
            else if (grayScaleType == GrayScaleType.Average)
            {
                return (color.R + color.G + color.B) / 3;
            }
            else if (grayScaleType == GrayScaleType.Luminosity)
            {
                return (int)((color.R * 0.21) + (color.G * 0.72) + (color.B * 0.07));
            }
            else
                return 0;
        }

        #endregion

        #endregion

        #endregion

        #region Get Chapter Components Name
        internal static int getCurrentChaptersCount(Document doc)
        {
            int i = 0;
            PageIDs[] vpns =
            {
                PageIDs._page_Chapter1,
                PageIDs._page_Chapter2,
                PageIDs._page_Chapter3,
                PageIDs._page_Chapter4,
                PageIDs._page_Chapter5,
                PageIDs._page_Chapter6,
                PageIDs._page_Chapter7,
                PageIDs._page_Chapter8,
                PageIDs._page_Chapter9,
                PageIDs._page_Chapter10,
            };

            foreach (PageIDs vpn in vpns)
            {
                if (DedicatedFunctions.isStaticVariableExist(doc, vpn.ToString()))
                {
                    i++;
                }
            }
            return i;
        }

        #endregion

        #region get Bookmark Names, ContentControl Names
        internal static Dictionary<string, ContentControlNames> getRelationContentControlAndVariableFieldNames()
        {
            Dictionary<string, ContentControlNames> relations = new Dictionary<string, ContentControlNames>
            {
                { VariableFieldIDs._variable_field_Advisor_En.ToString() ,ContentControlNames._field_Advisor_En },
                { VariableFieldIDs._variable_field_Advisor_Fa.ToString() ,ContentControlNames._field_Advisor_Fa },
                { VariableFieldIDs._variable_field_Author_En.ToString(),ContentControlNames._field_Author_En },
                { VariableFieldIDs._variable_field_Author_Fa.ToString(),ContentControlNames._field_Author_Fa },
                { VariableFieldIDs._variable_field_DefenseDate_En.ToString(),ContentControlNames._field_DefenseDate_En },
                { VariableFieldIDs._variable_field_DefenseDate_Fa.ToString(),ContentControlNames._field_DefenseDate_Fa },
                { VariableFieldIDs._variable_field_Department_En.ToString(),ContentControlNames._field_Department_En },
                { VariableFieldIDs._variable_field_Department_Fa.ToString(),ContentControlNames._field_Department_Fa },
                { VariableFieldIDs._variable_field_FieldOfStudy_En.ToString(),ContentControlNames._field_FieldOfStudy_En },
                { VariableFieldIDs._variable_field_FieldOfStudy_Fa.ToString(),ContentControlNames._field_FieldOfStudy_Fa },
                { VariableFieldIDs._variable_field_AreaOfStudy_En.ToString(),ContentControlNames._field_AreaOfStudy_En },
                { VariableFieldIDs._variable_field_AreaOfStudy_Fa.ToString(),ContentControlNames._field_AreaOfStudy_Fa },
                { VariableFieldIDs._variable_field_Group_En.ToString(),ContentControlNames._field_Group_En },
                { VariableFieldIDs._variable_field_Group_Fa.ToString(),ContentControlNames._field_Group_Fa },
                { VariableFieldIDs._variable_field_AcademicDegree_En.ToString(),ContentControlNames._field_AcademicDegree_En },
                { VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString(),ContentControlNames._field_AcademicDegree_Fa },
                { VariableFieldIDs._variable_field_NameOfCourse_Fa.ToString(),ContentControlNames._field_NameOfCourse_Fa },
                { VariableFieldIDs._variable_field_Supervisor_En.ToString(),ContentControlNames._field_Supervisor_En },
                { VariableFieldIDs._variable_field_Supervisor_Fa.ToString(),ContentControlNames._field_Supervisor_Fa },
                { VariableFieldIDs._variable_field_Title_En.ToString(),ContentControlNames._field_Title_En },
                { VariableFieldIDs._variable_field_Title_Fa.ToString(),ContentControlNames._field_Title_Fa },
                { VariableFieldIDs._variable_field_University_En.ToString(),ContentControlNames._field_University_En },
                { VariableFieldIDs._variable_field_University_Fa.ToString(),ContentControlNames._field_University_Fa },
                { VariableFieldIDs._variable_field_Branch_En.ToString(),ContentControlNames._field_Branch_En },
                { VariableFieldIDs._variable_field_Branch_Fa.ToString(),ContentControlNames._field_Branch_Fa },
            };
            return relations;
        }
        #endregion

        #region Page IDs
        //internal static void setPageIDIndexes(Document doc , string _pageID , List<int> values)
        //{
        //	values.Sort();
        //	string sectionNumber = "";
        //
        //	foreach(int value in values)
        //	{
        //		sectionNumber = sectionNumber + value.ToString() + ",";
        //	}
        //	sectionNumber = sectionNumber.TrimEnd(',');
        //
        //	DedicatedFunctions.setORAddStaticVariableValue(doc , _pageID , sectionNumber);
        //}

        internal static void removeUnCorrectContentControlInPageID(Document doc, string pageID)
        {
            string contentControlID = DedicatedFunctions.getStaticVariableValue(doc, pageID);

            ContentControls contentControls = doc.SelectContentControlsByTag(pageID);

            //check ContentControlWithIDExist
            ContentControl onContentControlWithIDExist = null;

            //get all content control created by user  and tag is Page ID,so should be removed
            List<ContentControl> wrongContentControls = new List<ContentControl>();

            if (contentControls != null)
            {
                //check list onContentControlWithIDExist and lastContentControl
                for (int i = 1; i <= contentControls.Count; i++)
                {
                    if (contentControls[i].ID == contentControlID)
                    {
                        onContentControlWithIDExist = contentControls[i];
                    }
                    else if (i == contentControls.Count)//is last ContentControl
                    {
                        if (onContentControlWithIDExist != null)
                            wrongContentControls.Add(contentControls[i]);
                        //else
                        // remain last ContentControl
                    }
                    else
                    {
                        //also delete ContentControls in LastSection
                        wrongContentControls.Add(contentControls[i]);
                    }
                }
            }

            //delete content control created by user and tag is Page ID
            for (int i = wrongContentControls.Count - 1; i >= 0; i--)
            {
                wrongContentControls[i].LockContentControl = false;
                wrongContentControls[i].LockContents = false;
                wrongContentControls[i].Delete(true);
            }
        }
        internal static int getPageIDIndex(Document doc, string pageID)
        {
            removeUnCorrectContentControlInPageID(doc, pageID);

            string contentControlID = DedicatedFunctions.getStaticVariableValue(doc, pageID);
            if (contentControlID == "LastPage")
            {
                return doc.Sections[doc.Sections.Count].Index;
            }
            else
            {
                ContentControls contentControls = doc.SelectContentControlsByTag(pageID);

                if (contentControls != null && contentControls.Count == 1)
                    return contentControls[1].Range.Sections[1].Index;
                else
                {
                    DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره برای گرفتن آخرین index از صفحه رخ داد", email: StringConstant.SupportEmail);
                    throw new Exception("ContentControl relationed by> " + pageID + " not exist");
                }
            }
        }

        #endregion

        #region Section
        internal enum WhichIndex
        {
            Next,
            Previous,
        }
        internal static PageIDs getLastSectionPageID(Document doc, Universities university, TemplateTypes templateTypes, DocumentTypes documentType, PageIDs pageID, WhichIndex whichIndex)
        {
            int orderCurrentPageID = 0;
            PageIDs specifiedPageID = pageID;

            //get available Variable Pages
            Dictionary<PageIDs, int> orders = new Dictionary<PageIDs, int> { };
            foreach (var pageIDOrder in TemplateAccess.getTemplateRelationshipModelList(university, templateTypes, documentType, true))
            {
                if (pageIDOrder.PageID == pageID)
                    orderCurrentPageID = pageIDOrder.Order;
                else if (DedicatedFunctions.isStaticVariableExist(doc, pageIDOrder.PageID.ToString()))
                    orders.Add(pageIDOrder.PageID, pageIDOrder.Order);
            }

            //sort 
            var sortedDict = (from entry in orders orderby entry.Value ascending select entry).ToList();

            //get Previous or Next PageID
            int lastIndex = sortedDict.Count - 1;
            for (int i = 0; i < sortedDict.Count; i++)
            {
                if (sortedDict[i].Value > orderCurrentPageID)
                {
                    if (whichIndex == WhichIndex.Next)
                    {
                        specifiedPageID = sortedDict[i].Key;
                    }
                    else if (whichIndex == WhichIndex.Previous)
                    {
                        specifiedPageID = sortedDict[i - 1].Key;
                    }
                    break;
                }
                else if (i == lastIndex)
                {
                    if (sortedDict[lastIndex].Value < orderCurrentPageID)
                        specifiedPageID = sortedDict[lastIndex].Key;
                    break;
                }
            }

            //get section index and PageID Specified Variable
            return specifiedPageID;
        }

        internal static void deleteSection(Document doc, Section section)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("حذف بخش " + section.Index);
            int lastSectionIndex = doc.Sections.Count;
            if (lastSectionIndex == section.Index)
            {
                foreach (HeaderFooter hf in section.Headers)
                {
                    hf.LinkToPrevious = false;
                    hf.Range.Delete();
                }
                foreach (HeaderFooter hf in section.Footers)
                {
                    hf.LinkToPrevious = false;
                    hf.Range.Delete();
                }
            }

            DedicatedFunctions.unProtectingContentControlsSection(section);

            if (doc.Sections.Count != section.Index)
            {
                section.Range.Delete();
            }
            else//Last Section
            {
                Range rng = doc.Sections[doc.Sections.Count].Range;
                rng.Select();
                rng.MoveStart(WdUnits.wdCharacter, -1);
                rng.Delete();
            }
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }
        #endregion

        #region Insert Page
        internal static void getLockStatusContentControl(string ccTag, out bool lockContentControl, out bool lockContents)
        {
            //bool[0]> lockContentControl, bool[1]> lockContents
            Dictionary<string, bool[]> lockStates = new Dictionary<string, bool[]>
            {
                { ContentControlNames._field_Advisor_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Advisor_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Author_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Author_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_DefenseDate_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_DefenseDate_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Department_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Department_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_FieldOfStudy_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_FieldOfStudy_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_AreaOfStudy_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_AreaOfStudy_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Group_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Group_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Icon_University.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_AcademicDegree_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_AcademicDegree_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_InTheNameOfAllah.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_NameOfCourse_Fa.ToString() , new bool[] { true , true } },

                { ContentControlNames._field_Advisor_Title_En.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Advisor_Title_Fa.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_AreaOfStudy_Title_Fa.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_AreaOfStudy_Title_En.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Author_Title_En.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Author_Title_Fa.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Supervisor_Title_Fa.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Supervisor_Title_En.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Title_Title_Fa.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Title_Title_En.ToString() , new bool[] { true , false } },

                { ContentControlNames._field_Supervisor_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Supervisor_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Title_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Title_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Type_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Type_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_University_En.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_University_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Branch_Fa.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Branch_En.ToString() , new bool[] { true , true } },

                { ContentControlNames._field_Chapter1_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter2_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter3_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter4_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter5_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter6_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter7_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter8_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter9_Title.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Chapter10_Title.ToString() , new bool[] { true , false } },

                { ContentControlNames._field_Hidden_Chapter1_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter2_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter3_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter4_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter5_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter6_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter7_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter8_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter9_Title.ToString() , new bool[] { true , true } },
                { ContentControlNames._field_Hidden_Chapter10_Title.ToString() , new bool[] { true , true } },

                { ContentControlNames._field_Abstract_En.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Abstract_Fa.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Keywords_En.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Keywords_Fa.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Dedication_Fa.ToString() , new bool[] { true , false } },
                { ContentControlNames._field_Acknowledgment_Fa.ToString() , new bool[] { true , false } },

                { ContentControlNames._field_DefenseLocation_Fa.ToString() , new bool[] { false , false } },
                { ContentControlNames._field_Examiner_Fa.ToString() , new bool[] { false , false } },
                { ContentControlNames._field_Examiner_Title_Fa.ToString() , new bool[] { false , false } },
            };

            if (lockStates.ContainsKey(ccTag))
            {
                lockStates.TryGetValue(ccTag, out var value);
                lockContentControl = value[0];
                lockContents = value[1];
                return;
            }
            else
            {
                lockContentControl = false;
                lockContents = false;
            }
        }
        internal static string getContentControlTitle(string tag)
        {
            if (ContentControlNames._field_Type_Fa.ToString() == tag)
            {
                return "نوع سند";
            }
            if (ContentControlNames._field_Type_En.ToString() == tag)
            { return "Document Type"; }
            else if (ContentControlNames._field_Icon_University.ToString() == tag)
            { return "نماد دانشگاه"; }
            else if (ContentControlNames._field_University_Fa.ToString() == tag)
            { return "دانشگاه"; }
            else if (ContentControlNames._field_Branch_Fa.ToString() == tag)
            { return "واحد"; }
            else if (ContentControlNames._field_Department_Fa.ToString() == tag)
            { return "دانشکده"; }
            else if (ContentControlNames._field_Group_Fa.ToString() == tag)
            { return "گروه"; }
            else if (ContentControlNames._field_FieldOfStudy_Fa.ToString() == tag)
            { return "رشته تحصیلی"; }
            else if (ContentControlNames._field_AreaOfStudy_Fa.ToString() == tag)
            { return "گرایش تحصیلی"; }
            else if (ContentControlNames._field_Title_Fa.ToString() == tag)
            { return "عنوان"; }
            else if (ContentControlNames._field_Supervisor_Fa.ToString() == tag)
            { return "استاد راهنما"; }
            else if (ContentControlNames._field_Advisor_Fa.ToString() == tag)
            { return "استاد مشاور"; }
            else if (ContentControlNames._field_Author_Fa.ToString() == tag)
            { return "نگارنده"; }
            else if (ContentControlNames._field_DefenseDate_Fa.ToString() == tag)
            { return "تاریخ دفاع"; }
            else if (ContentControlNames._field_AcademicDegree_Fa.ToString() == tag)
            { return "مقطع تحصیلی"; }
            else if (ContentControlNames._field_NameOfCourse_Fa.ToString() == tag)
            { return "نام درس"; }
            else if (ContentControlNames._field_InTheNameOfAllah.ToString() == tag)
            { return "به نام خدا"; }
            else if (ContentControlNames._field_Abstract_Fa.ToString() == tag)
            { return "چکیده"; }
            else if (ContentControlNames._field_Dedication_Fa.ToString() == tag)
            { return "تقدیم"; }
            else if (ContentControlNames._field_Acknowledgment_Fa.ToString() == tag)
            { return "سپاس و قدردانی"; }
            else if (ContentControlNames._field_Keywords_Fa.ToString() == tag)
            { return "کلید‌‌واژه ها"; }
            else if (ContentControlNames._field_Abstract_En.ToString() == tag)
            { return "Abstract"; }
            else if (ContentControlNames._field_Keywords_En.ToString() == tag)
            { return "Keywords"; }
            else if (ContentControlNames._field_University_En.ToString() == tag)
            { return "University"; }
            else if (ContentControlNames._field_Branch_En.ToString() == tag)
            { return "Branch"; }
            else if (ContentControlNames._field_Department_En.ToString() == tag)
            { return "Department"; }
            else if (ContentControlNames._field_Group_En.ToString() == tag)
            { return "Group"; }
            else if (ContentControlNames._field_FieldOfStudy_En.ToString() == tag)
            { return "Field of Study"; }
            else if (ContentControlNames._field_AreaOfStudy_En.ToString() == tag)
            { return "Area of Study"; }
            else if (ContentControlNames._field_Title_En.ToString() == tag)
            { return "Title"; }
            else if (ContentControlNames._field_Supervisor_En.ToString() == tag)
            { return "Supervisor"; }
            else if (ContentControlNames._field_Advisor_En.ToString() == tag)
            { return "Advisor"; }
            else if (ContentControlNames._field_Author_En.ToString() == tag)
            { return "Author"; }
            else if (ContentControlNames._field_DefenseDate_En.ToString() == tag)
            { return "Defense Date"; }
            else if (ContentControlNames._field_AcademicDegree_En.ToString() == tag)
            { return "Academic AcademicDegree"; }

            else if (ContentControlNames._field_Chapter1_Title.ToString() == tag)
            { return "فصل اول"; }
            else if (ContentControlNames._field_Chapter2_Title.ToString() == tag)
            { return "فصل دوم"; }
            else if (ContentControlNames._field_Chapter3_Title.ToString() == tag)
            { return "فصل سوم"; }
            else if (ContentControlNames._field_Chapter4_Title.ToString() == tag)
            { return "فصل چهارم"; }
            else if (ContentControlNames._field_Chapter5_Title.ToString() == tag)
            { return "فصل پنجم"; }
            else if (ContentControlNames._field_Chapter6_Title.ToString() == tag)
            { return "فصل ششم"; }
            else if (ContentControlNames._field_Chapter7_Title.ToString() == tag)
            { return "فصل هفتم"; }
            else if (ContentControlNames._field_Chapter8_Title.ToString() == tag)
            { return "فصل هشتم"; }
            else if (ContentControlNames._field_Chapter9_Title.ToString() == tag)
            { return "فصل نهم"; }
            else if (ContentControlNames._field_Chapter10_Title.ToString() == tag)
            { return "فصل دهم"; }
            else
            {
                return "";
            }
        }

        internal static void setProperties(Document doc, Section section, int differentSections, DocumentFormat.OpenXml.CustomProperties.Properties properties, DocumentFormat.OpenXml.Wordprocessing.Body body)
        {
            if (body != null)
            {
                if (differentSections != 0)
                {
                    int i = differentSections;
                    foreach (var paragraph in body.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>())
                    {
                        var sectionProperties = paragraph.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties>()?.SectionProperties;

                        if (sectionProperties != null)
                        {
                            if (differentSections > 0)
                            {
                                setSectionProperties(doc.Sections[section.Index - i], sectionProperties);
                                i--;
                            }
                            else
                            {
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره از اسناد الگو، اسناد الگو Section هاش درست مشخص نیست");
                                throw new Exception("خطای غیر منتظره از اسناد الگو، اسناد الگو Section هاش درست مشخص نیست");
                            }
                        }
                    }
                }
                foreach (var sectionProperties in body.Elements<DocumentFormat.OpenXml.Wordprocessing.SectionProperties>())
                {
                    if (sectionProperties != null)
                        setSectionProperties(section, sectionProperties);
                }
            }

            if (properties != null)
            {
                foreach (DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty property in properties)
                {
                    if (property.Name.Value.Equals(Constants.CustomProperties.Section_Break_EvenPage.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        doc.ActiveWindow.Selection.InsertBreak(WdBreakType.wdSectionBreakEvenPage);
                    }
                    else if (property.Name.Value.Equals(Constants.CustomProperties.Section_Break_OddPage.ToString(), StringComparison.OrdinalIgnoreCase))
                    {
                        doc.ActiveWindow.Selection.InsertBreak(WdBreakType.wdSectionBreakOddPage);
                    }
                }
            }
        }
        private static void setSectionProperties(Section section, DocumentFormat.OpenXml.Wordprocessing.SectionProperties sectionProperties)
        {
            var differnetFirstPage = sectionProperties.Descendants<DocumentFormat.OpenXml.Wordprocessing.TitlePage>().FirstOrDefault();
            if (differnetFirstPage == null)
                section.PageSetup.DifferentFirstPageHeaderFooter = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
            else
                section.PageSetup.DifferentFirstPageHeaderFooter = (int)Microsoft.Office.Core.MsoTriState.msoTrue;

            //set PageNumbers
            //یکبار شماره صفحه در سکشن تنظیم بشه کافیه، در کل section اعمال میشود
            HeaderFooter firstHeaderFooter = null;
            foreach (HeaderFooter headerFooter in section.Headers)
            {
                //if(headerFooter.PageNumbers.Count >= 1)
                if (headerFooter.Exists)
                {
                    firstHeaderFooter = headerFooter;
                    break;
                }
            }

            if (firstHeaderFooter == null)
            {
                foreach (HeaderFooter headerFooter in section.Footers)
                {
                    //if(headerFooter.PageNumbers.Count >= 1)
                    if (headerFooter.Exists)
                    {
                        firstHeaderFooter = headerFooter;
                        break;
                    }
                }
            }

            var pageNumberType = sectionProperties.Descendants<DocumentFormat.OpenXml.Wordprocessing.PageNumberType>().FirstOrDefault();
            if (pageNumberType != null)
            {
                if (pageNumberType.Format == null)
                    firstHeaderFooter.PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic;
                else if (pageNumberType.Format?.Value == DocumentFormat.OpenXml.Wordprocessing.NumberFormatValues.ArabicAlpha)
                    firstHeaderFooter.PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabicLetter1;
                else if (pageNumberType.Format?.Value == DocumentFormat.OpenXml.Wordprocessing.NumberFormatValues.ArabicAbjad)
                    firstHeaderFooter.PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabicLetter2;
                else if (pageNumberType.Format?.Value == DocumentFormat.OpenXml.Wordprocessing.NumberFormatValues.UpperRoman)
                    firstHeaderFooter.PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleUppercaseRoman;
                else if (pageNumberType.Format?.Value == DocumentFormat.OpenXml.Wordprocessing.NumberFormatValues.LowerRoman)
                    firstHeaderFooter.PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman;
                else if (pageNumberType.Format?.Value == DocumentFormat.OpenXml.Wordprocessing.NumberFormatValues.UpperLetter)
                    firstHeaderFooter.PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleUppercaseLetter;
                else if (pageNumberType.Format?.Value == DocumentFormat.OpenXml.Wordprocessing.NumberFormatValues.LowerLetter)
                    firstHeaderFooter.PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseLetter;

                if (pageNumberType.Start == null)
                    firstHeaderFooter.PageNumbers.RestartNumberingAtSection = false;
                else
                {
                    firstHeaderFooter.PageNumbers.RestartNumberingAtSection = true;
                    firstHeaderFooter.PageNumbers.StartingNumber = pageNumberType.Start.Value;
                }

                if (pageNumberType.ChapterStyle != null)
                    firstHeaderFooter.PageNumbers.HeadingLevelForChapter = pageNumberType.ChapterStyle;

                if (pageNumberType.ChapterSeparator != null)
                {
                    WdSeparatorType wdSeparatorType = WdSeparatorType.wdSeparatorColon;
                    if (pageNumberType.ChapterSeparator.Value == DocumentFormat.OpenXml.Wordprocessing.ChapterSeparatorValues.Colon)
                        wdSeparatorType = WdSeparatorType.wdSeparatorColon;
                    else if (pageNumberType.ChapterSeparator.Value == DocumentFormat.OpenXml.Wordprocessing.ChapterSeparatorValues.Hyphen)
                        wdSeparatorType = WdSeparatorType.wdSeparatorHyphen;
                    else if (pageNumberType.ChapterSeparator.Value == DocumentFormat.OpenXml.Wordprocessing.ChapterSeparatorValues.EmDash)
                        wdSeparatorType = WdSeparatorType.wdSeparatorEmDash;
                    else if (pageNumberType.ChapterSeparator.Value == DocumentFormat.OpenXml.Wordprocessing.ChapterSeparatorValues.EnDash)
                        wdSeparatorType = WdSeparatorType.wdSeparatorEnDash;
                    else if (pageNumberType.ChapterSeparator.Value == DocumentFormat.OpenXml.Wordprocessing.ChapterSeparatorValues.Period)
                        wdSeparatorType = WdSeparatorType.wdSeparatorPeriod;

                    firstHeaderFooter.PageNumbers.ChapterPageSeparator = wdSeparatorType;
                }

            }
            else
            {
                firstHeaderFooter.PageNumbers.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic;
                firstHeaderFooter.PageNumbers.RestartNumberingAtSection = false;
            }
        }


        internal static void resetHeaderFooter(Microsoft.Office.Interop.Word.Section section, bool afterInsertFile)
        {

            //headerFooter.Exists = false;
            if (afterInsertFile)
            {
                foreach (HeaderFooter headerFooter in section.Headers)
                {
                    if (headerFooter.Range.Text.Contains(SettingValues.NullInHeaderFooterTemplates))
                    {
                        headerFooter.LinkToPrevious = false;
                        headerFooter.Range.Delete();
                    }
                }
                foreach (HeaderFooter headerFooter in section.Footers)
                {
                    if (headerFooter.Range.Text.Contains(SettingValues.NullInHeaderFooterTemplates))
                    {
                        headerFooter.LinkToPrevious = false;
                        headerFooter.Range.Delete();
                    }
                }
            }
            else
            {
                foreach (HeaderFooter headerFooter in section.Headers)
                {
                    try
                    {
                        headerFooter.LinkToPrevious = false;
                        headerFooter.Range.Delete();
                    }
                    catch (Exception e)
                    {
                        //-2146823683 > System.Runtime.InteropServices.COMException: 'This method or property is not available because there is no previous section.'
                        if (e.HResult != -2146823683)
                            throw;
                    }

                }
                foreach (HeaderFooter headerFooter in section.Footers)
                {
                    try
                    {
                        headerFooter.LinkToPrevious = false;
                        headerFooter.Range.Delete();
                    }
                    catch (Exception e)
                    {
                        //-2146823683 > System.Runtime.InteropServices.COMException: 'This method or property is not available because there is no previous section.'
                        if (e.HResult != -2146823683)
                            throw;
                    }
                }
            }
        }


        #endregion

        #region Custom Styles
        internal static void createSectionBreakStyle(Document doc)
        {
            Style style;
            try
            {
                style = doc.Styles.Add(StyleNames.CustomSectionBreakStyle, WdStyleType.wdStyleTypeLinked);
            }
            catch (Exception)
            {
                style = doc.Styles[StyleNames.CustomSectionBreakStyle];
            }

            style.set_BaseStyle("");
            style.set_NextParagraphStyle(doc.Styles[StyleNames.styleNormal]);
            style.AutomaticallyUpdate = true;
            style.Font.Size = 8;
            style.Font.SizeBi = 8;
            //style.NoSpaceBetweenParagraphsOfSameStyle
            //style.ParagraphFormat
            style.Priority = 20;
            style.QuickStyle = false;
            style.NoSpaceBetweenParagraphsOfSameStyle = false;

            style.ParagraphFormat.SpaceBeforeAuto = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
            style.ParagraphFormat.SpaceAfterAuto = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
            style.ParagraphFormat.SpaceBefore = 0;
            style.ParagraphFormat.SpaceAfter = 0;
            style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly;
            style.ParagraphFormat.LineSpacing = 8;
            style.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            style.ParagraphFormat.WidowControl = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            style.ParagraphFormat.KeepWithNext = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            style.ParagraphFormat.KeepTogether = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            style.ParagraphFormat.NoLineNumber = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            style.ParagraphFormat.Hyphenation = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            style.ParagraphFormat.LineUnitBefore = 0;
            style.ParagraphFormat.LineUnitAfter = 0;
            style.ParagraphFormat.CollapsedByDefault = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
        }

        #endregion

        #region for Create and AddRemove Template
        internal static WdBreakType getBreakTypeFromCustomDocumentProperties(DocumentFormat.OpenXml.CustomProperties.Properties properties)
        {
            if (properties != null)
            {
                foreach (DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty property in properties)
                {
                    if (property.Name.Value.Equals(Constants.CustomProperties.Section_Break_EvenPage.ToString(), StringComparison.OrdinalIgnoreCase))
                        return WdBreakType.wdSectionBreakEvenPage;
                    else if (property.Name.Value.Equals(Constants.CustomProperties.Section_Break_OddPage.ToString(), StringComparison.OrdinalIgnoreCase))
                        return WdBreakType.wdSectionBreakOddPage;
                }
            }
            return WdBreakType.wdSectionBreakNextPage;
        }

        internal static void setSectionStartFromBreakType(Section section, WdBreakType wdBreakType)
        {
            if (wdBreakType == WdBreakType.wdSectionBreakNextPage)
            {
                section.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage;
            }
            else if (wdBreakType == WdBreakType.wdSectionBreakContinuous)
            {
                section.PageSetup.SectionStart = WdSectionStart.wdSectionContinuous;
            }
            else if (wdBreakType == WdBreakType.wdSectionBreakOddPage)
            {
                section.PageSetup.SectionStart = WdSectionStart.wdSectionOddPage;
            }
            else if (wdBreakType == WdBreakType.wdSectionBreakEvenPage)
            {
                section.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage;
            }
        }
        internal static void insertSectionBreak(Document doc, DocumentFormat.OpenXml.CustomProperties.Properties properties)
        {
            if (doc.ActiveWindow.Selection.ContentControls.Count > 0)
                doc.ActiveWindow.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);

            if (doc.ActiveWindow.Selection.Paragraphs.Count == 1 && doc.ActiveWindow.Selection.Paragraphs[1].Range.Text.Trim('\r').Count() != 0)
            {
                doc.ActiveWindow.Selection.InsertParagraphAfter();
                doc.ActiveWindow.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
            }

            doc.ActiveWindow.Selection.InsertParagraphAfter();
            doc.ActiveWindow.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);

            doc.ActiveWindow.Selection.InsertBreak(getBreakTypeFromCustomDocumentProperties(properties));
        }
        internal static string protectSectionBreak(Document doc, string pageID)
        {
            //select section break, and set style CustomSectionBreakStyle
            doc.ActiveWindow.Selection.HomeKey(WdUnits.wdLine);
            doc.ActiveWindow.Selection.EndKey(WdUnits.wdLine, Extend: true);
            doc.ActiveWindow.Selection.set_Style(StyleNames.CustomSectionBreakStyle);

            ////create Protector for created SectionBreak
            doc.ActiveWindow.Selection.ContentControls.Add(WdContentControlType.wdContentControlRichText);
            doc.ActiveWindow.Selection.ParentContentControl.Tag = pageID;
            doc.ActiveWindow.Selection.ParentContentControl.Appearance = WdContentControlAppearance.wdContentControlHidden;
            doc.ActiveWindow.Selection.ParentContentControl.Color = WdColor.wdColorDarkRed;
            doc.ActiveWindow.Selection.ParentContentControl.Title = "Section Break (Locked)";
            doc.ActiveWindow.Selection.ParentContentControl.LockContentControl = true;
            doc.ActiveWindow.Selection.ParentContentControl.LockContents = true;
            doc.ActiveWindow.Selection.ParentContentControl.set_DefaultTextStyle(StyleNames.CustomSectionBreakStyle);
            string contentControlID = doc.ActiveWindow.Selection.ParentContentControl.ID;

            return contentControlID;
        }

        #endregion

        #region Virastar

        /// <summary>
        /// 
        /// </summary>
        /// <param name="frm"></param>
        /// <param name="rng"></param>
        /// <param name="listFind"></param>
        /// <param name="listReplace"></param>
        /// <param name="listWildcard"></param>
        /// <param name="matchCase">اگر فعال شود به حروف کوچک و بزرگ حساس میشود</param>
        /// <param name="matchWholeWord">اگر فعال باشد تنها همین کلمه را پیدا میکند، مثلا کلمه جدی رو در جدید پیدا نمیکند</param>
        /// <param name="matchKashida"></param>
        /// <param name="matchDiacritics"></param>
        /// <param name="matchAlefHamza"></param>
        internal static void StringCorrection(LoadingForm frm, Range rng, List<SearchReplaceModel> correctionModels)
        {
            for (int i = correctionModels.Count - 1; i >= 0; i--)
            {
                if (correctionModels[i].Search == "")
                {
                    correctionModels.RemoveAt(i);
                }
            }

            Thread.Sleep(500);
            int countOfChanges = 0;

            try
            {
                int startRange = rng.Start;
                int endRange = rng.End;
                rng.Find.ClearFormatting();
                rng.Find.Replacement.ClearFormatting();

                int previousEndRange = 0;
                foreach (SearchReplaceModel cm in correctionModels)
                {
                    rng.SetRange(startRange, endRange);

                    while (rng.Find.Execute(FindText: cm.Search,
                                    ReplaceWith: cm.Replace,
                                    Replace: WdReplace.wdReplaceOne,
                                    MatchWildcards: cm.Wildcard == "1" ? true : false,
                                    MatchCase: cm.MatchCase,
                                    MatchWholeWord: cm.MatchWholeWord,
                                    MatchKashida: cm.MatchKashida,
                                    MatchDiacritics: cm.MatchDiacritics,
                                    MatchAlefHamza: cm.MatchAlefHamza,
                                    Wrap: WdFindWrap.wdFindStop))
                    {
                        if (!rng.Find.Found)
                            break;

                        if (rng.End == previousEndRange)
                            break;
                        else
                            previousEndRange = rng.End;

                        Debug.WriteLine("Text>" + rng.Text + "\n\nrng.Start>" + rng.Start + " previousRange.End>" + rng.End + " endRange>" + endRange);
                        if (cm.CounterUp != "0")
                            countOfChanges++;

                        rng.SetRange(rng.End, endRange);
                    }
                }
                rng.Find.ClearFormatting();
                rng.Find.Replacement.ClearFormatting();
            }
            catch (System.Exception e)
            {
                DedicatedFunctions.ShowErrorMessage("خطایی در اعمال جایگذاری در کل سند به وجود آمده" + "\nپیغام خطا:\n" + e.Message);
            }
            frm?.closeForm(countOfChanges);
        }
        internal static void StringCorrection(LoadingForm frm, Microsoft.Office.Interop.Word.Document doc, List<SearchReplaceModel> correctionModels)
        {
            for (int i = correctionModels.Count - 1; i >= 0; i--)
            {
                if (correctionModels[i].Search == "")
                {
                    correctionModels.RemoveAt(i);
                }
            }
            Thread.Sleep(500);
            System.Collections.Generic.List<Range> storyRanges = new System.Collections.Generic.List<Range>();
            //storyRanges.Add(doc.StoryRanges[WdStoryType.wdCommentsStory]);

            try
            {
                storyRanges.Add(doc.StoryRanges[WdStoryType.wdMainTextStory]);
            }
            catch (System.Exception) { }

            if (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionFootnote)
            {

                try
                {
                    storyRanges.Add(doc.StoryRanges[WdStoryType.wdEndnotesStory]);
                }
                catch (System.Exception)
                { }
                try
                {
                    storyRanges.Add(doc.StoryRanges[WdStoryType.wdFootnotesStory]);
                }
                catch (System.Exception)
                { }

            }
            if (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionHeadersFooters)
            {
                //try
                //{
                //    storyRanges.Add(doc.StoryRanges[WdStoryType.wdPrimaryHeaderStory]);
                //}
                //catch (System.Exception) { }
                //try
                //{
                //    storyRanges.Add(doc.StoryRanges[WdStoryType.wdPrimaryFooterStory]);
                //}
                //catch (System.Exception) { }
                //
                //try
                //{
                //    storyRanges.Add(doc.StoryRanges[WdStoryType.wdFirstPageHeaderStory]);
                //}
                //catch (System.Exception) { }
                //try
                //{
                //    storyRanges.Add(doc.StoryRanges[WdStoryType.wdFirstPageFooterStory]);
                //}
                //catch (System.Exception) { }
                //
                //try
                //{
                //    storyRanges.Add(doc.StoryRanges[WdStoryType.wdEvenPagesHeaderStory]);
                //}
                //catch (System.Exception) { }
                //try
                //{
                //    storyRanges.Add(doc.StoryRanges[WdStoryType.wdEvenPagesFooterStory]);
                //}
                //catch (System.Exception) { }

                foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
                {
                    foreach (HeaderFooter item in section.Headers)
                    {
                        storyRanges.Add(item.Range);
                    }
                    foreach (HeaderFooter item in section.Footers)
                    {
                        storyRanges.Add(item.Range);
                    }
                }
            }
            if (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionSpecialFields)
            {
                try
                {
                    storyRanges.Add(doc.StoryRanges[WdStoryType.wdTextFrameStory]);
                }
                catch (System.Exception) { }

                //foreach (Shape shape in doc.Shapes)
                //{
                //    try
                //    {
                //        storyRanges.Add(shape.TextFrame.TextRange);
                //    }
                //    catch (System.Exception) { }
                //}
                //foreach (InlineShape shape in doc.InlineShapes)
                //{
                //    try
                //    {
                //        storyRanges.Add(shape.Range);
                //    }
                //    catch (System.Exception) { }
                //}
            }

            int countOfChanges = 0;
            foreach (Range rng in storyRanges)
            {
                if (rng.Text.Trim().Length > 2)
                {
                    try
                    {
                        int startRange = rng.Start;
                        int endRange = rng.End;
                        rng.Find.ClearFormatting();
                        rng.Find.Replacement.ClearFormatting();

                        int previousEndRange = 0;
                        foreach (SearchReplaceModel cm in correctionModels)
                        {
                            rng.SetRange(startRange, endRange);

                            while (rng.Find.Execute(FindText: cm.Search,
                                            ReplaceWith: cm.Replace,
                                            Replace: WdReplace.wdReplaceOne,
                                            MatchWildcards: cm.Wildcard == "1" ? true : false,
                                            MatchCase: cm.MatchCase,
                                            MatchWholeWord: cm.MatchWholeWord,
                                            MatchKashida: cm.MatchKashida,
                                            MatchDiacritics: cm.MatchDiacritics,
                                            MatchAlefHamza: cm.MatchAlefHamza,
                                            Forward: true,
                                            Wrap: WdFindWrap.wdFindStop))
                            {
                                if (!rng.Find.Found)
                                    break;

                                if (rng.End == previousEndRange)
                                    break;
                                else
                                    previousEndRange = rng.End;

                                Debug.WriteLine("Text>" + rng.Text + "\n\nrng.Start>" + rng.Start + " previousRange.End>" + rng.End + " endRange>" + endRange);

                                if (cm.CounterUp != "0")
                                    countOfChanges++;

                                rng.SetRange(rng.End, endRange);
                            }
                        }
                        rng.Find.ClearFormatting();
                        rng.Find.Replacement.ClearFormatting();
                    }
                    catch (System.Exception e)
                    {
                        DedicatedFunctions.ShowErrorMessage("خطایی در اعمال جایگذاری در کل سند به وجود آمده" + "\nپیغام خطا:\n" + e.Message);
                    }
                }
            }
            frm?.closeForm(countOfChanges);
        }

        #endregion

        #region Document
        internal static void closeDocument(Document doc, WdSaveOptions wdSaveOptions, bool withApplicationQuit = true)
        {
            //Globals.ThisAddIn.Application.NormalTemplate.Saved = true;
            if (Globals.ThisAddIn.Application.Documents.Count > 1)
            {
                doc.Close(wdSaveOptions);//System.Runtime.InteropServices.COMException: 'Command failed'
            }
            else
            {
                Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                Ribbon.RibbonControlsVisibility(false);

                doc.Close(wdSaveOptions);

                if (withApplicationQuit)
                    Globals.ThisAddIn.Application.Quit(wdSaveOptions);
            }
            //Globals.ThisAddIn.Application.NormalTemplate.Saved = true;
        }
        #endregion

        #region Server
        internal static void uploadDocument(Document doc, bool showMessage)
        {
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                DedicatedFunctions.saveDocument(doc);

                JsonObject jsonVariables = DedicatedFunctions.variablesToJsonServer(doc);

                Microsoft.Office.Interop.Word.ContentControl[] abstractContentControl = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_Fa.ToString());
                if (abstractContentControl != null && abstractContentControl.Length != 0)
                {
                    Range rangeAbstract = abstractContentControl[0].Range;
                    if (rangeAbstract != null)
                    {
                        if (jsonVariables.ContainsKey(VariableFieldIDs._variable_field_Abstract_Fa.ToString()))
                            jsonVariables[VariableFieldIDs._variable_field_Abstract_Fa.ToString()] = rangeAbstract.Text;
                        else
                            jsonVariables.Add(VariableFieldIDs._variable_field_Abstract_Fa.ToString(), rangeAbstract.Text);
                    }
                }

                string token = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UserToken.ToString());
                string documentID = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString());
                string urlParameters = "save/parsanegar/1?id=" + documentID + "&config=" + jsonVariables.ToString();

                //add File
                var formData = new MultipartFormDataContent();
                var fileStream = new FileStream(doc.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var fileContent = new StreamContent(fileStream);
                //formData.Add(fileContent , "file" , doc.Name);
                formData.Add(fileContent, "file", "documentfile.docx");

                DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
                OnResult =>
                {
                    JsonDocument document = JsonDocument.Parse(OnResult);
                    JsonElement root = document.RootElement;

                    root.TryGetProperty("updated", out JsonElement updatedAtElement);
                    string updatedAt = updatedAtElement.GetString();
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedAt.ToString(), updatedAt);
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedFile.ToString(), updatedAt);
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedConfig.ToString(), updatedAt);
                    DedicatedFunctions.saveDocument(doc);

                    if (showMessage)
                        DedicatedFunctions.ShowMessage("سند با موفقیت آپلود شد", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                },
                OnFailed =>
                {
                    if (showMessage)
                        DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable + "\n" +
                            OnFailed.StatusCode + "> " + OnFailed.ReasonPhrase);
                }, formData);
            }
        }
        internal static void sendBugReport(int reportType, string text, StreamContent fileContent = null, string fileExtension = null)
        {
            string urlParameters = "report";

            var multipartContent = new MultipartFormDataContent
            {
                { new StringContent(reportType.ToString()) , "type" } ,
                { new StringContent(text) , "text" }
            };

            if (fileContent != null)
                multipartContent.Add(fileContent, "file", "reportFile" + fileExtension);

            //Start checking the server
            DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, null,
            onResult =>
            {
            },
            OnFailed =>
            {
            }, multipartContent);
        }

        internal static async Task<bool> checkingUpdate()
        {
            UpdateResultModel updateResult = await DedicatedFunctions.checkUpdate();

            bool result;
            if (updateResult != null)
            {
                if (updateResult.Result == UpdateResultModel.Results.ServerResult_NoUpdateAvailable)
                {
                    result = true;
                }
                else if (updateResult.Result == UpdateResultModel.Results.ServerResult_UpdateOptional)
                {
                    DialogResult dialogResult = DedicatedFunctions.ShowMessage("نسخه جدید از افزونه ارائه شده است و نصب آن اختیاری میباشد" + "\nآیا تمایل به نصب نسخه جدید دارید؟", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (dialogResult == DialogResult.Yes)
                    {
                        updateResult.Content.TryGetProperty("link", out JsonElement link);
                        System.Diagnostics.Process.Start(link.GetString());
                        result = false;
                    }
                    else
                    {
                        result = true;
                    }
                }
                else if (updateResult.Result == UpdateResultModel.Results.ServerResult_UpdateMandatory)
                {
                    DedicatedFunctions.ShowMessage("نسخه جدیدی از افزونه منتشر شده و لازم است که برای ادامه کار نسخه جدید را نصب نمایید, برای اینکار " + "\nگزینه OK را کلیک کرده تا وارد صفحه پرداخت شوید", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                    updateResult.Content.TryGetProperty("link", out JsonElement link);
                    System.Diagnostics.Process.Start(link.GetString());
                    result = false;
                }
                else //ServerProblem and ServerNotAvailable
                {
                    DedicatedFunctions.ShowMessage(ErrorMessages.ErrorServiceUnavailable, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    result = false;
                }
            }
            else
                result = false;

            return result;
        }

        internal static async Task<UpdateResultModel> checkUpdate()
        {
            string version = BugReport.AssemblyVersion.Replace(".", "");

            string URL = StringConstant.PrimaryServerApiBaseAddress;
            string urlParameters = "update?ver=" + version;

            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(URL);

            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            //limit timeout response
            client.Timeout = TimeSpan.FromSeconds(10);

            HttpResponseMessage response;
            try
            {
                response = await client.GetAsync(urlParameters);  // Blocking call! Program will wait here until a response is received or a timeout occurs.
                response.EnsureSuccessStatusCode();
            }
            catch (Exception)
            {
                return new UpdateResultModel(UpdateResultModel.Results.ServerNotAvailable, new JsonElement());
            }

            if (response.IsSuccessStatusCode)
            {
                string dataObjects = await response.Content.ReadAsStringAsync();  //Make sure to add a reference to System.Net.Http.Formatting.dll

                JsonDocument document = JsonDocument.Parse(dataObjects);
                JsonElement root = document.RootElement;

                UpdateResultModel.Results result;

                if (root.TryGetProperty("status", out JsonElement status) && status.GetString() == "1")
                {
                    if (root.TryGetProperty("important", out JsonElement important) && important.GetInt32() == 1)
                    {
                        result = UpdateResultModel.Results.ServerResult_UpdateMandatory;
                    }
                    else if (root.TryGetProperty("important", out JsonElement important2) && important2.GetInt32() == 0)
                    {
                        result = UpdateResultModel.Results.ServerResult_UpdateOptional;
                    }
                    else
                    {
                        result = UpdateResultModel.Results.ServerProblem;
                    }
                }
                else if (root.TryGetProperty("status", out JsonElement status2) && status2.GetString() == "0")
                {
                    result = UpdateResultModel.Results.ServerResult_NoUpdateAvailable;
                }
                else
                    result = UpdateResultModel.Results.ServerProblem;

                return new UpdateResultModel(result, root);
            }
            else
            {
                return new UpdateResultModel(UpdateResultModel.Results.ServerProblem, new JsonElement());
            }
        }

        internal static System.Threading.Tasks.Task<bool> getDocumentServerAndUpdateVariables(Document doc, string token, string documentID)
        {
            string urlParameters = "get-document/" + documentID;

            System.Threading.Tasks.Task<bool> s = DedicatedFunctions.httpGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;

                if (root.TryGetProperty("status", out JsonElement status) && status.GetString() == "1")
                {
                    if (root.TryGetProperty("document", out JsonElement jsonElement))
                    {
                        if (jsonElement.TryGetProperty("config", out JsonElement objectConfig))
                        {
                            JsonElement jsonVariables = JsonDocument.Parse(objectConfig.GetString()).RootElement;

                            string[] variableFields = Enum.GetNames(typeof(VariableFieldIDs));

                            foreach (var serverVariable in jsonVariables.EnumerateObject())
                            {
                                foreach (var localVariable in variableFields)
                                {
                                    try
                                    {
                                        if (serverVariable.Name == localVariable)
                                        {
                                            DedicatedFunctions.setORAddStaticVariableValue(doc, serverVariable.Name, serverVariable.Value.GetString(), true);
                                            break;
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        DedicatedFunctions.ShowErrorMessage("خطایی غیر منتظره رخ داد", email: StringConstant.SupportEmail);
                                    }
                                }
                            }

                            Dictionary<string, ContentControlNames> contentControlAndVariableFieldNames = DedicatedFunctions.getRelationContentControlAndVariableFieldNames();
                            foreach (string variable in variableFields)
                            {
                                try
                                {
                                    bool isSuccessfull = contentControlAndVariableFieldNames.TryGetValue(variable, out ContentControlNames relationContentControl);
                                    if (isSuccessfull)
                                    {
                                        string tempContent = DedicatedFunctions.getStaticVariableValue(doc, variable);
                                        if (tempContent == SettingValues.NotExist)
                                            DedicatedFunctions.setORAddStaticVariableValue(doc, variable, "");

                                        DedicatedFunctions.changeContentControlContents(
                                            doc,
                                            relationContentControl.ToString(), tempContent);
                                    }
                                    else
                                    {
                                        //throw new Exception("variable not exist. ERR: 1050");
                                    }
                                }
                                catch (Exception)
                                {
                                    DedicatedFunctions.ShowErrorMessage("خطایی غیر منتظره رخ داد", email: StringConstant.SupportEmail);
                                }
                            }
                        }
                    }
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("پاسخ سرور به برنامه دارای اشکال است",
                        (int)ErrorCodes.UnexpectedServerResponse, StringConstant.SupportEmail);
                }
                //DedicatedFunctions.ShowErrorMessage("اطلاعات دریافتی از سرور دارای اشکال است.");
            },
            OnFailed =>
            {
            });

            return s;
        }

        internal static JsonObject variablesToJsonServer(Document doc)
        {
            //TODO: variable.Name == VariableOtherNames._variable_other_DocID.ToString())

            JsonObject jsonObject = new JsonObject();
            foreach (Variable variable in doc.Variables)
            {
                string decryptName = "";
                string decryptValue = "";

                try
                {
                    decryptName = DecryptString(variable.Name);

                    if (!decryptName.Contains(InitialVariables.initialVariableFieldNames) &&
                        !decryptName.Contains(InitialVariables.initialVariableIDNames) &&
                        !decryptName.Contains(InitialVariables.initialVariableVersionNames) &&
                        !decryptName.Contains(InitialVariables.initialVariableTypeNames))
                        continue;
                }
                catch (Exception)
                {
                    decryptName = variable.Name;
                }

                try
                {
                    decryptValue = DecryptString(variable.Value);
                }
                catch (Exception)
                {
                    decryptValue = variable.Value;
                }

                jsonObject.Add(decryptName, decryptValue);
            }
            return jsonObject;
        }
        internal static async void httpAsyncPostRequest(string URL, string urlParameters, string bearerToken, Action<string> onResult, Action<HttpResponseMessage> onFailed, HttpContent httpContent = null)
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(URL);

            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            if (bearerToken != null)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + bearerToken);

            if (httpContent == null)
            {
                httpContent = new StringContent("", System.Text.Encoding.UTF8, "application/json");
            }

            HttpResponseMessage response;
            try
            {
                response = await client.PostAsync(urlParameters, httpContent);  // Blocking call! Program will wait here until a response is received or a timeout occurs.
                response.EnsureSuccessStatusCode();
                if (response.IsSuccessStatusCode)
                {
                    string dataObjects = await response.Content.ReadAsStringAsync();  //Make sure to add a reference to System.Net.Http.Formatting.dll
                    onResult(dataObjects);
                }
                else
                {
                    onFailed(response);
                }
            }
            catch (Exception e)
            {
                onFailed(new HttpResponseMessage(System.Net.HttpStatusCode.ServiceUnavailable));
            }
        }
        internal static async void httpAsyncGetRequest(string URL, string urlParameters, string bearerToken, Action<string> onResult, Action<HttpResponseMessage> onFailed)
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(URL);

            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            if (bearerToken != null)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + bearerToken);

            HttpResponseMessage response;
            try
            {
                response = await client.GetAsync(urlParameters);
                response.EnsureSuccessStatusCode();
                if (response.IsSuccessStatusCode)
                {
                    string dataObjects = await response.Content.ReadAsStringAsync();
                    onResult(dataObjects);
                }
                else
                {
                    onFailed(response);
                }
            }
            catch (Exception)
            {
                onFailed(new HttpResponseMessage(System.Net.HttpStatusCode.ServiceUnavailable));
            }
        }
        internal static async Task<bool> httpGetRequest(string URL, string urlParameters, string bearerToken, Action<string> onResult, Action<HttpResponseMessage> onFailed)
        {
            HttpClient client = new HttpClient();
            client.BaseAddress = new Uri(URL);

            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            if (bearerToken != null)
                client.DefaultRequestHeaders.Add("Authorization", "Bearer " + bearerToken);

            HttpResponseMessage response;
            try
            {
                //client.Timeout = TimeSpan.FromSeconds(3);
                response = await client.GetAsync(urlParameters);  // Blocking call! Program will wait here until a response is received or a timeout occurs.
                response.EnsureSuccessStatusCode();
            }
            catch (Exception)
            {
                onFailed(new HttpResponseMessage(System.Net.HttpStatusCode.ServiceUnavailable));
                return false;
            }

            if (response.IsSuccessStatusCode)
            {
                string dataObjects = await response.Content.ReadAsStringAsync();  //Make sure to add a reference to System.Net.Http.Formatting.dll
                onResult(dataObjects);
                return true;
            }
            else
            {
                onFailed(response);
                return false;
            }
        }
        #endregion

        #region OMath
        internal static void recursiveOMathFunctions(OMath oMath)
        {
            foreach (OMathFunction oMathFunction in oMath.Functions)
            {
                OMathArgs oMathArgs = oMathFunction.Args;
                for (int i = 1; i <= oMathArgs.Count; i++)
                {
                    OMath argOMath = oMathArgs.Item(i);
                    argOMath.ConvertToNormalText();

                    Debug.WriteLine("   recursiveOMathArgs, type > " + argOMath.Range.Text);
                    if (oMath.Functions.Count > 1)
                        recursiveOMathFunctions(argOMath);
                }

            }
        }
        #endregion

        #region Poem
        internal static void poemOneColumn(Document doc)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("حالت شعر تک ستونه");

            int type = Properties.Settings.Default.VirastarSettings_CreatePoemType;

            Range previousRange = doc.ActiveWindow.Selection.Range;

            doc.ActiveWindow.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
            if (type == (int)PoemTypes.CreatePoemUsingTextColumns)
                DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection);

            Range afterRange = doc.ActiveWindow.Selection.Range;

            previousRange.Select();

            if (type == (int)PoemTypes.CreatePoemUsingTable)
                poemOneColumnUsingTable(doc);
            else if (type == (int)PoemTypes.CreatePoemUsingTextColumns)
                poemOneColumnUsingTextColumns(doc);

            afterRange.Select();
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        internal static void poemTwoColumn(Document doc)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("حالت شعر دو ستونه");

            int type = Properties.Settings.Default.VirastarSettings_CreatePoemType;

            Range previousRange = doc.ActiveWindow.Selection.Range;

            doc.ActiveWindow.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
            Range afterRange = doc.ActiveWindow.Selection.Range;

            previousRange.Select();

            doc.ActiveWindow.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            doc.ActiveWindow.Selection.ParagraphFormat.SpaceBefore = 0;
            doc.ActiveWindow.Selection.ParagraphFormat.SpaceAfter = 0;

            if (type == (int)PoemTypes.CreatePoemUsingTable)
                poemTwoColumnUsingTable(doc);
            else if (type == (int)PoemTypes.CreatePoemUsingTextColumns)
                poemTwoColumnUsingTextColumns(doc);

            afterRange.Select();
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }

        public static void poemOneColumnUsingTextColumns(Document doc)
        {
            Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
            Selection selection = doc.ActiveWindow.Selection;

            if (selection.Paragraphs.Count <= 1)
            {
                DedicatedFunctions.ShowErrorMessage("بیشتر از یک سطر را باید انتخاب کنید");
                return;
            }

            //TextColumns withBlock1 = selection.PageSetup.TextColumns;

            // "^p">Paragragh Mark	"^l">Manual Line Break
            DedicatedFunctions.replaceText(selection.Range, "^p", "^l", WdReplace.wdReplaceAll, WdFindWrap.wdFindStop, true);

            selection.ParagraphFormat.FirstLineIndent = 0;
            selection.ParagraphFormat.LeftIndent = application.CentimetersToPoints(4.5f);
            selection.ParagraphFormat.RightIndent = application.CentimetersToPoints(4.5f);
            selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustifyLow;
            selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            selection.ParagraphFormat.SpaceBefore = 0;
            selection.ParagraphFormat.SpaceAfter = 0;

            // set textColumens properties
            TextColumns tc = selection.PageSetup.TextColumns;
            tc.SetCount(NumColumns: 1);
            tc.FlowDirection = WdFlowDirection.wdFlowRtl;
        }

        public static void poemTwoColumnUsingTextColumns(Document doc)
        {
            Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
            Selection selection = doc.ActiveWindow.Selection;
            Window win = doc.ActiveWindow;

            if (selection.Paragraphs.Count <= 1)
            {
                DedicatedFunctions.ShowErrorMessage("بیشتر از یک سطر را باید انتخاب کنید");
                return;
            }

            //It is used in the Two Column button function in Word itself.
            if (win.View.SplitSpecial != WdSpecialPane.wdPaneNone)
            {
                win.Panes[2].Close();
            }
            if (win.ActivePane.View.Type != WdViewType.wdPrintView)
            {
                win.ActivePane.View.Type = WdViewType.wdPrintView;
            }

            //If the poem had two or less stanzas(two paragraph)
            int paragraphCounts = selection.Text.Count(s => s == ParagraphAndTextWrapMarks.ParagraphMarkInTextChar);
            if (paragraphCounts <= 2)
            {//Divide the first paragraph into two so that the second stanza of the poem goes to the next line
                string addTwoParagraph = selection.Text;
                if (paragraphCounts == 2)
                {
                    addTwoParagraph = addTwoParagraph.Insert(addTwoParagraph.IndexOf(ParagraphAndTextWrapMarks.ParagraphMarkInText), ParagraphAndTextWrapMarks.ParagraphMarkInText);
                    addTwoParagraph = addTwoParagraph.TrimEnd(ParagraphAndTextWrapMarks.ParagraphMarkInTextChar);
                    selection.Text = addTwoParagraph;
                }
                else if (paragraphCounts == 1)
                {
                    //addTwoParagraph = addTwoParagraph.Insert(addTwoParagraph.IndexOf('\r') , ParagraphAndTextWrapMarks.ParagraphMarkInText);
                    selection.Text = addTwoParagraph;
                }
                else if (paragraphCounts == 0)//selected text inside paragraph
                {
                    addTwoParagraph = addTwoParagraph.Insert(0, "\r\r");
                    addTwoParagraph = addTwoParagraph.Insert(addTwoParagraph.Length, ParagraphAndTextWrapMarks.ParagraphMarkInText);
                    selection.Text = addTwoParagraph;
                    selection.Start = selection.Start + 2;
                }
            }
            else // more than 3 paragragh(organize text for poem)
            {
                string text = "";

                //get paragraphs
                List<string> paragraphs = selection.Text.Split(ParagraphAndTextWrapMarks.ParagraphMarkInTextChar).ToList();

                //remove empty last paragraph if exist
                if (string.IsNullOrEmpty(paragraphs[paragraphs.Count - 1].Trim()))
                    paragraphs.RemoveAt(paragraphs.Count - 1);

                //split paragraph in half
                int half = (int)paragraphs.Count / 2;

                foreach (string even in paragraphs.Where((item, index) => index % 2 == 0))
                {
                    text += even + ParagraphAndTextWrapMarks.ParagraphMarkInText;
                }
                foreach (string odd in paragraphs.Where((item, index) => index % 2 != 0))
                {
                    text += odd + ParagraphAndTextWrapMarks.ParagraphMarkInText;
                }

                //remove Additional paragraph break
                text = text.TrimEnd(ParagraphAndTextWrapMarks.ParagraphMarkInTextChar);

                //set to selection text
                selection.Text = text;
            }

            // insert SectionBreakContinuous start and End
            DedicatedFunctions.insertPageBreak(doc, doc.Range(Start: selection.Start, End: selection.Start), WdBreakType.wdSectionBreakContinuous, false);

            selection.Start = selection.Start + 1;
            selection.PageSetup.SectionDirection = WdSectionDirection.wdSectionDirectionRtl;

            //if end of selection text contains new Line(Paragraph), give up.
            if (selection.Text.EndsWith(ParagraphAndTextWrapMarks.ParagraphMarkInText))
            {
                selection.SetRange(selection.Start, selection.End - 1);
            }
            selection.ParagraphFormat.FirstLineIndent = 0;
            selection.InsertParagraphAfter();
            selection.ParagraphFormat.FirstLineIndent = 0;
            DedicatedFunctions.insertPageBreak(doc, doc.Range(Start: selection.End, End: selection.End), WdBreakType.wdSectionBreakContinuous, false);

            // change paragraph marker
            DedicatedFunctions.replaceText(selection.Range, ParagraphAndTextWrapMarks.ParagraphMarkInFind, ParagraphAndTextWrapMarks.TextWrapMarkInFind, WdReplace.wdReplaceAll, WdFindWrap.wdFindStop, true);

            // set Paragraph Alignment
            selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustifyLow;
            selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            selection.ParagraphFormat.SpaceBefore = 0;
            selection.ParagraphFormat.SpaceAfter = 0;


            // set textColumens properties
            TextColumns tc = selection.PageSetup.TextColumns;
            tc.SetCount(NumColumns: 2);
            tc.EvenlySpaced = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            tc.LineBetween = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
            tc.FlowDirection = WdFlowDirection.wdFlowRtl;
            //tc.Spacing = application.CentimetersToPoints(1.5f);
            //withBlock1.Width = application.CentimetersToPoints(10f);
        }

        public static void poemOneColumnUsingTable(Document doc)
        {
            Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
            Selection selection = doc.ActiveWindow.Selection;

            doc.ActiveWindow.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            doc.ActiveWindow.Selection.ParagraphFormat.SpaceBefore = 0;
            doc.ActiveWindow.Selection.ParagraphFormat.SpaceAfter = 0;

            if (selection.Paragraphs.Count <= 1)
            {
                DedicatedFunctions.ShowErrorMessage("بیشتر از یک سطر را باید انتخاب کنید");
                return;
            }

            //convet text to table with 1 column
            Table table = selection.ConvertToTable(WdTableFieldSeparator.wdSeparateByParagraphs, NumColumns: 1, ApplyBorders: false);
            table.TableDirection = WdTableDirection.wdTableDirectionRtl;
            table.Borders.Enable = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
            table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            table.ID = TableIDs.tableOneColumnPoem;

            //merge cells columns
            table.Columns[1].Cells.Merge();

            //get poem cells
            Cell PoemCell1 = table.Columns[1].Cells[1];

            //insert paragraph in end line
            PoemCell1.Range.InsertParagraphAfter();

            //replace paragraph to textwrap(manual line break)
            DedicatedFunctions.replaceText(PoemCell1.Range, ParagraphAndTextWrapMarks.ParagraphMarkInFind, ParagraphAndTextWrapMarks.TextWrapMarkInFind, WdReplace.wdReplaceAll, WdFindWrap.wdFindStop, true);

            //set ParagraphFormat Options
            //table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
            table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustifyLow;
            table.Range.ParagraphFormat.FirstLineIndent = 0;

            //set table width
            float pageWidth = getActiveAreaDocumentSize(doc).Width;
            pageWidth = application.PointsToInches(pageWidth) - 2;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
            table.PreferredWidth = application.InchesToPoints(pageWidth);
        }

        public static void poemTwoColumnUsingTable(Document doc)
        {
            Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
            Selection selection = doc.ActiveWindow.Selection;

            doc.ActiveWindow.Selection.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
            doc.ActiveWindow.Selection.ParagraphFormat.SpaceBefore = 0;
            doc.ActiveWindow.Selection.ParagraphFormat.SpaceAfter = 0;

            if (selection.Paragraphs.Count <= 1)
            {
                DedicatedFunctions.ShowErrorMessage("بیشتر از یک سطر را باید انتخاب کنید");
                return;
            }

            //convet text to table with 2 column
            Table table = selection.ConvertToTable(WdTableFieldSeparator.wdSeparateByParagraphs, NumColumns: 2, ApplyBorders: false);
            table.TableDirection = WdTableDirection.wdTableDirectionRtl;
            table.Borders.Enable = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
            table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            table.ID = TableIDs.tableTwoColumnPoem;

            // add space beetwen two column poem
            table.Columns.Add(table.Columns[1]);
            table.Columns[2].Width = application.MillimetersToPoints(10);

            //fit to page
            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
            table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);

            //merge cells columns
            table.Columns[1].Cells.Merge();
            table.Columns[2].Cells.Merge();
            table.Columns[3].Cells.Merge();

            //get poem cells
            Cell PoemCell1 = table.Columns[1].Cells[1];
            Cell PoemCell2 = table.Columns[3].Cells[1];

            //insert paragraph in end line
            PoemCell1.Range.InsertParagraphAfter();
            PoemCell2.Range.InsertParagraphAfter();

            //replace paragraph to textwrap(manual line break)
            DedicatedFunctions.replaceText(PoemCell1.Range, ParagraphAndTextWrapMarks.ParagraphMarkInFind, ParagraphAndTextWrapMarks.TextWrapMarkInFind, WdReplace.wdReplaceAll, WdFindWrap.wdFindStop, true);
            DedicatedFunctions.replaceText(PoemCell2.Range, ParagraphAndTextWrapMarks.ParagraphMarkInFind, ParagraphAndTextWrapMarks.TextWrapMarkInFind, WdReplace.wdReplaceAll, WdFindWrap.wdFindStop, true);

            //set ParagraphFormat Options
            //table.Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpace1pt5;
            table.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustifyLow;
            table.Range.ParagraphFormat.FirstLineIndent = 0;

            //set table width
            float pageWidth = getActiveAreaDocumentSize(doc).Width;
            pageWidth = application.PointsToInches(pageWidth) - 1f;
            table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints;
            table.PreferredWidth = application.InchesToPoints(pageWidth);
        }
        #endregion

        #region get Size
        public static System.Drawing.SizeF getActiveAreaDocumentSize(Range range)
        {
            return new System.Drawing.SizeF(range.PageSetup.PageWidth - range.PageSetup.LeftMargin - range.PageSetup.RightMargin, range.PageSetup.PageHeight - range.PageSetup.TopMargin - range.PageSetup.BottomMargin);
        }
        public static System.Drawing.SizeF getPageDocumentSize(Range range)
        {
            return new System.Drawing.SizeF(range.PageSetup.PageWidth, range.PageSetup.PageHeight);
        }

        public static System.Drawing.SizeF getActiveAreaDocumentSize(Document doc)
        {
            return new System.Drawing.SizeF(doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin, doc.PageSetup.PageHeight - doc.PageSetup.TopMargin - doc.PageSetup.BottomMargin);
        }
        public static System.Drawing.SizeF getPageDocumentSize(Document doc)
        {
            return new System.Drawing.SizeF(doc.PageSetup.PageWidth, doc.PageSetup.PageHeight);
        }
        #endregion

        #region Find
        //Active Selection

        internal static void replaceText(Range rng, string find, string replace, WdReplace replaceType, WdFindWrap wrapType, bool forward)
        {
            rng.Find.ClearFormatting();
            rng.Find.Replacement.ClearFormatting();
            if (wrapType != WdFindWrap.wdFindAsk)
                rng.Find.Execute(FindText: find, ReplaceWith: replace, Replace: replaceType, Wrap: wrapType, Forward: forward);
            rng.Find.Replacement.ClearFormatting();
            rng.Find.ClearFormatting();
        }

        internal static void clearFormmattingFind(Range rng)
        {
            rng.Find.ClearFormatting();
            rng.Find.Replacement.ClearFormatting();
        }
        internal static void changeToStandardDigit(Range rng, bool toPersian)
        {
            List<string> englishNumber = new List<string>() { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9" };
            List<string> persianNumber = new List<string>() { "۰", "۱", "۲", "۳", "۴", "۵", "۶", "۷", "۸", "۹" };

            if (Properties.Settings.Default.VirastarSettings_IncludeChangeDigitCharacters)
            {
                if (toPersian)
                {
                    for (int i = 0; i < persianNumber.Count; i++)
                    {
                        rng.Find.Execute(englishNumber[i], Replace: WdReplace.wdReplaceAll, ReplaceWith: persianNumber[i]);
                    }
                }
                else
                {
                    for (int i = 0; i < persianNumber.Count; i++)
                    {
                        rng.Find.Execute(persianNumber[i], Replace: WdReplace.wdReplaceAll, ReplaceWith: englishNumber[i]);
                    }
                }
            }
            else
            {
                for (int i = 0; i < persianNumber.Count; i++)
                {
                    rng.Find.Execute(persianNumber[i], Replace: WdReplace.wdReplaceAll, ReplaceWith: englishNumber[i]);
                }
            }

        }
        internal static void changeIntegerNumbers(Range rng, bool toPersian)
        {
            string defaultPersianFont = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleNormal].Font.NameBi;
            string defaultEnglishFont = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleNormal].Font.Name;

            if (toPersian)
            {
                rng.Find.Replacement.LanguageID = WdLanguageID.wdPersian;
                rng.Find.Replacement.Font.Name = defaultPersianFont;
            }
            else
            {
                rng.Find.Replacement.LanguageID = WdLanguageID.wdEnglishUS;
                rng.Find.Replacement.Font.Name = defaultEnglishFont;
            }
            rng.Find.Execute("^#", ReplaceWith: "", MatchWildcards: false, Forward: true, Replace: WdReplace.wdReplaceAll, Wrap: WdFindWrap.wdFindStop);
        }
        internal static void changeDoubleNumbers(Range rng, bool toPersian)
        {
            rng.Find.Wrap = WdFindWrap.wdFindStop;

            if (toPersian)
            {
                rng.Find.Execute("([0-9]{1,}).([0-9]{1,})", ReplaceWith: @"\2/\1", MatchWildcards: true, Forward: true, Replace: WdReplace.wdReplaceAll);
            }
            else
            {
                //selection.Find.LanguageID = WdLanguageID.wdPersian;
                rng.Find.Execute("([0-9]{1,})/([0-9]{1,})", ReplaceWith: @"\2.\1", MatchWildcards: true, Forward: true, Replace: WdReplace.wdReplaceAll);
            }
        }

        #endregion

        #region Insert

        internal static void insertParagraph(Selection selection, int count = 1)
        {
            for (int i = 0; i < count; i++)
            {
                selection.TypeParagraph();
            }
        }
        internal static void insertParagraph(Selection selection, string text)
        {
            selection.TypeParagraph();
            selection.TypeText(text);
        }
        internal static void insertParagraph(Selection selection, int count, WdParagraphAlignment paragraphAlignment)
        {
            for (int i = 0; i < count; i++)
            {
                selection.TypeParagraph();
                selection.ParagraphFormat.Alignment = paragraphAlignment;
            }
        }

        internal static void insertPageBreak(Document doc, Range rng, WdBreakType wdBreakType = WdBreakType.wdPageBreak, bool withTypeParagraph = true)
        {
            if (withTypeParagraph)
                rng.InsertParagraphAfter();

            rng.InsertBreak(wdBreakType);

        }

        internal static Range insertText(Document doc, Selection selection, string text)
        {
            selection.TypeText(text);
            if (text.Length != 0)
                return doc.Range(selection.Previous(WdUnits.wdCharacter, text.Length).Start, selection.Range.Start);
            else
                return selection.Range;
        }
        internal static Range insertText(Document doc, Selection selection, string text, string fontFamily, bool isInline = false)
        {
            if (!isInline)
                insertParagraph(selection, 1);

            selection.Font.Name = fontFamily;
            selection.TypeText(text);
            if (text.Length != 0)
                return doc.Range(selection.Previous(WdUnits.wdCharacter, text.Length).Start, selection.Range.Start);
            else
                return selection.Range;

        }
        internal static Range insertText(Document doc, Selection selection, string text, string fontFamily, bool isInline = false, TextTypes textType = TextTypes.BOTH, int fontSize = -1, bool bold = false, bool italic = false)
        {

            if (!isInline)
                insertParagraph(selection, 1);

            selection.Font.Name = fontFamily;
            if (textType == TextTypes.RTL)
            {
                if (fontSize != -1)
                    selection.Font.SizeBi = fontSize;
                selection.Font.BoldBi = Convert.ToInt32(bold);
                selection.Font.ItalicBi = Convert.ToInt32(italic);
            }
            else if (textType == TextTypes.LTR)
            {
                if (fontSize != -1)
                    selection.Font.Size = fontSize;
                selection.Font.Bold = Convert.ToInt32(bold);
                selection.Font.Italic = Convert.ToInt32(italic);
            }
            else if (textType == TextTypes.BOTH)
            {
                if (fontSize != -1)
                    selection.Font.Size = fontSize;
                selection.Font.Bold = Convert.ToInt32(bold);
                selection.Font.Italic = Convert.ToInt32(italic);

                if (fontSize != -1)
                    selection.Font.SizeBi = fontSize;
                selection.Font.BoldBi = Convert.ToInt32(bold);
                selection.Font.ItalicBi = Convert.ToInt32(italic);
            }


            selection.TypeText(text);
            if (text.Length != 0)
                return doc.Range(selection.Previous(WdUnits.wdCharacter, text.Length).Start, selection.Range.Start);
            else
                return selection.Range;
        }
        internal static Range insertPicture(Selection selection, string imagePath, WdWrapType WrapType = WdWrapType.wdWrapTopBottom)
        {
            try
            {
                Microsoft.Office.Interop.Word.Shape shape = selection.InlineShapes.AddPicture(imagePath).ConvertToShape();
                shape.Top = 0;
                shape.Left = 0;
                shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                shape.WrapFormat.Type = WrapType;
                shape.WrapFormat.Side = WdWrapSideType.wdWrapBoth;
                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                shape.Height = Globals.ThisAddIn.Application.MillimetersToPoints(28);
                shape.WrapFormat.DistanceTop = Globals.ThisAddIn.Application.MillimetersToPoints(0);
                shape.WrapFormat.DistanceBottom = Globals.ThisAddIn.Application.MillimetersToPoints(0);
                selection.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters, 1);
                selection.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignMiddles, 1);


                shape.Select();
                Range rng = selection.Range;
                selection.EndKey();
                return rng;
            }
            catch (Exception e)
            {
                ShowErrorMessage("فایل عکس مورد نظر یافت نشد!" + "\n" + "پیغام خطا:" + "\n" + e.Message);
                return null;
            }
        }

        internal static Shape insertFormPicture(Selection selection, string imagePath, WdWrapType wrapType = WdWrapType.wdWrapBehind)
        {
            try
            {
                Microsoft.Office.Interop.Word.Shape shape = selection.InlineShapes.AddPicture(imagePath).ConvertToShape();

                shape.Top = 0;
                shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                shape.WrapFormat.Type = wrapType;
                shape.WrapFormat.Side = WdWrapSideType.wdWrapBoth;
                shape.WrapFormat.DistanceTop = Globals.ThisAddIn.Application.MillimetersToPoints(0);
                shape.WrapFormat.DistanceBottom = Globals.ThisAddIn.Application.MillimetersToPoints(0);
                selection.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignCenters, 1);
                selection.ShapeRange.Align(Microsoft.Office.Core.MsoAlignCmd.msoAlignMiddles, 1);

                return shape;
            }
            catch (Exception e)
            {
                ShowErrorMessage("فایل عکس مورد نظر یافت نشد!" + "\n" + "پیغام خطا:" + "\n" + e.Message);
            }
            return null;
        }

        internal static void insertNormalTable(Selection selection, Range range, int row, int column, bool paragraphBeforeAfter, WdDefaultTableBehavior wdDefaultTableBehavior = WdDefaultTableBehavior.wdWord8TableBehavior, WdAutoFitBehavior wdAutoFitBehavior = WdAutoFitBehavior.wdAutoFitWindow)
        {
            if (paragraphBeforeAfter)
                insertParagraph(selection);
            range.Tables.Add(range, row, column, wdDefaultTableBehavior, wdAutoFitBehavior);
            if (paragraphBeforeAfter)
                insertParagraph(selection);
        }

        internal static void insertSpecialTable(Document doc, Selection selection, TableTypes tableType, string captionText, string tofCaption = "", string tofTableID = "")
        {
            string defaultPersianFont = doc.Styles[StyleNames.styleNormal].Font.NameBi;

            changeFontOptions(selection.Range, defaultPersianFont, 14, TextTypes.BOTH, true);
            insertParagraph(selection, 2);
            insertText(doc, selection, captionText, defaultPersianFont, false);

            insertParagraph(selection);
            //start text
            selection.ParagraphFormat.TabStops.Add(Globals.ThisAddIn.Application.MillimetersToPoints(1), WdAlignmentTabAlignment.wdLeft, WdTabLeader.wdTabLeaderSpaces);
            insertText(doc, selection, "\t");
            insertText(doc, selection, "عنوان");
            //end text
            float getLastWorkDocumentPosition = getActiveAreaDocumentSize(doc).Width - selection.PageSetup.CharsLine;
            selection.ParagraphFormat.TabStops.Add(getLastWorkDocumentPosition, WdAlignmentTabAlignment.wdRight, WdTabLeader.wdTabLeaderSpaces);
            insertText(doc, selection, "\t");
            selection.TypeText("صفحه");

            insertParagraph(selection, 1);
            selection.ParagraphFormat.TabStops.ClearAll();

            TableOfFigures tableOfFigures = null;

            if (tableType == TableTypes.TableOfContents)
            {
                //goto link http://www.vbaexpress.com/forum/archive/index.php/t-7028.html
                doc.Fields.Add(selection.Range, WdFieldType.wdFieldTOC);
                //Globals.ThisAddIn.Application.ActiveDocument.Fields.Add(selection.Range, WdFieldType.wdFieldTOC, @"\h \z  \t SaNaLevel2,SaNaLevel3,1,SaNaLevel4,1,SaNaLevel5,1,SaNaLevel1,1", false);
            }
            else if (tableType == TableTypes.TableOfFigures)
            {
                changeFontOptions(selection.Range, defaultPersianFont, 14, TextTypes.BOTH, false);
                //Globals.ThisAddIn.Application.ActiveDocument.TablesOfFigures.Add(selection.Range, tofCaption, true, false, 1, 3, true, tofTableID, true, true, tofAddedStyle, true, true);
                tableOfFigures = doc.TablesOfFigures.Add(selection.Range, tofCaption, true, TableID: tofTableID);
                tableOfFigures.TabLeader = WdTabLeader.wdTabLeaderDashes;
                doc.TablesOfFigures.Format = (WdTofFormat)WdIndexType.wdIndexIndent;
            }

            insertParagraph(selection, 2);
            changeTextDirection(selection, TextDirection.RTL);
            if (tableType == TableTypes.TableOfContents)
            {
                doc.TablesOfContents[doc.TablesOfContents.Count].Range.Font.Size = 12;
                doc.TablesOfContents[doc.TablesOfContents.Count].Range.Font.SizeBi = 12;
                doc.TablesOfContents[doc.TablesOfContents.Count].Range.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
                doc.TablesOfContents[doc.TablesOfContents.Count].Range.Font.BoldBi = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
                doc.TablesOfContents[doc.TablesOfContents.Count].Range.Font.Italic = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
                doc.TablesOfContents[doc.TablesOfContents.Count].Range.Font.ItalicBi = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
            }
            else if (!tableOfFigures.Equals(null))
            {
                tableOfFigures.Range.Font.Size = 12;
                tableOfFigures.Range.Font.SizeBi = 12;
                tableOfFigures.Range.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
                tableOfFigures.Range.Font.BoldBi = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
                tableOfFigures.Range.Font.Italic = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
                tableOfFigures.Range.Font.ItalicBi = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
            }
        }

        internal static void insertHalfSpace(Selection selection)
        {
            if (selection != null)
            {
                selection.RtlPara();
                selection.TypeText("\u200c");
            }
            else
                return;
        }
        #endregion

        #region Set Style
        internal static void textFormating(Selection selection, object style)
        {
            if (selection != null)
                selection.set_Style(style);
            else
                return;

            DedicatedFunctions.changeKeyboardLanguage(selection, KeyboardLanguage.Persian, TextDirection.RTL);
        }
        internal static bool changeStyleInSelection(Selection selection, object styleName)
        {
            try
            {
                selection.set_Style(styleName);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        #endregion

        #region Caption Labels
        internal static CaptionLabel setCaptionLabel(Document doc, string nameOfCaption)
        {
            CaptionLabel caption = Globals.ThisAddIn.Application.CaptionLabels.Add(nameOfCaption);

            caption.NumberStyle = WdCaptionNumberStyle.wdCaptionNumberStyleArabic;

            // 🔴 حذف کامل شماره فصل
            caption.IncludeChapterNumber = true;

            // فقط برای این که خروجی بشه: شکل -1
            caption.Separator = WdSeparatorType.wdSeparatorEnDash;

            return caption;
        }

        internal static void insertCaption(Document doc, Selection selection, string textCaption)
        {
            setCaptionLabel(doc, textCaption);

            try
            {
                if (textCaption == Constants.CaptionLabels.captionFigure)
                {
                    selection.InsertCaption(
                        Constants.CaptionLabels.captionFigure,
                        Position: WdCaptionPosition.wdCaptionPositionBelow
                    );
                    selection.TypeText("– ");
                }
                else if (textCaption == Constants.CaptionLabels.captionTable)
                {
                    selection.InsertCaption(
                        Constants.CaptionLabels.captionTable,
                        Position: WdCaptionPosition.wdCaptionPositionAbove
                    );
                    selection.TypeText("– ");
                }
                else if (textCaption == Constants.CaptionLabels.captionFormula)
                {
                    selection.TypeText("(");
                    selection.InsertCaption(Constants.CaptionLabels.captionFormula);
                    selection.TypeText(")");
                    selection.TypeText("\t\t");
                    selection.OMaths.Add(selection.Range);
                }

                CustomTaskPane customTaskPane =
                    DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CrossReference);

                if (customTaskPane != null && customTaskPane.Visible)
                    ((CrossReferenceTaskPane)customTaskPane.Control)
                        .crossReferenceControl.RefreshList();
            }
            catch (Exception e)
            {
                if (e.HResult == -2146823683)
                {
                    DedicatedFunctions.ShowErrorMessage(
                        "عنوان نمیتواند در سر صفحه یا پاصفحه قرار بگیرد"
                    );
                }
            }
        }


        internal static List<string> getCaptionLabels(Document doc, bool accessGranted)
        {
            List<string> captionLabels = new List<string>();

            if (accessGranted)
            {
                Universities university = DedicatedFunctions.getUniversity(doc);
                TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);
                DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);

                List<TemplateRelationshipModel> templateModels = TemplateAccess.getTemplateRelationshipModelList(university, templateType, documentType, true);

                bool isFigureExist = false;
                bool isTableExist = false;
                bool isMapExist = false;
                foreach (var item in templateModels)
                {
                    if (item.SubTemplateID == TemplateRelationshipModel.SubTemplateIDs.TableOfFigures)
                        isFigureExist = true;
                    else if (item.SubTemplateID == TemplateRelationshipModel.SubTemplateIDs.TableOfTables)
                        isTableExist = true;
                    else if (item.SubTemplateID == TemplateRelationshipModel.SubTemplateIDs.TableOfMaps)
                        isMapExist = true;
                }
                if (isFigureExist)
                    captionLabels.Add(DedicatedFunctions.setCaptionLabel(doc, Constants.CaptionLabels.captionFigure).Name);

                if (isTableExist)
                    captionLabels.Add(DedicatedFunctions.setCaptionLabel(doc, Constants.CaptionLabels.captionTable).Name);

                if (isMapExist)
                    captionLabels.Add(DedicatedFunctions.setCaptionLabel(doc, Constants.CaptionLabels.captionMap).Name);

                captionLabels.Add(DedicatedFunctions.setCaptionLabel(doc, Constants.CaptionLabels.captionFormula).Name);
            }
            else
            {
                captionLabels.Add(DedicatedFunctions.setCaptionLabel(doc, Constants.CaptionLabels.captionFigure).Name);
                captionLabels.Add(DedicatedFunctions.setCaptionLabel(doc, Constants.CaptionLabels.captionTable).Name);
                captionLabels.Add(DedicatedFunctions.setCaptionLabel(doc, Constants.CaptionLabels.captionFormula).Name);
            }

            return captionLabels;
        }

        internal static void removeAllCrossReferences(Document doc)
        {
            AccessType accessType = DedicatedFunctions.hasAccess(doc);
            bool accessGranted = accessType == AccessType.AccessGranted;
            List<string> captionLabels = getCaptionLabels(doc, accessGranted);

            foreach (string captionLabel in captionLabels)
            {
                List<Range> captions = new List<Range>();

                //get number of caption
                bool includeChapterNumber = Globals.ThisAddIn.Application.CaptionLabels[captionLabel].IncludeChapterNumber;
                int chapterStyleLevel = Globals.ThisAddIn.Application.CaptionLabels[captionLabel].ChapterStyleLevel;

                if (includeChapterNumber)
                {
                    foreach (Field field in doc.Fields)
                    {
                        if (field.Type == WdFieldType.wdFieldStyleRef)
                        {
                            if (field.Code.Text.ToLower()
                                .Replace(" ", "")
                                .Contains("\\s" + chapterStyleLevel.ToString()) || field.Code.Text.ToLower().Replace(" ", "").Contains(chapterStyleLevel.ToString() + "\\s"))
                            {
                                if (field.Next.Type == WdFieldType.wdFieldSequence)
                                {
                                    if (field.Next.Code.Text.ToLower().Contains(captionLabel.ToLower()) && field.Next.Code.Text.ToLower().Replace(" ", "").Contains("\\s" + chapterStyleLevel.ToString()))
                                    {
                                        captions.Add(field.Next.Result);
                                    }
                                }
                            }
                        }
                    }
                }

                foreach (Range rng in captions)
                {
                    rng.Select();
                    doc.ActiveWindow.Selection.HomeKey(WdUnits.wdLine);
                    doc.ActiveWindow.Selection.EndKey(WdUnits.wdLine, WdMovementType.wdExtend);
                    if (doc.ActiveWindow.Selection.OMaths.Count != 0)
                    {
                        Range rngSelection = doc.ActiveWindow.Selection.Range;
                        Range oMathRng = rngSelection.OMaths[1].Range;
                        rngSelection.SetRange(rngSelection.Start, oMathRng.Start);
                        rngSelection.Delete();
                    }
                    doc.ActiveWindow.Selection.Delete();

                }
            }
        }
        #endregion

        #region Font
        internal static void changeDefaultFont(Document doc, string persianFont = "", string englishFont = "")
        {
            try
            {
                if (persianFont != "" && englishFont != "")
                {
                    foreach (Style item in doc.Styles)
                    {
                        item.Font.Name = englishFont;
                        item.Font.NameBi = persianFont;
                    }
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_DefaultPersianFont.ToString(), persianFont);
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_DefaultEnglishFont.ToString(), englishFont);
                }
                else if (persianFont != "")
                {
                    foreach (Style item in doc.Styles)
                    {
                        item.Font.NameBi = persianFont;
                    }
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_DefaultPersianFont.ToString(), persianFont);
                }
                else if (englishFont != "")
                {
                    foreach (Style item in doc.Styles)
                    {
                        item.Font.Name = persianFont;
                    }
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_DefaultEnglishFont.ToString(), englishFont);
                }
            }
            catch (System.Exception e)
            {
                DedicatedFunctions.ShowErrorMessage(e.Message);
            }

        }

        internal static void changeFontToDefaultOptions(Range range, TextTypes textType = TextTypes.BOTH)
        {
            string defaultEnglishFont = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleNormal].Font.Name;
            string defaultPersianFont = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleNormal].Font.NameBi;

            range.Font.Name = defaultPersianFont;
            range.Font.NameBi = defaultEnglishFont;
            if (textType == TextTypes.RTL)
            {
                range.Font.SizeBi = 14;
                range.Font.BoldBi = 1;
                range.Font.ItalicBi = 1;
            }
            else if (textType == TextTypes.LTR)
            {
                range.Font.Size = 14;
                range.Font.Bold = 1;
                range.Font.Italic = 1;
            }
            else if (textType == TextTypes.BOTH)
            {
                range.Font.Size = 14;
                range.Font.Bold = 1;
                range.Font.Italic = 1;

                range.Font.SizeBi = 14;
                range.Font.BoldBi = 1;
                range.Font.ItalicBi = 1;
            }
        }
        internal static void changeFontSize(Range range, TextTypes textType = TextTypes.BOTH, int fontSize = 10)
        {
            if (textType == TextTypes.RTL)
                range.Font.SizeBi = fontSize;
            else if (textType == TextTypes.LTR)
                range.Font.Size = fontSize;
            else if (textType == TextTypes.BOTH)
            {
                range.Font.Size = fontSize;
                range.Font.SizeBi = fontSize;
            }
        }
        internal static void changeFontOptions(Range range, string fontname, int fontSize = -1, TextTypes textType = TextTypes.BOTH, bool bold = false, bool italic = false)
        {
            if (textType == TextTypes.RTL)
            {
                range.Font.NameBi = fontname;
                if (fontSize != -1)
                    range.Font.SizeBi = fontSize;
                range.Font.BoldBi = Convert.ToInt32(bold);
                range.Font.ItalicBi = Convert.ToInt32(italic);
            }
            else if (textType == TextTypes.LTR)
            {
                range.Font.Name = fontname;
                if (fontSize != -1)
                    range.Font.Size = fontSize;
                range.Font.Bold = Convert.ToInt32(bold);
                range.Font.Italic = Convert.ToInt32(italic);
            }
            else if (textType == TextTypes.BOTH)
            {
                range.Font.Name = fontname;
                if (fontSize != -1)
                    range.Font.Size = fontSize;
                range.Font.Bold = Convert.ToInt32(bold);
                range.Font.Italic = Convert.ToInt32(italic);

                range.Font.NameBi = fontname;
                if (fontSize != -1)
                    range.Font.SizeBi = fontSize;
                range.Font.BoldBi = Convert.ToInt32(bold);
                range.Font.ItalicBi = Convert.ToInt32(italic);
            }
        }

        internal static void setEmbedFonts(Document doc, bool embed = true)
        {
            doc.EmbedTrueTypeFonts = embed;
            doc.SaveSubsetFonts = embed;
            doc.DoNotEmbedSystemFonts = true;
        }
        #endregion

        #region Dialog
        internal static void ShowErrorMessage(string message, int errorCode = (int)ErrorCodes.Nothing, string email = "", string mobile = "", bool sendErrorToServer = false)
        {
            string errorMessage = message;
            if (errorCode != (int)ErrorCodes.Nothing)
                errorMessage += "\nکد خطا: " + (int)errorCode;

            if (!message.Contains("سرور") || sendErrorToServer)
            {
                System.Threading.Tasks.Task.Run(() =>
                {
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        string text = errorMessage;

                        bool isUserLogined = !string.IsNullOrEmpty(Properties.Settings.Default.UserToken);

                        text += getSettingsInformation(isUserLogined);
                        text += getInformation();

                        //Dispatcher.Invoke(() =>
                        //{
                        sendBugReport((int)ReportTypes.ApplicationReport, text);
                        //});
                    });
                });
            }

            if (email != "")
                errorMessage += "\nمورد را به این یارانامه ارسال نموده و توضیحات لازم را ارائه نمایید:\n" + email;

            if (mobile != "")
                errorMessage += "\nمیتوانید مورد را از طریق شماره موبایل " + mobile + " با ما در جریان بگذارید";

            MessageBox.Show(errorMessage, StringConstant.NameOfProject, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }
        internal static DialogResult ShowMessage(string message,
            MessageBoxButtons messageBoxButtons = MessageBoxButtons.OK,
            MessageBoxIcon messageBoxIcon = MessageBoxIcon.Information,
            MessageBoxDefaultButton messageBoxDefaultButton = MessageBoxDefaultButton.Button2,
            string caption = StringConstant.NameOfProject)
        {
            return MessageBox.Show(message, caption, messageBoxButtons, messageBoxIcon, messageBoxDefaultButton, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }
        #endregion

        #region get Informations
        internal static string getSettingsInformation(bool isUserLogined)
        {
            string version = BugReport.AssemblyVersion.Replace(".", "");

            string informations = "";
            string separator = "----------------------------";
            string tab = "\t";
            string newLine = "\n";

            informations += newLine + separator + newLine;
            informations += "اطلاعات ذخیره شده:" + newLine;

            informations += tab + "شماره نسخه: " + version + newLine;

            informations += tab + "حساب کاربری: " + (isUserLogined ? Properties.Settings.Default.Mobile : "خالی") + newLine;
            //informations += tab + "توکن: " + (isUserLogined ? Properties.Settings.Default.UserToken : "خالی") + newLine;
            informations += tab + "آدرس پوشه محیط کاری: " + (!string.IsNullOrEmpty(Properties.Settings.Default.WorkSpaceDirectory) ? Properties.Settings.Default.WorkSpaceDirectory : "خالی") + newLine;

            informations += tab + "تنظیمات عنوان - نمایش دیالوگ انتخاب عکس: " + (Properties.Settings.Default.CaptionSettings_ShowDialogCaptionFigure ? "فعال" : "غیرفعال") + newLine;
            informations += tab + "تنظیمات ویراستار - شامل شدن محدوده پانوشت در تصحیح نیم فاصله : " + (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionFootnote ? "فعال" : "غیرفعال") + newLine;
            informations += tab + "تنظیمات ویراستار - شامل شدن محدوده سرصفحه و پاصفحه در تصحیح نیم فاصله : " + (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionHeadersFooters ? "فعال" : "غیرفعال") + newLine;
            informations += tab + "تنظیمات ویراستار - شامل شدن محدوده فیلد های خاص در تصحیح نیم فاصله : " + (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionSpecialFields ? "فعال" : "غیرفعال") + newLine;
            informations += tab + "تنظیمات ویراستار - کاراکتر های اعداد فارسی استاندارد در تبدیل اعداد جایگزین شود: " + (Properties.Settings.Default.VirastarSettings_IncludeChangeDigitCharacters ? "فعال" : "غیرفعال") + newLine;

            int poemType = Properties.Settings.Default.VirastarSettings_CreatePoemType;
            string poemTypeString;
            if (poemType == (int)PoemTypes.CreatePoemUsingTable)
                poemTypeString = "ساخت شعر با استفاده از Table";
            else if (poemType == (int)PoemTypes.CreatePoemUsingTextColumns)
                poemTypeString = "ساخت شعر با استفاده از Text Columns";
            else
                throw new Exception("مقدار نادرست در VirastarSettings_CreatePoemType\n مقدار: " + poemType.ToString());

            informations += tab + "تنظیمات ویراستار - نوع حالت شعر: " + poemTypeString + newLine;

            return informations;
        }
        internal static string getInformation()
        {
            string version = BugReport.AssemblyVersion;

            //Globals.ThisAddIn.Application.Dialogs[WdWordDialog.]

            string informations = "";
            string separator = "----------------------------";
            string tab = "\t";
            string newLine = "\n";

            string wordSoftwarePath = "";
            if (File.Exists(Globals.ThisAddIn.Application.Path + "\\winword.exe"))
            {
                wordSoftwarePath = Globals.ThisAddIn.Application.Path + "\\winword.exe";
            }

            List<string> addInList = new List<string>() { };
            foreach (Microsoft.Office.Interop.Word.AddIn addIn in Globals.ThisAddIn.Application.AddIns)
            {
                addInList.Add(addIn.Name);
            }
            foreach (Microsoft.Office.Core.COMAddIn addIn in Globals.ThisAddIn.Application.COMAddIns)
            {
                addInList.Add(addIn.Description);
            }

            var wmi = new ManagementObjectSearcher("select * from Win32_OperatingSystem").Get().Cast<ManagementObject>().First();
            var cpu = new ManagementObjectSearcher("select * from Win32_Processor").Get().Cast<ManagementObject>().First();
            informations += separator + newLine;
            informations += "اطلاعات پایه:" + newLine;
            informations += tab + "تاریخ و ساعت محلی سیستم: " + System.Management.ManagementDateTimeConverter.ToDateTime((string)wmi["LocalDateTime"]).ToString();
            informations += newLine + separator + newLine;

            informations += "اطلاعات افزونه:" + newLine;


            object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);

            informations += tab + "نام افزونه: " + (attributes.Length == 0 ? "" : ((AssemblyProductAttribute)attributes[0]).Product) + newLine;
            informations += tab + "نسخه افزونه: " + version + newLine;
            informations += separator + newLine;
            informations += "اطلاعات نرم افزار آفیس:" + newLine;
            informations += tab + "آدرس مجموعه آفیس: " + Globals.ThisAddIn.Application.Path + newLine;
            informations += wordSoftwarePath != "" ? tab + "آدرس نرم افزار وُرد: " + wordSoftwarePath + newLine : "";
            informations += tab + "نسخه نرم افزار وُرد: " + Globals.ThisAddIn.Application.Version + " (" + Globals.ThisAddIn.Application.Build + ")\n";
            if (addInList.Count != 0)
            {
                informations += tab + "افزونه های نصب شده:\n";
                foreach (var addIn in addInList)
                {
                    informations += tab + tab + addIn + newLine;
                }
            }

            //informations += tab + "نام Office: " + Globals.ThisAddIn.Application.Options + newLine;
            //informations += tab + "نام Office: " + Globals.ThisAddIn.Application.PickerDialog + newLine;
            informations += separator + newLine;
            informations += "مشخصات سیستم:" + newLine;
            informations += tab + "سیستم عامل: " + ((string)wmi["Caption"]).Trim() + newLine;
            informations += tab + "نسخه سیستم عامل: " + (string)wmi["Version"] + " (" + (string)wmi["OSArchitecture"] + ")" + newLine;

            Graphics graphics = Graphics.FromHwnd(IntPtr.Zero);
            float dpi = (graphics.DpiX * 100) / 96;

            informations += tab + "وضوح نمایش: " + Globals.ThisAddIn.Application.System.HorizontalResolution + "x" + Globals.ThisAddIn.Application.System.VerticalResolution + " (مقیاس: %" + dpi.ToString() + ")" + newLine;
            informations += tab + "حافظه موقت: " + ((UInt64)wmi["TotalVisibleMemorySize"] / 1024).ToString() + "MB (در دسترس: " + ((UInt64)wmi["FreePhysicalMemory"] / 1024).ToString() + "MB)\n";
            informations += tab + "مشخصات پردازنده: " + (string)cpu["Name"] + newLine;
            informations += separator + newLine + "پایان پیام";

            return informations;
        }
        #endregion

        #region Keyboard And Direction
        internal static InputLanguage getInputLanguageByName(string inputName)
        {
            foreach (InputLanguage lang in InputLanguage.InstalledInputLanguages)
            {
                if (lang.Culture.EnglishName.ToLower().Contains(inputName.ToLower()))
                {
                    return lang;
                }
            }
            return null;
        }

        internal static void changeKeyboardLanguage(Constants.KeyboardLanguage keyboardLanguage)
        {
            Globals.ThisAddIn.Application.Keyboard((int)keyboardLanguage);
        }
        internal static bool changeKeyboardLanguage(Selection selection, Constants.KeyboardLanguage keyboardLanguage, TextDirection textDirection)
        {
            Globals.ThisAddIn.Application.Keyboard((int)keyboardLanguage);
            try
            {
                if (textDirection == TextDirection.RTL)
                {
                    selection.RtlPara();
                }
                else if (textDirection == TextDirection.LTR)
                {
                    selection.LtrPara();
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        internal static bool changeTextDirection(Selection selection, TextDirection textDirection)
        {
            try
            {
                if (textDirection == TextDirection.RTL)
                {
                    selection.RtlPara();
                }
                else if (textDirection == TextDirection.LTR)
                {
                    selection.LtrPara();
                }
                return true;
            }
            catch (Exception)
            {
                ShowErrorMessage("خطا در تغییر جهت متن");
                return false;
            }
        }
        #endregion

        #region Paragraph and Page
        internal static void changeParagraphAlignment(Selection selection, WdParagraphAlignment paragraphAlignment)
        {
            selection.ParagraphFormat.Alignment = paragraphAlignment;
        }
        internal static void changeVerticalAlignment(Selection selection, WdVerticalAlignment wdVerticalAlignment)
        {
            selection.PageSetup.VerticalAlignment = wdVerticalAlignment;
        }
        internal static void changeParagraphIndent(Selection selection, float firstLineIndent = -1, float leftIndent = -1, float rightIndent = -1)
        {
            if (firstLineIndent != -1)
            {
                selection.ParagraphFormat.FirstLineIndent = firstLineIndent;
            }
            if (leftIndent != -1)
            {
                selection.ParagraphFormat.LeftIndent = leftIndent;
            }
            if (rightIndent != -1)
            {
                selection.ParagraphFormat.RightIndent = rightIndent;
            }
        }

        #endregion

        #region Working with File
        internal static void getAllWorkSpaceFile(string rootPath, List<FileInfo> rootFilePaths, out List<FileInfo> filePaths)
        {
            foreach (string directory in Directory.GetDirectories(rootPath))
            {
                //except archive and documentTemplate and virastar folders
                if (directory != Properties.Settings.Default.WorkSpaceDirectory + StringConstant.ArchiveFolder.TrimEnd('\\') && directory != Properties.Settings.Default.WorkSpaceDirectory + StringConstant.DocumentsTemplateFolder.TrimEnd('\\') && directory != Properties.Settings.Default.WorkSpaceDirectory + StringConstant.VirastarFolder.TrimEnd('\\'))
                {
                    DirectoryInfo info = new DirectoryInfo(directory);
                    rootFilePaths.AddRange(info.GetFiles());

                    getAllWorkSpaceFile(directory, rootFilePaths, out filePaths);
                }
            }
            filePaths = rootFilePaths;
        }

        internal static void removeFileFromSystem(string filePath)
        {
            if (File.Exists(filePath))
            {
                File.SetAttributes(filePath, FileAttributes.Normal);
                File.Delete(filePath);
            }
        }
        internal static string copyFileToTempFolder(byte[] resources, string FileName, string postFix, string customDirectoryName = "")
        {
            string tempPath = System.IO.Path.GetTempPath();
            string filePath;

            if (string.IsNullOrEmpty(customDirectoryName))
            {
                customDirectoryName = StringConstant.AppTempFolder;
            }

            System.IO.Directory.CreateDirectory(tempPath + customDirectoryName);


            filePath = tempPath.TrimEnd('\\') + "\\" + customDirectoryName.TrimEnd('\\') + "\\" + FileName + postFix;

            System.IO.File.WriteAllBytes(filePath, resources);

            return filePath;
        }

        internal static Stream getStream(string resourcePath)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream streamResource = assembly.GetManifestResourceStream(resourcePath);

            if (streamResource == null)
            {
                throw new Exception("خطای غیر منتظره در گرفتن Stream از فایل!");
            }
            return streamResource;
        }

        internal static string copyFileToTempFolder(Stream stream, string FileName, string customDirectoryName = "")
        {
            string tempPath = System.IO.Path.GetTempPath();
            string filePath;

            if (string.IsNullOrEmpty(customDirectoryName))
            {
                customDirectoryName = StringConstant.AppTempFolder;
            }
            System.IO.Directory.CreateDirectory(tempPath + customDirectoryName);

            //filePath = tempPath + customDirectoryName + FileName + postFix;
            filePath = tempPath.TrimEnd('\\') + "\\" + customDirectoryName.TrimEnd('\\') + "\\" + FileName;

            try
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }
                return filePath;
            }
            catch (Exception e)
            {
                throw new Exception("خطای غیر منتظره در کپی سند الگو\n" + e.Message);
            }
            throw new Exception("خطای غیر منتظره در کپی سند الگو");
        }
        internal static string copyFileToFolder(Stream stream, string FileName, string directoryName)
        {
            System.IO.Directory.CreateDirectory(directoryName);
            string filePath;
            filePath = directoryName.TrimEnd('\\') + "\\" + FileName;

            return copyFileToFolder(stream, filePath);
        }
        internal static string copyFileToFolder(Stream stream, string filePath)
        {
            //unload AttachedTemplate for replace Template
            List<Document> doucmentsAttachedToTemplate = new List<Document>();
            foreach (Document docAttached in Globals.ThisAddIn.Application.Documents)
            {
                Microsoft.Office.Interop.Word.Template attachedTemplate = docAttached.get_AttachedTemplate() as Microsoft.Office.Interop.Word.Template;
                if (attachedTemplate != null && attachedTemplate.FullName.Equals(filePath, StringComparison.OrdinalIgnoreCase))
                {
                    doucmentsAttachedToTemplate.Add(docAttached);
                    docAttached.set_AttachedTemplate("");
                }
            }

            try
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                {
                    stream.CopyTo(fileStream);
                }
                //load again AttachedTemplate
                foreach (Document docAttached in doucmentsAttachedToTemplate)
                {
                    docAttached.set_AttachedTemplate(filePath);
                }
                return filePath;
            }
            catch (Exception e)
            {
                //load again AttachedTemplate
                foreach (Document docAttached in doucmentsAttachedToTemplate)
                {
                    docAttached.set_AttachedTemplate(filePath);
                }

                throw new Exception("خطای غیر منتظره در کپی سند الگو\n" + e.Message);
            }
        }
        internal static string copyFileToFolder(byte[] resources, string FileName, string postFix, string directoryName)
        {
            System.IO.Directory.CreateDirectory(directoryName);

            string filePath = directoryName.TrimEnd('\\') + "\\" + FileName + postFix;

            System.IO.File.WriteAllBytes(filePath, resources);

            return filePath;
        }
        #endregion

        #region Image
        internal static string copyImageToTempFolder(Bitmap bitmap, string customFileName = "filename", string postFix = ".jpg")
        {
            string tempPath = System.IO.Path.GetTempPath();
            System.IO.Directory.CreateDirectory(tempPath + StringConstant.AppTempFolder);
            string filePath = tempPath + StringConstant.AppTempFolder + customFileName + postFix;
            System.IO.File.WriteAllBytes(filePath, ImageToByte(bitmap));
            return filePath;
        }

        internal static byte[] ImageToByte(System.Drawing.Image img)
        {
            ImageConverter converter = new ImageConverter();
            return (byte[])converter.ConvertTo(img, typeof(byte[]));
        }
        #endregion

        #region Header And Footer
        internal static void setHeaderChapter(Selection selection, Document doc, string ChapterNumber, string ChapterTitle)
        {
            //string defaultPersianFont = doc.Styles[StyleNames.styleNormal].Font.NameBi;

            selection.Range.PageSetup.DifferentFirstPageHeaderFooter = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            //range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false

            //selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Name = defaultPersianFont;
            selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
            selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font.Size = 14;
            //doc.Fields.Add(selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, WdFieldType.wdFieldStyleRef, "\"" + Globals.ThisAddIn.Application.ActiveDocument.Styles[WdBuiltinStyle.wdStyleHeading1].NameLocal + "\"");
            selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = "\n" + ChapterNumber + ChapterTitle;
            selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleThinThickLargeGap;
        }
        internal static void clearHeaderChapter(Selection selection)

        {
            //range.PageSetup.DifferentFirstPageHeaderFooter = (int)WdConstants.wdToggle;
            selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            //selection.Range.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            selection.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;
            //selection.Sections[1].Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = false;

            selection.Range.PageSetup.DifferentFirstPageHeaderFooter = (int)Microsoft.Office.Core.MsoTriState.msoFalse;

            //range.ClearFormatting();
            selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Borders.Enable = 0;
            selection.Range.Sections[1].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = ParagraphAndTextWrapMarks.ParagraphMarkInText;
        }

        #endregion

        #region Bookmark
        internal static bool addBookmark(Document doc, string id, Range rng)
        {
            if (rng != null)
            {
                if (doc.Bookmarks.Exists(id))
                {
                    return false;
                }
                else
                {
                    doc.Bookmarks.Add(id, rng);
                    return true;
                }
            }
            else
                return false;
        }
        internal static bool changeBookmarkContents(Document doc, string id, string newText)
        {
            if (doc.Bookmarks.Exists(id))
            {
                Range rng = doc.Bookmarks[id].Range;
                if (doc.Bookmarks[id].Range.Text == null)
                    removeBookmark(doc, id);

                rng.Text = newText;
                //deleted previous Bookmark,becuase all text has been deleted,so lets to addBookmark
                if (addBookmark(doc, id, rng))
                    return true;
                else
                    return false;
            }
            else
            {
                return false;
            }
        }
        internal static bool removeBookmark(Document doc, string id)
        {
            if (doc.Bookmarks.Exists(id))
            {
                doc.Bookmarks[id].Delete();
                return true;
            }
            else
            {
                return false;
            }

        }
        internal static bool isBookmarkExist(Document doc, string bookmarkName)
        {
            try
            {
                string s = doc.Bookmarks[bookmarkName].Range.Text;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion

        #region Content Control
        internal static void ContentControlToImage(Document doc, Microsoft.Office.Interop.Word.ContentControl cc)
        {
            float fontSize = cc.Range.Font.Size > 200 ? cc.Range.Font.Size : 300;
            System.Drawing.Font font = new System.Drawing.Font(cc.Range.Font.Name, fontSize);
            Bitmap testBitmap = new Bitmap(1000, 1000);//for get Width And Height
            Graphics testGraphics = Graphics.FromImage(testBitmap);//for get Width And Height
            System.Drawing.Size size = System.Windows.Forms.TextRenderer.MeasureText(testGraphics, cc.Range.Text, font, new System.Drawing.Size(), TextFormatFlags.NoPadding);
            int width = size.Width;
            int height = size.Height;

            using (Bitmap bitmap = new Bitmap(width, height))
            {
                using (Graphics graphic = Graphics.FromImage(bitmap))
                {
                    System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(0, 0, width, height);

                    graphic.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                    StringFormat format = new StringFormat();
                    format.LineAlignment = StringAlignment.Center;
                    format.Alignment = StringAlignment.Center;

                    graphic.DrawString(cc.Range.Text, font, System.Drawing.Brushes.Black, rectangle, format);
                    graphic.DrawRectangle(Pens.Transparent, rectangle);
                }
                try
                {
                    cc.Range.Select();
                    cc.LockContentControl = false;
                    cc.LockContents = false;
                    cc.Delete(true);
                    string tempImageUniversityPath = DedicatedFunctions.copyImageToTempFolder(bitmap, "TextBesmellah");
                    Selection selection = doc.ActiveWindow.Selection;
                    DedicatedFunctions.insertFormPicture(selection, DedicatedFunctions.copyImageToTempFolder(bitmap, "TextBesmellah"), WdWrapType.wdWrapInline);
                    DedicatedFunctions.removeFileFromSystem(tempImageUniversityPath);
                    DedicatedFunctions.changeStyleInSelection(doc.ActiveWindow.Selection, StyleNames.styleNormal);
                }
                catch (Exception)
                {
                }
            }
        }
        internal static void BesmellahToImage(Selection selection, float fontSi, string fontName, string text)
        {
            float fontSize = fontSi > 200 ? fontSi : 300;
            System.Drawing.Font font = new System.Drawing.Font(fontName, fontSize);
            Bitmap testBitmap = new Bitmap(1500, 1500);//for get Width And Height
            Graphics testGraphics = Graphics.FromImage(testBitmap);//for get Width And Height
            System.Drawing.Size size = System.Windows.Forms.TextRenderer.MeasureText(testGraphics, text, font, new System.Drawing.Size(), TextFormatFlags.NoPadding);
            int width = size.Width;
            int height = size.Height;

            int finalWidth = (int)(width * 1.3);
            int finalHeight = (int)(height * 1.3);
            int xOffset = (int)(width * 0.15);
            int yOffset = (int)(height * 0.15);


            using (Bitmap bitmap = new Bitmap(finalWidth, finalHeight))
            {
                using (Graphics graphic = Graphics.FromImage(bitmap))
                {
                    graphic.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                    StringFormat format = new StringFormat();
                    format.LineAlignment = StringAlignment.Center;
                    format.Alignment = StringAlignment.Center;
                    format.FormatFlags = StringFormatFlags.NoClip;  // <-- فقط این خط جدید

                    System.Drawing.Rectangle rectangle = new System.Drawing.Rectangle(xOffset, yOffset, width, height);
                    graphic.DrawString(text, font, System.Drawing.Brushes.Black, rectangle, format);

                }
                try
                {
                    string tempImagePath = copyImageToTempFolder(bitmap, "TextBesmellah");

                    selection.Range.Select();
                    InlineShape inlineShape = selection.InlineShapes.AddPicture(tempImagePath);

                    removeFileFromSystem(tempImagePath);

                    Shape shape = inlineShape.ConvertToShape();

                    shape.WrapFormat.Type = WdWrapType.wdWrapSquare;
                    shape.WrapFormat.Side = WdWrapSideType.wdWrapBoth;

                    shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;

                    float pageWidth = selection.PageSetup.PageWidth;
                    float pageHeight = selection.PageSetup.PageHeight;

                    shape.Left = (pageWidth - shape.Width) / 2;
                    shape.Top = (pageHeight - shape.Height) / 2;

                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                    shape.Select();
                    changeStyleInSelection(selection, StyleNames.styleNormal);
                }
                catch (Exception)
                {
                }
            }
        }

        internal static bool changeContentControlContents(Document doc, string tag, string content, bool allowCarriageReturn = false)
        {
            try
            {
                ContentControls ccs = doc.SelectContentControlsByTag(tag);
                if (ccs != null)
                {
                    foreach (Microsoft.Office.Interop.Word.ContentControl cc in ccs)
                    {
                        bool tempLockContent = cc.LockContents;

                        if (content == SettingValues.NotExist)
                        {
                            content = "";
                        }

                        try
                        {
                            cc.LockContents = false;
                            cc.Range.Text = content;
                        }
                        catch (Exception)
                        {
                            try
                            {
                                if (allowCarriageReturn)
                                {
                                    cc.Range.Text = content.Replace(ParagraphAndTextWrapMarks.ParagraphMarkInText, "").Replace("\n", "\v");
                                }
                                else
                                {
                                    cc.Range.Text = content.Replace(ParagraphAndTextWrapMarks.ParagraphMarkInText, "").Replace("\n", "\t");
                                }
                            }
                            catch (Exception)
                            {
                                return false;
                            }
                        }

                        cc.LockContents = tempLockContent;
                    }
                }
                return true;
            }
            catch (Exception)
            {

            }
            return false;
        }
        internal static bool changeContentControlFontName(Document doc, string tag, string fontName, bool allowCarriageReturn = false)
        {
            try
            {
                ContentControls ccs = doc.SelectContentControlsByTag(tag);
                if (ccs != null)
                {
                    foreach (Microsoft.Office.Interop.Word.ContentControl cc in ccs)
                    {
                        unProtectImportantText(doc, tag, true, false);
                        try
                        {
                            cc.Range.Font.Name = fontName;
                            cc.Range.Font.NameBi = fontName;
                        }
                        catch (Exception)
                        {
                            return false;
                        }

                        ProtectImportantText(doc, tag);
                        return true;
                    }
                }
            }
            catch (Exception)
            {

            }
            return false;
        }
        internal static Microsoft.Office.Interop.Word.ContentControl[] getContentControls(Document doc, string tag)
        {
            List<Microsoft.Office.Interop.Word.ContentControl> cc = new List<Microsoft.Office.Interop.Word.ContentControl>();

            ContentControls ccs = doc.SelectContentControlsByTag(tag);
            if (ccs != null)
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl item in ccs)
                {
                    cc.Add(item);
                }
            }

            return cc.ToArray();
        }
        internal static Microsoft.Office.Interop.Word.ContentControl[] getContentControls(Section section, string tag)
        {
            try
            {
                List<Microsoft.Office.Interop.Word.ContentControl> cc = new List<Microsoft.Office.Interop.Word.ContentControl>();
                foreach (Microsoft.Office.Interop.Word.ContentControl item in section.Range.ContentControls)
                {
                    if (item.Tag == tag)
                        cc.Add(item);
                }
                if (cc.Count > 0)
                    return cc.ToArray();
            }
            catch (Exception)
            {

            }
            return null;
        }
        #endregion

        #region Encryption,Decryption
        internal static string EncryptString(string plainText)
        {
            byte[] Key = EncryptionKeys.Key;
            byte[] IV = EncryptionKeys.IV;
            // Check arguments.
            if (Key == null || Key.Length <= 0)
                throw new ArgumentNullException("Key");
            if (IV == null || IV.Length <= 0)
                throw new ArgumentNullException("IV");
            byte[] encrypted;
            // Create an RijndaelManaged object
            // with the specified key and IV.

            //or ICryptoTransform 
            using (RijndaelManaged rijAlg = new RijndaelManaged())
            {
                rijAlg.Key = Key;
                rijAlg.IV = IV;

                // Create an encryptor to perform the templateXMLStream transform.
                ICryptoTransform encryptor = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV);

                // Create the streams used for encryption.
                using (MemoryStream msEncrypt = new MemoryStream())
                {
                    using (CryptoStream csEncrypt = new CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write))
                    {
                        using (StreamWriter swEncrypt = new StreamWriter(csEncrypt))
                        {

                            //Write all data to the templateXMLStream.
                            swEncrypt.Write(plainText);
                        }
                        encrypted = msEncrypt.ToArray();
                    }
                }
            }
            // Return the encrypted bytes from the memory templateXMLStream.
            //return encrypted;

            return Convert.ToBase64String(encrypted);
        }

        internal static string DecryptString(string cipherText)
        {
            try
            {

                byte[] Key = EncryptionKeys.Key;
                byte[] IV = EncryptionKeys.IV;

                if (cipherText == null || cipherText.Length == 0)
                    return null;

                // Check arguments.
                if (Key == null || Key.Length == 0)
                    throw new ArgumentNullException("Key");
                if (IV == null || IV.Length == 0)
                    throw new ArgumentNullException("IV");

                // Declare the string used to hold
                // the decrypted text.
                string plaintext = null;

                // Create an RijndaelManaged object
                // with the specified key and IV.
                using (RijndaelManaged rijAlg = new RijndaelManaged())
                {
                    rijAlg.Key = Key;
                    rijAlg.IV = IV;

                    // Create a decryptor to perform the templateXMLStream transform.
                    ICryptoTransform decryptor = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV);

                    // Create the streams used for decryption.
                    using (MemoryStream msDecrypt = new MemoryStream(Convert.FromBase64String(cipherText)))
                    {
                        using (CryptoStream csDecrypt = new CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read))
                        {
                            using (StreamReader srDecrypt = new StreamReader(csDecrypt))
                            {
                                // Read the decrypted bytes from the decrypting templateXMLStream
                                // and place them in a string.
                                plaintext = srDecrypt.ReadToEnd();
                            }
                        }
                    }
                }

                return plaintext;
            }
            catch (Exception)
            {
                return cipherText;
            }
        }
        #endregion

        #region Variables (in Document)
        internal static string getStaticVariableValue(Document doc, string name)
        {
            try
            {
                //decrypt
                string decryptedvalue;
                try
                {
                    string encryptedname = DedicatedFunctions.EncryptString(name);
                    decryptedvalue = DedicatedFunctions.DecryptString(doc.Variables[encryptedname].Value);
                }
                catch (Exception)
                {
                    decryptedvalue = doc.Variables[name].Value;
                }

                return decryptedvalue;
            }
            catch (Exception)
            {
                return SettingValues.NotExist;
            }
        }
        internal static bool isStaticVariableExist(Document doc, string name)
        {
            try
            {
                string encryptedname = DedicatedFunctions.EncryptString(name);

                string test = doc.Variables[encryptedname].Name;
                //foreach(Variable variable in doc.Variables)
                //{
                //	if(variable.Name == encryptedname)
                //	{
                //		return true;
                //	}
                //}
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        internal static void setORAddStaticVariableValue(Document doc, string name, string value, bool valueIsEncrypted = false)
        {
            //Encrypt
            string encryptedname = DedicatedFunctions.EncryptString(name);
            string encryptedvalue = value;
            if (!valueIsEncrypted)
                encryptedvalue = DedicatedFunctions.EncryptString(value);

            if (isStaticVariableExist(doc, name))
            {
                doc.Variables[encryptedname].Value = encryptedvalue;
            }
            else
            {
                doc.Variables.Add(encryptedname, encryptedvalue);
            }
        }
        internal static bool addVariable(Document doc, string name, string value)
        {
            //Encrypt
            string encryptedname = DedicatedFunctions.EncryptString(name);
            string encryptedvalue = DedicatedFunctions.EncryptString(value);

            try
            {
                doc.Variables.Add(encryptedname, encryptedvalue);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        internal static bool removeVariable(Document doc, string name)
        {
            string encryptedname = DedicatedFunctions.EncryptString(name);

            try
            {
                doc.Variables[encryptedname].Delete();
                return true;
            }
            catch (Exception)
            {
                ShowErrorMessage("اطلاعات شما در متغیر های سند اضافه نشد", (int)ErrorCodes.addVariable, StringConstant.SupportEmail);
                return false;
            }
        }

        internal static bool removeAllDocumentVariables(Document doc)
        {
            //get all VariableNames for Addin and put into one array
            string[] vfn = Enum.GetNames(typeof(VariableFieldIDs));
            string[] vpn = Enum.GetNames(typeof(PageIDs));
            string[] vin = Enum.GetNames(typeof(VariableIdentifierIDs));
            string[] vsn = Enum.GetNames(typeof(VariableServerIDs));
            string[] vdn = Enum.GetNames(typeof(VariableDocumentIDs));
            string[] vtn = Enum.GetNames(typeof(VariableTypeIDs));
            string[] vvn = Enum.GetNames(typeof(VariableVersionIDs));
            string[] von = Enum.GetNames(typeof(VariableOptionIDs));
            string[] variables = vfn.Concat(vpn).ToArray().Concat(vsn).ToArray().Concat(vdn).ToArray().Concat(vtn).ToArray().Concat(vvn).ToArray().Concat(von).ToArray().Concat(vin).ToArray();

            try
            {
                foreach (string item in variables)
                {
                    try
                    {
                        string encryptedname = DedicatedFunctions.EncryptString(item);

                        doc.Variables[encryptedname].Delete();
                    }
                    catch (Exception)
                    {
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion

        #region Tables And Update Everything

        internal static void updateChapters(Document doc)
        {
            if (DedicatedFunctions.getStaticVariableValue(doc, VariableDocumentIDs._variable_document_SharedDocument.ToString()) != "1")
            {
                Dictionary<ContentControlNames, ContentControlNames> chapterContentControlTags = new Dictionary<ContentControlNames, ContentControlNames>()
            {
                {ContentControlNames._field_Chapter1_Title ,ContentControlNames._field_Hidden_Chapter1_Title},
                {ContentControlNames._field_Chapter2_Title ,ContentControlNames._field_Hidden_Chapter2_Title},
                {ContentControlNames._field_Chapter3_Title ,ContentControlNames._field_Hidden_Chapter3_Title},
                {ContentControlNames._field_Chapter4_Title ,ContentControlNames._field_Hidden_Chapter4_Title},
                {ContentControlNames._field_Chapter5_Title ,ContentControlNames._field_Hidden_Chapter5_Title},
                {ContentControlNames._field_Chapter6_Title ,ContentControlNames._field_Hidden_Chapter6_Title},
                {ContentControlNames._field_Chapter7_Title ,ContentControlNames._field_Hidden_Chapter7_Title},
                {ContentControlNames._field_Chapter8_Title ,ContentControlNames._field_Hidden_Chapter8_Title},
                {ContentControlNames._field_Chapter9_Title ,ContentControlNames._field_Hidden_Chapter9_Title},
                {ContentControlNames._field_Chapter10_Title ,ContentControlNames._field_Hidden_Chapter10_Title},
            };
                foreach (var chapterContentControlTag in chapterContentControlTags)
                {
                    ContentControls ccs = doc.SelectContentControlsByTag(chapterContentControlTag.Key.ToString());
                    ContentControls ccs2 = doc.SelectContentControlsByTag(chapterContentControlTag.Value.ToString());
                    if (ccs != null && ccs2 != null)
                    {
                        foreach (ContentControl chapter in ccs)
                        {
                            foreach (ContentControl hiddenChapter in ccs2)
                            {
                                hiddenChapter.LockContents = false;
                                hiddenChapter.Range.Text = chapter.Range.Text.Replace(ParagraphAndTextWrapMarks.TextWrapMarkInText, " ").Replace(ParagraphAndTextWrapMarks.ParagraphMarkInText, "");
                                hiddenChapter.LockContents = true;
                            }
                        }
                    }
                }
            }
        }
        internal static void updateTables(Document doc, Selection selection, AccessType accessType)//TODO: Fix This, edit styles,
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("بروزرسانی فهرست و جداول");
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            //string defaultPersianFont = doc.Styles[StyleNames.styleNormal].Font.NameBi;

            Range previousRng = selection.Range;

            updateChapters(doc);

            doc.Fields.Update();
            doc.PrintPreview();
            doc.ClosePrintPreview();
            foreach (Index index in doc.Indexes)
            {
                index.Update();
            }

            //bool isTableExist = false;
            //bool isFigExist = false;
            //bool isFormulaExist = false;
            //foreach(TableOfFigures table in doc.TablesOfFigures)
            //{
            //	if(table.Caption == Constants.CaptionLabels.captionFigure)
            //	{
            //		isFigExist = true;
            //	}
            //	if(table.Caption == Constants.CaptionLabels.captionTable)
            //	{
            //		isTableExist = true;
            //	}
            //	if(table.Caption == Constants.CaptionLabels.captionFormula)
            //	{
            //		isFormulaExist = true;
            //	}
            //	table.Update();
            //	//table.Range.Font.Name = defaultPersianFont;
            //	table.Range.Font.Size = 12;
            //}

            //TODO:Check again  rebuildSpecialTables
            //if(doc.TablesOfContents.Count < 1 && getStaticVariableValue(doc , PageIDs._page_TableOfContents.ToString()) != SettingValues.NotExist)
            //{
            //	rebuildSpecialTables(doc , selection , TableTypes.TableOfContents , TitleOfTables.TableOfContents);
            //}
            //if(!isFigExist && getStaticVariableValue(doc , PageIDs._page_TableOfFigures.ToString()) != SettingValues.NotExist)
            //{
            //	rebuildSpecialTables(doc , selection , TableTypes.TableOfFigures , TitleOfTables.TableOfFigures_Figures , Constants.CaptionLabels.captionFigure , Constants.CaptionLabels.Nothing);
            //}
            //if(!isTableExist && getStaticVariableValue(doc , PageIDs._page_TableOfTables.ToString()) != SettingValues.NotExist)
            //{
            //	if(getStaticVariableValue(doc , PageIDs._page_TableOfFigures.ToString()) != SettingValues.NotExist)
            //	{
            //		rebuildSpecialTables(doc , selection , TableTypes.TableOfFigures , TitleOfTables.TableOfFigures_Tables , Constants.CaptionLabels.captionTable , Constants.CaptionLabels.captionFigure);
            //	}
            //	else
            //	{
            //		rebuildSpecialTables(doc , selection , TableTypes.TableOfFigures , TitleOfTables.TableOfFigures_Tables , Constants.CaptionLabels.captionTable , Constants.CaptionLabels.Nothing);
            //	}
            //}


            //if(!isFormulaExist && getStaticVariableValue(doc , VariablePageIDs._variable_page_Abbreviations.ToString()) != SettingValues.NotExist)
            //{
            //	if(getStaticVariableValue(doc , VariablePageIDs._variable_page_TableOfTables.ToString()) != SettingValues.NotExist)
            //	{
            //		rebuildSpecialTables(doc , selection , TableTypes.TableOfFigures , TitleOfTables.TableOfFigures_Formula , Constants.CaptionLabels.captionFormula , Constants.CaptionLabels.captionTable , StyleNames.styleSaNaFormula);
            //	}
            //	else if(getStaticVariableValue(doc , VariablePageIDs._variable_page_TableOfFigures.ToString()) != SettingValues.NotExist)
            //	{
            //		rebuildSpecialTables(doc , selection , TableTypes.TableOfFigures , TitleOfTables.TableOfFigures_Formula , Constants.CaptionLabels.captionFormula , Constants.CaptionLabels.captionFigure , StyleNames.styleSaNaFormula);
            //	}
            //	else
            //	{
            //		rebuildSpecialTables(doc , selection , TableTypes.TableOfFigures , TitleOfTables.TableOfFigures_Formula , Constants.CaptionLabels.captionFormula , Constants.CaptionLabels.Nothing , StyleNames.styleSaNaFormula);
            //	}
            //}

            foreach (TableOfContents table in doc.TablesOfContents)
            {
                //table.Range.Select();
                table.Update();
                //table.Range.Font.Name = defaultPersianFont;
                table.Range.Font.Size = 12;
            }
            foreach (TableOfFigures table in doc.TablesOfFigures)
            {
                table.Update();

                if (table.Range.Text.ToLower().Trim() == SpecialTablesMessage.OldTableOfFiguresNoItemMessage.ToLower().Trim())
                {
                    string previousText = SpecialTablesMessage.OldTableOfFiguresNoItemMessage.ToLower().Trim();
                    table.Range.Select();
                    selection.EndKey();
                    selection.MoveLeft(WdUnits.wdCharacter, 1);
                    for (int i = 0; i < previousText.Length - 1; i++)
                    {
                        selection.TypeBackspace();
                    }
                    selection.TypeText(SpecialTablesMessage.NewTableOfFiguresNoItemMessage.Replace(SpecialTablesMessage.replace1, table.Caption));

                }
                //if (table.Caption == TextCaptions.captionFigure)
                //{
                //    table.Range.Font.Name = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleTableOfFigures].Font.Name;
                //    table.Range.Font.Size = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleTableOfFigures].Font.Size;
                //}
                //if (table.Caption == TextCaptions.captionTable)
                //{
                //    table.Range.Font.Name = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleSaNaTable].Font.Name;
                //    table.Range.Font.Size = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleSaNaTable].Font.Size;
                //}
                //if (table.Caption == TextCaptions.captionFormula)
                //{
                //    table.Range.Font.Name = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleSaNaFormula].Font.Name;
                //    table.Range.Font.Size = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleSaNaFormula].Font.Size;
                //}
            }

            updateBibliography(doc, selection, accessType);

            previousRng.Select();
            doc.ActiveWindow.ScrollIntoView(previousRng);
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }
        internal static void updateBibliography(Document doc, Selection selection, AccessType accessType)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            Range previousRng = selection.Range;

            foreach (Source item in Globals.ThisAddIn.Application.ActiveDocument.Bibliography.Sources)
            {
                char firstChar = item.Field["Author"].Trim()[0];
                char firstCharWithoutZeroWidthCharacter = item.Field["Author"].Trim().Replace("​", "")[0];

                if (DedicatedFunctions.IsPersianChar(firstChar))
                    item.Field["Author"] = "​" + item.Field["Author"];

                if (DedicatedFunctions.IsPersianChar(firstCharWithoutZeroWidthCharacter))
                    item.Field["LCID"] = ((int)WdLanguageID.wdPersian).ToString();//fa-IR
                else if (DedicatedFunctions.IsEnglishChar(firstCharWithoutZeroWidthCharacter))
                    item.Field["LCID"] = ((int)WdLanguageID.wdEnglishUS).ToString();//en-US
            }

            //TODO: disabled Bibliography
            //foreach(Microsoft.Office.Interop.Word.Source item in doc.Bibliography.Sources)
            //{
            //	if(item.Cited)
            //	{
            //		item.Field["LCID"] = ((int)WdLanguageID.wdPersian).ToString();
            //	}
            //	else
            //	{
            //		item.Field["LCID"] = ((int)WdLanguageID.wdNoProofing).ToString();
            //	}
            //}

            foreach (Field item in doc.Fields)
            {
                if (item.Type == WdFieldType.wdFieldBibliography)
                {
                    //TODO: disabled Bibliography
                    //if(!item.Code.Text.Contains("\\f "))
                    //	item.Code.Text += " \\f " + (int)WdLanguageID.wdPersian;


                    //if (!item.Code.Text.Contains("\\l "))
                    //    item.Code.Text += " \\l " + (int)WdLanguageID.wdPersian;
                    item.Update();

                    if (accessType == AccessType.AccessGranted)
                    {
                        if (DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_BibliographyStyle.ToString()) == "IEEE")
                        {
                            foreach (Table item2 in item.Result.Tables)
                            {
                                item2.TableDirection = WdTableDirection.wdTableDirectionRtl;
                                item2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                                item2.Range.ParagraphFormat.FirstLineIndent = 0;
                            }
                        }
                    }

                    if (item.Result.Text.ToLower().Trim() == SpecialTablesMessage.OldBibliographyNoItemMessage.ToLower().Trim())
                    {
                        string previousText = SpecialTablesMessage.OldBibliographyNoItemMessage.ToLower().Trim();
                        item.Result.Select();

                        selection.EndKey();
                        selection.MoveLeft(WdUnits.wdCharacter, 1);
                        for (int i = 0; i < previousText.Length - 1; i++)
                        {
                            selection.TypeBackspace();
                        }
                        selection.TypeText(SpecialTablesMessage.NewBibliographyNoItemMessage);
                    }
                }
            }
            previousRng.Select();
            doc.ActiveWindow.ScrollIntoView(previousRng);
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }

        /// <summary>
        /// rebuild TablesOfContents or TablesOfFigures if deleted by User
        /// </summary>
        /// <param name="tableType">The table type we want to rebuild</param>
        /// <param name="label">The caption that is written for the tables at the top of the page, for example "فهرست مطالب"</param>
        /// <param name="textCaptions">The caption that is displayed in TableOfFigures tables</param>
        /// <param name="previousTextCaptions">The previous caption that is displayed in TableOfFigures tables (Nothing means that the table is of TablesOfContents type)</param>
        internal static void rebuildSpecialTables(Document doc, Selection selection, TableTypes tableType, string label, string textCaptions = Constants.CaptionLabels.captionFormula, string previousTextCaptions = Constants.CaptionLabels.captionFigure)
        {
            if (tableType == TableTypes.TableOfContents)
            {
                if (getStaticVariableValue(doc, PageIDs._page_PersianAbstract.ToString()) != SettingValues.NotExist)//TODO:It works only when the >PagePersianAbstract< is available in the thesis
                {
                    ContentControl[] abstractContentControl = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_Fa.ToString());
                    if (abstractContentControl != null && abstractContentControl.Length != 0)
                    {
                        Range rangeAbstract = abstractContentControl[0].Range;
                        if (rangeAbstract != null)
                        {
                            rangeAbstract.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToNext, 1).Select();
                            doc.Bookmarks["\\Page"].Select();
                            selection.Delete();

                            changeStyleInSelection(selection, StyleNames.styleNormal);
                            changeParagraphAlignment(selection, WdParagraphAlignment.wdAlignParagraphCenter);
                            insertSpecialTable(doc, selection, TableTypes.TableOfContents, label);
                            insertPageBreak(doc, selection.Range);
                        }
                    }
                }
            }
            else if (tableType == TableTypes.TableOfFigures)//TODO:It works only when the >tableOfContents< is available in the thesis and it is >mandatory<
            {
                int whichTableOfFigures = 1;
                if (previousTextCaptions == Constants.CaptionLabels.Nothing && doc.TablesOfContents.Count > 0)
                {
                    doc.TablesOfContents[1].Range.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToNext, 1).Select();
                    //selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToNext, 1);
                    //doc.TablesOfContents[1].Range.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToNext, 1).Select();
                }
                else
                {
                    if (doc.TablesOfFigures.Count > 0)
                    {
                        foreach (TableOfFigures item in doc.TablesOfFigures)
                        {
                            if (item.Caption == previousTextCaptions)
                            {
                                break;
                            }
                            whichTableOfFigures++;
                        }
                    }
                    else
                        return;

                    //DedicatedFunctions.ShowMessage(doc.TablesOfFigures[whichTableOfFigures].Range.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToNext, 1).Text);
                    doc.TablesOfFigures[whichTableOfFigures].Range.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToNext, 1).Select();
                    doc.TablesOfFigures[whichTableOfFigures].Range.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToNext, 1).Select();
                }

                if (selection.Find.Execute(label))
                {
                    doc.Bookmarks["\\Page"].Select();
                    selection.Delete();
                }

                changeStyleInSelection(selection, StyleNames.styleNormal);
                changeParagraphAlignment(selection, WdParagraphAlignment.wdAlignParagraphCenter);
                insertSpecialTable(doc, selection, TableTypes.TableOfFigures, label, textCaptions, "A");
                insertPageBreak(doc, selection.Range);
            }
        }
        internal static void selectTableOfContentsStyle(Document doc, Selection selection, int withIndent)
        {
            //if(doc.TablesOfContents.Count < 1 && DedicatedFunctions.getStaticVariableValue(doc , PageIDs._page_TableOfContents.ToString()) != SettingValues.NotExist)
            //{
            //	DedicatedFunctions.rebuildSpecialTables(doc , selection , TableTypes.TableOfContents , TitleOfTables.TableOfContents);
            //}
            changeIndentTOCStyles(doc, withIndent);
        }
        internal static void changeIndentTOCStyles(Document doc, int withIntent)
        {
            Microsoft.Office.Interop.Word.Application application = Globals.ThisAddIn.Application;
            try
            {

                List<WdBuiltinStyle> listStyles = new List<WdBuiltinStyle>()
              {
                  //StyleNames.styleTOC1,
                  StyleNames.styleTOC2,
                  StyleNames.styleTOC3,
                  StyleNames.styleTOC4,
                  StyleNames.styleTOC5,
                  StyleNames.styleTOC6,
                  StyleNames.styleTOC7,
                  StyleNames.styleTOC8,
                  StyleNames.styleTOC9
              };

                foreach (WdBuiltinStyle selectedStyle in listStyles)
                {
                    Style style;

                    try
                    {
                        style = doc.Styles[selectedStyle];
                    }
                    catch (System.Exception)
                    {
                        style = doc.Styles.Add(doc.Styles[selectedStyle].NameLocal, WdStyleType.wdStyleTypeParagraph);
                    }

                    if (withIntent == 1)
                    {
                        style.ParagraphFormat.FirstLineIndent = application.MillimetersToPoints(5);

                        if (selectedStyle == StyleNames.styleTOC1)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(0f);
                        else if (selectedStyle == StyleNames.styleTOC2)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(4.23f);
                        else if (selectedStyle == StyleNames.styleTOC3)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(8.5f);
                        else if (selectedStyle == StyleNames.styleTOC4)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(12.7f);
                        else if (selectedStyle == StyleNames.styleTOC5)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(16.9f);
                        else if (selectedStyle == StyleNames.styleTOC6)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(21.2f);
                        else if (selectedStyle == StyleNames.styleTOC7)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(25.4f);
                        else if (selectedStyle == StyleNames.styleTOC8)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(29.6f);
                        else if (selectedStyle == StyleNames.styleTOC9)
                            style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(33.9f);
                    }
                    else
                    {
                        style.ParagraphFormat.FirstLineIndent = application.MillimetersToPoints(0);
                        style.ParagraphFormat.LeftIndent = application.MillimetersToPoints(0);
                    }
                    updateTOCStyle(doc, selectedStyle);
                }
            }
            catch (System.Exception e)
            {
                DedicatedFunctions.ShowErrorMessage(e.Message);
            }

        }


        internal static void updateTOCStyle(Document doc, object styleName)
        {
            try
            {

                Style style = doc.Styles[styleName];

                //Font(Latin)
                style.Font.ColorIndex = WdColorIndex.wdBlack;

                //Font(Complex)
                style.Font.ColorIndexBi = WdColorIndex.wdBlack;

                style.AutomaticallyUpdate = true;
            }
            catch (System.Exception e)
            {
                DedicatedFunctions.ShowErrorMessage(e.Message);
            }

        }

        #endregion

        #region Protect and Unprotect Texts
        internal static void ProtectImportantText(ContentControl cc, string content, string title, string placeHolderText, bool lockContentControl, bool lockContents, bool allowCarriageReturn, bool mustBeEmpty)
        {

            //cc.SetPlaceholderText(Text: placeHolderText);

            if (cc.Tag.Contains(InitialContentControls.initialHidden))
                cc.Appearance = WdContentControlAppearance.wdContentControlHidden;
            else
                cc.Appearance = WdContentControlAppearance.wdContentControlBoundingBox;
            cc.Title = title;
            cc.Temporary = false;
            cc.LockContents = false;

            try
            {
                if (cc.Type != WdContentControlType.wdContentControlPicture && content != Constants.Dictionaries.iconMark && (content != "" || mustBeEmpty))
                    cc.Range.Text = content;
            }
            catch (Exception)
            {
                try
                {
                    if (allowCarriageReturn)
                    {
                        if (cc.Type != WdContentControlType.wdContentControlPicture && content != Constants.Dictionaries.iconMark && content != "")
                            cc.Range.Text = content.Replace(ParagraphAndTextWrapMarks.ParagraphMarkInText, "").Replace("\n", "\v");
                    }
                    else
                    {
                        if (cc.Type != WdContentControlType.wdContentControlPicture && content != Constants.Dictionaries.iconMark && content != "")
                            cc.Range.Text = content.Replace(ParagraphAndTextWrapMarks.ParagraphMarkInText, "").Replace("\n", "\t");
                    }
                }
                catch (Exception e)
                {
                    throw e;
                }
            }
            cc.LockContentControl = lockContentControl;
            cc.LockContents = lockContents;
        }
        internal static bool ProtectImportantText(Document doc, string tag)
        {
            try
            {
                ContentControls ccs = doc.SelectContentControlsByTag(tag);
                if (ccs != null)
                {
                    foreach (Microsoft.Office.Interop.Word.ContentControl cc in ccs)
                    {
                        if (cc.Tag == tag)
                        {
                            cc.LockContentControl = true;
                            cc.LockContents = true;
                            cc.Appearance = WdContentControlAppearance.wdContentControlBoundingBox;
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception)
            {
                return false;
            }
        }
        internal static bool unProtectingImportants(Document doc, bool removeContentControl)
        {
            List<string> contentControlList = new List<string>() { };
            try
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl item in doc.ContentControls)
                {
                    if (removeContentControl)
                    {
                        item.LockContentControl = false;
                        item.LockContents = false;
                        contentControlList.Add(item.Tag);
                    }
                    else
                    {
                        item.LockContentControl = false;
                        item.LockContents = false;
                    }
                }
                if (removeContentControl)
                {
                    foreach (string item in contentControlList)
                    {
                        ContentControls ccs = doc.SelectContentControlsByTag(item);
                        if (ccs != null)
                        {
                            foreach (Microsoft.Office.Interop.Word.ContentControl item2 in ccs)
                            {
                                item2.Delete(false);
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        internal static bool unProtectingContentControlsSection(Section section)
        {
            try
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl item in section.Range.ContentControls)
                {
                    item.LockContentControl = false;
                    item.LockContents = false;
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        internal static bool unProtectImportantText(Document doc, string tag, bool LockContentControl, bool LockContents, bool removeContentControl = false)
        {
            try
            {
                ContentControls ccs = doc.SelectContentControlsByTag(tag);
                if (ccs != null)
                {
                    foreach (Microsoft.Office.Interop.Word.ContentControl item in ccs)
                    {
                        if (removeContentControl)
                        {
                            item.LockContentControl = false;
                            item.LockContents = false;
                            item.Delete(false);
                            return true;
                        }
                        else
                        {
                            item.LockContentControl = LockContentControl;
                            item.LockContents = LockContents;
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return false;
        }
        internal static bool unProtectContentControlEditing(Document doc, string tag)
        {
            try
            {
                ContentControls ccs = doc.SelectContentControlsByTag(tag);
                if (ccs != null)
                {
                    foreach (Microsoft.Office.Interop.Word.ContentControl item in ccs)
                    {
                        item.LockContents = false;
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }
            return false;
        }
        #endregion

        #region Footnote
        internal static Range insertFootnote(Document doc, Selection selection, AccessType accessType, KeyboardLanguage language)
        {
            try
            {
                //string defaultPersianFont = doc.Styles[StyleNames.styleNormal].Font.NameBi;
                //string defaultEnglishFont = doc.Styles[StyleNames.styleNormal].Font.Name;

                selection.FootnoteOptions.Location = WdFootnoteLocation.wdBottomOfPage;
                selection.FootnoteOptions.NumberingRule = WdNumberingRule.wdRestartPage;
                selection.FootnoteOptions.StartingNumber = 1;
                selection.FootnoteOptions.NumberStyle = WdNoteNumberStyle.wdNoteNumberStyleArabic;
                selection.EndnoteOptions.Location = WdEndnoteLocation.wdEndOfSection;

                DedicatedFunctions.changeKeyboardLanguage(language);
                Range rng;
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                Footnote footnote = selection.Footnotes.Add(selection.Range, "");
                rng = footnote.Range;

                if (accessType == AccessType.AccessGranted)
                {
                    DedicatedFunctions.setNumberFootnoteStyle(doc, selection, footnote,
                        int.Parse(DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_FootnoteNumberStyle.ToString())));
                }

                Globals.ThisAddIn.Application.ScreenUpdating = true;

                DedicatedFunctions.changeKeyboardLanguage(language);
                if (language == KeyboardLanguage.Persian)
                {
                    selection.RtlPara();
                    //selection.Font.NameBi = defaultPersianFont;
                }
                else
                {
                    selection.LtrPara();
                    //selection.Font.Name = defaultEnglishFont;
                }
                return rng;
            }
            catch (Exception)
            {
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                DedicatedFunctions.ShowErrorMessage("موقعیت نشانگر مناسب نیست. نشانگر را بعد از کلمه ای که میخواهید پانویس شود قرار دهید.");
                return null;
            }

            //selection.MoveLeft(WdUnits.wdCharacter, 1);
            //selection.TypeText("-");
            //selection.TypeText(" ");

            //selection.MoveRight(WdUnits.wdCharacter, 1);
            //selection.TypeText("-");
            //selection.TypeText(" ");
        }
        internal static void selectFootnoteStyle(Document doc, Selection selection, string indexInVariable)
        {
            try
            {
                int index = -1;
                if (indexInVariable != SettingValues.NotExist && indexInVariable != "")
                {
                    try
                    {
                        index = int.Parse(indexInVariable);
                    }
                    catch (Exception)
                    {
                    }
                }

                if (index == 0)
                {
                    selection.FootnoteOptions.LayoutColumns = 1;
                }
                else if (index == 1)
                {
                    selection.FootnoteOptions.LayoutColumns = 2;
                }
                else
                {
                    if (index == -1)
                    {
                        DedicatedFunctions.setORAddStaticVariableValue(doc,
                            VariableOptionIDs._variable_option_FootnoteNumberStyle.ToString(), "0");
                        selectFootnoteStyle(doc, selection, "0");
                    }
                }
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("در تنظیم نوع پانویس مشکلی پیش آمده است", email: StringConstant.SupportEmail);
            }
        }
        internal static void selectFootnoteSeparator(Document doc, string indexInVariable)
        {
            try
            {
                int index = -1;
                if (indexInVariable != SettingValues.NotExist && indexInVariable != "")
                {
                    try
                    {
                        index = int.Parse(indexInVariable);
                    }
                    catch (Exception)
                    {
                    }
                }

                if (doc.Footnotes.Count > 0)
                {
                    doc.Footnotes.Separator.Delete();
                    if (index == 1)//Left
                    {
                        doc.Footnotes.Separator.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        doc.Footnotes.Separator.ParagraphFormat.LeftIndent = Globals.ThisAddIn.Application.InchesToPoints(3);
                    }
                    else if (index == 2)//Right
                    {
                        doc.Footnotes.Separator.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        doc.Footnotes.Separator.ParagraphFormat.RightIndent = Globals.ThisAddIn.Application.InchesToPoints(3);
                    }
                    else if (index == 3)//Center
                    {
                        doc.Footnotes.Separator.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                        doc.Footnotes.Separator.ParagraphFormat.LeftIndent = Globals.ThisAddIn.Application.InchesToPoints(1);
                        doc.Footnotes.Separator.ParagraphFormat.RightIndent = Globals.ThisAddIn.Application.InchesToPoints(1);
                    }
                    else if (index == 4)//Continuous
                    {
                        doc.Footnotes.Separator.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleSingle;
                    }
                    else//Default
                    {
                        if (index == -1)
                        {
                            DedicatedFunctions.setORAddStaticVariableValue(doc,
                                VariableOptionIDs._variable_option_FootnoteSeparatorType.ToString(), "0");
                        }

                        index = 0;
                        doc.Footnotes.ResetSeparator();
                    }
                }
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("در تنظیم خط جدا کننده پانویس مشکلی پیش آمده است", email: StringConstant.SupportEmail);
            }
        }
        internal static void setAllNumberFootnoteStyle(Document doc, Selection selection, int setToNormal)
        {
            if (doc.Footnotes.Count > 0)
            {
                foreach (Footnote item in doc.Footnotes)
                {
                    setNumberFootnoteStyle(doc, selection, item, setToNormal);
                }
            }
        }
        internal static bool setNumberFootnoteStyle(Document doc, Selection selection, Footnote footnote, int setToNormal)
        {
            try
            {
                if (setToNormal == 0)//Default
                {
                    footnote.Range.Select();
                    selection.HomeKey();
                    selection.Expand(WdUnits.wdCharacter);
                    selection.Font.Superscript = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                    selection.Expand(WdUnits.wdCharacter);
                    selection.Font.Superscript = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
                    if (selection.Text != "-")
                    {
                        if (selection.Text != " ")
                            selection.Collapse(WdCollapseDirection.wdCollapseStart);
                        selection.TypeText(" ");
                    }
                    else
                    {
                        selection.TypeBackspace();
                        selection.Expand(WdUnits.wdCharacter);
                        if (selection.Text != " ")
                            selection.Collapse(WdCollapseDirection.wdCollapseStart);
                        selection.TypeText(" ");
                    }
                    return true;
                }
                else if (setToNormal == 1)//1=Normal
                {
                    footnote.Range.Select();
                    selection.HomeKey();
                    selection.Expand(WdUnits.wdCharacter);
                    selection.Font.Superscript = (int)Microsoft.Office.Core.MsoTriState.msoFalse;
                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                    selection.Expand(WdUnits.wdCharacter);
                    if (selection.Text != "-")
                    {
                        if (selection.Text != " ")
                            selection.Collapse(WdCollapseDirection.wdCollapseStart);
                        selection.TypeText("- ");
                    }
                    else
                    {
                        selection.Collapse(WdCollapseDirection.wdCollapseEnd);
                        selection.Expand(WdUnits.wdCharacter);
                        if (selection.Text != " ")
                            selection.Collapse(WdCollapseDirection.wdCollapseStart);
                        selection.TypeText(" ");
                    }
                    return true;
                }
                else
                {
                    if (setToNormal == -1)
                    {
                        DedicatedFunctions.setORAddStaticVariableValue(doc,
                            VariableOptionIDs._variable_option_FootnoteNumberStyle.ToString(), "0");
                        return setNumberFootnoteStyle(doc, selection, footnote, 0);
                    }
                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
        #endregion

        #region get Unique HardWare ID (May not correct)
        internal string GetMachineGuid()
        {
            string location = @"SOFTWARE\Microsoft\Cryptography";
            string name = "MachineGuid";

            using (RegistryKey localMachineX64View =
                RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                using (RegistryKey rk = localMachineX64View.OpenSubKey(location))
                {
                    if (rk == null)
                        throw new KeyNotFoundException(
                            string.Format("Key Not Found: {0}", location));

                    object machineGuid = rk.GetValue(name);
                    if (machineGuid == null)
                        throw new IndexOutOfRangeException(
                            string.Format("Index Not Found: {0}", name));

                    return machineGuid.ToString();
                }
            }
        }
        internal static string getUUID()//Recommended
        {
            ManagementObjectSearcher mos = new ManagementObjectSearcher("SELECT UUID FROM Win32_ComputerSystemProduct");
            ManagementObjectCollection mbsList = mos.Get();
            string systemId = string.Empty;
            foreach (ManagementBaseObject mo in mbsList)
            {
                systemId = mo["UUID"] as string;
            }
            return systemId;
        }
        #endregion

        #region Accesibility
        internal enum AccessType
        {
            AccessGranted,
            AccessGranted_Administrator,
            AccessDenied_NotExistInWorksapce,
            AccessDenied_NoHasAccessing,
            AccessDenied_OpenedAsReadOnly,
        }

        internal static AccessType hasAccess(Document doc)
        {
            if (!doc.ReadOnly)
            {

                if (DedicatedFunctions.getStaticVariableValue(doc, VariableIdentifierIDs._variable_id_GUID.ToString()) == StringConstant.GUID &&
                    (DedicatedFunctions.getStaticVariableValue(doc, VariableTypeIDs._variable_type_Document.ToString()) != SettingValues.NotExist ||
                    DedicatedFunctions.getStaticVariableValue(doc, VariableTypeIDs._variable_type_Document.ToString()) != ""))
                {
                    //if (Path.GetFullPath(new FileInfo(doc.FullName).Directory.FullName).TrimEnd('\\')
                    //    .Contains(Path.GetFullPath(Properties.Settings.Default.WorkSpaceDirectory).TrimEnd('\\')))
                    //{
                    //    return AccessType.AccessGranted;
                    //}
                    //else
                    //{
                    //    return AccessType.AccessDenied_NotExistInWorksapce;
                    //}
                    return AccessType.AccessGranted;
                }
                else if (StringConstant.AdministratorAccounts.Contains(Properties.Settings.Default.Mobile))
                {
                    return AccessType.AccessGranted_Administrator;
                }
                else
                {
                    return AccessType.AccessDenied_NoHasAccessing;
                }
            }
            else
                return AccessType.AccessDenied_OpenedAsReadOnly;
        }
        #endregion

        #region Task Pane
        internal static Microsoft.Office.Tools.CustomTaskPane getTaskPane(Document doc, string taskPaneTitle)
        {
            try
            {
                foreach (Microsoft.Office.Tools.CustomTaskPane item in Globals.ThisAddIn.CustomTaskPanes)
                {
                    if (item.Title == taskPaneTitle && item.Window == doc.ActiveWindow)
                    {
                        return item;
                    }
                }
            }
            catch (Exception e)
            {
                if (e.HResult == -2146232798)
                {
                    // System.ObjectDisposedException: 'Cannot access a disposed object.'
                }
            }
            return null;
        }

        internal static CustomTaskPane AddTaskPane(System.Windows.Forms.UserControl uc, Document doc, string taskPaneTitle)
        {
            DedicatedFunctions.RemoveOrphanedTaskPanes();
            return Globals.ThisAddIn.CustomTaskPanes.Add(uc, taskPaneTitle, doc.ActiveWindow);
            //Error: System.InvalidCastException: 'Unable to cast COM object of type 'System.__ComObject' to interface type 'Microsoft.VisualStudio.Tools.Office.Runtime.Interop.ICustomTaskPaneSite'. This operation failed because the QueryInterface call on the COM component for the interface with IID '{3CA8CD11-274A-41B6-A999-28562DAB3AA2}' failed due to the following error: No such interface supported (Exception from HRESULT: 0x80004002 (E_NOINTERFACE)).'
        }
        internal static void RemoveAllTaskPanes(Document doc)
        {
            List<Microsoft.Office.Tools.CustomTaskPane> ctp = new List<Microsoft.Office.Tools.CustomTaskPane>() { };

            foreach (var item in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (item.Window == doc.ActiveWindow)
                {
                    ctp.Add(item);
                }
            }
            foreach (var item in ctp)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(item);
            }
        }
        internal static void RemoveTaskPane(Document doc, string taskPaneTitle)
        {
            List<Microsoft.Office.Tools.CustomTaskPane> ctp = new List<Microsoft.Office.Tools.CustomTaskPane>() { };

            foreach (var item in Globals.ThisAddIn.CustomTaskPanes)//also remove duplicated TaskPanes
            {
                if (item.Title == taskPaneTitle && item.Window == doc.ActiveWindow)
                {
                    ctp.Add(item);
                }
            }
            foreach (var item in ctp)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(item);
            }
        }
        internal static void RemoveOrphanedTaskPanes()
        {
            List<Microsoft.Office.Tools.CustomTaskPane> ctp = new List<Microsoft.Office.Tools.CustomTaskPane>() { };

            foreach (var item in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (item.Window is null)
                {
                    ctp.Add(item);
                }
            }
            foreach (var item in ctp)
            {
                Globals.ThisAddIn.CustomTaskPanes.Remove(item);
            }
        }
        #endregion

        #region using OpenXml
        internal static void InitializeOpenXml()
        {
            try
            {

                String[] files = Directory.GetFiles(Properties.Settings.Default.WorkSpaceDirectory);

                foreach (string filePath in files)
                {
                    FileInfo fileInfo = new FileInfo(filePath);
                    if (fileInfo.Extension.ToLower() == ".docx" || fileInfo.Extension.ToLower() == ".docm")
                    {
                        try
                        {
                            DocumentFormat.OpenXml.Packaging.WordprocessingDocument document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(filePath, false);
                            var contentControls = document?.MainDocumentPart?.Document?.Body;
                            document.Dispose();
                            break;
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            catch (Exception e)
            {
                if (e.HResult != -2146233040)
                {
                    DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای در مقدار دهی اولیه به وجود آمده" + "\nپیغام خطا:\n" + e.Message, email: StringConstant.SupportEmail);
                }
            }
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.DocumentVariables GetVariables(DocumentFormat.OpenXml.Packaging.WordprocessingDocument document)
        {
            DocumentFormat.OpenXml.Wordprocessing.DocumentVariables documentVariables = (DocumentFormat.OpenXml.Wordprocessing.DocumentVariables)document.MainDocumentPart.DocumentSettingsPart.Settings.Elements<DocumentFormat.OpenXml.Wordprocessing.DocumentVariables>().FirstOrDefault();

            if (documentVariables != null)
            {
                return documentVariables;
            }
            return null;
        }
        internal static string GetVariableValue(string name, DocumentFormat.OpenXml.Wordprocessing.DocumentVariables documentVariables)
        {
            // check if the variables are not null
            if (documentVariables != null)
            {
                string encryptName = DedicatedFunctions.EncryptString(name);
                string value = documentVariables.Elements<DocumentFormat.OpenXml.Wordprocessing.DocumentVariable>().Where(v => v.Name == encryptName).FirstOrDefault()?.Val.Value;
                return DedicatedFunctions.DecryptString(value);
            }
            return null;
        }

        internal static IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SdtBlock> GetContentControls(DocumentFormat.OpenXml.Packaging.WordprocessingDocument document)
        {
            var contentControls = document.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.SdtBlock>();
            if (contentControls != null)
            {
                return contentControls;
            }
            return null;
        }
        internal static string GetContentControlByTag(string tag, IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SdtBlock> contentControls)
        {
            //var contentControls = document.MainDocumentPart.Document.Body.Descendants<SdtBlock>().Where(sdt => sdt.SdtProperties.Elements<Tag>().FirstOrDefault().Val == tag)?.FirstOrDefault();

            var contentControl = contentControls.Where(r => r.SdtProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Tag>()?.Val.Value.Contains(tag) == true).FirstOrDefault();


            if (contentControl != null)
            {
                return contentControl.SdtContentBlock?.FirstOrDefault().InnerText;
            }
            else
                return "محتوا یافت نشد!";
        }

        internal static DocumentFormat.OpenXml.CustomProperties.Properties GetCustomProperties(string documentPath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(documentPath, false))
            {
                CustomFilePropertiesPart customProps = wordDocument.CustomFilePropertiesPart;

                if (customProps == null)
                    return null;
                else
                    return customProps.Properties;
            }
        }
        internal static DocumentFormat.OpenXml.CustomProperties.Properties GetCustomPropertiesWithStream(string documentPath)
        {
            using (FileStream fs = new FileStream(documentPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Open a WordprocessingDocument for read-only access based on a templateXMLStream.
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(fs, false))
                {
                    CustomFilePropertiesPart customProps = wordDocument.CustomFilePropertiesPart;

                    if (customProps == null)
                        return null;
                    else
                        return customProps.Properties;
                }
            }
        }

        internal static DocumentFormat.OpenXml.Wordprocessing.Body GetBodyPart(string documentPath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(documentPath, false))
            {
                DocumentFormat.OpenXml.Wordprocessing.Body bodyPart = wordDocument.MainDocumentPart.Document.Body;

                if (bodyPart == null)
                    return null;
                else
                    return bodyPart;
            }
        }
        #endregion

        #region Initials
        internal static void initialSettings()
        {
            if (!Globals.ThisAddIn.accessedInOpen && !Globals.ThisAddIn.accessedInStartup)
            {
                Globals.ThisAddIn.setInitialSettings = true;
                try
                {
                    //Globals.ThisAddIn.Application.ActiveDocument.ReadOnlyRecommended = false;

                    Globals.ThisAddIn.previousMeasurementUnits = Globals.ThisAddIn.Application.Options.MeasurementUnit;
                    Globals.ThisAddIn.previousAraSpeller = Globals.ThisAddIn.Application.Options.ArabicMode;
                    Globals.ThisAddIn.previousArabicNumeral = Globals.ThisAddIn.Application.Options.ArabicNumeral;
                    Globals.ThisAddIn.previousDocumentViewDirection = Globals.ThisAddIn.Application.Options.DocumentViewDirection;
                    //Globals.ThisAddIn.previousBibliographyStyle = Globals.ThisAddIn.Application.Options.BibliographyStyle;
                    Globals.ThisAddIn.previousDisplayAlerts = Globals.ThisAddIn.Application.DisplayAlerts;

                    Registry.SetValue(@"HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Common\Toolbars", "ScreenTipScheme", 0, RegistryValueKind.DWord);
                    Globals.ThisAddIn.Application.DisplayScreenTips = true;
                    Globals.ThisAddIn.Application.DisplayAlerts = WdAlertLevel.wdAlertsAll;//TODO:disable for release 
                    Globals.ThisAddIn.Application.Options.MeasurementUnit = WdMeasurementUnits.wdMillimeters;
                    Globals.ThisAddIn.Application.Options.ArabicMode = WdAraSpeller.wdBoth;
                    Globals.ThisAddIn.Application.Options.ArabicNumeral = WdArabicNumeral.wdNumeralContext;
                    Globals.ThisAddIn.Application.Options.DocumentViewDirection = WdDocumentViewDirection.wdDocumentViewRtl;
                    //Application.Options.OMathAutoBuildUp = true;//TODO:omath Options
                    //Globals.ThisAddIn.Application.Options.BibliographyStyle = DedicatedFunctions.getStaticVariableValue(doc,VariableOtherNames._variable_other_BibliographyStyle.ToString());//TODO:check and resolve Bibliography
                    //Application.Options.AutoKeyboardSwitching = true;
                    //Application.Options.DefaultFilePath[WdDefaultFilePath.wdAutoRecoverPath]
                }
                catch (Exception e)
                {
                    DedicatedFunctions.ShowErrorMessage("خطا در تنظیم کردن تنظیمات اولیه" + "\n" + e.Message);
                }
            }
        }
        internal static void restoreInitialSettings()
        {
            if (Globals.ThisAddIn.setInitialSettings)
            {
                Globals.ThisAddIn.Application.DisplayAlerts = Globals.ThisAddIn.previousDisplayAlerts;//TODO:disable for release 

                Globals.ThisAddIn.Application.Options.MeasurementUnit = Globals.ThisAddIn.previousMeasurementUnits;
                Globals.ThisAddIn.Application.Options.ArabicMode = Globals.ThisAddIn.previousAraSpeller;
                Globals.ThisAddIn.Application.Options.ArabicNumeral = Globals.ThisAddIn.previousArabicNumeral;
                Globals.ThisAddIn.Application.Options.DocumentViewDirection = Globals.ThisAddIn.previousDocumentViewDirection;
                //Globals.ThisAddIn.Application.Options.BibliographyStyle = Globals.ThisAddIn.previousBibliographyStyle;
            }
        }
        #endregion

        #region Forms And Dialogs
        /// <summary>
        /// closing dialog with buttons
        /// </summary>
        /// <param name="captionDialog">caption of the window(be exactly the same)</param>
        /// <param name="actionButtonCaption">caption of the button(be exactly the same)</param>
        /// <param name="ifTextExistDoAction">caption of the button(It is not case sensitive)</param>
        /// <param name="ifSecondaryTextExistDoAction">caption of the button(It is not case sensitive and space and tab and break line)</param>
        /// <param name="repeatTrying">time trying to find dialog</param>
        /// <param name="repeatTryingChild">time trying to find childs of dialog</param>
        internal static void closeDialog(string captionDialog, string actionButtonCaption, string ifTextExistDoAction, int repeatTrying = 100, int repeatTryingChild = 10)
        {
            IntPtr Ret, ChildRet;
            int tryWindow = 0;

            while (FindWindow(null, captionDialog).ToInt64() == 0)
            {
                if (tryWindow >= repeatTrying)
                    break;
                Thread.Sleep(1);
                tryWindow++;
            }
            Ret = FindWindow(null, captionDialog);

            if (Ret != IntPtr.Zero)
            {
                tryWindow = 0;
                ChildRet = IntPtr.Zero;
                do
                {
                    while (FindWindowEx(Ret, ChildRet, "Static", null).ToInt64() == 0)
                    {
                        if (tryWindow >= repeatTryingChild)
                            break;
                        Thread.Sleep(1);
                        tryWindow++;
                    }
                    ChildRet = FindWindowEx(Ret, ChildRet, "Static", null);

                    if (ChildRet != IntPtr.Zero)
                    {
                        int length = SendMessage(ChildRet, WM_GETTEXTLENGTH, 0, 0);
                        StringBuilder messageText = new StringBuilder(length);
                        int hr = SendMessage(ChildRet, WM_GETTEXT, length, messageText);

                        if (messageText.ToString()
                            .ToLower()
                            .Replace(" ", "")
                            .Replace("\n", "")
                            .Replace("\r", "")
                            .Replace("\t", "")
                            .Contains(ifTextExistDoAction.ToLower()
                                .Replace(" ", "")
                                .Replace("\n", "")
                                .Replace("\r", "")
                                .Replace("\t", "")))
                        {
                            SetWindowPos(Ret, IntPtr.Zero, 0, 0, 0, 0, SWP_NOZORDER | SWP_NOCOPYBITS | SWP_NOACTIVATE | SWP_HIDEWINDOW | SWP_DEFERERASE);
                            tryWindow = 0;
                            while (FindWindowEx(Ret, IntPtr.Zero, "Button", actionButtonCaption).ToInt64() == 0)
                            {
                                if (tryWindow >= repeatTryingChild)
                                    break;
                                Thread.Sleep(1);
                                tryWindow++;
                            }
                            IntPtr ChildRet2 = FindWindowEx(Ret, IntPtr.Zero, "Button", actionButtonCaption);
                            if (ChildRet2 != IntPtr.Zero)
                            {
                                SendMessage(ChildRet2, BM_CLICK, 0, null);
                                SendMessage(ChildRet2, BM_CLICK, 0, null);
                            }
                            break;
                        }
                    }

                } while (ChildRet != IntPtr.Zero);
            }
        }
        internal static void displayCrossReference(Document doc, Selection selection, string caption)
        {
            int oldFiledCount = doc.Fields.Count;

            dynamic dlg = Globals.ThisAddIn.Application.Dialogs[WdWordDialog.wdDialogInsertCrossReference];

            dlg.ReferenceType = caption;
            SendKeys.Send(caption[0] + "{TAB}{DOWN 2}{ENTER}{TAB}{TAB}");// for dlg.ReferenceKind = WdReferenceKind.wdOnlyLabelAndNumber
            dlg.InsertAsHyperlink = true;
            dlg.Show();
            if (caption == Constants.CaptionLabels.captionFormula && doc.Fields.Count != oldFiledCount)
            {
                DedicatedFunctions.insertText(doc, selection, ")");
            }
        }

        internal static void showSplashScreen(int durationTime = 2500)
        {
            Forms.SplashScreen splash = new Forms.SplashScreen(durationTime);
            splash.ShowDialog();
        }

        #endregion

        #region Check Font and Persian Keyboard
        internal static bool checkPersianKeyboardLayout()
        {
            bool keyboardPersianExist = false;
            foreach (InputLanguage lang in InputLanguage.InstalledInputLanguages)
            {
                if (lang.Culture.EnglishName.ToLower().Contains(KeyboardLanguage.Persian.ToString().ToLower()))
                {
                    keyboardPersianExist = true;
                }
            }
            if (!keyboardPersianExist)
            {
                DedicatedFunctions.ShowMessage("زبان صفحه کلید فارسی موجود نیست!\nصفحه ای که باز خواهد شد برای تنظیم زبان صفحه کلید است.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                try
                {
                    string registry_key = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion";
                    Microsoft.Win32.RegistryKey key = Registry.LocalMachine.OpenSubKey(registry_key);

                    string keyContent = key.GetValue("ProductName").ToString().ToLower().Replace(" ", "");
                    if (keyContent.Contains("windows11"))
                    {
                        Process.Start("ms-settings:regionlanguage");
                    }
                    else if (keyContent.Contains("windows10"))
                    {
                        Process.Start("ms-settings:regionlanguage-languageoptions");
                    }
                    else if (keyContent.Contains("windows7") || keyContent.Contains("windows8")) // Windows 7 and 8 and 8.1
                    {
                        var cplPath = System.IO.Path.Combine(Environment.SystemDirectory, "control.exe");
                        Process.Start(cplPath, "/name Microsoft.RegionalAndLanguageOptions");
                    }
                }
                catch (Exception)
                {
                }
                return false;
            }
            else
                return true;
        }
        internal static bool checkingFonts()
        {
            List<string> fontNotInstalled = checkFonts();
            if (fontNotInstalled.Count != 0)
            {
                DialogResult dialog = DedicatedFunctions.ShowMessage("برخی از فونت ها روی سیستم نصب نیست ، مایل به نصب فونت ها هستید؟ ", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

                if (dialog == DialogResult.Yes)
                {
                    installFonts(fontNotInstalled);
                    DedicatedFunctions.ShowMessage("برای اعمال فونت ها، بستن کامل نرم افزار word لازم است", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return false;
            }
            return true;
        }
        internal static List<string> checkFonts()
        {
            bool isBNazaninExist = false;
            bool isBZarExist = false;
            bool isBLotusExist = false;
            bool isBYagutExist = false;
            bool isBBadrExist = false;
            bool isBesmellahExist = false;
            bool isIranNastaliq = false;
            bool isTimesNewRomanExist = false;
            bool isVazirExist = false;
            List<string> fontPaths = new List<string>();

            InstalledFontCollection fontsCollection = new InstalledFontCollection();
            foreach (var fontFamily in fontsCollection.Families)
            {
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBNazanin.Trim().ToLower())
                    isBNazaninExist = true;
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBLotus.Trim().ToLower())
                    isBLotusExist = true;
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBYagut.Trim().ToLower())
                    isBYagutExist = true;
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBZar.Trim().ToLower())
                    isBZarExist = true;
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBBadr.Trim().ToLower())
                    isBBadrExist = true;
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontIranNastaliq.Trim().ToLower())
                    isIranNastaliq = true;
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBesmellah1.Trim().ToLower() ||
                    fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBesmellah2.Trim().ToLower() ||
                    fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBesmellah3.Trim().ToLower() ||
                    fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontBesmellah4.Trim().ToLower())
                    isBesmellahExist = true;
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontTimesNewRoman.Trim().ToLower())
                    isTimesNewRomanExist = true;
                if (fontFamily.Name.Trim().ToLower() == Constants.FontNames.fontVazirmatn.Trim().ToLower())
                    isVazirExist = true;
            }

            if (!isBNazaninExist)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Nazanin), nameof(EmbeddedResourceNames.B_Nazanin), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Nazanin_Bold), nameof(EmbeddedResourceNames.B_Nazanin_Bold), ".ttf"));
            }
            if (!isBZarExist)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Zar), nameof(EmbeddedResourceNames.B_Zar), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Zar_Bold), nameof(EmbeddedResourceNames.B_Zar_Bold), ".ttf"));
            }
            if (!isBLotusExist)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Lotus), nameof(EmbeddedResourceNames.B_Lotus), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Lotus_Bold), nameof(EmbeddedResourceNames.B_Lotus_Bold), ".ttf"));
            }
            if (!isBYagutExist)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Yagut), nameof(EmbeddedResourceNames.B_Yagut), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Yagut_Bold), nameof(EmbeddedResourceNames.B_Yagut_Bold), ".ttf"));
            }
            if (!isBBadrExist)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Badr), nameof(EmbeddedResourceNames.B_Badr), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.B_Badr_Bold), nameof(EmbeddedResourceNames.B_Badr_Bold), ".ttf"));
            }
            if (!isIranNastaliq)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.IranNastaliq), nameof(EmbeddedResourceNames.IranNastaliq), ".ttf"));
            }
            if (!isBesmellahExist)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.Besmellah_1), nameof(EmbeddedResourceNames.Besmellah_1), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.Besmellah_2), nameof(EmbeddedResourceNames.Besmellah_2), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.Besmellah_3), nameof(EmbeddedResourceNames.Besmellah_3), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.Besmellah_4), nameof(EmbeddedResourceNames.Besmellah_4), ".ttf"));
            }
            if (!isTimesNewRomanExist)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.times), nameof(EmbeddedResourceNames.times), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.times_BD), nameof(EmbeddedResourceNames.times_BD), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.times_BI), nameof(EmbeddedResourceNames.times_BI), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.times_I), nameof(EmbeddedResourceNames.times_I), ".ttf"));
            }
            if (!isVazirExist)
            {
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.Vazirmatn_Regular), nameof(EmbeddedResourceNames.Vazirmatn_Regular), ".ttf"));
                fontPaths.Add(DedicatedFunctions.copyFileToTempFolder(getStream(EmbeddedResourceNames.Vazirmatn_Bold), nameof(EmbeddedResourceNames.Vazirmatn_Bold), ".ttf"));
            }
            return fontPaths;
        }
        internal static void installFonts(List<string> fonts)
        {
            foreach (string font in fonts)
            {
                try
                {
                    Process process = Process.Start(font);
                    process.WaitForExit();
                }
                catch (Exception)
                {
                }

                DedicatedFunctions.removeFileFromSystem(font);
            }
        }
        #endregion

        #region Scrolling
        internal static void scrollToPage(Microsoft.Office.Interop.Word.Window window, Selection selection, int pageNumber)
        {
            try
            {
                selection.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToFirst, pageNumber);
                window.ScrollIntoView(selection.Range);
                //window.PageScroll(pageNumber - 1);
            }
            catch (Exception)
            {
            }
        }

        #endregion

        #region detect Persian or English chars
        static bool IsPersianChar(char c)
        {
            return (c >= '\u0600' && c <= '\u06FF') || // Arabic (Including Persian letters)
                   (c >= '\u0750' && c <= '\u077F') || // Arabic Supplement
                   (c >= '\u08A0' && c <= '\u08FF') || // Arabic Extended-A
                   (c >= '\uFB50' && c <= '\uFDFF') || // Arabic Presentation Forms-A
                   (c >= '\uFE70' && c <= '\uFEFF');   // Arabic Presentation Forms-B
        }
        static bool IsEnglishChar(char c)
        {
            return (c >= 'A' && c <= 'Z') || // حروف بزرگ
                   (c >= 'a' && c <= 'z');   // حروف کوچک
        }
        #endregion

    }


    #region new func
    public static class CsvDownloader
    {
        public static List<string> GetCsvFilesByCode(string code)
        {
            string virastarFolder =
                Properties.Settings.Default.WorkSpaceDirectory +
                StringConstant.VirastarFolder;


            if (!Directory.Exists(virastarFolder))
                return new List<string>();


            return Directory.GetFiles(virastarFolder, "*.csv")
                .Where(file =>
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);

                    return fileName.EndsWith(code);
                })
                .ToList();
        }


        private static readonly HttpClient HttpClient = new HttpClient();

        private const string CsvListUrl = "https://shivanegar.ir/api/getcsv";


        public static async Task DownloadAllCsvFilesAsync()
        {
            try
            {
                // مسیر پوشه‌ای که کدهای فعلی از آن می‌خوانند
                string virastarFolder =
                    Properties.Settings.Default.WorkSpaceDirectory +
                    StringConstant.VirastarFolder;


                // اگر پوشه وجود نداشت بساز
                if (!Directory.Exists(virastarFolder))
                {
                    Directory.CreateDirectory(virastarFolder);
                }


                // گرفتن لیست فایل‌ها از API
                List<string> csvUrls = await GetCsvUrlsAsync();


                foreach (string csvUrl in csvUrls)
                {
                    try
                    {
                        await DownloadFileAsync(csvUrl, virastarFolder);
                    }
                    catch (Exception ex)
                    {
                        // اگر یک فایل مشکل داشت، بقیه ادامه پیدا کنند
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("خطا در دانلود CSV ها : " + ex.Message);
            }
        }



        private static async Task<List<string>> GetCsvUrlsAsync()
        {
            string json = await HttpClient.GetStringAsync(CsvListUrl);
           

            var response =
                System.Text.Json.JsonSerializer.Deserialize<CsvListResponse>(
                json,
                new System.Text.Json.JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });


            if (response == null ||
                response.Status != "1" ||
                response.Csv == null)
            {
                throw new Exception("دریافت لیست CSV ها ناموفق بود.");
            }


            return response.Csv;
        }


        private static async Task DownloadFileAsync(
        string url,
        string folder)
        {
            byte[] data =
                await HttpClient.GetByteArrayAsync(url);


            string fileName =
                Path.GetFileName(new Uri(url).LocalPath);


            string filePath =
                Path.Combine(folder, fileName);

            File.WriteAllBytes(filePath, data);


            Console.WriteLine("Downloaded : " + fileName);
        }



        private class CsvListResponse
        {
            public string Status { get; set; }

            public List<string> Csv { get; set; }
        }
    }
    #endregion


    #region CSVCode

    public static class VirastarCodes
    {
        public const string HalfSpace = "1002";
        public const string Signs = "1003";
        public const string SpellingCorrection = "1001";
        public const string Standard = "1004";
        public const string Tanvin = "1000";
    }

    #endregion


}