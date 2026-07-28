using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using CsvHelper;
using DocumentFormat.OpenXml.ExtendedProperties;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using ShivaNegar.Constants;
using ShivaNegar.Forms;
using ShivaNegar.Forms.AddBibliography;
using ShivaNegar.Forms.AddRemovePages;
using ShivaNegar.Forms.BesmellahPage;
using ShivaNegar.Forms.CaptionSettings;
using ShivaNegar.Forms.ChangeContents;
using ShivaNegar.Forms.CitationSettings;
using ShivaNegar.Forms.DocumentSettings;
using ShivaNegar.Forms.FootnoteSettings;
using ShivaNegar.Forms.FormatSettings;
using ShivaNegar.Forms.ShivaNegarManager;
using ShivaNegar.Forms.VirastarSettings;
using ShivaNegar.Models;
using ShivaNegar.TaskPanes.ChatBoxNetworking;
using ShivaNegar.TaskPanes.CrossReference;
using ShivaNegar.TaskPanes.DefenseAnnouncements;
using ShivaNegar.TaskPanes.DesigningCD;
using ShivaNegar.TaskPanes.InsertCitation;
using ShivaNegar.TaskPanes.InsertDedicate;
using ShivaNegar.TaskPanes.InsertNahjBalaghe;
using ShivaNegar.TaskPanes.InsertQuran;
using ShivaNegar.Templates;
using static ShivaNegar.DedicatedFunctions;

namespace ShivaNegar
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public partial class ThisAddIn : Interfaces.IKeyboardShortcuts
    {
        #region Not Tested
        //rng.PageSetup;
        //rng.ParagraphFormat;
        //sourceDoc.Merge
        //sourceDoc.RunAutoMacro
        //sourceDoc.RunLetterWizard
        //sourceDoc.SetLetterContent

        //rng.FormattedText;
        //rng.StoryType;

        //rng.TextRetrievalMode.IncludeHiddenText = true;

        //rng.Subdocuments.Select();

        //rng.ParentContentControl.AllowInsertDeleteSection = true;

        //Globals.ThisAddIn.Application.Selection.GoToEditableRange
        //Globals.ThisAddIn.Application.Selection.PreviousSubdocument
        //Globals.ThisAddIn.Application.Selection.WholeStory
        //Globals.ThisAddIn.Application.Selection.Words.First.Select();
        //Globals.ThisAddIn.Application.Selection.Sections
        //Globals.ThisAddIn.Application.Selection.Sentences;
        //Globals.ThisAddIn.Application.Selection.Paragraphs;
        //Globals.ThisAddIn.Application.Selection.Characters;
        //Globals.ThisAddIn.Application.Selection.InsertCaption
        //Globals.ThisAddIn.Application.Selection.InsertBreak
        //Globals.ThisAddIn.Application.Selection.Tables;
        //Globals.ThisAddIn.Application.Selection.Document;
        //Globals.ThisAddIn.Application.Selection.Frames;
        //Globals.ThisAddIn.Application.Selection.Fields;


        //Globals.ThisAddIn.Application.Selection.PageSetup;
        //Globals.ThisAddIn.Application.Selection.ParagraphFormat;
        //Globals.ThisAddIn.Application.Selection.FormattedText;
        //Globals.ThisAddIn.Application.Selection.StoryType;
        //Globals.ThisAddIn.Application.Selection.TextRetrievalMode.IncludeHiddenText = true;
        //Globals.ThisAddIn.Application.Selection.Subdocuments.Select();
        //Globals.ThisAddIn.Application.Selection.ParentContentControl.AllowInsertDeleteSection = true;
        //Globals.ThisAddIn.Application.ActiveWindow.View.
        //Globals.ThisAddIn.Application.ActiveWindow.ActivePane
        //Globals.ThisAddIn.Application.ActiveWindow.Panes
        //Globals.ThisAddIn.Application.OnTime();
        //Globals.ThisAddIn.Application.Caption = "tested";//for Table
        //Globals.ThisAddIn.Application.KeysBoundTo.
        //Globals.ThisAddIn.Application.EnableCancelKey
        //Globals.ThisAddIn.Application.ActiveWindow.ActivePane.View.SeekView
        //Globals.ThisAddIn.Application.Selection.Font.Hidden
        //Globals.ThisAddIn.Application.Selection.ClearParagraphAllFormatting();
        //Globals.ThisAddIn.Application.Selection.Bookmarks["\\Page"].Range.Editors.Add(WdEditorType.wdEditorEveryone);
        //Microsoft.Office.Tools.Word.Controls.Button button = this.Controls.AddButton(Application.Selection.Range, Application.MillimetersToPoints(50), Application.MillimetersToPoints(50), "salesButton");


        //DedicatedFunctions.ShowMessage(Globals.ThisAddIn.Application.ActiveDocument.BuiltInDocumentProperties);
        //DedicatedFunctions.ShowMessage(Globals.ThisAddIn.Application.ActiveDocument.Creator.ToString());
        //DedicatedFunctions.ShowMessage(Globals.ThisAddIn.Application.ActiveDocument.CurrentRsid.ToString());
        //DedicatedFunctions.ShowMessage(Globals.ThisAddIn.Application.ActiveDocument.CustomDocumentProperties);
        //DedicatedFunctions.ShowMessage(Globals.ThisAddIn.Application.ActiveDocument.CustomXMLParts.Count.ToString());


        //foreach (Microsoft.Office.Core.WorkflowTask item in Globals.ThisAddIn.Application.ActiveDocument.GetWorkflowTasks())
        //{
        //    DedicatedFunctions.ShowMessage(item.Description);
        //}


        //Importants
        //Globals.ThisAddIn.Application.ActiveDocument.PresentIt();

        //DedicatedFunctions.ShowMessage(document.Reload);

        //Document document = new Document();
        //document.Broadcast
        //document.CodeName
        //document.ContentTypeProperties
        //document.DataForm
        //document.DisableFeatures
        //document.DisableFeaturesIntroducedAfter
        //document.DocumentLibraryVersions
        //document.GetType().
        //DedicatedFunctions.ShowMessage(document.IsInAutosave.ToString());
        //DedicatedFunctions.ShowMessage(document.KerningByAlgorithm.ToString());
        //DedicatedFunctions.ShowMessage(document.LanguageDetected.ToString());
        //DedicatedFunctions.ShowMessage(document.ReadabilityStatistics[]);


        //document.RemoveDateAndTime
        //document.RemoveDocumentInformation
        //document.RemoveLockedStyles
        //document.RemovePersonalInformation

        //document.get_AttachedTemplate
        //document.set_AttachedTemplate
        //document.UpdateStyles
        //document.UpdateStylesOnOpen


        //document.SaveSubsetFonts

        //document.UserControl
        //document.Versions
        //DedicatedFunctions.ShowMessage(document._CodeName);

        //document.EncryptionProvider
        //DedicatedFunctions.ShowMessage(document.PasswordEncryptionFileProperties.ToString());
        //document.Unprotect
        //document.WritePassword
        //DedicatedFunctions.ShowMessage(document.PasswordEncryptionAlgorithm);
        //document.TextEncoding
        //document.SetPasswordEncryptionOptions
        //DedicatedFunctions.ShowMessage(document.ProtectionType);
        //DedicatedFunctions.ShowMessage(document.Permission[]);
        //DedicatedFunctions.ShowMessage(document.Protect);


        //DedicatedFunctions.ShowMessage(Globals.ThisAddIn.Application.ActiveDocument.DocID.ToString());
        //document.Final

        #endregion
        //Microsoft.Office.Tools.Word.Document extendedDocument = Globals.Factory.GetVstoObject(doc);
        //var firstParagraph = extendedDocument.Controls.AddPlainTextContentControl(extendedDocument.Paragraphs[1].Range, "FirstParagraph");
        //var firstParagraph2 = extendedDocument.Controls.AddRichTextContentControl(extendedDocument.Paragraphs[2].Range, "FirstParagraph2");

        #region Variable and Properties

        //Form
        private ShivaNegarForm documentManagerForm;
        private ChangeContentsForm changeContentsForm;

        public bool DocumentManagerFormVisible { get; set; } = false;
        public bool ChangeContentFormVisible { get; set; } = false;


        //Server
        public bool ManualyDisableServer { get; set; } = false;

        //Restore Initial Settings
        public WdMeasurementUnits previousMeasurementUnits;
        public WdAraSpeller previousAraSpeller;
        public WdArabicNumeral previousArabicNumeral;
        public WdDocumentViewDirection previousDocumentViewDirection;
        public WdAlertLevel previousDisplayAlerts;
        public string previousBibliographyStyle;
        public bool accessedInStartup = false;
        public bool accessedInOpen = false;


        //Disable Events
        public bool DisableEvents { get; set; } = false;
        public bool DisableSelectionChangedEvent { get; set; } = false;
        public bool DisableSaveEvent { get; set; } = false;
        public bool DisableBeforeCloseEvent { get; set; } = false;

        public bool setInitialSettings = false;
        #endregion


        internal bool SetKeyBindingStatus { get; set; } = false;

        #region VSTO Designer generated code
        private void InternalStartup()
        {
            try
            {
                this.Startup += ThisAddIn_Startup;
                this.Application.DocumentOpen += Application_DocumentOpen;
                this.Application.WindowActivate += Application_WindowActivate;
                this.Application.WindowSelectionChange += Application_WindowSelectionChange;
                ((ApplicationEvents4_Event)this.Application).NewDocument += Application_DocumentOpen;
                this.Application.DocumentBeforeSave += Application_DocumentBeforeSave;
                this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;
                this.Shutdown += ThisAddIn_Shutdown;
                //((ApplicationEvents4_Event)this.Application).DocumentChange += ThisAddIn_DocumentChange;
                //((ApplicationEvents4_Event)this.Application).Quit += ThisAddIn_Quit;
                //this.Application.ActiveDocument.New

                //Correct WorkSpaceDirectoryPath
                System.Threading.Tasks.Task.Run(() =>
                {
                    if (string.IsNullOrEmpty(Properties.Settings.Default.WorkSpaceDirectory))
                    {
                        Properties.Settings.Default.WorkSpaceDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), StringConstant.DefaultWorkspaceName) + "\\";
                        Properties.Settings.Default.Save();

                    }
                    if (!Properties.Settings.Default.WorkSpaceDirectory.EndsWith("\\"))
                    {
                        Properties.Settings.Default.WorkSpaceDirectory += "\\";
                        Properties.Settings.Default.Save();
                    }
                    Directory.CreateDirectory(Properties.Settings.Default.WorkSpaceDirectory);
                });

                //Stopwatch sw = Stopwatch.StartNew();

                //var t = new System.Threading.Thread(o =>
                //{
                //ShivaNegarForm frm = new ShivaNegarForm(false);
                //frm.Close();

                //Initialize Forms And OpenXML
                //System.Windows.Forms.Integration.ElementHost elementHost1 = new System.Windows.Forms.Integration.ElementHost();
                //elementHost1.Dispose();
                ShivaNegarControl shivaNegarControl = new ShivaNegarControl(false);
                //DocumentManagerControl documentManagerControl = new DocumentManagerControl();
                //CreateDocumentControl createDocumentControl = new CreateDocumentControl();

                //MessageBox.Show(sw.ElapsedMilliseconds.ToString());
                //sw.Stop();

                //ShivaNegarForm acs = new ShivaNegarForm(false);

                //System.Windows.Threading.Dispatcher.Run();

                //});
                //t.SetApartmentState(System.Threading.ApartmentState.STA);
                //t.Start();

                System.Threading.Tasks.Task.Run(() =>
                {
                    DedicatedFunctions.InitializeOpenXml();
                });

                //initialize Tsl support in Http Client
                ServicePointManager.Expect100Continue = true;
                System.Net.ServicePointManager.SecurityProtocol |=
                    SecurityProtocolType.Tls12 |
                    SecurityProtocolType.Tls11 |
                    SecurityProtocolType.Tls; // comparable to modern browsers
            }
            catch (Exception e)
            {
                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای در مقدار دهی اولیه به وجود آمده" + "\nپیغام خطا:\n" + e.Message, (int)ErrorCodes.InternalStartup, StringConstant.SupportEmail);
            }

            //What?
            //foreach(Document doc in Globals.ThisAddIn.Application.Documents)
            //{
            //	if(DedicatedFunctions.hasAccess(doc) == DedicatedFunctions.AccessType.AccessGranted)
            //		doc.ActiveWindow.Visible = true;
            //}
        }

        // for expose to VBA (Shortcut)
        protected override object RequestComAddInAutomationService()
        {
            return this;
        }

        #endregion

        #region Events

        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                await CsvDownloader.DownloadAllCsvFilesAsync();
                //unload already exists dedicated keyboard shortcuts

                if (!Globals.ThisAddIn.SetKeyBindingStatus)
                    Ribbon.unsetKeyBinding(Ribbon.getKeyboardRelations());

                ////load keyboard hooking
                //KeyboardHooking.SetHook();

                if (Globals.ThisAddIn.Application.Documents.Count != 0)//if start Page is not Show
                {
                    Document doc;
                    try
                    {
                        doc = Globals.ThisAddIn.Application.ActiveDocument;
                    }
                    catch (Exception)
                    {
                        return;
                    }
                    DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

                    doc.ContentControlOnExit += Doc_ContentControlOnExit;

                    if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
                    {
                        DisableEvents = true;
                        DedicatedFunctions.checkingFonts();
                        DedicatedFunctions.checkPersianKeyboardLayout();

                        Ribbon.InitializeRibbon(StringConstant.NameOfProject);
                        Ribbon.loadKeyboardShortcut();
                        DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);
                        DisableEvents = false;

                        accessedInStartup = true;
                    }
                    else if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                    {
                        Globals.ThisAddIn.Application.ScreenUpdating = false;

                        DedicatedFunctions.showSplashScreen();

                        bool checkResult = await DedicatedFunctions.checkingUpdate();

                        if (checkResult)
                        {
                            DisableEvents = true;

                            DedicatedFunctions.checkingFonts();
                            DedicatedFunctions.checkPersianKeyboardLayout();

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

                            #region initial Settings
                            DedicatedFunctions.initialSettings();
                            try
                            {
                                Ribbon.InitializeRibbon(StringConstant.NameOfProject + "(" + DedicatedFunctions.getDocumentTypePersianName(doc) + ")");
                            }
                            catch (Exception)
                            {
                                Ribbon.InitializeRibbon(StringConstant.NameOfProject);
                            }
                            Ribbon.loadKeyboardShortcut();
                            #endregion

                            Globals.ThisAddIn.Application.ScreenUpdating = true;
                            DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);
                            DedicatedFunctions.saveDocument(doc);
                            DisableEvents = false;

                            accessedInStartup = true;
                        }
                        else
                        {
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
                else
                    Ribbon.RibbonControlsVisibility(false);
            }
            catch (Exception ex)
            {
                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای در بارگیری افزونه رخ داده است\n افزونه غیرفعال میشود\n پیغام خطا:\n" + ex.Message,
                    (int)ErrorCodes.StartupProblem, StringConstant.SupportEmail);
            }
        }


        private void Application_DocumentOpen(Document doc)
        {
            if (Globals.ThisAddIn.Application.Caption.EndsWith(" (Exported)"))
            {
                Globals.ThisAddIn.Application.Caption = Globals.ThisAddIn.Application.Caption.Substring(0,
                    Globals.ThisAddIn.Application.Caption.Length - " (Exported)".Length);
            }
            if (!DisableEvents)
            {
                checkAndInitializeOpenedDocument(doc);
            }
        }
        private void Application_WindowActivate(Document doc, Microsoft.Office.Interop.Word.Window Wn)
        {
            if (!DisableEvents)
            {
                DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
                if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                {
                    string ribbonTitle = StringConstant.NameOfProject + "(" + DedicatedFunctions.getDocumentTypePersianName(doc) + ")";

                    Ribbon.setTabProperties(ribbonTitle, true);
                    Ribbon.RibbonControlsVisibility(true);
                }
                else if (Globals.ThisAddIn.Application.Documents.Count != 0 && accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
                {
                    Ribbon.setTabProperties(StringConstant.NameOfProject, true);
                    Ribbon.RibbonControlsVisibility(true);
                }
                else
                {
                    Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                    Ribbon.RibbonControlsVisibility(false);
                }
            }
        }
        private void Application_DocumentBeforeSave(Microsoft.Office.Interop.Word.Document doc, ref bool SaveAsUI, ref bool Cancel)
        {
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                updateChapters(doc);
            }

            if (!DisableEvents && !DisableSaveEvent)
            {
                if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                {
                    if (SaveAsUI)
                    {
                        Cancel = true;
                        DedicatedFunctions.ShowErrorMessage("امکان save as کردن سند وجود ندارد");
                    }
                }
            }
        }
        private void Application_DocumentBeforeClose(Microsoft.Office.Interop.Word.Document doc, ref bool Cancel)
        {
            if (!DisableEvents && !DisableBeforeCloseEvent)
            {
                DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
                if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
                {
                    //check if Create CD Task pane already Open
                    CustomTaskPane customTaskPane = DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD);
                    if (customTaskPane != null && customTaskPane.Visible)
                    {
                        DialogResult dialogTaskPaneCD = DedicatedFunctions.ShowMessage("عملیات ساخت لوح فشرده در حال انجام است، آیا مایل به بستن برنامه هستید؟", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (dialogTaskPaneCD == DialogResult.Yes)
                        {
                            Globals.ThisAddIn.Application.ScreenUpdating = false;

                            DesigningCDTaskPane.DeleteAndDisableActionPane(doc, customTaskPane, null);

                            Globals.ThisAddIn.Application.ScreenUpdating = true;
                        }
                        else
                        {
                            Cancel = true;
                            return;
                        }
                    }

                    if (!doc.Saved)
                    {
                        // نمایش فرم اختصاصی
                        string message = "آیا می خواهید تغییرات خود را در " + "\"" + doc.Name + "\"" + " ذخیره کنید؟";
                        //message += "\n\n" + "اگر روی \"NO\" کلیک کنید، یک نسخه از این فایل به طور موقت در دسترس خواهد بود.";
                        DialogResult closeDialog = DedicatedFunctions.ShowMessage(message, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button3);

                        if (closeDialog == DialogResult.Yes)
                        {
                            //save currentPagePosition
                            int currentPagePosition = (int)this.Application.Selection.Information[WdInformation.wdActiveEndPageNumber];
                            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_PositionCurrentPage.ToString(), currentPagePosition.ToString());

                            //DedicatedFunctions.saveDocument(doc);
                            if (Properties.Settings.Default.AutoSaveOnClose)
                            {
                                DedicatedFunctions.uploadDocument(doc, false);
                            }
                            DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdSaveChanges);
                        }
                        else if (closeDialog == DialogResult.No)
                        {
                            DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges);
                        }
                        else
                        {
                            Cancel = true;
                            return;
                        }
                    }
                    else
                    {
                        if (Globals.ThisAddIn.Application.Documents.Count == 1)
                        {
                            try
                            {
                                Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                                Ribbon.RibbonControlsVisibility(false);
                            }
                            catch (Exception)
                            {
                                //MessageBox.Show("خطا رخ داد در غیرفعال کردن تب" + "\nپیغام خطا:\n" + e.Message);
                            }
                        }

                        if (Properties.Settings.Default.AutoSaveOnClose)
                        {
                            DedicatedFunctions.uploadDocument(doc, false);
                        }
                    }
                }

                if (Globals.ThisAddIn.Application.Documents.Count <= 1)
                {
                    if (accessedInOpen || accessedInStartup)
                    {
                        accessedInOpen = false;
                        accessedInStartup = false;

                        try
                        {
                            if (File.Exists(doc.FullName))
                            {
                                DedicatedFunctions.restoreInitialSettings();
                                DisableBeforeCloseEvent = true;
                                DedicatedFunctions.saveDocument(doc);
                                DisableBeforeCloseEvent = false;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
        }
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.RememberUser == false)
            {
                Properties.Settings.Default.UserToken = "";
                Properties.Settings.Default.Mobile = "";
                Properties.Settings.Default.Save();
            }

            if (!DisableEvents)
            {
                // disabled Addin
                if (Application.Documents.Count > 0)
                {
                    if (accessedInOpen || accessedInStartup)
                    {
                        DedicatedFunctions.restoreInitialSettings();
                        //DedicatedFunctions.unsetKeyBinding(DedicatedFunctions.getKeyboardRelations());
                    }
                }

                // Unload Template Shortcut
                Ribbon.unloadKeyboardShortcut();

                ////unload keyboard hooking
                //KeyboardHooking.ReleaseHook();

                // disabled Addin
                if (Application.Documents.Count > 0)
                {
                    bool hasDedicatedDocuments = false;
                    List<Document> docs = new List<Document>();
                    foreach (Document doc in Application.Documents)
                    {
                        if (DedicatedFunctions.hasAccess(doc) == DedicatedFunctions.AccessType.AccessGranted)
                        {
                            hasDedicatedDocuments = true;
                            docs.Add(doc);
                        }
                    }

                    if (hasDedicatedDocuments)
                    {
                        DisableBeforeCloseEvent = true;
                        DialogResult dr = DedicatedFunctions.ShowMessage("آیا اسناد " + StringConstant.NameOfProject + " ذخیره شود؟\n در صورتی که Yes کلیک شود تغییرات ذخیره شده و در صورتی که No کلیک شود بدون ذخیره تنظیمات سند بسته میشود", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                        if (dr == DialogResult.Yes)
                        {
                            foreach (var item in docs)
                            {
                                DedicatedFunctions.closeDocument(item, WdSaveOptions.wdSaveChanges);
                            }
                        }
                        else
                        {
                            foreach (var item in docs)
                            {
                                DedicatedFunctions.closeDocument(item, WdSaveOptions.wdDoNotSaveChanges);
                            }
                        }
                    }
                }
            }
        }

        //enable or disable ControlProperties
        //if(DedicatedFunctions.hasAccess(doc) == DedicatedFunctions.AccessType.AccessGranted)
        //{
        //	if(contentControl.Tag.Contains(InitialContentControls.initialContentControlNames) || contentControl.Tag.Contains(InitialContentControls.initialPageID))
        //		Globals.Ribbon.ribbonControlledInCommands.Where(p => p.Id == RibbonControlNames.builtInControlProperties).First().Enable = false;
        //	else
        //		Globals.Ribbon.ribbonControlledInCommands.Where(p => p.Id == RibbonControlNames.builtInControlProperties).First().Enable = true;
        //}
        //else
        //{
        //	Globals.Ribbon.ribbonControlledInCommands.Where(p => p.Id == RibbonControlNames.builtInControlProperties).First().Enable = true;
        //}
        //Globals.Ribbon.ribbon?.Invalidate();

        Range previousSelection;
        private void Application_WindowSelectionChange(Selection sel)
        {

            if (previousSelection == null)
                previousSelection = sel.Range;


            if (!DisableEvents && !DisableSelectionChangedEvent)
            {
                Document doc;
                try
                {
                    doc = Globals.ThisAddIn.Application.ActiveDocument;
                }
                catch (Exception)
                {
                    return;
                }
                if (DedicatedFunctions.hasAccess(doc) == DedicatedFunctions.AccessType.AccessGranted)
                {
                    //for protect, to prevent delete SectionBreak(protected by ContentControl)
                    string text = "";
                    try
                    {
                        text = sel.Text;
                    }
                    catch (Exception)
                    {
                        return;
                    }

                    if (text == "\f")//is section break only (or page break)
                    {
                        if (sel.Previous(WdUnits.wdParagraph, 1).ContentControls.Count == 1)
                        {
                            Microsoft.Office.Interop.Word.ContentControl contentControl = sel.Previous(WdUnits.wdParagraph, 1).ContentControls[1];

                            if (contentControl.Tag.Contains(InitialContentControls.initialPageID))
                            {
                                Range rng = contentControl.Range;
                                rng.Collapse(WdCollapseDirection.wdCollapseStart);
                                rng.MoveEnd(WdUnits.wdParagraph, -2);
                                rng.Select();
                                sel.EndKey();

                                //sel.MoveUp(WdUnits.wdParagraph , 2);
                                //rng.MoveStart(WdUnits.wdParagraph , -2);
                                //sel.HomeKey();
                                //sel.MoveUp(WdUnits.wdParagraph , 2);
                                //rng.Collapse(WdCollapseDirection.wdCollapseEnd);
                                //rng.Select();
                            }
                        }
                    }
                    else if (sel.ContentControls.Count > 0)
                    {
                        //filter selection
                        foreach (Microsoft.Office.Interop.Word.ContentControl cc in sel.ContentControls)
                        {
                            if (cc != null && cc.Tag != null)
                            {
                                if (cc.Tag.Contains(InitialContentControls.initialPageID.ToString()))
                                {
                                    try
                                    {
                                        //if(sel.Start < cc.Range.Start - 1)
                                        if (previousSelection.End > cc.Range.End && sel.End > cc.Range.End)
                                        {
                                            this.Application.ActiveDocument.Range(cc.Range.End + 2, sel.End).Select();
                                        }
                                        else if (previousSelection.Start < cc.Range.Start && sel.Start < cc.Range.Start)
                                        {
                                            this.Application.ActiveDocument.Range(sel.Start, cc.Range.Start - 1).Select();
                                        }
                                        else
                                        {
                                            this.Application.ActiveDocument.Range(cc.Range.Start - 2, cc.Range.Start - 2).Select();
                                        }
                                        //break;
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                                else
                                {
                                    if (sel.ContentControls.Count == 1)
                                    {
                                        //string[] values = Enum.GetNames(typeof(ContentControlNames));


                                        //Microsoft.Office.Interop.Word.ContentControl cc2 = sel.ContentControls[1];

                                        //if(values.Contains(cc2.Tag))
                                        //{
                                        //	//Globals.ThisAddIn.Application.Selection.EndKey();
                                        //	//break;
                                        //}
                                        //else
                                        //{

                                        //}
                                    }
                                    else
                                    {

                                    }
                                }
                            }
                        }
                    }

                    //    if (Sel.Text != "" && Sel.Text.Length > 2)//remove uncorrect halfSpace and replace it
                    //    {
                    //        Range allContents = doc.Content;
                    //        allContents.Find.ClearFormatting();
                    //        allContents.Find.Execute(FindText: "^-", ReplaceWith: "\u200c", Replace: WdReplace.wdReplaceAll);
                    //    }

                    if (sel.OMaths.Count == 1)
                    {
                        OMath parent = sel.OMaths[1];
                        //DedicatedFunctions.recursiveOMathFunctions(parent);
                        sel.OMaths[1].ConvertToNormalText();
                    }
                }
            }

            try
            {
                previousSelection = sel.Range;
            }
            catch (Exception)
            {
            }
        }

        internal void Doc_ContentControlOnExit(Microsoft.Office.Interop.Word.ContentControl contentControl, ref bool Cancel)
        {
            if (contentControl.Tag != null && contentControl.Tag.Contains(InitialContentControls.initialChapter))
            {
                Document doc;
                try
                {
                    doc = Globals.ThisAddIn.Application.ActiveDocument;
                }
                catch (Exception)
                {
                    return;
                }
                updateChapters(doc);
            }
        }

        #endregion

        #region Interfaces
        public void documentsManager()
        {
            //Stopwatch sw = Stopwatch.StartNew();
            try
            {
                if (DocumentManagerFormVisible)
                {
                    if (documentManagerForm != null)
                        documentManagerForm.Focus();
                    else
                    {
                        documentManagerForm = new ShivaNegarForm(true);
                        documentManagerForm.Show();
                        DocumentManagerFormVisible = true;
                    }
                }
                else
                {
                    documentManagerForm = new ShivaNegarForm(true);
                    documentManagerForm.Show();
                    DocumentManagerFormVisible = true;
                }
            }
            catch (Exception)
            {
            }
            //MessageBox.Show(sw.ElapsedMilliseconds.ToString());
            //sw.Stop();
        }

        public void documentsManagerArchive()
        {
            try
            {
                if (DocumentManagerFormVisible && documentManagerForm != null)
                {
                    documentManagerForm.ShowArchiveDocuments();
                    documentManagerForm.Focus();
                }
                else
                {
                    documentManagerForm = new ShivaNegarForm(true);
                    documentManagerForm.Show();
                    DocumentManagerFormVisible = true;
                    documentManagerForm.ShowArchiveDocuments();
                }
            }
            catch (Exception)
            {
            }
        }

        public void documentsManagerCreate()
        {
            try
            {
                if (DocumentManagerFormVisible && documentManagerForm != null)
                {
                    var mainControl = documentManagerForm.GetMainControl();
                    if (mainControl != null)
                    {
                        mainControl.CreateDocumentPage();
                    }
                    documentManagerForm.Focus();
                }
                else
                {
                    documentManagerForm = new ShivaNegarForm(true);
                    documentManagerForm.Show();
                    DocumentManagerFormVisible = true;

                    var mainControl = documentManagerForm.GetMainControl();
                    if (mainControl != null)
                    {
                        mainControl.CreateDocumentPage();
                    }
                }
            }
            catch (Exception)
            {
            }
        }
        public void ShowAbout()
        {
            try
            {
                
                var aboutControl = new ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View.AboutUs(false);

                
                System.Windows.Forms.Integration.ElementHost elementHost =
                    new System.Windows.Forms.Integration.ElementHost();
                elementHost.Child = aboutControl;
                elementHost.Dock = System.Windows.Forms.DockStyle.Fill;

                
                System.Windows.Forms.Form form = new System.Windows.Forms.Form();
                form.Text = "درباره پارسانگار";
                form.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                form.Size = new System.Drawing.Size(550, 450);
                form.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
                form.RightToLeftLayout = true;
                form.Controls.Add(elementHost);

                form.ShowDialog();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "خطا در نمایش درباره ما:\n" + ex.Message,
                    "خطا",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
        public void chatBoxNetworking()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            //string sharedDocumentValue = DedicatedFunctions.getStaticVariableValue(doc, VariableDocumentIDs._variable_document_SharedDocument.ToString());
            //if (sharedDocumentValue == "1")
            //{
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                if (DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD) == null ||
                !DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD).Visible)
                {
                    DedicatedFunctions.RemoveAllTaskPanes(doc);
                    CustomTaskPane customTaskPane = DedicatedFunctions.AddTaskPane(new ChatBoxNetworkingTaskPane(), doc, TaskPaneTitles.ChatBoxNetworking);
                    ((ChatBoxNetworkingTaskPane)customTaskPane.Control).enableActionPane(customTaskPane);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("عملیات ساخت لوح فشرده در حال انجام است، برای ارجاع دادن، عملیات ساخت لوح فشرده را لغو کنید");
                }
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }
            //}
        }
        public void insertQuran()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                if (DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD) == null ||
                !DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD).Visible)
                {
                    DedicatedFunctions.RemoveAllTaskPanes(doc);
                    //DedicatedFunctions.RemoveTaskPane(Globals.ThisAddIn.Application.ActiveDocument, TaskPaneTitle.CrossReference);
                    CustomTaskPane customTaskPane = DedicatedFunctions.AddTaskPane(new InsertQuranTaskPane(), doc, TaskPaneTitles.InsertQuran);
                    ((InsertQuranTaskPane)customTaskPane.Control).enableActionPane(customTaskPane);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("عملیات ساخت لوح فشرده در حال انجام است، برای ارجاع دادن، عملیات ساخت لوح فشرده را لغو کنید");
                }
            }
        }
        public void insertNahjBalaghe()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                if (DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD) == null ||
!DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD).Visible)
                {
                    DedicatedFunctions.RemoveAllTaskPanes(doc);
                    //DedicatedFunctions.RemoveTaskPane(Globals.ThisAddIn.Application.ActiveDocument, TaskPaneTitle.CrossReference);
                    CustomTaskPane customTaskPane = DedicatedFunctions.AddTaskPane(new InsertNahjBalagheTaskPane(), doc, TaskPaneTitles.InsertNahjBalaghe);
                    ((InsertNahjBalagheTaskPane)customTaskPane.Control).enableActionPane(customTaskPane);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("عملیات ساخت لوح فشرده در حال انجام است، برای ارجاع دادن، عملیات ساخت لوح فشرده را لغو کنید");
                }
            }

        }

        public void insertDedicate()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                if (DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD) == null ||
                !DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD).Visible)
                {
                    DedicatedFunctions.RemoveAllTaskPanes(doc);
                    //DedicatedFunctions.RemoveTaskPane(Globals.ThisAddIn.Application.ActiveDocument, TaskPaneTitle.CrossReference);
                    CustomTaskPane customTaskPane = DedicatedFunctions.AddTaskPane(new InsertDedicateTaskPane(), doc, TaskPaneTitles.InsertDedicate);
                    ((InsertDedicateTaskPane)customTaskPane.Control).enableActionPane(customTaskPane);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("عملیات ساخت لوح فشرده در حال انجام است، برای ارجاع دادن، عملیات ساخت لوح فشرده را لغو کنید");
                }
            }
        }

        public void defenseAnnouncements()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                if (DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD) == null ||
                !DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD).Visible)
                {
                    DedicatedFunctions.RemoveAllTaskPanes(doc);
                    CustomTaskPane customTaskPane = DedicatedFunctions.AddTaskPane(new DefenseAnnouncementsTaskPane(), doc, TaskPaneTitles.DefenseAnnouncements);
                    ((DefenseAnnouncementsTaskPane)customTaskPane.Control).enableActionPane(customTaskPane);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("عملیات ساخت لوح فشرده در حال انجام است، برای ارجاع دادن، عملیات ساخت لوح فشرده را لغو کنید");
                }
            }
        }

        public void besmellahPageForm()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                BesmellahPageForm bpf = new BesmellahPageForm();
                bpf.ShowDialog();
            }
        }
        public void createProposal()
        {
            return;

            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                Universities university = DedicatedFunctions.getUniversity(doc);
                TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

                List<ProposalRelationshipModel> proposals = TemplateAccess.getProposalModelList(university, templateType);

                if (proposals == null)
                {
                    DedicatedFunctions.ShowMessage("پروپوزالی برای این دانشگاه اضافه نشده است.");
                }
                else
                {

                    AcademicDegrees academicDegree = DedicatedFunctions.getAcademicDegreeID(doc);
                    bool found = false;
                    foreach (ProposalRelationshipModel proposal in proposals)
                    {
                        if (proposal.AcademicDegreeID == academicDegree)
                        {
                            found = true;

                            System.Windows.Forms.SaveFileDialog fileDialog = new System.Windows.Forms.SaveFileDialog();
                            fileDialog.Title = "مکان و نام خروجی فایل خود را مشخص کنید";
                            fileDialog.AddExtension = false;
                            fileDialog.Filter = "Word Document (*.docx)|*.docx";
                            if (fileDialog.ShowDialog() == DialogResult.OK)
                            {
                                if (fileDialog.FileName != doc.FullName)
                                {
                                    string resourcePath = proposal.ResourcePath + proposal.FileName;

                                    Stream stream = DedicatedFunctions.getStream(resourcePath);
                                    string templatePath = DedicatedFunctions.copyFileToFolder(stream, fileDialog.FileName);

                                    Globals.ThisAddIn.DisableEvents = true;
                                    Document specifiedDocument = Globals.ThisAddIn.Application.Documents.Open(FileName: templatePath, Visible: false, ReadOnly: false, OpenAndRepair: false, NoEncodingDialog: true);
                                    Globals.ThisAddIn.DisableEvents = false;

                                    foreach (Microsoft.Office.Interop.Word.ContentControl ccSpecified in specifiedDocument.ContentControls)
                                    {
                                        foreach (Microsoft.Office.Interop.Word.ContentControl ccCurrent in doc.ContentControls)
                                        {
                                            if (ccCurrent.Tag == ccSpecified.Tag)
                                            {
                                                ccSpecified.Range.Text = ccCurrent.Range.Text;
                                                ccSpecified.Delete(false);
                                                break;
                                            }
                                        }
                                    }
                                    specifiedDocument.ActiveWindow.Visible = true;
                                }
                                else
                                {
                                    DedicatedFunctions.ShowErrorMessage("سند خروجی نمیتواند با سند اختصاصی شما جایگزین شود، لطفا با نامی دیگر سند را ذخیره کنید");
                                }
                                break;
                            }
                        }
                    }

                    if (!found)
                    {
                        DedicatedFunctions.ShowMessage("پروپوزالی برای مقطع شما یافت نشد!");
                    }
                }
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }
        }
        public void changeContent()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                DedicatedFunctions.saveDocument(doc);
                try
                {
                    if (ChangeContentFormVisible)
                    {
                        if (changeContentsForm != null)
                            changeContentsForm.Focus();
                        else
                        {
                            changeContentsForm = new ChangeContentsForm(doc);
                            changeContentsForm.Show();
                            ChangeContentFormVisible = true;
                        }
                    }
                    else
                    {
                        changeContentsForm = new ChangeContentsForm(doc);
                        changeContentsForm.Show();
                        ChangeContentFormVisible = true;
                    }

                }
                catch (Exception e)
                {
                    DedicatedFunctions.ShowErrorMessage(e.Message);
                }
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }
        }
        public void addRemovePages()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                DedicatedFunctions.saveDocument(doc);

                AddRemovePagesForm addRemovePagesForm = new AddRemovePagesForm(doc);
                addRemovePagesForm.ShowDialog();
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }

        }

        #region Set Style
        public void setNormalStyle()
        {
            Selection selection = Globals.ThisAddIn.Application.Selection;

            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("قالب‌بندی متن پایان‌نامه");
            DedicatedFunctions.textFormating(selection, StyleNames.BodyText);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }
        public void setHeadingStyle(int index)
        {
            Selection selection = Globals.ThisAddIn.Application.Selection;

            if (index == 2)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("قالب‌بندی عنوان سطح دو");
                DedicatedFunctions.textFormating(selection, StyleNames.styleHeading2);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
            else if (index == 3)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("قالب‌بندی عنوان سطح سه");
                DedicatedFunctions.textFormating(selection, StyleNames.styleHeading3);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
            else if (index == 4)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("قالب‌بندی عنوان سطح چهار");
                DedicatedFunctions.textFormating(selection, StyleNames.styleHeading4);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
            else if (index == 5)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("قالب‌بندی عنوان سطح پنج");
                DedicatedFunctions.textFormating(selection, StyleNames.styleHeading5);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        #endregion

        #region Caption and CrossReference
        public void InsertCaption(int index) // 0 > Shape , 1 > table , 2 > Formula
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            Selection selection = Globals.ThisAddIn.Application.Selection;
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                if (selection.InlineShapes.Count != 1 && selection.ShapeRange.Count != 1)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("ایجاد خط جدید");

                    if (selection.Characters.Count > 1 || selection.ContentControls.Count > 0)
                    {
                        selection.EndKey();
                        DedicatedFunctions.insertParagraph(selection);
                    }
                    else if (selection.Paragraphs.Count > 2 || selection.Paragraphs[1].Range.Text.Length > 1)
                    {
                        selection.Paragraphs[selection.Paragraphs.Count].Range.Select();
                        selection.EndKey();
                        DedicatedFunctions.insertParagraph(selection);
                    }
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                }

                if (index == 0)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("درج عنوان شکل");

                    if (selection.InlineShapes.Count == 1)
                    {
                        InlineShape shape = selection.InlineShapes[1];
                        shape.Select();
                        if (selection.Paragraphs[1].Range.Text.Length > 2)
                        {
                            selection.EndKey();
                            DedicatedFunctions.insertParagraph(selection);
                            shape.Select();
                        }
                        if (selection.Paragraphs[1].Range.Text.Length > 2)
                        {
                            selection.HomeKey();
                            DedicatedFunctions.insertParagraph(selection);
                            shape.Select();
                        }
                    }
                    else if (selection.ShapeRange.Count == 1)
                    {
                        Shape shape = selection.ShapeRange[1];
                        shape.Select();
                    }
                    else
                    {
                        if (Properties.Settings.Default.CaptionSettings_ShowDialogCaptionFigure)
                        {
                            System.Windows.Forms.OpenFileDialog openFile = new System.Windows.Forms.OpenFileDialog();
                            openFile.Title = "عکس مورد نظر خود را انتخاب کنید";
                            openFile.Multiselect = false;
                            openFile.AddExtension = false;
                            openFile.Filter = "All Pictures (*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.eps;*.tif;*.tiff;*.jfif;*.jepg;*.jpe;*.emf;*.wmf)|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.eps;*.tif;*.tiff;*.jfif;*.jepg;*.jpe;*.emf;*.wmf";

                            if (openFile.ShowDialog() == DialogResult.OK)
                            {
                                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("درج شکل");
                                InlineShape shape = selection.InlineShapes.AddPicture(openFile.FileName);
                                DedicatedFunctions.changeParagraphAlignment(selection, WdParagraphAlignment.wdAlignParagraphCenter);
                                shape.Select();
                                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                            }

                            //doc.ActiveWindow.SetFocus();
                            doc.ActiveWindow.Activate();
                        }
                    }
                    DedicatedFunctions.insertCaption(doc, selection, Constants.CaptionLabels.captionFigure);
                    DedicatedFunctions.changeKeyboardLanguage(selection, KeyboardLanguage.Persian, TextDirection.RTL);
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                }
                else if (index == 1)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("درج عنوان جدول");
                    DedicatedFunctions.insertCaption(doc, selection, Constants.CaptionLabels.captionTable);
                    DedicatedFunctions.changeKeyboardLanguage(selection, KeyboardLanguage.Persian, TextDirection.RTL);
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                }
                else if (index == 2)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("درج عنوان معادله(معادله نویسی)");
                    DedicatedFunctions.insertCaption(doc, selection, Constants.CaptionLabels.captionFormula);
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.English);
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                }
            }
        }
        public void ShowCrossReferenceMenu()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                if (DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD) == null ||
                !DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD).Visible)
                {
                    DedicatedFunctions.RemoveAllTaskPanes(doc);
                    //DedicatedFunctions.RemoveTaskPane(Globals.ThisAddIn.Application.ActiveDocument, TaskPaneTitle.CrossReference);
                    CustomTaskPane customTaskPane = DedicatedFunctions.AddTaskPane(new CrossReferenceTaskPane(), doc, TaskPaneTitles.CrossReference);
                    ((CrossReferenceTaskPane)customTaskPane.Control).enableActionPane(customTaskPane, doc);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("عملیات ساخت لوح فشرده در حال انجام است، برای ارجاع دادن، عملیات ساخت لوح فشرده را لغو کنید");
                }
            }
        }
        #endregion

        public void insertFootnote(bool Persian)
        {

            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            Selection selection = Globals.ThisAddIn.Application.Selection;
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                if (Persian)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("درج پانویس فارسی");

                    if (accessType == AccessType.AccessGranted)
                    {
                        DedicatedFunctions.selectFootnoteSeparator(doc, DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_FootnoteSeparatorType.ToString()));

                        DedicatedFunctions.selectFootnoteStyle(doc, selection,
                            DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_FootnoteStyle.ToString()));
                    }

                    Range rng = DedicatedFunctions.insertFootnote(doc, selection, accessType, KeyboardLanguage.Persian);
                    doc.ActiveWindow.ScrollIntoView(rng);
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                }
                else
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("درج پانویس انگلیسی");

                    if (accessType == AccessType.AccessGranted)
                    {
                        DedicatedFunctions.selectFootnoteSeparator(doc,
                        DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_FootnoteSeparatorType.ToString()));

                        DedicatedFunctions.selectFootnoteStyle(doc, selection,
                            DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_FootnoteStyle.ToString()));
                    }

                    Range rng = DedicatedFunctions.insertFootnote(doc, selection, accessType, KeyboardLanguage.English);
                    doc.ActiveWindow.ScrollIntoView(rng);
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                }
            }
        }

        #region Virastar

        public void insertHalfSpace()
        {
            Selection selection = Globals.ThisAddIn.Application.Selection;
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("درج نیم‌فاصله");
            DedicatedFunctions.insertHalfSpace(selection);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }
        public void halfSpaceCorrection()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

            string virastarFolder = Properties.Settings.Default.WorkSpaceDirectory + StringConstant.VirastarFolder;

            if (accessType == AccessType.AccessGranted || accessType == AccessType.AccessGranted_Administrator)
            {
                List<SearchReplaceModel> standardModels = new List<SearchReplaceModel>();

                if (Directory.Exists(virastarFolder))
                {
                    string[] files = Directory.GetFiles(
                        virastarFolder,
                        $"*{StringConstant.ButtonCodeSigns}.csv");

                    foreach (string file in files)
                    {
                        using (StreamReader reader = new StreamReader(file, Encoding.UTF8))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            standardModels.AddRange(csv.GetRecords<SearchReplaceModel>());
                        }
                    }
                }


                LoadingForm loadingForm = new LoadingForm();
                if (doc.ActiveWindow.Selection.Words.Count >= 2)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تصحیح نیم‌فاصله‌");

                    System.Threading.Tasks.Task.Run(() => DedicatedFunctions.StringCorrection(loadingForm, doc.ActiveWindow.Selection.Range, standardModels));
                }
                else
                {
                    if (DedicatedFunctions.ShowMessage("متنی برای رعایت نیم‌فاصله انتخاب نشده؛ در کل سند اعمال شود؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تصحیح تمامی نیم‌فاصله‌ها");

                        System.Threading.Tasks.Task.Run(() => DedicatedFunctions.StringCorrection(loadingForm, doc, standardModels));
                    }
                    else
                        return;
                }
                loadingForm.ShowDialog();
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        public void spellingCorrection()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == AccessType.AccessGranted || accessType == AccessType.AccessGranted_Administrator)
            {
                string virastarFolder = Properties.Settings.Default.WorkSpaceDirectory + StringConstant.VirastarFolder;

                List<SearchReplaceModel> standardModels = new List<SearchReplaceModel>();
                if (Directory.Exists(virastarFolder))
                {
                    string[] files = Directory.GetFiles(
                        virastarFolder,
                        $"*{StringConstant.ButtonCodeSpelling}.csv");

                    foreach (string file in files)
                    {
                        using (StreamReader reader = new StreamReader(file, Encoding.UTF8))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            standardModels.AddRange(csv.GetRecords<SearchReplaceModel>());
                        }
                    }
                }



                LoadingForm loadingForm = new LoadingForm();
                if (doc.ActiveWindow.Selection.Words.Count >= 2)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تصحیح غلط املایی‌");

                    System.Threading.Tasks.Task.Run(() => DedicatedFunctions.StringCorrection(loadingForm, doc.ActiveWindow.Selection.Range, standardModels));
                }
                else
                {
                    if (DedicatedFunctions.ShowMessage("متنی برای تصحیح غلط املایی انتخاب نشده؛ در کل سند اعمال شود؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تصحیح تمامی غلط املایی ها");

                        System.Threading.Tasks.Task.Run(() => DedicatedFunctions.StringCorrection(loadingForm, doc, standardModels));
                    }
                    else
                        return;
                }
                loadingForm.ShowDialog();
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        public void neshanehGozariCorrection()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == AccessType.AccessGranted || accessType == AccessType.AccessGranted_Administrator)
            {

                string virastarFolder = Properties.Settings.Default.WorkSpaceDirectory + StringConstant.VirastarFolder;

                List<SearchReplaceModel> standardModels = new List<SearchReplaceModel>();

                if (Directory.Exists(virastarFolder))
                {
                    string[] files = Directory.GetFiles(
                        virastarFolder,
                        $"*{StringConstant.ButtonCodeSigns}.csv");

                    foreach (string file in files)
                    {
                        using (StreamReader reader = new StreamReader(file, Encoding.UTF8))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            standardModels.AddRange(csv.GetRecords<SearchReplaceModel>());
                        }
                    }
                }

                // correct Spaces
                SearchReplaceModel whiteSpaceCorrection = new SearchReplaceModel()
                {
                    Search = "^w",
                    Replace = " ",
                    Wildcard = "0",
                    MatchCase = "1",
                    MatchWholeWord = "1",
                    MatchKashida = "0",
                    MatchDiacritics = "0",
                    MatchAlefHamza = "0",
                    CounterUp = "0"
                };
                standardModels.Add(whiteSpaceCorrection);


                LoadingForm loadingForm = new LoadingForm();

                if (doc.ActiveWindow.Selection.Words.Count >= 2)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تصحیح نشانه گذاری (سجاوندی)");

                    System.Threading.Tasks.Task.Run(() => DedicatedFunctions.StringCorrection(loadingForm, doc.ActiveWindow.Selection.Range, standardModels));
                }
                else
                {
                    DialogResult dr = DedicatedFunctions.ShowMessage("متنی برای رعایت نشانه گذاری (سجاوندی) انتخاب نشده؛ در کل سند اعمال شود؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);

                    if (dr == DialogResult.Yes)
                    {
                        Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تصحیح تمامی نشانه گذاری (سجاوندی) ها");

                        System.Threading.Tasks.Task.Run(() => DedicatedFunctions.StringCorrection(loadingForm, doc, standardModels));
                    }
                    else
                        return;
                }

                loadingForm.ShowDialog();
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        public void standardCorrection()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == AccessType.AccessGranted || accessType == AccessType.AccessGranted_Administrator)
            {
                string virastarFolder = Properties.Settings.Default.WorkSpaceDirectory + StringConstant.VirastarFolder;

                List<SearchReplaceModel> standardModels = new List<SearchReplaceModel>();

                if (Directory.Exists(virastarFolder))
                {
                    string[] files = Directory.GetFiles(
                        virastarFolder,
                        $"*{StringConstant.ButtonCodeStandard}.csv"
                        );

                    foreach (string file in files)
                    {
                        using (StreamReader reader = new StreamReader(file, Encoding.UTF8))
                        using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                        {
                            standardModels.AddRange(csv.GetRecords<SearchReplaceModel>());
                        }
                    }
                }

                LoadingForm loadingForm = new LoadingForm();

                if (doc.ActiveWindow.Selection.Words.Count >= 2)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("زبان معیار");

                    System.Threading.Tasks.Task.Run(() => DedicatedFunctions.StringCorrection(loadingForm, doc.ActiveWindow.Selection.Range, standardModels));
                }
                else
                {
                    if (DedicatedFunctions.ShowMessage("متنی برای زبان معیار انتخاب نشده؛ در کل سند اعمال شود؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تبدیل تمام سند به زبان معیار");

                        System.Threading.Tasks.Task.Run(() => DedicatedFunctions.StringCorrection(loadingForm, doc, standardModels));
                    }
                    else
                        return;
                }
                loadingForm.ShowDialog();
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }

        public void convertNumber(bool toPersian)
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            Selection selection = Globals.ThisAddIn.Application.Selection;

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == AccessType.AccessGranted || accessType == AccessType.AccessGranted_Administrator)
            {
                if (selection.Text == null)
                    return;

                Globals.ThisAddIn.DisableEvents = true;

                if (toPersian)
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تبدیل اعداد از انگلیسی به فارسی");
                else
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تبدیل اعداد از فارسی به انگلیسی");

                //for prevent crash Word when document is openend and do and nothing has been done, crashed using UndoRecord and Find.Execute with Repllace All
                Variable tempVariable = doc.Variables.Add("test", "test");
                tempVariable.Delete();


                bool containsDigit = selection.Text.Any(char.IsDigit);
                //MessageBox.Show(selection.StoryType.ToString());

                if (!containsDigit)// || selection.Text.Length == 0
                {
                    Range rng;
                    string rangeStatus = "";
                    if (selection.StoryType == WdStoryType.wdMainTextStory)
                    {
                        rng = doc.Content;
                        rangeStatus = "در کل سند اعمال شود؟";

                    }
                    else if (selection.StoryType == WdStoryType.wdPrimaryFooterStory ||
                        selection.StoryType == WdStoryType.wdFirstPageFooterStory ||
                        selection.StoryType == WdStoryType.wdEvenPagesFooterStory)
                    {
                        rng = selection.HeaderFooter.Range;
                        rangeStatus = "در پاصفحه(Footer) مربوطه اعمال شود؟";

                    }
                    else if (selection.StoryType == WdStoryType.wdPrimaryHeaderStory ||
                        selection.StoryType == WdStoryType.wdFirstPageHeaderStory ||
                        selection.StoryType == WdStoryType.wdEvenPagesHeaderStory)
                    {
                        rng = selection.HeaderFooter.Range;
                        rangeStatus = "در سرصفحه(Header) مربوطه اعمال شود؟";

                    }
                    else
                    {
                        rng = doc.Content;
                        rangeStatus = "در کل سند اعمال شود؟";
                    }

                    if (DedicatedFunctions.ShowMessage("عددی انتخاب نشده؛ " + rangeStatus, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
                    {
                        bool contentContainsDigit = rng.Text.Any(char.IsDigit);

                        if (contentContainsDigit)
                        {
                            DedicatedFunctions.clearFormmattingFind(rng);
                            DedicatedFunctions.changeToStandardDigit(rng, toPersian);

                            DedicatedFunctions.clearFormmattingFind(rng);
                            DedicatedFunctions.changeIntegerNumbers(rng, toPersian);

                            DedicatedFunctions.clearFormmattingFind(rng);
                            DedicatedFunctions.changeDoubleNumbers(rng, toPersian);

                            DedicatedFunctions.clearFormmattingFind(rng);

                            System.Threading.Tasks.Task.Run(() =>
                            {
                                foreach (Field field in rng.Fields)
                                {
                                    field.Update();
                                }
                            });
                        }
                        else
                        {
                            DedicatedFunctions.ShowMessage("عددی در سند یافت نشد!");
                        }
                    }
                }
                else
                {
                    DedicatedFunctions.clearFormmattingFind(selection.Range);
                    DedicatedFunctions.changeToStandardDigit(selection.Range, toPersian);

                    DedicatedFunctions.clearFormmattingFind(selection.Range);
                    DedicatedFunctions.changeIntegerNumbers(selection.Range, toPersian);

                    DedicatedFunctions.clearFormmattingFind(selection.Range);
                    DedicatedFunctions.changeDoubleNumbers(selection.Range, toPersian);

                    DedicatedFunctions.clearFormmattingFind(selection.Range);
                }

                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

                //update Fields if changed numbers, so update and resolve it.
                Globals.ThisAddIn.DisableEvents = false;
            }
        }
        public void poemMode(int column)
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == AccessType.AccessGranted || accessType == AccessType.AccessGranted_Administrator)
            {
                if (column == 1)
                    DedicatedFunctions.poemOneColumn(doc);
                else if (column == 2)
                    DedicatedFunctions.poemTwoColumn(doc);
            }
        }
        #endregion

        #region Bibliography and Sources
        public void sourceManagementDialog()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            try
            {
                Dialog dlg = Globals.ThisAddIn.Application.Dialogs[WdWordDialog.wdDialogSourceManager];
                dlg.Show();
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("نشانگر متن در مکانی غیر قابل ویرایش قرار دارد، لطفا مکانی دیگر از متن را انتخاب کرده و سپس گزینه مورد نظر خود را انتخاب نمایید", (int)ErrorCodes.CursorInLockedSpace);
            }
        }
        public void insertSources()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            //doc.ActiveWindow.SetFocus();
            doc.ActiveWindow.Activate();
            dynamic dlg = Globals.ThisAddIn.Application.Dialogs[WdWordDialog.wdDialogCreateSource];
            //SendKeys.Send("j{ENTER}{TAB}{TAB}");
            dlg.Show();
        }
        public void importSourcesDialog()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                //close Ribbon Gallery
                SendKeys.Send("{ESC}");

                doc.ActiveWindow.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateNormal;
                AddBibliographyForm addBibliography = new AddBibliographyForm();
                addBibliography.ShowDialog();
                doc.Activate();
            }
        }
        public void goToBibliography()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            bool found = false;
            foreach (Field item in doc.Fields)
            {
                if (item.Type == WdFieldType.wdFieldBibliography)
                {
                    found = true;
                    item.Select();
                    Globals.ThisAddIn.Application.Selection.EndKey();
                    Globals.ThisAddIn.Application.Selection.Select();
                    break;
                }
            }

            if (!found)
                DedicatedFunctions.ShowMessage("فهرست منابعی در سند یافت نشد");
        }
        public void insertCitationMenu()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            Selection selection = Globals.ThisAddIn.Application.Selection;
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                if (DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD) == null ||
                !DedicatedFunctions.getTaskPane(doc, TaskPaneTitles.CreateCD).Visible)
                {
                    DedicatedFunctions.RemoveAllTaskPanes(doc);
                    //DedicatedFunctions.RemoveTaskPane(Globals.ThisAddIn.Application.ActiveDocument, TaskPaneTitle.CrossReference);
                    CustomTaskPane customTaskPane = DedicatedFunctions.AddTaskPane(new InsertCitationTaskPane(), doc, TaskPaneTitles.InsertCitation);
                    ((InsertCitationTaskPane)customTaskPane.Control).enableActionPane(customTaskPane);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("عملیات ساخت لوح فشرده در حال انجام است، برای ارجاع دادن، عملیات ساخت لوح فشرده را لغو کنید");
                }
            }
        }
        #endregion

        public void updateDocument()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            LoadingForm loadingForm = new LoadingForm();

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    Selection selection = Globals.ThisAddIn.Application.Selection;

                    DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
                    if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == AccessType.AccessGranted_Administrator)
                    {
                        try
                        {
                            DedicatedFunctions.updateTables(doc, selection, accessType);
                            loadingForm?.closeForm(successfull: true);

                        }
                        catch (Exception)
                        {
                            Globals.ThisAddIn.Application.ScreenUpdating = true;
                            loadingForm?.closeForm(successfull: false);
                        }
                    }
                }
                catch (Exception)
                {
                    loadingForm?.closeForm(successfull: false);
                }
            });
            loadingForm.ShowDialog();
        }

        public void uploadDocument()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            LoadingForm loadingForm = new LoadingForm();

            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

                    if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                    {
                        DedicatedFunctions.uploadDocument(doc, true);
                        loadingForm?.closeForm(successfull: true);
                    }
                    else
                    {
                        DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
                        loadingForm?.closeForm(successfull: false);
                    }
                }
                catch (Exception e)
                {
                    DedicatedFunctions.ShowErrorMessage("خطا غیر منتظره ای در آپلود سند رخ داد" + "\nپیغام خطا:\n" + e.Message);
                    loadingForm?.closeForm(successfull: false);
                }
            });
            loadingForm.ShowDialog();
        }
        #region Export
        public void exportToWord()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.saveDocument(doc);

                System.Windows.Forms.SaveFileDialog fileDialog = new System.Windows.Forms.SaveFileDialog();
                fileDialog.Title = "مکان و نام خروجی فایل خود را مشخص کنید";
                fileDialog.AddExtension = false;
                fileDialog.Filter = "Word Document (*.docx)|*.docx";
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (fileDialog.FileName != doc.FullName)
                    {
                        LoadingForm loadingForm = new LoadingForm();

                        System.Threading.Tasks.Task.Run(async () =>
                        {
                            try
                            {
                                if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                                {
                                    while (!await DedicatedFunctions.getDocumentServerAndUpdateVariables(
                                    doc,
                                    DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UserToken.ToString()),
                                    DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString())))
                                    {
                                        DialogResult dialogResult = DedicatedFunctions.ShowMessage(
                                            ErrorMessages.ErrorServiceUnavailable,
                                            MessageBoxButtons.RetryCancel,
                                            MessageBoxIcon.Error);
                                        if (dialogResult != DialogResult.Retry)
                                        {
                                            return;
                                        }
                                    }
                                }

                                DedicatedFunctions.saveDocument(doc);
                                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("گرفتن خروجی");

                                DedicatedFunctions.saveAsDocument(doc, fileDialog.FileName);

                                if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                                {
                                    Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                                    Ribbon.RibbonControlsVisibility(false);
                                }
                                Globals.ThisAddIn.Application.Caption = Globals.ThisAddIn.Application.Caption.Replace(" (Exported)", "") + " (Exported)";

                                Microsoft.Office.Interop.Word.ContentControl[] nameOFAllahChars = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_InTheNameOfAllah.ToString());
                                if (nameOFAllahChars != null && nameOFAllahChars.Length != 0)
                                {
                                    foreach (var nameOFAllahChar in nameOFAllahChars)
                                    {
                                        DedicatedFunctions.ContentControlToImage(doc, nameOFAllahChar);
                                    }
                                }

                                DedicatedFunctions.setEmbedFonts(doc, true);
                                DedicatedFunctions.unProtectingImportants(doc, true);
                                //Globals.ThisDocument.RemoveDocumentInformation(WdRemoveDocInfoType.wdRDIDocumentProperties);
                                DedicatedFunctions.removeAllDocumentVariables(doc);
                                doc.Password = "";
                                doc.UndoClear();
                                doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, Name: 1).Select();
                                Globals.ThisAddIn.Application.ScreenUpdating = true;
                                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

                                DedicatedFunctions.saveDocument(doc);

                                Globals.ThisAddIn.DisableEvents = false;
                                loadingForm?.closeForm(successfull: true);
                                //Process.Start("explorer.exe" , "/select, \"" + fileDialog.FileName + "\"");
                            }
                            catch (Exception e)
                            {
                                doc.UndoClear();
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای رخ داد \n" + e.Message, email: StringConstant.SupportEmail);
                                loadingForm?.closeForm(successfull: false);
                            }
                        });
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("سند خروجی نمیتواند با سند اختصاصی شما جایگزین شود، لطفا با نامی دیگر سند را ذخیره کنید");
                    }
                }
            }
        }
        public void exportToIdentification()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                DedicatedFunctions.saveDocument(doc);

                System.Windows.Forms.SaveFileDialog fileDialog = new System.Windows.Forms.SaveFileDialog();
                fileDialog.Title = "مکان و نام خروجی فایل خود را مشخص کنید";
                fileDialog.AddExtension = false;
                fileDialog.Filter = "Word Document (*.docx)|*.docx";

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (fileDialog.FileName != doc.FullName)
                    {
                        LoadingForm loadingForm = new LoadingForm();

                        Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("گرفتن خروجی");
                        Globals.ThisAddIn.Application.ScreenUpdating = false;
                        Globals.ThisAddIn.DisableEvents = true;

                        string fileName = fileDialog.FileName;
                        System.Threading.Tasks.Task.Run(async () =>
                        {
                            try
                            {
                                while (!await DedicatedFunctions.getDocumentServerAndUpdateVariables(
                                    doc,
                                    DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UserToken.ToString()),
                                    DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString())))
                                {
                                    DialogResult dialogResult = DedicatedFunctions.ShowMessage(ErrorMessages.ErrorServiceUnavailable, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                                    if (dialogResult != DialogResult.Retry)
                                    {
                                        return;
                                    }
                                }
                                //doc.Save();

                                //save to another file
                                DedicatedFunctions.saveAsDocument(doc, fileName);

                                //Unprotect Content Controls for deleting Sections
                                DedicatedFunctions.unProtectingImportants(doc, false);

                                #region main Part
                                DedicatedFunctions.exportIdentification(doc);
                                #endregion

                                //delete All ContentControls and Variables
                                DedicatedFunctions.removeAllDocumentVariables(doc);
                                DedicatedFunctions.unProtectingImportants(doc, true);
                                //final
                                doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, Name: 1).Select();

                                doc.Password = "";
                                //save changes
                                DedicatedFunctions.saveDocument(doc);

                                //set ui
                                Ribbon.RibbonControlsVisibility(false);

                                Globals.ThisAddIn.Application.Caption = Globals.ThisAddIn.Application.Caption.Replace(" (Exported)", "") + " (Exported)";

                                //clear undo
                                doc.UndoClear();
                                loadingForm?.closeForm(successfull: true);
                            }
                            catch (Exception e)
                            {
                                doc.UndoClear();
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای رخ داد \n" + e.Message, email: StringConstant.SupportEmail);
                                loadingForm?.closeForm(successfull: false);
                            }

                        });
                        loadingForm.ShowDialog();

                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                        Globals.ThisAddIn.DisableEvents = false;
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("سند خروجی نمیتواند با سند اختصاصی شما جایگزین شود، لطفا با نامی دیگر سند را ذخیره کنید");
                    }
                }
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }
        }
        public void exportAsGrayscale()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.saveDocument(doc);

                System.Windows.Forms.SaveFileDialog fileDialog = new System.Windows.Forms.SaveFileDialog();
                fileDialog.Title = "مکان و نام خروجی فایل خود را مشخص کنید";
                fileDialog.AddExtension = false;
                fileDialog.Filter = "Word Document (*.docx)|*.docx";

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (fileDialog.FileName != doc.FullName)
                    {
                        LoadingForm loadingForm = new LoadingForm();

                        Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("گرفتن خروجی");
                        Globals.ThisAddIn.Application.ScreenUpdating = false;
                        Globals.ThisAddIn.DisableEvents = true;

                        string fileName = fileDialog.FileName;

                        System.Threading.Tasks.Task.Run(async () =>
                        {
                            try
                            {
                                if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                                {
                                    while (!await DedicatedFunctions.getDocumentServerAndUpdateVariables(doc, DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UserToken.ToString()), DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString())))
                                    {
                                        DialogResult dialogResult = DedicatedFunctions.ShowMessage(ErrorMessages.ErrorServiceUnavailable, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                                        if (dialogResult != DialogResult.Retry)
                                        {
                                            return;
                                        }
                                    }
                                }

                                //save to another file
                                DedicatedFunctions.saveAsDocument(doc, fileName);

                                //remove specialized Addin
                                DedicatedFunctions.unProtectingImportants(doc, true);
                                DedicatedFunctions.removeAllDocumentVariables(doc);

                                #region main Part
                                //set as Grayscale Contents
                                DedicatedFunctions.convertToGrayScaleContent(doc);
                                DedicatedFunctions.convertToGrayScaleHeaderFooterContent(doc);
                                DedicatedFunctions.convertToGrayScaleOtherContent(doc);

                                DedicatedFunctions.convertToGrayScaleTable(doc);

                                //set as Grayscale Shapes and InlineShapes
                                foreach (InlineShape iShape in Globals.ThisAddIn.Application.ActiveDocument.InlineShapes)
                                {
                                    DedicatedFunctions.convertToGrayScaleShape(iShape);
                                }
                                foreach (Shape shape in Globals.ThisAddIn.Application.ActiveDocument.Shapes)
                                {
                                    DedicatedFunctions.convertToGrayScaleShape(shape);
                                }

                                //set grayscale shapes in header and footer
                                DedicatedFunctions.convertToGrayScaleHeaderFooterShape(doc);
                                //set grayscale shapes in Footnotes and Endnotes and Comments and TextFrame
                                DedicatedFunctions.convertToGrayScaleOtherShape(doc);

                                #endregion

                                //final
                                doc.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, Name: 1).Select();
                                doc.Password = "";

                                //save changes
                                DedicatedFunctions.saveDocument(doc);

                                //set ui
                                if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                                {
                                    Ribbon.RibbonControlsVisibility(false);
                                    Globals.ThisAddIn.Application.Caption = Globals.ThisAddIn.Application.Caption.Replace(" (Exported)", "") + " (Exported)";
                                }

                                //clear undo
                                doc.UndoClear();
                                loadingForm?.closeForm(successfull: true);

                            }
                            catch (Exception e)
                            {
                                doc.UndoClear();
                                doc.Content.Delete();
                                doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای رخ داد \n" + e.Message, email: StringConstant.SupportEmail);
                                loadingForm?.closeForm(successfull: false);
                            }

                        });
                        loadingForm.ShowDialog();

                        Globals.ThisAddIn.Application.ScreenUpdating = true;
                        Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                        Globals.ThisAddIn.DisableEvents = false;
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("سند خروجی نمیتواند با سند اختصاصی شما جایگزین شود، لطفا با نامی دیگر سند را ذخیره کنید");
                    }
                }
            }
            //}
        }
        public void exportToPDF()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.saveDocument(doc);

                System.Windows.Forms.SaveFileDialog fileDialog = new System.Windows.Forms.SaveFileDialog();
                fileDialog.Title = "مکان و نام خروجی فایل خود را مشخص کنید";
                fileDialog.AddExtension = false;
                fileDialog.Filter = "Portable Document Format (*.pdf)|*.pdf";

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                    {
                        LoadingForm loadingForm = new LoadingForm();

                        System.Threading.Tasks.Task.Run(async () =>
                        {
                            try
                            {
                                while (!await DedicatedFunctions.getDocumentServerAndUpdateVariables(
                                        doc,
                                        DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UserToken.ToString()),
                                        DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString())))
                                {
                                    DialogResult dialogResult = DedicatedFunctions.ShowMessage(ErrorMessages.ErrorServiceUnavailable, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                                    if (dialogResult != DialogResult.Retry)
                                    {
                                        loadingForm?.closeForm(successfull: false);
                                        return;
                                    }
                                }
                                //Save File
                                doc.ExportAsFixedFormat(fileDialog.FileName, WdExportFormat.wdExportFormatPDF, true, WdExportOptimizeFor.wdExportOptimizeForPrint, WdExportRange.wdExportAllDocument, 1, 1, WdExportItem.wdExportDocumentContent, true, true, WdExportCreateBookmarks.wdExportCreateNoBookmarks, true, true, false);
                                loadingForm?.closeForm(successfull: true);
                            }
                            catch (Exception)
                            {
                                loadingForm?.closeForm(successfull: false);
                            }
                        });
                        loadingForm.ShowDialog();
                    }
                }
            }
        }
        public void exportToLatex()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType != DedicatedFunctions.AccessType.AccessGranted &&
                accessType != DedicatedFunctions.AccessType.AccessGranted_Administrator)
                return;

            DedicatedFunctions.saveDocument(doc);

            SaveFileDialog fileDialog = new SaveFileDialog();
            fileDialog.Title = "مکان و نام خروجی فایل LaTeX را مشخص کنید";
            fileDialog.AddExtension = true;
            fileDialog.Filter = "LaTeX File (*.tex)|*.tex";

            if (fileDialog.ShowDialog() != DialogResult.OK)
                return;

            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                LoadingForm loadingForm = new LoadingForm();

                System.Threading.Tasks.Task.Run(async () =>
                {
                    try
                    {
                        while (!await DedicatedFunctions.getDocumentServerAndUpdateVariables(
                                doc,
                                DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UserToken.ToString()),
                                DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString())))
                        {
                            DialogResult dialogResult = DedicatedFunctions.ShowMessage(ErrorMessages.ErrorServiceUnavailable, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                            if (dialogResult != DialogResult.Retry)
                            {
                                loadingForm?.closeForm(successfull: false);
                                return;
                            }
                        }

                        // تلاش اولیه: ذخیره موقت HTML با retry (برای مواقعی که Word موقتا پاسخگو نیست)
                        string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");
                        bool savedHtml = false;

                        try
                        {
                            RetryAction(() =>
                            {
                                // SaveAs2 ممکن است خطا دهد اگر دسترسی نداشته باشیم؛ اینجا با retry امتحان می‌کنیم
                                doc.SaveAs2(tempHtmlPath, WdSaveFormat.wdFormatFilteredHTML);
                            }, 5, 200);

                            savedHtml = File.Exists(tempHtmlPath);
                        }
                        catch (Exception saveEx)
                        {
                            // اگر SaveAs2 خطا داد، savedHtml=false خواهد بود و به fallback می‌رویم
                            savedHtml = false;
                        }

                        if (savedHtml)
                        {
                            // خواندن HTML و تبدیل ساده به LaTeX
                            string htmlContent = File.ReadAllText(tempHtmlPath, Encoding.UTF8);
                            string latexContent = ConvertHtmlToLatex(htmlContent);

                            File.WriteAllText(fileDialog.FileName, latexContent, Encoding.UTF8);

                            // پاک کردن فایل موقت
                            try { File.Delete(tempHtmlPath); } catch { }
                            loadingForm?.closeForm(successfull: true);
                            return;
                        }
                        else
                        {
                            // ---------- fallback: ساخت LaTeX مستقیم از Document بدون ذخیره HTML ----------
                            string latexContent = BuildLatexFromDocument(doc);
                            File.WriteAllText(fileDialog.FileName, latexContent, Encoding.UTF8);
                            loadingForm?.closeForm(successfull: true);
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        loadingForm?.closeForm(successfull: false);
                        System.Windows.Forms.MessageBox.Show("خطا در تولید خروجی LaTeX:\n" + ex.Message);
                    }
                });

                loadingForm.ShowDialog();
            }
        }
        // تابع RetryAction برای مقابله با RPC_E_CALL_REJECTED و موارد موقتی
        private void RetryAction(Action action, int retries = 3, int delayMs = 200)
        {
            for (int i = 0; i < retries; i++)
            {
                try
                {
                    action();
                    return;
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    const uint RPC_E_CALL_REJECTED = 0x80010001;
                    if ((uint)ex.ErrorCode == RPC_E_CALL_REJECTED)
                    {
                        Thread.Sleep(delayMs);
                        continue;
                    }
                    throw;
                }
            }
            throw new Exception("Word پاسخ نمی‌دهد یا اجازه‌ی ذخیره‌سازی را نمی‌دهد.");
        }

        // تبدیل ساده HTML -> LaTeX (همان نسخه ساده قبلی)
        private string ConvertHtmlToLatex(string html)
        {
            string latex = html;

            latex = latex.Replace("<b>", "\\textbf{").Replace("</b>", "}");
            latex = latex.Replace("<strong>", "\\textbf{").Replace("</strong>", "}");
            latex = latex.Replace("<i>", "\\textit{").Replace("</i>", "}");
            latex = latex.Replace("<em>", "\\textit{").Replace("</em>", "}");
            latex = latex.Replace("<u>", "\\underline{").Replace("</u>", "}");
            latex = latex.Replace("<br>", "\\\\").Replace("<br/>", "\\\\");
            latex = latex.Replace("<p>", "\n\n").Replace("</p>", "\n");
            latex = latex.Replace("&nbsp;", " ");

            // حذف تگ‌های باقیمانده
            latex = System.Text.RegularExpressions.Regex.Replace(latex, "<.*?>", string.Empty);

            // ساختار پایه LaTeX (پیشنهادی برای استفاده با xelatex)
            var sb = new StringBuilder();
            sb.AppendLine("\\documentclass{article}");
            sb.AppendLine("\\usepackage{fontspec}");
            sb.AppendLine("\\usepackage{polyglossia}");
            sb.AppendLine("\\setmainlanguage{english}");
            sb.AppendLine("\\setotherlanguage{farsi}");
            sb.AppendLine("\\newfontfamily\\persianfont{XB Niloofar}"); // کاربر باید فونت فارسی مناسب را نصب کند یا تغییر دهد
            sb.AppendLine("\\begin{document}");
            sb.AppendLine(latex);
            sb.AppendLine("\\end{document}");

            return sb.ToString();
        }

        // fallback: ساخت LaTeX مستقیماً از Document (متن و ساختار پاراگراف‌ها)
        private string BuildLatexFromDocument(Document doc)
        {
            var sb = new StringBuilder();

            sb.AppendLine("\\documentclass{article}");
            sb.AppendLine("\\usepackage{fontspec}");
            sb.AppendLine("\\usepackage{polyglossia}");
            sb.AppendLine("\\setmainlanguage{english}");
            sb.AppendLine("\\setotherlanguage{farsi}");
            sb.AppendLine("\\setlength{\\parskip}{0.5em}");
            sb.AppendLine("\\begin{document}");
            sb.AppendLine();

            // تلاش برای نگهداری سطوح heading به صورت section/subsection
            foreach (Paragraph para in doc.Paragraphs)
            {
                try
                {
                    string txt = para.Range.Text ?? "";
                    txt = EscapeLatex(txt).TrimEnd('\r', '\n');

                    // سعی می‌کنیم براساس Style نام Heading را تشخیص دهیم
                    string styleName = "";
                    try { styleName = para.get_Style() is Style s ? s.NameLocal : ""; } catch { styleName = ""; }

                    if (!string.IsNullOrEmpty(styleName) && styleName.ToLower().Contains("heading 1"))
                    {
                        sb.AppendLine("\\section{" + txt + "}");
                    }
                    else if (!string.IsNullOrEmpty(styleName) && styleName.ToLower().Contains("heading 2"))
                    {
                        sb.AppendLine("\\subsection{" + txt + "}");
                    }
                    else if (!string.IsNullOrEmpty(styleName) && styleName.ToLower().Contains("heading 3"))
                    {
                        sb.AppendLine("\\subsubsection{" + txt + "}");
                    }
                    else
                    {
                        // اگر متن کوتاه است، فقط یک پاراگراف عادی
                        sb.AppendLine(txt + "\n");
                    }
                }
                catch
                {
                    // بی‌خیال پاراگرافی که خطا داد
                    continue;
                }
            }

            sb.AppendLine();
            sb.AppendLine("\\end{document}");
            return sb.ToString();
        }

        // تابع کمک‌کننده برای Escape کردن کاراکترهای خاص LaTeX
        private string EscapeLatex(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            return input
                .Replace(@"\", @"\textbackslash{}")
                .Replace("&", @"\&")
                .Replace("%", @"\%")
                .Replace("$", @"\$")
                .Replace("#", @"\#")
                .Replace("_", @"\_")
                .Replace("{", @"\{")
                .Replace("}", @"\}")
                .Replace("~", @"\textasciitilde{}")
                .Replace("^", @"\^{}")
                .Replace("—", "--")
                .Replace("–", "-");
        }


        public void exportCDMenu()
        {
            Document doc;
            Selection selection;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
                selection = doc.ActiveWindow.Selection;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted || accessType == AccessType.AccessGranted_Administrator)
            {
                if (File.Exists(doc.FullName))
                    DedicatedFunctions.saveDocument(doc);

                Application.ScreenUpdating = false;
                Range previousRange = Application.Selection.Range;
                Application.UndoRecord.StartCustomRecord("طراحی طرح سیدی");

                //در صورت وجود طرح سیدی قبلی، طرح قبلی پاک شود
                try
                {
                    string tested = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].Name;
                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].Select();

                    if (doc.ActiveWindow.Selection.Sections.Last.Range.Text.Length < 20)
                    {
                        DesigningCDTaskPane.DeleteAndDisableActionPane(doc, null, null);
                    }
                    else
                    {
                        if (doc.Bookmarks["\\Page"].Range.Text.Trim().Replace("\n", "").Replace("\r", "") == "/")
                        {
                            doc.Bookmarks["\\Page"].Range.Delete();
                            selection.TypeBackspace();
                            selection.TypeBackspace();
                        }
                        else
                            doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].Delete();
                    }
                }
                catch (Exception)
                {
                }

                //رفتن به آخر سند
                doc.Bookmarks["\\EndOfDoc"].Range.InsertBreak(WdBreakType.wdSectionBreakNextPage);
                doc.Bookmarks["\\EndOfDoc"].Select();
                doc.Bookmarks["\\EndOfDoc"].Range.set_Style(StyleNames.styleNormal);
                doc.Bookmarks["\\EndOfDoc"].Range.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                DedicatedFunctions.insertParagraph(selection);

                try
                {
                    foreach (HeaderFooter item in doc.Sections.Last.Headers)
                    {
                        item.LinkToPrevious = false;
                        //item.Range.Text = "";
                    }
                    foreach (HeaderFooter item in doc.Sections.Last.Footers)
                    {
                        item.LinkToPrevious = false;
                        //item.Range.Text = "";
                    }
                }
                catch (Exception)
                {
                }
                doc.Sections.Last.PageSetup.BottomMargin = 0;
                doc.Sections.Last.PageSetup.LeftMargin = 0;
                doc.Sections.Last.PageSetup.RightMargin = 0;
                doc.Sections.Last.PageSetup.TopMargin = 0;
                doc.Sections.Last.PageSetup.HeaderDistance = 0;
                doc.Sections.Last.PageSetup.FooterDistance = 0;

                DedicatedFunctions.changeParagraphAlignment(selection, WdParagraphAlignment.wdAlignParagraphCenter);
                DedicatedFunctions.changeParagraphIndent(selection, 0, 0, 0);

                //ساخت سیدی و بارگیری Taskpane
                DedicatedFunctions.RemoveAllTaskPanes(doc);
                //DedicatedFunctions.RemoveTaskPane(Globals.ThisAddIn.Application.ActiveDocument, TaskPaneTitle.CrossReference);
                CustomTaskPane customTaskPane = DedicatedFunctions.AddTaskPane(new DesigningCDTaskPane(doc, previousRange), doc, TaskPaneTitles.CreateCD);
                DesigningCDTaskPane taskPaneControl = (DesigningCDTaskPane)customTaskPane.Control;
                taskPaneControl.enableActionPane(customTaskPane);
            }
        }
        #endregion

        #region Settings
        public void formatSettings()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                FormatSettingsForm form = new FormatSettingsForm(Globals.ThisAddIn.Application.ActiveDocument, accessType);
                form.ShowDialog();
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }

        }
        public void captionSettings()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                CaptionsReferSettingsForm form = new CaptionsReferSettingsForm(Globals.ThisAddIn.Application.ActiveDocument);
                form.Show();
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }

        }
        public void virastarSettings()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            //if(accessType == DedicatedFunctions.AccessType.AccessGranted)
            //{
            VirastarSettingsForm form = new VirastarSettingsForm(Globals.ThisAddIn.Application.ActiveDocument);
            form.ShowDialog();
            //}
            //else if(accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            //{
            //	DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            //}

        }
        public void footnoteSettings()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                FootnoteSettingsForm form = new FootnoteSettingsForm(Globals.ThisAddIn.Application.ActiveDocument);
                form.ShowDialog();
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }

        }
        public void listsManagerSettings()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                DocumentSettingsForm form = new DocumentSettingsForm(Globals.ThisAddIn.Application.ActiveDocument);
                form.ShowDialog();
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }

        }
        public void citationSettings()
        {
            Document doc;
            try
            {
                doc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (Exception)
            {
                return;
            }
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);
            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                CitationSettingsForm form = new CitationSettingsForm(Globals.ThisAddIn.Application.ActiveDocument, accessType);
                form.ShowDialog();
            }
            else if (accessType == DedicatedFunctions.AccessType.AccessGranted_Administrator)
            {
                DedicatedFunctions.ShowMessage(DialogBoxMessages.RequiredDedicatedDocument);
            }
        }
        #endregion

        #endregion
    }
}
