using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Forms.AddRemovePages.Models;
using ShivaNegar.Forms.ShivaNegarManager.CreateDocument;
using ShivaNegar.Templates;
using static ShivaNegar.DedicatedFunctions;
using static ShivaNegar.Models.TemplateRelationshipModel;

namespace ShivaNegar.Forms.AddRemovePages
{
    /// <summary>
    /// Interaction logic for AddRemovePagesControl.xaml
    /// </summary>
    public partial class AddRemovePagesControl : UserControl, Interfaces.IStatusFormRequest
    {
        public List<AddRemovePageRelationModel> pageRelations;

        public List<AddRemovePageRelationModel> chapterPageRelations;

        int lastExistingChapterIndex = 0;
        int lastChapterIndex = 0;
        //public List<AddRemovePageRelationModel> chapterNotMadePageRelations;
        //public List<AddRemovePageRelationModel> preMadeChapterDeletedRelations;

        Document doc;

        public Action CloseFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action NormalStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }

        bool isUpdatingDocument = false;
        #region constructor

        private void InitializeMaterialDesign()
        {
            // Create dummy objects to force the MaterialDesign assemblies to be loaded
            // from this assembly, which causes the MaterialDesign assemblies to be searched
            // relative to this assembly's path. Otherwise, the MaterialDesign assemblies
            // are searched relative to Eclipse's path, so they're not found.
            var card = new Card();
            var hue = new Hue("Dummy", Colors.Black, Colors.White);
        }

        public AddRemovePagesControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();
        }
        public AddRemovePagesControl(Document doc)
        {
            InitializeMaterialDesign();
            InitializeComponent();

            pageRelations = new List<AddRemovePageRelationModel>();

            chapterPageRelations = new List<AddRemovePageRelationModel>();
            //chapterNotMadePageRelations = new List<AddRemovePageRelationModel>();
            //preMadeChapterDeletedRelations = new List<AddRemovePageRelationModel>();

            this.doc = doc;

            btnCloseApp.Click += BtnCloseApp_Click;
            //btnMaximize.Click += BtnMaximize_Click;
            //btnMinimize.Click += BtnMinimize_Click;

            btnAddChapter.Click += BtnAddChapter_Click;
            btnRemoveChapter.Click += BtnRemoveChapter_Click;
            btnConfirmChanges.Click += BtnConfirmChanges_Click;


            Universities university = DedicatedFunctions.getUniversity(doc);
            DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
            TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

            initialization(doc, university, templateType, documentType);
        }


        #endregion

        #region Functions
        internal void insertPage(Document doc, Universities university, TemplateTypes templateType, DocumentTypes documentType, AddRemovePageRelationModel model)
        {
            Globals.ThisAddIn.DisableSelectionChangedEvent = true;
            try
            {
                //copy and get specified file template path
                Stream stream = DedicatedFunctions.getStream(model.ResourcePath + model.FileName);
                string path = DedicatedFunctions.copyFileToTempFolder(stream, model.FileName);

                //get previous PageID
                PageIDs previousPageID = getLastSectionPageID(doc, university, templateType, documentType, model.PageID, WhichIndex.Previous);
                int previousSectionIndex = DedicatedFunctions.getPageIDIndex(doc, previousPageID.ToString());
                string previousVariableValuePageID = DedicatedFunctions.getStaticVariableValue(doc, previousPageID.ToString());

                //previous is LastPage, so add Section Break for Previous Page and no SectionBreak for Current PageID
                if (previousVariableValuePageID == "LastPage")
                {
                    //go end of previous Page
                    Range rng = doc.Sections[previousSectionIndex].Range;
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd);
                    rng.Select();

                    DedicatedFunctions.insertSectionBreak(doc, DedicatedFunctions.GetCustomProperties(path));

                    //select line and reduce font size for protecting
                    Range previousRange = doc.ActiveWindow.Selection.Range;
                    doc.ActiveWindow.Selection.MoveLeft(WdUnits.wdCharacter, 1);
                    string contentControlID = DedicatedFunctions.protectSectionBreak(doc, previousPageID.ToString());
                    //restore Selection to previous Range
                    previousRange.Select();

                    DedicatedFunctions.setORAddStaticVariableValue(doc, previousPageID.ToString(), contentControlID);
                    DedicatedFunctions.setORAddStaticVariableValue(doc, model.PageID.ToString(), "LastPage");
                }
                else
                {
                    //get previousContentControl

                    ContentControls ccs = doc.SelectContentControlsByTag(previousPageID.ToString());
                    if (ccs != null)
                    {
                        Microsoft.Office.Interop.Word.ContentControl previousContentControl = ccs[1];
                        //previousContentControl is now currentContentControl
                        previousContentControl.Tag = model.PageID.ToString();
                        DedicatedFunctions.setORAddStaticVariableValue(doc, model.PageID.ToString(), previousContentControl.ID);

                        Range rng = previousContentControl.Range;
                        //go to up of protected SectionBreak for create another SectionBreak
                        rng.Collapse(WdCollapseDirection.wdCollapseStart);
                        rng.MoveEnd(WdUnits.wdCharacter, -2);
                        rng.Select();
                        DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection);
                    }

                    //now create SectionBreak and ContentControl for previousPageID
                    DedicatedFunctions.insertSectionBreak(doc, DedicatedFunctions.GetCustomProperties(path));

                    Range previousRange = doc.ActiveWindow.Selection.Range;
                    doc.ActiveWindow.Selection.MoveLeft(WdUnits.wdCharacter, 1);
                    string contentControlID = DedicatedFunctions.protectSectionBreak(doc, previousPageID.ToString());
                    //restore Selection to previous Range
                    previousRange.Select();

                    DedicatedFunctions.setORAddStaticVariableValue(doc, previousPageID.ToString(), contentControlID);
                }

                //copy and create specified page with file template
                DocumentFormat.OpenXml.CustomProperties.Properties properties = GetCustomPropertiesWithStream(path);
                insertDocumentFileAddRemove(doc, path, properties);

                //remove specified file template
                if (!Path.GetFullPath(new FileInfo(path).Directory.FullName).TrimEnd('\\').Contains(Path.GetFullPath(Properties.Settings.Default.WorkSpaceDirectory + StringConstant.DocumentsTemplateFolder).TrimEnd('\\')))
                {
                    DedicatedFunctions.removeFileFromSystem(path);
                }
            }
            catch (Exception e)
            {
                throw new FileLoadException("خطا در ایجاد صفحه" + "\n" + e.Message);
            }
            Globals.ThisAddIn.DisableSelectionChangedEvent = false;
        }
        //internal void insertPage(Document doc , PageIDs _pageID , int _order , SubTemplateIDs stt)
        //{
        //	Globals.ThisAddIn.DisableSelectionChangedEvent = true;
        //	try
        //	{
        //		//copy and get specified file template path
        //		Stream templateXMLStream = DedicatedFunctions.getTemplateDocumentStream(stt , currentUniversity);
        //		string path = DedicatedFunctions.copyFileToTempFolder(templateXMLStream , DedicatedFunctions.getTemplateFileName(stt , currentUniversity));

        //		//get previous PageID
        //		PageIDs previousPageID = getLastSectionPageID(doc , _pageID , WhichIndex.Previous);
        //		int previousSectionIndex = DedicatedFunctions.getPageIDIndex(doc , previousPageID.ToString());
        //		string previousVariableValuePageID = DedicatedFunctions.getStaticVariableValue(doc , previousPageID.ToString());

        //		//previous is LastPage, so add Section Break for Previous Page and no SectionBreak for Current PageID
        //		if(previousVariableValuePageID == "LastPage")
        //		{
        //			//go end of previous Page
        //			doc.Sections[previousSectionIndex].Range.Select();
        //			doc.ActiveWindow.Selection.EndKey();
        //			DedicatedFunctions.insertSectionBreak(doc , DedicatedFunctions.GetCustomProperties(path));

        //			//select line and reduce font size for protecting
        //			Range previousRange = doc.ActiveWindow.Selection.Range;
        //			doc.ActiveWindow.Selection.MoveLeft(WdUnits.wdCharacter , 1);
        //			string contentControlID = DedicatedFunctions.protectSectionBreak(doc , previousPageID.ToString());
        //			//restore Selection to previous Range
        //			previousRange.Select();

        //			DedicatedFunctions.setORAddStaticVariableValue(doc , previousPageID.ToString() , contentControlID);
        //			DedicatedFunctions.setORAddStaticVariableValue(doc , _pageID.ToString() , "LastPage");
        //		}
        //		else
        //		{
        //			//go start of after Page
        //			doc.Sections[previousSectionIndex + 1].Range.Select();
        //			doc.ActiveWindow.Selection.HomeKey();
        //			Range previousRange = doc.ActiveWindow.Selection.Range;

        //			DedicatedFunctions.insertSectionBreak(doc , DedicatedFunctions.GetCustomProperties(path));

        //			//sett LinkToPrevious false in after page
        //			foreach(HeaderFooter headerFooter in doc.Sections[previousSectionIndex + 2].Headers)
        //			{
        //				if(headerFooter.Range.Text.Contains(SettingValues.NullInHeaderFooterTemplates))
        //				{
        //					headerFooter.LinkToPrevious = false;
        //				}
        //			}
        //			foreach(HeaderFooter headerFooter in doc.Sections[previousSectionIndex + 2].Footers)
        //			{
        //				if(headerFooter.Range.Text.Contains(SettingValues.NullInHeaderFooterTemplates))
        //				{
        //					headerFooter.LinkToPrevious = false;
        //				}
        //			}

        //			//go to start of created Section
        //			previousRange.Select();
        //		}

        //		//copy and create specified page with file template
        //		insertDocumentFileAddRemove(doc , path);

        //		if(previousVariableValuePageID != "LastPage")
        //		{
        //			//select new Section
        //			doc.Sections[previousSectionIndex + 1].Range.Select();
        //			doc.ActiveWindow.Selection.EndKey();
        //			string contentControlID = DedicatedFunctions.protectSectionBreak(doc , _pageID.ToString());

        //			DedicatedFunctions.setORAddStaticVariableValue(doc , _pageID.ToString() , contentControlID);
        //		}

        //		//remove specified file template
        //		if(!Path.GetFullPath(new FileInfo(path).Directory.FullName).TrimEnd('\\').Contains(Path.GetFullPath(Properties.Settings.Default.WorkSpaceDirectory + StringConstant.DocumentsTemplateFolder).TrimEnd('\\')))
        //		{
        //			DedicatedFunctions.removeFileFromSystem(path);
        //		}

        //		//if(sectionCount == -1)//error 
        //		//{
        //		//	DedicatedFunctions.ShowErrorMessage("خطایی رخ داده، لطفا مورد را به ما گزارش دهید");
        //		//	throw new Exception("Unexpected error");
        //		//}
        //		//else
        //		//{
        //		//	int currentSectionIndex = DedicatedFunctions.getPageIDIndex(doc , _pageID.ToString());
        //		//
        //		//	notifySectionIndexChanged(doc , currentSectionIndex , sectionCount , _pageID.ToString());
        //		//}

        //	}
        //	catch(Exception e)
        //	{
        //		throw new FileLoadException("خطا در ایجاد صفحه" + "\n" + e.Message);
        //	}
        //	Globals.ThisAddIn.DisableSelectionChangedEvent = false;
        //}

        internal void deletePage(Document doc, Universities university, TemplateTypes templateType, DocumentTypes documentType, PageIDs pageID, int order)
        {
            Globals.ThisAddIn.DisableSelectionChangedEvent = true;

            //get previousLastSectionIndex

            PageIDs previousPageID = getLastSectionPageID(doc, university, templateType, documentType, pageID, WhichIndex.Previous);
            int previousLastSectionIndex = DedicatedFunctions.getPageIDIndex(doc, previousPageID.ToString());

            //get currentSectionIndex
            int lastSectionIndex = DedicatedFunctions.getPageIDIndex(doc, pageID.ToString());
            int firstSectionIndex = previousLastSectionIndex + 1;

            //delete ContentControlPageID for _pageID and data inside not delete

            ContentControls ccs = doc.SelectContentControlsByTag(pageID.ToString());
            if (ccs != null)
            {
                foreach (Microsoft.Office.Interop.Word.ContentControl currentContentControlPageID in ccs)
                {
                    currentContentControlPageID.LockContentControl = false;
                    currentContentControlPageID.LockContents = false;
                    currentContentControlPageID.Delete(false);
                }
            }
            DedicatedFunctions.removeVariable(doc, pageID.ToString());

            int diffreneceSections = lastSectionIndex - firstSectionIndex;
            //delete Sections between previous last page Section and current last page Section
            for (int i = diffreneceSections; i >= 0; i--)
            {
                DedicatedFunctions.deleteSection(doc, doc.Sections[firstSectionIndex]);
            }

            //is last Section
            if (doc.Sections.Count == previousLastSectionIndex)
            {
                ContentControls ccs2 = doc.SelectContentControlsByTag(previousPageID.ToString());
                if (ccs2 != null)
                {
                    foreach (Microsoft.Office.Interop.Word.ContentControl previousContentControlPageID in ccs2)
                    {
                        previousContentControlPageID.LockContentControl = false;
                        previousContentControlPageID.LockContents = false;
                        Range rng = doc.Range(previousContentControlPageID.Range.Start, doc.Content.End);
                        previousContentControlPageID.Delete(true);
                        rng.Delete();
                    }
                }
                DedicatedFunctions.setORAddStaticVariableValue(doc, previousPageID.ToString(), "LastPage");
            }
            Globals.ThisAddIn.DisableSelectionChangedEvent = false;
        }
        //internal static void deletePage(Document doc , VariablePageIDs vpn)
        //{
        //	List<int> sectionIndex = DedicatedFunctions.getVariablePageNameIndexes(doc , vpn.ToString());

        //	for(int i = sectionIndex.Count - 1 ; i >= 0 ; i--)
        //	{
        //		DedicatedFunctions.deleteSection(doc , doc.Sections[sectionIndex[i]]);
        //	}
        //	notifySectionIndexChanged(doc , sectionIndex[0] , -sectionIndex.Count , vpn.ToString());//TODO: not work if index not Continuous,example: 1,3,4,6
        //	DedicatedFunctions.removeVariable(doc , vpn.ToString());
        //}

        internal static void insertDocumentFileAddRemove(Document doc, string path, DocumentFormat.OpenXml.CustomProperties.Properties properties)
        {
            Selection selection = doc.ActiveWindow.Selection;

            DedicatedFunctions.setSectionStartFromBreakType(selection.Sections[1], DedicatedFunctions.getBreakTypeFromCustomDocumentProperties(properties));


            selection.InsertFile(path);

            Globals.ThisAddIn.DisableEvents = true;
            Document specifiedDocumentPage = Globals.ThisAddIn.Application.Documents.Open(path, Visible: false);
            Globals.ThisAddIn.DisableEvents = false;

            int templateSectionCount = specifiedDocumentPage.Sections.Count;
            int lastSectionIndex = selection.Sections[1].Index;//TODO:what? Sections[1].Index ? LastIndex?

            //Doc.Sections Index start at 1, use (templateSectionCount - 1) converts Index to 0
            int i = templateSectionCount - 1;

            foreach (Section specifiedSection in specifiedDocumentPage.Sections)
            {
                Section currentSection = doc.Sections[lastSectionIndex - i];
                currentSection.PageSetup.DifferentFirstPageHeaderFooter = specifiedSection.PageSetup.DifferentFirstPageHeaderFooter;
                //currentSection.Borders = specifiedSection.Borders;

                //set custom Page Setup Setted in Custom Properties
                if (properties != null)
                {
                    foreach (DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty property in properties)
                    {
                        if (property.Name.Value.Equals(Constants.CustomProperties.Custom_PageSetup.ToString(), StringComparison.OrdinalIgnoreCase))
                        {
                            currentSection.PageSetup = specifiedSection.PageSetup;
                        }
                    }
                }

                HeaderFooter firstCurrentHeaderFooter = null;
                HeaderFooter firstTemplateHeaderFooter = null;
                for (int x = 1; x <= 3; x++)//Headers
                {
                    WdHeaderFooterIndex whfi = (WdHeaderFooterIndex)x;

                    if (specifiedSection.Headers[whfi].Exists)
                    {
                        HeaderFooter hf = currentSection.Headers[whfi];
                        hf.LinkToPrevious = false;
                        //hf.Range.Delete();
                        hf.Range.FormattedText = specifiedSection.Headers[whfi].Range.FormattedText;

                        if (hf.Range.Text.Contains(SettingValues.NullInHeaderFooterTemplates))
                        {
                            hf.Range.Delete();
                        }
                        else
                        {
                            if (hf.Range.Paragraphs.Count >= 2)
                                hf.Range.Paragraphs[hf.Range.Paragraphs.Count].Range.Delete();
                        }

                        if (firstCurrentHeaderFooter == null)
                        {
                            firstCurrentHeaderFooter = hf;
                            firstTemplateHeaderFooter = specifiedSection.Headers[whfi];
                        }
                    }
                }
                for (int x = 1; x <= 3; x++)//Footers
                {
                    WdHeaderFooterIndex whfi = (WdHeaderFooterIndex)x;

                    if (specifiedSection.Footers[whfi].Exists)
                    {
                        HeaderFooter hf = currentSection.Footers[whfi];
                        hf.LinkToPrevious = false;
                        //hf.Range.Delete();
                        hf.Range.FormattedText = specifiedSection.Footers[whfi].Range.FormattedText;

                        if (hf.Range.Text.Contains(SettingValues.NullInHeaderFooterTemplates))
                        {
                            hf.Range.Delete();
                        }
                        else
                        {
                            if (hf.Range.Paragraphs.Count >= 2)
                                hf.Range.Paragraphs[hf.Range.Paragraphs.Count].Range.Delete();
                        }

                        if (firstCurrentHeaderFooter == null)
                        {
                            firstCurrentHeaderFooter = hf;
                            firstTemplateHeaderFooter = specifiedSection.Headers[whfi];
                        }
                    }
                }

                if (firstCurrentHeaderFooter != null)
                {
                    firstCurrentHeaderFooter.PageNumbers.ChapterPageSeparator = firstTemplateHeaderFooter.PageNumbers.ChapterPageSeparator;
                    //firstCurrentHeaderFooter.PageNumbers.DoubleQuote = firstTemplateHeaderFooter.PageNumbers.DoubleQuote;
                    //firstCurrentHeaderFooter.PageNumbers.IncludeChapterNumber = firstTemplateHeaderFooter.PageNumbers.IncludeChapterNumber;
                    firstCurrentHeaderFooter.PageNumbers.HeadingLevelForChapter = firstTemplateHeaderFooter.PageNumbers.HeadingLevelForChapter;
                    firstCurrentHeaderFooter.PageNumbers.NumberStyle = firstTemplateHeaderFooter.PageNumbers.NumberStyle;
                    firstCurrentHeaderFooter.PageNumbers.RestartNumberingAtSection = firstTemplateHeaderFooter.PageNumbers.RestartNumberingAtSection;
                    firstCurrentHeaderFooter.PageNumbers.StartingNumber = firstTemplateHeaderFooter.PageNumbers.StartingNumber;
                }

                i--;
            }
            DedicatedFunctions.closeDocument(specifiedDocumentPage, WdSaveOptions.wdDoNotSaveChanges, false);
        }

        #endregion

        private Dictionary<string, string> getValueList(Document doc, Universities university, TemplateTypes templateType, DocumentTypes documentType)
        {
            Microsoft.Office.Interop.Word.ContentControl[] nameOfAllahContentControls = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_InTheNameOfAllah.ToString());

            string nameOfAllah;
            if (nameOfAllahContentControls != null && nameOfAllahContentControls.Length > 0)
            {
                nameOfAllah = nameOfAllahContentControls.First().Range.Text;
            }
            else
            {
                nameOfAllah = "";
            }

            string documentTypePersian = DedicatedFunctions.getDocumentTypePersianName(documentType);
            string documentTypeEnglish = DedicatedFunctions.getDocumentTypeEnglishName(documentType);
            string templateTypePersian = DedicatedFunctions.getTemplateTypePersianName(templateType);
            string templateTypeEnglish = DedicatedFunctions.getTemplateTypePersianName(templateType);

            string universityFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_University_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string universityEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_University_En.ToString()).Replace(SettingValues.NotExist, "");
            string branchFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Branch_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string branchEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Branch_En.ToString()).Replace(SettingValues.NotExist, "");
            string departmentFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Department_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string departmentEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Department_En.ToString()).Replace(SettingValues.NotExist, "");
            string groupFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Group_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string groupEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Group_En.ToString()).Replace(SettingValues.NotExist, "");
            string fieldOfStudyFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_FieldOfStudy_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string fieldOfStudyEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_FieldOfStudy_En.ToString()).Replace(SettingValues.NotExist, "");
            string areaOfStudyFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_AreaOfStudy_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string areaOfStudyEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_AreaOfStudy_En.ToString()).Replace(SettingValues.NotExist, "");
            string academicDegreeFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string academicDegreeEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_AcademicDegree_En.ToString()).Replace(SettingValues.NotExist, "");

            string nameOfCourseFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_NameOfCourse_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string titleFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Title_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string titleEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Title_En.ToString()).Replace(SettingValues.NotExist, "");
            string authorFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Author_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string authorEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Author_En.ToString()).Replace(SettingValues.NotExist, "");
            string advisorFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Advisor_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string advisorEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Advisor_En.ToString()).Replace(SettingValues.NotExist, "");
            string supervisorFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Supervisor_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string supervisorEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Supervisor_En.ToString()).Replace(SettingValues.NotExist, "");
            string defenseDateFa = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_DefenseDate_Fa.ToString()).Replace(SettingValues.NotExist, "");
            string defenseDateEn = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_DefenseDate_En.ToString()).Replace(SettingValues.NotExist, "");


            string advisorTitleEn = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Advisor_Title_En);
            string advisorTitleFa = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Advisor_Title_Fa);
            string areaOfStudyTitleEn = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_AreaOfStudy_Title_En);
            string areaOfStudyTitleFa = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_AreaOfStudy_Title_Fa);

            //string authorTitleEn = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Author_Title_En);
            //string authorTitleFa = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Author_Title_Fa);
            //string supervisorTitleEn = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Supervisor_Title_En);
            //string supervisorTitleFa = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Supervisor_Title_Fa);
            //string titleTitleEn = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Title_Title_En);
            //string titleTitleFa = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Title_Title_Fa);

            if (string.IsNullOrEmpty(advisorEn.Trim()))
                advisorTitleEn = "";
            if (string.IsNullOrEmpty(advisorFa.Trim()))
                advisorTitleFa = "";
            if (string.IsNullOrEmpty(areaOfStudyEn.Trim()))
                areaOfStudyTitleEn = "";
            if (string.IsNullOrEmpty(areaOfStudyFa.Trim()))
                areaOfStudyTitleFa = "";

            return new Dictionary<string, string>()
                {
                    { ContentControlNames._field_Chapter1_Title.ToString(),""},
                    { ContentControlNames._field_Chapter2_Title.ToString(),""},
                    { ContentControlNames._field_Chapter3_Title.ToString(),""},
                    { ContentControlNames._field_Chapter4_Title.ToString(),""},
                    { ContentControlNames._field_Chapter5_Title.ToString(),""},
                    { ContentControlNames._field_Chapter6_Title.ToString(),""},
                    { ContentControlNames._field_Chapter7_Title.ToString(),""},
                    { ContentControlNames._field_Chapter8_Title.ToString(),""},
                    { ContentControlNames._field_Chapter9_Title.ToString(),""},
                    { ContentControlNames._field_Chapter10_Title.ToString(),""},

                    { ContentControlNames._field_Hidden_Chapter1_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter2_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter3_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter4_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter5_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter6_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter7_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter8_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter9_Title.ToString(),""},
                    { ContentControlNames._field_Hidden_Chapter10_Title.ToString(),""},

                    { ContentControlNames._field_Abstract_En.ToString(),""},
                    { ContentControlNames._field_Abstract_Fa.ToString(),""},
                    { ContentControlNames._field_Keywords_En.ToString(),""},
                    { ContentControlNames._field_Keywords_Fa.ToString(),""},
                    { ContentControlNames._field_Dedication_Fa.ToString(),""},
                    { ContentControlNames._field_Acknowledgment_Fa.ToString(),""},
                    { ContentControlNames._field_Icon_University.ToString(),Constants.Dictionaries.iconMark},

                    { ContentControlNames._field_Advisor_En.ToString(),DedicatedFunctions.getStaticVariableValue(doc,VariableFieldIDs._variable_field_Advisor_En.ToString())},
                    { ContentControlNames._field_Advisor_Fa.ToString(),advisorFa},
                    { ContentControlNames._field_Author_En.ToString(),authorEn},
                    { ContentControlNames._field_Author_Fa.ToString(),authorFa},
                    { ContentControlNames._field_DefenseDate_En.ToString(),defenseDateEn},
                    { ContentControlNames._field_DefenseDate_Fa.ToString(),defenseDateFa},
                    { ContentControlNames._field_Department_En.ToString(),departmentEn},
                    { ContentControlNames._field_Department_Fa.ToString(),departmentFa},
                    { ContentControlNames._field_FieldOfStudy_En.ToString(),fieldOfStudyEn},
                    { ContentControlNames._field_FieldOfStudy_Fa.ToString(),fieldOfStudyFa},
                    { ContentControlNames._field_AreaOfStudy_En.ToString(),areaOfStudyEn},
                    { ContentControlNames._field_AreaOfStudy_Fa.ToString(),areaOfStudyFa},
                    { ContentControlNames._field_Group_En.ToString(),groupEn},
                    { ContentControlNames._field_Group_Fa.ToString(),groupFa},
                    { ContentControlNames._field_AcademicDegree_En.ToString(),academicDegreeEn},
                    { ContentControlNames._field_AcademicDegree_Fa.ToString(),academicDegreeFa},
                    { ContentControlNames._field_InTheNameOfAllah.ToString(),nameOfAllah},
                    { ContentControlNames._field_NameOfCourse_Fa.ToString(),nameOfCourseFa},

                    { ContentControlNames._field_Advisor_Title_En.ToString(),advisorTitleEn},
                    { ContentControlNames._field_Advisor_Title_Fa.ToString(),advisorTitleFa},
                    { ContentControlNames._field_AreaOfStudy_Title_Fa.ToString(),areaOfStudyTitleFa},
                    { ContentControlNames._field_AreaOfStudy_Title_En.ToString(),areaOfStudyTitleEn},

					//{ ContentControlNames._field_Author_Title_En.ToString(),authorTitleEn},
					//{ ContentControlNames._field_Author_Title_Fa.ToString(),authorTitleFa},
					//{ ContentControlNames._field_Supervisor_Title_Fa.ToString(),supervisorTitleFa},
					//{ ContentControlNames._field_Supervisor_Title_En.ToString(),supervisorTitleEn},
					//{ ContentControlNames._field_Title_Title_En.ToString(),titleTitleEn},
					//{ ContentControlNames._field_Title_Title_Fa.ToString(),titleTitleFa},

					{ ContentControlNames._field_Supervisor_En.ToString(),supervisorEn},
                    { ContentControlNames._field_Supervisor_Fa.ToString(),supervisorFa},
                    { ContentControlNames._field_Title_En.ToString(),titleEn},
                    { ContentControlNames._field_Title_Fa.ToString(),titleFa},
                    { ContentControlNames._field_University_En.ToString(),universityEn},
                    { ContentControlNames._field_University_Fa.ToString(),universityFa},
                    { ContentControlNames._field_Branch_En.ToString(),branchEn},
                    { ContentControlNames._field_Branch_Fa.ToString(),branchFa},
                    { ContentControlNames._field_Type_En.ToString(),documentTypeEnglish},
                    { ContentControlNames._field_Type_Fa.ToString(),documentTypePersian},

                    { ContentControlNames._field_DefenseLocation_Fa.ToString(),""},
                    { ContentControlNames._field_Examiner_Fa.ToString(),""},
                    { ContentControlNames._field_Examiner_Title_Fa.ToString(),""},


					//{ BookmarkNames.bookmark_Faculty_Fa,""},
					//{ BookmarkNames.bookmark_Faculty_En,""},
				};
        }


        #region buttons
        private async void BtnConfirmChanges_Click(object sender, System.Windows.RoutedEventArgs es)
        {
            try
            {
                btnConfirmChanges.IsEnabled = false;

                isUpdatingDocument = true;
                loadingCreating.IsEnabled = true;
                loadingCreating.Tag = "لطفا منتظر بمانید";
                await System.Threading.Tasks.Task.Run(() => Thread.Sleep(1000));
                bool onChanged = false;
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                Universities university = DedicatedFunctions.getUniversity(doc);
                DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
                TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

                #region Pages
                foreach (AddRemovePageRelationModel pr in pageRelations)
                {
                    if (pr.Control is ToggleButton toggleButton)
                    {
                        if (((bool)pr.IsPageExist) != toggleButton.IsChecked)//Changed
                        {
                            if (toggleButton.IsChecked == true)//create Page
                            {
                                loadingCreating.Tag = "افزودن " + pr.PageTitle;

                                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("افزودن " + pr.PageTitle);
                                await System.Threading.Tasks.Task.Run(() => insertPage(doc, university, templateType, documentType, pr));
                                //insertPage(doc , pr.VariableTag , pr.SubTemplateID , getValuesForBookmark);
                                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                                onChanged = true;
                            }
                            else // remove Page
                            {
                                loadingCreating.Tag = "حذف " + pr.PageTitle;

                                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("حذف " + pr.PageTitle);
                                await System.Threading.Tasks.Task.Run(() => deletePage(doc, university, templateType, documentType, pr.PageID, pr.Order));
                                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                                onChanged = true;
                            }
                        }
                    }
                    else
                    {
                        throw new Exception();
                    }
                }


                if (lastChapterIndex < lastExistingChapterIndex)//remove
                {
                    for (int i = lastExistingChapterIndex; i > lastChapterIndex; i--)
                    {
                        loadingCreating.Tag = "حذف " + chapterPageRelations[i].PageTitle;

                        Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("حذف " + chapterPageRelations[i].PageTitle);
                        await System.Threading.Tasks.Task.Run(() => deletePage(doc, university, templateType, documentType, chapterPageRelations[i].PageID, chapterPageRelations[i].Order));
                        Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                        onChanged = true;
                    }
                }
                else if (lastChapterIndex > lastExistingChapterIndex)//add
                {
                    for (int i = lastExistingChapterIndex + 1; i <= lastChapterIndex; i++)
                    {
                        loadingCreating.Tag = "ایجاد " + chapterPageRelations[i].PageTitle;

                        Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("ایجاد " + chapterPageRelations[i].PageTitle);
                        await System.Threading.Tasks.Task.Run(() => insertPage(doc, university, templateType, documentType, chapterPageRelations[i]));
                        Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                        onChanged = true;
                    }
                }

                //foreach(AddRemovePageRelationModel preChapterDeleted in preMadeChapterDeletedRelations)//Remove pre-made Chapter Detected
                //{
                //	loadingCreating.Tag = "حذف " + preChapterDeleted.PageTitle;

                //	Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("حذف " + preChapterDeleted.PageTitle);
                //	await System.Threading.Tasks.Task.Run(() => deletePage(doc , preChapterDeleted.PageID , preChapterDeleted.Order));
                //	Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                //	onChanged = true;
                //}

                //foreach(AddRemovePageRelationModel cpr in chapterPageRelations)
                //{
                //	if(cpr.SubTemplateType == SubTemplateTypes.ChapterNotMade)//New Chapter Detected
                //	{
                //		loadingCreating.Tag = "ایجاد " + cpr.PageTitle;

                //		Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("ایجاد " + cpr.PageTitle);
                //		await System.Threading.Tasks.Task.Run(() => insertPage(doc , university , documentType , cpr.PageID , cpr.Order , cpr.SubTemplateID));
                //		Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                //		onChanged = true;
                //	}
                //}
                #endregion
                if (onChanged)
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("پیکربندی نهایی عملیات حذف و اضافه ");

                    Dictionary<string, string> listOfContents = getValueList(doc, university, templateType, documentType);
                    string[] contentControlNames = Enum.GetNames(typeof(ContentControlNames));
                    foreach (string contentControlName in contentControlNames)
                    {
                        bool mustBeEmpty = false;

                        DedicatedFunctions.getLockStatusContentControl(contentControlName, out bool lockContentControl, out bool lockContent);
                        string content;
                        try
                        {
                            content = listOfContents[contentControlName];
                        }
                        catch (Exception)
                        {
                            ContentControls ccs = doc.SelectContentControlsByTag(contentControlName);
                            if (ccs != null)
                            {
                                foreach (Microsoft.Office.Interop.Word.ContentControl cc in ccs)
                                {
                                    cc.LockContents = lockContent;
                                    cc.LockContentControl = lockContentControl;
                                }
                            }
                            continue;
                        }
                        string title = DedicatedFunctions.getContentControlTitle(contentControlName);

                        // ContentControlNames._field_Author_Title_En
                        // ContentControlNames._field_Author_Title_Fa
                        // ContentControlNames._field_Supervisor_Title_Fa
                        // ContentControlNames._field_Supervisor_Title_En
                        // ContentControlNames._field_Title_Title_En
                        // ContentControlNames._field_Title_Title_Fa

                        if (contentControlName == ContentControlNames._field_Advisor_Title_Fa.ToString() ||
                            contentControlName == ContentControlNames._field_Advisor_Title_En.ToString() ||
                            contentControlName == ContentControlNames._field_AreaOfStudy_Title_Fa.ToString() ||
                            contentControlName == ContentControlNames._field_AreaOfStudy_Title_En.ToString() ||
                            contentControlName == ContentControlNames._field_Advisor_Fa.ToString() ||
                            contentControlName == ContentControlNames._field_Advisor_En.ToString() ||
                            contentControlName == ContentControlNames._field_AreaOfStudy_Fa.ToString() ||
                            contentControlName == ContentControlNames._field_AreaOfStudy_En.ToString()
                            )
                        {
                            if (string.IsNullOrEmpty(content))
                            {
                                mustBeEmpty = true;
                            }
                        }

                        try
                        {
                            ContentControls ccs = doc.SelectContentControlsByTag(contentControlName);
                            if (ccs != null)
                            {
                                foreach (Microsoft.Office.Interop.Word.ContentControl cc in ccs)
                                {
                                    if (cc.Tag != ContentControlNames._field_InTheNameOfAllah.ToString())
                                    {
                                        DedicatedFunctions.ProtectImportantText(cc, content, title, title, lockContentControl, lockContent, false, mustBeEmpty);

                                        if (university == Universities.YazdUniversity && mustBeEmpty)
                                        {
                                            CreateDocumentSlide6.customActionInUniversity(doc, contentControlName);
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {
                            DedicatedFunctions.ShowErrorMessage("خطایی غیر منتظره رخ داد", (int)ErrorCodes.ContentControlProtect, mobile: StringConstant.SupportMobile, sendErrorToServer: true);
                            CloseFormRequest?.Invoke();
                            DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges, false);
                        }
                    }

                    try
                    {
                        #region Update Document via Template
                        string shivanegarTemplatesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates", "ShivaNegarTemplates");
                        Directory.CreateDirectory(shivanegarTemplatesPath);
                        string templateName = TemplateAccess.getTemplateFileName(university, templateType);

                        //unload AttachedTemplate for replace Template
                        List<Document> doucmentsAttachedToTemplate = new List<Document>();
                        foreach (Document docAttached in Globals.ThisAddIn.Application.Documents)
                        {
                            if (docAttached == doc)
                                continue;

                            Microsoft.Office.Interop.Word.Template attachedTemplate = docAttached.get_AttachedTemplate() as Microsoft.Office.Interop.Word.Template;
                            if (attachedTemplate != null && attachedTemplate.FullName.Equals(shivanegarTemplatesPath.TrimEnd('\\') + "\\" + templateName, StringComparison.OrdinalIgnoreCase))
                            {
                                doucmentsAttachedToTemplate.Add(docAttached);
                                docAttached.set_AttachedTemplate("");
                            }
                        }
                        doc.set_AttachedTemplate("");

                        //foreach(Microsoft.Office.Interop.Word.Template template in Globals.ThisAddIn.Application.Templates)
                        //{
                        //	if(template.FullName.Equals(path , StringComparison.OrdinalIgnoreCase))
                        //	{
                        //		// بستن فایل Template
                        //		template.Unload();
                        //		break;
                        //	}
                        //}

                        //copy Template File for UpdateDocument based on Template
                        Stream stream = TemplateAccess.getTemplateFileStream(university, templateType);
                        string templatePath = DedicatedFunctions.copyFileToFolder(stream, templateName, shivanegarTemplatesPath);
                        stream.Dispose();

                        //save last edit Template to document
                        doc.UpdateStylesOnOpen = true;
                        doc.set_AttachedTemplate(templatePath);
                        doc.UpdateStylesOnOpen = false;
                        //load again AttachedTemplate
                        foreach (Document docAttached in doucmentsAttachedToTemplate)
                        {
                            docAttached.set_AttachedTemplate(templatePath);
                        }

                        #endregion
                    }
                    catch (Exception e)
                    {
                        DedicatedFunctions.ShowErrorMessage("خطایی در اعمال Style و Page Layout با استفاده از Template سند مورد نظر رخ داد\nبا اینحال برنامه به کارخود ادامه میدهد\n" + e.Message, email: StringConstant.SupportEmail);
                    }
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

                    loadingCreating.Tag = "عملیات با موفقیت انجام شد";
                    await System.Threading.Tasks.Task.Run(() => Thread.Sleep(2000));
                }
                else
                {
                    loadingCreating.Tag = "تغییری توسط شما اعمال نشد";
                    await System.Threading.Tasks.Task.Run(() => Thread.Sleep(2000));
                }

                Globals.ThisAddIn.Application.ScreenUpdating = true;
                loadingCreating.IsEnabled = false;
                isUpdatingDocument = false;
                doc.UndoClear();
                DedicatedFunctions.saveDocument(doc);
                CloseFormRequest?.Invoke();
            }
            catch (Exception e)
            {
                loadingCreating.IsEnabled = false;
                isUpdatingDocument = false;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                if (Globals.ThisAddIn.Application.UndoRecord.IsRecordingCustomRecord)
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای در ایجاد و حذف به وجود آمده" + "\nپیغام خطا:\n" + e.Message);
                CloseFormRequest?.Invoke();
                DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges, false);
            }
        }


        private void BtnAddChapter_Click(object sender, RoutedEventArgs e)
        {
            btnRemoveChapter.IsEnabled = true;
            if (lastChapterIndex >= 8)
            {
                btnAddChapter.IsEnabled = false;
            }

            AddRemovePageRelationModel specifiedRelation = chapterPageRelations[lastChapterIndex + 1];
            initializeChapterControl(specifiedRelation);
            lastChapterIndex++;
        }
        private void BtnRemoveChapter_Click(object sender, RoutedEventArgs e)
        {
            btnAddChapter.IsEnabled = true;
            if (lastChapterIndex <= 1)
            {
                btnRemoveChapter.IsEnabled = false;
            }

            AddRemovePageRelationModel specifiedRelation = chapterPageRelations[lastChapterIndex];
            Card card = (Card)specifiedRelation.Control;
            stackPanelChapters.Children.Remove(card);
            lastChapterIndex--;
        }


        private void BtnMinimize_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            MinimizeStateFormRequest?.Invoke();
        }
        private void BtnMaximize_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            MaximizeStateFormRequest?.Invoke();
        }
        private void BtnCloseApp_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (!isUpdatingDocument)
                CloseFormRequest?.Invoke();
        }
        #endregion


        #region initialization
        private void initialization(Document doc, Universities university, TemplateTypes templateType, DocumentTypes documentType)
        {
            foreach (AddRemovePageRelationModel model in TemplateAccess.getAddRemovePageRelationModelList(doc, university, templateType, documentType))
            {
                if (model.SubTemplateType == SubTemplateTypes.Required || model.SubTemplateType == SubTemplateTypes.Optional)
                {
                    pageRelations.Add(model);
                    initializeControl(model);
                }
                else if (model.SubTemplateType == SubTemplateTypes.Chapter || model.SubTemplateType == SubTemplateTypes.ChapterNotMade)
                {
                    lastChapterIndex = DedicatedFunctions.getCurrentChaptersCount(doc) - 1;
                    lastExistingChapterIndex = lastChapterIndex;
                    chapterPageRelations.Add(model);

                    if (model.SubTemplateType == SubTemplateTypes.Chapter)
                        initializeChapterControl(model);
                }
            }

            if (lastChapterIndex < 1)
            {
                btnRemoveChapter.IsEnabled = false;
            }
            else if (lastChapterIndex >= 9)
            {
                btnAddChapter.IsEnabled = false;
            }
        }

        private void initializeControl(AddRemovePageRelationModel pageRelation)
        {
            //ToggleButton
            ToggleButton toggleButton = new ToggleButton
            {
                Margin = new Thickness(0, 0, 10, 0),
                HorizontalAlignment = HorizontalAlignment.Left,
                FlowDirection = FlowDirection.LeftToRight,
                IsEnabled = (pageRelation.SubTemplateType == SubTemplateTypes.Required) ? false : true,
                IsChecked = (bool)pageRelation.IsPageExist,
                UseLayoutRounding = false
            };
            pageRelation.Control = toggleButton;

            //TextBlock
            TextBlock textBlock = new TextBlock
            {
                Text = pageRelation.PageTitle,
                TextWrapping = TextWrapping.Wrap
            };

            //Grid
            var cd1 = new ColumnDefinition();
            cd1.Width = new GridLength(1, GridUnitType.Auto);
            var cd2 = new ColumnDefinition();
            cd2.Width = new GridLength(1, GridUnitType.Star);
            Grid grid = new Grid
            {
                Margin = new Thickness(0, 5, 0, 5),
                UseLayoutRounding = false
            };
            grid.ColumnDefinitions.Add(cd1);
            grid.ColumnDefinitions.Add(cd2);
            grid.Children.Add(toggleButton);
            grid.Children.Add(textBlock);
            Grid.SetColumn(toggleButton, 0);
            Grid.SetColumn(textBlock, 1);

            if (pageRelation.SubTemplateType == SubTemplateTypes.Required)
            {
                grid.ToolTip = "امکان حدف و اضافه این مورد وجود ندارد، چون صفحه الزامی میباشد";
                stackPanelEssentials.Children.Add(grid);
            }
            else if (pageRelation.SubTemplateType == SubTemplateTypes.Optional)
                stackPanelOptionals.Children.Add(grid);
        }

        private void initializeChapterControl(AddRemovePageRelationModel pageRelation)
        {
            string text = pageRelation.PageTitle;

            if (pageRelation.SubTemplateType == SubTemplateTypes.Chapter)
                text += " (ساخته شده)";
            else if (pageRelation.SubTemplateType == SubTemplateTypes.ChapterNotMade)
                text += " (جدید)";

            //TextBlock
            TextBlock textBlock = new TextBlock
            {
                Text = text,
                TextWrapping = TextWrapping.Wrap,
            };
            Card card = new Card
            {
                Margin = new Thickness(0, 0, 0, 10),
                Padding = new Thickness(5),
                UniformCornerRadius = 5
            };
            card.Content = textBlock;
            //if(pageRelation.SubTemplateType == SubTemplateTypes.Chapters)
            //	textBlock.IsEnabled = false;
            pageRelation.Control = card;

            stackPanelChapters.Children.Add(card);
        }
        #endregion
    }

}
