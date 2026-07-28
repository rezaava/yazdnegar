using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Threading;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View;
using ShivaNegar.Interfaces;
using ShivaNegar.Models;
using ShivaNegar.Templates;
using static ShivaNegar.DedicatedFunctions;
using static ShivaNegar.Models.TemplateRelationshipModel;

namespace ShivaNegar.Forms.ShivaNegarManager.CreateDocument
{
    public partial class CreateDocumentSlide6 : UserControl, IChangeTransitionDocumentManager, ICloseForm
    {
        public Action TransitionDocumentManagerRequest { get; set; }
        public Action CloseForm { get; set; }

        Document specifiedDocument;
        Window specifiedWindow;

        Document previousDocument;

        #region Progress
        System.Windows.Media.BrushConverter converter = new System.Windows.Media.BrushConverter();
        System.Windows.Media.Brush successfullBackgroundProgress;
        System.Windows.Media.Brush successfullForegroundProgress;
        System.Windows.Media.Brush proccesingBackgroundProgress;
        System.Windows.Media.Brush proccesingForegroundProgress;
        System.Windows.Media.Brush failedBackgroundProgress;
        System.Windows.Media.Brush failedForegroundProgress;
        string testStatus = "در حال آماده سازی اولیه";
        #endregion

        #region Variables
        internal string DocumentName { get; private set; }


        //Slide1
        private string nameOfAllah;
        private int nameOfAllahFontType;

        //Slide2
        private DocumentTypes? documentType;
        private TemplateTypes? templateType;
        string documentTypePersian;
        string documentTypeEnglish;
        string templateTypePersian;
        string templateTypeEnglish;

        //Slide3
        private string universityFa;
        private string universityEn;
        private string branchFa;
        private string branchEn;
        private Universities university;

        private string departmentFa;
        private string departmentEn;

        private string groupFa;
        private string groupEn;

        private string fieldOfStudyFa;
        private string fieldOfStudyEn;

        private string areaOfStudyFa;
        private string areaOfStudyEn;

        private string academicDegreeFa;
        private string academicDegreeEn;


        //Slide4
        private string nameOfCourseFa;

        private string titleFa;
        private string titleEn;

        private string authorFa;
        private string authorEn;

        private string advisorFa;
        private string advisorEn;

        private string supervisorFa;
        private string supervisorEn;

        private string defenseDateFa;
        private string defenseDateEn;

        #endregion

        #region TransitionSlides
        int lastSlideIndex = 0;
        bool allowClick = true;
        #endregion

        #region Timers
        System.Timers.Timer timerTransitionClick;
        System.Timers.Timer timerChangeSlide;
        System.Timers.Timer timerStart;

        System.Timers.Timer timerProgressStart;
        #endregion

        public System.Threading.Thread threadCreateDocument;
        bool isCompleted = false;
        bool closeFormClicked = false;
        public CreateDocumentSlide6()
        {
            InitializeComponent();

            #region Progress
            successfullBackgroundProgress = System.Windows.Media.Brushes.Green;
            successfullForegroundProgress = System.Windows.Media.Brushes.DarkGreen;
            //successfullBackgroundProgress = (System.Windows.Media.Brush)converter.ConvertFromString("#FF00BB40");
            //successfullForegroundProgress = (System.Windows.Media.Brush)converter.ConvertFromString("#FF009900");
            //System.Windows.Media.Brushes.DarkGreen

            failedBackgroundProgress = System.Windows.Media.Brushes.Red;
            failedForegroundProgress = System.Windows.Media.Brushes.DarkRed;

            proccesingBackgroundProgress = (Resources["LightBackgroundColor"] as System.Windows.Media.Brush);
            proccesingForegroundProgress = (System.Windows.Media.Brush)converter.ConvertFromString("#FF007AC1");

            ButtonProgressAssist.SetMinimum(btnProgress, 0);
            ButtonProgressAssist.SetMaximum(btnProgress, 100);
            ButtonProgressAssist.SetIsIndeterminate(btnProgress, false);
            #endregion

            #region Timers
            timerChangeSlide = new System.Timers.Timer
            {
                Interval = 5000,
                Enabled = false,
            };
            timerChangeSlide.Elapsed += TimerChangeSlide_Elapsed;

            timerTransitionClick = new System.Timers.Timer
            {
                Interval = 500,
                Enabled = false,
            };
            timerTransitionClick.Elapsed += TimerClick_Elapsed;


            timerStart = new System.Timers.Timer
            {
                Interval = 3000,
                Enabled = false,
            };
            timerStart.Elapsed += TimerStart_Elapsed;


            timerProgressStart = new System.Timers.Timer
            {
                Interval = 1000,
                Enabled = false,
            };
            timerProgressStart.Elapsed += TimerProgressStart_Elapsed;
            #endregion

            btnForward.Click += BtnForward_Click;
            btnBackward.Click += BtnBackward_Click;
            lastSlideIndex = transitionSildes.Items.Count - 1;

        }

        private void TimerProgressStart_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Dispatcher?.Invoke(() =>
            {
                ButtonProgressAssist.SetIsIndeterminate(btnProgress, true);
                timerProgressStart.Stop();
            });
        }

        #region Events

        #region Timers

        private void TimerStart_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Dispatcher?.Invoke(() =>
            {
                ButtonProgressAssist.SetIsIndeterminate(btnProgress, false);
                timerStart.Stop();
            });
            createNewDocument();
        }

        private void TimerChangeSlide_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Dispatcher?.Invoke(() =>
            {
                if (transitionSildes.SelectedIndex == lastSlideIndex)
                {
                    transitionSildes.SelectedIndex = 0;
                }
                else
                    transitionSildes.SelectedIndex++;
            });
        }

        private void TimerClick_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Dispatcher?.Invoke(() =>
            {
                allowClick = true;
                timerTransitionClick.Stop();
            });
        }
        #endregion

        #region Buttons
        private void BtnBackward_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (allowClick)
            {
                allowClick = false;
                timerTransitionClick.Start();
                if (transitionSildes.SelectedIndex == 0)
                {
                    transitionSildes.SelectedIndex = lastSlideIndex;
                }
                else
                    transitionSildes.SelectedIndex--;

                timerChangeSlide.Stop();
                timerChangeSlide.Start();
            }
        }

        private void BtnForward_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (allowClick)
            {
                allowClick = false;
                timerTransitionClick.Start();

                if (transitionSildes.SelectedIndex == lastSlideIndex)
                {
                    transitionSildes.SelectedIndex = 0;
                }
                else
                    transitionSildes.SelectedIndex++;

                timerChangeSlide.Stop();
                timerChangeSlide.Start();
            }
        }


        #endregion
        #endregion
        #region Functions
        public void close()
        {
            if (!isCompleted)
            {
                closeFormClicked = true;
                if (threadCreateDocument != null)
                    threadCreateDocument.Abort();

                Dispatcher.Invoke(() =>
                {
                    timerStart.Stop();

                    btnProgress.Content = "ساخت سند لغو شد";
                    ButtonProgressAssist.SetValue(btnProgress, 0);
                    ButtonProgressAssist.SetIsIndeterminate(btnProgress, false);
                    ButtonProgressAssist.SetIsIndeterminate(btnProgress, true);
                    ButtonProgressAssist.SetIndicatorForeground(btnProgress, failedForegroundProgress);
                    ButtonProgressAssist.SetIndicatorBackground(btnProgress, failedBackgroundProgress);
                    btnProgress.Background = failedBackgroundProgress;
                });

                System.Timers.Timer timerToExit = new System.Timers.Timer
                {
                    Interval = 3000,
                    Enabled = false,
                };
                timerToExit.Start();
                timerToExit.Elapsed += (a, w) =>
                {
                    timerToExit.Stop();
                    timerToExit.Enabled = false;
                    timerToExit.Dispose();
                    timerChangeSlide.Stop();
                    timerChangeSlide.Dispose();
                    cancelCreateDocument();
                };
            }
        }

        public static void SetPercent(Button button, double percentage, TimeSpan duration)
        {
            DoubleAnimation animation = new DoubleAnimation(percentage, duration);

            button.BeginAnimation(ButtonProgressAssist.ValueProperty, animation);
        }

        public void resetControls()
        {
            #region Variables
            isCompleted = false;
            specifiedDocument = null;
            specifiedWindow = null;
            previousDocument = null;
            lastSlideIndex = 0;
            allowClick = true;

            DocumentName = null;

            nameOfAllah = null;
            nameOfAllahFontType = 0;
            documentType = null;
            templateType = null;
            documentTypePersian = null;
            documentTypeEnglish = null;
            templateTypePersian = null;
            templateTypeEnglish = null;

            universityFa = null;
            universityEn = null;
            branchFa = null;
            branchEn = null;
            departmentFa = null;
            departmentEn = null;
            groupFa = null;
            groupEn = null;
            fieldOfStudyFa = null;
            fieldOfStudyEn = null;
            areaOfStudyFa = null;
            areaOfStudyEn = null;
            academicDegreeFa = null;
            academicDegreeEn = null;

            nameOfCourseFa = null;
            titleFa = null;
            titleEn = null;
            authorFa = null;
            authorEn = null;
            supervisorFa = null;
            supervisorEn = null;
            advisorFa = null;
            advisorEn = null;
            defenseDateFa = null;
            defenseDateEn = null;
            #endregion

            Dispatcher.Invoke(() =>
            {
                #region Controls
                btnProgress.Content = testStatus;

                ButtonProgressAssist.SetValue(btnProgress, -1);
                ButtonProgressAssist.SetIsIndeterminate(btnProgress, true);
                ButtonProgressAssist.SetIndicatorForeground(btnProgress, proccesingForegroundProgress);
                ButtonProgressAssist.SetIndicatorBackground(btnProgress, proccesingBackgroundProgress);
                btnProgress.Background = proccesingBackgroundProgress;

                timerTransitionClick.Stop();
                timerChangeSlide.Stop();
                timerStart.Stop();
                timerProgressStart.Stop();
                #endregion
            });
        }

        public void initializeVariables(CreateDocumentSlide1 slide1, CreateDocumentSlide2 slide2, CreateDocumentSlide3 slide3, CreateDocumentSlide4 slide4)
        {
            #region start Timers
            timerChangeSlide.Start();
            timerStart.Start();
            timerProgressStart.Start();
            #endregion

            university = slide3.University;
            threadCreateDocument = new Thread(() => create());

            DocumentName = slide2.DocumentName;

            //Slide1
            nameOfAllah = slide1.NameOfAllah;
            nameOfAllahFontType = slide1.NameOfAllahFontType;
            //Slide2
            documentType = slide2.DocumentType;
            templateType = slide2.TemplateType;
            documentTypePersian = DedicatedFunctions.getDocumentTypePersianName((DocumentTypes)documentType);
            documentTypeEnglish = DedicatedFunctions.getDocumentTypeEnglishName((DocumentTypes)documentType);
            templateTypePersian = DedicatedFunctions.getTemplateTypePersianName((TemplateTypes)templateType);
            templateTypeEnglish = DedicatedFunctions.getTemplateTypeEnglishName((TemplateTypes)templateType);

            //Slide3
            universityFa = slide3.UniversityFa;
            universityEn = slide3.UniversityEn;
            branchFa = slide3.BranchFa;
            branchEn = slide3.BranchEn;
            departmentFa = slide3.DepartmentFa;
            departmentEn = slide3.DepartmentEn;
            groupFa = slide3.GroupFa;
            groupEn = slide3.GroupEn;
            fieldOfStudyFa = slide3.FieldOfStudyFa;
            fieldOfStudyEn = slide3.FieldOfStudyEn;
            areaOfStudyFa = slide3.AreaOfStudyFa;
            areaOfStudyEn = slide3.AreaOfStudyEn;
            academicDegreeFa = slide3.AcademicDegreeFa;
            academicDegreeEn = slide3.AcademicDegreeEn;

            //Slide4
            nameOfCourseFa = slide4.NameOFCourseFa;
            titleFa = slide4.TitleFa;
            titleEn = slide4.TitleEn;
            authorFa = slide4.AuthorFa;
            authorEn = slide4.AuthorEn;
            supervisorFa = slide4.SupervisorFa;
            supervisorEn = slide4.SupervisorEn;
            advisorFa = slide4.AdvisorFa;
            advisorEn = slide4.AdvisorEn;
            defenseDateFa = slide4.DefenseDateFa;
            defenseDateEn = slide4.DefenseDateEn;

            //Slide5
        }
        #endregion

        #region The main part of the work

        private void createNewDocument()
        {
            System.Threading.ThreadState threadNotStartAndAbort = System.Threading.ThreadState.Unstarted | System.Threading.ThreadState.AbortRequested;
            if (threadCreateDocument.ThreadState != System.Threading.ThreadState.Aborted && threadCreateDocument.ThreadState != threadNotStartAndAbort)
            {
                Globals.ThisAddIn.DisableEvents = true;

                //btnProgress.Content = "آماده سازی اولیه";TODO:Uncomment This
                //Globals.ThisAddIn.Application.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                //Globals.ThisDocument.RemovePersonalInformation = true;
                //_ = System.Threading.Tasks.Task.Run(() =>
                //{
                //    DedicatedFunctions.closeDialog("Microsoft Word", "OK", "that can't be removed by the Document Inspector", 500, 10);
                //});//close Document Inspector Dialog
                //Globals.ThisDocument.Save();
                //Globals.ThisDocument.RemovePersonalInformation = false;
                //Globals.ThisAddIn.Application.DisplayAlerts = WdAlertLevel.wdAlertsAll;

                if (Globals.ThisAddIn.Application.Documents.Count != 0)
                {
                    previousDocument = Globals.ThisAddIn.Application.ActiveDocument;
                    if (!File.Exists(previousDocument.FullName) && previousDocument.Characters.Count < 2)
                    {
                        previousDocument.ActiveWindow.Visible = false;
                    }
                }
                foreach (Microsoft.Office.Interop.Word.Window window in Globals.ThisAddIn.Application.Windows)
                {
                    window.Visible = false;
                }

                //copy Template File for createDocument based on Template
                string templateName = TemplateAccess.getTemplateFileName(university, (TemplateTypes)templateType);

                using (Stream stream = TemplateAccess.getTemplateFileStream(university, (TemplateTypes)templateType))
                {
                    string shivanegarTemplatesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates", "ShivaNegarTemplates");
                    Directory.CreateDirectory(shivanegarTemplatesPath);
                    string templatePath = DedicatedFunctions.copyFileToFolder(stream, templateName, shivanegarTemplatesPath);

                    Globals.ThisAddIn.DisableEvents = true;
                    specifiedDocument = Globals.ThisAddIn.Application.Documents.Add(templatePath);
                    Globals.ThisAddIn.DisableEvents = false;
                    specifiedWindow = specifiedDocument.ActiveWindow;
                    specifiedWindow.Visible = false;
                }


                threadCreateDocument.Start();
            }
        }

        private Dictionary<string, string> getValueList(Universities university, TemplateTypes templateType, DocumentTypes documentType)
        {
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

                    { ContentControlNames._field_Advisor_En.ToString(),advisorEn},
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

                    { ContentControlNames._field_Advisor_Title_En.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_Advisor_Title_En)},
                    { ContentControlNames._field_Advisor_Title_Fa.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_Advisor_Title_Fa)},
                    { ContentControlNames._field_AreaOfStudy_Title_Fa.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_AreaOfStudy_Title_Fa)},
                    { ContentControlNames._field_AreaOfStudy_Title_En.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_AreaOfStudy_Title_En)},

					//{ ContentControlNames._field_Author_Title_En.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_Author_Title_En)},
					//{ ContentControlNames._field_Author_Title_Fa.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_Author_Title_Fa)},
					//{ ContentControlNames._field_Supervisor_Title_Fa.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_Supervisor_Title_Fa)},
					//{ ContentControlNames._field_Supervisor_Title_En.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_Supervisor_Title_En)},
					//{ ContentControlNames._field_Title_Title_En.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_Title_Title_En)},
					//{ ContentControlNames._field_Title_Title_Fa.ToString(),TemplateAccess.getCustomTitle(university,templateType,ContentControlNames._field_Title_Title_Fa)},

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

        private void create()
        {
            List<TemplateRelationshipModel> subTemplateModels = new List<TemplateRelationshipModel>();
            foreach (var model in TemplateAccess.getTemplateRelationshipModelList(university, (TemplateTypes)templateType, (DocumentTypes)documentType, false))
            {
                if (model.SubTemplateType == SubTemplateTypes.Required || model.SubTemplateType == SubTemplateTypes.Chapter)
                    subTemplateModels.Add(model);
            }
            #region Progress
            double progressStep = 100d / (subTemplateModels.Count + 2);
            int progressAnimationSpeed = 200;

            Dispatcher.Invoke(() =>
            {
                btnProgress.Content = "انجام تنظیمات اولیه";
                SetPercent(btnProgress, ButtonProgressAssist.GetValue(btnProgress) + progressStep + 1, TimeSpan.FromMilliseconds(progressAnimationSpeed * 2));
            });
            Thread.Sleep(progressAnimationSpeed * 2);
            #endregion

            //create Section Break Style
            createSectionBreakStyle(specifiedDocument);

            #region set Variables
            if (documentType == DocumentTypes.SchoolResearch)
                DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_NameOfCourse_Fa.ToString(), nameOfCourseFa);

            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Advisor_Fa.ToString(), advisorFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_University_En.ToString(), universityEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Branch_En.ToString(), branchEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Group_En.ToString(), groupEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Department_En.ToString(), departmentEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_AcademicDegree_En.ToString(), academicDegreeEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_FieldOfStudy_En.ToString(), fieldOfStudyEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_AreaOfStudy_En.ToString(), areaOfStudyEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Title_En.ToString(), titleEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Supervisor_En.ToString(), supervisorEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Advisor_En.ToString(), advisorEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Author_En.ToString(), authorEn);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_DefenseDate_En.ToString(), defenseDateEn);

            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_University_Fa.ToString(), universityFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Branch_Fa.ToString(), branchFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Department_Fa.ToString(), departmentFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Group_Fa.ToString(), groupFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString(), academicDegreeFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_FieldOfStudy_Fa.ToString(), fieldOfStudyFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_AreaOfStudy_Fa.ToString(), areaOfStudyFa);

            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Title_Fa.ToString(), titleFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Supervisor_Fa.ToString(), supervisorFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_Author_Fa.ToString(), authorFa);
            DedicatedFunctions.addVariable(specifiedDocument, VariableFieldIDs._variable_field_DefenseDate_Fa.ToString(), defenseDateFa);


            DedicatedFunctions.addVariable(specifiedDocument, VariableTypeIDs._variable_type_Document.ToString(), ((int)documentType).ToString());
            #endregion

            #region create Pages of Document

            Stopwatch sw = new Stopwatch();
            Stopwatch sw2 = new Stopwatch();
            sw.Start();
            int counterForDetectLast = 0;

            foreach (TemplateRelationshipModel relations in subTemplateModels)
            {
                sw2.Restart();

                #region Progress
                Dispatcher.Invoke(() =>
                {
                    btnProgress.Content = "ایجاد " + relations.PageTitle;
                    SetPercent(btnProgress, ButtonProgressAssist.GetValue(btnProgress) + progressStep, TimeSpan.FromMilliseconds(progressAnimationSpeed));
                });
                Thread.Sleep(progressAnimationSpeed);
                #endregion

                //copy and get specified file template templatePath
                Stream stream = DedicatedFunctions.getStream(relations.ResourcePath + relations.FileName);
                string templatePath = DedicatedFunctions.copyFileToTempFolder(stream, relations.FileName);

                try
                {
                    insertDocumentFile(specifiedDocument, templatePath, DedicatedFunctions.GetCustomProperties(templatePath));
                }
                catch (Exception e)
                {
                    if (!closeFormClicked)
                    {
                        DedicatedFunctions.ShowErrorMessage("مشکلی در ساخت سند به وجود آمد\nساخت سند شما لغو میشود و میتوانید مجدد برای ساخت سند اقدام کنید \nخطا:" + e.Message, email: StringConstant.SupportEmail);
                        cancelCreateDocument();
                    }
                    return;
                }

                //specifiedDocument.ActiveWindow.Selection.Collapse();
                if (counterForDetectLast != subTemplateModels.Count - 1)
                {
                    Globals.ThisAddIn.DisableSelectionChangedEvent = true;
                    //insert SectionBreak
                    //specifiedDocument.ActiveWindow.Selection.EndKey();
                    specifiedDocument.ActiveWindow.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                    DedicatedFunctions.insertSectionBreak(specifiedDocument, DedicatedFunctions.GetCustomProperties(templatePath));

                    //select line and reduce font size for protecting
                    Range previousRange = specifiedDocument.ActiveWindow.Selection.Range;
                    specifiedDocument.ActiveWindow.Selection.MoveLeft(WdUnits.wdCharacter, 1);
                    string contentControlID = DedicatedFunctions.protectSectionBreak(specifiedDocument, relations.PageID.ToString());
                    //restore Selection to previous Range
                    previousRange.Select();

                    DedicatedFunctions.addVariable(specifiedDocument, relations.PageID.ToString(), contentControlID);
                    Globals.ThisAddIn.DisableSelectionChangedEvent = false;
                }
                else
                {
                    //lastPage
                    DedicatedFunctions.addVariable(specifiedDocument, relations.PageID.ToString(), "LastPage");
                }

                //remove specified file template
                if (!System.IO.Path.GetFullPath(new FileInfo(templatePath).Directory.FullName).TrimEnd('\\').Contains(System.IO.Path.GetFullPath(Properties.Settings.Default.WorkSpaceDirectory + StringConstant.DocumentsTemplateFolder).TrimEnd('\\')))
                {
                    DedicatedFunctions.removeFileFromSystem(templatePath);
                }

                counterForDetectLast++;
                Debug.WriteLine("\t" + relations.PageTitle + " , Ellapsed Time> " + sw2.ElapsedMilliseconds + "ms");
            }
            sw.Stop();
            Debug.WriteLine("Finished, Ellapsed Time> " + sw.ElapsedMilliseconds + "ms");
            #endregion

            #region Progress
            Dispatcher.Invoke(() =>
            {
                SetPercent(btnProgress, 100, TimeSpan.FromMilliseconds(progressAnimationSpeed));
                btnProgress.Content = "پیکربندی نهایی";
            });
            #endregion

            #region final Configuration
            sw.Restart();
            foreach (Section section in specifiedDocument.Sections)
            {
                DedicatedFunctions.resetHeaderFooter(section, true);
            }

            Dictionary<string, string> listOfContents = getValueList(university, (TemplateTypes)templateType, (DocumentTypes)documentType);
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
                    // اگر ContentControl موجود در ContentControlNames در listOfContents وجود نداشت
                    ContentControls ccs2 = specifiedDocument.SelectContentControlsByTag(contentControlName);
                    if (ccs2 != null)
                    {
                        foreach (Microsoft.Office.Interop.Word.ContentControl cc in ccs2)
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

                if (contentControlName == ContentControlNames._field_Advisor_Title_Fa.ToString() || contentControlName == ContentControlNames._field_Advisor_Fa.ToString())
                {
                    if (string.IsNullOrEmpty(advisorFa.Trim()))
                    {
                        content = "";
                        mustBeEmpty = true;
                    }
                }
                else if (contentControlName == ContentControlNames._field_Advisor_Title_En.ToString() || contentControlName == ContentControlNames._field_Advisor_En.ToString())
                {
                    if (string.IsNullOrEmpty(advisorEn.Trim()))
                    {
                        content = "";
                        mustBeEmpty = true;
                    }
                }
                else if (contentControlName == ContentControlNames._field_AreaOfStudy_Title_Fa.ToString() || contentControlName == ContentControlNames._field_AreaOfStudy_Fa.ToString())
                {
                    if (string.IsNullOrEmpty(areaOfStudyFa.Trim()))
                    {
                        content = "";
                        mustBeEmpty = true;
                    }
                }
                else if (contentControlName == ContentControlNames._field_AreaOfStudy_Title_En.ToString() || contentControlName == ContentControlNames._field_AreaOfStudy_En.ToString())
                {
                    if (string.IsNullOrEmpty(areaOfStudyEn.Trim()))
                    {
                        content = "";
                        mustBeEmpty = true;
                    }
                }

                ContentControls ccs = specifiedDocument.SelectContentControlsByTag(contentControlName);
                if (ccs != null)
                {
                    foreach (Microsoft.Office.Interop.Word.ContentControl cc in ccs)
                    {
                        if (cc.Tag == ContentControlNames._field_InTheNameOfAllah.ToString())
                        {
                            string besmellahFont;
                            if (nameOfAllahFontType == 1)
                                besmellahFont = Constants.FontNames.fontBesmellah1;
                            else if (nameOfAllahFontType == 2)
                                besmellahFont = Constants.FontNames.fontBesmellah2;
                            else if (nameOfAllahFontType == 3)
                                besmellahFont = Constants.FontNames.fontBesmellah3;
                            else if (nameOfAllahFontType == 4)
                                besmellahFont = Constants.FontNames.fontBesmellah4;
                            else
                                throw new Exception("Besmellah font wrong");

                            cc.Range.Font.Name = besmellahFont;
                            cc.Range.Font.NameBi = besmellahFont;
                        }

                        DedicatedFunctions.ProtectImportantText(cc, content, title, title, lockContentControl, lockContent, false, mustBeEmpty);

                        if (university == Universities.YazdUniversity && mustBeEmpty)
                        {
                            customActionInUniversity(specifiedDocument, contentControlName);
                        }
                    }
                }
            }
            sw.Stop();
            Debug.WriteLine("final Configuration , Ellapsed Time> " + sw.ElapsedMilliseconds + "ms");
            #endregion

            #region Progress
            Dispatcher.Invoke(() =>
            {
                ButtonProgressAssist.SetIsIndeterminate(btnProgress, true);
                btnProgress.Content = "لطفا منتظر بمانید...";
            });
            Thread.Sleep(1500);
            #endregion

            #region Save As

            #region add Variable
            string version = BugReport.AssemblyVersion.Replace(".", "");

            DedicatedFunctions.addVariable(specifiedDocument, VariableServerIDs._variable_server_VersionNumber.ToString(), version);
            DedicatedFunctions.addVariable(specifiedDocument, VariableServerIDs._variable_server_UserToken.ToString(), Properties.Settings.Default.UserToken);
            DedicatedFunctions.addVariable(specifiedDocument, VariableVersionIDs._variable_version_AddIn.ToString(), Properties.Settings.Default.VersionAddin.ToString());
            DedicatedFunctions.addVariable(specifiedDocument, VariableVersionIDs._variable_version_Template.ToString(), Properties.Settings.Default.VersionTemplate.ToString());
            DedicatedFunctions.addVariable(specifiedDocument, VariableTypeIDs._variable_type_Template.ToString(), ((int)((TemplateTypes)templateType)).ToString());
            DedicatedFunctions.addVariable(specifiedDocument, VariableIdentifierIDs._variable_id_Template.ToString(), getTemplateID((TemplateTypes)templateType, university).ToString());

            DedicatedFunctions.addVariable(specifiedDocument, VariableIdentifierIDs._variable_id_AcademicDegree.ToString(), ((int)DedicatedFunctions.getAcademicDegreeID(academicDegreeFa)).ToString());
            DedicatedFunctions.addVariable(specifiedDocument, VariableIdentifierIDs._variable_id_GUID.ToString(), StringConstant.GUID);
            DedicatedFunctions.addVariable(specifiedDocument, VariableIdentifierIDs._variable_id_University.ToString(), ((int)university).ToString());
            DedicatedFunctions.addVariable(specifiedDocument, VariableIdentifierIDs._variable_id_Document.ToString(), specifiedDocument.DocID.ToString());
            DedicatedFunctions.addVariable(specifiedDocument, VariableIdentifierIDs._variable_id_Hardware.ToString(), DedicatedFunctions.getUUID());
            DedicatedFunctions.addVariable(specifiedDocument, VariableOptionIDs._variable_option_BibliographyStyle.ToString(), "APA");
            #endregion

            //#region Scroll
            ////first section index in chapter1
            //PageIDs previousPageID = getLastSectionPageID(specifiedDocument, university, (TemplateTypes)templateType, (DocumentTypes)documentType, PageIDs._page_Chapter1, WhichIndex.Previous);
            //int previousPageLastIndex = DedicatedFunctions.getPageIDIndex(specifiedDocument, previousPageID.ToString());
            //int pageChapter1 = (int)specifiedDocument.Sections[previousPageLastIndex].Range.Information[WdInformation.wdActiveEndPageNumber] + 3;
            //DedicatedFunctions.setORAddStaticVariableValue(specifiedDocument, VariableDocumentIDs._variable_document_PositionCurrentPage.ToString(), pageChapter1.ToString());
            //DedicatedFunctions.scrollToPage(specifiedWindow, specifiedDocument.ActiveWindow.Selection, pageChapter1);
            //#endregion
            #region Scroll - Go to First Page
            try
            {
                // رفتن به صفحه اول سند
                specifiedDocument.GoTo(WdGoToItem.wdGoToPage, WdGoToDirection.wdGoToAbsolute, 1).Select();
                specifiedDocument.ActiveWindow.Selection.HomeKey();
                specifiedWindow.ScrollIntoView(specifiedDocument.ActiveWindow.Selection.Range);

                // ذخیره شماره صفحه جاری
                DedicatedFunctions.setORAddStaticVariableValue(specifiedDocument, VariableDocumentIDs._variable_document_PositionCurrentPage.ToString(), "1");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error scrolling to first page: " + ex.Message);
            }
            #endregion


            string path = Properties.Settings.Default.WorkSpaceDirectory.TrimEnd('\\') + "\\" + DocumentName + ".docx";
            Directory.CreateDirectory(Properties.Settings.Default.WorkSpaceDirectory);

            int currentAddinVersion = int.Parse(Properties.Settings.Default.VersionAddin.ToString());
            if (currentAddinVersion == 1)
                specifiedDocument.Password = StringConstant.DocumentPassword;
            //if(currentAddinVersion == 2)
            //	specifiedDocument.Password = StringConstant.DocumentPassword2;

            specifiedDocument.UndoClear();
            #region Update Tables
            try
            {
                DedicatedFunctions.updateTables(specifiedDocument, specifiedDocument.ActiveWindow.Selection, AccessType.AccessGranted);
            }
            catch (Exception)
            {
            }
            #endregion

            DedicatedFunctions.saveAsDocument(specifiedDocument, path);
            #endregion

            #region Progress
            Dispatcher.Invoke(() =>
            {
                isCompleted = true;
                btnProgress.Content = "ذخیره اطلاعات در سرور";
            });
            Thread.Sleep(1500);
            #endregion

            saveToServer(specifiedDocument, Properties.Settings.Default.UserToken);
        }

        internal static void customActionInUniversity(Microsoft.Office.Interop.Word.Document doc, string ccName)
        {
            if (ccName == ContentControlNames._field_Advisor_En.ToString())
            {
                Range previousRange = doc.ActiveWindow.Selection.Range;
                Microsoft.Office.Interop.Word.ContentControl[] ccsAdvisorEn = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Advisor_Title_En.ToString());

                foreach (Microsoft.Office.Interop.Word.ContentControl contentControl in ccsAdvisorEn)
                {

                    bool previousLockState = contentControl.LockContents;

                    contentControl.LockContents = false;
                    contentControl.Range.Select();

                    doc.ActiveWindow.Selection.HomeKey();
                    doc.ActiveWindow.Selection.TypeBackspace();

                    contentControl.LockContents = previousLockState;

                    previousRange.Select();

                }
            }
        }
        private void saveToServer(Microsoft.Office.Interop.Word.Document doc, string token)
        {
            DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
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

            string urlParameters = "save?type=" + (int)documentType + "&name=" + DocumentName + "&config=" + jsonVariables.ToString();
            var formData = new MultipartFormDataContent();
            var fileStream = new FileStream(doc.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var fileContent = new StreamContent(fileStream);
            //fileContent.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/octet-templateXMLStream");
            //formData.Add(new ByteArrayContent(file , 0 , file.Length) , "documentFile" , fileName + ".docx");
            //byte[] fileData = File.ReadAllBytes(filePath);
            //formData.Add(new ByteArrayContent(fileData , 0 , fileData.Length) , "file" , Path.GetFileName(filePath));
            //formData.Add(fileContent , "file" , doc.Name);
            formData.Add(fileContent, "file", "documentfile.docx");

            DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            OnResult =>
            {
                try
                {
                    JsonDocument document = JsonDocument.Parse(OnResult);
                    JsonElement root = document.RootElement;
                    root.TryGetProperty("id", out JsonElement idElement);
                    int DocuementID = idElement.GetInt32();
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString(), DocuementID.ToString());

                    root.TryGetProperty("updated", out JsonElement updatedAtElement);
                    string updatedAt = updatedAtElement.GetString();
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedAt.ToString(), updatedAt);
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedFile.ToString(), updatedAt);
                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedConfig.ToString(), updatedAt);
                }
                catch (Exception)
                {
                }

                Dispatcher.Invoke(() =>
                {
                    ButtonProgressAssist.SetValue(btnProgress, 0);
                    ButtonProgressAssist.SetIsIndeterminate(btnProgress, false);
                    ButtonProgressAssist.SetIsIndeterminate(btnProgress, true);
                    ButtonProgressAssist.SetIndicatorForeground(btnProgress, successfullForegroundProgress);
                    ButtonProgressAssist.SetIndicatorBackground(btnProgress, successfullBackgroundProgress);
                    btnProgress.Background = successfullBackgroundProgress;

                    btnProgress.Content = "سند شما با موفقیت ایجاد شد";
                });
                Thread.Sleep(1500);

                onFinishTask();

                Dispatcher.Invoke(() =>
                {
                    timerChangeSlide.Stop();
                    timerChangeSlide.Dispose();
                    timerStart.Stop();
                    timerStart.Dispose();
                    timerProgressStart.Stop();
                    timerProgressStart.Dispose();
                    CloseForm?.Invoke();
                });
            },
            OnFailed =>
            {
                System.Windows.Forms.DialogResult dr = System.Windows.Forms.DialogResult.None;

                Dispatcher.Invoke(() =>
                {
                    lblServerDialogMessage.Text = ErrorMessages.ErrorServiceUnavailable + "\n" + OnFailed.ReasonPhrase;
                    dialogServerError.IsOpen = true;
                    btnServerTryAgainDialog.Click += (a, x) =>
                    {
                        dr = System.Windows.Forms.DialogResult.Retry;
                    };
                    btnServerCancelDialog.Click += (a, x) =>
                    {
                        dr = System.Windows.Forms.DialogResult.Cancel;
                    };
                });
                while (dr == System.Windows.Forms.DialogResult.None)
                {//TODO:better way

                }
                Dispatcher.Invoke(() =>
                { dialogServerError.IsOpen = false; });
                Thread.Sleep(1000);

                if (dr == System.Windows.Forms.DialogResult.Retry)
                {
                    saveToServer(doc, token);
                }
                else if (dr == System.Windows.Forms.DialogResult.Cancel)
                {
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        cancelCreateDocument();
                    });
                }
            }, formData);
        }
        private void cancelCreateDocument()
        {
            #region Progress

            Dispatcher.Invoke(() =>
            {
                timerChangeSlide.Stop();
                timerStart.Stop();
                timerProgressStart.Stop();

                btnProgress.Content = "ساخت سند لغو شد";
                ButtonProgressAssist.SetValue(btnProgress, 0);
                ButtonProgressAssist.SetIsIndeterminate(btnProgress, false);
                ButtonProgressAssist.SetIsIndeterminate(btnProgress, true);
                ButtonProgressAssist.SetIndicatorForeground(btnProgress, failedForegroundProgress);
                ButtonProgressAssist.SetIndicatorBackground(btnProgress, failedBackgroundProgress);
                btnProgress.Background = failedBackgroundProgress;
            });
            if (specifiedDocument != null)
            {
                try
                {
                    string path = specifiedDocument.FullName;
                    DedicatedFunctions.closeDocument(specifiedDocument, WdSaveOptions.wdDoNotSaveChanges, false);
                    if (File.Exists(path))
                    {
                        try
                        {
                            File.Delete(path);
                        }
                        catch (Exception)
                        {
                            DedicatedFunctions.ShowErrorMessage("خطا در حذف سند");
                        }
                    }
                }
                catch (Exception)
                {
                    DedicatedFunctions.ShowErrorMessage("خطا در بستن سند");
                }
            }

            Thread.Sleep(3000);
            Dispatcher.Invoke(() =>
            {
                CloseForm?.Invoke();
                //TransitionDocumentManagerRequest?.Invoke();
            });

            #region restate Documents visible And ScreenUpdating
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            foreach (Microsoft.Office.Interop.Word.Window window in Globals.ThisAddIn.Application.Windows)
            {
                window.Visible = true;
            }
            #endregion

            Globals.ThisAddIn.DisableEvents = false;
            #endregion
        }
        private void onFinishTask()
        {
            specifiedDocument.ContentControlOnExit += Globals.ThisAddIn.Doc_ContentControlOnExit;
            if (!Globals.ThisAddIn.accessedInStartup)
            {
                DedicatedFunctions.initialSettings();
                Ribbon.loadKeyboardShortcut();
                Globals.ThisAddIn.accessedInStartup = true;
            }

            #region restate Documents visible And ScreenUpdating
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            foreach (Microsoft.Office.Interop.Word.Window window in Globals.ThisAddIn.Application.Windows)
            {
                window.Visible = true;
            }
            try
            {
                if (previousDocument != null)
                {
                    if (!File.Exists(previousDocument.FullName) && previousDocument.Characters.Count < 2)
                    {
                        DedicatedFunctions.closeDocument(previousDocument, WdSaveOptions.wdDoNotSaveChanges);
                    }
                }
            }
            catch (Exception)
            {
            }
            #endregion


            DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);
            specifiedWindow.View.Zoom.Percentage = 100;
            specifiedWindow.Activate();
            specifiedWindow.Visible = true;
            Ribbon.InitializeRibbon(StringConstant.NameOfProject + "(" + DedicatedFunctions.getDocumentTypePersianName(specifiedDocument) + ")");
            Globals.ThisAddIn.DisableEvents = true;
            DedicatedFunctions.saveDocument(specifiedDocument);
            Globals.ThisAddIn.DisableEvents = false;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        #endregion

        internal static void insertDocumentFile(Microsoft.Office.Interop.Word.Document doc, string path, DocumentFormat.OpenXml.CustomProperties.Properties properties)
        {
            Selection selection = doc.ActiveWindow.Selection;

            DedicatedFunctions.setSectionStartFromBreakType(selection.Sections[1], DedicatedFunctions.getBreakTypeFromCustomDocumentProperties(properties));

            int preSectionCount = doc.Sections.Count;
            DedicatedFunctions.resetHeaderFooter(selection.Sections[1], false);
            selection.InsertFile(path);
            int differentSections = doc.Sections.Count - preSectionCount;

            DedicatedFunctions.setProperties(doc, selection.Sections[1], differentSections, DedicatedFunctions.GetCustomProperties(path), DedicatedFunctions.GetBodyPart(path));
        }
    }
}