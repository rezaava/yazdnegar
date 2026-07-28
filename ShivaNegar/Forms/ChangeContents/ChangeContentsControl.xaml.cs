using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Constants.ComboBoxData;
using ShivaNegar.Forms.ChangeContents.Models;
using ShivaNegar.Templates;

namespace ShivaNegar.Forms.ChangeContents
{
    /// <summary>
    /// Interaction logic for ChangeContentsControl.xaml
    /// </summary>
    public partial class ChangeContentsControl : UserControl, Interfaces.IStatusFormRequest
    {
        private List<ChangeContentsModel> contents;

        Document doc;
        public Action CloseFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }
        public Action NormalStateFormRequest { get; set; }

        #region Besmellah
        public int NameOfAllahFontType { get; private set; }

        string besmellahFontNameInUI1 = "B e s m e l l a h 1";
        string besmellahFontNameInUI2 = "B e s m e l l a h 2";
        string besmellahFontNameInUI3 = "B e s m e l l a h 3";
        //string besmellahFontNameInUI4 = "B e s m e l l a h 4";
        #endregion

        Universities university;

        #region constructor
        public ChangeContentsControl()
        {
            InitializeComponent();
        }
        public ChangeContentsControl(Document doc)
        {
            InitializeComponent();


            contents = new List<ChangeContentsModel>();

            this.doc = doc;

            btnCloseApp.Click += BtnCloseApp_Click;
            //btnMaximize.Click += BtnMaximize_Click;
            //btnMinimize.Click += BtnMinimize_Click;
            btnConfirmChanges.Click += BtnConfirmChanges_Click;

            txtBoxSupervisorEn.PreviewKeyDown += TextBoxMultiLine_PreviewKeyDown;
            txtBoxSupervisorFa.PreviewKeyDown += TextBoxMultiLine_PreviewKeyDown;
            txtBoxAdvisorEn.PreviewKeyDown += TextBoxMultiLine_PreviewKeyDown;
            txtBoxAdvisorFa.PreviewKeyDown += TextBoxMultiLine_PreviewKeyDown;

            expanderNameOfAllah.Expanded += Expander_Expanded;
            expanderEducationDetails.Expanded += Expander_Expanded;
            expanderDocumentDetails.Expanded += Expander_Expanded;
            expanderAbstractAndKeywords.Expanded += Expander_Expanded;

            loadingControl.IsEnabled = true;
            getDocumentDataFromServer(
                            DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UserToken.ToString()),
                            DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString()));

            university = DedicatedFunctions.getUniversity(doc);
        }

        private void getDocumentDataFromServer(string token, string documentID)
        {
            string urlParameters = "get-document/" + documentID;

            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;

                if (root.TryGetProperty("status", out JsonElement status) && status.GetString() == "1")
                {
                    if (root.TryGetProperty("document", out JsonElement jsonElement))
                    {
                        JsonElement objectConfig;

                        if (jsonElement.TryGetProperty("config", out objectConfig))
                        {
                            try
                            {
                                initialization(doc, objectConfig);
                            }
                            catch (Exception e)
                            {
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای در مقداردهی رخ داده\n" + e.Message, email: StringConstant.SupportEmail);
                                this.CloseFormRequest();
                            }

                            try
                            {
                                initializeControls();
                            }
                            catch (Exception e)
                            {
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای رخ داده\n" + e.Message, email: StringConstant.SupportEmail);
                                this.CloseFormRequest();
                            }
                            try
                            {
                                initializeControlNameOfAllah(doc);
                            }
                            catch (Exception e)
                            {
                                DedicatedFunctions.ShowErrorMessage("خطا در مقدار دهی اولیه کلمه بسم الله\n" + e.Message);
                                this.CloseFormRequest();
                            }

                            try
                            {
                                validateControls();
                            }
                            catch (Exception e)
                            {
                                DedicatedFunctions.ShowErrorMessage("خطا در صحت اطلاعات اولیه\n" + e.Message);
                                this.CloseFormRequest();
                            }
                        }
                        loadingControl.IsEnabled = false;
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("پاسخ سرور به برنامه دارای اشکال است",
                            (int)ErrorCodes.UnexpectedServerResponse, StringConstant.SupportEmail);
                        this.CloseFormRequest();
                        loadingControl.IsEnabled = false;
                    }
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("شناسه سندی که ساخته شده در سرور وجود ندارد!",
                        (int)ErrorCodes.UnexpectedServerResponse, StringConstant.SupportEmail);
                    this.CloseFormRequest();
                    loadingControl.IsEnabled = false;
                }
            },
            OnFailed =>
            {
                System.Windows.Forms.DialogResult dialogResult = DedicatedFunctions.ShowMessage(ErrorMessages.ErrorServiceUnavailable,
                    System.Windows.Forms.MessageBoxButtons.RetryCancel,
                    System.Windows.Forms.MessageBoxIcon.Error);

                if (dialogResult == System.Windows.Forms.DialogResult.Retry)
                {
                    getDocumentDataFromServer(token, documentID);
                }
                else
                {
                    this.CloseFormRequest();
                }

            });
        }

        #endregion
        private void Expander_Expanded(object sender, RoutedEventArgs e)
        {
            if (((Expander)sender).Name == expanderNameOfAllah.Name)
            {
                expanderEducationDetails.IsExpanded = false;
                expanderDocumentDetails.IsExpanded = false;
                expanderAbstractAndKeywords.IsExpanded = false;
                lstBoxNameOfAllah.ScrollIntoView(lstBoxNameOfAllah.SelectedItem);
                lstBoxNameOfAllah.Focus();

            }
            else if (((Expander)sender).Name == expanderEducationDetails.Name)
            {
                expanderNameOfAllah.IsExpanded = false;
                expanderDocumentDetails.IsExpanded = false;
                expanderAbstractAndKeywords.IsExpanded = false;
            }
            else if (((Expander)sender).Name == expanderDocumentDetails.Name)
            {
                expanderNameOfAllah.IsExpanded = false;
                expanderEducationDetails.IsExpanded = false;
                expanderAbstractAndKeywords.IsExpanded = false;
            }
            else if (((Expander)sender).Name == expanderAbstractAndKeywords.Name)
            {
                expanderNameOfAllah.IsExpanded = false;
                expanderEducationDetails.IsExpanded = false;
                expanderDocumentDetails.IsExpanded = false;
            }

        }

        #region buttons
        private void BtnConfirmChanges_Click(object sender, System.Windows.RoutedEventArgs es)
        {
            if (validateControls())
            {
                loadingControl.Tag = "لطفا منتظر بمانید";
                loadingControl.IsEnabled = true;
                try
                {
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر اطلاعات");

                    #region Change Besmellah
                    String fontName;
                    if (NameOfAllahFontType == 1)
                        fontName = Constants.FontNames.fontBesmellah1;
                    else if (NameOfAllahFontType == 2)
                        fontName = Constants.FontNames.fontBesmellah2;
                    else if (NameOfAllahFontType == 3)
                        fontName = Constants.FontNames.fontBesmellah3;
                    else if (NameOfAllahFontType == 4)
                        fontName = Constants.FontNames.fontBesmellah4;
                    else
                        fontName = Constants.FontNames.fontBesmellah1;

                    DedicatedFunctions.changeContentControlFontName(doc, ContentControlNames._field_InTheNameOfAllah.ToString(), fontName);
                    DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_InTheNameOfAllah.ToString(), lblNameOfAllah.Text);
                    #endregion

                    Universities university = DedicatedFunctions.getUniversity(doc);
                    DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
                    TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

                    //TextBox and ComboBox
                    foreach (ChangeContentsModel content in contents)
                    {
                        if (content.ContentControlName == ContentControlNames._field_Advisor_Fa.ToString())
                        {
                            string title = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Advisor_Title_Fa);

                            if (!string.IsNullOrEmpty(title))
                            {
                                if (!string.IsNullOrEmpty(((TextBox)content.Control).Text.Trim()))
                                    DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Advisor_Title_Fa.ToString(), title);
                                else
                                    DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Advisor_Title_Fa.ToString(), "");
                            }
                        }
                        else if (content.ContentControlName == ContentControlNames._field_Advisor_En.ToString())
                        {
                            if (!string.IsNullOrEmpty(((TextBox)content.Control).Text.Trim()))
                            {
                                if (university == Universities.YazdUniversity)
                                {
                                    Range previousRange = doc.ActiveWindow.Selection.Range;
                                    Microsoft.Office.Interop.Word.ContentControl[] ccsAdvisorEn = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Advisor_Title_En.ToString());

                                    foreach (Microsoft.Office.Interop.Word.ContentControl contentControl in ccsAdvisorEn)
                                    {

                                        bool previousLockState = contentControl.LockContents;

                                        contentControl.LockContents = false;
                                        contentControl.Range.Select();

                                        doc.ActiveWindow.Selection.SetRange(doc.ActiveWindow.Selection.Range.Start - 1, doc.ActiveWindow.Selection.Range.Start - 1);
                                        doc.ActiveWindow.Selection.TypeParagraph();

                                        contentControl.LockContents = previousLockState;

                                        previousRange.Select();

                                    }
                                }

                                string title = TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_Advisor_Title_En);

                                if (!string.IsNullOrEmpty(title))
                                {
                                    DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Advisor_Title_En.ToString(), title);
                                }
                            }
                            else
                            {
                                if (university == Universities.YazdUniversity)
                                {
                                    Range previousRange = doc.ActiveWindow.Selection.Range;
                                    Microsoft.Office.Interop.Word.ContentControl[] ccsAdvisorEn = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Advisor_Title_En.ToString());

                                    if (ccsAdvisorEn != null && !string.IsNullOrEmpty(ccsAdvisorEn[ccsAdvisorEn.Length - 1].Range.Text.Trim()))
                                    {
                                        ccsAdvisorEn[ccsAdvisorEn.Length - 1].Range.Select();

                                        doc.ActiveWindow.Selection.HomeKey();
                                        doc.ActiveWindow.Selection.TypeBackspace();

                                        previousRange.Select();
                                    }
                                }
                                DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Advisor_Title_En.ToString(), "");
                            }

                        }
                        else if (content.ContentControlName == ContentControlNames._field_AreaOfStudy_Fa.ToString())
                        {
                            if (!string.IsNullOrEmpty(((TextBox)content.Control).Text.Trim()))
                                DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_AreaOfStudy_Title_Fa.ToString(), TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_AreaOfStudy_Title_Fa));
                            else
                                DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_AreaOfStudy_Title_Fa.ToString(), "");

                        }
                        else if (content.ContentControlName == ContentControlNames._field_AreaOfStudy_En.ToString())
                        {
                            if (!string.IsNullOrEmpty(((TextBox)content.Control).Text.Trim()))
                                DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_AreaOfStudy_Title_En.ToString(), TemplateAccess.getCustomTitle(university, templateType, ContentControlNames._field_AreaOfStudy_Title_En));
                            else
                                DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_AreaOfStudy_Title_En.ToString(), "");
                        }

                        if (content.Control is TextBox textBox)
                        {
                            try
                            {
                                if (textBox.IsEnabled && textBox.Visibility == Visibility.Visible)
                                {
                                    if (content.useVariable())
                                    {
                                        DedicatedFunctions.setORAddStaticVariableValue(doc, content.VariableName, textBox.Text);
                                        DedicatedFunctions.changeContentControlContents(doc, content.ContentControlName, textBox.Text);
                                    }
                                    else
                                    {
                                        DedicatedFunctions.changeContentControlContents(doc, content.ContentControlName, textBox.Text);
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره در تغییر اطلاعات در " + HintAssist.GetHint(textBox) + " به وجود آمده، لطفا خطا را گزارش دهید\n" + e.Message);
                                loadingControl.IsEnabled = false;
                            }
                        }
                        else if (content.Control is ComboBox comboBox)
                        {
                            try
                            {
                                if (comboBox.IsEnabled && comboBox.Visibility == Visibility.Visible)
                                {
                                    if (content.useVariable())
                                    {
                                        string tempNothing = comboBox.Text;
                                        if (tempNothing == ComboBoxData.nothingFa)
                                            tempNothing = "";

                                        DedicatedFunctions.setORAddStaticVariableValue(doc, content.VariableName, tempNothing);
                                        DedicatedFunctions.changeContentControlContents(doc, content.ContentControlName, tempNothing);

                                        if (content.Control == comboUniversity)
                                        {
                                            if (content.VariableName == VariableFieldIDs._variable_field_University_Fa.ToString())
                                            {

                                                DedicatedFunctions.setORAddStaticVariableValue(doc, VariableFieldIDs._variable_field_Branch_Fa.ToString(), UniversitiesData.getBranchFaOfUniversity(university));
                                                DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Branch_Fa.ToString(), DedicatedFunctions.getStaticVariableValue(doc, content.VariableName));

                                                DedicatedFunctions.setORAddStaticVariableValue(doc, VariableFieldIDs._variable_field_Branch_En.ToString(), UniversitiesData.getBranchEnOfUniversity(university));
                                                DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Branch_En.ToString(), DedicatedFunctions.getStaticVariableValue(doc, content.VariableName));
                                            }


                                            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableFieldIDs._variable_field_University_En.ToString(), UniversitiesData.getEnglishUniversities()[comboBox.SelectedIndex]);
                                            DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_University_En.ToString(), UniversitiesData.getEnglishUniversities()[comboBox.SelectedIndex]);
                                        }
                                        //else if (cccbm.VariableName == VariableFieldNames._variable_field_FacultyDepartment_Fa.ToString())
                                        //{
                                        //    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableFieldNames._variable_field_FacultyDepartment_En.ToString(), DataComboBox.Faculty_En[((ComboBox)control).SelectedIndex]);
                                        //    DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_FacultyDepartment_En.ToString(), DataComboBox.Faculty_En[((ComboBox)control).SelectedIndex]);
                                        //}
                                        else if (content.Control == comboDepartment && txtBoxCustomDepartmentEn.Visibility == Visibility.Collapsed)
                                        {
                                            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableFieldIDs._variable_field_Department_En.ToString(), DepartmentsData.getEnglistDepartments(university, true)[comboBox.SelectedIndex]);
                                            DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Department_En.ToString(), DepartmentsData.getEnglistDepartments(university, true)[comboBox.SelectedIndex]);
                                        }
                                        else if (content.Control == comboGroup && txtBoxCustomGroupEn.Visibility == Visibility.Collapsed)
                                        {
                                            string temp = DepartmentsData.getEnglishGroups(university, comboDepartment.SelectedItem as string)[comboBox.SelectedIndex];
                                            if (temp == ComboBoxData.nothingEn)
                                                temp = "";

                                            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableFieldIDs._variable_field_Group_En.ToString(), temp);
                                            DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Group_En.ToString(), temp);
                                        }
                                        else if (content.Control == comboAcademicDegree)
                                        {
                                            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableFieldIDs._variable_field_AcademicDegree_En.ToString(), ComboBoxDataAcademicDegree.AcademicDegree_En[comboBox.SelectedIndex]);
                                            DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_AcademicDegree_En.ToString(), ComboBoxDataAcademicDegree.AcademicDegree_En[comboBox.SelectedIndex]);
                                        }
                                    }
                                    else
                                    {
                                        DedicatedFunctions.changeContentControlContents(doc, content.ContentControlName, comboBox.Text);

                                        if (content.Control == comboUniversity)
                                        {
                                            DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_University_En.ToString(), UniversitiesData.getEnglishUniversities()[comboBox.SelectedIndex]);
                                        }
                                        //else if (cccbm.ContentControlName == ContentControlNames._field_FacultyDepartment_Fa.ToString())
                                        //    DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_FacultyDepartment_En.ToString(), DataComboBox.Faculty_En[((ComboBox)control).SelectedIndex]);
                                        if (content.Control == comboDepartment && txtBoxCustomDepartmentEn.Visibility == Visibility.Collapsed)
                                        {
                                            DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Department_En.ToString(), DepartmentsData.getEnglistDepartments(university, true)[comboBox.SelectedIndex]);
                                        }
                                        if (content.Control == comboGroup && txtBoxCustomGroupEn.Visibility == Visibility.Collapsed)
                                        {
                                            string temp = DepartmentsData.getEnglishGroups(university, comboDepartment.SelectedItem as string)[comboBox.SelectedIndex];
                                            if (temp == ComboBoxData.nothingEn)
                                                temp = "";

                                            DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Group_En.ToString(), temp);
                                        }
                                        else if (content.Control == comboAcademicDegree)
                                        {
                                            DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_AcademicDegree_En.ToString(), ComboBoxDataAcademicDegree.AcademicDegree_En[comboBox.SelectedIndex]);
                                        }

                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره در تغییر اطلاعات در " + HintAssist.GetHint(comboBox) + " به وجود آمده، لطفا خطا را گزارش دهید\n" + e.Message);
                                loadingControl.IsEnabled = false;
                            }
                        }
                    }

                    JsonObject jsonVariables = DedicatedFunctions.variablesToJsonServer(doc);

                    //add Abstract to Json 
                    Microsoft.Office.Interop.Word.ContentControl[] abstractContentControl = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_Fa.ToString());
                    if (abstractContentControl != null && abstractContentControl[0].Range != null && abstractContentControl.Length != 0)
                    {
                        string abstractText = abstractContentControl[0].Range.Text;
                        if (jsonVariables.ContainsKey(VariableFieldIDs._variable_field_Abstract_Fa.ToString()))
                            jsonVariables[VariableFieldIDs._variable_field_Abstract_Fa.ToString()] = abstractText;
                        else
                            jsonVariables.Add(VariableFieldIDs._variable_field_Abstract_Fa.ToString(), abstractText);
                    }

                    string token = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UserToken.ToString());
                    string documentID = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString());
                    string urlParameters = "save/parsanegar/1?id=" + documentID + "&config=" + jsonVariables.ToString();

                    Dispatcher.Invoke(() =>
                    {
                        DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
                        OnResult =>
                        {
                            JsonDocument document = JsonDocument.Parse(OnResult);
                            JsonElement root = document.RootElement;

                            root.TryGetProperty("updated", out JsonElement updatedAtElement);
                            string updatedAt = updatedAtElement.GetString();
                            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedAt.ToString(), updatedAt);
                            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedConfig.ToString(), updatedAt);

                            doc.UndoClear();
                            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                            DedicatedFunctions.saveDocument(doc);

                            CloseFormRequest?.Invoke();
                        },
                        OnFailed =>
                        {
                            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                            DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable + "\n" +
                                OnFailed.StatusCode + "> " + OnFailed.ReasonPhrase);
                            loadingControl.IsEnabled = false;
                            CloseFormRequest?.Invoke();
                            DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges, false);
                        });
                    });
                }
                catch (Exception e)
                {
                    DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره در تغییر اطلاعات به وجود آمده\n" + e.Message, email: StringConstant.SupportEmail);
                    loadingControl.IsEnabled = false;
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                    DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges, false);
                }
            }
            else
            {
                DedicatedFunctions.ShowErrorMessage("برخی فیلد ها یا خالی هستند یا به درستی وارد نشده اند.");
                loadingControl.IsEnabled = false;
            }
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
            CloseFormRequest?.Invoke();
        }
        #endregion

        #region events
        private void TextBoxMultiLine_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (e.Key == Key.Enter)
            {
                int lineCount = textBox.Text.Split(new[] { '\n' }, StringSplitOptions.None).Length - 1;
                if (lineCount >= 2)
                {
                    e.Handled = true;
                }
            }
        }

        private void ComboBesmellahFontType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            NameOfAllahFontType = comboBesmellahFontType.SelectedIndex + 1;

            lstBoxNameOfAllah.SelectedIndex = 0;
            lstBoxNameOfAllah.ScrollIntoView(lstBoxNameOfAllah.SelectedItem);
            lstBoxNameOfAllah.Focus();

            if (comboBesmellahFontType.SelectedIndex == 0)
            {
                lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
                lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
            }
            else if (comboBesmellahFontType.SelectedIndex == 1)
            {

                lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI2);
                lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI2);
            }
            else if (comboBesmellahFontType.SelectedIndex == 2)
            {
                lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI3);
                lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI3);
            }
            //else if(comboBesmellahFontType.SelectedIndex == 3)
            //{
            //	lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI4);
            //	lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI4);
            //}
        }

        private void ComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            ChangeContentsModel changeContentsModel = contents.Where(content => content.Control == comboBox).First();
            if (!changeContentsModel.IsOptional)
            {
                if (comboBox.SelectedIndex == -1)
                {
                    if (string.IsNullOrEmpty(comboBox.Text))
                    {
                        HintAssist.SetHelperText((ComboBox)sender, "فیلد نباید خالی باشد");
                        comboBox.Foreground = System.Windows.Media.Brushes.Red;
                        //comboBox.Margin = new Thickness(5 , 5 , 5 , 20);
                    }
                    else
                    {
                        HintAssist.SetHelperText((ComboBox)sender, "");
                        comboBox.Foreground = System.Windows.Media.Brushes.Black;
                        //comboBox.Margin = new Thickness(5);
                    }
                }
                else
                {
                    HintAssist.SetHelperText((ComboBox)sender, "");
                    comboBox.Foreground = System.Windows.Media.Brushes.Black;
                    //comboBox.Margin = new Thickness(5);
                }
            }

        }
        private void ComboBox_GotFocus(object sender, RoutedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            string tag = comboBox.Tag.ToString();

            if (!string.IsNullOrEmpty(tag.Trim()))
            {
                if (tag == "Persian")
                {
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);
                }
                else if (tag == "English")
                {
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.English);
                }
            }
        }
        private void ComboBox_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if (!comboBox.IsDropDownOpen)
            {
                if (comboBox.IsFocused)
                {
                    scrollContent.Focus();
                    e.Handled = true;
                }
            }
        }
        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if (comboBox.SelectedIndex == -1)
            {
                return;
            }

            if (comboBox == comboUniversity && comboUniversity.SelectedIndex != -1)
            {
                university = (Universities)comboUniversity.SelectedIndex;

                comboDepartment.ItemsSource = DepartmentsData.getPersianDepartments(university, true);

                resetComboBox(comboDepartment);
                resetComboBox(comboGroup);
                changeVisibilityEnglishControls(txtBoxCustomDepartmentEn, gridDepartment, Visibility.Collapsed);
                changeVisibilityEnglishControls(txtBoxCustomGroupEn, gridGroup, Visibility.Collapsed);

            }
            else if (comboBox == comboDepartment)
            {
                comboGroup.ItemsSource = DepartmentsData.getPersianGroups(university, comboDepartment.SelectedItem as string);

                resetComboBox(comboGroup);
                changeVisibilityEnglishControls(txtBoxCustomGroupEn, gridGroup, Visibility.Collapsed);
            }

            ChangeContentsModel changeContentsModel = contents.Where(content => content.Control == comboBox).First();

            //Show Message on not Validate
            if (!changeContentsModel.IsOptional)
            {
                if (comboBox.SelectedIndex == -1)
                {
                    if (!comboBox.IsEditable)
                    {
                        HintAssist.SetHelperText((ComboBox)sender, "موردی انتخاب نشده است");
                        comboBox.Foreground = System.Windows.Media.Brushes.Red;
                        //comboBox.Margin = new Thickness(5 , 5 , 5 , 20);
                    }
                }
                else
                {
                    HintAssist.SetHelperText((ComboBox)sender, "");
                    comboBox.Foreground = System.Windows.Media.Brushes.Black;
                    //comboBox.Margin = new Thickness(5);
                }
            }



            TextBox textBoxInside = (comboBox.Template.FindName("PART_EditableTextBox", comboBox) as TextBox);

            if (textBoxInside != null)
            {
                textBoxInside.TextWrapping = TextWrapping.Wrap;
                textBoxInside.AcceptsReturn = false;
                textBoxInside.AcceptsTab = false;
            }

            if (comboBox.SelectedItem.ToString() != ComboBoxData.otherFa)
            {
                comboBox.IsEditable = false;
                comboBox.IsTextSearchEnabled = true;

                if (comboBox == comboDepartment)
                {
                    changeVisibilityEnglishControls(txtBoxCustomDepartmentEn, gridDepartment, Visibility.Collapsed);
                    txtBoxCustomDepartmentEn.Text = "";
                    HintAssist.SetHelperText(txtBoxCustomDepartmentEn, "");
                    txtBoxCustomDepartmentEn.Foreground = System.Windows.Media.Brushes.Black;
                }
                else if (comboBox == comboGroup)
                {
                    changeVisibilityEnglishControls(txtBoxCustomGroupEn, gridGroup, Visibility.Collapsed);
                    txtBoxCustomGroupEn.Text = "";
                    HintAssist.SetHelperText(txtBoxCustomGroupEn, "");
                    txtBoxCustomGroupEn.Foreground = System.Windows.Media.Brushes.Black;
                }
            }
            else// selected Other
            {
                comboBox.IsEditable = true;
                comboBox.IsTextSearchEnabled = false;

                textBoxInside.Focus();
                textBoxInside.Text = "";
                //textBox.Select(0, textBox.Text.Length);

                if (comboBox == comboDepartment)
                {
                    changeVisibilityEnglishControls(txtBoxCustomDepartmentEn, gridDepartment, Visibility.Visible);
                    HintAssist.SetHelperText(txtBoxCustomDepartmentEn, "");
                    txtBoxCustomDepartmentEn.Foreground = System.Windows.Media.Brushes.Black;
                }
                else if (comboBox == comboGroup)
                {
                    changeVisibilityEnglishControls(txtBoxCustomGroupEn, gridGroup, Visibility.Visible);
                    HintAssist.SetHelperText(txtBoxCustomGroupEn, "");
                    txtBoxCustomGroupEn.Foreground = System.Windows.Media.Brushes.Black;
                }
            }

            //validateControls();
        }


        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = ((TextBox)sender);

            ChangeContentsModel changeContentsModel = contents.Where(content => content.Control == textBox).First();


            if (!changeContentsModel.IsOptional)
            {
                if (string.IsNullOrEmpty(textBox.Text))
                {
                    HintAssist.SetHelperText(textBox, "فیلد نباید خالی باشد");
                    textBox.Foreground = System.Windows.Media.Brushes.Red;
                    //textBox.Margin = new Thickness(5 , 5 , 5 , 20);
                }
                else
                {
                    HintAssist.SetHelperText(textBox, "");
                    textBox.Foreground = System.Windows.Media.Brushes.Black;
                    //textBox.Margin = new Thickness(5);
                }
            }
        }
        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = ((TextBox)sender);

            string tag = textBox.Tag.ToString();

            if (!string.IsNullOrEmpty(tag.Trim()))
            {
                if (tag == "Persian")
                {
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);
                }
                else if (tag == "English")
                {
                    DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.English);
                }
            }
        }

        #endregion

        #region initialization

        private void initialization(Document doc, JsonElement jsonConfig)
        {
            JsonElement jsonVariables = JsonDocument.Parse(jsonConfig.GetString()).RootElement;
            //new Content(ContentControlNames._field_Type)
            //new Content(ContentControlNames._field_Icon_University)
            JsonElement jsonElement;

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_University_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(comboUniversity, ContentControlNames._field_University_Fa.ToString(),
                VariableFieldIDs._variable_field_University_Fa.ToString()));

                comboUniversity.ItemsSource = UniversitiesData.getPersianUniversities();
                comboUniversity.SelectedItem = jsonElement.GetString();
                comboUniversity.IsEnabled = false;
            }
            else
            {
                comboUniversity.Visibility = Visibility.Collapsed;
                comboUniversity.IsEnabled = false;
            }

            //contents.Add(
            //new ChangeContentsModel(ContentControlNames._field_FacultyDepartment_Fa.ToString(),
            //VariableFieldNames._variable_field_FacultyDepartment_Fa.ToString(),
            //ChangeContents.Content.CExpander.ExpanderEducationDetails_Fa,
            //ChangeContents.Content.CContentType.Variable,
            //ChangeContents.Content.CControlType.ComboBox,
            //true,
            //"پردیس",
            //contentsComboBox: DataComboBox.Faculty_Fa,
            //selectItemComboBox: DedicatedFunctions.getStaticVariableValue(doc, VariableFieldNames._variable_field_FacultyDepartment_Fa.ToString()))
            //    );

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Department_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(comboDepartment, ContentControlNames._field_Department_Fa.ToString(),
                VariableFieldIDs._variable_field_Department_Fa.ToString()));

                comboDepartment.ItemsSource = DepartmentsData.getPersianDepartments(DedicatedFunctions.getUniversity(doc), true);

                string selectedItem = jsonElement.GetString();

                setComboBoxSelectedItem(comboDepartment, txtBoxCustomDepartmentEn, gridDepartment, selectedItem);
            }
            else
            {
                gridDepartment.Visibility = Visibility.Collapsed;
                comboDepartment.Visibility = Visibility.Collapsed;
                comboDepartment.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Department_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxCustomDepartmentEn, ContentControlNames._field_Department_En.ToString(),
                VariableFieldIDs._variable_field_Department_En.ToString()
                ));
                txtBoxCustomDepartmentEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxCustomDepartmentEn.Visibility = Visibility.Collapsed;
                txtBoxCustomDepartmentEn.IsEnabled = false;
            }
            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Group_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(comboGroup, ContentControlNames._field_Group_Fa.ToString(),
                VariableFieldIDs._variable_field_Group_Fa.ToString()
                ));

                comboGroup.ItemsSource = DepartmentsData.getPersianGroups(university, comboDepartment.SelectedItem as string);
                string selectedItem = jsonElement.GetString();
                setComboBoxSelectedItem(comboGroup, txtBoxCustomGroupEn, gridGroup, selectedItem);
            }
            else
            {
                gridGroup.Visibility = Visibility.Collapsed;
                comboGroup.Visibility = Visibility.Collapsed;
                comboGroup.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Group_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxCustomGroupEn, ContentControlNames._field_Group_En.ToString(),
                VariableFieldIDs._variable_field_Group_En.ToString()
                ));
                txtBoxCustomGroupEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxCustomGroupEn.Visibility = Visibility.Collapsed;
                txtBoxCustomGroupEn.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_FieldOfStudy_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxFieldOfStudy, ContentControlNames._field_FieldOfStudy_Fa.ToString(),
                VariableFieldIDs._variable_field_FieldOfStudy_Fa.ToString()
                ));
                txtBoxFieldOfStudy.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxFieldOfStudy.Visibility = Visibility.Collapsed;
                txtBoxFieldOfStudy.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_FieldOfStudy_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxFieldOfStudyEn, ContentControlNames._field_FieldOfStudy_En.ToString(),
                VariableFieldIDs._variable_field_FieldOfStudy_En.ToString()
                ));
                txtBoxFieldOfStudyEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxFieldOfStudyEn.Visibility = Visibility.Collapsed;
                txtBoxFieldOfStudyEn.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_AreaOfStudy_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxAreaOfStudy, ContentControlNames._field_AreaOfStudy_Fa.ToString(),
                VariableFieldIDs._variable_field_AreaOfStudy_Fa.ToString(),
                isOptional: true));
                txtBoxAreaOfStudy.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxAreaOfStudy.Visibility = Visibility.Collapsed;
                txtBoxAreaOfStudy.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_AreaOfStudy_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxAreaOfStudyEn, ContentControlNames._field_AreaOfStudy_En.ToString(),
                VariableFieldIDs._variable_field_AreaOfStudy_En.ToString(),
                isOptional: true));
                txtBoxAreaOfStudyEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxAreaOfStudyEn.Visibility = Visibility.Collapsed;
                txtBoxAreaOfStudyEn.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(comboAcademicDegree, ContentControlNames._field_AcademicDegree_Fa.ToString(),
                VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString()
                ));
                comboAcademicDegree.ItemsSource = ComboBoxDataAcademicDegree.AcademicDegree_Fa;
                comboAcademicDegree.SelectedItem = jsonElement.GetString();
                comboAcademicDegree.IsEnabled = false;
            }
            else
            {
                comboAcademicDegree.Visibility = Visibility.Collapsed;
                comboAcademicDegree.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_NameOfCourse_Fa.ToString(), out jsonElement) && DedicatedFunctions.getDocumentType(doc) == DocumentTypes.SchoolResearch)
            {
                contents.Add(
                new ChangeContentsModel(txtBoxNameOfCourseFa, ContentControlNames._field_NameOfCourse_Fa.ToString(),
                VariableFieldIDs._variable_field_NameOfCourse_Fa.ToString()
                ));
                txtBoxNameOfCourseFa.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxNameOfCourseFa.Visibility = Visibility.Collapsed;
                txtBoxNameOfCourseFa.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Title_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxTitleFa, ContentControlNames._field_Title_Fa.ToString(),
                VariableFieldIDs._variable_field_Title_Fa.ToString()
                ));
                txtBoxTitleFa.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxTitleFa.Visibility = Visibility.Collapsed;
                txtBoxTitleFa.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Title_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxTitleEn, ContentControlNames._field_Title_En.ToString(),
                VariableFieldIDs._variable_field_Title_En.ToString()
                ));
                txtBoxTitleEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxTitleEn.Visibility = Visibility.Collapsed;
                txtBoxTitleEn.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Supervisor_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxSupervisorFa, ContentControlNames._field_Supervisor_Fa.ToString(),
                VariableFieldIDs._variable_field_Supervisor_Fa.ToString()
                ));
                txtBoxSupervisorFa.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxSupervisorFa.Visibility = Visibility.Collapsed;
                txtBoxSupervisorFa.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Supervisor_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxSupervisorEn, ContentControlNames._field_Supervisor_En.ToString(),
                VariableFieldIDs._variable_field_Supervisor_En.ToString()
                ));
                txtBoxSupervisorEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxSupervisorEn.Visibility = Visibility.Collapsed;
                txtBoxSupervisorEn.IsEnabled = false;
            }


            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Advisor_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxAdvisorFa, ContentControlNames._field_Advisor_Fa.ToString(),
                VariableFieldIDs._variable_field_Advisor_Fa.ToString(), isOptional: true
                ));
                txtBoxAdvisorFa.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxAdvisorFa.Visibility = Visibility.Collapsed;
                txtBoxAdvisorFa.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Advisor_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxAdvisorEn, ContentControlNames._field_Advisor_En.ToString(),
                VariableFieldIDs._variable_field_Advisor_En.ToString(), isOptional: true
                ));
                txtBoxAdvisorEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxAdvisorEn.Visibility = Visibility.Collapsed;
                txtBoxAdvisorEn.IsEnabled = false;
            }


            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Author_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxAuthorFa, ContentControlNames._field_Author_Fa.ToString(),
                VariableFieldIDs._variable_field_Author_Fa.ToString()
                ));
                txtBoxAuthorFa.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxAuthorFa.Visibility = Visibility.Collapsed;
                txtBoxAuthorFa.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_Author_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxAuthorEn, ContentControlNames._field_Author_En.ToString(),
                VariableFieldIDs._variable_field_Author_En.ToString()
                ));
                txtBoxAuthorEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxAuthorEn.Visibility = Visibility.Collapsed;
                txtBoxAuthorEn.IsEnabled = false;
            }


            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_DefenseDate_Fa.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxDefenseDateFa, ContentControlNames._field_DefenseDate_Fa.ToString(),
                VariableFieldIDs._variable_field_DefenseDate_Fa.ToString(), isOptional: true
                ));
                txtBoxDefenseDateFa.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxDefenseDateFa.Visibility = Visibility.Collapsed;
                txtBoxDefenseDateFa.IsEnabled = false;
            }

            if (jsonVariables.TryGetProperty(VariableFieldIDs._variable_field_DefenseDate_En.ToString(), out jsonElement))
            {
                contents.Add(
                new ChangeContentsModel(txtBoxDefenseDateEn, ContentControlNames._field_DefenseDate_En.ToString(),
                VariableFieldIDs._variable_field_DefenseDate_En.ToString(), isOptional: true
                ));
                txtBoxDefenseDateEn.Text = jsonElement.GetString();
            }
            else
            {
                txtBoxDefenseDateEn.Visibility = Visibility.Collapsed;
                txtBoxDefenseDateEn.IsEnabled = false;
            }


            if (DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_Fa.ToString()) != null)
            {
                contents.Add(
                new ChangeContentsModel(txtBoxAbstractFa, ContentControlNames._field_Abstract_Fa.ToString(), null
                ));

                Microsoft.Office.Interop.Word.ContentControl[] abstractContentControl = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_Fa.ToString());
                if (abstractContentControl != null && abstractContentControl.Length != 0)
                {
                    Range rangeAbstract = abstractContentControl[0].Range;
                    if (rangeAbstract != null)
                        txtBoxAbstractFa.Text = abstractContentControl[0].Range.Text;
                }
            }
            else
            {
                txtBoxAbstractFa.Visibility = Visibility.Collapsed;
                txtBoxAbstractFa.IsEnabled = false;
            }

            if (DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_En.ToString()) != null)
            {
                contents.Add(
                new ChangeContentsModel(txtBoxAbstractEn, ContentControlNames._field_Abstract_En.ToString(), null
                ));

                Range abstractRange = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Abstract_En.ToString())[0].Range;
                if (abstractRange != null)
                    txtBoxAbstractEn.Text = abstractRange.Text;
            }
            else
            {
                txtBoxAbstractEn.Visibility = Visibility.Collapsed;
                txtBoxAbstractEn.IsEnabled = false;
            }


            if (DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Keywords_Fa.ToString()) != null)
            {
                contents.Add(
                new ChangeContentsModel(txtBoxKeywordFa, ContentControlNames._field_Keywords_Fa.ToString(), null
                ));

                Range keywordRange = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Keywords_Fa.ToString())[0].Range;
                if (keywordRange != null)
                    txtBoxKeywordFa.Text = keywordRange.Text;
            }
            else
            {
                txtBoxKeywordFa.Visibility = Visibility.Collapsed;
                txtBoxKeywordFa.IsEnabled = false;
            }

            if (DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Keywords_En.ToString()) != null)
            {
                contents.Add(
                new ChangeContentsModel(txtBoxKeywordEn, ContentControlNames._field_Keywords_En.ToString(), null
                ));

                Range keywordRange = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Keywords_En.ToString())[0].Range;

                if (keywordRange != null)
                    txtBoxKeywordEn.Text = keywordRange.Text;
            }
            else
            {
                txtBoxKeywordEn.Visibility = Visibility.Collapsed;
                txtBoxKeywordEn.IsEnabled = false;
            }

        }
        private void initializeControls()
        {
            foreach (ChangeContentsModel content in contents)
            {
                if (content.Control is TextBox textBox)
                {
                    textBox.TextChanged += TextBox_TextChanged;
                    textBox.GotFocus += TextBox_GotFocus;
                }
                else if (content.Control is ComboBox comboBox)
                {
                    comboBox.SelectionChanged += ComboBox_SelectionChanged;
                    comboBox.LostFocus += ComboBox_LostFocus;
                    comboBox.GotFocus += ComboBox_GotFocus;
                    comboBox.PreviewMouseWheel += ComboBox_PreviewMouseWheel;
                }
            }
        }

        private void initializeControlNameOfAllah(Document doc)
        {

            comboBesmellahFontType.SelectionChanged += ComboBesmellahFontType_SelectionChanged;
            //lstBoxNameOfAllah.SelectionChanged += LstBoxNameOfAllah_SelectionChanged;

            comboBesmellahFontType.Items.Clear();
            comboBesmellahFontType.Items.Add("قالب نوع اول");
            comboBesmellahFontType.Items.Add("قالب نوع دوم");
            comboBesmellahFontType.Items.Add("قالب نوع سوم");
            //comboBesmellahFontType.Items.Add("قالب نوع چهارم");//Disabled

            lstBoxNameOfAllah.Items.Clear();
            lstBoxNameOfAllah.ItemsSource = ComboBoxData.InTheNameOfAllah;
            lstBoxNameOfAllah.SelectedIndex = 0;



            //string nameOfAllah = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_NameOfAllah.ToString())[0].Range.Text;
            Microsoft.Office.Interop.Word.ContentControl[] nameOfAllahContentControls = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_InTheNameOfAllah.ToString());

            if (nameOfAllahContentControls != null)
            {
                string nameOfAllah = nameOfAllahContentControls[0].Range.Text;

                string currentNameOfAllahFontName = nameOfAllahContentControls[0].Range.Font.Name;

                if (currentNameOfAllahFontName == Constants.FontNames.fontBesmellah1)
                {
                    lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
                    lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
                    NameOfAllahFontType = 1;
                    comboBesmellahFontType.SelectedIndex = 0;
                }
                else if (currentNameOfAllahFontName == Constants.FontNames.fontBesmellah2)
                {
                    lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI2);
                    lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI2);
                    NameOfAllahFontType = 2;
                    comboBesmellahFontType.SelectedIndex = 1;
                }
                else if (currentNameOfAllahFontName == Constants.FontNames.fontBesmellah3)
                {
                    lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI3);
                    lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI3);
                    NameOfAllahFontType = 3;
                    comboBesmellahFontType.SelectedIndex = 2;
                }
                //else if(currentNameOfAllahFontName == Constants.FontNames.fontBesmellah4)
                //{
                //	lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI4);
                //	lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI4);
                //	NameOfAllahFontType = 4;
                //	comboBesmellahFontType.SelectedIndex = 3;
                //}

                if (lstBoxNameOfAllah.Items.Count != 0 && nameOfAllah != null)
                    lstBoxNameOfAllah.SelectedItem = nameOfAllah;
            }
        }

        #endregion

        #region Methods
        private bool validateControls()
        {
            bool isValid = true;

            foreach (ChangeContentsModel content in contents)
            {
                if (content.Control is TextBox textBox)
                {
                    if (!content.IsOptional)
                    {
                        if (textBox.Visibility == Visibility.Visible && string.IsNullOrEmpty(textBox.Text))
                        {
                            isValid = false;
                            HintAssist.SetHelperText(textBox, "فیلد نباید خالی باشد");
                            textBox.Foreground = System.Windows.Media.Brushes.Red;
                            //(textBox.Margin = new Thickness(5 , 5 , 5 , 20);
                        }
                        else
                        {
                            HintAssist.SetHelperText(textBox, "");
                            textBox.Foreground = System.Windows.Media.Brushes.Black;
                            //textBox.Margin = new Thickness(5);
                        }
                    }
                }
                else if (content.Control is ComboBox comboBox)
                {
                    if (!content.IsOptional)
                    {
                        if (comboBox.Visibility == Visibility.Visible && comboBox.SelectedIndex == -1)
                        {
                            if (!comboBox.IsEditable)
                            {
                                isValid = false;
                                HintAssist.SetHelperText(comboBox, "موردی انتخاب نشده است");
                                comboBox.Foreground = System.Windows.Media.Brushes.Red;
                                //comboBox.Margin = new Thickness(5 , 5 , 5 , 20);
                            }
                        }
                        else
                        {
                            HintAssist.SetHelperText(comboBox, "");
                            comboBox.Foreground = System.Windows.Media.Brushes.Black;
                            //comboBox.Margin = new Thickness(5);
                        }
                    }
                }
            }
            return isValid;
        }
        private void changeVisibilityEnglishControls(TextBox englishTextBox, Grid grid, Visibility visibility)
        {
            englishTextBox.Visibility = visibility;

            if (visibility == Visibility.Collapsed)
            {
                grid.ColumnDefinitions[1].Width = new GridLength(1, GridUnitType.Auto);
            }
            else
            {
                grid.ColumnDefinitions[1].Width = new GridLength(1, GridUnitType.Star);
            }
        }

        private void setComboBoxSelectedItem(ComboBox comboBox, TextBox englishTextBox, Grid grid, string text)
        {
            if (text == "" && comboBox == comboGroup)
                text = ComboBoxData.nothingFa;

            if (comboBox.Items.Contains(text) && text != ComboBoxData.otherFa)
            {
                comboBox.IsEditable = false;
                comboBox.IsTextSearchEnabled = true;
                comboBox.SelectedItem = text;

                changeVisibilityEnglishControls(englishTextBox, grid, Visibility.Collapsed);
            }
            else
            {
                if (comboBox.Items.Contains(ComboBoxData.otherFa))
                {
                    comboBox.SelectedItem = ComboBoxData.otherFa;
                }
                comboBox.IsEditable = true;
                comboBox.IsTextSearchEnabled = false;
                comboBox.Text = text;

                changeVisibilityEnglishControls(englishTextBox, grid, Visibility.Visible);
            }
        }
        private void resetComboBox(ComboBox comboBox)
        {
            comboBox.IsEditable = false;
            comboBox.IsTextSearchEnabled = true;
            comboBox.Text = "";
            comboBox.SelectedIndex = -1;
            HintAssist.SetHelperText(comboBox, "");
            comboBox.Foreground = System.Windows.Media.Brushes.Black;
        }

        #endregion

        /// <summary>
        /// Finds a Child of a given item in the visual tree. 
        /// </summary>
        /// <param name="parent">A direct parent of the queried item.</param>
        /// <typeparam name="T">The type of the queried item.</typeparam>
        /// <param name="childName">x:Name or Name of child. </param>
        /// <returns>The first parent item that matches the submitted type parameter. 
        /// If not matching item can be found, 
        /// a null parent is being returned.</returns>
        public static T FindChild<T>(DependencyObject parent, string childName)
           where T : DependencyObject
        {
            // Confirm parent and childName are valid. 
            if (parent == null) return null;

            T foundChild = null;

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                // If the child is not of the request child type child
                T childType = child as T;
                if (childType == null)
                {
                    // recursively drill down the tree
                    foundChild = FindChild<T>(child, childName);

                    // If the child is found, break so we do not overwrite the found child. 
                    if (foundChild != null) break;
                }
                else if (!string.IsNullOrEmpty(childName))
                {
                    var frameworkElement = child as FrameworkElement;
                    // If the child's name is set for search
                    if (frameworkElement != null && frameworkElement.Name == childName)
                    {
                        // if the child's name is of the request name
                        foundChild = (T)child;
                        break;
                    }
                }
                else
                {
                    // child element found.
                    foundChild = (T)child;
                    break;
                }
            }

            return foundChild;
        }

        #region set Property
        //
        //public static readonly DependencyProperty DependencyContentProperty = DependencyProperty.Register("Contents" , typeof(object) , typeof(UIElement) , new FrameworkPropertyMetadata(null , FrameworkPropertyMetadataOptions.AffectsRender));
        //public static void setContentProperty(UIElement uIElement , object value)
        //{
        //	uIElement.SetValue(DependencyContentProperty , value);
        //}
        //public static object getContentProperty(UIElement uIElement)
        //{
        //	return (object)uIElement.GetValue(DependencyContentProperty);
        //}

        #endregion

    }
}
