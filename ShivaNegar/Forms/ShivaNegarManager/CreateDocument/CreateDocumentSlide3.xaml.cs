using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MaterialDesignThemes.Wpf;
using ShivaNegar.Constants;
using ShivaNegar.Constants.ComboBoxData;
using ShivaNegar.Forms.ShivaNegarManager.CreateDocument.Models;
using static ShivaNegar.Forms.ShivaNegarManager.CreateDocument.Models.CreateDocumentControlModel;

namespace ShivaNegar.Forms.ShivaNegarManager.CreateDocument
{
    public partial class CreateDocumentSlide3 : UserControl
    {
        private List<CreateDocumentControlModel> textBoxControlModels;
        private List<CreateDocumentControlModel> comboBoxcontrolModels;

        private DocumentTypes documentType;

        //Properties
        public Universities University { get; private set; }

        public string UniversityFa { get; private set; }
        public string UniversityEn { get; private set; }
        public string BranchFa { get; private set; }
        public string BranchEn { get; private set; }
        public string DepartmentFa { get; private set; }
        public string DepartmentEn { get; private set; }
        public string GroupFa { get; private set; }
        public string GroupEn { get; private set; }
        public string FieldOfStudyFa { get; private set; }
        public string FieldOfStudyEn { get; private set; }
        public string AreaOfStudyFa { get; private set; }
        public string AreaOfStudyEn { get; private set; }
        public string AcademicDegreeFa { get; private set; }
        public string AcademicDegreeEn { get; private set; }

        internal static readonly string[] AcademicDegree_Fa2 =
        {
            AcademicDegreeValues.AcademicDegree_AssociateOfScienceFa,
            //AcademicDegreeValues.AcademicDegree_BachelorOfScienceFa,
            AcademicDegreeValues.AcademicDegree_PartTimeBachelorOfScienceFa,
            //AcademicDegreeValues.AcademicDegree_MasterOfScienceFa,
            //AcademicDegreeValues.AcademicDegree_DoctoralFa,
        };

        public CreateDocumentSlide3()
        {
            InitializeComponent();

            btnForward.IsEnabled = false;

            //comboDepartment.ItemsSource = DepartmentsData.getPersianDepartments();
            comboAcademicDegree.ItemsSource = AcademicDegree_Fa2;

            List<string> universities = UniversitiesData.getPersianUniversities();
            for (int i = 0; i < universities.Count; i++)
            {
                string branchFa = UniversitiesData.getBranchFaOfUniversity((Universities)i);

                if (!string.IsNullOrEmpty(branchFa))
                    universities[i] = universities[i] + " " + branchFa;
            }
            comboUniversity.ItemsSource = universities;


            comboBoxcontrolModels = new List<CreateDocumentControlModel>()
            {
                new CreateDocumentControlModel(comboUniversity,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(comboDepartment,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(comboGroup,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(comboAcademicDegree,CreateDocumentControlModel.ControlLevels.Essential),
            };
            foreach (var controlModel in comboBoxcontrolModels)
            {
                ComboBox comboBox = controlModel.Control as ComboBox;

                comboBox.SelectionChanged += ComboBox_SelectionChanged;
                comboBox.PreviewMouseWheel += ComboBox_PreviewMouseWheel;
                comboBox.GotFocus += ComboBox_GotFocus;
                //control.LostFocus += ComboBox_LostFocus;
            }

            textBoxControlModels = new List<CreateDocumentControlModel>()
            {
                new CreateDocumentControlModel(txtBoxCustomDepartmentEn,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxCustomGroupEn,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxFieldOfStudy,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxFieldOfStudyEn,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxAreaOfStudy,CreateDocumentControlModel.ControlLevels.Optional),
                new CreateDocumentControlModel(txtBoxAreaOfStudyEn,CreateDocumentControlModel.ControlLevels.Optional),
            };

            foreach (var controlModel in textBoxControlModels)
            {
                TextBox textBox = controlModel.Control as TextBox;

                textBox.TextChanged += TextBox_TextChanged;
                textBox.GotFocus += TextBox_GotFocus;
                textBox.LostFocus += TextBox_LostFocus;
            }
        }

        #region Events

        #region TextBox
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            for (int i = 0; i < textBoxControlModels.Count; i++)
            {
                if (textBox == (TextBox)textBoxControlModels[i].Control)
                {
                    textBoxControlModels[i].Validate = validateTextBox(textBox, textBoxControlModels[i].ControlLevel, false);
                    validateControls();
                    return;
                }
            }
        }
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            #region on Optional Fields, one language typed, change to Essential for complete other language field

            if (textBox == txtBoxAreaOfStudy)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudyEn).FirstOrDefault().ControlLevel = ControlLevels.Essential;
                }
                else
                {
                    textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudyEn).FirstOrDefault().ControlLevel = ControlLevels.Optional;
                }

                textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudyEn).FirstOrDefault().Validate = validateTextBox(txtBoxAreaOfStudyEn, textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudyEn).FirstOrDefault().ControlLevel, false);
                validateControls();
            }
            else if (textBox == txtBoxAreaOfStudyEn)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudy).FirstOrDefault().ControlLevel = ControlLevels.Essential;
                }
                else
                {
                    textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudy).FirstOrDefault().ControlLevel = ControlLevels.Optional;
                }

                textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudy).FirstOrDefault().Validate = validateTextBox(txtBoxAreaOfStudy, textBoxControlModels.Where(a => a.Control == txtBoxAreaOfStudy).FirstOrDefault().ControlLevel, false);
                validateControls();
            }
            #endregion

            for (int i = 0; i < textBoxControlModels.Count; i++)
            {
                if (textBox == (TextBox)textBoxControlModels[i].Control)
                {
                    textBoxControlModels[i].Validate = validateTextBox(textBox, textBoxControlModels[i].ControlLevel, true);
                    if (textBoxControlModels[i].Validate)
                        validateControls();
                    return;
                }
            }
        }

        private void TextBox_GotFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

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

        #region TextBox inside ComboBox
        private void TextBoxInsideComboDepartment_TextChanged(object sender, TextChangedEventArgs e)
        {
            for (int i = 0; i < comboBoxcontrolModels.Count; i++)
            {
                if (comboBoxcontrolModels[i].Control == comboDepartment)
                {
                    comboBoxcontrolModels[i].Validate = validateComboBox(comboDepartment, comboBoxcontrolModels[i].ControlLevel, true);
                    if (comboBoxcontrolModels[i].Validate)
                        validateControls();
                    return;
                }
            }
        }
        private void TextBoxInsideComboGroup_TextChanged(object sender, TextChangedEventArgs e)
        {
            for (int i = 0; i < comboBoxcontrolModels.Count; i++)
            {
                if (comboBoxcontrolModels[i].Control == comboGroup)
                {
                    comboBoxcontrolModels[i].Validate = validateComboBox(comboGroup, comboBoxcontrolModels[i].ControlLevel, true);
                    if (comboBoxcontrolModels[i].Validate)
                        validateControls();
                    return;
                }
            }
        }
        private void TextBoxInsideComboDepartment_LostFocus(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < comboBoxcontrolModels.Count; i++)
            {
                if (comboBoxcontrolModels[i].Control == comboDepartment)
                {
                    comboBoxcontrolModels[i].Validate = validateComboBox(comboDepartment, comboBoxcontrolModels[i].ControlLevel, true);
                    validateControls();
                    return;
                }
            }
        }
        private void TextBoxInsideComboGroup_LostFocus(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < comboBoxcontrolModels.Count; i++)
            {
                if (comboBoxcontrolModels[i].Control == comboGroup)
                {
                    comboBoxcontrolModels[i].Validate = validateComboBox(comboGroup, comboBoxcontrolModels[i].ControlLevel, false);
                    validateControls();
                    return;
                }
            }
        }

        #endregion


        #region ComboBox
        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if (comboBox.SelectedIndex == -1)
            {
                return;
            }

            if (comboBox == comboUniversity && comboUniversity.SelectedIndex != -1)
            {
                University = (Universities)comboUniversity.SelectedIndex;

                comboDepartment.ItemsSource = DepartmentsData.getPersianDepartments(University, true);

                comboBoxcontrolModels.Where(a => a.Control == comboBox).FirstOrDefault().Validate = false;
                comboBoxcontrolModels.Where(a => a.Control == comboDepartment).FirstOrDefault().Validate = false;
                comboBoxcontrolModels.Where(a => a.Control == comboGroup).FirstOrDefault().Validate = false;

                resetComboBox(comboDepartment);
                resetComboBox(comboGroup);
                changeVisibilityEnglishControls(txtBoxCustomDepartmentEn, gridDepartment, Visibility.Collapsed);
                changeVisibilityEnglishControls(txtBoxCustomGroupEn, gridGroup, Visibility.Collapsed);
            }
            else if (comboBox == comboDepartment)
            {
                comboGroup.ItemsSource = DepartmentsData.getPersianGroups(University, comboDepartment.SelectedItem as string);

                comboBoxcontrolModels.Where(a => a.Control == comboBox).FirstOrDefault().Validate = false;
                comboBoxcontrolModels.Where(a => a.Control == comboDepartment).FirstOrDefault().Validate = false;
                comboBoxcontrolModels.Where(a => a.Control == comboGroup).FirstOrDefault().Validate = false;

                resetComboBox(comboGroup);
                changeVisibilityEnglishControls(txtBoxCustomGroupEn, gridGroup, Visibility.Collapsed);
            }

            TextBox textBoxInside = (comboBox.Template.FindName("PART_EditableTextBox", comboBox) as TextBox);
            textBoxInside.TextWrapping = TextWrapping.Wrap;
            textBoxInside.AcceptsReturn = false;
            textBoxInside.AcceptsTab = false;

            if (comboBox.SelectedItem.ToString() != ComboBoxData.otherFa)
            {
                comboBox.IsEditable = false;
                comboBox.IsTextSearchEnabled = true;

                bool isValid = validateComboBox(comboBox, comboBoxcontrolModels.Where(a => a.Control == comboBox).FirstOrDefault().ControlLevel, false);
                comboBoxcontrolModels.Where(a => a.Control == comboBox).FirstOrDefault().Validate = isValid;

                if (comboBox == comboDepartment)
                {

                    gridDepartment.ColumnDefinitions[1].Width = new GridLength(1, GridUnitType.Auto);

                    txtBoxCustomDepartmentEn.Visibility = Visibility.Collapsed;
                    resetTextBoxControl(txtBoxCustomDepartmentEn);

                    textBoxInside.TextChanged -= TextBoxInsideComboDepartment_TextChanged;
                    textBoxInside.LostFocus -= TextBoxInsideComboDepartment_LostFocus;

                    textBoxControlModels.Where(a => a.Control == txtBoxCustomDepartmentEn).FirstOrDefault().Validate = isValid;
                }
                else if (comboBox == comboGroup)
                {

                    gridGroup.ColumnDefinitions[1].Width = new GridLength(1, GridUnitType.Auto);

                    txtBoxCustomGroupEn.Visibility = Visibility.Collapsed;
                    resetTextBoxControl(txtBoxCustomGroupEn);

                    textBoxInside.TextChanged -= TextBoxInsideComboGroup_TextChanged;
                    textBoxInside.LostFocus -= TextBoxInsideComboGroup_LostFocus;

                    textBoxControlModels.Where(a => a.Control == txtBoxCustomGroupEn).FirstOrDefault().Validate = isValid;
                }
            }
            else// selected Other
            {
                comboBox.IsEditable = true;
                comboBox.IsTextSearchEnabled = false;

                textBoxInside.Focus();
                textBoxInside.Text = "";
                //textBox.Select(0, textBox.Text.Length);

                comboBoxcontrolModels.Where(a => a.Control == comboBox).FirstOrDefault().Validate = false;

                if (comboBox == comboDepartment)
                {
                    changeVisibilityEnglishControls(txtBoxCustomDepartmentEn, gridDepartment, Visibility.Visible);

                    textBoxInside.TextChanged += TextBoxInsideComboDepartment_TextChanged;
                    textBoxInside.LostFocus += TextBoxInsideComboDepartment_LostFocus;

                    textBoxControlModels.Where(a => a.Control == txtBoxCustomDepartmentEn).FirstOrDefault().Validate = false;
                }
                else if (comboBox == comboGroup)
                {
                    changeVisibilityEnglishControls(txtBoxCustomGroupEn, gridGroup, Visibility.Visible);

                    textBoxInside.TextChanged += TextBoxInsideComboGroup_TextChanged;
                    textBoxInside.LostFocus += TextBoxInsideComboGroup_LostFocus;

                    textBoxControlModels.Where(a => a.Control == txtBoxCustomGroupEn).FirstOrDefault().Validate = false;
                }
            }

            validateControls();
        }

        private void ComboBox_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
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
        private void ComboBox_GotFocus(object sender, RoutedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;

            if (comboBox.FlowDirection == FlowDirection.RightToLeft)
            {
                DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.Persian);
            }
            else
            {
                DedicatedFunctions.changeKeyboardLanguage(KeyboardLanguage.English);
            }
        }
        #endregion

        #endregion

        #region Validators
        private bool validateControls()
        {
            bool isValid = true;
            foreach (CreateDocumentControlModel controlModel in comboBoxcontrolModels)
            {
                if (!controlModel.Validate && controlModel.ControlLevel != CreateDocumentControlModel.ControlLevels.Optional)
                {
                    isValid = false;
                    break;
                }
            }
            foreach (CreateDocumentControlModel controlModel in textBoxControlModels)
            {
                if (!controlModel.Validate && controlModel.ControlLevel != CreateDocumentControlModel.ControlLevels.Optional)
                {
                    isValid = false;
                    break;
                }
            }

            if (isValid)
            {
                UniversityFa = UniversitiesData.getPersianUniversities()[comboUniversity.SelectedIndex];
                UniversityEn = UniversitiesData.getEnglishUniversities()[comboUniversity.SelectedIndex];

                University = (Universities)comboUniversity.SelectedIndex;
                BranchFa = UniversitiesData.getBranchFaOfUniversity(University);
                BranchEn = UniversitiesData.getBranchEnOfUniversity(University);

                if (!comboDepartment.IsEditable)
                {
                    DepartmentFa = DepartmentsData.getPersianDepartments(University, true)[comboDepartment.SelectedIndex];
                    DepartmentEn = DepartmentsData.getEnglistDepartments(University, true)[comboDepartment.SelectedIndex];
                }
                else
                {
                    DepartmentFa = comboDepartment.Text;
                    DepartmentEn = txtBoxCustomDepartmentEn.Text;
                }

                if (!comboGroup.IsEditable)
                {
                    string tempFa = DepartmentsData.getPersianGroups(University, comboDepartment.SelectedItem as string)[comboGroup.SelectedIndex];
                    string tempEn = DepartmentsData.getEnglishGroups(University, comboDepartment.SelectedItem as string)[comboGroup.SelectedIndex];
                    if (tempFa == ComboBoxData.nothingFa || tempEn == ComboBoxData.nothingEn)
                    {
                        tempFa = "";
                        tempEn = "";
                    }

                    GroupFa = tempFa;
                    GroupEn = tempEn;
                }
                else
                {
                    GroupFa = comboGroup.Text;
                    GroupEn = txtBoxCustomGroupEn.Text;
                }

                if (documentType == DocumentTypes.Project)
                {
                    AcademicDegreeFa = ComboBoxDataAcademicDegree.AcademicDegree_Project_Fa[comboAcademicDegree.SelectedIndex];
                    AcademicDegreeEn = ComboBoxDataAcademicDegree.AcademicDegree_Project_En[comboAcademicDegree.SelectedIndex];
                }
                else if (documentType == DocumentTypes.Thesis)
                {
                    AcademicDegreeFa = ComboBoxDataAcademicDegree.AcademicDegree_Thesis_Fa[comboAcademicDegree.SelectedIndex];
                    AcademicDegreeEn = ComboBoxDataAcademicDegree.AcademicDegree_Thesis_En[comboAcademicDegree.SelectedIndex];
                }
                else if (documentType == DocumentTypes.Dissertation)
                {
                    AcademicDegreeFa = ComboBoxDataAcademicDegree.AcademicDegree_Dissertation_Fa[comboAcademicDegree.SelectedIndex];
                    AcademicDegreeEn = ComboBoxDataAcademicDegree.AcademicDegree_Dissertation_En[comboAcademicDegree.SelectedIndex];
                }
                else
                {
                    AcademicDegreeFa = AcademicDegree_Fa2[comboAcademicDegree.SelectedIndex];
                    AcademicDegreeEn = ComboBoxDataAcademicDegree.AcademicDegree_En[comboAcademicDegree.SelectedIndex];
                }
                FieldOfStudyFa = txtBoxFieldOfStudy.Text;
                FieldOfStudyEn = txtBoxFieldOfStudyEn.Text;
                AreaOfStudyFa = txtBoxAreaOfStudy.Text;
                AreaOfStudyEn = txtBoxAreaOfStudyEn.Text;

                btnForward.IsEnabled = true;
                return true;
            }
            else
            {
                UniversityFa = "";
                UniversityEn = "";
                DepartmentFa = "";
                DepartmentEn = "";
                GroupFa = "";
                GroupEn = "";
                AcademicDegreeFa = "";
                AcademicDegreeEn = "";
                FieldOfStudyFa = "";
                FieldOfStudyEn = "";
                AreaOfStudyFa = "";
                AreaOfStudyEn = "";

                btnForward.IsEnabled = false;
                return false;
            }
        }

        private bool validateComboBox(ComboBox comboBox, ControlLevels controlLevel, bool onlyReturn)
        {
            if (controlLevel == ControlLevels.Optional)
            {
                normalControl(comboBox);
                return true;
            }

            if (!comboBox.IsEditable)
            {
                if (comboBox.SelectedIndex != -1)
                {
                    normalControl(comboBox);
                    return true;
                }
                else
                {
                    if (!onlyReturn)
                        errorControl(comboBox, "موردی انتخاب نشده است");
                    return false;
                }
            }
            else
            {
                var textBox = (comboBox.Template.FindName("PART_EditableTextBox", comboBox) as TextBox);

                if (!string.IsNullOrEmpty(textBox.Text) && !string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text != ComboBoxData.otherFa && textBox.Text.Trim().Length > 5)
                {
                    normalControl(comboBox);
                    return true;
                }
                else if (string.IsNullOrEmpty(textBox.Text) || string.IsNullOrWhiteSpace(textBox.Text))
                {
                    if (!onlyReturn)
                        errorControl(comboBox, "فیلد نباید خالی باشد");

                    return false;
                }
                else if (textBox.Text == ComboBoxData.otherFa)
                {
                    normalControl(comboBox);
                    return true;
                }
                else if (textBox.Text.Trim().Length <= 5)
                {
                    if (!onlyReturn)
                        errorControl(comboBox, "بیشتر از 5 حرف میبایست باشد");

                    return false;
                }
                else
                    throw new System.Exception("unexpected error, value is >\n\t" + textBox.Text);
            }
        }
        private bool validateTextBox(TextBox textBox, ControlLevels controlLevel, bool onlyReturn)
        {
            if (controlLevel == ControlLevels.Optional)
            {
                normalControl(textBox);
                return true;
            }

            if (!string.IsNullOrEmpty(textBox.Text) && !string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text.Length > 5)
            {
                normalControl(textBox);
                return true;
            }
            else if (string.IsNullOrEmpty(textBox.Text) || string.IsNullOrWhiteSpace(textBox.Text))
            {
                if (!onlyReturn)
                    errorControl(textBox, "فیلد نباید خالی باشد");

                return false;
            }
            else
            {
                if (!onlyReturn)
                    errorControl(textBox, "حروف بیشتر از 5 حرف میبایست باشد");

                return false;
            }

        }
        #endregion

        #region Functions

        internal void resetControls()
        {
            #region variables
            foreach (CreateDocumentControlModel controlModel in comboBoxcontrolModels)
            {
                controlModel.Validate = false;
            }
            foreach (CreateDocumentControlModel controlModel in textBoxControlModels)
            {
                controlModel.Validate = false;
            }

            UniversityFa = "";
            UniversityEn = "";

            DepartmentFa = "";
            DepartmentEn = "";

            GroupFa = "";
            GroupEn = "";

            FieldOfStudyFa = "";
            FieldOfStudyEn = "";

            AreaOfStudyFa = "";
            AreaOfStudyEn = "";

            AcademicDegreeFa = "";
            AcademicDegreeEn = "";
            #endregion

            #region controls
            Dispatcher.Invoke(() =>
            {
                btnForward.IsEnabled = false;

                comboUniversity.SelectedIndex = -1;
                comboDepartment.ItemsSource = null;
                comboGroup.ItemsSource = null;
                comboAcademicDegree.ItemsSource = AcademicDegree_Fa2;

                foreach (CreateDocumentControlModel controlModel in comboBoxcontrolModels)
                {
                    ComboBox comboBox = (ComboBox)controlModel.Control;

                    resetComboBox(comboBox);
                    comboBoxcontrolModels.Where(a => a.Control == comboBox).FirstOrDefault().Validate = false;
                }

                foreach (CreateDocumentControlModel controlModel in textBoxControlModels)
                {
                    TextBox textBox = (TextBox)controlModel.Control;

                    if (textBox == txtBoxCustomDepartmentEn)
                    {
                        changeVisibilityEnglishControls(txtBoxCustomDepartmentEn, gridDepartment, Visibility.Collapsed);
                    }
                    else if (textBox == txtBoxCustomGroupEn)
                    {
                        changeVisibilityEnglishControls(txtBoxCustomGroupEn, gridGroup, Visibility.Collapsed);
                    }

                    resetTextBoxControl(textBox);
                }
            });
            #endregion
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


        private void resetTextBoxControl(TextBox control)
        {
            control.Text = "";
            normalControl(control);
        }
        private void resetComboBox(ComboBox control)
        {
            control.IsEditable = false;
            control.IsTextSearchEnabled = true;
            control.Text = "";
            control.SelectedIndex = -1;
            normalControl(control);
        }

        private void errorControl(Control control, string hintText)
        {
            HintAssist.SetHelperText(control, hintText);
            control.Foreground = System.Windows.Media.Brushes.Red;

            Thickness margin = new Thickness(0, 0, 0, 20);
            if (control == comboUniversity)
                gridUniversity.Margin = margin;
            else if (control == comboDepartment || control == txtBoxCustomDepartmentEn)
                gridDepartment.Margin = margin;
            else if (control == comboGroup || control == txtBoxCustomGroupEn)
                gridGroup.Margin = margin;
            else if (control == comboAcademicDegree)
                gridAcademicDegree.Margin = margin;
            else if (control == txtBoxFieldOfStudy || control == txtBoxFieldOfStudyEn)
                gridFieldOfStudy.Margin = margin;
            else if (control == txtBoxAreaOfStudy || control == txtBoxAreaOfStudyEn)
                gridAreaOfStudy.Margin = margin;
        }
        private void normalControl(Control control)
        {
            HintAssist.SetHelperText(control, "");
            control.Foreground = System.Windows.Media.Brushes.Black;

            Thickness margin = new Thickness(0);
            if (control == comboUniversity)
                gridUniversity.Margin = margin;
            else if (control == comboDepartment && txtBoxCustomDepartmentEn.Foreground != System.Windows.Media.Brushes.Red)
                gridDepartment.Margin = margin;
            else if (control == txtBoxCustomDepartmentEn && comboDepartment.Foreground != System.Windows.Media.Brushes.Red)
                gridDepartment.Margin = margin;
            else if (control == comboGroup && txtBoxCustomGroupEn.Foreground != System.Windows.Media.Brushes.Red)
                gridGroup.Margin = margin;
            else if (control == txtBoxCustomGroupEn && comboGroup.Foreground != System.Windows.Media.Brushes.Red)
                gridGroup.Margin = margin;
            else if (control == comboAcademicDegree)
                gridAcademicDegree.Margin = margin;
            else if (control == txtBoxFieldOfStudy && txtBoxFieldOfStudyEn.Foreground != System.Windows.Media.Brushes.Red)
                gridFieldOfStudy.Margin = margin;
            else if (control == txtBoxFieldOfStudyEn && txtBoxFieldOfStudy.Foreground != System.Windows.Media.Brushes.Red)
                gridFieldOfStudy.Margin = margin;
            else if (control == txtBoxAreaOfStudy && txtBoxAreaOfStudyEn.Foreground != System.Windows.Media.Brushes.Red)
                gridAreaOfStudy.Margin = margin;
            else if (control == txtBoxAreaOfStudyEn && txtBoxAreaOfStudy.Foreground != System.Windows.Media.Brushes.Red)
                gridAreaOfStudy.Margin = margin;
        }

        internal void initializeVariables(DocumentTypes documentType)
        {
            this.documentType = documentType;

            string previousSelectedAcademicDegree = "";

            if (comboAcademicDegree.SelectedIndex != -1)
            {
                previousSelectedAcademicDegree = comboAcademicDegree.SelectedItem as string;
            }
            comboAcademicDegree.SelectedIndex = -1;


            if (documentType == DocumentTypes.Project)
            {
                comboAcademicDegree.ItemsSource = ComboBoxDataAcademicDegree.AcademicDegree_Project_Fa;
            }
            else if (documentType == DocumentTypes.Thesis)
            {
                comboAcademicDegree.ItemsSource = ComboBoxDataAcademicDegree.AcademicDegree_Thesis_Fa;
            }
            else if (documentType == DocumentTypes.Dissertation)
            {
                comboAcademicDegree.ItemsSource = ComboBoxDataAcademicDegree.AcademicDegree_Dissertation_Fa;
            }
            else
            {
                comboAcademicDegree.ItemsSource = AcademicDegree_Fa2;
            }

            if (comboAcademicDegree.Items.Contains(previousSelectedAcademicDegree))
            {
                comboAcademicDegree.SelectedItem = previousSelectedAcademicDegree;
            }

            comboBoxcontrolModels.Where(a => a.Control == comboAcademicDegree).FirstOrDefault().Validate = validateComboBox(comboAcademicDegree, ControlLevels.Essential, true);

            validateControls();
        }
        #endregion
    }
}