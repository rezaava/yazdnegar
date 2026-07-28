using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MaterialDesignThemes.Wpf;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.CreateDocument.Models;
using static ShivaNegar.Forms.ShivaNegarManager.CreateDocument.Models.CreateDocumentControlModel;

namespace ShivaNegar.Forms.ShivaNegarManager.CreateDocument
{
    public partial class CreateDocumentSlide4 : UserControl
    {
        private List<CreateDocumentControlModel> controlModels;

        //Properties
        public string NameOFCourseFa { get; private set; }
        public string TitleFa { get; private set; }
        public string TitleEn { get; private set; }
        public string AuthorFa { get; private set; }
        public string AuthorEn { get; private set; }
        public string AdvisorFa { get; private set; }
        public string AdvisorEn { get; private set; }
        public string SupervisorFa { get; private set; }
        public string SupervisorEn { get; private set; }
        public string DefenseDateFa { get; private set; }
        public string DefenseDateEn { get; private set; }

        public CreateDocumentSlide4()
        {
            InitializeComponent();

            btnForward.IsEnabled = false;

            controlModels = new List<CreateDocumentControlModel>()
            {
                new CreateDocumentControlModel(txtBoxNameOfCourseFa,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxTitleFa,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxTitleEn,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxAuthorFa,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxAuthorEn,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxSupervisorFa,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxSupervisorEn,CreateDocumentControlModel.ControlLevels.Essential),
                new CreateDocumentControlModel(txtBoxAdvisorFa,CreateDocumentControlModel.ControlLevels.Optional),
                new CreateDocumentControlModel(txtBoxAdvisorEn,CreateDocumentControlModel.ControlLevels.Optional),
                new CreateDocumentControlModel(txtBoxDefenseDateFa,CreateDocumentControlModel.ControlLevels.Optional),
                new CreateDocumentControlModel(txtBoxDefenseDateEn,CreateDocumentControlModel.ControlLevels.Optional),
            };

            txtBoxSupervisorFa.PreviewKeyDown += TextBoxMultiLine_PreviewKeyDown;
            txtBoxSupervisorEn.PreviewKeyDown += TextBoxMultiLine_PreviewKeyDown;
            txtBoxAdvisorFa.PreviewKeyDown += TextBoxMultiLine_PreviewKeyDown;
            txtBoxAdvisorEn.PreviewKeyDown += TextBoxMultiLine_PreviewKeyDown;

            foreach (CreateDocumentControlModel cdcm in controlModels)
            {
                ((TextBox)cdcm.Control).TextChanged += TextBox_TextChanged;
                ((TextBox)cdcm.Control).GotFocus += TextBox_GotFocus;
                ((TextBox)cdcm.Control).LostFocus += TextBox_LostFocus;
            }
        }

        #region Events

        #region TextBox
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
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            for (int i = 0; i < controlModels.Count; i++)
            {
                if (textBox == (TextBox)controlModels[i].Control)
                {
                    controlModels[i].Validate = validateTextBox(textBox, controlModels[i].ControlLevel, false);
                    validateControls();
                    return;
                }
            }
        }
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox textBox = (TextBox)sender;

            if (textBox == txtBoxAdvisorFa)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {

                    controlModels.Where(a => a.Control == txtBoxAdvisorEn).FirstOrDefault().ControlLevel = ControlLevels.Essential;
                }
                else
                {
                    controlModels.Where(a => a.Control == txtBoxAdvisorEn).FirstOrDefault().ControlLevel = ControlLevels.Optional;
                }

                controlModels.Where(a => a.Control == txtBoxAdvisorEn).FirstOrDefault().Validate = validateTextBox(txtBoxAdvisorEn, controlModels.Where(a => a.Control == txtBoxAdvisorEn).FirstOrDefault().ControlLevel, false);
                validateControls();
            }
            else if (textBox == txtBoxAdvisorEn)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    controlModels.Where(a => a.Control == txtBoxAdvisorFa).FirstOrDefault().ControlLevel = ControlLevels.Essential;
                }
                else
                {
                    controlModels.Where(a => a.Control == txtBoxAdvisorFa).FirstOrDefault().ControlLevel = ControlLevels.Optional;
                }

                controlModels.Where(a => a.Control == txtBoxAdvisorFa).FirstOrDefault().Validate = validateTextBox(txtBoxAdvisorFa, controlModels.Where(a => a.Control == txtBoxAdvisorFa).FirstOrDefault().ControlLevel, false);
                validateControls();
            }

            if (textBox == txtBoxDefenseDateFa)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {

                    controlModels.Where(a => a.Control == txtBoxDefenseDateEn).FirstOrDefault().ControlLevel = ControlLevels.Essential;
                }
                else
                {
                    controlModels.Where(a => a.Control == txtBoxDefenseDateEn).FirstOrDefault().ControlLevel = ControlLevels.Optional;
                }

                controlModels.Where(a => a.Control == txtBoxDefenseDateEn).FirstOrDefault().Validate = validateTextBox(txtBoxDefenseDateEn, controlModels.Where(a => a.Control == txtBoxDefenseDateEn).FirstOrDefault().ControlLevel, false);
                validateControls();
            }
            else if (textBox == txtBoxDefenseDateEn)
            {
                if (!string.IsNullOrEmpty(textBox.Text))
                {
                    controlModels.Where(a => a.Control == txtBoxDefenseDateFa).FirstOrDefault().ControlLevel = ControlLevels.Essential;
                }
                else
                {
                    controlModels.Where(a => a.Control == txtBoxDefenseDateFa).FirstOrDefault().ControlLevel = ControlLevels.Optional;
                }

                controlModels.Where(a => a.Control == txtBoxDefenseDateFa).FirstOrDefault().Validate = validateTextBox(txtBoxDefenseDateFa, controlModels.Where(a => a.Control == txtBoxDefenseDateFa).FirstOrDefault().ControlLevel, false);
                validateControls();
            }

            for (int i = 0; i < controlModels.Count; i++)
            {
                if (textBox == (TextBox)controlModels[i].Control)
                {
                    controlModels[i].Validate = validateTextBox(textBox, controlModels[i].ControlLevel, true);

                    if (controlModels[i].Validate)
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

        #endregion

        #region Validators
        private bool validateControls()
        {
            bool isValid = true;
            foreach (CreateDocumentControlModel controlModel in controlModels)
            {
                if (!controlModel.Validate && ((UIElement)controlModel.Control).Visibility == Visibility.Visible && controlModel.ControlLevel != CreateDocumentControlModel.ControlLevels.Optional)
                {
                    isValid = false;
                    break;
                }
            }

            if (isValid)
            {
                NameOFCourseFa = txtBoxNameOfCourseFa.Text;

                TitleFa = txtBoxTitleFa.Text;
                TitleEn = txtBoxTitleEn.Text;

                AuthorFa = txtBoxAuthorFa.Text;
                AuthorEn = txtBoxAuthorEn.Text;

                SupervisorFa = txtBoxSupervisorFa.Text;
                SupervisorEn = txtBoxSupervisorEn.Text;

                AdvisorFa = txtBoxAdvisorFa.Text;
                AdvisorEn = txtBoxAdvisorEn.Text;

                DefenseDateFa = txtBoxDefenseDateFa.Text;
                DefenseDateEn = txtBoxDefenseDateEn.Text;

                btnForward.IsEnabled = true;
                return true;
            }
            else
            {
                NameOFCourseFa = "";

                TitleFa = "";
                TitleEn = "";

                AuthorFa = "";
                AuthorEn = "";

                SupervisorFa = "";
                SupervisorEn = "";

                AdvisorFa = "";
                AdvisorEn = "";

                DefenseDateFa = "";
                DefenseDateEn = "";

                btnForward.IsEnabled = false;
                return false;
            }
        }
        private bool validateTextBox(TextBox textBox, ControlLevels controlLevel, bool onlyReturn)
        {
            if (controlLevel == ControlLevels.Optional)
            {
                normalControl(textBox);
                return true;
            }

            if (!string.IsNullOrEmpty(textBox.Text) && !string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text.Length > 4)
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
                    errorControl(textBox, "حروف بیشتر از 4 حرف میبایست باشد");
                return false;
            }

        }
        #endregion

        #region Functions
        internal void resetControls()
        {
            #region variables

            foreach (CreateDocumentControlModel controlModel in controlModels)
            {
                controlModel.Validate = false;
            }

            NameOFCourseFa = "";

            TitleFa = "";
            TitleEn = "";

            AuthorFa = "";
            AuthorEn = "";

            SupervisorFa = "";
            SupervisorEn = "";

            AdvisorFa = "";
            AdvisorEn = "";

            DefenseDateFa = "";
            DefenseDateEn = "";
            #endregion


            #region controls
            Dispatcher.Invoke(() =>
            {

                btnForward.IsEnabled = false;
                txtBoxNameOfCourseFa.Visibility = Visibility.Collapsed;

                foreach (CreateDocumentControlModel controlModel in controlModels)
                {
                    TextBox textBox = (TextBox)controlModel.Control;

                    textBox.Text = "";
                    normalControl(textBox);
                }
            });
            #endregion
        }

        private void errorControl(Control control, string hintText)
        {
            HintAssist.SetHelperText(control, hintText);
            control.Foreground = System.Windows.Media.Brushes.Red;

            Thickness margin = new Thickness(0, 0, 0, 20);
            if (control == txtBoxNameOfCourseFa)
                gridNameOfCourse.Margin = margin;
            else if (control == txtBoxTitleFa || control == txtBoxTitleEn)
                gridTitle.Margin = margin;
            else if (control == txtBoxAuthorFa || control == txtBoxAuthorEn)
                gridAuthor.Margin = margin;
            else if (control == txtBoxSupervisorFa || control == txtBoxSupervisorEn)
                gridSupervisor.Margin = margin;
            else if (control == txtBoxAdvisorFa || control == txtBoxAdvisorEn)
                gridAdvisor.Margin = margin;
            else if (control == txtBoxDefenseDateFa || control == txtBoxDefenseDateEn)
                gridDefenseDate.Margin = margin;
        }
        private void normalControl(Control control)
        {
            HintAssist.SetHelperText(control, "");
            control.Foreground = System.Windows.Media.Brushes.Black;

            Thickness margin = new Thickness(0);
            if (control == txtBoxNameOfCourseFa)
                gridNameOfCourse.Margin = margin;
            else if (control == txtBoxTitleFa && txtBoxTitleEn.Foreground != System.Windows.Media.Brushes.Red)
                gridTitle.Margin = margin;
            else if (control == txtBoxTitleEn && txtBoxTitleFa.Foreground != System.Windows.Media.Brushes.Red)
                gridTitle.Margin = margin;
            else if (control == txtBoxAuthorFa && txtBoxAuthorEn.Foreground != System.Windows.Media.Brushes.Red)
                gridAuthor.Margin = margin;
            else if (control == txtBoxAuthorEn && txtBoxAuthorFa.Foreground != System.Windows.Media.Brushes.Red)
                gridAuthor.Margin = margin;
            else if (control == txtBoxSupervisorFa && txtBoxSupervisorEn.Foreground != System.Windows.Media.Brushes.Red)
                gridSupervisor.Margin = margin;
            else if (control == txtBoxSupervisorEn && txtBoxSupervisorFa.Foreground != System.Windows.Media.Brushes.Red)
                gridSupervisor.Margin = margin;
            else if (control == txtBoxAdvisorFa && txtBoxAdvisorEn.Foreground != System.Windows.Media.Brushes.Red)
                gridAdvisor.Margin = margin;
            else if (control == txtBoxAdvisorEn && txtBoxAdvisorFa.Foreground != System.Windows.Media.Brushes.Red)
                gridAdvisor.Margin = margin;
            else if (control == txtBoxDefenseDateFa && txtBoxDefenseDateEn.Foreground != System.Windows.Media.Brushes.Red)
                gridDefenseDate.Margin = margin;
            else if (control == txtBoxDefenseDateEn && txtBoxDefenseDateFa.Foreground != System.Windows.Media.Brushes.Red)
                gridDefenseDate.Margin = margin;
        }

        public void initializeVariables(DocumentTypes documentType)
        {
            if (documentType == DocumentTypes.SchoolResearch)
                txtBoxNameOfCourseFa.Visibility = Visibility.Visible;
            else
                txtBoxNameOfCourseFa.Visibility = Visibility.Collapsed;

            //string initialTitle = "مشخصات ";
            //if(documentType == DocumentTypes.Project)
            //	lblTitle.Text = initialTitle + DocumentTypeValues.DocumentType_ProjectFa;
            //else if(documentType == DocumentTypes.Thesis)
            //	lblTitle.Text = initialTitle + DocumentTypeValues.DocumentType_ThesisFa;
            //else if(documentType == DocumentTypes.Dissertation)
            //	lblTitle.Text = initialTitle + DocumentTypeValues.DocumentType_DissertationFa;
            //else if(documentType == DocumentTypes.SchoolResearch)
            //	lblTitle.Text = initialTitle + "تحقیق درسی";
        }
        #endregion

    }
}