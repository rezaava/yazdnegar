using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Media;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;

namespace ShivaNegar.Forms.DocumentSettings
{
    /// <summary>
    /// Interaction logic for ChangeContentsControl.xaml
    /// </summary>
    public partial class DocumentSettingsControl : System.Windows.Controls.UserControl, Interfaces.IStatusFormRequest
    {
        Document doc;
        public Action CloseFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }

        #region constructor
        public DocumentSettingsControl(Document doc)
        {
            InitializeComponent();

            this.doc = doc;

            btnCloseApp.Click += BtnCloseApp_Click;
            //btnMaximize.Click += BtnMaximize_Click;
            //btnMinimize.Click += BtnMinimize_Click;

            string tableOfContentsStyle = DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_TableOfContentsStyle.ToString());

            //if(tableOfContentsStyle == SettingNames.NotExist)
            //	DedicatedFunctions.addVariable(doc , VariableOtherNames._variable_other_TableOfContentsStyle.ToString() , "0");

            if (tableOfContentsStyle != SettingValues.NotExist)
                comboTableOfContentIndentionStyle.SelectedIndex = int.Parse(tableOfContentsStyle);

            comboTableOfContentIndentionStyle.SelectionChanged += ComboTableOfContentIndentionStyle_SelectionChanged;

            btnChangeDefaultFontEnglish.Click += BtnChangeDefaultFontEnglish_Click;
            btnChangeDefaultFontPersian.Click += BtnChangeDefaultFontPersian_Click;

            tbtnAutoSaveOnClose.Click += TbtnAutoSaveOnClose_Click;
        }

        private void TbtnAutoSaveOnClose_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.AutoSaveOnClose = (bool)tbtnAutoSaveOnClose.IsChecked;
            Properties.Settings.Default.Save();
        }

        private void BtnChangeDefaultFontPersian_Click(object sender, RoutedEventArgs e)
        {
            string defaultPersianFont = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleNormal].Font.NameBi;

            FontDialog fontDialog = new FontDialog();
            fontDialog.ShowColor = false;
            fontDialog.ShowEffects = false;
            fontDialog.ShowHelp = false;
            fontDialog.Font = new System.Drawing.Font(defaultPersianFont, 12);
            fontDialog.ShowApply = false;
            DialogResult result = fontDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                LoadingForm loadingForm = new LoadingForm();
                System.Threading.Tasks.Task.Run(() => changeStyles(Globals.ThisAddIn.Application.ActiveDocument, loadingForm, true, fontDialog.Font.FontFamily.Name));
                loadingForm.ShowDialog();
            }
        }

        private void BtnChangeDefaultFontEnglish_Click(object sender, RoutedEventArgs e)
        {
            string defaultEnglishFont = Globals.ThisAddIn.Application.ActiveDocument.Styles[StyleNames.styleNormal].Font.Name;

            FontDialog fontDialog = new FontDialog();
            fontDialog.ShowColor = false;
            fontDialog.ShowEffects = false;
            fontDialog.ShowHelp = false;
            fontDialog.Font = new System.Drawing.Font(defaultEnglishFont, 12);
            fontDialog.ShowApply = false;
            DialogResult result = fontDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                LoadingForm loadingForm = new LoadingForm();
                System.Threading.Tasks.Task.Run(() => changeStyles(Globals.ThisAddIn.Application.ActiveDocument, loadingForm, false, fontDialog.Font.FontFamily.Name));
                loadingForm.ShowDialog();
            }
        }

        private void ComboTableOfContentIndentionStyle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تنظیم سبک فهرست مطالب");
            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableOptionIDs._variable_option_TableOfContentsStyle.ToString(), comboTableOfContentIndentionStyle.SelectedIndex.ToString());
            DedicatedFunctions.selectTableOfContentsStyle(doc, doc.ActiveWindow.Selection, comboTableOfContentIndentionStyle.SelectedIndex);
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }


        #endregion

        #region Functions

        internal void changeStyles(Document doc, LoadingForm frm, bool isPersian, string fontName)
        {
            string statusUndo;
            if (isPersian)
                statusUndo = "فارسی";
            else
                statusUndo = "انگلیسی";

            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر فونت پیشفرض " + statusUndo);

            if (isPersian)
                DedicatedFunctions.changeDefaultFont(doc, persianFont: fontName);
            else
                DedicatedFunctions.changeDefaultFont(doc, englishFont: fontName);

            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

            frm?.closeForm();
        }
        #endregion

        #region buttons

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

        public Action NormalStateFormRequest { get; set; }
    }
}
