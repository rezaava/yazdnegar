using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;

namespace ShivaNegar.Forms.CitationSettings
{
    /// <summary>
    /// Interaction logic for ChangeContentsControl.xaml
    /// </summary>
    public partial class CitationSettingsControl : UserControl, Interfaces.IStatusFormRequest
    {
        Document doc;
        public Action CloseFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }

        DedicatedFunctions.AccessType accessType;
        #region constructor
        internal CitationSettingsControl(Document doc, DedicatedFunctions.AccessType accessType)
        {
            InitializeComponent();

            this.accessType = accessType;
            this.doc = doc;

            btnCloseApp.Click += BtnCloseApp_Click;
            //btnMaximize.Click += BtnMaximize_Click;
            //btnMinimize.Click += BtnMinimize_Click;

            string bibliographyStyle = DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_BibliographyStyle.ToString());

            //if(tableOfContentsStyle == SettingNames.NotExist)
            //	DedicatedFunctions.addVariable(doc , VariableOtherNames._variable_other_TableOfContentsStyle.ToString() , "0");

            int index = -1;
            if (bibliographyStyle == "APA")
                index = 0;
            else if (bibliographyStyle == "IEEE")
                index = 1;

            if (bibliographyStyle != SettingValues.NotExist && !string.IsNullOrEmpty(bibliographyStyle))
                comboCitationDisplayStyle.SelectedIndex = index;

            comboCitationDisplayStyle.SelectionChanged += ComboCitationDisplayStyle_SelectionChanged;
        }

        private void ComboCitationDisplayStyle_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ComboBoxItem cbi = (ComboBoxItem)comboCitationDisplayStyle.SelectedItem;
            string value = cbi.Content.ToString();


            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر سبک نمایش منابع به " + value);

            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableOptionIDs._variable_option_BibliographyStyle.ToString(), value);
            Globals.ThisAddIn.Application.Options.BibliographyStyle = value;
            doc.Bibliography.BibliographyStyle = value;

            DedicatedFunctions.updateBibliography(doc, doc.ActiveWindow.Selection, accessType);

            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
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
