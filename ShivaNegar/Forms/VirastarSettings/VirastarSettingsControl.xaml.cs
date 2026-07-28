using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Office.Interop.Word;

namespace ShivaNegar.Forms.VirastarSettings
{
    /// <summary>
    /// Interaction logic for ChangeContentsControl.xaml
    /// </summary>
    public partial class VirastarSettingsControl : UserControl, Interfaces.IStatusFormRequest
    {
        Document doc;
        public Action CloseFormRequest { get; set; }
        public Action NormalStateFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }

        #region constructor
        public VirastarSettingsControl(Document doc)
        {
            InitializeComponent();

            this.doc = doc;

            btnCloseApp.Click += BtnCloseApp_Click;
            //btnMaximize.Click += BtnMaximize_Click;
            //btnMinimize.Click += BtnMinimize_Click;

            tbtnIncludeHalfSpaceCorrectionFootnote.IsChecked = Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionFootnote;
            tbtnIncludeHalfSpaceCorrectionFootnote.Click += TbtnIncludeHalfSpaceCorrectionFootnote_Click;

            tbtnIncludeHalfSpaceCorrectionHeadersFooters.IsChecked = Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionHeadersFooters;
            tbtnIncludeHalfSpaceCorrectionHeadersFooters.Click += TbtnIncludeHalfSpaceCorrectionHeadersFooters_Click;

            tbtnIncludeHalfSpaceCorrectionSpecialFields.IsChecked = Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionSpecialFields;
            tbtnIncludeHalfSpaceCorrectionSpecialFields.Click += TbtnIncludeHalfSpaceCorrectionSpecialFields_Click;

            tbIncludeChangeDigitCharacters.IsChecked = Properties.Settings.Default.VirastarSettings_IncludeChangeDigitCharacters;
            tbIncludeChangeDigitCharacters.Click += TbIncludeChangeDigitCharacters_Click; ;

            comboCreatePoemType.SelectionChanged += ComboCreatePoemType_SelectionChanged;

            //store in variable
            //string poemType = DedicatedFunctions.getStaticVariableValue(doc , VariableOtherNames._variable_other_VirastarSettings_CreatePoemType.ToString());
            //
            //if(string.IsNullOrEmpty(poemType) || poemType == SettingNames.NotExist)
            //{
            //	comboCreatePoemType.SelectedIndex = 0;
            //}
            //else
            //{
            //	comboCreatePoemType.SelectedIndex = int.Parse(poemType);
            //}

            //store in addin
            int type = Properties.Settings.Default.VirastarSettings_CreatePoemType;
            comboCreatePoemType.SelectedIndex = type;
        }

        private void TbIncludeChangeDigitCharacters_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.VirastarSettings_IncludeChangeDigitCharacters = (bool)((System.Windows.Controls.Primitives.ToggleButton)sender).IsChecked;
            Properties.Settings.Default.Save();
        }

        private void ComboCreatePoemType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Properties.Settings.Default.VirastarSettings_CreatePoemType = comboCreatePoemType.SelectedIndex;
            Properties.Settings.Default.Save();
            //DedicatedFunctions.setORAddStaticVariableValue(doc , VariableOtherNames._variable_other_VirastarSettings_CreatePoemType.ToString() , comboCreatePoemType.SelectedIndex.ToString());
        }

        #endregion

        #region ToggleButtons
        private void TbtnIncludeHalfSpaceCorrectionFootnote_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionFootnote = (bool)((System.Windows.Controls.Primitives.ToggleButton)sender).IsChecked;
            Properties.Settings.Default.Save();
        }
        private void TbtnIncludeHalfSpaceCorrectionHeadersFooters_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionHeadersFooters = (bool)((System.Windows.Controls.Primitives.ToggleButton)sender).IsChecked;
            Properties.Settings.Default.Save();
        }

        private void TbtnIncludeHalfSpaceCorrectionSpecialFields_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionSpecialFields = (bool)((System.Windows.Controls.Primitives.ToggleButton)sender).IsChecked;
            Properties.Settings.Default.Save();
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

    }
}
