using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants.ComboBoxData;

namespace ShivaNegar.Forms.BesmellahPage
{
    /// <summary>
    /// Interaction logic for ChangeContentsControl.xaml
    /// </summary>
    public partial class BesmellahPageControl : UserControl, Interfaces.IStatusFormRequest
    {
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

        public BesmellahPageControl()
        {
            InitializeComponent();

            btnCloseApp.Click += BtnCloseApp_Click;
            //btnMaximize.Click += BtnMaximize_Click;
            //btnMinimize.Click += BtnMinimize_Click;
            btnConfirmChanges.Click += BtnConfirmChanges_Click;

            initializeControlNameOfAllah();
        }

        private void BtnConfirmChanges_Click(object sender, System.Windows.RoutedEventArgs es)
        {
            string fontName = Constants.FontNames.fontBesmellah1;
            if (comboBesmellahFontType.SelectedIndex == 0)
                fontName = Constants.FontNames.fontBesmellah1;
            else if (comboBesmellahFontType.SelectedIndex == 1)
                fontName = Constants.FontNames.fontBesmellah2;
            else if (comboBesmellahFontType.SelectedIndex == 2)
                fontName = Constants.FontNames.fontBesmellah3;

            loadingControl.Tag = "در حال ایجاد سند جدید";
            loadingControl.IsEnabled = true;

            try
            {
                // ایجاد سند جدید ورد
                Globals.ThisAddIn.DisableEvents = true;
                Document newDocument = Globals.ThisAddIn.Application.Documents.Add();
                Globals.ThisAddIn.DisableEvents = false;

                // استفاده از نام کامل برای Window
                Microsoft.Office.Interop.Word.Window window = newDocument.ActiveWindow;
                Selection selection = window.Selection;

                // مخفی کردن پنجره موقتاً برای سرعت بیشتر
                window.Visible = false;

                // اضافه کردن بسم الله به سند جدید
                DedicatedFunctions.BesmellahToImage(selection, 200, fontName, lblNameOfAllah.Text);

                // تنظیمات نهایی سند
                selection.HomeKey();
                selection.HomeKey();
                selection.Font.Size = 12;
                selection.Font.SizeBi = 12;
                selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                selection.ParagraphFormat.Space1();
                selection.InsertParagraph();
                selection.InsertParagraph();
                selection.InsertParagraph();

                // نمایش سند
                window.Visible = true;

                // بستن فرم فعلی
                CloseFormRequest?.Invoke();
            }
            catch (Exception ex)
            {
                DedicatedFunctions.ShowErrorMessage("خطا در ایجاد سند جدید: " + ex.Message);
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

        private void initializeControlNameOfAllah()
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

            lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
            lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
            NameOfAllahFontType = 1;
            comboBesmellahFontType.SelectedIndex = 0;
        }

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
