using System;
using System.Drawing.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using Application = Microsoft.Office.Interop.Word.Application;
using Style = Microsoft.Office.Interop.Word.Style;
using System.Linq;
using System.Collections.Generic;

namespace ShivaNegar.Forms.FormatSettings
{
    /// <summary>
    /// Interaction logic for ChangeContentsControl.xaml
    /// </summary>
    public partial class FormatSettingsControl : UserControl, Interfaces.IStatusFormRequest
    {
        Document doc;
        public Action CloseFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }
        private Application wordApp;

        private List<string> allFonts;

        private List<WordStyleItem> predefinedStyles = new List<WordStyleItem>();

        DedicatedFunctions.AccessType accessType;
        #region constructor
        internal FormatSettingsControl(Document doc, DedicatedFunctions.AccessType accessType)
        {
            InitializeComponent();
            LoadInstalledFonts();


            predefinedStyles = new List<WordStyleItem>
            {
                new WordStyleItem { DisplayName = "عنوان سطح 1", StyleKey = "Heading 1" },
                new WordStyleItem { DisplayName = "عنوان سطح 2", StyleKey = "Heading 2" },
                new WordStyleItem { DisplayName = "عنوان سطح 3", StyleKey = "Heading 3" },
                new WordStyleItem { DisplayName = "عنوان سطح 4", StyleKey = "Heading 4" },
                new WordStyleItem { DisplayName = "عنوان سطح 5", StyleKey = "Heading 5" },
                new WordStyleItem { DisplayName = "متن سند", StyleKey = "Normal" },
                new WordStyleItem { DisplayName = "عنوان شکل", StyleKey = "Caption" },
                new WordStyleItem { DisplayName = "عنوان جدول", StyleKey = "Table Heading" },
            };

            comboHeading.ItemsSource = predefinedStyles;
            comboHeading.SelectionChanged += ComboHeading_SelectionChanged;

            this.accessType = accessType;
            this.doc = doc;


            btnCloseApp.Click += BtnCloseApp_Click;
            //btnMaximize.Click += BtnMaximize_Click;
            //btnMinimize.Click += BtnMinimize_Click;
            btnApply.Click += BtnApply_Click;



            string bibliographyStyle = DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_BibliographyStyle.ToString());

            //if(tableOfContentsStyle == SettingNames.NotExist)
            //	DedicatedFunctions.addVariable(doc , VariableOtherNames._variable_other_TableOfContentsStyle.ToString() , "0");
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
        private void BtnApply_Click(object sender, RoutedEventArgs e)
        {


            string fontName = comboFontFamily.SelectedItem as string
                  ?? comboFontFamily.Text
                  ?? "B Lotus";
            string fontNameBi = comboFontFamilyBi.SelectedItem as string
                  ?? comboFontFamilyBi.Text
                  ?? "Times New Roman";

            if (!float.TryParse(txtFontSize.Text, out float fontSize))
            {
                MessageBox.Show("اندازه فونت فارسی معتبر نیست.");
                return;
            }
            if (!float.TryParse(txtFontSizeBi.Text, out float fontSizeBi))
            {
                MessageBox.Show("اندازه فونت انگلیسی معتبر نیست.");
                return;
            }

            var colorItem = comboFontColor.SelectedItem as ComboBoxItem;
            System.Drawing.Color color = System.Drawing.Color.Black;
            if (colorItem?.Tag != null)
            {
                string[] rgb = colorItem.Tag.ToString().Split(',');
                color = System.Drawing.Color.FromArgb(int.Parse(rgb[0]), int.Parse(rgb[1]), int.Parse(rgb[2]));
            }

            try
            {
                var wordApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application") as Application;
                if (wordApp == null)
                {
                    MessageBox.Show("برنامه Word باز نیست.");
                    return;
                }

                var doc = wordApp.ActiveDocument;
                //Style style = doc.Styles[headingName];


                var selectedItem = comboHeading.SelectedItem as WordStyleItem;
                string styleName = null;

                if (selectedItem != null)
                {
                    // مقدار واقعی استایل (مثل "Heading 1")
                    styleName = selectedItem.StyleKey;
                }
                else if (!string.IsNullOrWhiteSpace(comboHeading.Text))
                {
                    // اگر کاربر خودش تایپ کرده باشه
                    styleName = comboHeading.Text;
                }

                // حالا مطمئن می‌شیم که استایل واقعاً وجود داره
                if (!string.IsNullOrEmpty(styleName))
                {
                    Style style = doc.Styles.Cast<Style>().FirstOrDefault(s => s.NameLocal == styleName);
                    if (style == null)
                    {
                        MessageBox.Show($"❌ استایل '{styleName}' در این سند وجود ندارد!");
                        return;
                    }


                    // 🔹 فونت انگلیسی
                    style.Font.Name = fontNameBi;
                    style.Font.Size = fontSizeBi;

                    // 🔹 فونت فارسی
                    style.Font.NameBi = fontName;
                    style.Font.SizeBi = fontSize;

                    

                    // 🔹 رنگ
                    style.Font.Color = (WdColor)(color.R + 0x100 * color.G + 0x10000 * color.B);

                    // 🔹 فاصله‌ها
                    if (float.TryParse(txtLineSpacing.Text, out float lineSpacing))
                        style.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple;

                    style.ParagraphFormat.LineSpacing = wordApp.LinesToPoints(lineSpacing);

                    if (float.TryParse(txtSpaceBefore.Text, out float spaceBefore))
                        style.ParagraphFormat.SpaceBefore = spaceBefore;

                    if (float.TryParse(txtSpaceAfter.Text, out float spaceAfter))
                        style.ParagraphFormat.SpaceAfter = spaceAfter;

                    CloseFormRequest?.Invoke();
                }
                else
                {
                    MessageBox.Show("❌ لطفاً یک استایل معتبر انتخاب کنید!");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در تغییر استایل: " + ex.Message);
            }
        }

        private void LoadInstalledFonts()
        {
            try
            {
                InstalledFontCollection fontsCollection = new InstalledFontCollection();
                allFonts = fontsCollection.Families
                    .Select(f => f.Name)
                    .OrderBy(n => n)
                    .ToList();

                comboFontFamily.ItemsSource = allFonts;
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در بارگذاری فونت‌ها: " + ex.Message);
            }
        }

        private void comboFontFamily_GotFocus(object sender, RoutedEventArgs e)
        {
            // وقتی فوکوس شد، لیست را باز کن
            if (comboFontFamily.ItemsSource != null && comboFontFamily.Items.Count > 0)
            {
                comboFontFamily.IsDropDownOpen = true;
            }
            else
            {
                // اگر آیتم‌ها لود نشده‌اند، ابتدا لود کن سپس باز کن
                comboFontFamily.ItemsSource = allFonts;
                comboFontFamily.IsDropDownOpen = true;
            }
        }
        private void ComboFontFamily_TextChanged(object sender, TextChangedEventArgs e)
        {
            // مقدار تایپ‌شده توسط کاربر
            string filter = comboFontFamily.Text?.ToLower() ?? "";

            if (string.IsNullOrWhiteSpace(filter))
            {
                comboFontFamily.ItemsSource = allFonts;
            }
            else
            {
                comboFontFamily.ItemsSource = allFonts
                    .Where(f => f.ToLower().Contains(filter))
                    .ToList();
            }

            comboFontFamily.IsDropDownOpen = true; // باز نگه داشتن لیست
            comboFontFamily.Items.Refresh();
        }
        private void comboFontFamilyBi_GotFocus(object sender, RoutedEventArgs e)
        {
            // وقتی فوکوس شد، لیست را باز کن
            if (comboFontFamilyBi.ItemsSource != null && comboFontFamilyBi.Items.Count > 0)
            {
                comboFontFamilyBi.IsDropDownOpen = true;
            }
            else
            {
                // اگر آیتم‌ها لود نشده‌اند، ابتدا لود کن سپس باز کن
                comboFontFamilyBi.ItemsSource = allFonts;
                comboFontFamilyBi.IsDropDownOpen = true;
            }
        }
        private void ComboFontFamilyBi_TextChanged(object sender, TextChangedEventArgs e)
        {
            // مقدار تایپ‌شده توسط کاربر
            string filter = comboFontFamilyBi.Text?.ToLower() ?? "";

            if (string.IsNullOrWhiteSpace(filter))
            {
                comboFontFamilyBi.ItemsSource = allFonts;
            }
            else
            {
                comboFontFamilyBi.ItemsSource = allFonts
                    .Where(f => f.ToLower().Contains(filter))
                    .ToList();
            }

            comboFontFamilyBi.IsDropDownOpen = true; // باز نگه داشتن لیست
            comboFontFamilyBi.Items.Refresh();
        }
        
        private void ComboHeading_TextChanged(object sender, TextChangedEventArgs e)
        {
            string filter = comboHeading.Text?.Trim().ToLower() ?? "";
            if (string.IsNullOrEmpty(filter))
            {
                comboHeading.ItemsSource = predefinedStyles;
            }
            else
            {
                var filtered = predefinedStyles
                    .Where(s => s.DisplayName.ToLower().Contains(filter) ||
                                s.StyleKey.ToLower().Contains(filter))
                    .OrderBy(s => s.DisplayName.StartsWith(filter) ? 0 : 1)
                    .ThenBy(s => s.DisplayName)
                    .ToList();

                comboHeading.ItemsSource = filtered;
            }

            comboHeading.IsDropDownOpen = true;
            comboHeading.Items.Refresh();
        }

        private void ComboHeading_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                var selectedItem = comboHeading.SelectedItem as WordStyleItem;
                if (selectedItem == null)
                    return;

                string styleKey = selectedItem.StyleKey;

                if (doc == null)
                {
                    MessageBox.Show("سند Word فعال نیست ❌");
                    return;
                }

                // پیدا کردن استایل در سند
                Style style = doc.Styles.Cast<Style>().FirstOrDefault(s => s.NameLocal == styleKey);
                if (style == null)
                {
                    MessageBox.Show("استایل یافت نشد ❌");
                    return;
                }

                // 🔹 فونت فارسی و انگلیسی
                string fontFa = style.Font.NameBi;
                string fontEn = style.Font.Name;

                // 🔹 اندازه فونت‌ها (اگر مقدار 0 بود از دیگری استفاده می‌کنیم)
                float fontSizeFa = (float)(style.Font.SizeBi > 0 ? style.Font.SizeBi : style.Font.Size);
                float fontSizeEn = (float)(style.Font.Size > 0 ? style.Font.Size : style.Font.SizeBi);

                // 🔹 رنگ متن
                WdColor color = style.Font.Color;

                // 🔹 فاصله بین خطوط
                float lineSpacing = (float)Math.Round(style.ParagraphFormat.LineSpacing / 12, 2);
                if (lineSpacing < 0) lineSpacing = Math.Abs(lineSpacing);

                // 🔹 فاصله قبل و بعد از پاراگراف (در Word بعضی استایل‌ها فاصله منفی دارن)
                float spaceBefore = Math.Max(0, (float)Math.Round(style.ParagraphFormat.SpaceBefore, 2));
                float spaceAfter = Math.Max(0, (float)Math.Round(style.ParagraphFormat.SpaceAfter, 2));

                // 🔹 پر کردن فیلدها
                comboFontFamily.Text = !string.IsNullOrWhiteSpace(fontFa) ? fontFa : fontEn;
                txtFontSize.Text = fontSizeFa.ToString();

                comboFontFamilyBi.Text = !string.IsNullOrWhiteSpace(fontEn) ? fontEn : fontFa;
                txtFontSizeBi.Text = fontSizeEn.ToString();

                txtLineSpacing.Text = lineSpacing.ToString();
                txtSpaceBefore.Text = spaceBefore.ToString();
                txtSpaceAfter.Text = spaceAfter.ToString();

                // 🔹 رنگ
                if (color == WdColor.wdColorAutomatic)
                    comboFontColor.SelectedIndex = 0;
                else if (color == WdColor.wdColorBlack)
                    comboFontColor.SelectedIndex = 1;
                else if (color == WdColor.wdColorGray50)
                    comboFontColor.SelectedIndex = 2;
                else
                    comboFontColor.Text = color.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show("خطا در دریافت تنظیمات استایل: " + ex.Message);
            }
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
        public class WordStyleItem
        {
            public string DisplayName { get; set; } // نام فارسی برای نمایش
            public string StyleKey { get; set; }     // نام واقعی استایل در ورد

            public override string ToString() => DisplayName;
        }

    }
}
