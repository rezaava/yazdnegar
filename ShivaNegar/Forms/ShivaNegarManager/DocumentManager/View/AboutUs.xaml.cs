using System;
using System.Reflection;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View
{
    /// <summary>
    /// Interaction logic for Login.xaml
    /// </summary>
    public partial class AboutUs : System.Windows.Controls.UserControl
    {
        private bool _showBackButton = true;

        // سازنده پیش‌فرض (برای استفاده در جاهای دیگر - با دکمه)
        public AboutUs() : this(true)
        {
        }

        // سازنده با پارامتر
        public AboutUs(bool showBackButton)
        {
            InitializeComponent();

            _showBackButton = showBackButton;

            if (!_showBackButton)
            {
                // مخفی کردن دکمه بازگشت
                btnBack.Visibility = System.Windows.Visibility.Collapsed;
            }
            else
            {
                // نمایش دکمه و اتصال رویداد
                btnBack.Visibility = System.Windows.Visibility.Visible;
                btnBack.Click += BtnBack_Click;
            }

            // تنظیم مقادیر متون
            lblName.Text = AssemblyProduct;
            lblVersion.Text = AssemblyVersion;
            lblCopyright.Text = AssemblyCopyright;
            lblCompany.Text = AssemblyCompany;
            lblDescription.Text = AssemblyDescription;
        }

        #region Assembly Attribute Accessors

        public string AssemblyTitle
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyTitleAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyTitleAttribute titleAttribute = (AssemblyTitleAttribute)attributes[0];
                    if (titleAttribute.Title != "")
                    {
                        return titleAttribute.Title;
                    }
                }
                return System.IO.Path.GetFileNameWithoutExtension(Assembly.GetExecutingAssembly().CodeBase);
            }
        }

        public string AssemblyVersion
        {
            get
            {
                String version = BugReport.AssemblyVersion;
                version = version.Replace("0", "۰").Replace("1", "۱").Replace("2", "۲").Replace("3", "۳").Replace("4", "۴").Replace("5", "۵").Replace("6", "۶").Replace("7", "۷").Replace("8", "۸").Replace("9", "۹");

                return version;
            }
        }

        public string AssemblyDescription
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyDescriptionAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyDescriptionAttribute)attributes[0]).Description;
            }
        }

        public string AssemblyProduct
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyProductAttribute)attributes[0]).Product;
            }
        }

        public string AssemblyCopyright
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCopyrightAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCopyrightAttribute)attributes[0]).Copyright;
            }
        }

        public string AssemblyCompany
        {
            get
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute), false);
                if (attributes.Length == 0)
                {
                    return "";
                }
                return ((AssemblyCompanyAttribute)attributes[0]).Company;
            }
        }
        #endregion


        private void BtnBack_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            bool isUserLogined = !string.IsNullOrEmpty(Properties.Settings.Default.UserToken);

            ((AboutUsVM)DataContext).goBack(isUserLogined);
        }
    }
}
