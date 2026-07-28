using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.Http;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View
{
    /// <summary>
    /// Interaction logic for BugReport.xaml
    /// </summary>
    public partial class BugReport : System.Windows.Controls.UserControl
    {
        StreamContent fileContent = null;
        string extension = null;

        bool isUserLogined = false;

        bool validateTxtReportText = false;
        public BugReport()
        {
            InitializeComponent();

            btnSendBugReport.Click += BtnSendBugReport_Click;
            btnSendBugReport.IsEnabled = false;

            btnChooseFile.Click += BtnChooseFile_Click;
            btnBack.Click += BtnBack_Click;

            txtReportText.LostFocus += TxtReportText_LostFocus;
            txtReportText.TextChanged += TxtReportText_TextChanged; ;

            isUserLogined = !string.IsNullOrEmpty(Properties.Settings.Default.UserToken);
        }

        private void TxtReportText_TextChanged(object sender, TextChangedEventArgs e)
        {
            validateTxtReportText = validateReportTextBox(txtReportText, true);
            validateControls(validateTxtReportText);
        }

        #region Events
        private void TxtReportText_LostFocus(object sender, System.Windows.RoutedEventArgs e)
        {
            validateTxtReportText = validateReportTextBox(txtReportText, false);
            validateControls(validateTxtReportText);
        }


        private void BtnBack_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            ((BugReportVM)DataContext).goBack(isUserLogined);
        }

        private void BtnChooseFile_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog fileDialog = new System.Windows.Forms.OpenFileDialog();
            fileDialog.Title = "فایل مد نظر خود را انتخاب نمایید";
            fileDialog.AddExtension = false;
            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                FileInfo fileInfo = new FileInfo(fileDialog.FileName);
                extension = fileInfo.Extension;
                string filePath = fileDialog.FileName;
                var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                fileContent = new StreamContent(fileStream);

                txtFilePath.Text = fileDialog.FileName;
            }
        }

        private void BtnSendBugReport_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            string text = txtReportText.Text;

            System.Threading.Tasks.Task.Run(() =>
            {
                text += getSettingsInformation(isUserLogined);
                text += getInformation();

                Dispatcher.Invoke(() =>
                {
                    ((BugReportVM)DataContext).sendBugReport((int)ReportTypes.UserReport, text, isUserLogined, fileContent, extension);
                });
            });
        }
        #endregion

        #region Validate Functions

        private void validateControls(bool validateTxtReportText)
        {
            btnSendBugReport.IsEnabled = validateTxtReportText;
        }
        private bool validateReportTextBox(TextBox textBox, bool onlyReturn)
        {
            if (!string.IsNullOrEmpty(textBox.Text) && !string.IsNullOrWhiteSpace(textBox.Text) && textBox.Text.Length > 10)
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
                    errorControl(textBox, "حروف بیشتر از 10 کاراکتر میبایست باشد");

                return false;
            }
        }

        private void errorControl(Control control, string hintText)
        {
            HintAssist.SetHelperText(control, hintText);
            control.Foreground = System.Windows.Media.Brushes.Red;

            Thickness margin = new Thickness(0, 0, 0, 20);
            control.Margin = margin;
        }
        private void normalControl(Control control)
        {
            HintAssist.SetHelperText(control, "");
            control.Foreground = System.Windows.Media.Brushes.Black;

            Thickness margin = new Thickness(0);
            control.Margin = margin;
        }
        #endregion

        #region Functions

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
        public static string AssemblyVersion
        {
            get
            {
                String version = Assembly.GetExecutingAssembly().GetName().Version.ToString();

                // عدد نسخه رو سه قسمتی کردم دستی، چون عدد نسخه پیشفرض 4 قسمتیه
                version = version.Substring(0, version.LastIndexOf("."));

                return version;
            }
        }

        internal static string getSettingsInformation(bool isUserLogined)
        {
            string version = AssemblyVersion.Replace(".", "");

            string informations = "";
            string separator = "----------------------------";
            string tab = "\t";
            string newLine = "\n";

            informations += newLine + separator + newLine;
            informations += "اطلاعات ذخیره شده:" + newLine;

            informations += tab + "نسخه: " + AssemblyVersion + newLine;
            informations += tab + "شماره نسخه: " + version + newLine;

            informations += tab + "حساب کاربری: " + (isUserLogined ? Properties.Settings.Default.Mobile : "خالی") + newLine;
            //informations += tab + "توکن: " + (isUserLogined ? Properties.Settings.Default.UserToken : "خالی") + newLine;
            informations += tab + "آدرس پوشه محیط کاری: " + (!string.IsNullOrEmpty(Properties.Settings.Default.WorkSpaceDirectory) ? Properties.Settings.Default.WorkSpaceDirectory : "خالی") + newLine;

            informations += tab + "تنظیمات عنوان - نمایش دیالوگ انتخاب عکس: " + (Properties.Settings.Default.CaptionSettings_ShowDialogCaptionFigure ? "فعال" : "غیرفعال") + newLine;
            informations += tab + "تنظیمات ویراستار - شامل شدن محدوده پانوشت در تصحیح نیم فاصله : " + (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionFootnote ? "فعال" : "غیرفعال") + newLine;
            informations += tab + "تنظیمات ویراستار - شامل شدن محدوده سرصفحه و پاصفحه در تصحیح نیم فاصله : " + (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionHeadersFooters ? "فعال" : "غیرفعال") + newLine;
            informations += tab + "تنظیمات ویراستار - شامل شدن محدوده فیلد های خاص در تصحیح نیم فاصله : " + (Properties.Settings.Default.VirastarSettings_IncludeHalfSpaceCorrectionSpecialFields ? "فعال" : "غیرفعال") + newLine;
            informations += tab + "تنظیمات ویراستار - کاراکتر های اعداد فارسی استاندارد در تبدیل اعداد جایگزین شود: " + (Properties.Settings.Default.VirastarSettings_IncludeChangeDigitCharacters ? "فعال" : "غیرفعال") + newLine;

            int poemType = Properties.Settings.Default.VirastarSettings_CreatePoemType;
            string poemTypeString;
            if (poemType == (int)PoemTypes.CreatePoemUsingTable)
                poemTypeString = "ساخت شعر با استفاده از Table";
            else if (poemType == (int)PoemTypes.CreatePoemUsingTextColumns)
                poemTypeString = "ساخت شعر با استفاده از Text Columns";
            else
                throw new Exception("مقدار نادرست در VirastarSettings_CreatePoemType\n مقدار: " + poemType.ToString());

            informations += tab + "تنظیمات ویراستار - نوع حالت شعر: " + poemTypeString + newLine;

            return informations;
        }
        internal static string getInformation()
        {
            //Globals.ThisAddIn.Application.Dialogs[WdWordDialog.]

            string informations = "";
            string separator = "----------------------------";
            string tab = "\t";
            string newLine = "\n";

            string wordSoftwarePath = "";
            if (File.Exists(Globals.ThisAddIn.Application.Path + "\\winword.exe"))
            {
                wordSoftwarePath = Globals.ThisAddIn.Application.Path + "\\winword.exe";
            }

            List<string> addInList = new List<string>() { };
            foreach (AddIn addIn in Globals.ThisAddIn.Application.AddIns)
            {
                addInList.Add(addIn.Name);
            }
            foreach (COMAddIn addIn in Globals.ThisAddIn.Application.COMAddIns)
            {
                addInList.Add(addIn.Description);
            }

            var wmi = new ManagementObjectSearcher("select * from Win32_OperatingSystem").Get().Cast<ManagementObject>().First();
            var cpu = new ManagementObjectSearcher("select * from Win32_Processor").Get().Cast<ManagementObject>().First();
            informations += separator + newLine;
            informations += "اطلاعات پایه:" + newLine;
            informations += tab + "تاریخ و ساعت محلی سیستم: " + System.Management.ManagementDateTimeConverter.ToDateTime((string)wmi["LocalDateTime"]).ToString();
            informations += newLine + separator + newLine;

            informations += "اطلاعات افزونه:" + newLine;


            object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute), false);

            informations += tab + "نام افزونه: " + (attributes.Length == 0 ? "" : ((AssemblyProductAttribute)attributes[0]).Product) + newLine;
            informations += tab + "نسخه افزونه: " + AssemblyVersion + newLine;
            informations += separator + newLine;
            informations += "اطلاعات نرم افزار آفیس:" + newLine;
            informations += tab + "آدرس مجموعه آفیس: " + Globals.ThisAddIn.Application.Path + newLine;
            informations += wordSoftwarePath != "" ? tab + "آدرس نرم افزار وُرد: " + wordSoftwarePath + newLine : "";
            informations += tab + "نسخه نرم افزار وُرد: " + Globals.ThisAddIn.Application.Version + " (" + Globals.ThisAddIn.Application.Build + ")\n";
            if (addInList.Count != 0)
            {
                informations += tab + "افزونه های نصب شده:\n";
                foreach (var addIn in addInList)
                {
                    informations += tab + tab + addIn + newLine;
                }
            }

            //informations += tab + "نام Office: " + Globals.ThisAddIn.Application.Options + newLine;
            //informations += tab + "نام Office: " + Globals.ThisAddIn.Application.PickerDialog + newLine;
            informations += separator + newLine;
            informations += "مشخصات سیستم:" + newLine;
            informations += tab + "سیستم عامل: " + ((string)wmi["Caption"]).Trim() + newLine;
            informations += tab + "نسخه سیستم عامل: " + (string)wmi["Version"] + " (" + (string)wmi["OSArchitecture"] + ")" + newLine;

            Graphics graphics = Graphics.FromHwnd(IntPtr.Zero);
            float dpi = (graphics.DpiX * 100) / 96;

            informations += tab + "وضوح نمایش: " + Globals.ThisAddIn.Application.System.HorizontalResolution + "x" + Globals.ThisAddIn.Application.System.VerticalResolution + " (مقیاس: %" + dpi.ToString() + ")" + newLine;
            informations += tab + "حافظه موقت: " + ((UInt64)wmi["TotalVisibleMemorySize"] / 1024).ToString() + "MB (در دسترس: " + ((UInt64)wmi["FreePhysicalMemory"] / 1024).ToString() + "MB)\n";
            informations += tab + "مشخصات پردازنده: " + (string)cpu["Name"] + newLine;
            informations += separator + newLine + "پایان پیام";

            return informations;
        }
        #endregion
    }
}
