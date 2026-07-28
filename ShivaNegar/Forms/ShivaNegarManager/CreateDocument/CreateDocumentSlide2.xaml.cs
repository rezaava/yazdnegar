using System;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using DocumentFormat.OpenXml.ExtendedProperties;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View;
using ShivaNegar.Interfaces;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Collections.Generic;
using System.Net.Http;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Diagnostics; // برای پارس JSON

namespace ShivaNegar.Forms.ShivaNegarManager.CreateDocument
{
    public partial class CreateDocumentSlide2 : UserControl
    {

        private readonly HttpClient _httpClient = new HttpClient();
        //Properties
        public string DocumentName { get; private set; }
        public DocumentTypes DocumentType { get; private set; }
        public TemplateTypes TemplateType { get; private set; }
        public Action CloseForm { get; set; }

        //validate Variables
        private bool validateDocumentName = false;
        private bool validateDocumentNameInServer = false;
        private bool validateDocumentType = false;
        private bool validateTemplateType = false;
        public CreateDocumentSlide2()
        {
            InitializeComponent();

            comboDocumentType.Items.Clear();
            //comboDocumentType.Items.Add(DedicatedFunctions.getDocumentTypePersianName(3));
            comboDocumentType.Items.Add("تحقیق درسی، موضوع ویژه و گزارش کارورزی");
            comboDocumentType.Items.Add(DedicatedFunctions.getDocumentTypePersianName(0));
            comboDocumentType.Items.Add(DedicatedFunctions.getDocumentTypePersianName(1));
            comboDocumentType.Items.Add(DedicatedFunctions.getDocumentTypePersianName(2));

            comboDocumentType.SelectedIndex = -1;
            validateDocumentType = false;


            //old
            //comboDocumentType.Items.Add(DedicatedFunctions.getDocumentTypePersianName(4));
            //comboDocumentType.Items.Add(DedicatedFunctions.getDocumentTypePersianName(5));
            //comboDocumentType.Items.Add(DedicatedFunctions.getDocumentTypePersianName(6));


            comboTemplateType.Items.Clear();
            comboTemplateType.Items.Add(DedicatedFunctions.getTemplateTypePersianName(-1));
            comboTemplateType.Items.Add(DedicatedFunctions.getTemplateTypePersianName(1));
            comboTemplateType.Items.Add(DedicatedFunctions.getTemplateTypePersianName(2));
            comboTemplateType.Items.Add(DedicatedFunctions.getTemplateTypePersianName(3));
            //comboTemplateType.Items.Add(DedicatedFunctions.getTemplateTypePersianName(2));
            //comboTemplateType.SelectedIndex = 0;
            //validateTemplateType = true;

            //Events
            comboDocumentType.SelectionChanged += ComboDocumentType_SelectionChanged;
            comboTemplateType.SelectionChanged += ComboTemplateType_SelectionChanged;
            //Dispatcher.InvokeAsync(() =>
            //{
            //    ValidateDocumentTypeWithServer(comboDocumentType.SelectedIndex.ToString());
            //});


            txtDocumentName.TextChanged += TextBox_TextChanged;
            txtDocumentName.GotFocus += TextBox_GotFocus;
            txtDocumentName.LostFocus += TextBox_LostFocus;

            btnForward.Click += BtnForward_Click;
        }

        private void BtnGoToWebsite_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // باز کردن سایت در مرورگر پیشفرض
                Process.Start(new ProcessStartInfo
                {
                    FileName = "https://shivanegar.ir", // آدرس سایت خودتون رو بذارید
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در باز کردن سایت: {ex.Message}");
            }
        }
        private void BtnForward_Click(object sender, RoutedEventArgs e)
        {
            if (txtDocumentName.Text != txtDocumentName.Text.Trim())
            {
                DocumentName = txtDocumentName.Text.Trim();
                txtDocumentName.Text = DocumentName;
            }
            else
            {
                DocumentName = txtDocumentName.Text;
            }
        }
        public System.Threading.Thread threadCreateDocument;
        Microsoft.Office.Interop.Word.Document previousDocument;
        Microsoft.Office.Interop.Word.Document specifiedDocument;
        Microsoft.Office.Interop.Word.Window specifiedWindow;
        Microsoft.Office.Interop.Word.Window Window;
        private async void btnStore_Click(object sender, RoutedEventArgs e)
        {
            var btn = sender as Button;
            btn.IsEnabled = false;

            try
            {
                // روش صحیح استفاده از Task.Run
                await System.Threading.Tasks.Task.Run(() =>
                {
                    CreateDocumentFromTemplate();
                });

                // یا به صورت خلاصه:
                // await Task.Run(CreateDocumentFromTemplate);
            }
            finally
            {
                btn.IsEnabled = true;
            }

            //try
            //{
            //    // 1. بررسی نام صحیح Embedded Resource
            //    string resourceName = "ShivaNegar.Templates.Blank.docx";

            //    // اگر نام پروژه Shivanegar است و فایل در پوشه Templates است:
            //    // resourceName = "Shivanegar.Templates.Blank.docx";

            //    // 2. دریافت Assembly جاری
            //    Assembly assembly = Assembly.GetExecutingAssembly();

            //    // 3. لیست تمام Embedded Resources (برای دیباگ)
            //    string[] resourceNames = assembly.GetManifestResourceNames();
            //    System.Diagnostics.Debug.WriteLine("موجودیت منابع:");
            //    foreach (var name in resourceNames)
            //    {
            //        System.Diagnostics.Debug.WriteLine(name);
            //    }

            //    // 4. خواندن Resource
            //    using (Stream resourceStream = assembly.GetManifestResourceStream(resourceName))
            //    {
            //        if (resourceStream == null)
            //        {
            //            System.Windows.Forms.MessageBox.Show("فایل قالب پیدا نشد.\nنام مورد جستجو: " + resourceName +
            //                                              "\nموجودیت منابع:\n" + string.Join("", resourceNames));
            //            return;
            //        }

            //        // 5. ایجاد فایل موقت
            //        string tempFilePath = Path.Combine(Path.GetTempPath(), "template_temp_" + Guid.NewGuid().ToString() + ".docx");

            //        using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
            //        {
            //            resourceStream.CopyTo(fileStream);
            //        }

            //        // 6. باز کردن سند در ورد
            //        var wordApp = Globals.ThisAddIn.Application;
            //        Document templateDoc = wordApp.Documents.Open(tempFilePath, ReadOnly: false);

            //        // 7. تمیزکاری (اختیاری)
            //        try { File.Delete(tempFilePath); } catch { }


            //        if (this.DataContext is NavigationVM dc)
            //        {
            //            var documentVM = dc.CurrentView as ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel.DocumentsVM;
            //            if (documentVM != null)
            //            {
            //                documentVM.threadFillList.Abort();
            //            }

            //        }
            //        CloseForm?.Invoke();
            //        System.Windows.Forms.MessageBox.Show("سند با موفقیت ساخته شد.");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    System.Windows.Forms.MessageBox.Show("خطا در ایجاد سند جدید:\n" + ex.ToString());
            //}
        }
        private void CreateDocumentFromTemplate()
        {
            try
            {
                var wordApp = Globals.ThisAddIn.Application;

                // 1. بستن یا مخفی کردن تمام اسناد باز
                foreach (Microsoft.Office.Interop.Word.Window win in wordApp.Windows)
                {
                    win.Visible = false;
                }

                
                // 3. ایجاد سند جدید از قالب
                string resourceName = "ShivaNegar.Templates.Blank.docx";
                Assembly assembly = Assembly.GetExecutingAssembly();

                using (Stream resourceStream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (resourceStream == null)
                    {
                        MessageBox.Show("فایل قالب پیدا نشد!");
                        return;
                    }

                    // ایجاد فایل موقت
                    string tempFilePath = Path.Combine(Path.GetTempPath(), $"template_{Guid.NewGuid()}.docx");

                    using (FileStream fileStream = new FileStream(tempFilePath, FileMode.Create))
                    {
                        resourceStream.CopyTo(fileStream);
                    }

                    // باز کردن سند جدید
                    Document newDoc = wordApp.Documents.Open(tempFilePath);

                    // حذف فایل موقت (اختیاری)
                    try { File.Delete(tempFilePath); } catch { }


                    // فعال کردن پنجره ورد
                    wordApp.Visible = true;
                    wordApp.Activate();
                    newDoc.Activate();
                }


                // ribbon controll
                Ribbon.InitializeRibbon(StringConstant.NameOfProject);
                Ribbon.loadKeyboardShortcut();
                Globals.ThisAddIn.DisableEvents = true;
                DedicatedFunctions.saveDocument(specifiedDocument);
                Globals.ThisAddIn.DisableEvents = false;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                // 4. بستن فرم WPF

            }
            catch (Exception ex)
            {
                MessageBox.Show($"خطا در ایجاد سند: {ex.Message}");
            }
        }
        private Microsoft.Office.Interop.Word.Window GetParentWindow()
        {
            DependencyObject parent = this;
            while (parent != null && !(parent is Microsoft.Office.Interop.Word.Window))
            {
                parent = VisualTreeHelper.GetParent(parent);
            }
            return parent as Microsoft.Office.Interop.Word.Window;
        }


        #region Events
        private async void ComboDocumentType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            validateDocumentType = validateComboBox(comboDocumentType);
            validateControls();
            ValidateDocumentTypeWithServer(comboDocumentType.SelectedIndex.ToString());
        }
        private async System.Threading.Tasks.Task ValidateDocumentTypeWithServer(string documentType)
        {
            btnForward.IsEnabled = false;
            validationText.Text = "در حال بررسی...";
            validationText.Foreground = Brushes.Blue;
            String token = Properties.Settings.Default.UserToken;
            string urlParameters = "get-package/parsanegar/1?package="+comboDocumentType.SelectedIndex.ToString();
            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            OnResult =>
            {
                JsonDocument document = JsonDocument.Parse(OnResult);
                JsonElement root = document.RootElement;

                bool nameExist = false;
                string docuemntName = "";
                Dispatcher.Invoke(() => docuemntName = txtDocumentName.Text.Trim());

                if (root.TryGetProperty("status", out JsonElement statusElement) && statusElement.GetString() == "1")
                {
                    validationText.Text = "پکیج معتبر است ✓";
                    validationText.Foreground = Brushes.Green;
                    packageCard.Visibility = Visibility.Collapsed;

                    // ✅ فقط اگه اسم سند پر باشه فعال کن
                    btnForward.IsEnabled = !string.IsNullOrWhiteSpace(txtDocumentName.Text);
                }
                else
                {
                    validationText.Text = "پکیج معتبر نیست.";
                    validationText.Foreground = Brushes.Red;
                    packageCard.Visibility = Visibility.Visible;
                    btnForward.IsEnabled = false;
                }
            },
            OnFailed =>
            {
                btnForward.IsEnabled = false;
                //System.Windows.Forms.DialogResult dr = DedicatedFunctions.ShowMessage("خطا در بررسی سرور" + "\nپیغام خطا:\n" + OnFailed.ReasonPhrase, System.Windows.Forms.MessageBoxButtons.RetryCancel, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1);

                //if (dr == System.Windows.Forms.DialogResult.Retry)
                //{
                //    checkNameNotExist();
                //}
            });
        }
        private void ComboTemplateType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var selectedItem = comboTemplateType.SelectedIndex;

            if (selectedItem == 2 || selectedItem == 3)
            {
                comboTemplateType.SelectedIndex = -1;
            }
            validateTemplateType = validateComboBox(comboTemplateType);
            validateControls();
        }

        private void checkNameNotExist()
        {
            String token = Properties.Settings.Default.UserToken;
            string urlParameters = "get-data/parsanegar/1";
            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            OnResult =>
            {
                JsonDocument document = JsonDocument.Parse(OnResult);
                JsonElement root = document.RootElement;

                bool nameExist = false;
                string docuemntName = "";
                Dispatcher.Invoke(() => docuemntName = txtDocumentName.Text.Trim());

                if (root.TryGetProperty("status", out JsonElement statusElement) && statusElement.GetString() == "1")
                {
                    root.TryGetProperty("documents", out JsonElement documentsElement);
                    foreach (JsonElement item in documentsElement.EnumerateArray())
                    {
                        string name = item.GetProperty("name").GetString();
                        if (name == docuemntName)
                        {
                            nameExist = true;
                            break;
                        }
                    }
                    Dispatcher.Invoke(() =>
                    {
                        if (nameExist)
                        {
                            errorControl(txtDocumentName, "نام مشابهی در اسناد شما در سرور وجود دارد");
                            validateDocumentNameInServer = false;
                            validateControls();
                        }
                        else
                        {
                            if (txtDocumentName.Foreground != System.Windows.Media.Brushes.Red)
                                normalControl(txtDocumentName);

                            validateDocumentNameInServer = true;
                            validateControls();
                        }
                    });
                }
            },
            OnFailed =>
            {
                System.Windows.Forms.DialogResult dr = DedicatedFunctions.ShowMessage("خطا در بررسی سرور" + "\nپیغام خطا:\n" + OnFailed.ReasonPhrase, System.Windows.Forms.MessageBoxButtons.RetryCancel, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1);

                if (dr == System.Windows.Forms.DialogResult.Retry)
                {
                    checkNameNotExist();
                }
            });
        }
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            checkNameNotExist();

            validateDocumentName = validateDocumentNameTextBox(txtDocumentName, false);
            validateControls();
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
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            validateDocumentNameInServer = false;
            validateDocumentName = validateDocumentNameTextBox(txtDocumentName, true);

            if (validationText.Text.Contains("معتبر"))
            {
                btnForward.IsEnabled = !string.IsNullOrWhiteSpace(txtDocumentName.Text);
            }

            validateControls();
        }

        #endregion

        #region Validators
        private bool validateControls()
        {
            if (validateDocumentType && validateTemplateType && validateDocumentName && validateDocumentNameInServer)
            {
                DocumentType = DedicatedFunctions.getDocumentType(comboDocumentType.SelectedItem as string);
                TemplateType = DedicatedFunctions.getTemplateType(comboTemplateType.SelectedItem as string);

                // ✅ اینجا هم دکمه رو فعال کن
                btnForward.IsEnabled = !string.IsNullOrWhiteSpace(txtDocumentName.Text);

                if (comboTemplateType.SelectedIndex == 0)
                {
                    btnForward.Visibility = Visibility.Collapsed;
                    btnStore.Visibility = Visibility.Visible;
                    btnStore.IsEnabled = true;
                }
                else
                {
                    btnStore.Visibility = Visibility.Hidden;
                    btnForward.Visibility = Visibility.Visible;
                }
                return true;
            }
            else
            {
                btnForward.IsEnabled = false;  // ✅ این رو هم اضافه کن
                return false;
            }
        }

        private bool isCorrectPath(string path, string documentName, bool onlyReturn)
        {
            //lblPath.TextDecorations = null;
            //lblPath.Foreground = System.Windows.Media.Brushes.Black;
            documentName = documentName.ToUpper();

            bool isCorrectName = documentName.IndexOfAny(System.IO.Path.GetInvalidPathChars()) == -1 &&
                !documentName.Contains("/") &&
                !documentName.Contains("\\") &&
                !documentName.Contains(":") &&
                !documentName.Contains("*") &&
                !documentName.Contains("?") &&
                !documentName.Contains("\"") &&
                !documentName.Contains("<") &&
                !documentName.Contains(">") &&
                !documentName.Contains("|");

            bool isCorrectFullName =
                documentName != "CON" &&
                documentName != "PRN" &&
                documentName != "AUX" &&
                documentName != "NUL" &&
                documentName != "COM1" &&
                documentName != "COM2" &&
                documentName != "COM3" &&
                documentName != "COM4" &&
                documentName != "COM5" &&
                documentName != "COM6" &&
                documentName != "COM7" &&
                documentName != "COM8" &&
                documentName != "COM9" &&
                documentName != "LPT1" &&
                documentName != "LPT2" &&
                documentName != "LPT3" &&
                documentName != "LPT4" &&
                documentName != "LPT5" &&
                documentName != "LPT6" &&
                documentName != "LPT7" &&
                documentName != "LPT8" &&
                documentName != "LPT9";

            if (documentName.Length == 0)
            {
                if (!onlyReturn)
                    errorControl(txtDocumentName, "نامی برای سند انتخاب نشده است");
                return false;
            }
            else if (!isCorrectFullName)
            {
                if (!onlyReturn)
                    errorControl(txtDocumentName, "نام سند نامعتبر است، لطفا از نام مناسب استفاده نمایید");
                return false;
            }
            else if (!isCorrectName)
            {
                if (!onlyReturn)
                    errorControl(txtDocumentName, "نام سند نامعتبر است، لطفا از کاراکتر های مناسب استفاده نمایید");
                return false;
            }
            else if (File.Exists(path))
            {
                if (!onlyReturn)
                    errorControl(txtDocumentName, "سندی با این نام وجود دارد، نام دیگری را انتخاب نمایید");
                return false;
            }
            else
            {
                normalControl(txtDocumentName);
                return true;
            }
        }

        private bool validateDocumentNameTextBox(TextBox textBox, bool onlyReturn)
        {
            string path = (Properties.Settings.Default.WorkSpaceDirectory.TrimEnd('\\') + "\\").Replace("\\", "/") + textBox.Text + ".docx";
            //lblPath.Text = path;
            textBox.Text = textBox.Text.TrimStart(' ');

            if (string.IsNullOrEmpty(textBox.Text) || string.IsNullOrWhiteSpace(textBox.Text))
            {
                if (!onlyReturn)
                    errorControl(textBox, "فیلد نمی‌تواند خالی باشد.");

                return false;
            }
            else if (!isCorrectPath(path, textBox.Text, onlyReturn))
            {
                return false;
            }
            else
            {
                normalControl(textBox);
                return true;
            }
        }

        // متد ارسال درخواست به سرور
        private async Task<bool> ValidateFromServer(string baseUrl, string value)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    string url = $"{baseUrl}?selectedValue={Uri.EscapeDataString(value)}";

                    HttpResponseMessage response = await client.GetAsync(url);

                    if (!response.IsSuccessStatusCode)
                    {
                        MessageBox.Show($"خطا در اتصال به سرور: {response.StatusCode}");
                        return false;
                    }

                    string responseBody = await response.Content.ReadAsStringAsync();

                    if (string.IsNullOrWhiteSpace(responseBody))
                    {
                        MessageBox.Show("پاسخی از سرور دریافت نشد.");
                        return false;
                    }

                    JObject json;
                    try
                    {
                        json = JObject.Parse(responseBody);
                    }
                    catch
                    {
                        MessageBox.Show("پاسخ سرور معتبر نیست.");
                        return false;
                    }

                    if (json["status"] == null)
                    {
                        MessageBox.Show("پاسخ سرور مقدار status ندارد.");
                        return false;
                    }

                    return (bool)json["status"];
                }
                catch (Exception ex)
                {
                    MessageBox.Show("خطا در ارسال درخواست: " + ex.Message);
                    return false;
                }
            }
        }

        private bool validateComboBox(ComboBox comboBox)
        {
            if (comboBox.SelectedIndex != -1 && comboBox.SelectedItem != null)
            {
                normalControl(comboBox);
                return true;
            }
            else
            {
                errorControl(comboBox, "موردی انتخاب نشده است");
                return false;
            }
        }
        
        #endregion

        #region Functions
        internal void resetControls()
        {
            #region variables
            validateDocumentType = false;
            validateTemplateType = false;

            DocumentType = DocumentTypes.Project;
            TemplateType = TemplateTypes.UniversityTemplate;//TODO: not set Null
            DocumentName = "";
            #endregion

            #region controls

            Dispatcher.Invoke(() =>
            {
                btnForward.IsEnabled = false;

                comboDocumentType.SelectedIndex = -1;
                validateDocumentType = false;
                normalControl(comboDocumentType);

                comboTemplateType.SelectedIndex = -1;
                normalControl(comboTemplateType);

                txtDocumentName.Text = "";
                normalControl(txtDocumentName);
            });
            #endregion
        }

        private void errorControl(Control control, string hintText)
        {
            HintAssist.SetHelperText(control, hintText);
            control.Foreground = System.Windows.Media.Brushes.Red;
            control.Margin = new Thickness(5, 5, 5, 20);
        }
        private void normalControl(Control control)
        {
            HintAssist.SetHelperText(control, "");
            control.Foreground = System.Windows.Media.Brushes.Black;
            control.Margin = new Thickness(5);
        }

        #endregion
    }
}
// کلاس برای پاسخ سرور (در همان فایل)
public class ValidationResponse
{
    public bool Success { get; set; }
    public string Message { get; set; }
}