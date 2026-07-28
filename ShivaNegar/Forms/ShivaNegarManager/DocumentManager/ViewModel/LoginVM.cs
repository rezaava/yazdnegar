using System;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using MaterialDesignThemes.Wpf;
using MaterialDesignThemes.Wpf.Transitions;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Utilities;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel
{
    public class LoginVM : ViewModelBase, IGoToDocumentManager
    {
        public Action GoToMain { get; set; }
        public Action GoToDocuments { get; set; }
        public Action GoToLogin { get; set; }

        internal static int loginPageIndex = 0;
        internal static int passwordPageIndex = 1;
        internal static int changePasswordPageIndex = 2;
        internal static int changePassswordResultPageIndex = 3;
        internal static int RegisterResultPageIndex = 4;

        public LoginVM()
        {
        }

        private void login(string token, string mobile)
        {
            Properties.Settings.Default.UserToken = token;
            Properties.Settings.Default.Mobile = mobile;
            Properties.Settings.Default.Save();

            if (Globals.ThisAddIn.Application.Documents.Count != 0 && StringConstant.AdministratorAccounts.Contains(Properties.Settings.Default.Mobile))
            {
                Ribbon.setTabProperties(StringConstant.NameOfProject, true);
                Ribbon.RibbonControlsVisibility(true);
            }
            else
            {
                Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                Ribbon.RibbonControlsVisibility(false);
            }
            //GoToMain.Invoke();
            GoToDocuments.Invoke();
        }

        #region Server Requests
        internal void checkMobileAndRegister(TextBox mobileTextBox, PasswordBox passwordBox, Transitioner transition)
        {
            string urlParameters = "test?mobile=" + mobileTextBox.Text;
            //Start checking the server
            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, null,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;
                root.TryGetProperty("status", out JsonElement statusElement);

                if (statusElement.GetString() == "1")//already exist mobile
                {
                    transition.SelectedIndex = passwordPageIndex;
                    passwordBox.Focus();
                }
                else if (statusElement.GetString() == "0")//new mobile
                {
                    registerRequest(mobileTextBox, transition);
                }
            },
            OnFailed =>
            {

                if (OnFailed.StatusCode == System.Net.HttpStatusCode.ServiceUnavailable)
                {
                    DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable);
                }
                else
                {
                    int resultCode = (int)OnFailed.StatusCode;
                    if (resultCode != (int)System.Net.HttpStatusCode.RequestTimeout || (resultCode >= 400 && resultCode < 500))
                    {
                        DedicatedFunctions.ShowErrorMessage("لطفا مجدد ورود کنید");

                        Properties.Settings.Default.UserToken = "";
                        Properties.Settings.Default.Mobile = "";
                        Properties.Settings.Default.Save();
                        Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                        Ribbon.RibbonControlsVisibility(false);
                        GoToLogin.Invoke();
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره از سمت سرور", (int)OnFailed.StatusCode, email: StringConstant.SupportEmail, mobile: StringConstant.SupportMobile, false);
                    }
                }
            });
        }
        internal void registerRequest(TextBox mobileTextBox, Transitioner transition)
        {
            string urlParameters = "login?mobile=" + mobileTextBox.Text;
            //Start checking the server
            DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, null,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;
                // Mobile and Password is Correct, Allow access.
                root.TryGetProperty("status", out JsonElement statusElement);

                if (statusElement.GetString() == "2")//Register
                {
                    transition.SelectedIndex = RegisterResultPageIndex;
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("پاسخ غیر منتظره از سمت سرور", email: StringConstant.SupportEmail);
                }
            },
            OnFailed =>
            {
                if (OnFailed.StatusCode == System.Net.HttpStatusCode.ServiceUnavailable)
                {
                    DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable);
                }
                else
                {
                    int resultCode = (int)OnFailed.StatusCode;
                    if (resultCode != (int)System.Net.HttpStatusCode.RequestTimeout || (resultCode >= 400 && resultCode < 500))
                    {
                        DedicatedFunctions.ShowErrorMessage("لطفا مجدد ورود کنید");

                        Properties.Settings.Default.UserToken = "";
                        Properties.Settings.Default.Mobile = "";
                        Properties.Settings.Default.Save();
                        Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                        Ribbon.RibbonControlsVisibility(false);
                        GoToLogin.Invoke();
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره از سمت سرور", (int)OnFailed.StatusCode, email: StringConstant.SupportEmail, mobile: StringConstant.SupportMobile, false);
                    }
                }
            });
        }
        internal void forgetPasswordRequest(TextBox mobileTextBox, Transitioner transition)
        {
            string urlParameters = "forget";
            //Start checking the server
            var multipartContent = new MultipartFormDataContent
            {
                { new StringContent(mobileTextBox.Text) ,"mobile"}
            };

            DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, null,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;
                // Mobile and Password is Correct, Allow access.
                root.TryGetProperty("status", out JsonElement statusElement);

                if (statusElement.GetString() == "1")//Send New Password
                {
                    transition.SelectedIndex = changePassswordResultPageIndex;
                }
                else if (statusElement.GetString() == "0")//Password is wrong
                {
                    HintAssist.SetHelperText(mobileTextBox, "شماره موبایل وارد شده وجود ندارد!");
                    mobileTextBox.Foreground = System.Windows.Media.Brushes.Red;
                    mobileTextBox.Margin = new Thickness(0, 20, 0, 20);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("پاسخ غیر منتظره از سمت سرور", email: StringConstant.SupportEmail);
                }
            },
            OnFailed =>
            {

                if (OnFailed.StatusCode == System.Net.HttpStatusCode.ServiceUnavailable)
                {
                    DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable);
                }
                else
                {
                    int resultCode = (int)OnFailed.StatusCode;
                    if (resultCode != (int)System.Net.HttpStatusCode.RequestTimeout || (resultCode >= 400 && resultCode < 500))
                    {
                        DedicatedFunctions.ShowErrorMessage("لطفا مجدد ورود کنید");

                        Properties.Settings.Default.UserToken = "";
                        Properties.Settings.Default.Mobile = "";
                        Properties.Settings.Default.Save();
                        Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                        Ribbon.RibbonControlsVisibility(false);
                        GoToLogin.Invoke();
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره از سمت سرور", (int)OnFailed.StatusCode, email: StringConstant.SupportEmail, mobile: StringConstant.SupportMobile, false);
                    }
                }
            }, multipartContent);
        }
        internal void loginRequest(TextBox mobileTextBox, PasswordBox passwordBox)
        {
            string token;
            string urlParameters = "login?mobile=" + mobileTextBox.Text + "&password=" + passwordBox.Password;

            //Start checking the server
            DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, null,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;
                if (root.TryGetProperty("token", out JsonElement tokenElement) && !string.IsNullOrEmpty(tokenElement.GetString()))
                {
                    // Mobile and Password is Correct, Allow access.
                    root.TryGetProperty("status", out JsonElement statusElement);

                    //get Bearer Token from server
                    token = tokenElement.GetString();

                    if (statusElement.GetString() == "1")//already exist mobile
                    {
                        login(token, mobileTextBox.Text);
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("پاسخ غیر منتظره از سمت سرور", email: StringConstant.SupportEmail);
                    }
                }
                else if (root.TryGetProperty("status", out JsonElement jsonElement2) && jsonElement2.GetString() == "0")
                {
                    // Mobile is Correct , only Password not Correct
                    HintAssist.SetHelperText(passwordBox, "رمز عبور وارد شده اشتباه است");
                    passwordBox.Foreground = System.Windows.Media.Brushes.Red;
                    passwordBox.Margin = new Thickness(0, 20, 0, 20);
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("پاسخ سرور به برنامه دارای اشکال است",
                        (int)ErrorCodes.UnexpectedServerResponse, StringConstant.SupportEmail);
                }
            },
            OnFailed =>
            {

                if (OnFailed.StatusCode == System.Net.HttpStatusCode.ServiceUnavailable)
                {
                    DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable);
                }
                else
                {
                    int resultCode = (int)OnFailed.StatusCode;
                    if (resultCode != (int)System.Net.HttpStatusCode.RequestTimeout || (resultCode >= 400 && resultCode < 500))
                    {
                        DedicatedFunctions.ShowErrorMessage("لطفا مجدد ورود کنید");

                        Properties.Settings.Default.UserToken = "";
                        Properties.Settings.Default.Mobile = "";
                        Properties.Settings.Default.Save();
                        Ribbon.setTabProperties(StringConstant.NameOfProject, false);
                        Ribbon.RibbonControlsVisibility(false);
                        GoToLogin.Invoke();
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره از سمت سرور", (int)OnFailed.StatusCode, email: StringConstant.SupportEmail, mobile: StringConstant.SupportMobile, false);
                    }
                }
            });
        }
        #endregion

    }
}
