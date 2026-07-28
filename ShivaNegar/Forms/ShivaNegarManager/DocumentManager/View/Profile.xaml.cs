using System.Collections.Generic;
using System.Text.Json;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View
{
    /// <summary>
    /// Interaction logic for Profile.xaml
    /// </summary>
    public partial class Profile : System.Windows.Controls.UserControl
    {
        public Profile()
        {
            InitializeComponent();

            btnSaveChanges.Click += BtnSaveChanges_Click;
            getUserInformation(Properties.Settings.Default.UserToken);
        }

        private void BtnSaveChanges_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Dictionary<string, string> keyValuePairs = new Dictionary<string, string>()
            {
                {"name",txtName.Text.Trim() },
                {"family",txtFamily.Text.Trim() },
            };

            Dispatcher.Invoke(() => { ((ProfileVM)DataContext).setUserInformation(Properties.Settings.Default.UserToken, keyValuePairs); });
        }

        private void getUserInformation(string userToken)
        {
            string urlParameters = "profile";
            //Start checking the server
            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, userToken,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;
                root.TryGetProperty("status", out JsonElement statusElement);

                if (statusElement.GetString() == "1")//user exist.
                {
                    root.TryGetProperty("user", out JsonElement userElement);

                    userElement.TryGetProperty("mobile", out JsonElement mobileElement);
                    txtMobile.Text = mobileElement.GetString();

                    userElement.TryGetProperty("name", out JsonElement nameElement);
                    txtName.Text = nameElement.GetString();

                    userElement.TryGetProperty("family", out JsonElement familyElement);
                    txtFamily.Text = familyElement.GetString();
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
                        ((ProfileVM)DataContext).GoToLogin.Invoke();
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره از سمت سرور", (int)OnFailed.StatusCode, email: StringConstant.SupportEmail, mobile: StringConstant.SupportMobile, false);
                    }
                }
            });
        }

    }
}
