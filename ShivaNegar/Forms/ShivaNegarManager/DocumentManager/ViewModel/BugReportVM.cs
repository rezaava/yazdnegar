using System;
using System.Net.Http;
using System.Text.Json;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Utilities;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel
{
    public class BugReportVM : ViewModelBase, IGoToDocumentManager
    {
        public Action GoToMain { get; set; }
        public Action GoToDocuments { get; set; }
        public Action GoToLogin { get; set; }

        public BugReportVM()
        {
        }

        internal void goBack(bool isUserLogined)
        {
            if (isUserLogined)
                GoToMain.Invoke();
            else
                GoToLogin.Invoke();
        }
        internal void sendBugReport(int reportType, string text, bool isUserLogined, StreamContent fileContent = null, string fileExtension = null)
        {
            string urlParameters = "report";

            var multipartContent = new MultipartFormDataContent
            {
                { new StringContent(reportType.ToString()) , "type" } ,
                { new StringContent(text) , "text" }
            };

            if (fileContent != null)
                multipartContent.Add(fileContent, "file", "reportFile" + fileExtension);

            //Start checking the server
            DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, null,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;

                root.TryGetProperty("status", out JsonElement statusElement);
                if (statusElement.GetString() == "1")//successful
                {
                    DedicatedFunctions.ShowMessage("گزارش شما با موفقیت ارسال شد");
                    if (isUserLogined)
                        GoToMain.Invoke();
                    else
                        GoToLogin.Invoke();
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
    }
}
