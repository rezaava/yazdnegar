using System;
using System.Collections.ObjectModel;
using System.Net.Http;
using System.Text.Json;
using MaterialDesignThemes.Wpf;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Utilities;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel
{
    public class NetworkingVM : ViewModelBase, IGoToDocumentManager
    {
        public ObservableCollection<NetworkingRequestModel> NetworkingRequestModels { get; set; }

        public Action GoToMain { get; set; }
        public Action GoToDocuments { get; set; }
        public Action GoToLogin { get; set; }

        private bool showEmptyList;

        private Badged badged = null;

        public bool ShowEmptyList
        {
            get
            {
                return showEmptyList;
            }
            set
            {
                showEmptyList = value;
                OnPropertyChanged("ShowEmptyList");
            }
        }


        public NetworkingVM()
        {
            NetworkingRequestModels = new ObservableCollection<NetworkingRequestModel>();
        }
        internal void getBadgedButton(Badged controlBadged)
        {
            badged = controlBadged;
        }
        internal void getNetworkingRequests()
        {
            ShowEmptyList = false;

            NetworkingRequestModels.Clear();

            string token = Properties.Settings.Default.UserToken;
            string urlParameters = "shares";

            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;

                if (root.TryGetProperty("status", out JsonElement statusResult) && statusResult.GetString() == "1")
                {
                    root.TryGetProperty("shares0", out JsonElement jsonElement);

                    foreach (JsonElement item in jsonElement.EnumerateArray())
                    {
                        int status = item.GetProperty("status").GetInt32();

                        if (status == 0)
                        {
                            int id = item.GetProperty("id").GetInt32();
                            int role = item.GetProperty("role").GetInt32();

                            item.TryGetProperty("doc", out JsonElement docElement);
                            int docType = docElement.GetProperty("type").GetInt32();

                            docElement.TryGetProperty("config", out JsonElement docConfig);
                            JsonElement jsonVariables = JsonDocument.Parse(docConfig.GetString()).RootElement;

                            string docAuthor = jsonVariables.GetProperty(VariableFieldIDs._variable_field_Author_Fa.ToString()).GetString();
                            string docTitle = jsonVariables.GetProperty(VariableFieldIDs._variable_field_Title_Fa.ToString()).GetString();


                            string typeText = DedicatedFunctions.getDocumentTypePersianName(docType);
                            string message = "برای انجام " + typeText + " با عنوان " + docTitle + " توسط " + docAuthor + " دعوت به همکاری شده اید. لطفا نظر خود را اعلام بفرمایید.";
                            NetworkingRequestModel nrm = new NetworkingRequestModel(id, status, role, message);
                            NetworkingRequestModels.Add(nrm);
                        }
                        else
                        {
                            DedicatedFunctions.ShowErrorMessage("وضعیت اشتباهی در لیست درخواست ها ثبت شده است", sendErrorToServer: true);
                        }
                    }
                    if (NetworkingRequestModels.Count == 0)
                    {
                        if (badged != null)
                            badged.Badge = null;

                        ShowEmptyList = true;
                    }
                    else
                    {
                        ShowEmptyList = false;
                        string shareCounts = NetworkingRequestModels.Count.ToString();
                        shareCounts = shareCounts.Replace("0", "۰").Replace("1", "۱").Replace("2", "۲").Replace("3", "۳").Replace("4", "۴").Replace("5", "۵").Replace("6", "۶").Replace("7", "۷").Replace("8", "۸").Replace("9", "۹");

                        if ((string)badged.Badge != shareCounts)
                        {
                            if (badged != null)
                                badged.Badge = shareCounts;
                        }
                    }
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره در دریافت داده از سرور رخ داد",
                        (int)ErrorCodes.UnexpectedServerResponse, StringConstant.SupportEmail);

                    try
                    {
                        GoToMain?.Invoke();
                    }
                    catch (System.Exception)
                    {
                    }
                }
            },
            OnFailed =>
            {
                System.Windows.Forms.DialogResult dialogResult = DedicatedFunctions.ShowMessage(ErrorMessages.ErrorServiceUnavailable,
                    System.Windows.Forms.MessageBoxButtons.RetryCancel,
                System.Windows.Forms.MessageBoxIcon.Error);

                try
                {
                    GoToMain?.Invoke();
                }
                catch (System.Exception)
                {
                }
            });
        }
        internal void setStatus(NetworkingRequestModel model, int status)
        {
            string token = Properties.Settings.Default.UserToken;
            string urlParameters = "share-accept";
            var multipartContent = new MultipartFormDataContent
                {
                    { new StringContent(model.ID.ToString()) , "id" } ,
                    { new StringContent(status.ToString()) , "status" },
                };

            DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            OnResult =>
            {
                JsonDocument document = JsonDocument.Parse(OnResult);
                JsonElement root = document.RootElement;

                root.TryGetProperty("status", out JsonElement statusResult);
                if (statusResult.GetString() == "1")
                {
                    //if (status == 1)
                    //    DedicatedFunctions.ShowMessage("درخواست با موفقیت ثبت شد و ارتباط ایجاد شد.");
                    //else
                    //    DedicatedFunctions.ShowMessage("درخواست به خواست شما لغو شد.");


                    NetworkingRequestModels.Remove(model);
                    ShowEmptyList = true;
                    if (NetworkingRequestModels.Count == 0)
                    {
                        if (badged != null)
                            badged.Badge = null;

                        ShowEmptyList = true;
                    }
                    else
                    {
                        ShowEmptyList = false;
                        string shareCounts = NetworkingRequestModels.Count.ToString();
                        shareCounts = shareCounts.Replace("0", "۰").Replace("1", "۱").Replace("2", "۲").Replace("3", "۳").Replace("4", "۴").Replace("5", "۵").Replace("6", "۶").Replace("7", "۷").Replace("8", "۸").Replace("9", "۹");

                        if ((string)badged.Badge != shareCounts)
                        {
                            if (badged != null)
                                badged.Badge = shareCounts;
                        }
                    }
                }
                else
                {
                    DedicatedFunctions.ShowMessage("درخواست ارسال نشد! لطفا دوباره امتحان کنید");
                    GoToMain?.Invoke();
                }
            },
            OnFailed =>
            {
                DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable + "\n" +
                    OnFailed.StatusCode + "> " + OnFailed.ReasonPhrase);
                GoToMain?.Invoke();
            }, multipartContent);
        }

    }

    public class NetworkingRequestModel : ViewModelBase
    {
        private int _id;
        private int _role;
        private string _content;

        public NetworkingRequestModel(int id, int status, int role, string content)
        {
            ID = id;
            Role = role;
            Content = content;
        }

        public int ID
        {
            get => _id;
            set => RaisePropertyChanged(ref _id, value);
        }

        public int Role
        {
            get => _role;
            set => RaisePropertyChanged(ref _role, value);
        }

        public string Content
        {
            get => _content;
            set => RaisePropertyChanged(ref _content, value);
        }

    }
    public class NetworkingUserModel : ViewModelBase
    {
        private int _userID;
        private int _requestID;
        private int _roleID;

        private string _userMobile;
        private string _userName;
        private string _roleName;

        public NetworkingUserModel(int userID, int requestID, int roleID, string userMobile, string userName)
        {
            UserID = userID;
            RequestID = requestID;
            RoleID = roleID;
            UserMobile = userMobile;
            UserName = userName;

            if (RoleID == (int)RoleTypes.Supervisor1)
                RoleName = RoleNames.Supervisor1;
            else if (RoleID == (int)RoleTypes.Supervisor2)
                RoleName = RoleNames.Supervisor2;
            else if (RoleID == (int)RoleTypes.Advisor1)
                RoleName = RoleNames.Advisor1;
            else if (RoleID == (int)RoleTypes.Advisor2)
                RoleName = RoleNames.Advisor2;
            else if (RoleID == (int)RoleTypes.Analyst)
                RoleName = RoleNames.Analyst;
        }

        public int UserID
        {
            get => _userID;
            set => RaisePropertyChanged(ref _userID, value);
        }

        public int RequestID
        {
            get => _requestID;
            set => RaisePropertyChanged(ref _requestID, value);
        }

        public int RoleID
        {
            get => _roleID;
            set => RaisePropertyChanged(ref _roleID, value);
        }

        public string UserMobile
        {
            get => _userMobile;
            set => RaisePropertyChanged(ref _userMobile, value);
        }

        public string UserName
        {
            get => _userName;
            set => RaisePropertyChanged(ref _userName, value);
        }

        public string RoleName
        {
            get => _roleName;
            set => RaisePropertyChanged(ref _roleName, value);
        }
    }
}
