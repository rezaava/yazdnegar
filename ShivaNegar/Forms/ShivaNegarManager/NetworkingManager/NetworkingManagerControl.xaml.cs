using System;
using System.Collections.ObjectModel;
using System.Net.Http;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MaterialDesignThemes.Wpf;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Models;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;

namespace ShivaNegar.Forms.ShivaNegarManager.NetworkingManager
{
    /// <summary>
    /// Interaction logic for Profile.xaml
    /// </summary>
    public partial class NetworkingManagerControl : System.Windows.Controls.UserControl, Interfaces.IStatusFormRequest, Interfaces.IChangeTransitionDocumentManager
    {
        public Action TransitionDocumentManagerRequest { get; set; }

        public Action CloseFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action NormalStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }

        public ObservableCollection<NetworkingUserModel> NetworkingUserModels { get; set; }

        FileModel currentFileModel;
        public NetworkingManagerControl()
        {
            InitializeComponent();

            comboRoleType.Items.Add(RoleNames.Supervisor1);
            comboRoleType.Items.Add(RoleNames.Supervisor2);
            comboRoleType.Items.Add(RoleNames.Advisor1);
            comboRoleType.Items.Add(RoleNames.Advisor2);
            comboRoleType.Items.Add(RoleNames.Analyst);

            toggleButtonAlwaysOnTop.Click += ToggleButtonAlwaysOnTop_Click;
            btnExitCreateDocument.Click += (b1, e) => { TransitionDocumentManagerRequest?.Invoke(); };
            btnMinimize.Click += (b1, e) => { MinimizeStateFormRequest?.Invoke(); };
            btnMaximize.Click += (b1, e) => { MaximizeStateFormRequest?.Invoke(); };

            txtMobile.PreviewTextInput += PhoneNumberTextBox_PreviewTextInput;
            txtMobile.PreviewKeyDown += PreventTypeSpace_PreviewKeyDown;
            btnSendNetworkingRequest.Click += BtnSendNetworkingRequest_Click;
            btnRefreshList.Click += BtnRefreshList_Click;

        }

        internal void initializeNetworkingManagerControl(FileModel fileModel)
        {
            currentFileModel = fileModel;
            getNetworkingUsers();

        }
        private void BtnRefreshList_Click(object sender, RoutedEventArgs e)
        {
            getNetworkingUsers();
        }

        private void ToggleButtonAlwaysOnTop_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (toggleButtonAlwaysOnTop.IsChecked == true)
                AlwaysOnTopEnableRequest?.Invoke();
            else
                AlwaysOnTopDisableRequest?.Invoke();
        }

        private void BtnSendNetworkingRequest_Click(object sender, RoutedEventArgs e)
        {
            if (validateMobile(txtMobile) && validateRole(comboRoleType))
            {
                string token = Properties.Settings.Default.UserToken;
                string urlParameters = "share";

                int roleID = 1;
                string comboValue = comboRoleType.SelectedItem as string;
                if (comboValue == RoleNames.Supervisor1)
                    roleID = (int)RoleTypes.Supervisor1;
                else if (comboValue == RoleNames.Supervisor2)
                    roleID = (int)RoleTypes.Supervisor2;
                else if (comboValue == RoleNames.Advisor1)
                    roleID = (int)RoleTypes.Advisor1;
                else if (comboValue == RoleNames.Advisor2)
                    roleID = (int)RoleTypes.Advisor2;
                else if (comboValue == RoleNames.Analyst)
                    roleID = (int)RoleTypes.Analyst;

                var multipartContent = new MultipartFormDataContent
                {
                    { new StringContent(txtMobile.Text) , "mobile" } ,
                    { new StringContent(currentFileModel.ID.ToString()) , "id" },
                    { new StringContent(roleID.ToString()) , "role" },
                };


                Dispatcher.Invoke(() =>
                {
                    DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
                    OnResult =>
                    {
                        JsonDocument document = JsonDocument.Parse(OnResult);
                        JsonElement root = document.RootElement;

                        root.TryGetProperty("status", out JsonElement statusResult);
                        string status = statusResult.GetString();
                        if (status == "1")
                        {
                            DedicatedFunctions.ShowMessage("درخواست ارسال شد، لطفا تا تایید کاربر برای ورود به شبکه منتظر بمانید");
                            txtMobile.Text = "";
                            comboRoleType.SelectedIndex = -1;
                        }
                        else
                        {
                            DedicatedFunctions.ShowMessage("درخواستی از قبل ثبت شده است!");
                        }
                    },
                    OnFailed =>
                    {
                        DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable + "\n" +
                            OnFailed.StatusCode + "> " + OnFailed.ReasonPhrase);
                        TransitionDocumentManagerRequest?.Invoke();
                    }, multipartContent);
                });
            }
        }

        public void getNetworkingUsers()
        {
            lblNoUser.Visibility = Visibility.Collapsed;

            if (NetworkingUserModels == null)
            {
                NetworkingUserModels = new ObservableCollection<NetworkingUserModel>();
                lstNetworkingUsers.ItemsSource = NetworkingUserModels;
            }
            else
                NetworkingUserModels.Clear();


            string token = Properties.Settings.Default.UserToken;
            string urlParameters = "shares/" + currentFileModel.ID.ToString();

            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            onResult =>
            {
                JsonDocument document = JsonDocument.Parse(onResult);
                JsonElement root = document.RootElement;

                if (root.TryGetProperty("status", out JsonElement statusResult) && statusResult.GetString() == "1")
                {
                    root.TryGetProperty("shares", out JsonElement jsonElement);

                    foreach (JsonElement item in jsonElement.EnumerateArray())
                    {
                        int status = item.GetProperty("status").GetInt32();
                        int id = item.GetProperty("id").GetInt32();
                        int role = item.GetProperty("role").GetInt32();

                        if (status == 1)
                        {
                            item.TryGetProperty("user", out JsonElement docElement);
                            string userName = docElement.GetProperty("name").GetString();
                            string userFamily = docElement.GetProperty("family").GetString();
                            string userMobile = docElement.GetProperty("mobile").GetString();

                            NetworkingUserModel nrm = new NetworkingUserModel(id, role, userMobile, userName + " " + userFamily);
                            NetworkingUserModels.Add(nrm);
                        }
                    }
                    if (NetworkingUserModels.Count == 0)
                    {
                        lblNoUser.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        lblNoUser.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره در دریافت داده از سرور رخ داد",
                        (int)ErrorCodes.UnexpectedServerResponse, StringConstant.SupportEmail);

                    try
                    {
                        ((NetworkingVM)this.DataContext).GoToMain?.Invoke();
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
                    ((NetworkingVM)this.DataContext).GoToMain?.Invoke();
                }
                catch (System.Exception)
                {
                }
            });
        }

        private void btnDeleteAccess_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                var task = button.DataContext as NetworkingUserModel;
                Dispatcher.Invoke(() =>
                {
                    setStatus(task, 0);
                });
            }
            else
            {
                return;
            }

            if (NetworkingUserModels.Count == 0)
            {
                lblNoUser.Visibility = Visibility.Visible;
            }
            else
            {
                lblNoUser.Visibility = Visibility.Collapsed;
            }
        }

        internal void setStatus(NetworkingUserModel model, int status)
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
                    //DedicatedFunctions.ShowMessage("درخواست شما با موفقیت انجام شد");

                    ((ObservableCollection<NetworkingUserModel>)lstNetworkingUsers.ItemsSource).Remove(model);

                    if (NetworkingUserModels.Count == 0)
                    {
                        lblNoUser.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        lblNoUser.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    DedicatedFunctions.ShowMessage("درخواست ارسال نشد! لطفا دوباره امتحان کنید");
                }
            },
            OnFailed =>
            {
                DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable + "\n" +
                    OnFailed.StatusCode + "> " + OnFailed.ReasonPhrase);
            }, multipartContent);
        }



        private bool validateMobile(TextBox textBox)
        {
            if (string.IsNullOrEmpty(textBox.Text) || string.IsNullOrWhiteSpace(textBox.Text))
            {
                HintAssist.SetHelperText(textBox, "شماره موبایل نمی‌تواند خالی باشد.");
                textBox.Foreground = System.Windows.Media.Brushes.Red;

                return false;
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(textBox.Text, @"^\d+$") || textBox.Text.Length < 11 || !textBox.Text.StartsWith("09"))
            {
                HintAssist.SetHelperText(textBox, "شماره موبایل نادرست است");
                textBox.Foreground = System.Windows.Media.Brushes.Red;

                return false;
            }
            else
            {
                HintAssist.SetHelperText(textBox, "");
                textBox.Foreground = System.Windows.Media.Brushes.Black;
                return true;
            }
        }

        private bool validateRole(ComboBox comboBox)
        {
            if (comboBox.SelectedIndex == -1)
            {
                HintAssist.SetHelperText(comboBox, "نقشی را انتخاب کنید");
                comboBox.Foreground = System.Windows.Media.Brushes.Red;

                return false;
            }
            else
            {
                HintAssist.SetHelperText(comboBox, "");
                comboBox.Foreground = System.Windows.Media.Brushes.Black;
                return true;
            }
        }

        #region TextBox Event for Control
        private void PreventTypeSpace_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // اگر کلید Space فشرده شود
            if (e.Key == Key.Space)
            {
                e.Handled = true; // جلوگیری از ورود Space
            }
        }
        private void PhoneNumberTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");

            //((TextBox)sender).Text = ((TextBox)sender).Text.Replace(" " , "");
            //((TextBox)sender).
            e.Handled = regex.IsMatch(e.Text);
        }
        #endregion

    }
    public class NetworkingRequestModel
    {
        private int _status;
        private int _type;
        private int _role;
        private string _content;

        public NetworkingRequestModel(int status, int type, int role, string content)
        {
            Status = status;
            Type = type;
            Role = role;
            Content = content;
        }

        public int Status
        {
            get => _status;
            set => _status = value;
        }

        public int Type
        {
            get => _type;
            set => _type = value;
        }

        public int Role
        {
            get => _role;
            set => _role = value;
        }

        public string Content
        {
            get => _content;
            set => _content = value;
        }

    }
    public class NetworkingUserModel
    {
        private int _id;
        private int _roleID;
        private string _roleName;

        private string _userMobile;
        private string _userName;

        public NetworkingUserModel(int id, int roleID, string userMobile, string userName)
        {
            ID = id;
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

        public int ID
        {
            get => _id;
            set => _id = value;
        }
        public int RoleID
        {
            get => _roleID;
            set => _roleID = value;
        }
        public string RoleName
        {
            get => _roleName;
            set => _roleName = value;
        }

        public string UserMobile
        {
            get => _userMobile;
            set => _userMobile = value;
        }

        public string UserName
        {
            get => _userName;
            set => _userName = value;
        }

    }

}
