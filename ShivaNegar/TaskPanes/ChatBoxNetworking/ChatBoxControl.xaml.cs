using System.Collections.ObjectModel;
using System.Net.Http;
using System.Text.Json;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;

namespace ShivaNegar.TaskPanes.ChatBoxNetworking
{
    /// <summary>
    /// Interaction logic for CrossReferenceControl.xaml
    /// </summary>
    public partial class ChatBoxNetworkingControl : UserControl
    {
        internal ObservableCollection<MessageItemModel> MessagesItems { get; set; } = new ObservableCollection<MessageItemModel>();
        public ChatBoxNetworkingControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();

            btnInsert.Click += BtnInsert_Click;
            txtMessage.PreviewKeyDown += TxtMessage_PreviewKeyDown;
            btnRefreshList.Click += BtnRefreshList_Click;

            lstMessages.ItemsSource = MessagesItems;
        }

        internal void getChatContents(Document doc, int docID)
        {
            MessagesItems.Clear();
            lblNoMessage.Visibility = System.Windows.Visibility.Collapsed;
            string token = Properties.Settings.Default.UserToken;
            string urlParameters = "messages/" + docID.ToString();

            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
                onResult =>
                {
                    JsonDocument document = JsonDocument.Parse(onResult);
                    JsonElement root = document.RootElement;

                    if (root.TryGetProperty("status", out JsonElement statusResult) && statusResult.GetString() == "1")
                    {
                        root.TryGetProperty("messages", out JsonElement jsonElement);

                        foreach (JsonElement item in jsonElement.EnumerateArray())
                        {
                            int id = item.GetProperty("id").GetInt32();
                            //int role = item.GetProperty("role").GetInt32();
                            string messageText = item.GetProperty("text").GetString();
                            RoleTypes role = (RoleTypes)item.GetProperty("role").GetInt32();

                            string caption = "";
                            if (role == RoleTypes.CurrentUser_Chat)
                            {
                                caption = "شما";
                            }
                            else
                            {
                                string roleName = "";
                                if (role == RoleTypes.Supervisor1)
                                    roleName = RoleNames.Supervisor1;
                                else if (role == RoleTypes.Supervisor2)
                                    roleName = RoleNames.Supervisor2;
                                else if (role == RoleTypes.Advisor1)
                                    roleName = RoleNames.Advisor1;
                                else if (role == RoleTypes.Advisor2)
                                    roleName = RoleNames.Advisor2;
                                else if (role == RoleTypes.Analyst)
                                    roleName = RoleNames.Analyst;
                                else if (role == RoleTypes.Creator)
                                    roleName = RoleNames.Creator;

                                item.TryGetProperty("user", out JsonElement docElement);
                                //string userName = docElement.GetProperty("name").GetString();
                                //string userFamily = docElement.GetProperty("family").GetString();
                                string userMobile = docElement.GetProperty("mobile").GetString();
                                //caption = userName + " " + userFamily + " " + "(" + userMobile + ")";
                                caption = roleName + " " + "(" + userMobile + ")";
                            }

                            MessageItemModel nrm = new MessageItemModel(role, messageText, caption);
                            MessagesItems.Add(nrm);
                        }
                        if (MessagesItems.Count == 0)
                        {
                            lblNoMessage.Visibility = System.Windows.Visibility.Visible;
                        }
                        else
                        {
                            lblNoMessage.Visibility = System.Windows.Visibility.Collapsed;
                            if (VisualTreeHelper.GetChildrenCount(lstMessages) > 0)
                            {
                                System.Windows.Controls.Border border = (System.Windows.Controls.Border)VisualTreeHelper.GetChild(lstMessages, 0);
                                ScrollViewer scrollViewer = (ScrollViewer)VisualTreeHelper.GetChild(border, 0);
                                scrollViewer.ScrollToBottom();
                            }
                        }
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره در دریافت داده از سرور رخ داد",
                            (int)ErrorCodes.UnexpectedServerResponse, StringConstant.SupportEmail);
                    }
                },
                OnFailed =>
                {
                    System.Windows.Forms.DialogResult dialogResult = DedicatedFunctions.ShowMessage(ErrorMessages.ErrorServiceUnavailable,
                        System.Windows.Forms.MessageBoxButtons.RetryCancel,
                    System.Windows.Forms.MessageBoxIcon.Error);
                });

        }

        private void BtnRefreshList_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            string id = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString());
            if (!string.IsNullOrEmpty(id.Trim()) && id != SettingValues.NotExist)
            {
                int docID = int.Parse(id.Trim());
                getChatContents(doc, docID);
            }
        }

        private void InitializeMaterialDesign()
        {
            // Create dummy objects to force the MaterialDesign assemblies to be loaded
            // from this assembly, which causes the MaterialDesign assemblies to be searched
            // relative to this assembly's path. Otherwise, the MaterialDesign assemblies
            // are searched relative to Eclipse's path, so they're not found.
            var card = new Card();
            var hue = new Hue("Dummy", Colors.Black, Colors.White);
        }

        internal void initializeChatBoxControl()
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;

            string id = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString());
            if (!string.IsNullOrEmpty(id.Trim()) && id != SettingValues.NotExist)
            {
                int docID = int.Parse(id.Trim());
                getChatContents(doc, docID);
            }
        }

        private void insert()
        {
            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            string docID = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString());
            string message = txtMessage.Text.Trim();
            if (!string.IsNullOrEmpty(message) && !string.IsNullOrEmpty(docID.Trim()) && docID != SettingValues.NotExist)
            {
                string token = Properties.Settings.Default.UserToken;
                string urlParameters = "message";

                var multipartContent = new MultipartFormDataContent
                {
                    { new StringContent(message) , "text" } ,
                    { new StringContent(docID) , "id" },
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
                            //string userName = docElement.GetProperty("name").GetString();
                            //string userFamily = docElement.GetProperty("family").GetString();
                            //string userMobile = docElement.GetProperty("mobile").GetString();

                            //string caption = userName + userFamily + "(" + userMobile + ")";

                            MessagesItems.Add(new MessageItemModel(RoleTypes.CurrentUser_Chat, message, "شما"));
                            txtMessage.Text = "";

                            if (VisualTreeHelper.GetChildrenCount(lstMessages) > 0)
                            {
                                System.Windows.Controls.Border border = (System.Windows.Controls.Border)VisualTreeHelper.GetChild(lstMessages, 0);
                                ScrollViewer scrollViewer = (ScrollViewer)VisualTreeHelper.GetChild(border, 0);
                                scrollViewer.ScrollToBottom();
                            }
                        }
                        else
                        {
                            DedicatedFunctions.ShowMessage("خطا در ارسال پیام");
                            txtMessage.Text = "";
                        }
                    },
                    OnFailed =>
                    {
                        DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable + "\n" +
                            OnFailed.StatusCode + "> " + OnFailed.ReasonPhrase);
                    }, multipartContent);
                });
            }
        }

        #region Events

        private void TxtMessage_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter && (Keyboard.IsKeyDown(Key.LeftShift) || Keyboard.IsKeyDown(Key.RightShift)))
            {
                e.Handled = true;
                insert();
            }
        }
        private void BtnInsert_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            insert();
        }
        #endregion
    }
    internal class MessageItemModel
    {
        public MessageItemModel(RoleTypes roleType, string message, string caption)
        {
            RoleType = roleType;
            Message = message;
            Caption = caption;

            IsCurrentUser = RoleType == RoleTypes.CurrentUser_Chat;
        }

        public bool IsCurrentUser { get; set; }
        public RoleTypes RoleType { get; set; }
        public string Message { get; set; }
        public string Caption { get; set; }
    }
}
