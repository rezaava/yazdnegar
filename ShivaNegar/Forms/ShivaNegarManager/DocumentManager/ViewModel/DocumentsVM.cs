using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Models;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Utilities;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel
{
    public class DocumentsVM : ViewModelBase, Interfaces.IStatusFormRequest
    {
        SynchronizationContext uiContext = SynchronizationContext.Current;

        internal FileModel selectedFileModel = null;

        public Action TransitionNetworkingManagerRequest { get; set; }

        List<FileModel> allFileModels = new List<FileModel>();
        public ObservableCollection<FileModel> FileModels { get; set; }

        public Action CloseFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action NormalStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }

        public System.Threading.Thread threadFillList = null;

        public enum MessageType
        {
            ConnectionError,
            NoDocumentExist,
            NoDocumentExistInArchive,
            DocumentsAlreadyOpened,
            SearchNotFound,

        }
        MessageType messageType = MessageType.NoDocumentExist;


        private string message;
        public string Message
        {
            get
            {
                return message;
            }
            set
            {
                message = value;
                OnPropertyChanged("Message");
            }
        }

        private bool showMessage;
        public bool ShowMessage
        {
            get
            {
                return showMessage;
            }
            set
            {
                showMessage = value;
                if (showMessage)
                {
                    if (messageType == MessageType.ConnectionError)
                    {
                        Message = "اتصال اینترنت خود را بررسی کنید";
                    }
                    else if (messageType == MessageType.NoDocumentExist)
                    {
                        Message = "سندی ساخته نشده است";
                    }
                    else if (messageType == MessageType.NoDocumentExistInArchive)
                    {
                        Message = "سندی در بایگانی وجود ندارد";
                    }
                    else if (messageType == MessageType.DocumentsAlreadyOpened)
                    {
                        Message = "تمامی اسناد هم اکنون باز است";
                    }
                    else if (messageType == MessageType.SearchNotFound)
                    {
                        Message = "سندی با عبارت مورد نظر یافت نشد";
                    }
                }
                OnPropertyChanged("ShowMessage");
            }
        }

        private bool setAsArchive;
        public bool SetAsArchive
        {
            get
            {
                return setAsArchive;
            }
            set
            {
                setAsArchive = value;
                OnPropertyChanged("SetAsArchive");
            }
        }


        private string openedDocumentMessage;
        public string OpenedDocumentMessage
        {
            get
            {
                return openedDocumentMessage;
            }
            set
            {
                openedDocumentMessage = value;
                OnPropertyChanged("OpenedDocumentMessage");
            }
        }

        private bool isDocumentsOpened;
        public bool IsDocumentsOpened
        {
            get
            {
                return isDocumentsOpened;
            }
            set
            {
                isDocumentsOpened = value;
                OnPropertyChanged("IsDocumentsOpened");
            }
        }

        private bool showLoading;
        public bool ShowLoading
        {
            get
            {
                return showLoading;
            }
            set
            {
                showLoading = value;
                OnPropertyChanged("ShowLoading");
            }
        }

        public DocumentsVM(bool isArchive)
        {
            FileModels = new ObservableCollection<FileModel>();
            SetAsArchive = isArchive;
            ShowLoading = true;
            ShowMessage = false;
            IsDocumentsOpened = false;


            //List<FileInfo> allFiles = new List<FileInfo>();
            //List<FileInfo> files = new List<FileInfo>();
            //if(isArchive)
            //{
            //	string path = Properties.Settings.Default.WorkSpaceDirectory + StringConstant.ArchiveFolder;
            //	Directory.CreateDirectory(path);
            //	DirectoryInfo info = new DirectoryInfo(path);
            //	files.AddRange(info.GetFiles());

            //	DedicatedFunctions.getAllWorkSpaceFile(path , files , out allFiles);
            //}
            //else
            //{
            //	string path = Properties.Settings.Default.WorkSpaceDirectory;
            //	Directory.CreateDirectory(path);
            //	DirectoryInfo info = new DirectoryInfo(path);
            //	files.AddRange(info.GetFiles());

            //	DedicatedFunctions.getAllWorkSpaceFile(path , files , out allFiles);
            //}
            //threadFillList = new Thread(() => fillList(allFiles.ToArray()));
            //threadFillList.Start();

            threadFillList = new Thread(() => fillList());
            threadFillList.Start();
        }

        internal void goToNetworkingManager(FileModel fileModel)
        {
            selectedFileModel = fileModel;
            TransitionNetworkingManagerRequest?.Invoke();
        }
        private static bool isFileLocked(string filePath)
        {
            FileStream stream = null;

            try
            {
                // سعی کنید فایل را باز کنید
                stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            }
            catch (IOException)
            {
                // اگر IOException دریافت کردید، فایل در حال استفاده است
                return true;
            }
            finally
            {
                // اگر فایل باز شده بود، آن را ببندید
                stream?.Close();
            }

            // اگر فایل باز نشد، در حال استفاده نیست
            return false;
        }

        private void setArchiveStatus(FileModel fileModel, bool archive)
        {
            string archiveFolder = Properties.Settings.Default.WorkSpaceDirectory + StringConstant.ArchiveFolder;
            string documentFolder = Properties.Settings.Default.WorkSpaceDirectory;

            if (archive)
            {
                if (File.Exists(documentFolder + fileModel.Name) && isFileLocked(documentFolder + fileModel.Name))
                {
                    DedicatedFunctions.ShowErrorMessage("ابتدا سند خود را بسته و سپس اقدام به بایگانی کردن کنید");
                    return;
                }
            }
            else
            {
                if (File.Exists(archiveFolder + fileModel.Name) && isFileLocked(archiveFolder + fileModel.Name))
                {
                    DedicatedFunctions.ShowErrorMessage("ابتدا سند خود را بسته و سپس اقدام به خارج کردن سند از بایگانی کنید");
                    return;
                }
            }

            int archiveStatus = archive ? 1 : 0;

            string token = Properties.Settings.Default.UserToken;
            int documentID = fileModel.ID;
            string urlParameters = "save/parsanegar/1?id=" + documentID.ToString() + "&archive=" + archiveStatus.ToString();

            DedicatedFunctions.httpAsyncPostRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            OnResult =>
            {
                Directory.CreateDirectory(documentFolder);
                Directory.CreateDirectory(archiveFolder);

                try
                {
                    if (archive)
                    {
                        if (File.Exists(documentFolder + fileModel.Name))
                            File.Move(documentFolder + fileModel.Name, archiveFolder + fileModel.Name);
                    }
                    else
                    {
                        if (File.Exists(archiveFolder + fileModel.Name))
                            File.Move(archiveFolder + fileModel.Name, documentFolder + fileModel.Name);
                    }

                    FileModels.Remove(fileModel);
                    allFileModels.Remove(fileModel);

                    if (FileModels.Count == 0)
                    {
                        int documentsAlreadyOpenedCount = 0;
                        foreach (Microsoft.Office.Interop.Word.Document doc in Globals.ThisAddIn.Application.Documents)
                        {
                            if (DedicatedFunctions.hasAccess(doc) == DedicatedFunctions.AccessType.AccessGranted)
                            {
                                documentsAlreadyOpenedCount++;
                            }
                        }

                        if (documentsAlreadyOpenedCount != 0)
                        {
                            messageType = MessageType.DocumentsAlreadyOpened;
                        }
                        else
                        {
                            if (SetAsArchive)
                                messageType = MessageType.NoDocumentExistInArchive;
                            else
                                messageType = MessageType.NoDocumentExist;
                        }

                        ShowMessage = true;
                    }
                }
                catch (System.Exception error)
                {
                    DedicatedFunctions.ShowErrorMessage(error.Message);
                }
            },
            OnFailed =>
            {
                DedicatedFunctions.ShowErrorMessage(ErrorMessages.ErrorServiceUnavailable + "\n" +
                    OnFailed.StatusCode + "> " + OnFailed.ReasonPhrase);
            });

        }
        public void fillList()
        {
            Thread.Sleep(500);

            String token = Properties.Settings.Default.UserToken;
            string urlParameters = "get-data/parsanegar/1";
            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            OnResult =>
            {
                JsonDocument document = JsonDocument.Parse(OnResult);
                JsonElement root = document.RootElement;

                if (root.TryGetProperty("status", out JsonElement statusElement) && statusElement.GetString() == "1")
                {
                    root.TryGetProperty("documents", out JsonElement documentsElement);
                    root.TryGetProperty("shares", out JsonElement sharedDocumentsElement);

                    List<JsonElement> documentList = new List<JsonElement>();

                    documentList.AddRange(documentsElement.EnumerateArray());
                    int documentsCount = documentList.Count;
                    int tempCounter = 0;

                    documentList.AddRange(sharedDocumentsElement.EnumerateArray());

                    List<FileModel> fileModels = new List<FileModel>();
                    foreach (JsonElement item in documentList)
                    {
                        try
                        {
                            int id = item.GetProperty("id").GetInt32();
                            string name = item.GetProperty("name").GetString();
                            name += name == null ? "" : ".docx";
                            string config = item.GetProperty("config").GetString();
                            string fileURL = item.GetProperty("file").GetString();

                            bool archive = item.GetProperty("archive").GetInt16() == 1 ? true : false;
                            if (fileURL != null)
                            {
                                fileURL = StringConstant.PrimaryServerBaseAddress + fileURL;
                            }

                            int networkedDocument = 0;

                            bool sharedDocuemnt = false;
                            if (tempCounter >= documentsCount)
                                sharedDocuemnt = true;

                            if (archive == SetAsArchive)
                            {
                                string uncertainText = "نامشخص";

                                string type = DedicatedFunctions.getDocumentTypePersianName(item.GetProperty("type").GetInt32());
                                DateTime createdAt = item.GetProperty("created_at").GetDateTime();
                                DateTime updatedAt = item.GetProperty("updated_at").GetDateTime();
                                //DateTime createdAt = DateTime.Parse(item.GetProperty("created_at").GetString());
                                //DateTime updatedAt = DateTime.Parse(item.GetProperty("updated_at").GetString());
                                DateTime updatedFile = DateTime.Parse(item.GetProperty("updated_file").GetString());
                                DateTime updatedConfig = DateTime.Parse(item.GetProperty("updated_config").GetString());
                                JsonElement configElement = JsonDocument.Parse(config).RootElement;

                                String versionAddin = "1";
                                versionAddin = configElement.TryGetProperty(VariableVersionIDs._variable_version_AddIn.ToString(), out JsonElement versionAddinElement) ? versionAddinElement.GetString() : "1";

                                string abstractFa = configElement.TryGetProperty(VariableFieldIDs._variable_field_Abstract_Fa.ToString(), out JsonElement abstractElement) ? abstractElement.GetString() : uncertainText;

                                string university = configElement.TryGetProperty(VariableFieldIDs._variable_field_University_Fa.ToString(), out JsonElement universityElement) ? universityElement.GetString() : uncertainText;
                                string department = configElement.TryGetProperty(VariableFieldIDs._variable_field_Department_Fa.ToString(), out JsonElement departmentElement) ? departmentElement.GetString() : uncertainText;
                                string group = configElement.TryGetProperty(VariableFieldIDs._variable_field_Group_Fa.ToString(), out JsonElement groupElement) ? groupElement.GetString() : uncertainText;
                                string author = configElement.TryGetProperty(VariableFieldIDs._variable_field_Author_Fa.ToString(), out JsonElement authorElement) ? authorElement.GetString() : uncertainText;
                                string supervisor = configElement.TryGetProperty(VariableFieldIDs._variable_field_Supervisor_Fa.ToString(), out JsonElement supervisorElement) ? supervisorElement.GetString() : uncertainText;
                                string advisor = configElement.TryGetProperty(VariableFieldIDs._variable_field_Advisor_Fa.ToString(), out JsonElement advisorElement) ? advisorElement.GetString() : uncertainText;
                                string fieldOfStudy = configElement.TryGetProperty(VariableFieldIDs._variable_field_FieldOfStudy_Fa.ToString(), out JsonElement fieldOfStudyElement) ? fieldOfStudyElement.GetString() : uncertainText;
                                string defenseDate = configElement.TryGetProperty(VariableFieldIDs._variable_field_DefenseDate_Fa.ToString(), out JsonElement defenseDateElement) ? defenseDateElement.GetString() : uncertainText;
                                string academicDegree = configElement.TryGetProperty(VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString(), out JsonElement academicDegreeElement) ? academicDegreeElement.GetString() : uncertainText;
                                string title = configElement.TryGetProperty(VariableFieldIDs._variable_field_Title_Fa.ToString(), out JsonElement titleElement) ? titleElement.GetString() : uncertainText;
                                string areaOfStudy = configElement.TryGetProperty(VariableFieldIDs._variable_field_AreaOfStudy_Fa.ToString(), out JsonElement areaOfStudyElement) ? areaOfStudyElement.GetString() : uncertainText;


                                FileModel fileModel = new FileModel();
                                fileModel.VersionAddin = versionAddin;
                                fileModel.ID = id;
                                fileModel.Abstract = abstractFa;
                                fileModel.Name = name;
                                fileModel.DocumentURL = fileURL;
                                fileModel.Type = type;
                                fileModel.TypeHolder = type;
                                fileModel.CreatedAt = createdAt;
                                fileModel.UpdatedAt = updatedAt;
                                fileModel.UpdatedFile = updatedFile;
                                fileModel.UpdatedConfig = updatedConfig;
                                fileModel.Image = Icon.FromHandle(Properties.ResourceRibbonIcons.ShivaNegarDocuments.GetHicon());
                                fileModel.TitleSearched = "";
                                fileModel.Title = title;
                                fileModel.Author = author;
                                fileModel.Supervisor = supervisor;
                                fileModel.FieldOfStudy = fieldOfStudy;
                                fileModel.UniversityName = university;
                                fileModel.DepartmentOfUniversity = department;
                                fileModel.GroupOFUniversity = group;

                                fileModel.NetworkedDocument = networkedDocument == 1 ? true : false;
                                fileModel.SharedDocuemnt = sharedDocuemnt;

                                fileModels.Add(fileModel);
                            }
                        }
                        catch (Exception e)
                        {
                            DedicatedFunctions.ShowErrorMessage("خطا در گرفتن اطلاعات سند" + "\nپیغام خطا:\n" + e.Message);
                            //uiContext.Send(x => ShowLoading = false , null);
                        }

                        tempCounter++;
                    }

                    if (fileModels.Count == 0)
                    {
                        uiContext.Send(x => ShowLoading = false, null);
                        Thread.Sleep(500);
                        uiContext.Send(x => messageType = MessageType.NoDocumentExist, null);
                        uiContext.Send(x => ShowMessage = true, null);
                        return;
                    }
                    else
                    {
                        int documentsAlreadyOpenedCount = 0;
                        int fileModelCount = fileModels.Count;
                        foreach (Microsoft.Office.Interop.Word.Document doc in Globals.ThisAddIn.Application.Documents)
                        {
                            if (DedicatedFunctions.hasAccess(doc) == DedicatedFunctions.AccessType.AccessGranted)
                            {
                                string documentID = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString());
                                FileModel fileModel = fileModels.Where(model => model.ID == int.Parse(documentID)).FirstOrDefault();
                                if (fileModel != null)
                                {
                                    fileModels.Remove(fileModel);
                                    documentsAlreadyOpenedCount++;
                                }
                            }
                        }

                        if (fileModelCount == 0)
                        {
                            uiContext.Send(x => ShowLoading = false, null);
                            Thread.Sleep(500);
                            if (SetAsArchive)
                                uiContext.Send(x => messageType = MessageType.NoDocumentExistInArchive, null);
                            else
                                uiContext.Send(x => messageType = MessageType.NoDocumentExist, null);
                            uiContext.Send(x => ShowMessage = true, null);

                            return;
                        }
                        else if (documentsAlreadyOpenedCount == fileModelCount)
                        {
                            uiContext.Send(x => ShowLoading = false, null);
                            Thread.Sleep(500);
                            uiContext.Send(x => messageType = MessageType.DocumentsAlreadyOpened, null);
                            uiContext.Send(x => ShowMessage = true, null);
                            return;
                        }
                        else
                        {
                            foreach (var item in fileModels.OrderBy(p => p.UpdatedAt).Reverse())
                            {
                                uiContext.Send(x => FileModels.Add(item), null);
                            }
                            uiContext.Send(x => allFileModels = FileModels.ToList(), null);
                            Thread.Sleep(500);
                            uiContext.Send(x => ShowLoading = false, null);
                        }

                        if (documentsAlreadyOpenedCount != 0)
                        {
                            Thread.Sleep(500);
                            uiContext.Send(x => OpenedDocumentMessage = documentsAlreadyOpenedCount + " عدد از سند های شما هم اکنون باز است", null);
                            uiContext.Send(x => IsDocumentsOpened = true, null);
                        }
                    }
                }
            },
            OnFailed =>
            {
                DedicatedFunctions.ShowErrorMessage("خطا در گرفتن لیست اسناد");
                try
                {
                    uiContext.Send(x => ShowLoading = false, null);
                    Thread.Sleep(500);
                    uiContext.Send(x => messageType = MessageType.ConnectionError, null);
                    uiContext.Send(x => ShowMessage = true, null);
                }
                catch (Exception)
                {
                }
            });
        }

        public void moveToArchive(FileModel fileModel)
        {
            setArchiveStatus(fileModel, true);
        }
        public void returnFromArchive(FileModel fileModel)
        {
            setArchiveStatus(fileModel, false);
        }
        public void searchList(string searchText)
        {
            if (allFileModels.Count >= 1)
            {
                ShowMessage = false;
                FileModels.Clear();
                if (string.IsNullOrEmpty(searchText))
                {
                    foreach (FileModel item in allFileModels)
                    {
                        FileModels.Add(item);
                    }
                }
                else
                {
                    foreach (FileModel item in allFileModels)
                    {
                        if (item.Title.Contains(searchText))
                        {
                            FileModel newFileModel = item;
                            newFileModel.TitleSearched = searchText;
                            FileModels.Add(newFileModel);
                        }
                    }

                    if (FileModels.Count == 0)
                    {
                        messageType = MessageType.SearchNotFound;
                        ShowMessage = true;
                    }
                }
            }
        }

    }

}
