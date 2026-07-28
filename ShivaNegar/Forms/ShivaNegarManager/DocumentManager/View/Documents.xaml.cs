using System;
using System.IO;
using System.Net;
using System.Windows.Controls;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Models;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;
namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View
{
    /// <summary>
    /// Interaction logic for Thesis.xaml
    /// </summary>
    public partial class Documents : UserControl
    {
        Document doc = null;
        FileModel selectedFileModel = null;

        public Documents()
        {
            InitializeComponent();
            btnOpenDocument.IsEnabled = false;
            btnTurnNetworkingDocument.IsEnabled = false;
            btnMoveToArchive.IsEnabled = false;
            btnReturnFromArchive.IsEnabled = false;


            btnOpenDocument.Click += BtnOpenDocument_Click;
            btnMoveToArchive.Click += BtnMoveToArchive_Click;
            btnReturnFromArchive.Click += BtnReturnFromArchive_Click;
            btnTurnNetworkingDocument.Click += BtnTurnNetworkingDocument_Click;
            txtSearchDocument.TextChanged += TxtSearchDocument_TextChanged;

            btnUseCurrentDocument.Click += BtnUseCurrentDocument_Click;
            btnNewDocument.Click += BtnNewDocument_Click;

            lstBoxFiles.KeyDown += List_KeyDown;
            lstBoxFiles.MouseDoubleClick += list_MouseDoubleClick;

        }

        private void BtnNewDocument_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (doc != null && selectedFileModel != null)
            {
                ((DocumentsVM)DataContext).ShowLoading = true;
                dialogMain.IsOpen = false;

                string filePathToRemove = doc.FullName;
                DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges, false);
                try
                {
                    File.Delete(filePathToRemove);
                    downloadAndOpenDocument(doc, selectedFileModel, filePathToRemove);
                    //checkAndOpenDocument(doc);

                    //DedicatedFunctions.ShowMessage("سند جدید باز شد");
                }
                catch (Exception)
                {
                    DedicatedFunctions.ShowErrorMessage("فایل " + filePathToRemove + " در حال اجرا است، لطفا سند را ببندید");
                }
            }
        }

        private async void BtnUseCurrentDocument_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (doc != null && selectedFileModel != null)
            {
                ((DocumentsVM)DataContext).ShowLoading = true;
                dialogMain.IsOpen = false;

                await DedicatedFunctions.getDocumentServerAndUpdateVariables(doc, Properties.Settings.Default.UserToken, selectedFileModel.ID.ToString());

                //resume
                ((DocumentsVM)DataContext).CloseFormRequest?.Invoke();

                try
                {
                    doc.ActiveWindow.Visible = true;
                }
                catch (Exception)
                {
                }

                DedicatedFunctions.checkAndInitializeOpenedDocument(doc);

                //DedicatedFunctions.ShowMessage("اطلاعات سند بروز شد");
            }
        }

        private void TxtSearchDocument_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (DataContext is DocumentsVM)
            {
                ((DocumentsVM)DataContext).searchList(txtSearchDocument.Text);
            }
        }

        private void lstBoxFiles_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            clickOnceBtnOpenDocument = true;
            if (lstBoxFiles.Items.Count == 0)
            {
                btnOpenDocument.IsEnabled = false;
                btnMoveToArchive.IsEnabled = false;
                btnReturnFromArchive.IsEnabled = false;
                btnTurnNetworkingDocument.IsEnabled = false;
            }
            else
            {
                btnOpenDocument.IsEnabled = true;
                btnMoveToArchive.IsEnabled = true;
                btnReturnFromArchive.IsEnabled = true;

                FileModel fm = (FileModel)lstBoxFiles.SelectedItem;
                bool archiveDocuments = ((DocumentsVM)DataContext).SetAsArchive;

                if (!archiveDocuments)
                {
                    if (!fm.SharedDocuemnt)
                    {
                        btnTurnNetworkingDocument.IsEnabled = true;
                        btnTurnNetworkingDocument.Visibility = System.Windows.Visibility.Visible;
                    }
                    else
                    {
                        btnTurnNetworkingDocument.IsEnabled = false;
                        btnTurnNetworkingDocument.Visibility = System.Windows.Visibility.Collapsed;
                    }
                }
            }
        }

        private void BtnTurnNetworkingDocument_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (lstBoxFiles.SelectedIndex == -1)
                return;


            selectedFileModel = lstBoxFiles.SelectedItem as FileModel;
            ((DocumentsVM)DataContext).goToNetworkingManager(selectedFileModel);
            //FileModel fm = ((FileModel)lstBoxFiles.SelectedItem);
            //bool archiveDocuments = ((DocumentsVM)DataContext).SetAsArchive;

            //if (!archiveDocuments)
            //{
            //    if (fm.IsEditor)
            //    {
            //        btnTurnNetworkingDocument.IsEnabled = true;
            //        btnTurnNetworkingDocument.Visibility = System.Windows.Visibility.Visible;
            //        if (!fm.NetworkedDocument)
            //        {
            //            setIconTurnNetworking(false);
            //            ((DocumentsVM)DataContext).turnNetworking((FileModel)lstBoxFiles.SelectedItem, true);
            //            fm.NetworkedDocument = true;
            //        }
            //        else
            //        {
            //            setIconTurnNetworking(true);
            //            ((DocumentsVM)DataContext).turnNetworking((FileModel)lstBoxFiles.SelectedItem, false);
            //            fm.NetworkedDocument = false;
            //        }
            //    }
            //    else
            //    {
            //        btnTurnNetworkingDocument.IsEnabled = false;
            //        btnTurnNetworkingDocument.Visibility = System.Windows.Visibility.Collapsed;
            //    }
            //}
        }

        private void BtnMoveToArchive_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DataContext is DocumentsVM)
            {
                Dispatcher.Invoke(() =>
                {
                    ((DocumentsVM)DataContext).moveToArchive((FileModel)lstBoxFiles.SelectedItem);
                });

                btnOpenDocument.IsEnabled = false;
                btnMoveToArchive.IsEnabled = false;
                btnReturnFromArchive.IsEnabled = false;
                btnTurnNetworkingDocument.IsEnabled = false;
            }
        }

        private void BtnReturnFromArchive_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (DataContext is DocumentsVM)
            {
                ((DocumentsVM)DataContext).returnFromArchive((FileModel)lstBoxFiles.SelectedItem);
                btnOpenDocument.IsEnabled = false;
                btnMoveToArchive.IsEnabled = false;
                btnReturnFromArchive.IsEnabled = false;
                btnTurnNetworkingDocument.IsEnabled = false;
            }
        }


        private void List_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                openDocument();
            }
        }
        private void list_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            openDocument();
        }


        private bool clickOnceBtnOpenDocument = true;
        private void BtnOpenDocument_Click(object sender, System.Windows.RoutedEventArgs ea)
        {
            openDocument();
        }

        private async void openDocument()
        {
            if (!btnOpenDocument.IsEnabled && btnOpenDocument.Visibility != System.Windows.Visibility.Visible)
                return;

            Document previousDoc = null;

            if (Globals.ThisAddIn.Application.Documents.Count > 0)
                previousDoc = Globals.ThisAddIn.Application.ActiveDocument;

            if (clickOnceBtnOpenDocument)
            {
                clickOnceBtnOpenDocument = false;
                if (lstBoxFiles.SelectedIndex != -1)
                {
                    selectedFileModel = (FileModel)lstBoxFiles.SelectedItem;

                    if (!String.IsNullOrEmpty(((FileModel)lstBoxFiles.SelectedItem).Name.Trim()) && !String.IsNullOrEmpty(((FileModel)lstBoxFiles.SelectedItem).DocumentURL.Trim()))
                    {
                        string filePath = Properties.Settings.Default.WorkSpaceDirectory.Trim('\\') + "\\" + ((FileModel)lstBoxFiles.SelectedItem).Name;

                        if (File.Exists(filePath))
                        {
                            #region Open Document
                            Globals.ThisAddIn.DisableEvents = true;
                            try
                            {
                                string password = StringConstant.DocumentPassword;
                                //int version;
                                //try
                                //{
                                //	version = int.Parse(selectedFileModel.VersionAddin);
                                //}
                                //catch(Exception)
                                //{
                                //	version = 1;
                                //}

                                ////Password Selection
                                //if(version == 1)
                                //	password = StringConstant.DocumentPassword;
                                //else if(version == 2)
                                //	password = StringConstant.DocumentPassword2;


                                doc = Globals.ThisAddIn.Application.Documents.Open(FileName: filePath, PasswordDocument: password, Visible: false, ReadOnly: false, AddToRecentFiles: false, OpenAndRepair: false, NoEncodingDialog: true);
                                if (selectedFileModel.SharedDocuemnt)
                                {
                                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_SharedDocument.ToString(), "1");
                                    try
                                    {
                                        if (doc.ProtectionType == WdProtectionType.wdNoProtection)
                                            doc.Protect(WdProtectionType.wdAllowOnlyReading, false, StringConstant.DocumentProtectionPassword);
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                                else
                                {
                                    DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_SharedDocument.ToString(), "0");
                                    try
                                    {
                                        if (doc.ProtectionType != WdProtectionType.wdNoProtection)
                                            doc.Unprotect(StringConstant.DocumentProtectionPassword);
                                    }
                                    catch (Exception)
                                    {
                                    }
                                }
                            }
                            catch (Exception e)
                            {
                                if (e.HResult == -2146822880)
                                {
                                    DedicatedFunctions.ShowErrorMessage("سند دارای رمز عبور ناشناخته ای میباشد و برنامه قادر به بازکردن آن و قرار دادن آن به لیست اسناد نیست");
                                }
                                else
                                    throw new Exception(e.Message, e);

                                clickOnceBtnOpenDocument = true;

                                if (doc != null)
                                    DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges, false);

                                ((DocumentsVM)DataContext).CloseFormRequest?.Invoke();
                                Globals.ThisAddIn.DisableEvents = false;
                                return;
                            }
                            Globals.ThisAddIn.DisableEvents = false;
                            #endregion

                            string docID = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString());
                            if (docID != "" && docID != SettingValues.NotExist && selectedFileModel.ID == int.Parse(docID))
                            {
                                //DedicatedFunctions.saveDocument(doc);
                                bool accessUpdate = false;
                                bool requireUpdate = false;
                                bool requireNewDocument = false;

                                // check update
                                String docStringUpdatedFile = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedFile.ToString());
                                String docStringUpdatedConfig = DedicatedFunctions.getStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedConfig.ToString());

                                //try
                                //{
                                //DateTime docUpdatedFile = DateTime.Parse(docStringUpdatedFile).ToUniversalTime();
                                //DateTime docUpdatedConfig = DateTime.Parse(docStringUpdatedConfig).ToUniversalTime();
                                DateTime docUpdatedFile = DateTime.Parse(docStringUpdatedFile);
                                DateTime docUpdatedConfig = DateTime.Parse(docStringUpdatedConfig);

                                DateTime lastUpdatedFile = selectedFileModel.UpdatedFile;
                                DateTime lastUpdatedConfig = selectedFileModel.UpdatedConfig;

                                if (docUpdatedConfig.Ticks < selectedFileModel.UpdatedConfig.Ticks)
                                {
                                    requireUpdate = true;

                                    if (docUpdatedFile.Ticks < selectedFileModel.UpdatedFile.Ticks)
                                        requireNewDocument = true;
                                }

                                accessUpdate = true;
                                //}
                                //catch (Exception)
                                //{
                                //    accessUpdate = false;
                                //}

                                // apply update
                                if (requireUpdate && accessUpdate)
                                {
                                    if (requireNewDocument)
                                    {
                                        dialogMain.IsOpen = true;
                                        //System.Windows.Forms.DialogResult dr = DedicatedFunctions.ShowMessage("سند جدیدی همرسانی شده است، آیا میخواهید سند جدید جایگزین سند قبلی شود؟", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                                        return;
                                    }
                                    else
                                    {
                                        await DedicatedFunctions.getDocumentServerAndUpdateVariables(doc, Properties.Settings.Default.UserToken, selectedFileModel.ID.ToString());
                                        //DedicatedFunctions.ShowMessage("اطلاعات سند بروز شد");
                                    }
                                }

                                //resume
                                ((DocumentsVM)DataContext).CloseFormRequest?.Invoke();

                                try
                                {
                                    doc.ActiveWindow.Visible = true;
                                }
                                catch (Exception)
                                {
                                }

                                DedicatedFunctions.checkAndInitializeOpenedDocument(doc);
                            }
                            else
                            {
                                clickOnceBtnOpenDocument = true;
                                DedicatedFunctions.ShowErrorMessage("سندی با این نام و غیر مرتبط با سند انتخاب شده در پوشه کاری وجود دارد، برای بازشدن سند لازم است سندی با نام مشابه وجود نداشته باشد");
                                DedicatedFunctions.closeDocument(doc, WdSaveOptions.wdDoNotSaveChanges, false);
                            }
                        }
                        else
                        {
                            downloadAndOpenDocument(doc, selectedFileModel, filePath);
                        }

                        if (previousDoc != null)
                        {
                            try
                            {
                                if (!File.Exists(previousDoc.FullName))
                                    if (previousDoc.Saved)
                                        DedicatedFunctions.closeDocument(previousDoc, WdSaveOptions.wdDoNotSaveChanges, false);

                            }
                            catch (Exception)
                            {
                            }
                        }
                    }
                    else
                    {
                        clickOnceBtnOpenDocument = true;
                        DedicatedFunctions.ShowErrorMessage("سند یا نامی برای سند در سرور ذخیره نشده است!", email: StringConstant.SupportEmail);
                    }
                }
            }
        }

        private void downloadAndOpenDocument(Document doc, FileModel fileModel, string filePath)
        {
            ((DocumentsVM)DataContext).CloseFormRequest?.Invoke();
            string url = fileModel.DocumentURL;
            WebClient webClient = new WebClient();
            webClient.DownloadFile(url, filePath);

            string password = StringConstant.DocumentPassword;

            //int version;
            //try
            //{
            //    version = int.Parse(fileModel.VersionAddin);
            //}
            //catch (Exception)
            //{
            //    version = 1;
            //}

            ////Password Selection
            //if (version == 1)
            //    password = StringConstant.DocumentPassword;
            //else if (version == 2)
            //    password = StringConstant.DocumentPassword2;

            selectedFileModel = fileModel;

            Globals.ThisAddIn.DisableEvents = false;
            doc = Globals.ThisAddIn.Application.Documents.Open(FileName: filePath, PasswordDocument: password, ReadOnly: false, AddToRecentFiles: false, OpenAndRepair: false, NoEncodingDialog: true);
            if (selectedFileModel.SharedDocuemnt)
            {
                DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_SharedDocument.ToString(), "1");
                if (doc.ProtectionType == WdProtectionType.wdNoProtection)
                    doc.Protect(WdProtectionType.wdAllowOnlyReading, false, StringConstant.DocumentProtectionPassword);
            }
            else
            {
                DedicatedFunctions.setORAddStaticVariableValue(doc, VariableDocumentIDs._variable_document_SharedDocument.ToString(), "0");
                try
                {
                    if (doc.ProtectionType != WdProtectionType.wdNoProtection)
                        doc.Unprotect(StringConstant.DocumentProtectionPassword);
                }
                catch (Exception)
                {
                }
            }

            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_DocumentID.ToString(), fileModel.ID.ToString());

            // این مورد شاید مشکل ساز شود برای بروزرسانی ها
            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedAt.ToString(), fileModel.UpdatedAt.ToString());
            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedFile.ToString(), fileModel.UpdatedFile.ToString());
            DedicatedFunctions.setORAddStaticVariableValue(doc, VariableServerIDs._variable_server_UpdatedConfig.ToString(), fileModel.UpdatedConfig.ToString());

            DedicatedFunctions.saveDocument(doc);
        }
    }
}
