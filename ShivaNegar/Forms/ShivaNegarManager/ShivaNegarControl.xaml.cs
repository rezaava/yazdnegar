using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Models;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;
using ShivaNegar.Models;
using System;
using System.Text.Json;
using System.Threading;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace ShivaNegar.Forms.ShivaNegarManager
{
    public partial class ShivaNegarControl : System.Windows.Controls.UserControl, Interfaces.ICloseForm
    {
        public Action CloseForm { get; set; }

        string updateLink = "";

        public ShivaNegarControl(bool initialize)
        {
            InitializeMaterialDesign();
            InitializeComponent();

            if (initialize)
            {
                // Invoke from Pages
                documentManager.TransitionCreateDocumentRequest += () =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        goToCreateDocumentPage();
                    });
                };
                documentManager.TransitionNetworkingManagerRequest += () =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        FileModel fileModel;
                        if (documentManager.NavigationVm.CurrentView is DocumentsVM vm)
                        {
                            fileModel = vm.selectedFileModel;
                            goToNetworkingManager(fileModel);
                        }
                    });
                };
                createDocument.TransitionDocumentManagerRequest += () =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        goToDocumentManagerPage();
                    });
                };
                networkingManagerControl.TransitionDocumentManagerRequest += () =>
                {
                    Dispatcher.Invoke(() =>
                    {
                        goToDocumentManagerPage();
                    });
                };


                btnServerTryAgainDialog.Click += BtnServerTryAgainDialog_Click;
                btnServerCancelDialog.Click += BtnServerCancelDialog_Click;
                btnUpdateMandatory.Click += BtnUpdate_Click;
                btnUpdateOptional.Click += BtnUpdate_Click;
                btnUpdateOptionalCancel.Click += BtnUpdateOptionalCancel_Click;


                System.Threading.Tasks.Task.Run(() =>
                {
                    Thread.Sleep(2000);
                    if (!Globals.ThisAddIn.ManualyDisableServer)
                    {
                        checkUpdate();
                    }
                });
            }
        }


        #region Functions
        internal void goToDocumentManagerPage()
        {
            transitionMain.SelectedIndex = 0;
        }
        internal void goToCreateDocumentPage()
        {
            transitionMain.SelectedIndex = 1;
        }
        internal void goToNetworkingManager(FileModel fileModel)
        {
            transitionMain.SelectedIndex = 2;
            networkingManagerControl.initializeNetworkingManagerControl(fileModel);
        }


        private void BtnUpdate_Click(object sender, RoutedEventArgs e)
        {
            dialogMain.IsOpen = false;
            CloseForm.Invoke();
            if (!string.IsNullOrEmpty(updateLink))
            {
                System.Diagnostics.Process.Start(updateLink);
            }
        }

        private void BtnUpdateOptionalCancel_Click(object sender, RoutedEventArgs e)
        {
            dialogMain.IsOpen = false;

            loadCtrlCheckingUpdated.IsEnabled = false;

            if (string.IsNullOrEmpty(Properties.Settings.Default.UserToken))
            {
                documentManager.goToLoginPage();
            }
            else
            {
                documentManager.goToDocumentsPage();
            }
        }

        private void BtnServerCancelDialog_Click(object sender, RoutedEventArgs e)
        {
            dialogMain.IsOpen = false;
            CloseForm.Invoke();
        }

        private void BtnServerTryAgainDialog_Click(object sender, RoutedEventArgs e)
        {
            dialogMain.IsOpen = false;

            System.Threading.Tasks.Task.Run(() =>
            {
                Thread.Sleep(1000);
                checkUpdate();
            });
        }

        private void showDialog(int dialogIndex)
        {
            gDialogServerError.Visibility = Visibility.Collapsed;
            gDialogUpdateOptional.Visibility = Visibility.Collapsed;
            gDialogUpdateMandatory.Visibility = Visibility.Collapsed;

            if (dialogIndex == 0)
            {
                gDialogServerError.Visibility = Visibility.Visible;
            }
            else if (dialogIndex == 1)
            {
                gDialogUpdateOptional.Visibility = Visibility.Visible;
            }
            else if (dialogIndex == 2)
            {
                gDialogUpdateMandatory.Visibility = Visibility.Visible;
            }

            dialogMain.IsOpen = true;
        }


        private async void checkUpdate()
        {
            UpdateResultModel updateResult = await DedicatedFunctions.checkUpdate();

            if (updateResult.Result == UpdateResultModel.Results.ServerResult_NoUpdateAvailable)
            {
                Dispatcher.Invoke(() =>
                {
                    if (string.IsNullOrEmpty(Properties.Settings.Default.UserToken))
                    {
                        documentManager.goToLoginPage();
                    }
                    else
                    {
                        documentManager.goToDocumentsPage();
                    }
                });

                Thread.Sleep(500);
                Dispatcher.Invoke(() =>
                {
                    dialogMain.IsOpen = false;
                    loadCtrlCheckingUpdated.IsEnabled = false;
                });
            }
            else if (updateResult.Result == UpdateResultModel.Results.ServerResult_UpdateOptional)
            {
                Dispatcher.Invoke(() =>
                {
                    showDialog(1);
                });

                updateResult.Content.TryGetProperty("link", out JsonElement link);
                updateLink = link.GetString();
            }
            else if (updateResult.Result == UpdateResultModel.Results.ServerResult_UpdateMandatory)
            {
                Dispatcher.Invoke(() =>
                {
                    showDialog(2);
                });

                updateResult.Content.TryGetProperty("link", out JsonElement link);
                updateLink = link.GetString();
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    showDialog(0);
                });
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
        #endregion

        private void loadCtrlCheckingUpdated_Loaded(object sender, RoutedEventArgs e)
        {

        }

        public DocumentManagerControl DocumentManagerControl => documentManager;

        internal void ShowDocumentManagerWithArchive()
        {
            transitionMain.SelectedIndex = 0;

            if (documentManager != null && documentManager.NavigationVm != null)
            {
                documentManager.NavigationVm.goToArchive = true;
            }


            documentManager.goToDocumentsPage();

        }

        internal void CreateDocumentPage()
        {
            transitionMain.SelectedIndex = 1;
        }
    }

}
