using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager
{
    interface IGoToDocumentManager
    {
        Action GoToMain { get; set; }
        Action GoToDocuments { get; set; }
        Action GoToLogin { get; set; }
    }

    public partial class DocumentManagerControl : System.Windows.Controls.UserControl, Interfaces.ICloseForm
    {
        public Action CloseForm { get; set; }
        public Action TransitionCreateDocumentRequest { get; set; }
        public Action TransitionNetworkingManagerRequest { get; set; }

        public DocumentManagerControl()
        {
            InitializeComponent();

            btnCloseApp.Click += btnCloseApp_Click;
            btnWorkSpaceDirectory.Click += BtnWorkSpaceDirectory_Click;
            btnLogout.Click += BtnLogout_Click;
            goToMainPage();

            NavigationVm.GoToLogin += () =>
            {
                goToLoginPage();
            };
            NavigationVm.GoToMain += () =>
            {
                goToMainPage();
            };
            NavigationVm.GoToDocuments += () =>
            {
                goToDocumentsPage();
            };

            NavigationVm.CurrentViewChanged = () =>
            {
                radioBtnNetworking.Foreground = System.Windows.Media.Brushes.White;
                radioBtnNetworking.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
                radioBtnAllDocuments.Foreground = System.Windows.Media.Brushes.White;
                radioBtnAllDocuments.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
                radioBtnArchives.Foreground = System.Windows.Media.Brushes.White;
                radioBtnArchives.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");

                if ((NavigationVm.CurrentView is NetworkingVM networkingVM))
                {
                    networkingVM.getBadgedButton(badgeNetworking);
                    networkingVM.getNetworkingRequests();
                    radioBtnNetworking.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
                    radioBtnNetworking.Background = System.Windows.Media.Brushes.White;
                }
                else if ((NavigationVm.CurrentView is DocumentsVM dvm))
                {
                    dvm.TransitionNetworkingManagerRequest = () =>
                    {
                        TransitionNetworkingManagerRequest?.Invoke();
                    };

                    if (NavigationVm.goToArchive)
                    {
                        radioBtnArchives.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
                        radioBtnArchives.Background = System.Windows.Media.Brushes.White;
                    }
                    else
                    {
                        radioBtnAllDocuments.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
                        radioBtnAllDocuments.Background = System.Windows.Media.Brushes.White;
                    }
                }
            };

            radioBtnAllDocuments.Click += RadioBtnAllDocuments_Click;
            radioBtnArchives.Click += RadioBtnArchives_Click;
            radioBtnCreateDocument.Click += RadioBtnCreateDocument_Click;
            radioBtnNetworking.Click += RadioBtnNetworking_Click;
        }

        public void SelectArchiveTab()
        {
            // تغییر رنگ دکمه‌ها
            radioBtnNetworking.Foreground = System.Windows.Media.Brushes.White;
            radioBtnNetworking.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");

            radioBtnArchives.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
            radioBtnArchives.Background = System.Windows.Media.Brushes.White;
            radioBtnAllDocuments.Foreground = System.Windows.Media.Brushes.White;
            radioBtnAllDocuments.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");

            // ✅ این خط مهم است - تغییر صفحه به بایگانی
            NavigationVm.goToArchive = true;
            NavigationVm.CurrentView = new DocumentsVM(true);
        }

        private void RadioBtnNetworking_Click(object sender, RoutedEventArgs e)
        {
            radioBtnNetworking.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
            radioBtnNetworking.Background = System.Windows.Media.Brushes.White;

            radioBtnAllDocuments.Foreground = System.Windows.Media.Brushes.White;
            radioBtnAllDocuments.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
            radioBtnArchives.Foreground = System.Windows.Media.Brushes.White;
            radioBtnArchives.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
        }

        private void RadioBtnArchives_Click(object sender, RoutedEventArgs e)
        {
            radioBtnNetworking.Foreground = System.Windows.Media.Brushes.White;
            radioBtnNetworking.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");

            radioBtnArchives.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
            radioBtnArchives.Background = System.Windows.Media.Brushes.White;
            radioBtnAllDocuments.Foreground = System.Windows.Media.Brushes.White;
            radioBtnAllDocuments.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
        }
        private void RadioBtnAllDocuments_Click(object sender, RoutedEventArgs e)
        {
            radioBtnNetworking.Foreground = System.Windows.Media.Brushes.White;
            radioBtnNetworking.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");

            radioBtnAllDocuments.Foreground = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
            radioBtnAllDocuments.Background = System.Windows.Media.Brushes.White;
            radioBtnArchives.Foreground = System.Windows.Media.Brushes.White;
            radioBtnArchives.Background = (System.Windows.Media.Brush)new System.Windows.Media.BrushConverter().ConvertFrom("#007ac1");
        }

        private void BtnLogout_Click(object sender, RoutedEventArgs e)
        {
            bool hasDedicatedDocuments = false;
            List<Document> docs = new List<Document>();
            foreach (Document doc in Globals.ThisAddIn.Application.Documents)
            {
                if (DedicatedFunctions.hasAccess(doc) == DedicatedFunctions.AccessType.AccessGranted)
                {
                    // TODO: اینجا ممکن هست برای خروجی سیدی طرح رو در سند ذخیره کند و Task Pane رو ببندد
                    DedicatedFunctions.RemoveAllTaskPanes(doc);
                    hasDedicatedDocuments = true;
                    docs.Add(doc);
                }
            }

            if (hasDedicatedDocuments)
            {
                DialogResult dr = DedicatedFunctions.ShowMessage("آیا اسناد " + StringConstant.NameOfProject + " ذخیره شود؟", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1);
                if (dr == DialogResult.OK)
                {
                    Globals.ThisAddIn.DisableBeforeCloseEvent = true;
                    foreach (var item in docs)
                    {
                        DedicatedFunctions.closeDocument(item, WdSaveOptions.wdSaveChanges, false);
                    }

                    if (Globals.ThisAddIn.accessedInOpen || Globals.ThisAddIn.accessedInStartup)
                    {
                        Globals.ThisAddIn.accessedInOpen = false;
                        Globals.ThisAddIn.accessedInStartup = false;
                        DedicatedFunctions.restoreInitialSettings();
                        //DedicatedFunctions.saveDocument(doc);
                    }
                    Globals.ThisAddIn.DisableBeforeCloseEvent = false;
                }
                else
                {
                    return;
                }
            }
            //TODO:upload document and remove
            Properties.Settings.Default.UserToken = "";
            Properties.Settings.Default.Mobile = "";
            Properties.Settings.Default.Save();
            Ribbon.setTabProperties(StringConstant.NameOfProject, false);
            Ribbon.RibbonControlsVisibility(false);
            goToLoginPage();
        }

        private void RadioBtnCreateDocument_Click(object sender, RoutedEventArgs e)
        {
            TransitionCreateDocumentRequest?.Invoke();
        }

        internal void goToLoginPage()
        {
            radioBtnAllDocuments.IsEnabled = false;
            radioBtnArchives.IsEnabled = false;
            radioBtnCreateDocument.IsEnabled = false;
            radioBtnProfile.IsEnabled = false;
            radioBtnSettings.IsEnabled = false;
            radioBtnNetworking.IsEnabled = false;
            btnLogout.IsEnabled = false;

            NavigationVm.CurrentView = new LoginVM();
        }
        internal void goToMainPage()
        {
            radioBtnAllDocuments.IsEnabled = true;
            radioBtnArchives.IsEnabled = true;
            radioBtnCreateDocument.IsEnabled = true;
            radioBtnProfile.IsEnabled = true;
            radioBtnSettings.IsEnabled = true;
            radioBtnNetworking.IsEnabled = true;
            btnLogout.IsEnabled = true;

            NavigationVm.CurrentView = new MainVM();
        }
        internal void goToDocumentsPage()
        {
            radioBtnAllDocuments.IsEnabled = true;
            radioBtnArchives.IsEnabled = true;
            radioBtnCreateDocument.IsEnabled = true;
            radioBtnProfile.IsEnabled = true;
            radioBtnSettings.IsEnabled = true;
            radioBtnNetworking.IsEnabled = true;
            btnLogout.IsEnabled = true;

            NavigationVm.CurrentView = new DocumentsVM(NavigationVm.goToArchive);

            checkNetworkingRequests();
        }

        private void BtnWorkSpaceDirectory_Click(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists(Properties.Settings.Default.WorkSpaceDirectory))
            {
                Directory.CreateDirectory(Properties.Settings.Default.WorkSpaceDirectory);
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = Properties.Settings.Default.WorkSpaceDirectory,
                UseShellExecute = true,
                Verb = "open"
            });
        }

        private void btnCloseApp_Click(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is NavigationVM dc)
            {
                var documentVM = dc.CurrentView as ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel.DocumentsVM;
                if (documentVM != null)
                {
                    documentVM.threadFillList.Abort();
                }

            }
            CloseForm?.Invoke();
        }

        private void checkNetworkingRequests()
        {
            String token = Properties.Settings.Default.UserToken;
            string urlParameters = "notif";
            DedicatedFunctions.httpAsyncGetRequest(StringConstant.PrimaryServerApiBaseAddress, urlParameters, token,
            OnResult =>
            {
                JsonDocument document = JsonDocument.Parse(OnResult);
                JsonElement root = document.RootElement;

                if (root.TryGetProperty("status", out JsonElement statusElement) && statusElement.GetString() == "1")
                {
                    root.TryGetProperty("shares", out JsonElement sharesCountElement);

                    if (sharesCountElement.GetInt32() == 0)
                    {
                        Dispatcher.Invoke(() =>
                        {
                            badgeNetworking.Badge = null;
                        });
                    }
                    else
                    {
                        string shareCounts = sharesCountElement.GetInt32().ToString();
                        shareCounts = shareCounts.Replace("0", "۰").Replace("1", "۱").Replace("2", "۲").Replace("3", "۳").Replace("4", "۴").Replace("5", "۵").Replace("6", "۶").Replace("7", "۷").Replace("8", "۸").Replace("9", "۹");

                        Dispatcher.Invoke(() =>
                        {
                            badgeNetworking.Badge = shareCounts;
                        });
                    }
                }
            },
            OnFailed =>
            {
            });
        }
    }

}
