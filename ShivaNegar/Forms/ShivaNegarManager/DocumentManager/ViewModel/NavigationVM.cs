using System;
using System.Windows.Input;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Utilities;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel
{

    class NavigationVM : ViewModelBase, IGoToDocumentManager
    {
        public Action GoToMain { get; set; }
        public Action GoToDocuments { get; set; }
        public Action GoToLogin { get; set; }
        public Action CurrentViewChanged { get; set; }

        public NavigationVM()
        {
            AboutUsCommand = new RelayCommand(AboutUs);
            LoginCommand = new RelayCommand(Login);
            AllDocumentsCommand = new RelayCommand(AllDocuments);
            ArchiveDocumentsCommand = new RelayCommand(ArchiveDocuments);
            SettingsCommand = new RelayCommand(Settings);
            ProfileCommand = new RelayCommand(Profile);
            BugReportCommand = new RelayCommand(BugReport);
            MainCommand = new RelayCommand(Main);
            NetworkingCommand = new RelayCommand(Networking);
        }


        private object _currentView;
        internal bool goToArchive = false;

        public object CurrentView
        {
            get { return _currentView; }
            set
            {
                var documentVM = _currentView as ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel.DocumentsVM;
                if (documentVM != null)
                {
                    documentVM.threadFillList.Abort();
                }

                if (value is LoginVM loginVM)
                {
                    loginVM.GoToMain = () =>
                    {
                        GoToMain.Invoke();
                    };
                    loginVM.GoToDocuments = () =>
                    {
                        GoToDocuments.Invoke();
                    };
                }

                if (value is ProfileVM profileVM)
                {
                    profileVM.GoToMain = () =>
                    {
                        GoToMain.Invoke();
                    };
                    profileVM.GoToDocuments = () =>
                    {
                        GoToDocuments.Invoke();
                    };
                    profileVM.GoToLogin = () =>
                    {
                        GoToLogin.Invoke();
                    };
                }

                if (value is BugReportVM bugReportVM)
                {
                    bugReportVM.GoToMain = () =>
                    {
                        GoToMain.Invoke();
                    };
                    bugReportVM.GoToDocuments = () =>
                    {
                        GoToDocuments.Invoke();
                    };
                    bugReportVM.GoToLogin = () =>
                    {
                        GoToLogin.Invoke();
                    };
                }

                if (value is AboutUsVM aboutUsVM)
                {
                    aboutUsVM.GoToMain = () =>
                    {
                        GoToMain.Invoke();
                    };
                    aboutUsVM.GoToDocuments = () =>
                    {
                        GoToDocuments.Invoke();
                    };
                    aboutUsVM.GoToLogin = () =>
                    {
                        GoToLogin.Invoke();
                    };
                }

                if (value is MainVM mainVM)
                {
                    mainVM.GoToMain = () =>
                    {
                        GoToMain.Invoke();
                    };
                    mainVM.GoToDocuments = () =>
                    {
                        GoToDocuments.Invoke();
                    };
                    mainVM.GoToLogin = () =>
                    {
                        GoToLogin.Invoke();
                    };
                }


                if (value is NetworkingVM networkingVM)
                {
                    networkingVM.GoToMain = () =>
                    {
                        GoToMain.Invoke();
                    };
                    networkingVM.GoToDocuments = () =>
                    {
                        GoToDocuments.Invoke();
                    };
                    networkingVM.GoToLogin = () =>
                    {
                        GoToLogin.Invoke();
                    };
                }


                _currentView = value;
                OnPropertyChanged();
                CurrentViewChanged?.Invoke();
            }
        }
        public ICommand AboutUsCommand { get; set; }
        public ICommand LoginCommand { get; set; }
        public ICommand AllDocumentsCommand { get; set; }
        public ICommand ArchiveDocumentsCommand { get; set; }
        public ICommand SettingsCommand { get; set; }
        public ICommand ProfileCommand { get; set; }
        public ICommand BugReportCommand { get; set; }
        public ICommand MainCommand { get; set; }
        public ICommand NetworkingCommand { get; set; }

        private void AboutUs(object value)
        {
            CurrentView = new AboutUsVM();
        }
        private void Login(object value)
        {
            CurrentView = new LoginVM();
        }
        private void AllDocuments(object value)
        {
            goToArchive = false;
            CurrentView = new DocumentsVM(false);
        }
        private void ArchiveDocuments(object value)
        {
            goToArchive = true;
            CurrentView = new DocumentsVM(true);
        }
        private void Settings(object value)
        {
            CurrentView = new SettingsVM();
        }
        private void Profile(object value)
        {
            CurrentView = new ProfileVM();
        }
        private void BugReport(object value)
        {
            CurrentView = new BugReportVM();
        }
        private void Main(object value)
        {
            CurrentView = new MainVM();
        }
        private void Networking(object value)
        {
            CurrentView = new NetworkingVM();
        }
    }
}
