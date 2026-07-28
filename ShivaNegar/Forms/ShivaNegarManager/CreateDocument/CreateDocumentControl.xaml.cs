using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Threading;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View;

namespace ShivaNegar.Forms.ShivaNegarManager.CreateDocument
{
    public partial class CreateDocumentControl : System.Windows.Controls.UserControl, Interfaces.IStatusFormRequest, Interfaces.IChangeTransitionDocumentManager, INotifyPropertyChanged
    {
        //implement interface
        public Action TransitionDocumentManagerRequest { get; set; }

        public Action CloseFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action NormalStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }

        //Transition Variables
        private int previousSelectedTransition = 0;
        private bool isTransitionMovementForward = false;


        private int m_progress;
        public int Progress
        {
            get { return m_progress; }
            set
            {
                m_progress = value;
                OnPropertyChanged("Progress");
            }
        }

        public ObservableCollection<string> Steps
        {
            get;
            set;
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public CreateDocumentControl()
        {
            InitializeComponent();

            Steps = new ObservableCollection<string>();
            Steps.Add("بسم الله");
            Steps.Add("مشخصات سند");
            Steps.Add("مشخصات تحصیلی");

            Constants.DocumentTypes documentType = DedicatedFunctions.getDocumentType(createDocumentSlide2.comboDocumentType.SelectedItem as string);
            string documentTypeString = DedicatedFunctions.getDocumentTypePersianName((int)documentType);
            Steps.Add("مشخصات " + documentTypeString);

            Steps.Add("تاییدیه");
            Steps.Add("ساخت سند");
            Progress = 1;
            DataContext = this;

            btnExitCreateDocument.Click += BtnExitCreateDocument_Click;
            toggleButtonAlwaysOnTop.Click += ToggleButtonAlwaysOnTop_Click;
            btnMinimize.Click += (b1, e) => { MinimizeStateFormRequest?.Invoke(); };
            btnMaximize.Click += (b1, e) => { MaximizeStateFormRequest?.Invoke(); };

            transitionCreateDocument.SelectionChanged += TransitionCreateDocument_SelectionChanged;

            //progressPage.Value = 0;
            //lblProgressPage.Text = "%0";
            //transitionCreateDocument.SelectedIndex = 5;//Test //Page 6 

            createDocumentSlide2.comboDocumentType.SelectionChanged += (s, a) =>
            {
                if (createDocumentSlide2.comboDocumentType.SelectedIndex != -1)
                {
                    documentType = DedicatedFunctions.getDocumentType(createDocumentSlide2.comboDocumentType.SelectedItem as string);
                    documentTypeString = DedicatedFunctions.getDocumentTypePersianName((int)documentType);
                    Steps[3] = "مشخصات " + documentTypeString;
                }
            };

            createDocumentSlide5.TransitionMoveNextCommand += () =>
            {
                transitionCreateDocument.SelectedIndex++;
            };
            createDocumentSlide6.TransitionDocumentManagerRequest += () =>
            {
                #region pages reset
                createDocumentSlide1.resetControls();
                createDocumentSlide2.resetControls();
                createDocumentSlide3.resetControls();
                createDocumentSlide4.resetControls();
                createDocumentSlide5.resetControls();
                createDocumentSlide6.resetControls();
                #endregion

                Dispatcher.Invoke(() =>
                {
                    transitionCreateDocument.SelectedIndex = 0;
                });
                this.TransitionDocumentManagerRequest?.Invoke();
            };
            createDocumentSlide6.CloseForm += () =>
            {
                CloseFormRequest?.Invoke();
            };


            string versionCustomized = BugReport.AssemblyVersion;
            versionCustomized = versionCustomized.Replace("0", "۰").Replace("1", "۱").Replace("2", "۲").Replace("3", "۳").Replace("4", "۴").Replace("5", "۵")
                .Replace("6", "۶").Replace("7", "۷").Replace("8", "۸").Replace("9", "۹");

            lblVersion.Text = "نسخه: " + versionCustomized;
        }

        private void ToggleButtonAlwaysOnTop_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (toggleButtonAlwaysOnTop.IsChecked == true)
                AlwaysOnTopEnableRequest?.Invoke();
            else
                AlwaysOnTopDisableRequest?.Invoke();
        }

        #region Events
        public void BtnExitCreateDocument_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (transitionCreateDocument.SelectedIndex == 5)
            {
                createDocumentSlide6.close();
                return;
            }
            #region pages reset
            createDocumentSlide1.resetControls();
            createDocumentSlide2.resetControls();
            createDocumentSlide3.resetControls();
            createDocumentSlide4.resetControls();
            createDocumentSlide5.resetControls();
            createDocumentSlide6.resetControls();
            #endregion

            transitionCreateDocument.SelectedIndex = 0;
            TransitionDocumentManagerRequest?.Invoke();
        }

        private void TransitionCreateDocument_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            if (transitionCreateDocument.SelectedIndex != previousSelectedTransition)//changed selected Page
            {
                if (transitionCreateDocument.SelectedIndex > previousSelectedTransition)
                    isTransitionMovementForward = true;
                else
                    isTransitionMovementForward = false;
            }
            else
                return;

            float step = 100f / (float)transitionCreateDocument.Items.Count;

            Progress = (int)step * (transitionCreateDocument.SelectedIndex + 1);

            //lblProgressPage.Text = "%" + (step * (transitionCreateDocument.SelectedIndex + 1)).ToString();
            //progressPage.Value = step * (transitionCreateDocument.SelectedIndex + 1);

            if (transitionCreateDocument.SelectedIndex == -1)//not selected any Page
            {
                previousSelectedTransition = transitionCreateDocument.SelectedIndex;
                return;
            }
            else if (transitionCreateDocument.SelectedIndex == 2)//page 3
            {
                if (isTransitionMovementForward)
                    createDocumentSlide3.initializeVariables(createDocumentSlide2.DocumentType);
            }
            else if (transitionCreateDocument.SelectedIndex == 3)//page 4
            {
                createDocumentSlide4.initializeVariables(createDocumentSlide2.DocumentType);
            }
            else if (transitionCreateDocument.SelectedIndex == 4)//page 6
            {
                createDocumentSlide5.initializeVariables(createDocumentSlide2, createDocumentSlide3, createDocumentSlide4);
            }
            else if (transitionCreateDocument.SelectedIndex == 5)//page 7
            {
                createDocumentSlide6.initializeVariables(createDocumentSlide1, createDocumentSlide2, createDocumentSlide3, createDocumentSlide4);
            }

            previousSelectedTransition = transitionCreateDocument.SelectedIndex;
        }
        #endregion

        #region Functions
        #endregion
    }

}
