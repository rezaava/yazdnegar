using System.Windows.Controls;
using System.Windows.Threading;
using ShivaNegar.Constants.ComboBoxData;

namespace ShivaNegar.Forms.ShivaNegarManager.CreateDocument
{
    public partial class CreateDocumentSlide1 : UserControl
    {
        //Properties
        public string NameOfAllah { get; private set; }
        public int NameOfAllahFontType { get; private set; }

        //validate Variables
        private bool validateNameOfAllah = false;

        string besmellahFontNameInUI1 = "B e s m e l l a h 1";
        string besmellahFontNameInUI2 = "B e s m e l l a h 2";
        string besmellahFontNameInUI3 = "B e s m e l l a h 3";
        //string besmellahFontNameInUI4 = "B e s m e l l a h 4";

        public CreateDocumentSlide1()
        {
            InitializeComponent();

            //reset
            resetControls();

            lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
            lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
            NameOfAllahFontType = 1;

            comboBesmellahFontType.SelectionChanged += ComboBesmellahFontType_SelectionChanged;
            lstBoxNameOfAllah.SelectionChanged += LstBoxNameOfAllah_SelectionChanged;

            comboBesmellahFontType.Items.Clear();
            comboBesmellahFontType.Items.Add("قالب نوع اول");
            comboBesmellahFontType.Items.Add("قالب نوع دوم");
            comboBesmellahFontType.Items.Add("قالب نوع سوم");
            //comboBesmellahFontType.Items.Add("قالب نوع چهارم");//Disabled

            comboBesmellahFontType.SelectedIndex = 0;

            lstBoxNameOfAllah.Items.Clear();
            lstBoxNameOfAllah.ItemsSource = ComboBoxData.InTheNameOfAllah;

            lstBoxNameOfAllah.SelectedIndex = 0;
            lstBoxNameOfAllah.ScrollIntoView(lstBoxNameOfAllah.SelectedItem);
            lstBoxNameOfAllah.Focus();
        }

        private void ComboBesmellahFontType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            NameOfAllahFontType = comboBesmellahFontType.SelectedIndex + 1;

            lstBoxNameOfAllah.SelectedIndex = 0;
            lstBoxNameOfAllah.ScrollIntoView(lstBoxNameOfAllah.SelectedItem);
            lstBoxNameOfAllah.Focus();

            if (comboBesmellahFontType.SelectedIndex == 0)
            {
                lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
                lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
            }
            else if (comboBesmellahFontType.SelectedIndex == 1)
            {

                lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI2);
                lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI2);
            }
            else if (comboBesmellahFontType.SelectedIndex == 2)
            {
                lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI3);
                lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI3);
            }
            //else if(comboBesmellahFontType.SelectedIndex == 3)
            //{
            //	lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI4);
            //	lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI4);
            //}
        }


        #region Events
        private void LstBoxNameOfAllah_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            validateNameOfAllah = validateListBoxNameOfAllah();
            validateControls();
        }
        #endregion

        #region Validators
        private bool validateControls()
        {
            if (validateNameOfAllah)
            {
                lblNameOfAllah.Text = lstBoxNameOfAllah.SelectedItem as string;
                NameOfAllah = lblNameOfAllah.Text;
                btnGoCreateDocument.IsEnabled = true;
                return true;
            }
            else
            {
                lblNameOfAllah.Text = "";
                NameOfAllah = "";
                btnGoCreateDocument.IsEnabled = false;
                return false;
            }
        }
        private bool validateListBoxNameOfAllah()
        {
            if (lstBoxNameOfAllah.SelectedIndex != -1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }


        #endregion

        #region Functions
        internal void resetControls()
        {
            Dispatcher.Invoke(() =>
            {
                #region variables
                validateNameOfAllah = false;

                NameOfAllah = "";
                #endregion

                #region controls
                //btnGoCreateDocument.IsEnabled = false;
                Dispatcher.Invoke(() =>
                {
                    lstBoxNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
                    lblNameOfAllah.FontFamily = new System.Windows.Media.FontFamily(besmellahFontNameInUI1);
                    NameOfAllahFontType = 1;
                    comboBesmellahFontType.SelectedIndex = 0;

                    lstBoxNameOfAllah.SelectedIndex = 0;
                    lstBoxNameOfAllah.ScrollIntoView(lstBoxNameOfAllah.SelectedItem);
                    lstBoxNameOfAllah.Focus();
                });
                #endregion
            });
        }
        #endregion
    }
}
