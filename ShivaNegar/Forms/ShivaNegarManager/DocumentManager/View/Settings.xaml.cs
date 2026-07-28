using System.Windows.Forms;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View
{
    /// <summary>
    /// Interaction logic for Settings.xaml
    /// </summary>
    public partial class Settings : System.Windows.Controls.UserControl
    {
        public Settings()
        {
            InitializeComponent();

            btnChangeWorkspace.Click += BtnChangeWorkspace_Click;
            txtWorkspace.Text = Properties.Settings.Default.WorkSpaceDirectory;
        }

        private void BtnChangeWorkspace_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.SelectedPath = Properties.Settings.Default.WorkSpaceDirectory;
            dialog.ShowNewFolderButton = true;
            DialogResult result = dialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                txtWorkspace.Text = dialog.SelectedPath.TrimEnd('\\') + "\\";

                Properties.Settings.Default.WorkSpaceDirectory = dialog.SelectedPath.TrimEnd('\\') + "\\";
                Properties.Settings.Default.Save();

                //DialogResult closeDialog = DedicatedFunctions.ShowMessage("اسناد موجود به پوشه جدید انتقال پیدا کند؟",
                //    StringConstant.NameOfProject,
                //    MessageBoxButtons.YesNoCancel,
                //    MessageBoxIcon.Warning,
                //    MessageBoxDefaultButton.Button1,
                //    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                //if (closeDialog == DialogResult.Yes)
                //{
                //
                //}

            }
        }
    }
}
