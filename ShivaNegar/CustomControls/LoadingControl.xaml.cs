using System.Windows.Controls;

namespace ShivaNegar.CustomControls
{
    /// <summary>
    /// Interaction logic for LoadingModel.xaml
    /// </summary>
    public partial class LoadingControl : UserControl
    {
        public LoadingControl()
        {
            InitializeComponent();
        }

        public void changeSizeLoading(double size)
        {
            progress.Width = size;
            progress.Height = size;
        }
        public void setProgressValue(int value)
        {
            progress.Value = value;
        }
        public void setMessage(string message)
        {
            lblMessage.Text = message;
        }
    }
}
