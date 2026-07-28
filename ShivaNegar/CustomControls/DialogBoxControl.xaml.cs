using System.Windows.Controls;

namespace ShivaNegar.CustomControls
{
    /// <summary>
    /// Interaction logic for LoadingModel.xaml
    /// </summary>
    public partial class DialogBoxControl : UserControl
    {
        public DialogBoxControl()
        {
            InitializeComponent();

            btnConfirm.Click += BtnConfirm_Click;
        }

        private void BtnConfirm_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            this.IsEnabled = false;
        }
    }
}
