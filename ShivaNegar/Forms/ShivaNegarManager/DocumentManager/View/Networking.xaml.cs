using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View
{
    /// <summary>
    /// Interaction logic for Profile.xaml
    /// </summary>
    public partial class Networking : System.Windows.Controls.UserControl
    {

        public Networking()
        {
            InitializeComponent();

            btnRefreshList.Click += BtnRefreshList_Click;
        }

        private void BtnRefreshList_Click(object sender, RoutedEventArgs e)
        {
            ((NetworkingVM)this.DataContext).getNetworkingRequests();
        }

        private void btnRequestAccept_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                var task = button.DataContext as NetworkingRequestModel;
                Dispatcher.Invoke(() =>
                {
                    ((NetworkingVM)this.DataContext).setStatus(task, 1);
                });
            }
            else
            {
                return;
            }
        }

        private void btnRequestReject_Click(object sender, RoutedEventArgs e)
        {
            var button = sender as Button;
            if (button != null)
            {
                var task = button.DataContext as NetworkingRequestModel;
                Dispatcher.Invoke(() =>
                {
                    ((NetworkingVM)this.DataContext).setStatus(task, 0);
                });
            }
            else
            {
                return;
            }
        }

        #region TextBox Event for Control
        private void PreventTypeSpace_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // اگر کلید Space فشرده شود
            if (e.Key == Key.Space)
            {
                e.Handled = true; // جلوگیری از ورود Space
            }
        }
        private void PhoneNumberTextBox_PreviewTextInput(object sender, System.Windows.Input.TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9]+");

            //((TextBox)sender).Text = ((TextBox)sender).Text.Replace(" " , "");
            //((TextBox)sender).
            e.Handled = regex.IsMatch(e.Text);
        }
        #endregion

    }
}
