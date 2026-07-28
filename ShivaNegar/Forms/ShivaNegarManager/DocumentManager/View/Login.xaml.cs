using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MaterialDesignThemes.Wpf;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View
{
    /// <summary>
    /// Interaction logic for Login.xaml
    /// </summary>
    public partial class Login : System.Windows.Controls.UserControl
    {
        public Login()
        {
            InitializeComponent();

            btnBackToPassword.Click += BtnBackToPassword_Click;
            btnBackToLogin.Click += BtnBackToLogin_Click;
            btnConfirmSendPassword.Click += BtnBackToLogin_Click;
            btnConfirmSendRegister.Click += BtnBackToLogin_Click;

            btnLogin.Click += BtnLogin_Click;
            btnLoginConfirm.Click += BtnLoginConfirm_Click;
            btnChangePassword.Click += BtnChangePassword_Click;
            btnRequestChangePassword.Click += BtnRequestChangePassword_Click;

            txtMobile.PreviewTextInput += PhoneNumberTextBox_PreviewTextInput;
            txtMobile.PreviewKeyDown += PreventTypeSpace_PreviewKeyDown;
            //txtMobile.LostFocus += PhoneNumberTextBox_LostFocus;

            txtMobileForChangePassword.PreviewTextInput += PhoneNumberTextBox_PreviewTextInput;
            txtMobileForChangePassword.PreviewKeyDown += PreventTypeSpace_PreviewKeyDown;
            //txtMobileForChangePassword.LostFocus += PhoneNumberTextBox_LostFocus;

            txtPassword.PreviewKeyDown += PreventTypeSpace_PreviewKeyDown;
            //txtPassword.LostFocus += TxtPassword_LostFocus;

            lblHyperLinkAgreement.Click += LblHyperLinkAgreement_Click;


        }

        private void LblHyperLinkAgreement_Click(object sender, RoutedEventArgs e)
        {
            //Show rules
            dialogBox.Tag = "قوانین ایجاد سند" + "\n\n\n" + "استفاده از این نرم افزار برای انجام هرگونه فعالیت غیر قانونی ممنوع است. این شامل مواردی مانند تبلیغات غیرمجاز، خرید و فروش غیرقانونی و سایر موارد ممنوعه طبق قوانین ایران می باشد " + "\n\n" + "مالکیت سند ایجاد شده متعلق به صاحب حساب " + txtMobile.Text + " می باشد که بدین وسیله متعهد می شود از قرار دادن سند، بخشی از سند و یا قالب آن به شخص ثالث خودداری کند." + "\n\n" + "این قوانین به منظور حفظ امنیت، حریم خصوصی و رعایت قوانین جاری در جمهوری اسلامی ایران تدوین شده اند. عدم رعایت این قوانین می تواند منجر به تعلیق یا محدودیت دسترسی کاربر به این نرم افزار شود.";
            dialogBox.IsEnabled = true;
        }

        #region Validates
        private bool validateMobile(TextBox textBox)
        {
            if (string.IsNullOrEmpty(textBox.Text) || string.IsNullOrWhiteSpace(textBox.Text))
            {
                HintAssist.SetHelperText(textBox, "شماره موبایل نمی‌تواند خالی باشد.");
                textBox.Foreground = System.Windows.Media.Brushes.Red;
                textBox.Margin = new Thickness(0, 20, 0, 20);

                return false;
            }
            else if (!System.Text.RegularExpressions.Regex.IsMatch(textBox.Text, @"^\d+$") || textBox.Text.Length < 11 || !textBox.Text.StartsWith("09"))
            {
                HintAssist.SetHelperText(textBox, "شماره موبایل نادرست است");
                textBox.Foreground = System.Windows.Media.Brushes.Red;
                textBox.Margin = new Thickness(0, 20, 0, 20);

                return false;
            }
            else
            {
                HintAssist.SetHelperText(textBox, "");
                textBox.Foreground = System.Windows.Media.Brushes.Black;
                textBox.Margin = new Thickness(0, 20, 0, 10);
                return true;
            }
        }
        private bool validatePassword(PasswordBox textBox)
        {
            if (string.IsNullOrEmpty(textBox.Password) || string.IsNullOrWhiteSpace(textBox.Password))
            {
                HintAssist.SetHelperText(textBox, "رمز عبور نمی‌تواند خالی باشد.");
                textBox.Foreground = System.Windows.Media.Brushes.Red;
                textBox.Margin = new Thickness(0, 20, 0, 20);

                return false;
            }
            else
            {
                HintAssist.SetHelperText(textBox, "");
                textBox.Foreground = System.Windows.Media.Brushes.Black;
                textBox.Margin = new Thickness(0, 20, 0, 10);
                return true;
            }
        }

        //private void TxtPassword_LostFocus(object sender , RoutedEventArgs e)
        //{
        //	validatePassword(sender as PasswordBox);
        //}
        //private void PhoneNumberTextBox_LostFocus(object sender , System.Windows.RoutedEventArgs e)
        //{
        //	validateMobile(sender as TextBox);
        //}
        #endregion

        #region TextBox Event for Control
        private void PreventTypeSpace_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            // اگر کلید Space فشرده شود
            if (e.Key == Key.Space)
            {
                e.Handled = true; // جلوگیری از ورود Space
            }

            if (e.Key == System.Windows.Input.Key.Enter)
            {
                if (sender is TextBox textBox)
                {
                    if (textBox.Name == txtMobile.Name)
                        login();
                    else if (textBox.Name == txtMobileForChangePassword.Name)
                        requestChangePassword();
                }
                else if (sender is PasswordBox passwordBox)
                {
                    if (passwordBox.Name == txtPassword.Name)
                        loginConfirm();
                }
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

        private void TxtHyperLinkAgreement_Click(object sender, RoutedEventArgs e)
        {
            //Show rules
            dialogBox.Tag = "قوانین ایجاد سند" + "\n\n\n" + "استفاده از این نرم افزار برای انجام هرگونه فعالیت غیر قانونی ممنوع است. این شامل مواردی مانند تبلیغات غیرمجاز، خرید و فروش غیرقانونی و سایر موارد ممنوعه طبق قوانین ایران می باشد " + "\n\n" + "مالکیت سند ایجاد شده متعلق به صاحب حساب " + txtMobile.Text + " می باشد که بدین وسیله متعهد می شود از قرار دادن سند، بخشی از سند و یا قالب آن به شخص ثالث خودداری کند." + "\n\n" + "این قوانین به منظور حفظ امنیت، حریم خصوصی و رعایت قوانین جاری در جمهوری اسلامی ایران تدوین شده اند. عدم رعایت این قوانین می تواند منجر به تعلیق یا محدودیت دسترسی کاربر به این نرم افزار شود.";
            dialogBox.IsEnabled = true;
        }

        #region Dynamic and Main Navigators

        private void BtnLogin_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            login();
        }
        private void BtnLoginConfirm_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            loginConfirm();
        }
        private void BtnRequestChangePassword_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            requestChangePassword();
        }
        #endregion


        #region Functions
        private void login()
        {
            if (validateMobile(txtMobile))
            {
                Dispatcher.Invoke(() =>
                {
                    ((LoginVM)DataContext).checkMobileAndRegister(txtMobile, txtPassword, transitionLoginRegister);
                });

                //lblMobileInRegister.Text = txtMobile.Text;
            }
        }

        private void loginConfirm()
        {
            if (validatePassword(txtPassword))
            {
                Dispatcher.Invoke(() => { ((LoginVM)DataContext).loginRequest(txtMobile, txtPassword); });
                Properties.Settings.Default.RememberUser = chkRememberUser.IsChecked == true ? true : false;
                Properties.Settings.Default.Save();
            }
        }

        private void requestChangePassword()
        {
            if (validateMobile(txtMobileForChangePassword))
            {
                Dispatcher.Invoke(() => { ((LoginVM)DataContext).forgetPasswordRequest(txtMobileForChangePassword, transitionLoginRegister); });
                lblMobileInChangePassword.Text = txtMobileForChangePassword.Text;
            }
        }
        #endregion


        #region Static Navigators
        private void BtnChangePassword_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            transitionLoginRegister.SelectedIndex = LoginVM.changePasswordPageIndex;
            txtMobileForChangePassword.Text = txtMobile.Text;
            txtMobileForChangePassword.Focus();
        }

        private void BtnBackToPassword_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            transitionLoginRegister.SelectedIndex = LoginVM.passwordPageIndex;

            HintAssist.SetHelperText(txtMobileForChangePassword, "");
            txtMobileForChangePassword.Foreground = System.Windows.Media.Brushes.Black;
            txtMobileForChangePassword.Margin = new Thickness(0, 20, 0, 10);
            txtPassword.Focus();
        }

        private void BtnBackToLogin_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            transitionLoginRegister.SelectedIndex = LoginVM.loginPageIndex;
            txtPassword.Password = "";
            HintAssist.SetHelperText(txtPassword, "");
            txtPassword.Foreground = System.Windows.Media.Brushes.Black;
            txtPassword.Margin = new Thickness(0, 20, 0, 10);
            txtMobile.Focus();
        }
        #endregion


    }
}
