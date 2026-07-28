using System;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Utilities;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.ViewModel
{
    public class AboutUsVM : ViewModelBase, IGoToDocumentManager
    {
        public Action GoToMain { get; set; }
        public Action GoToDocuments { get; set; }
        public Action GoToLogin { get; set; }

        public AboutUsVM()
        {
        }

        internal void goBack(bool isUserLogined)
        {
            if (isUserLogined)
                GoToMain.Invoke();
            else
                GoToLogin.Invoke();
        }
    }
}
