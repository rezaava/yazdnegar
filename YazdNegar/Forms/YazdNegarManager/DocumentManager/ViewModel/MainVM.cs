using System;
using YazdNegar.Forms.YazdNegarManager.DocumentManager.Utilities;

namespace YazdNegar.Forms.YazdNegarManager.DocumentManager.ViewModel
{
    public class MainVM : ViewModelBase, IGoToDocumentManager
    {
        public Action GoToMain { get; set; }
        public Action GoToDocuments { get; set; }
        public Action GoToLogin { get; set; }

        public MainVM()
        {
        }
    }
}
