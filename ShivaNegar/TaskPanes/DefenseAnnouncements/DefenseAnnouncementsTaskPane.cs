using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace ShivaNegar.TaskPanes.DefenseAnnouncements
{

    partial class DefenseAnnouncementsTaskPane : UserControl
    {
        private CustomTaskPane taskPane;

        internal int sideDocumentActionSize = 34;
        internal int defaultWidth = 350;
        internal DefenseAnnouncementsTaskPane()
        {
            InitializeComponent();
            defaultWidth = this.Width;
            this.Dock = DockStyle.Fill;

            defenseAnnouncementsControl.btnClose.Click += BtnClose_Click;
        }

        private void BtnClose_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            disableActionPane();
        }

        internal void disableActionPane()
        {
            try
            {
                //txtSearchCitation.Clear();

                //if (taskPane.Visible)
                //    taskPane.Visible = false;
                Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
            }
            catch (Exception)
            {
                //DedicatedFunctions.ShowErrorMessage("در بستن منو ارجاع دادن مشکلی پیش آمده است" , (int)ErrorCodes.ActionsPaneCrossReferenceDisabling , StringConstant.SupportEmail);
            }
        }
        internal void enableActionPane(CustomTaskPane _taskPane)
        {
            try
            {
                taskPane = _taskPane;
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Width = defaultWidth + sideDocumentActionSize;
                taskPane.Visible = true;

                defenseAnnouncementsControl.initializeDefenseAnnouncementControl();
            }
            catch (Exception)
            {
                //DedicatedFunctions.ShowErrorMessage(" منو درج منبع از کارافتاده است\n برای رفع مشکل برنامه، سند را بسته و دوباره بازنمایید" + "\nپیغام خطا:\n" + e.Message , (int)ErrorCodes.ActionsPaneCrossReferenceEnabling , StringConstant.SupportEmail);
                disableActionPane();
            }
        }
    }
}
