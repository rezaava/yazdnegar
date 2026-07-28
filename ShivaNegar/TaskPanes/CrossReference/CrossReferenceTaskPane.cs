using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;

namespace ShivaNegar.TaskPanes.CrossReference
{

    partial class CrossReferenceTaskPane : UserControl
    {
        private CustomTaskPane taskPane;

        internal int sideDocumentActionSize = 34;
        internal int defaultWidth = 350;
        internal CrossReferenceTaskPane()
        {
            InitializeComponent();
            defaultWidth = this.Width;
            this.Dock = DockStyle.Fill;

            crossReferenceControl.btnClose.Click += BtnClose_Click;
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
        internal void enableActionPane(CustomTaskPane _taskPane, Document doc)
        {
            try
            {
                taskPane = _taskPane;
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Width = defaultWidth + sideDocumentActionSize;
                taskPane.Visible = true;

                crossReferenceControl.useCrossReference(doc);
            }
            catch (Exception)
            {
                //DedicatedFunctions.ShowErrorMessage(" منو درج منبع از کارافتاده است\n برای رفع مشکل برنامه، سند را بسته و دوباره بازنمایید" + "\nپیغام خطا:\n" + e.Message , (int)ErrorCodes.ActionsPaneCrossReferenceEnabling , StringConstant.SupportEmail);
                disableActionPane();
            }
        }
    }
    internal class CaptionObject
    {
        private string captionType;
        private string captionText;

        internal CaptionObject(string captionType, string captionText)
        {
            this.captionType = captionType;
            this.captionText = captionText;
        }

        internal string CaptionType { get => captionType; set => captionType = value; }
        internal string CaptionText { get => captionText; set => captionText = value; }
    }
}
