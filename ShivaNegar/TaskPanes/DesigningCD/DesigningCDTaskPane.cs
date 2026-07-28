using System;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using ShivaNegar.Constants;

namespace ShivaNegar.TaskPanes.DesigningCD
{
    interface ITaskPaneRequests
    {
        Action CloseTaskPaneRequest { get; set; }
    }
    partial class DesigningCDTaskPane : UserControl, ITaskPaneRequests
    {
        private CustomTaskPane taskPane;

        internal int sideDocumentActionSize = 34;
        internal int defaultWidth = 300;

        internal Range previousRange;
        internal Document doc;

        public Action CloseTaskPaneRequest { get; set; }

        public DesigningCDTaskPane(Document doc, Range previousRange)
        {
            InitializeComponent();
            defaultWidth = this.Width;
            this.Paint += ActionPaneDesigningCD_Paint;
            this.Dock = DockStyle.Fill;

            this.doc = doc;
            this.previousRange = previousRange;

            userControl.CloseTaskPaneRequest += () =>
            {
                DeleteAndDisableActionPane(doc, taskPane, previousRange);
            };
        }

        #region reVisibleActionsPane
        private void ActionPaneDesigningCD_Paint(object sender, PaintEventArgs e)
        {
            System.Threading.Tasks.Task.Run(() => { reVisibleActionsPane(); });
        }

        internal void reVisibleActionsPane()
        {
            Thread.Sleep(500);
            if (!taskPane.Visible)
            {
                bool shapeIsExist = false;
                try
                {
                    int getID = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].ID;
                    shapeIsExist = true;
                }
                catch (Exception) { }

                if (shapeIsExist)
                {
                    taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                    taskPane.Width = defaultWidth + sideDocumentActionSize;
                    taskPane.Visible = true;
                }
            }
        }
        #endregion

        internal static void disableActionPane(CustomTaskPane taskPane)
        {
            try
            {
                if (taskPane != null)
                    Globals.ThisAddIn.CustomTaskPanes.Remove(taskPane);
            }
            catch (Exception)
            {
                //DedicatedFunctions.ShowErrorMessage("در بستن منو ارجاع دادن مشکلی پیش آمده است" , (int)ErrorCodes.ActionsPaneCrossReferenceDisabling , StringConstant.SupportEmail);
            }
        }
        internal static void DeleteAndDisableActionPane(Document doc, CustomTaskPane taskPane, Range previousRange)
        {
            try
            {
                doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].Select();
                doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].Delete();

                int sectionIndex = doc.ActiveWindow.Selection.Sections.Last.Index;
                Section cdSection = doc.Sections.Last;
                if (doc.Sections.Count >= 2)
                {
                    Section previousSection = doc.Sections[doc.Sections.Count - 1];


                    try
                    {
                        cdSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = previousSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious;
                        cdSection.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = previousSection.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious;
                        cdSection.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = previousSection.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious;
                        //cdSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = previousSection.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;
                        //cdSection.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Text = previousSection.Headers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Text;
                        //cdSection.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = previousSection.Headers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text;
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        cdSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = previousSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious;
                        cdSection.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious = previousSection.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].LinkToPrevious;
                        cdSection.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious = previousSection.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].LinkToPrevious;
                        //cdSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text = previousSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text;
                        //cdSection.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Text = previousSection.Footers[WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range.Text;
                        //cdSection.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text = previousSection.Footers[WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range.Text;
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        cdSection.PageSetup.BottomMargin = previousSection.PageSetup.BottomMargin;
                        cdSection.PageSetup.LeftMargin = previousSection.PageSetup.LeftMargin;
                        cdSection.PageSetup.RightMargin = previousSection.PageSetup.RightMargin;
                        cdSection.PageSetup.TopMargin = previousSection.PageSetup.TopMargin;
                        cdSection.PageSetup.HeaderDistance = previousSection.PageSetup.HeaderDistance;
                        cdSection.PageSetup.FooterDistance = previousSection.PageSetup.FooterDistance;
                    }
                    catch (Exception)
                    {
                    }
                }

                DedicatedFunctions.deleteSpecificSection(doc, sectionIndex);
                doc.ActiveWindow.Selection.MoveEnd(WdUnits.wdParagraph, 1);
                doc.ActiveWindow.Selection.MoveStart(WdUnits.wdCharacter, -1);
                doc.ActiveWindow.Selection.Delete();
                DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection);
            }
            catch (Exception) { }

            disableActionPane(taskPane);

            if (previousRange != null)
            {
                previousRange.Select();
                doc.ActiveWindow.ScrollIntoView(previousRange);
            }
            previousRange = null;
            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
        internal void enableActionPane(CustomTaskPane _taskPane)
        {
            try
            {
                taskPane = _taskPane;
                taskPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoHorizontal;
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
                taskPane.Width = defaultWidth + sideDocumentActionSize;
                taskPane.Visible = true;

                userControl.initializeUserControl(doc);
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage(" منو کناری طراحی لوح فشرده از کارافتاده است\n برای رفع مشکل برنامه، سند را بسته و دوباره بازنمایید", (int)ErrorCodes.ActionsPaneCDEnabling, StringConstant.SupportEmail);
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("لغو خروجی گرفتن از طرح سیدی");
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                DeleteAndDisableActionPane(doc, taskPane, previousRange);
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
    }
}
