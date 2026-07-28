using System;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Media;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using static ShivaNegar.DedicatedFunctions;

namespace ShivaNegar.Forms.CaptionSettings
{
    /// <summary>
    /// Interaction logic for ChangeContentsControl.xaml
    /// </summary>
    public partial class CaptionSettingsControl : System.Windows.Controls.UserControl, Interfaces.IStatusFormRequest
    {
        public Document doc;
        public Action CloseFormRequest { get; set; }
        public Action MinimizeStateFormRequest { get; set; }
        public Action MaximizeStateFormRequest { get; set; }
        public Action AlwaysOnTopEnableRequest { get; set; }
        public Action AlwaysOnTopDisableRequest { get; set; }

        #region constructor
        public CaptionSettingsControl(Document doc)
        {
            InitializeComponent();

            this.doc = doc;

            btnCloseApp.Click += BtnCloseApp_Click;
            //btnMaximize.Click += BtnMaximize_Click;
            //btnMinimize.Click += BtnMinimize_Click;

            tbtnShowDialogCaptionFigure.IsChecked = Properties.Settings.Default.CaptionSettings_ShowDialogCaptionFigure;
            tbtnShowDialogCaptionFigure.Click += TbtnShowDialogCaptionFigure_Click;


            btnChangeShapeSizeToContentArea.Click += BtnChangeShapeSizeToContentArea_Click;
            btnChangeTableSizeToPage.Click += BtnChangeTableSizeToPage_Click;
            btnChangeTableSizeToContentArea.Click += BtnChangeTableSizeToContentArea_Click;
            btnRotateTableToRight.Click += BtnRotateTableToRight_Click;
        }

        #endregion

        #region Events
        private void BtnMinimize_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            MinimizeStateFormRequest?.Invoke();
        }

        private void BtnMaximize_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            MaximizeStateFormRequest?.Invoke();
        }

        private void BtnCloseApp_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            CloseFormRequest?.Invoke();
        }

        private void BtnChangeShapeSizeToContentArea_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.ShapeRange.Count == 0 && Globals.ThisAddIn.Application.Selection.InlineShapes.Count == 0)
            {
                if (DedicatedFunctions.ShowMessage("شکلی انتخاب نشده؛ در کل شکل ها اعمال شود؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    changeShapesSize(Globals.ThisAddIn.Application.ActiveDocument, WidthType.Margin, (bool)tbtnKeepScaleShape.IsChecked, (bool)!tbtnNotOverlap.IsChecked);
            }
            else
            {
                changeShapesSize(Globals.ThisAddIn.Application.ActiveDocument, Globals.ThisAddIn.Application.Selection, WidthType.Margin, (bool)tbtnKeepScaleShape.IsChecked, (bool)!tbtnNotOverlap.IsChecked);
            }
        }

        private void BtnChangeTableSizeToPage_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Tables.Count == 0)
            {
                if (DedicatedFunctions.ShowMessage("جدولی انتخاب نشده؛ در کل جداول اعمال شود؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    changeTableSize(Globals.ThisAddIn.Application.ActiveDocument, WidthType.Page);
            }
            else
            {
                changeTableSize(Globals.ThisAddIn.Application.ActiveDocument, Globals.ThisAddIn.Application.Selection, WidthType.Page);
            }
        }

        private void BtnChangeTableSizeToContentArea_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.Selection.Tables.Count == 0)
            {
                if (DedicatedFunctions.ShowMessage("جدولی انتخاب نشده؛ در کل جداول اعمال شود؟", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    changeTableSize(Globals.ThisAddIn.Application.ActiveDocument, WidthType.Margin);
            }
            else
            {
                changeTableSize(Globals.ThisAddIn.Application.ActiveDocument, Globals.ThisAddIn.Application.Selection, WidthType.Margin);
            }
        }

        private void BtnRotateTableToRight_Click(object sender, EventArgs e)
        {
            char[,] a = new char[6, 6]
            {
            {'#', '#', '#', '%', '%', '%'},
            {'#', '#', '#', '%', '%', '%'},
            {'#', '#', '#', '%', '%', '%'},
            {'*', '*', '*', '+', '+', '+'},
            {'*', '*', '*', '+', '+', '+'},
            {'*', '*', '*', '+', '+', '+'},
            };
            Console.WriteLine("Rotate 90");
            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < 6; j++)
                {
                    Console.Write(" " + a[a.GetLength(0) - 1 - j, i]);
                }
                Console.WriteLine();
            }
            //Console.WriteLine();
            //Console.WriteLine("Rotate 180");
            //
            //for (int i = 0; i < 6; i++)
            //{
            //    for (int j = 0; j < 6; j++)
            //    {
            //        Console.Write(" " + a[a.GetLength(0) - 1 - i, a.GetLength(1) - 1 - j]);
            //    }
            //    Console.WriteLine();
            //}
            //Console.WriteLine();
            //Console.WriteLine("Rotate 270");
            //for (int i = 0; i < 6; i++)
            //{
            //    for (int j = 0; j < 6; j++)
            //    {
            //        Console.Write(" " + a[j, a.GetLength(1) - 1 - i]);
            //    }
            //    Console.WriteLine();
            //}
        }

        private void TbtnShowDialogCaptionFigure_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.CaptionSettings_ShowDialogCaptionFigure = (bool)((System.Windows.Controls.Primitives.ToggleButton)sender).IsChecked;
            Properties.Settings.Default.Save();
        }

        #endregion

        #region Functions
        internal enum WidthType
        {
            Page,
            Margin,
        }
        internal void changeShapesSize(Document doc, WidthType widthType, bool lockAspectRatio, bool allowOverlap)
        {
            //first section index in chapter1
            Universities university = DedicatedFunctions.getUniversity(doc);
            DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
            TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

            PageIDs previousPageID = getLastSectionPageID(doc, university, templateType, documentType, PageIDs._page_Chapter1, WhichIndex.Previous);
            int firstSectionIndexPageChapter1 = DedicatedFunctions.getPageIDIndex(doc, previousPageID.ToString());

            //last section index in last available section
            int chaptersCount = DedicatedFunctions.getCurrentChaptersCount(doc);
            int sectionNumberLastPageChapter = DedicatedFunctions.getPageIDIndex(doc, InitialVariables.initialPageChapters.ToString() + chaptersCount);


            for (int i = firstSectionIndexPageChapter1; i <= sectionNumberLastPageChapter; i++)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر اندازه شکل ها");


                foreach (object shapeObject in doc.Sections[i].Range.ShapeRange)
                {
                    Shape shape = shapeObject as Shape;
                    if (lockAspectRatio)
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    //shape.WrapFormat.Type = WdWrapType.wdWrapTopBottom;
                    if (allowOverlap)
                        shape.WrapFormat.AllowOverlap = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
                    else
                        shape.WrapFormat.AllowOverlap = (int)Microsoft.Office.Core.MsoTriState.msoFalse;


                    if (widthType == WidthType.Page)
                        shape.Width = DedicatedFunctions.getPageDocumentSize(doc.Sections[i].Range).Width;
                    else if (widthType == WidthType.Margin)
                    {
                        shape.Width = DedicatedFunctions.getActiveAreaDocumentSize(doc.Sections[i].Range).Width;
                        shape.Left = 0;
                    }
                }

                foreach (InlineShape inlineShape in doc.Sections[i].Range.InlineShapes)
                {
                    Shape shape = inlineShape.ConvertToShape();
                    if (lockAspectRatio)
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                    if (allowOverlap)
                        shape.WrapFormat.AllowOverlap = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
                    else
                        shape.WrapFormat.AllowOverlap = (int)Microsoft.Office.Core.MsoTriState.msoFalse;


                    shape.WrapFormat.Type = WdWrapType.wdWrapTopBottom;
                    //shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    //shape.ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                    //shape.RelativeHorizontalSize = WdRelativeHorizontalSize.wdRelativeHorizontalSizePage;

                    if (widthType == WidthType.Page)
                        shape.Width = DedicatedFunctions.getPageDocumentSize(inlineShape.Range).Width;
                    else if (widthType == WidthType.Margin)
                    {
                        shape.Width = DedicatedFunctions.getActiveAreaDocumentSize(inlineShape.Range).Width;
                        shape.Left = 0;
                    }

                }

                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        internal void changeShapesSize(Document doc, Selection selection, WidthType widthType, bool lockAspectRatio, bool allowOverlap)
        {
            //first section index in chapter1
            Universities university = DedicatedFunctions.getUniversity(doc);
            DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
            TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

            PageIDs previousPageID = getLastSectionPageID(doc, university, templateType, documentType, PageIDs._page_Chapter1, WhichIndex.Previous);
            int firstSectionIndexPageChapter1 = DedicatedFunctions.getPageIDIndex(doc, previousPageID.ToString());

            //last section index in last available section
            int chaptersCount = DedicatedFunctions.getCurrentChaptersCount(doc);
            int sectionNumberLastPageChapter = DedicatedFunctions.getPageIDIndex(doc, InitialVariables.initialPageChapters.ToString() + chaptersCount);

            for (int i = firstSectionIndexPageChapter1; i <= sectionNumberLastPageChapter; i++)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر اندازه شکل ها");


                foreach (object shapeObject in selection.Range.ShapeRange)
                {
                    Shape shape = shapeObject as Shape;
                    if (lockAspectRatio)
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    //shape.WrapFormat.Type = WdWrapType.wdWrapTopBottom;

                    if (allowOverlap)
                        shape.WrapFormat.AllowOverlap = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
                    else
                        shape.WrapFormat.AllowOverlap = (int)Microsoft.Office.Core.MsoTriState.msoFalse;

                    if (widthType == WidthType.Page)
                        shape.Width = DedicatedFunctions.getPageDocumentSize(selection.Range).Width;
                    else if (widthType == WidthType.Margin)
                    {
                        shape.Width = DedicatedFunctions.getActiveAreaDocumentSize(selection.Range).Width;
                        shape.Left = 0;
                    }

                }

                foreach (InlineShape inlineShape in selection.Range.InlineShapes)
                {
                    Shape shape = inlineShape.ConvertToShape();
                    if (lockAspectRatio)
                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    shape.WrapFormat.Type = WdWrapType.wdWrapTopBottom;
                    //shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    //shape.ScaleWidth(1, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromMiddle);
                    //shape.RelativeHorizontalSize = WdRelativeHorizontalSize.wdRelativeHorizontalSizePage;

                    if (allowOverlap)
                        shape.WrapFormat.AllowOverlap = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
                    else
                        shape.WrapFormat.AllowOverlap = (int)Microsoft.Office.Core.MsoTriState.msoFalse;

                    if (widthType == WidthType.Page)
                        shape.Width = DedicatedFunctions.getPageDocumentSize(inlineShape.Range).Width;
                    else if (widthType == WidthType.Margin)
                    {
                        shape.Width = DedicatedFunctions.getActiveAreaDocumentSize(inlineShape.Range).Width;
                        shape.Left = 0;
                    }
                }

                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }

        internal void changeTableSize(Document doc, WidthType widthType)
        {
            //first section index in chapter1
            Universities university = DedicatedFunctions.getUniversity(doc);
            DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
            TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

            PageIDs previousPageID = getLastSectionPageID(doc, university, templateType, documentType, PageIDs._page_Chapter1, WhichIndex.Previous);
            int firstSectionIndexPageChapter1 = DedicatedFunctions.getPageIDIndex(doc, previousPageID.ToString());

            //last section index in last available section
            int chaptersCount = DedicatedFunctions.getCurrentChaptersCount(doc);
            int sectionNumberLastPageChapter = DedicatedFunctions.getPageIDIndex(doc, InitialVariables.initialPageChapters.ToString() + chaptersCount);

            for (int i = firstSectionIndexPageChapter1; i <= sectionNumberLastPageChapter; i++)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر اندازه شکل ها");
                foreach (Table table in doc.Sections[i].Range.Tables)
                {
                    //table.AllowAutoFit
                    //table.AllowPageBreaks
                    //table.Range

                    if (widthType == WidthType.Page)
                    {
                        table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    }
                    else if (widthType == WidthType.Margin)
                    {
                        table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        internal void changeTableSize(Document doc, Selection selection, WidthType widthType)
        {
            //first section index in chapter1
            Universities university = DedicatedFunctions.getUniversity(doc);
            DocumentTypes documentType = DedicatedFunctions.getDocumentType(doc);
            TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

            PageIDs previousPageID = getLastSectionPageID(doc, university, templateType, documentType, PageIDs._page_Chapter1, WhichIndex.Previous);
            int firstSectionIndexPageChapter1 = DedicatedFunctions.getPageIDIndex(doc, previousPageID.ToString());

            //last section index in last available section
            int chaptersCount = DedicatedFunctions.getCurrentChaptersCount(doc);
            int sectionNumberLastPageChapter = DedicatedFunctions.getPageIDIndex(doc, InitialVariables.initialPageChapters.ToString() + chaptersCount);

            for (int i = firstSectionIndexPageChapter1; i <= sectionNumberLastPageChapter; i++)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر اندازه شکل ها");
                foreach (Table table in selection.Range.Tables)
                {
                    //table.AllowAutoFit
                    //table.AllowPageBreaks
                    //table.Range

                    if (widthType == WidthType.Page)
                    {
                        table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow);
                    }
                    else if (widthType == WidthType.Margin)
                    {
                        //table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        #endregion

        /// <summary>
        /// Finds a Child of a given item in the visual tree. 
        /// </summary>
        /// <param name="parent">A direct parent of the queried item.</param>
        /// <typeparam name="T">The type of the queried item.</typeparam>
        /// <param name="childName">x:Name or Name of child. </param>
        /// <returns>The first parent item that matches the submitted type parameter. 
        /// If not matching item can be found, 
        /// a null parent is being returned.</returns>
        public static T FindChild<T>(DependencyObject parent, string childName)
           where T : DependencyObject
        {
            // Confirm parent and childName are valid. 
            if (parent == null) return null;

            T foundChild = null;

            int childrenCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childrenCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                // If the child is not of the request child type child
                T childType = child as T;
                if (childType == null)
                {
                    // recursively drill down the tree
                    foundChild = FindChild<T>(child, childName);

                    // If the child is found, break so we do not overwrite the found child. 
                    if (foundChild != null) break;
                }
                else if (!string.IsNullOrEmpty(childName))
                {
                    var frameworkElement = child as FrameworkElement;
                    // If the child's name is set for search
                    if (frameworkElement != null && frameworkElement.Name == childName)
                    {
                        // if the child's name is of the request name
                        foundChild = (T)child;
                        break;
                    }
                }
                else
                {
                    // child element found.
                    foundChild = (T)child;
                    break;
                }
            }

            return foundChild;
        }

        public Action NormalStateFormRequest { get; set; }
    }
}
