using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using CsvHelper;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;

namespace ShivaNegar.TaskPanes.InsertDedicate
{
    /// <summary>
    /// Interaction logic for CrossReferenceControl.xaml
    /// </summary>
    public partial class InsertDedicateControl : UserControl
    {
        internal ObservableCollection<DedicateItem> DedicateItems { get; set; } = new ObservableCollection<DedicateItem>();

        List<string> comboBoxItems;

        bool setAsChangeDedicate = false;
        public InsertDedicateControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();

            comboBoxItems = new List<string>();

            comboSelectDedicateType.SelectionChanged += ComboSelectType_SelectionChanged;
            btnInsertAyeh.Click += BtnInsertAyeh_Click;
            lstDedicate.KeyDown += List_KeyDown;
            lstDedicate.MouseDoubleClick += list_MouseDoubleClick;
        }

        private void InitializeMaterialDesign()
        {
            // Create dummy objects to force the MaterialDesign assemblies to be loaded
            // from this assembly, which causes the MaterialDesign assemblies to be searched
            // relative to this assembly's path. Otherwise, the MaterialDesign assemblies
            // are searched relative to Eclipse's path, so they're not found.
            var card = new Card();
            var hue = new Hue("Dummy", Colors.Black, Colors.White);
        }

        internal void initializeInsertDedicateControl()
        {
            comboSelectDedicateType.SelectedIndex = -1;
            comboSelectDedicateType.Items.Clear();

            setDedicateTypes();

            Document doc = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Interop.Word.ContentControl[] ccs = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Dedication_Fa.ToString());

            if (ccs != null && ccs.Length > 0)
            {
                btnInsertAyeh.Content = "تغییر متن تقدیم";
                ccs[0].Range.Select();
                setAsChangeDedicate = true;
            }
            else
            {
                btnInsertAyeh.Content = "درج متن تقدیم";
                setAsChangeDedicate = false;
            }

            DedicateItems.Clear();
            lstDedicate.ItemsSource = DedicateItems;
        }

        private void insert(bool focusDocument)
        {
            if (lstDedicate.SelectedIndex != -1)
            {
                Document doc;
                try
                {
                    doc = Globals.ThisAddIn.Application.ActiveDocument;
                }
                catch (Exception)
                {
                    return;
                }

                if (focusDocument)
                {
                    IntPtr hWnd = new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.Hwnd);
                    DedicatedFunctions.SetFocus(hWnd);
                }

                string dedicateText = DedicateItems[lstDedicate.SelectedIndex].Dedicate;
                if (setAsChangeDedicate)
                {
                    Microsoft.Office.Interop.Word.ContentControl[] ccs = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Dedication_Fa.ToString());

                    if (ccs != null && ccs.Length > 0)
                    {
                        ccs[0].Range.Select();
                        DedicatedFunctions.changeContentControlContents(doc, ContentControlNames._field_Dedication_Fa.ToString(), dedicateText);
                    }
                    else
                    {
                        doc.ActiveWindow.Selection.RtlPara();

                        DedicatedFunctions.insertText(doc, doc.ActiveWindow.Selection, dedicateText);
                        DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection);
                    }
                }
                else
                {
                    doc.ActiveWindow.Selection.RtlPara();

                    DedicatedFunctions.insertText(doc, doc.ActiveWindow.Selection, dedicateText);
                    DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection);
                }
            }
        }

        private void setDedicateTypes()
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                using (var reader = new StreamReader(DedicatedFunctions.getStream(EmbeddedResourceNames.Dedicate), Encoding.UTF8))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var items = csv.GetRecords<DedicateItem>();
                    foreach (var item in items)
                    {
                        if (!comboBoxItems.Contains(item.Category.Trim()))
                            comboBoxItems.Add(item.Category.Trim());
                    }
                }
                Dispatcher.Invoke(() =>
                {
                    foreach (var item in comboBoxItems)
                    {
                        comboSelectDedicateType.Items.Add(item);
                    }
                });
            });
        }

        private void setDedicateItems(string dedicateTypeText)
        {
            DedicateItems.Clear();
            System.Threading.Tasks.Task.Run(() =>
            {
                using (var reader = new StreamReader(DedicatedFunctions.getStream(EmbeddedResourceNames.Dedicate), Encoding.UTF8, true))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var items = csv.GetRecords<DedicateItem>();
                    foreach (var item in items)
                    {
                        item.Category = item.Category.Trim();
                        if (item.Category == dedicateTypeText)
                        {
                            Dispatcher.Invoke(() =>
                            {
                                DedicateItems.Add(item);
                            });
                        }
                    }
                }
            });
        }

        #region Events
        private void List_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                insert(false);
            }
        }
        private void list_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            insert(false);
        }
        private void BtnInsertAyeh_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            insert(true);
        }


        private void ComboSelectType_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            setDedicateItems(comboBoxItems[comboSelectDedicateType.SelectedIndex]);
        }
        #endregion
    }
    internal class DedicateItem
    {
        public string Number { get; set; }
        public string Category { get; set; }
        public string Dedicate { get; set; }
        public string Writer { get; set; }
        public string BookName { get; set; }
    }
}
