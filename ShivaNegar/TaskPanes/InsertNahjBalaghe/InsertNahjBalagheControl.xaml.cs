using System;
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

namespace ShivaNegar.TaskPanes.InsertNahjBalaghe
{
    /// <summary>
    /// Interaction logic for CrossReferenceControl.xaml
    /// </summary>
    public partial class InsertNahjBalagheControl : UserControl
    {
        internal ObservableCollection<SubNahjBalagheItem> SubNahjBalagheItems { get; set; } = new ObservableCollection<SubNahjBalagheItem>();
        public InsertNahjBalagheControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();

            comboSelectNahjBalaghe.SelectionChanged += ComboSelectNahjBalaghe_SelectionChanged;
            btnInsert.Click += BtnInsert_Click;
            lstNahjBalaghe.KeyDown += List_KeyDown;
            lstNahjBalaghe.MouseDoubleClick += list_MouseDoubleClick;
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

        internal void initializeInsertNahjBalagheControl()
        {
            comboSelectNahjBalaghe.SelectedIndex = -1;
            comboSelectNahjBalaghe.Items.Clear();

            string text = "خطبه ها (241 خطبه)";
            comboSelectNahjBalaghe.Items.Add(text);
            text = "حکمت ها (360 حکمت)";
            comboSelectNahjBalaghe.Items.Add(text);
            text = "نامه ها (79 نامه)";
            comboSelectNahjBalaghe.Items.Add(text);

            lstNahjBalaghe.ItemsSource = null;
        }

        private void setSubItems(int comboBoxIndex)
        {
            if (comboBoxIndex != -1)
            {
                System.Threading.Tasks.Task.Run(() =>
                {
                    StreamReader reader;

                    if (comboBoxIndex == 0)
                        reader = new StreamReader(DedicatedFunctions.getStream(EmbeddedResourceNames.Khotbeh), Encoding.UTF8, true);
                    else if (comboBoxIndex == 1)
                        reader = new StreamReader(DedicatedFunctions.getStream(EmbeddedResourceNames.Hekmat), Encoding.UTF8, true);
                    else if (comboBoxIndex == 2)
                        reader = new StreamReader(DedicatedFunctions.getStream(EmbeddedResourceNames.Nameh), Encoding.UTF8, true);
                    else
                        return;

                    using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                    {
                        var Items = csv.GetRecords<SubNahjBalagheItem>();

                        Dispatcher.Invoke(() =>
                        {
                            if (comboSelectNahjBalaghe.SelectedIndex != -1)
                                lstNahjBalaghe.ItemsSource = Items;
                        });
                    }
                });
            }
        }
        private void insert(bool focusDocument)
        {
            if (lstNahjBalaghe.SelectedIndex != -1)
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
                doc.ActiveWindow.Selection.RtlPara();

                SubNahjBalagheItem selectedSubItem = lstNahjBalaghe.SelectedItem as SubNahjBalagheItem;
                DedicatedFunctions.insertText(doc, doc.ActiveWindow.Selection, selectedSubItem.Text);

                DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection);
            }
        }

        #region Events
        private void ComboSelectNahjBalaghe_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            setSubItems(comboSelectNahjBalaghe.SelectedIndex);

            if (comboSelectNahjBalaghe.SelectedIndex == 0)
            {
                btnInsert.Content = "درج خطبه";
            }
            else if (comboSelectNahjBalaghe.SelectedIndex == 1)
            {
                btnInsert.Content = "درج حکمت";
            }
            else if (comboSelectNahjBalaghe.SelectedIndex == 2)
            {
                btnInsert.Content = "درج نامه";
            }
        }
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
        private void BtnInsert_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            insert(true);
        }
        #endregion
    }
    internal class SubNahjBalagheItem
    {
        public string Number { get; set; }
        public string Text { get; set; }
        public string Caption { get; set; }
    }
}
