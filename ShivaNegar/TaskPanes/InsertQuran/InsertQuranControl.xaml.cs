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

namespace ShivaNegar.TaskPanes.InsertQuran
{
    /// <summary>
    /// Interaction logic for CrossReferenceControl.xaml
    /// </summary>
    public partial class InsertQuranControl : UserControl
    {
        internal ObservableCollection<AyehQuranItem> QuranItems { get; set; } = new ObservableCollection<AyehQuranItem>();

        List<SurahQuranItem> surahQuranItems;

        public InsertQuranControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();

            surahQuranItems = new List<SurahQuranItem>();

            comboSelectSurah.SelectionChanged += ComboSelectSurah_SelectionChanged;
            btnInsertAyeh.Click += BtnInsertAyeh_Click;
            lstQuran.KeyDown += List_KeyDown;
            lstQuran.MouseDoubleClick += list_MouseDoubleClick;
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

        internal void initializeInsertQuranControl()
        {
            comboSelectSurah.SelectedIndex = -1;
            comboSelectSurah.Items.Clear();

            setSurahItems();

            lstQuran.ItemsSource = null;
        }

        private void insert(bool focusDocument)
        {
            if (lstQuran.SelectedIndex != -1)
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
                AyehQuranItem selectedAyehItem = lstQuran.SelectedItem as AyehQuranItem;

                if (focusDocument)
                {
                    IntPtr hWnd = new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.Hwnd);
                    DedicatedFunctions.SetFocus(hWnd);
                }
                doc.ActiveWindow.Selection.RtlPara();

                if (tbtnInsertSurahAyeh.IsChecked == true)
                {
                    string surahName = surahQuranItems[comboSelectSurah.SelectedIndex].SurahName;
                    string ayehNumber = ((AyehQuranItem)lstQuran.SelectedItem).AyehNumber;
                    DedicatedFunctions.insertText(doc, doc.ActiveWindow.Selection, "سوره " + surahName + " آیه " + ayehNumber + ": ");
                }
                DedicatedFunctions.insertText(doc, doc.ActiveWindow.Selection, selectedAyehItem.AyehText + " ");
                DedicatedFunctions.insertParagraph(doc.ActiveWindow.Selection);
            }
        }
        private void setSurahItems()
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                using (var reader = new StreamReader(DedicatedFunctions.getStream(EmbeddedResourceNames.Surah), Encoding.UTF8))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var surahItems = csv.GetRecords<SurahQuranItem>();
                    foreach (var surahItem in surahItems)
                    {
                        surahQuranItems.Add(surahItem);
                    }
                }
                Dispatcher.Invoke(() =>
                {
                    foreach (var surahItem in surahQuranItems)
                    {
                        string text = surahItem.SurahNumber + "- سوره " + surahItem.SurahName + " (" + surahItem.AyehCount + " آیه)";
                        comboSelectSurah.Items.Add(text);
                    }
                });
            });
        }

        private void setAyehItems(int surahNumber)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                List<AyehQuranItem> ayehQuranItems = new List<AyehQuranItem>();
                using (var reader = new StreamReader(DedicatedFunctions.getStream(EmbeddedResourceNames.Ayeh), Encoding.UTF8, true))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    var ayehItems = csv.GetRecords<AyehQuranItem>();
                    foreach (var ayehItem in ayehItems)
                    {
                        if (int.Parse(ayehItem.SurahNumber) == surahNumber)
                            ayehQuranItems.Add(ayehItem);
                    }
                }

                Dispatcher.Invoke(() =>
                {
                    if (comboSelectSurah.SelectedIndex != -1)
                        lstQuran.ItemsSource = ayehQuranItems;
                });
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


        private void ComboSelectSurah_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            setAyehItems(comboSelectSurah.SelectedIndex + 1);
        }
        #endregion
    }
    internal class SurahQuranItem
    {
        //public SurahQuranItem(string surahName , string ayehNumber , string ayehCount)
        //{
        //	this.SurahName = surahName;
        //	this.SurahNumber = ayehNumber;
        //	this.AyehCount = ayehCount;
        //}

        public string SurahName { get; set; }
        public string SurahNumber { get; set; }
        public string AyehCount { get; set; }
    }
    internal class AyehQuranItem
    {
        //public AyehQuranItem(string surahNumber , string ayehNumber , string ayehText , string hezb , string joze , string isSajdeh , string surahName)
        //{
        //	SurahNumber = surahNumber;
        //	AyehNumber = ayehNumber;
        //	AyehText = ayehText;
        //	Hezb = hezb;
        //	Joze = joze;
        //	IsSajdeh = isSajdeh;
        //	SurahName = surahName;
        //}

        public string SurahNumber { get; set; }
        public string AyehNumber { get; set; }
        public string AyehText { get; set; }
        public string Hezb { get; set; }
        public string Joze { get; set; }
        public string IsSajdeh { get; set; }
        public string SurahName { get; set; }
    }
}
