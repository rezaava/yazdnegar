using System;
using System.Collections.ObjectModel;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using System.Xml;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;

namespace ShivaNegar.TaskPanes.InsertCitation
{
    /// <summary>
    /// Interaction logic for CrossReferenceControl.xaml
    /// </summary>
    public partial class InsertCitationControl : UserControl
    {
        internal ObservableCollection<ReferenceItem> ReferenceItems { get; set; } = new ObservableCollection<ReferenceItem>();

        Document doc;
        Microsoft.Office.Interop.Word.Selection selection;

        bool setAsChangeDedicate = false;
        public InsertCitationControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();

            btnInsert.Click += BtnInsert_Click;
            btnRefreshList.Click += BtnRefreshList_Click;
            lstCitation.KeyDown += List_KeyDown;
            lstCitation.MouseDoubleClick += list_MouseDoubleClick;
            txtSearchCitation.TextChanged += txtSearchCitation_TextChanged;

            btnImportSources.Click += BtnImportSources_Click;
            btnInsertSources.Click += BtnInsertSources_Click;
        }

        private void BtnInsertSources_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Globals.ThisAddIn.insertSources();
        }

        private void BtnImportSources_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Globals.ThisAddIn.importSourcesDialog();
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
        internal void initializeInsertCitationControl(Document _doc, Microsoft.Office.Interop.Word.Selection _selection)
        {
            doc = _doc;
            selection = _selection;

            ReferenceItems.Clear();
            lstCitation.ItemsSource = ReferenceItems;

            AddCitationsToList(doc);
            ObservableCollection<ReferenceItem> referenceItems = SearchInList(txtSearchCitation.Text);
            messagesManager(doc.Bibliography.Sources.Count, referenceItems.Count);
            txtSearchCitation.Clear();
        }

        private void insert(bool focusDocument)
        {
            doc = Globals.ThisAddIn.Application.ActiveDocument;
            selection = Globals.ThisAddIn.Application.Selection;
            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

            ReferenceItem selectedItem = lstCitation.SelectedItem as ReferenceItem;
            if (selectedItem != null)
            {
                string showTagInUndoList = selectedItem.Tag.Length > 20 ? selectedItem.Tag.Substring(0, 20) + "..." : showTagInUndoList = selectedItem.Tag;

                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("درج منبع " + showTagInUndoList + " به سند");

                string bibliographyStyle;
                if (accessType == DedicatedFunctions.AccessType.AccessGranted)
                {
                    bibliographyStyle = DedicatedFunctions.getStaticVariableValue(doc, VariableOptionIDs._variable_option_BibliographyStyle.ToString());
                    if (bibliographyStyle == SettingValues.NotExist || string.IsNullOrEmpty(bibliographyStyle))
                        bibliographyStyle = "APA";
                }
                else
                {
                    bibliographyStyle = "APA";
                }


                if (focusDocument)
                {
                    IntPtr hWnd = new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.Hwnd);
                    DedicatedFunctions.SetFocus(hWnd);
                }

                Globals.ThisAddIn.Application.Options.BibliographyStyle = bibliographyStyle;
                doc.Bibliography.BibliographyStyle = bibliographyStyle;
                Microsoft.Office.Interop.Word.Field s = selection.Fields.Add(selection.Range, WdFieldType.wdFieldCitation, selectedItem.Tag);

                DedicatedFunctions.updateBibliography(doc, selection, accessType);
                try
                {
                    s.Update();
                }
                catch (Exception)
                {
                    //DedicatedFunctions.ShowErrorMessage("منبع شما نباید درون یک منبع دیگر باشد، لطفا در مکانی دیگر منبع را اضافه نمایید");
                }

                //selection.EndKey();
                selection.Next(WdUnits.wdCharacter, 1).Select();

                //s.Select();
                //if (Globals.ThisAddIn.Application.Selection.ParagraphFormat.ReadingOrder == WdReadingOrder.wdReadingOrderRtl)
                //{
                //    DedicatedFunctions.ShowMessage(Globals.ThisAddIn.Application.Selection.Text);
                //    string pattern = "([ا-ی])";
                //    MatchCollection matches = Regex.Matches(Globals.ThisAddIn.Application.Selection.Text, pattern);
                //    if (matches.Count != 0)
                //    {
                //        Globals.ThisAddIn.Application.Selection.RtlRun();
                //        Globals.ThisAddIn.Application.Selection.EndKey();
                //    }
                //}
                //else
                //{
                //    DedicatedFunctions.ShowMessage(Globals.ThisAddIn.Application.Selection.Text);
                //    string pattern = "([a-zA-Z])";
                //    MatchCollection matches = Regex.Matches(Globals.ThisAddIn.Application.Selection.Text, pattern);
                //    if (matches.Count != 0)
                //    {
                //        Globals.ThisAddIn.Application.Selection.LtrRun();
                //        Globals.ThisAddIn.Application.Selection.EndKey();
                //    }
                //}
                //Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

            }
            else
            {
                DedicatedFunctions.ShowErrorMessage("منبعی انتخاب نشده است، لطفا منبع مورد نظر خود را انتخاب نمایید");
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

        private void txtSearchCitation_TextChanged(object sender, System.EventArgs e)
        {
            ObservableCollection<ReferenceItem> referenceItems = SearchInList(txtSearchCitation.Text);
            messagesManager(doc.Bibliography.Sources.Count, referenceItems.Count);
        }

        private void BtnRefreshList_Click(object sender, System.EventArgs e)
        {
            AddCitationsToList(doc);
            ObservableCollection<ReferenceItem> referenceItems = SearchInList(txtSearchCitation.Text);

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

            DedicatedFunctions.updateBibliography(doc, selection, accessType);
            messagesManager(doc.Bibliography.Sources.Count, referenceItems.Count);
        }

        internal void messagesManager(int bibliographySourcesCount, int seachedItemCount)
        {
            if (bibliographySourcesCount > 0)
            {
                cardNoSourcesFound.Visibility = System.Windows.Visibility.Collapsed;
                if (seachedItemCount > 0)
                {
                    txtSourceNotFound.Visibility = System.Windows.Visibility.Collapsed;
                }
                else
                {
                    txtSourceNotFound.Visibility = System.Windows.Visibility.Visible;
                }
                txtSearchCitation.IsEnabled = true;
            }
            else
            {
                cardNoSourcesFound.Visibility = System.Windows.Visibility.Visible;
                txtSourceNotFound.Visibility = System.Windows.Visibility.Visible;
                txtSearchCitation.IsEnabled = false;
            }

        }

        private void btnImportSources_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.importSourcesDialog();
        }

        private void btnInsertSources_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.insertSources();
        }

        internal ObservableCollection<ReferenceItem> SearchInList(string searchText)
        {
            ObservableCollection<ReferenceItem> searchList = new ObservableCollection<ReferenceItem>();
            if (searchText != "")
            {
                foreach (ReferenceItem item in ReferenceItems)
                {
                    if ((item.NameOfReference != null && item.NameOfReference.Contains(searchText)) || (item.Authors != null && item.Authors.Contains(searchText)))
                    {
                        //bMain.Items.Add(item);
                        searchList.Add(item);
                    }
                }
                lstCitation.ItemsSource = searchList;
                return searchList;
            }
            else
            {
                lstCitation.ItemsSource = ReferenceItems;
                return ReferenceItems;
            }
        }

        internal void AddCitationsToList(Document doc)
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    ReferenceItems.Clear();
                });

                foreach (Source source in doc.Bibliography.Sources)
                {
                    string xmlString = source.XML;

                    int startPointNameSpace = xmlString.IndexOf("\"", xmlString.IndexOf("xmlns:b"));
                    int EndPointNameSpace = xmlString.IndexOf("\"", startPointNameSpace + 1);
                    string nameSpace = xmlString.Substring(startPointNameSpace + 1, EndPointNameSpace - startPointNameSpace - 1);

                    XmlDocument xml = new XmlDocument();
                    xml.LoadXml(xmlString);
                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(xml.NameTable);
                    nsmgr.AddNamespace("b", nameSpace);

                    string authorsName = "";
                    foreach (XmlNode item in xml.SelectNodes("/b:Source/b:Author/b:Author/b:NameList/b:Person", nsmgr))
                    {
                        authorsName += item.InnerText + ";";
                    }
                    Dispatcher.Invoke(() =>
                    {
                        ReferenceItems.Add(new ReferenceItem
                        {
                            Tag = xml.SelectSingleNode("/b:Source/b:Tag", nsmgr)?.InnerText,
                            //NameOfReference = xml.SelectSingleNode("/b:Source/b:Caption" , nsmgr)?.InnerText ,
                            NameOfReference = xml.SelectSingleNode("/b:Source/b:Title", nsmgr)?.InnerText,
                            Authors = authorsName,
                            TypeOfReference = xml.SelectSingleNode("/b:Source/b:SourceType", nsmgr)?.InnerText,
                        });
                    });
                }
                Dispatcher.Invoke(() =>
                {
                    lstCitation.ItemsSource = ReferenceItems;
                });
            });
        }
    }
    public class ReferenceItem
    {
        public string Tag { get; set; }
        public string NameOfReference { get; set; }
        public string Authors { get; set; }
        public string TypeOfReference { get; set; }
    }
}
