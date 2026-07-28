using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Controls;
using System.Windows.Media;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using static ShivaNegar.DedicatedFunctions;

namespace ShivaNegar.TaskPanes.CrossReference
{
    /// <summary>
    /// Interaction logic for CrossReferenceControl.xaml
    /// </summary>
    public partial class CrossReferenceControl : UserControl
    {
        static string allItems = "همه موارد";

        bool? previousToggleButtonState;


        List<CaptionObject> crossReferenceList;
        List<string> captionLabels = new List<string>();

        string captionType = "";
        string captionText = "";
        string previewCrossReference = "";

        string selectedChapterStyleLevel;
        string selectedChapterSeq;

        //string previousTxtBoxBeforeText = "";
        //string previousTxtBoxAfterText = "";


        Document doc;


        public CrossReferenceControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();

            crossReferenceList = new List<CaptionObject> { };

            comboCrossReferenceTo.SelectionChanged += ComboCrossReferenceTo_SelectionChanged;
            btnRefer.Click += BtnRefer_Click;
            lstCaptionList.KeyDown += List_KeyDown;
            lstCaptionList.MouseDoubleClick += List_MouseDoubleClick;

            btnRefreshList.Click += BtnRefreshList_Click;
            txtSearchCaptions.TextChanged += TxtSearchCaptions_TextChanged;
            lstCaptionList.SelectionChanged += LstCaptionList_SelectionChanged;
            tbtnReferInsideParentheses.Click += TbtnReferInsideParentheses_Click;
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


        #region Events
        private void TbtnReferInsideParentheses_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            if (lstCaptionList.SelectedIndex != -1 && previewCrossReference != "")
            {
                if (previewCrossReference.Contains("(" + Constants.CaptionLabels.captionFormula))
                {
                    previousToggleButtonState = tbtnReferInsideParentheses.IsChecked;
                    tbtnReferInsideParentheses.IsChecked = true;
                    tbtnReferInsideParentheses.IsEnabled = false;

                    txtPreview.Text = previewCrossReference + ")";
                }
                else
                {
                    if (previousToggleButtonState != null)
                    {
                        tbtnReferInsideParentheses.IsChecked = previousToggleButtonState;
                        previousToggleButtonState = null;
                        tbtnReferInsideParentheses.IsEnabled = true;
                    }

                    if (tbtnReferInsideParentheses.IsChecked == true)
                        txtPreview.Text = "(" + previewCrossReference + ")";
                    else
                        txtPreview.Text = previewCrossReference;
                }
            }
        }

        private void LstCaptionList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (lstCaptionList.SelectedIndex != -1)
            {
                btnRefer.IsEnabled = true;

                captionText = lstCaptionList.SelectedItem.ToString();
                foreach (CaptionObject item in crossReferenceList)
                {
                    if (item.CaptionText == lstCaptionList.SelectedItem.ToString())
                    {
                        captionType = item.CaptionType;
                    }
                }

                bool includeChapterNumber = Globals.ThisAddIn.Application.CaptionLabels[captionType].IncludeChapterNumber;

                try
                {
                    MatchCollection matches = Regex.Matches(captionText, @"\d+");

                    if (includeChapterNumber)
                    {
                        string[] fetchSeqAndStyleNumber = matches.Cast<Match>()
                                                                 .Take(2)
                                                                 .Select(match => match.Value)
                                                                 .ToArray();

                        selectedChapterStyleLevel = fetchSeqAndStyleNumber[0];
                        selectedChapterSeq = fetchSeqAndStyleNumber[1];

                        //get preview
                        previewCrossReference = captionText.Substring(0, captionText.IndexOf(selectedChapterSeq, captionText.IndexOf(selectedChapterStyleLevel) + selectedChapterStyleLevel.Length) + selectedChapterSeq.Length);
                    }
                    else
                    {
                        string[] fetchSeqAndStyleNumber = matches.Cast<Match>()
                             .Take(1)
                             .Select(match => match.Value)
                             .ToArray();
                        selectedChapterStyleLevel = null;
                        selectedChapterSeq = fetchSeqAndStyleNumber[0];

                        //get preview
                        previewCrossReference = captionText.Substring(0, captionText.IndexOf(selectedChapterSeq) + selectedChapterSeq.Length);
                    }
                }
                catch (Exception)
                {
                    previewCrossReference = "???";
                    btnRefer.IsEnabled = false;
                }

                // set preview
                if (previewCrossReference != "")
                {
                    if (previewCrossReference.Contains("(" + Constants.CaptionLabels.captionFormula))
                    {
                        previousToggleButtonState = tbtnReferInsideParentheses.IsChecked;

                        tbtnReferInsideParentheses.IsChecked = true;
                        tbtnReferInsideParentheses.IsEnabled = false;

                        txtPreview.Text = previewCrossReference + ")";
                    }
                    else
                    {
                        if (previousToggleButtonState != null)
                        {
                            tbtnReferInsideParentheses.IsChecked = previousToggleButtonState;
                            previousToggleButtonState = null;
                            tbtnReferInsideParentheses.IsEnabled = true;
                        }

                        if (tbtnReferInsideParentheses.IsChecked == true)
                            txtPreview.Text = "(" + previewCrossReference + ")";
                        else
                            txtPreview.Text = previewCrossReference;
                    }

                }
                else
                    txtPreview.Text = "";
            }
            else
            {
                txtPreview.Text = "";
                previewCrossReference = "";
                captionType = "";
                captionText = "";
                btnRefer.IsEnabled = false;
            }
        }

        private void TxtSearchCaptions_TextChanged(object sender, TextChangedEventArgs e)
        {
            SearchInList(txtSearchCaptions.Text);
        }

        private void BtnRefreshList_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            RefreshList();
        }


        private void ComboCrossReferenceTo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboCrossReferenceTo.SelectedIndex != -1)
                SearchInList(txtSearchCaptions.Text);
        }
        private void insert(bool focusDocument)
        {
            if (!btnRefer.IsEnabled)
                return;

            try
            {
                if (lstCaptionList.SelectedIndex != -1)
                {
                    //get number of caption
                    bool includeChapterNumber = Globals.ThisAddIn.Application.CaptionLabels[captionType].IncludeChapterNumber;
                    int chapterStyleLevel = Globals.ThisAddIn.Application.CaptionLabels[captionType].ChapterStyleLevel;

                    int seqNumber = 0;
                    //Globals.ThisAddIn.Application.CaptionLabels[captionType].NumberStyle =;
                    if (includeChapterNumber)
                    {

                        int seqCounter = 0;
                        foreach (Field field in doc.Fields)
                        {
                            if (field.Type == WdFieldType.wdFieldStyleRef)
                            {
                                if (field.Code.Text.ToLower()
                                    .Replace(" ", "")
                                    .Contains("\\s" + chapterStyleLevel.ToString()) || field.Code.Text.ToLower().Replace(" ", "").Contains(chapterStyleLevel.ToString() + "\\s"))
                                {
                                    if (field.Next.Type == WdFieldType.wdFieldSequence)
                                    {
                                        if (field.Next.Code.Text.ToLower()
                                            .Contains(captionType.ToLower()) && field.Next.Code.Text.ToLower().Replace(" ", "").Contains("\\s" + chapterStyleLevel.ToString()))
                                        {
                                            seqCounter++;

                                            if (field.Result.Text.Trim().Replace("‏", "").Replace("‎", "") == selectedChapterStyleLevel.Trim() && field.Next.Result.Text.Trim() == selectedChapterSeq.Trim())
                                            {
                                                seqNumber = seqCounter;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        if (seqNumber == 0)
                            throw new Exception();
                    }
                    else
                    {
                        seqNumber = int.Parse(selectedChapterSeq);
                        if (seqNumber == 0)
                            throw new Exception();
                    }

                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("ارجاع دادن به " + captionType);

                    if (focusDocument)
                    {
                        IntPtr hWnd = new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.Hwnd);
                        DedicatedFunctions.SetFocus(hWnd);
                    }

                    Selection selection = Globals.ThisAddIn.Application.Selection;

                    selection.Collapse(WdCollapseDirection.wdCollapseStart);

                    if (captionType == Constants.CaptionLabels.captionFormula)
                    {
                        selection.InsertBefore(" ");
                    }
                    else if (tbtnReferInsideParentheses.IsChecked == true)
                        selection.InsertBefore("(");
                    //selection.InsertBefore(")");

                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                    selection.InsertCrossReference(captionType, WdReferenceKind.wdOnlyLabelAndNumber, seqNumber, true);

                    if (tbtnReferInsideParentheses.IsChecked == true)
                        selection.InsertAfter(")");
                    //selection.InsertAfter("(");

                    selection.Collapse(WdCollapseDirection.wdCollapseEnd);

                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();

                }
                else
                {
                    throw new Exception();
                }
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("عنوان " + captionType + " انتخاب شده قابل ارجاع دادن نیست، لطفا لیست عناوین را بروزرسانی کنید\nیا عنوان انتخابی خود را پاک کرده و مجدد اقدام به ساخت عنوان " + captionType + " خود نمایید");
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        private void List_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter)
            {
                insert(false);
            }
        }
        private void List_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            insert(false);
        }
        private void BtnRefer_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            insert(true);
        }
        #endregion

        #region Functions
        internal void useCrossReference(Document doc)
        {
            this.doc = doc;

            AccessType accessType = DedicatedFunctions.hasAccess(doc);
            bool accessGranted = accessType == AccessType.AccessGranted;
            captionLabels = DedicatedFunctions.getCaptionLabels(doc, accessGranted);

            setComboBoxData(captionLabels);

            if (comboCrossReferenceTo.Items.Count != 0)
                comboCrossReferenceTo.SelectedIndex = 0;
            //setListBoxData();
            if (lstCaptionList.Items.Count != 0)
                lstCaptionList.SelectedIndex = 0;
            else
                btnRefer.IsEnabled = false;
            txtSearchCaptions.Focus();
        }


        internal void setComboBoxData(List<string> captionLabels)
        {
            comboCrossReferenceTo.Items.Clear();
            comboCrossReferenceTo.Items.Add(allItems);

            foreach (var item in captionLabels)
            {
                comboCrossReferenceTo.Items.Add(item);
            }
            //comboCrossReferenceTo.Items.AddRange(builtinCaption.ToArray());
        }
        internal void setListBoxData()
        {
            crossReferenceList.Clear();
            lstCaptionList.Items.Clear();
            if (comboCrossReferenceTo.SelectedItem.ToString() == allItems)
            {
                List<string> captions = new List<string>();
                captions.AddRange(comboCrossReferenceTo.Items.Cast<string>().ToList());
                captions.Remove(allItems);

                foreach (string caption in captions)
                {
                    object crossReferenceItems = doc.GetCrossReferenceItems(caption);
                    string[] arr = ((IEnumerable)crossReferenceItems).Cast<object>()
                                .Select(x => x.ToString())
                                .ToArray();
                    foreach (string item in arr)
                    {
                        crossReferenceList.Add(new CaptionObject(caption, item));
                    }
                    foreach (var item in arr)
                    {
                        lstCaptionList.Items.Add(item);
                    }

                }
            }
            else
            {
                object crossReferenceItems = doc.GetCrossReferenceItems(comboCrossReferenceTo.SelectedItem.ToString());
                string[] arr = ((IEnumerable)crossReferenceItems).Cast<object>()
                            .Select(x => x.ToString())
                            .ToArray();
                foreach (string item in arr)
                {
                    crossReferenceList.Add(new CaptionObject(comboCrossReferenceTo.SelectedItem.ToString(), item));
                }
                foreach (var item in arr)
                {
                    lstCaptionList.Items.Add(item);
                }

            }
        }

        private void setListVisibility(int count)
        {
            if (count == 0)
            {
                lstCaptionList.Visibility = System.Windows.Visibility.Collapsed;
                btnRefer.IsEnabled = false;
            }
            else
            {
                lstCaptionList.Visibility = System.Windows.Visibility.Visible;
                btnRefer.IsEnabled = true;
            }
        }

        internal void RefreshList()
        {
            //int previousIndex = comboCrossReferenceTo.SelectedIndex;
            //if(captionLabels.Count != 0)
            //{
            //	setComboBoxData(captionLabels);
            //	if(comboCrossReferenceTo.Items.Count >= previousIndex - 1)
            //		comboCrossReferenceTo.SelectedIndex = previousIndex;
            //}
            SearchInList(txtSearchCaptions.Text);

            //setListBoxData();
            //SearchInList(txtSearchCaptions.Text);

        }

        internal void SearchInList(string searchText)
        {
            btnRefer.IsEnabled = false;
            List<String> searchList = new List<String>();
            if (searchText != "")
            {
                foreach (CaptionObject item in crossReferenceList)
                {
                    if (item.CaptionText.Contains(txtSearchCaptions.Text) || item.CaptionText.Contains(txtSearchCaptions.Text))
                    {
                        searchList.Add(item.CaptionText);
                    }
                }
                lstCaptionList.Items.Clear();
                foreach (var item in searchList)
                {
                    lstCaptionList.Items.Add(item);
                }
                setListVisibility(lstCaptionList.Items.Count);
            }
            else
            {
                setListBoxData();
                setListVisibility(crossReferenceList.Count);
            }

            if (lstCaptionList.Items.Count != 0)
            {
                btnRefer.IsEnabled = true;
                lstCaptionList.SelectedIndex = 0;
            }
            else
            {
                btnRefer.IsEnabled = false;
            }
        }


        #endregion
    }
}
