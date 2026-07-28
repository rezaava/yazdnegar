using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Templates;

namespace ShivaNegar.TaskPanes.DefenseAnnouncements
{
    /// <summary>
    /// Interaction logic for CrossReferenceControl.xaml
    /// </summary>
    public partial class DefenseAnnouncementsControl : UserControl
    {
        internal ObservableCollection<DefenseAnnouncementItem> DefenseAnnouncementItems { get; set; } = new ObservableCollection<DefenseAnnouncementItem>();


        bool setAsChangeDedicate = false;
        public DefenseAnnouncementsControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();
            btnCreate.Click += BtnInsertAyeh_Click;
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

        internal void initializeDefenseAnnouncementControl()
        {
            DefenseAnnouncementItems.Clear();
            lstBox.ItemsSource = DefenseAnnouncementItems;
            setDedicateTypes();
        }

        private void create()
        {
            if (lstBox.SelectedIndex != -1)
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

                Universities university = DedicatedFunctions.getUniversity(doc);
                TemplateTypes templateType = DedicatedFunctions.getTemplateType(doc);

                AcademicDegrees academicDegree = DedicatedFunctions.getAcademicDegreeID(doc);
                bool found = false;
                System.Windows.Forms.SaveFileDialog fileDialog = new System.Windows.Forms.SaveFileDialog();
                fileDialog.Title = "مکان و نام خروجی فایل خود را مشخص کنید";
                fileDialog.AddExtension = false;
                fileDialog.Filter = "Word Document (*.docx)|*.docx";
                if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    if (fileDialog.FileName != doc.FullName)
                    {
                        DefenseAnnouncementItem daModel = DefenseAnnouncementItems[lstBox.SelectedIndex];
                        Stream stream = DedicatedFunctions.getStream(daModel.ResourcePath);
                        string templatePath = DedicatedFunctions.copyFileToFolder(stream, fileDialog.FileName);

                        Globals.ThisAddIn.DisableEvents = true;
                        Document specifiedDocument = Globals.ThisAddIn.Application.Documents.Open(FileName: templatePath, Visible: false, ReadOnly: false, OpenAndRepair: false, NoEncodingDialog: true);
                        Globals.ThisAddIn.DisableEvents = false;


                        for (int i = specifiedDocument.ContentControls.Count; i > 0; i--)
                        {
                            //آرایه از این نوع از اندیس 1 شروع میشود
                            Microsoft.Office.Interop.Word.ContentControl ccSpecified = specifiedDocument.ContentControls[i];

                            if (ccSpecified.Tag == ContentControlNames._field_Examiner_Fa.ToString())
                            {
                                ccSpecified.LockContents = false;
                                ccSpecified.LockContentControl = false;
                                ccSpecified.Range.Text = txtExaminer.Text;
                                ccSpecified.Delete(false);
                            }
                            else if (ccSpecified.Tag == ContentControlNames._field_DefenseLocation_Fa.ToString())
                            {
                                ccSpecified.LockContents = false;
                                ccSpecified.LockContentControl = false;
                                ccSpecified.Range.Text = txtLocation.Text;
                                ccSpecified.Delete(false);
                            }
                            else if (ccSpecified.Tag == ContentControlNames._field_DefenseDate_Fa.ToString())
                            {
                                ccSpecified.LockContents = false;
                                ccSpecified.LockContentControl = false;
                                ccSpecified.Range.Text = txtDate.Text;
                                ccSpecified.Delete(false);
                            }
                            else if (ccSpecified.Tag == ContentControlNames._field_Icon_University.ToString())
                            {
                                if (ccSpecified.Type == WdContentControlType.wdContentControlPicture)
                                {
                                    ccSpecified.LockContents = false;
                                    ccSpecified.LockContentControl = false;
                                    if (ccSpecified.Range.InlineShapes.Count > 0)
                                    {
                                        if (university != Universities.Nothing && templateType != TemplateTypes.Nothing)
                                        {
                                            Bitmap universityIcon = TemplateAccess.getUniversityIcon(university, templateType);
                                            string tempImageUniversityPath = DedicatedFunctions.copyImageToTempFolder(universityIcon, nameof(universityIcon));

                                            InlineShape prevInlineShape = ccSpecified.Range.InlineShapes[1];
                                            float height = prevInlineShape.Height;
                                            InlineShape shape = ccSpecified.Range.InlineShapes.AddPicture(tempImageUniversityPath, Range: prevInlineShape.Range);
                                            shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                                            shape.Height = height;
                                            shape.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

                                            DedicatedFunctions.removeFileFromSystem(tempImageUniversityPath);
                                        }
                                    }
                                    ccSpecified.Delete(false);
                                }
                            }
                            else
                            {
                                if (ccSpecified.Tag != null)
                                {
                                    //آرایه از این نوع از اندیس 1 شروع میشود
                                    Microsoft.Office.Interop.Word.ContentControls ccs = doc.SelectContentControlsByTag(ccSpecified.Tag);
                                    if (ccs.Count >= 1)
                                    {
                                        ccSpecified.LockContents = false;
                                        ccSpecified.Range.Text = ccs[1].Range.Text;
                                    }
                                    ccSpecified.LockContentControl = false;
                                    ccSpecified.Delete(false);
                                }
                            }
                        }
                        specifiedDocument.ActiveWindow.Visible = true;
                        specifiedDocument.Save();
                    }
                    else
                    {
                        DedicatedFunctions.ShowErrorMessage("سند خروجی نمیتواند با سند اختصاصی شما جایگزین شود، لطفا با نامی دیگر سند را ذخیره کنید");
                    }
                }
            }
        }

        private void setDedicateTypes()
        {
            System.Threading.Tasks.Task.Run(() =>
            {
                Dispatcher.Invoke(() =>
                {
                    DefenseAnnouncementItems.Clear();
                });

                List<DefenseAnnouncementModel> lls = DefenseAnnouncementsAccess.getDefenseAnnouncementsModels();
                foreach (var item in lls)
                {
                    Bitmap bitmap = new Bitmap(DefenseAnnouncementsAccess.getCoverImageStream(item));
                    //bitmapImage = BitmapToBitmapImage(bitmap);

                    string resourceFilePath = item.Path + item.FileName;

                    DefenseAnnouncementItem daItem = new DefenseAnnouncementItem()
                    {
                        ID = item.Id,
                        Name = item.Name,
                        SourceFrom = item.SourceFrom,
                        CoverImage = bitmap,
                        ResourcePath = resourceFilePath,
                    };
                    Dispatcher.Invoke(() =>
                    {
                        DefenseAnnouncementItems.Add(daItem);
                    });
                }
            });

        }

        private void BtnInsertAyeh_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            create();
        }
        public BitmapImage BitmapToBitmapImage(Bitmap bitmap)
        {
            using (var memory = new MemoryStream())
            {
                bitmap.Save(memory, System.Drawing.Imaging.ImageFormat.Png);
                memory.Position = 0;

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }
    }
    internal class DefenseAnnouncementItem
    {
        public Bitmap CoverImage { get; set; }
        public string ID { get; set; }
        public string ResourcePath { get; set; }
        public string Name { get; set; }
        public string SourceFrom { get; set; }
    }
}
