using System;
using System.Drawing;
using System.IO;
using System.Windows.Controls;
using System.Windows.Media;
using MaterialDesignColors;
using MaterialDesignThemes.Wpf;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Templates;

namespace ShivaNegar.TaskPanes.DesigningCD
{
    /// <summary>
    /// Interaction logic for DesigningCDControl.xaml
    /// </summary>
    public partial class DesigningCDControl : UserControl, ITaskPaneRequests
    {
        public Action CloseTaskPaneRequest { get; set; }

        Microsoft.Office.Interop.Word.Document doc;
        Microsoft.Office.Interop.Word.Selection selection;

        Bitmap universityIcon = null;
        public DesigningCDControl()
        {
            InitializeMaterialDesign();
            InitializeComponent();

            btnInsert.Click += BtnInsert_Click;
            btnRefreshList.Click += BtnRefreshList_Click;
            btnClose.Click += BtnClose_Click;

            btnChangeBackgroundColor.Click += BtnChangeBackgroundColor_Click;
            btnChangeTextColor.Click += BtnChangeTextColor_Click;
            btnChangeIconColor.Click += BtnChangeIconColor_Click;
            btnChangeBorderColor.Click += BtnChangeBorderColor_Click;
            comboThickness.SelectionChanged += ComboThickness_SelectionChanged;
            tbtnEnableThickness.Click += TbtnEnableThickness_Click;
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
        internal void initializeUserControl(Microsoft.Office.Interop.Word.Document _doc)
        {
            doc = _doc;
            selection = _doc.ActiveWindow.Selection;

            DedicatedFunctions.AccessType accessType = DedicatedFunctions.hasAccess(doc);

            //get values
            Universities university;
            TemplateTypes templateType;
            Bitmap universityIcon = null;
            string department;
            string group;
            string fieldOfStudy;
            string title;
            string supervisor;
            string advisor;
            string author;
            string defenseDate;
            string academicDegree;
            string documentType;

            if (accessType == DedicatedFunctions.AccessType.AccessGranted)
            {
                university = DedicatedFunctions.getUniversity(doc);
                templateType = DedicatedFunctions.getTemplateType(doc);
                universityIcon = TemplateAccess.getUniversityIcon(university, templateType);
                department = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Department_Fa.ToString());
                group = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Group_Fa.ToString());
                fieldOfStudy = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_FieldOfStudy_Fa.ToString());
                title = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Title_Fa.ToString());
                supervisor = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Supervisor_Fa.ToString());
                advisor = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Advisor_Fa.ToString());
                author = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_Author_Fa.ToString());
                defenseDate = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_DefenseDate_Fa.ToString());
                academicDegree = DedicatedFunctions.getStaticVariableValue(doc, VariableFieldIDs._variable_field_AcademicDegree_Fa.ToString());
                documentType = DedicatedFunctions.getDocumentTypePersianName(doc);
            }
            else
            {
                university = Universities.YazdUniversity;
                templateType = TemplateTypes.UniversityTemplate;
                universityIcon = TemplateAccess.getUniversityIcon(university, templateType);
                department = "دانشکده";
                group = "گروه";
                fieldOfStudy = "رشته تحصیلی";
                title = "عنوان";
                supervisor = "استاد مشاور";
                advisor = "استاد راهنما";
                author = "نگارنده";
                defenseDate = "تاریخ دفاع";
                academicDegree = "مقطع تحصیلی";
                documentType = "نوع سند";
            }


            Microsoft.Office.Interop.Word.Shape shape = createCD(universityIcon, department, group, fieldOfStudy, title, supervisor, advisor, author, defenseDate, academicDegree, documentType);

            System.Drawing.Color colorBackground = System.Drawing.Color.FromArgb(doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].Fill.ForeColor.RGB);
            bdBackgroundColor.Background = toMediaBrush(System.Drawing.Color.FromArgb(255, colorBackground.B, colorBackground.G, colorBackground.R));
            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            doc.ActiveWindow.ScrollIntoView(shape);
        }

        public static System.Windows.Media.Color toMediaColor(System.Drawing.Color color)
        {
            return System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B);
        }
        public static System.Windows.Media.Brush toMediaBrush(System.Drawing.Color color)
        {
            return new SolidColorBrush(toMediaColor(System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B)));
        }
        private WdColor ConvertBrushToWdColor(System.Windows.Media.Brush brush)
        {
            System.Windows.Media.Color color = (brush as SolidColorBrush).Color;
            return (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B));
        }
        public System.Drawing.Color WdColorToColor(WdColor wdColor)
        {
            // تبدیل WdColor به رنگ RGB
            int rgb = (int)wdColor;

            // استخراج مقادیر رنگ قرمز، سبز و آبی
            int red = (rgb & 0xFF);
            int green = (rgb >> 8) & 0xFF;
            int blue = (rgb >> 16) & 0xFF;

            // ایجاد و بازگشت رنگ System.Drawing.Color
            return System.Drawing.Color.FromArgb(red, green, blue);
        }


        private void BtnInsert_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            insert();
        }

        private void BtnRefreshList_Click(object sender, System.EventArgs e)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("حالت پیشفرض طرح سیدی");

            changeStatusThickness(false);
            comboThickness.SelectedIndex = 0;
            comboCDCount.SelectedIndex = 0;

            changeTextColor(System.Drawing.Color.FromArgb(255, 255, 255, 255));
            changeBackgroundColor(System.Drawing.Color.FromArgb(255, 55, 47, 148));
            changeBorderColor(System.Drawing.Color.FromArgb(255, 0, 0, 0));
            changeIconColor(System.Drawing.Color.FromArgb(255, 255, 255, 255));

            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }
        private void BtnClose_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("لغو خروجی گرفتن از طرح سیدی");
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            CloseTaskPaneRequest.Invoke();

            Globals.ThisAddIn.Application.ScreenUpdating = true;
            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
        }


        private void TbtnEnableThickness_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            changeStatusThickness(tbtnEnableThickness.IsChecked == true ? true : false);
        }
        private void ComboThickness_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (comboThickness.SelectedIndex < 0)
                return;

            bool canvasIsExist = false;
            bool shapeIsExist = false;
            try
            {
                int getID3 = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].ID;
                canvasIsExist = true;
                int getID = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].ID;
                int getID2 = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].ID;
                shapeIsExist = true;
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("شکل ساخته شده توسط شما پاک شده و عملیات لغو میشود");
                CloseTaskPaneRequest.Invoke();
            }

            if (shapeIsExist)
            {
                int value = comboThickness.SelectedIndex + 1;

                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر ضخامت خط دور سیدی به " + value.ToString());

                doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].Line.Weight = value;
                doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].Line.Weight = value;
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }

        private void BtnChangeBackgroundColor_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            System.Windows.Forms.ColorDialog colorDialog = new System.Windows.Forms.ColorDialog();
            System.Windows.Forms.DialogResult result = colorDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                changeBackgroundColor(colorDialog.Color);
            }
        }
        private void BtnChangeBorderColor_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            System.Windows.Forms.ColorDialog colorDialog = new System.Windows.Forms.ColorDialog();
            System.Windows.Forms.DialogResult result = colorDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                changeBorderColor(colorDialog.Color);
            }
        }
        private void BtnChangeIconColor_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            System.Windows.Forms.ColorDialog colorDialog = new System.Windows.Forms.ColorDialog();
            System.Windows.Forms.DialogResult result = colorDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                changeIconColor(colorDialog.Color);
            }
        }
        private void BtnChangeTextColor_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            System.Windows.Forms.ColorDialog colorDialog = new System.Windows.Forms.ColorDialog();
            System.Windows.Forms.DialogResult result = colorDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                changeTextColor(colorDialog.Color);
            }
        }

        private void BtnInsertSources_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Globals.ThisAddIn.insertSources();
        }
        private void BtnImportSources_Click(object sender, System.Windows.RoutedEventArgs e)
        {
            Globals.ThisAddIn.importSourcesDialog();
        }


        private void changeTextColor(System.Drawing.Color color)
        {
            bool shapeIsExist = false;
            try
            {
                int getID = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].ID;
                shapeIsExist = true;
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("شکل ساخته شده توسط شما پاک شده و عملیات لغو میشود");
                CloseTaskPaneRequest.Invoke();
            }

            if (shapeIsExist)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر رنگ متن در طرح سیدی");

                WdColor wdColor = (WdColor)ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, color.B, color.G, color.R));

                System.Drawing.Color correctColor = WdColorToColor(wdColor);
                bdTextColor.Background = toMediaBrush(System.Drawing.Color.FromArgb(255, correctColor.R, correctColor.G, correctColor.B));

                foreach (Shape item in doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems)
                {
                    if (item.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                    {
                        item.TextFrame.TextRange.Font.Color = wdColor;
                    }
                }
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        private void changeBackgroundColor(System.Drawing.Color color)
        {
            bool canvasIsExist = false;
            bool shapeIsExist = false;
            try
            {
                int getID = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].ID;
                canvasIsExist = true;
                int getID2 = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].ID;
                shapeIsExist = true;
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("شکل ساخته شده توسط شما پاک شده و عملیات لغو میشود");
                CloseTaskPaneRequest.Invoke();
            }

            if (shapeIsExist)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر رنگ پس زمینه سیدی");
                bdBackgroundColor.Background = toMediaBrush(System.Drawing.Color.FromArgb(255, color.R, color.G, color.B));
                doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, color.B, color.G, color.R).ToArgb();
                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        private void changeBorderColor(System.Drawing.Color color)
        {
            bool canvasIsExist = false;
            bool shapeIsExist = false;
            try
            {
                int getID3 = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].ID;
                canvasIsExist = true;
                int getID = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].ID;
                int getID2 = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].ID;
                shapeIsExist = true;
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("شکل ساخته شده توسط شما پاک شده و عملیات لغو میشود");
                CloseTaskPaneRequest.Invoke();
            }
            if (shapeIsExist)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر رنگ پس زمینه سیدی");

                bdBorderColor.Background = toMediaBrush(System.Drawing.Color.FromArgb(255, color.R, color.G, color.B));
                doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].Line.ForeColor.RGB = System.Drawing.Color.FromArgb(255, color.B, color.G, color.R).ToArgb();
                doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].Line.ForeColor.RGB = System.Drawing.Color.FromArgb(255, color.B, color.G, color.R).ToArgb();

                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        private void changeIconColor(System.Drawing.Color color)
        {
            bool canvasIsExist = false;
            bool shapeIsExist = false;
            try
            {
                int getID2 = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].ID;
                canvasIsExist = true;
                int getID = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_UniversityIcon_.ToString()].ID;
                shapeIsExist = true;
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("شکل ساخته شده توسط شما پاک شده و عملیات لغو میشود");
                CloseTaskPaneRequest.Invoke();
            }

            if (shapeIsExist)
            {
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر رنگ آیکون دانشگاه در طرح سیدی");

                bdUniversityIconColor.Background = toMediaBrush(System.Drawing.Color.FromArgb(255, color.R, color.G, color.B));

                //float width = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_UniversityIcon_.ToString()].Width;
                //float height = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_UniversityIcon_.ToString()].Height;
                doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_UniversityIcon_.ToString()].Delete();

                float dimensionCD = Globals.ThisAddIn.Application.MillimetersToPoints(116);
                float lineWeight = Globals.ThisAddIn.Application.PixelsToPoints(1);

                // insert Icon
                universityIcon.MakeTransparent();

                for (int y = 0; (y <= (universityIcon.Height - 1)); y++)// for invert Colors
                {
                    for (int x = 0; (x <= (universityIcon.Width - 1)); x++)
                    {
                        System.Drawing.Color inv = universityIcon.GetPixel(x, y);
                        //inv = System.Drawing.Color.FromArgb(inv.A, (255 - inv.R), (255 - inv.G), (255 - inv.B));
                        inv = System.Drawing.Color.FromArgb(inv.A, color.R, color.G, color.B);
                        universityIcon.SetPixel(x, y, inv);
                    }
                }

                string tempImageUniversityPath = DedicatedFunctions.copyImageToTempFolder(universityIcon, nameof(universityIcon));

                float aspectRatio = Math.Min((float)universityIcon.Width / (float)universityIcon.Height, (float)universityIcon.Height / (float)universityIcon.Width);
                float scaleUniversityIcon = Globals.ThisAddIn.Application.PointsToPixels(Globals.ThisAddIn.Application.MillimetersToPoints(21)) / universityIcon.Height;
                float heightIcon = universityIcon.Height * scaleUniversityIcon;
                float widthIcon = Globals.ThisAddIn.Application.PixelsToPoints(heightIcon * aspectRatio);

                Microsoft.Office.Interop.Word.Shape cdUniversityIcon = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems.AddPicture(tempImageUniversityPath,
                                                                         false,
                                                                         true,
                                                                         (lineWeight + (dimensionCD / 2)) - (widthIcon / 2),
                                                                         lineWeight + Globals.ThisAddIn.Application.MillimetersToPoints(1),
                                                                         widthIcon,
                                                                         Globals.ThisAddIn.Application.PixelsToPoints(heightIcon));
                cdUniversityIcon.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                cdUniversityIcon.Name = ShapeIDs._CD_UniversityIcon_.ToString();
                DedicatedFunctions.removeFileFromSystem(tempImageUniversityPath);


                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }
        private void changeStatusThickness(bool enable)
        {
            bool canvasIsExist = false;
            bool shapeIsExist = false;
            try
            {
                int getID3 = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].ID;
                canvasIsExist = true;
                int getID = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].ID;
                int getID2 = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].ID;
                shapeIsExist = true;
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("شکل ساخته شده توسط شما پاک شده و عملیات لغو میشود");
                CloseTaskPaneRequest.Invoke();
            }
            if (shapeIsExist)
            {
                if (enable)
                {
                    tbtnEnableThickness.IsChecked = true;
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("فعال کردن خط دور سیدی");

                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].Line.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 0, 0, 0).ToArgb();
                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].Line.Weight = 1;

                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].Line.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 0, 0, 0).ToArgb();
                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].Line.Weight = 1;

                    comboThickness.SelectedIndex = 0;
                    bdBorderColor.Background = toMediaBrush(System.Drawing.Color.FromArgb(255, 0, 0, 0));
                    comboThickness.IsEnabled = true;
                    btnChangeBorderColor.IsEnabled = true;

                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                }
                else
                {
                    tbtnEnableThickness.IsChecked = false;
                    Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("غیر فعال کردن خط دور سیدی");
                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeOutline_.ToString()].Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].CanvasItems[ShapeIDs._CD_CirculeInline_.ToString()].Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                    comboThickness.IsEnabled = false;
                    btnChangeBorderColor.IsEnabled = false;
                    Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                }
            }
        }

        internal Microsoft.Office.Interop.Word.Shape createCD(Bitmap UniversityIcon,
                               string departmentFa,
                               string txtGroup,
                               string txtFieldFa,
                               string txtTitleFa,
                               string txtSupervisorFa,
                               string txtAdvisorFa,
                               string txtAuthorFa,
                               string txtDefenseDateFa,
                               string academicDegree,
                               string type)
        {
            this.universityIcon = UniversityIcon;
            float dimensionCD = Globals.ThisAddIn.Application.MillimetersToPoints(116);
            float dimensionInnerCircle = Globals.ThisAddIn.Application.MillimetersToPoints(23);
            float lineWeight = Globals.ThisAddIn.Application.PixelsToPoints(1);


            // add Canvas
            Microsoft.Office.Interop.Word.Shape cdCanvas = doc.Shapes.AddCanvas(0, 0, dimensionCD, dimensionCD);
            cdCanvas.ConvertToInlineShape();
            cdCanvas.Name = ShapeIDs._Pilot_CreateCD_.ToString();

            cdCanvas.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
            // add OutlineLine Circule
            Microsoft.Office.Interop.Word.Shape cdOutLineCircule = cdCanvas.CanvasItems.AddShape((int)Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                                                                   lineWeight / 2,
                                                                   lineWeight / 2,
                                                                   dimensionCD - lineWeight,
                                                                   dimensionCD - lineWeight);
            cdOutLineCircule.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 148, 47, 55).ToArgb();
            bdBackgroundColor.Background = toMediaBrush(System.Drawing.Color.FromArgb(255, 55, 47, 148));
            cdOutLineCircule.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            cdOutLineCircule.Name = ShapeIDs._CD_CirculeOutline_.ToString();

            // add inLine Circule
            float cdInLineTopLeft = (lineWeight + (dimensionCD / 2)) - dimensionInnerCircle / 2;
            Microsoft.Office.Interop.Word.Shape cdInLineCircule = cdCanvas.CanvasItems.AddShape((int)Microsoft.Office.Core.MsoAutoShapeType.msoShapeOval,
                                                                  cdInLineTopLeft,
                                                                  cdInLineTopLeft,
                                                                  dimensionInnerCircle - lineWeight,
                                                                  dimensionInnerCircle - lineWeight);
            cdInLineCircule.Fill.ForeColor.RGB = System.Drawing.Color.FromArgb(255, 255, 255, 255).ToArgb();
            cdInLineCircule.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            cdInLineCircule.Name = ShapeIDs._CD_CirculeInline_.ToString();



            // insert Icon
            bdUniversityIconColor.Background = toMediaBrush(System.Drawing.Color.FromArgb(255, 255, 255, 255));
            Bitmap universityIcon = UniversityIcon;
            universityIcon.MakeTransparent();
            for (int y = 0; (y <= (universityIcon.Height - 1)); y++)// for invert Colors
            {
                for (int x = 0; (x <= (universityIcon.Width - 1)); x++)
                {
                    System.Drawing.Color inv = universityIcon.GetPixel(x, y);
                    //inv = System.Drawing.Color.FromArgb(inv.A, (255 - inv.R), (255 - inv.G), (255 - inv.B));
                    inv = System.Drawing.Color.FromArgb(inv.A, 255, 255, 255);
                    universityIcon.SetPixel(x, y, inv);
                }
            }

            float aspectRatio = Math.Min((float)universityIcon.Width / (float)universityIcon.Height, (float)universityIcon.Height / (float)universityIcon.Width);
            float scaleUniversityIcon = Globals.ThisAddIn.Application.PointsToPixels(Globals.ThisAddIn.Application.MillimetersToPoints(21)) / universityIcon.Height;
            float heightIcon = universityIcon.Height * scaleUniversityIcon;
            float widthIcon = Globals.ThisAddIn.Application.PixelsToPoints(heightIcon * aspectRatio);
            float leftIcon = (lineWeight + (dimensionCD / 2)) - (widthIcon / 2);
            float topIcon = lineWeight + Globals.ThisAddIn.Application.MillimetersToPoints(1);
            try
            {
                string tempImageUniversityPath = DedicatedFunctions.copyImageToTempFolder(universityIcon, nameof(universityIcon));

                Microsoft.Office.Interop.Word.Shape cdUniversityIcon = cdCanvas.CanvasItems.AddPicture(tempImageUniversityPath,
                                                                         false,
                                                                         true,
                                                                         leftIcon,
                                                                         topIcon,
                                                                         widthIcon,
                                                                         Globals.ThisAddIn.Application.PixelsToPoints(heightIcon));
                cdUniversityIcon.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                cdUniversityIcon.Name = ShapeIDs._CD_UniversityIcon_.ToString();
                DedicatedFunctions.removeFileFromSystem(tempImageUniversityPath);
            }
            catch (Exception e)
            {
                DedicatedFunctions.ShowMessage("خطا در درج عکس نماد دانشگاه در طرح لوح فشرده" + "\nپیغام خطا:" + "\n" + e.Message);
            }


            //add university Informations
            float textWidth = Globals.ThisAddIn.Application.MillimetersToPoints(80);
            float textHeight = Globals.ThisAddIn.Application.MillimetersToPoints(25);
            float leftTextBox = (lineWeight + (dimensionCD / 2)) - textWidth / 2;
            float topTextBox = topIcon + Globals.ThisAddIn.Application.PixelsToPoints(heightIcon);
            Microsoft.Office.Interop.Word.Shape textBox = cdCanvas.CanvasItems.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                            leftTextBox,
                                                            topTextBox,
                                                            textWidth,
                                                            textHeight);
            textBox.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            textBox.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            Range rng = textBox.TextFrame.TextRange;
            rng.set_Style(StyleNames.styleNormal);
            rng.Font.Color = WdColor.wdColorWhite;
            rng.Font.Size = 7;
            rng.Font.SizeBi = 7;
            rng.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng.Font.BoldBi = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rng.ParagraphFormat.Space1();
            rng.ParagraphFormat.SpaceBefore = 0;
            rng.ParagraphFormat.SpaceAfter = 0;

            string text;
            //if (facultyFa.Trim().ToLower().Contains("پردیس".ToLower().Trim()) || facultyFa.Trim().ToLower().Contains("دانشکده".ToLower().Trim()))
            //    text = facultyFa + "\n";
            //else
            //    text = "پردیس " + facultyFa + "\n";

            if (departmentFa.Trim().ToLower().Contains("ساختمان".ToLower().Trim()) || departmentFa.Trim().ToLower().Contains("دانشکده".ToLower().Trim()))
                text = departmentFa + "\n";
            else
                text = "دانشکده " + departmentFa + "\n";

            if (txtGroup.Trim().ToLower().Contains("گروه".ToLower().Trim()))
                text += txtGroup + "\n";
            else
                text += "گروه " + txtGroup + "\n";

            text += type + "\n";
            text += "برای دریافت درجه " + academicDegree + "\n";
            text += "رشته " + txtFieldFa;
            rng.Text = text;


            //add caption of Those involved
            float textWidth2 = Globals.ThisAddIn.Application.MillimetersToPoints(45);
            float textHeight2 = Globals.ThisAddIn.Application.MillimetersToPoints(26);
            float leftTextBox2 = cdInLineTopLeft + dimensionInnerCircle + Globals.ThisAddIn.Application.MillimetersToPoints(2);
            float topTextBox2 = cdInLineTopLeft - Globals.ThisAddIn.Application.MillimetersToPoints(2);
            Microsoft.Office.Interop.Word.Shape textBox2 = cdCanvas.CanvasItems.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                             leftTextBox2,
                                                             topTextBox2,
                                                             textWidth2,
                                                             textHeight2);
            textBox2.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            textBox2.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            Range rng2 = textBox2.TextFrame.TextRange;
            rng2.set_Style(StyleNames.styleNormal);
            rng2.Font.Color = WdColor.wdColorWhite;
            rng2.Font.Size = 10;
            rng2.Font.SizeBi = 10;
            rng2.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng2.Font.BoldBi = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng2.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rng2.ParagraphFormat.Space1();
            rng2.ParagraphFormat.SpaceBefore = 0;
            rng2.ParagraphFormat.SpaceAfter = 0;
            rng2.Text = "استاد راهنما:" + "\n" + txtSupervisorFa.Replace(",", "\n").Replace(";", "\n");


            //add Those involved
            float textWidth3 = Globals.ThisAddIn.Application.MillimetersToPoints(45);
            float textHeight3 = Globals.ThisAddIn.Application.MillimetersToPoints(26);
            float leftTextBox3 = cdInLineTopLeft - textWidth3 - Globals.ThisAddIn.Application.MillimetersToPoints(2);
            float topTextBox3 = cdInLineTopLeft - Globals.ThisAddIn.Application.MillimetersToPoints(2);
            Microsoft.Office.Interop.Word.Shape textBox3 = cdCanvas.CanvasItems.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                             leftTextBox3,
                                                             topTextBox3,
                                                             textWidth3,
                                                             textHeight3);
            textBox3.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            textBox3.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            Range rng3 = textBox3.TextFrame.TextRange;
            rng3.set_Style(StyleNames.styleNormal);
            rng3.Font.Color = WdColor.wdColorWhite;
            rng3.Font.Size = 10;
            rng3.Font.SizeBi = 10;
            rng3.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng3.Font.BoldBi = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng3.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rng3.ParagraphFormat.Space1();
            rng3.ParagraphFormat.SpaceBefore = 0;
            rng3.ParagraphFormat.SpaceAfter = 0;
            rng3.Text = "استاد مشاور:" + "\n" + txtAdvisorFa.Replace(",", "\n").Replace(";", "\n");


            //add PersonName
            float textWidth6 = Globals.ThisAddIn.Application.MillimetersToPoints(70);
            float textHeight6 = Globals.ThisAddIn.Application.MillimetersToPoints(10);
            float leftTextBox6 = (lineWeight + (dimensionCD / 2)) - textWidth6 / 2;
            float topTextBox6 = topTextBox3 + textHeight3;
            Microsoft.Office.Interop.Word.Shape textBox6 = cdCanvas.CanvasItems.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                             leftTextBox6,
                                                             topTextBox6,
                                                             textWidth6,
                                                             textHeight6);
            textBox6.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            textBox6.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            Range rng6 = textBox6.TextFrame.TextRange;
            rng6.set_Style(StyleNames.styleNormal);
            rng6.Font.Color = WdColor.wdColorWhite;
            rng6.Font.Size = 10;
            rng6.Font.SizeBi = 10;
            rng6.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng6.Font.BoldBi = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng6.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rng6.ParagraphFormat.Space1();
            rng6.ParagraphFormat.SpaceBefore = 0;
            rng6.ParagraphFormat.SpaceAfter = 0;
            rng6.Text = "نگارش: " + txtAuthorFa;


            //add Caption
            float textWidth4 = Globals.ThisAddIn.Application.MillimetersToPoints(80);
            float textHeight4 = Globals.ThisAddIn.Application.MillimetersToPoints(27);
            float leftTextBox4 = (lineWeight + (dimensionCD / 2)) - textWidth4 / 2;
            float topTextBox4 = topTextBox6 + textHeight6 - Globals.ThisAddIn.Application.MillimetersToPoints(1);
            Microsoft.Office.Interop.Word.Shape textBox4 = cdCanvas.CanvasItems.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                             leftTextBox4,
                                                             topTextBox4,
                                                             textWidth4,
                                                             textHeight4);
            textBox4.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            textBox4.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            textBox4.Name = ShapeIDs._CD_Title_.ToString();

            Range rng4 = textBox4.TextFrame.TextRange;
            rng4.set_Style(StyleNames.styleNormal);
            rng4.Font.Color = WdColor.wdColorWhite;
            rng4.Font.Size = 10;
            rng4.Font.SizeBi = 10;
            rng4.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng4.Font.BoldBi = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng4.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rng4.ParagraphFormat.Space1();
            rng4.ParagraphFormat.SpaceBefore = 0;
            rng4.ParagraphFormat.SpaceAfter = 0;
            rng4.Text = "عنوان:" + "\n" + txtTitleFa;

            //add Date
            float textWidth5 = Globals.ThisAddIn.Application.MillimetersToPoints(40);
            float textHeight5 = Globals.ThisAddIn.Application.MillimetersToPoints(10);
            float leftTextBox5 = (lineWeight + (dimensionCD / 2)) - textWidth5 / 2;
            float topTextBox5 = topTextBox4 + textHeight4;
            Microsoft.Office.Interop.Word.Shape textBox5 = cdCanvas.CanvasItems.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                                                             leftTextBox5,
                                                             topTextBox5,
                                                             textWidth5,
                                                             textHeight5);
            textBox5.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            textBox5.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

            Range rng5 = textBox5.TextFrame.TextRange;
            rng5.set_Style(StyleNames.styleNormal);
            rng5.Font.Color = WdColor.wdColorWhite;
            rng5.Font.Size = 10;
            rng5.Font.SizeBi = 10;
            rng5.Font.Bold = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng5.Font.BoldBi = (int)Microsoft.Office.Core.MsoTriState.msoTrue;
            rng5.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            rng5.ParagraphFormat.Space1();
            rng5.ParagraphFormat.SpaceBefore = 0;
            rng5.ParagraphFormat.SpaceAfter = 0;
            rng5.Text = txtDefenseDateFa;

            return cdCanvas;
        }

        private void insert()
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;

            bool shapeIsExist = false;
            try
            {
                int getID = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()].ID;
                shapeIsExist = true;
            }
            catch (Exception)
            {
                DedicatedFunctions.ShowErrorMessage("شکل ساخته شده توسط شما پاک شده و عملیات لغو میشود");
                CloseTaskPaneRequest.Invoke();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }

            if (File.Exists(doc.FullName))
                DedicatedFunctions.saveDocument(doc);

            System.Windows.Forms.SaveFileDialog fileDialog = new System.Windows.Forms.SaveFileDialog();
            fileDialog.Title = "مکان و نام خروجی فایل خود را مشخص کنید";
            fileDialog.AddExtension = false;
            fileDialog.Filter = "Portable Document Format (*.pdf)|*.pdf";

            if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if (fileDialog.FileName != doc.FullName)
                {
                    if (shapeIsExist)
                    {
                        if (fileDialog.FileName != "")
                        {
                            int pageNumberStart = 0;
                            int pageNumberEnd = 0;

                            Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("گرفتن خروجی از طرح سیدی");
                            try
                            {
                                Shape shape = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()];
                                shape.Select();
                                Section cdSection = selection.Sections.Last;
                                pageNumberStart = (int)selection.Information[WdInformation.wdActiveEndPageNumber];
                                pageNumberEnd = pageNumberStart + cdSection.Range.ComputeStatistics(WdStatistic.wdStatisticPages) - 1;
                            }
                            catch (Exception)
                            {
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای رخ داد، عملیات لغو میشود.");
                                CloseTaskPaneRequest.Invoke();
                                Globals.ThisAddIn.Application.ScreenUpdating = true;
                                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                                return;
                            }

                            try
                            {
                                //اگه دو تا خروجی سیدی خواست
                                if (comboCDCount.SelectedIndex == 1)
                                {
                                    Shape shape = doc.Shapes[ShapeIDs._Pilot_CreateCD_.ToString()];
                                    shape.Select();
                                    Range rng = selection.FormattedText;
                                    selection.EndKey(WdUnits.wdLine);
                                    DedicatedFunctions.insertParagraph(selection, 2);
                                    DedicatedFunctions.changeStyleInSelection(selection, StyleNames.styleNormal);
                                    DedicatedFunctions.changeParagraphAlignment(selection, WdParagraphAlignment.wdAlignParagraphCenter);
                                    DedicatedFunctions.changeParagraphIndent(selection, 0, 0, 0);
                                    selection.FormattedText = rng;
                                    //if (doc.Characters.Count != 0)
                                    //    doc.Characters[2].Copy();
                                }
                            }
                            catch (Exception e)
                            {
                                DedicatedFunctions.ShowErrorMessage("خطای غیر منتظره ای رخ داد، عملیات لغو میشود.");
                                CloseTaskPaneRequest.Invoke();
                                Globals.ThisAddIn.Application.ScreenUpdating = true;
                                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                                return;
                            }
                            


                            try
                            {
                                doc.ExportAsFixedFormat(fileDialog.FileName,
                                                                         WdExportFormat.wdExportFormatPDF,
                                                                         true,
                                                                         WdExportOptimizeFor.wdExportOptimizeForPrint,
                                                                         WdExportRange.wdExportFromTo,
                                                                         pageNumberStart,
                                                                         pageNumberEnd,
                                                                         WdExportItem.wdExportDocumentContent,
                                                                         true,
                                                                         true,
                                                                         WdExportCreateBookmarks.wdExportCreateNoBookmarks,
                                                                         true,
                                                                         true,
                                                                         false
                                                                         );
                            }
                            catch (Exception)
                            {
                                DedicatedFunctions.ShowErrorMessage("مشکلی در گرفتن خروجی آمد، لطفا مجدد تلاش کنید", (int)ErrorCodes.CD_Exporting);
                                Globals.ThisAddIn.Application.ScreenUpdating = true;
                                return;
                            }


                            try
                            {
                                DesigningCDTaskPane.DeleteAndDisableActionPane(doc, null, null);
                            }
                            catch (Exception)
                            {
                                Globals.ThisAddIn.Application.ScreenUpdating = true;
                                DedicatedFunctions.ShowErrorMessage("برنامه قادر به پاک کردن صفحه ساخته شده سیدی در سند شما نشد! لطفا دستی صفحه را پاک نمایید.");
                            }

                            CloseTaskPaneRequest.Invoke();
                            Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
                        }
                        else
                        {
                            Globals.ThisAddIn.Application.ScreenUpdating = true;
                            DedicatedFunctions.ShowErrorMessage("مکان فایل انتخاب نشده است");
                        }
                    }
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("خروجی نمیتواند با سند اختصاصی شما جایگزین شود، لطفا با نامی دیگر خروجی را ذخیره کنید");
                }
            }

            Globals.ThisAddIn.Application.ScreenUpdating = true;
        }
    }
}
