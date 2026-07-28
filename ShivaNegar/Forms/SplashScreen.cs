using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.View;

namespace ShivaNegar.Forms
{

    public partial class SplashScreen : Form
    {
        [Flags]
        enum AnimateWindowFlags
        {
            AW_HOR_POSITIVE = 0x0000000,
            AW_HOR_NEGATIVE = 0x00000002,
            AW_VER_POSITIVE = 0x00000004,
            AW_VER_NEGATIVE = 0x00000008,
            AW_CENTER = 0x00000010,
            AW_HIDE = 0x00010000,
            AW_ACTIVATE = 0x00020000,
            AW_SLIDE = 0x00040000,
            AW_BLEND = 0x00080000
        }

        [DllImport("user32.dll")]
        static extern bool AnimateWindow(IntPtr hWnd, int time, AnimateWindowFlags flags);


        #region for Windows form and Controls
        private int formBorderRadius = 8;
        private int formBorderSize = 3;
        private Color formBorderColor = Color.FromArgb(6, 174, 244);
        //private Color formBorderColor = Color.FromArgb(59, 84, 164);
        byte dotShowCounter = 0;
        #endregion

        /// <summary>
        /// A manual splash screen
        /// </summary>
        /// <param name="DurationTime">Time in Milisecond</param>
        public SplashScreen(int DurationTime = 5000)
        {
            InitializeComponent();
            this.SuspendLayout();

            #region Form initial settings
            this.FormBorderStyle = FormBorderStyle.None;
            this.Padding = new Padding(formBorderSize);
            this.BackColor = formBorderColor;
            this.Opacity = 0;
            this.TopMost = true;
            this.DoubleBuffered = true;

            string versionCustomized = BugReport.AssemblyVersion;
            versionCustomized = versionCustomized.Replace("0", "۰").Replace("1", "۱").Replace("2", "۲").Replace("3", "۳").Replace("4", "۴").Replace("5", "۵")
                .Replace("6", "۶").Replace("7", "۷").Replace("8", "۸").Replace("9", "۹");

            lblVersion.Text = "نسخه: " + versionCustomized;
            #endregion

            this.FormClosing += (e, a) =>
            {
                AnimateWindow(this.Handle, 100, AnimateWindowFlags.AW_BLEND | AnimateWindowFlags.AW_HIDE);
                Globals.ThisAddIn.DocumentManagerFormVisible = false;
            };

            #region Panel initial settings
            panelMain.BorderRadius = formBorderRadius;
            panelMain.BorderSize = 1;
            panelMain.PaletteDrawBorder = CustomControls.PaletteDrawBorders.All;
            panelMain.BorderColor = formBorderColor;
            #endregion

            #region other initialize
            timerOpenAnimation.Enabled = true;
            timerLifeTime.Interval = DurationTime;
            timerLifeTime.Enabled = true;
            #endregion
            this.ResumeLayout(false);
        }


        #region (UI) WindowsForm

        #region Round Corner Windows Form

        private GraphicsPath getRoundedPath(Rectangle rect, float radius)
        {
            GraphicsPath graphicsPath = new GraphicsPath();
            float curveSize = radius * 2f;

            graphicsPath.StartFigure();
            graphicsPath.AddArc(rect.X, rect.Y, curveSize, curveSize, 180, 90);
            graphicsPath.AddArc(rect.Right - curveSize, rect.Y, curveSize, curveSize, 270, 90);
            graphicsPath.AddArc(rect.Right - curveSize, rect.Bottom - curveSize, curveSize, curveSize, 0, 90);
            graphicsPath.AddArc(rect.X, rect.Bottom - curveSize, curveSize, curveSize, 90, 90);
            graphicsPath.CloseFigure();
            return graphicsPath;
        }
        private void formRegionAndBorder(Form form, float radius, Graphics graph, Color borderColor, float borderSize)
        {
            if (this.WindowState != FormWindowState.Minimized)
            {
                using (GraphicsPath roundPath = getRoundedPath(form.ClientRectangle, radius))
                using (Pen penBorder = new Pen(borderColor, borderSize))
                using (Matrix transform = new Matrix())
                {
                    graph.SmoothingMode = SmoothingMode.AntiAlias;
                    form.Region = new Region(roundPath);
                    if (borderSize >= 1)
                    {
                        Rectangle rect = form.ClientRectangle;
                        float scaleX = 1.0f - ((borderSize + 1) / rect.Width);
                        float scaleY = 1.0f - ((borderSize + 1) / rect.Height);

                        transform.Scale(scaleX, scaleY);
                        transform.Translate(borderSize / 1.6f, borderSize / 1.6f);

                        graph.Transform = transform;
                        graph.DrawPath(penBorder, roundPath);
                    }
                }
            }
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            formRegionAndBorder(this, formBorderRadius, e.Graphics, formBorderColor, formBorderSize);
        }
        private void ControlRegionAndBorder(Control control, float radius, Graphics graph, Color borderColor)
        {
            using (GraphicsPath roundPath = getRoundedPath(control.ClientRectangle, radius))
            using (Pen penBorder = new Pen(borderColor, 1))
            {
                graph.SmoothingMode = SmoothingMode.AntiAlias;
                control.Region = new Region(roundPath);
                graph.DrawPath(penBorder, roundPath);
            }
        }

        //Smooth Round Corner
        private void drawPath(System.Drawing.Rectangle rect, Graphics graph, Color color)
        {
            using (GraphicsPath roundPath = getRoundedPath(rect, formBorderRadius))
            using (Pen penBorder = new Pen(color, 3))
            {
                graph.DrawPath(penBorder, roundPath);
            }
        }
        protected override void OnPaintBackground(PaintEventArgs e)
        {
            base.OnPaintBackground(e);
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            //Smooth outer Border
            System.Drawing.Rectangle rectForm = this.ClientRectangle;
            int mWidth = rectForm.Width / 2;
            int mHeight = rectForm.Height / 2;
            var fbColors = GetFormBoundsColors();

            //Top Left
            drawPath(rectForm, e.Graphics, fbColors.TopLeftColor);

            //Top Right
            Rectangle rectTopRight = new System.Drawing.Rectangle(mWidth, rectForm.Y, mWidth, mHeight);
            drawPath(rectTopRight, e.Graphics, fbColors.TopRightColor);

            //Bottom Left
            Rectangle rectBottomLeft = new System.Drawing.Rectangle(rectForm.X, rectForm.X + mHeight, mWidth, mHeight);
            drawPath(rectBottomLeft, e.Graphics, fbColors.BottomLeftColor);

            //Bottom Right
            Rectangle rectBottomRight = new System.Drawing.Rectangle(mWidth, rectForm.Y + mHeight, mWidth, mHeight);
            drawPath(rectBottomRight, e.Graphics, fbColors.BottomRightColor);

        }
        private struct FormBoundsColors
        {
            public Color TopLeftColor;
            public Color TopRightColor;
            public Color BottomLeftColor;
            public Color BottomRightColor;
        }
        private FormBoundsColors GetFormBoundsColors()
        {
            var fbColor = new FormBoundsColors();
            using (var bmp = new Bitmap(1, 1))
            using (Graphics graph = Graphics.FromImage(bmp))
            {
                Rectangle rectBmp = new Rectangle(0, 0, 1, 1);
                //Top Left
                rectBmp.X = this.Bounds.X - 1;
                rectBmp.Y = this.Bounds.Y;
                graph.CopyFromScreen(rectBmp.Location, Point.Empty, rectBmp.Size);
                fbColor.TopLeftColor = bmp.GetPixel(0, 0);
                //Top Right
                rectBmp.X = this.Bounds.Right;
                rectBmp.Y = this.Bounds.Y;
                graph.CopyFromScreen(rectBmp.Location, Point.Empty, rectBmp.Size);
                fbColor.TopRightColor = bmp.GetPixel(0, 0);
                //Bottom Left
                rectBmp.X = this.Bounds.X;
                rectBmp.Y = this.Bounds.Bottom;
                graph.CopyFromScreen(rectBmp.Location, Point.Empty, rectBmp.Size);
                fbColor.BottomLeftColor = bmp.GetPixel(0, 0);
                //Bottom Right
                rectBmp.X = this.Bounds.Right;
                rectBmp.Y = this.Bounds.Bottom;
                graph.CopyFromScreen(rectBmp.Location, Point.Empty, rectBmp.Size);
                fbColor.BottomRightColor = bmp.GetPixel(0, 0);
            }
            return fbColor;
        }

        #endregion

        #endregion

        private void timerOpenAnimation_Tick(object sender, EventArgs e)
        {
            Opacity += 0.1;

            if (Opacity == 1)
            {
                timerOpenAnimation.Enabled = false;
            }
        }

        private void timerDotAnimation_Tick(object sender, EventArgs e)
        {
            if (lblStatus.Text.Contains("در حال شروع برنامه") || lblStatus.Text.Contains("لطفا منتظر بمانید"))
            {
                if (lblStatus.Text.Contains("..."))
                {
                    dotShowCounter++;
                    lblStatus.Text = lblStatus.Text.Remove(lblStatus.Text.Length - 3);
                    if (dotShowCounter >= 2)
                    {
                        lblStatus.Text = "لطفا منتظر بمانید";
                    }
                }
                else
                {
                    lblStatus.Text += ".";
                }
            }
            else
            {
                timerDotAnimation.Stop();
                timerDotAnimation.Dispose();
            }
        }

        private void timerLifeTime_Tick(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
