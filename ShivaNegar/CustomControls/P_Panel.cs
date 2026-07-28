using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    public enum PaletteDrawBorders
    {
        TopLeft,
        TopRight,
        BottomLeft,
        BottomRight,
        TopBottomLeft,
        TopBottomRight,
        TopLeftRight,
        BottomLeftRight,
        All,
    }
    public class P_Panel : Panel
    {
        private Color borderColor = Color.MediumSlateBlue;
        private int borderRadius = 0;
        private int borderSize = 1;
        private PaletteDrawBorders paletteDrawBorders = PaletteDrawBorders.All;

        private DashStyle borderLineStyle = DashStyle.Solid;
        private DashCap borderCapStyle = DashCap.Flat;

        #region -> Properties
        [Category("Dedicated")]
        public Color BorderColor
        {
            get { return borderColor; }
            set
            {
                borderColor = value;
                this.Invalidate();
            }
        }
        [Category("Dedicated")]
        public int BorderSize
        {
            get { return borderSize; }
            set
            {
                if (value >= 0)
                {
                    borderSize = value;
                    this.Invalidate();
                }
            }
        }
        [Category("Dedicated")]
        public int BorderRadius
        {
            get { return borderRadius; }
            set
            {
                if (value >= 0)
                {
                    borderRadius = value;
                    this.Invalidate();//Redraw control
                }
            }
        }

        [Category("Dedicated")]
        public PaletteDrawBorders PaletteDrawBorder
        {
            get => paletteDrawBorders; set => paletteDrawBorders = value;
        }

        [Category("Dedicated")]
        public DashCap BorderCapStyle
        {
            get
            {
                return borderCapStyle;
            }

            set
            {
                borderCapStyle = value;
                this.Invalidate();
            }
        }
        [Category("Dedicated")]
        public DashStyle BorderLineStyle
        {
            get
            {
                return borderLineStyle;
            }

            set
            {
                borderLineStyle = value;
                this.Invalidate();
            }
        }
        #endregion

        public P_Panel()
        {
            borderColor = this.BackColor;
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            this.SuspendLayout();
            base.OnPaint(e);
            Graphics graph = e.Graphics;

            if (BorderSize >= 1)
            {
                if (borderRadius > 1)//Rounded TextBox
                {
                    //-Fields
                    var rectBorderSmooth = this.ClientRectangle;
                    var rectBorder = Rectangle.Inflate(rectBorderSmooth, -borderSize, -borderSize);
                    int smoothSize = borderSize > 0 ? borderSize : 1;

                    using (GraphicsPath pathBorderSmooth = GetFigurePath(rectBorderSmooth, borderRadius))
                    using (GraphicsPath pathBorder = GetFigurePath(rectBorder, borderRadius - borderSize))
                    using (Pen penBorderSmooth = new Pen(this.Parent.BackColor, smoothSize))
                    using (Pen penBorder = new Pen(BorderColor, borderSize))
                    {
                        //-Drawing
                        this.Region = new Region(pathBorderSmooth);//Set the rounded region of UserControl
                        graph.SmoothingMode = SmoothingMode.AntiAlias;
                        penBorder.Alignment = System.Drawing.Drawing2D.PenAlignment.Center;
                        penBorder.DashStyle = borderLineStyle;
                        penBorder.DashCap = borderCapStyle;


                        //Draw border smoothing
                        graph.DrawPath(penBorderSmooth, pathBorderSmooth);
                        //Draw border
                        graph.DrawPath(penBorder, pathBorder);
                    }
                }
                else //Square/Normal TextBox
                {
                    //Draw border
                    using (Pen penBorder = new Pen(BorderColor, borderSize))
                    {
                        this.Region = new Region(this.ClientRectangle);
                        penBorder.Alignment = PenAlignment.Inset;
                        penBorder.DashStyle = borderLineStyle;
                        penBorder.DashCap = borderCapStyle;
                        graph.DrawRectangle(penBorder, 0, 0, this.Width - 0.5F, this.Height - 0.5F);
                    }
                }
            }
            this.ResumeLayout(false);
        }

        private GraphicsPath GetFigurePath(Rectangle rect, int radius)
        {

            GraphicsPath borderPath = new GraphicsPath();
            float curveSize = radius * 2F;

            borderPath.StartFigure();
            RectangleF rectF = new RectangleF(rect.X, rect.Y, rect.Width, rect.Height);

            // Add only the border for drawing
            switch (paletteDrawBorders)
            {
                case PaletteDrawBorders.TopLeft:
                    borderPath.AddLine(rectF.Left, rectF.Bottom + 1, rectF.Left, rectF.Top + curveSize);
                    borderPath.AddArc(rectF.Left, rectF.Top, curveSize, curveSize, 180f, 90f);
                    borderPath.AddLine(rectF.Left + curveSize, rectF.Top, rectF.Right + 1, rectF.Top);
                    borderPath.AddLine(rectF.Right, rectF.Top, rectF.Right, rectF.Bottom);
                    break;
                case PaletteDrawBorders.TopRight:
                    borderPath.AddLine(rectF.Left - 1, rectF.Top, rectF.Right - curveSize, rectF.Top);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Top, curveSize, curveSize, -90f, 90f);
                    borderPath.AddLine(rectF.Right, rectF.Top + curveSize, rectF.Right, rectF.Bottom + 1);
                    borderPath.AddLine(rectF.Right, rectF.Bottom, rectF.Left, rectF.Bottom);
                    break;
                case PaletteDrawBorders.BottomRight:
                    borderPath.AddLine(rectF.Right, rectF.Top - 1, rectF.Right, rectF.Bottom - curveSize);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Bottom - curveSize, curveSize, curveSize, 0f, 90f);
                    borderPath.AddLine(rectF.Right - curveSize, rectF.Bottom, rectF.Left - 1, rectF.Bottom);
                    borderPath.AddLine(rectF.Left, rectF.Bottom, rectF.Left, rectF.Top);
                    break;
                case PaletteDrawBorders.BottomLeft:
                    borderPath.AddLine(rectF.Right + 1, rectF.Bottom, rectF.Left + curveSize, rectF.Bottom);
                    borderPath.AddArc(rectF.Left, rectF.Bottom - curveSize, curveSize, curveSize, 90f, 90f);
                    borderPath.AddLine(rectF.Left, rectF.Bottom - curveSize, rectF.Left, rectF.Top - 1);
                    borderPath.AddLine(rectF.Left, rectF.Top, rectF.Right, rectF.Top);
                    break;
                case PaletteDrawBorders.TopBottomLeft:
                    borderPath.AddLine(rectF.Right + 1, rectF.Bottom, rectF.Left + curveSize, rectF.Bottom);
                    borderPath.AddArc(rectF.Left, rectF.Bottom - curveSize, curveSize, curveSize, 90f, 90f);
                    borderPath.AddArc(rectF.Left, rectF.Top, curveSize, curveSize, 180f, 90f);
                    borderPath.AddLine(rectF.Left + curveSize, rectF.Top, rectF.Right + 1, rectF.Top);
                    break;
                case PaletteDrawBorders.TopBottomRight:
                    borderPath.AddLine(rectF.Left - 1, rectF.Top, rectF.Right - curveSize, rectF.Top);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Top, curveSize, curveSize, -90f, 90f);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Bottom - curveSize, curveSize, curveSize, 0f, 90f);
                    borderPath.AddLine(rectF.Right - curveSize, rectF.Bottom, rectF.Left - 1, rectF.Bottom);
                    break;
                case PaletteDrawBorders.TopLeftRight:
                    borderPath.AddLine(rectF.Left, rectF.Bottom + 1, rectF.Left, rectF.Top + curveSize);
                    borderPath.AddArc(rectF.Left, rectF.Top, curveSize, curveSize, 180f, 90f);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Top, curveSize, curveSize, -90f, 90f);
                    borderPath.AddLine(rectF.Right, rectF.Top + curveSize, rectF.Right, rectF.Bottom + 1);
                    break;
                case PaletteDrawBorders.BottomLeftRight:
                    borderPath.AddLine(rectF.Right, rectF.Top - 1, rectF.Right, rectF.Bottom - curveSize);
                    borderPath.AddArc(rectF.Right - curveSize, rectF.Bottom - curveSize, curveSize, curveSize, 0f, 90f);
                    borderPath.AddArc(rectF.Left, rectF.Bottom - curveSize, curveSize, curveSize, 90f, 90f);
                    borderPath.AddLine(rectF.Left, rectF.Bottom - curveSize, rectF.Left, rectF.Top - 1);
                    break;
                case PaletteDrawBorders.All:
                    //borderPath.AddArc(rectF.Left, rectF.Top, curveSize, curveSize, 180f, 90f);
                    //borderPath.AddArc(rectF.Right - curveSize, rectF.Top, curveSize, curveSize, 270f, 90f);
                    //borderPath.AddArc(rectF.Right - curveSize, rectF.Bottom - curveSize, curveSize, curveSize, 0f, 90f);
                    //borderPath.AddArc(rectF.Left, rectF.Bottom - curveSize, curveSize, curveSize, 90f, 90f);

                    borderPath.AddArc(rect.X, rect.Y, curveSize, curveSize, 180, 90);
                    borderPath.AddArc(rect.Right - curveSize, rect.Y, curveSize, curveSize, 270, 90);
                    borderPath.AddArc(rect.Right - curveSize, rect.Bottom - curveSize, curveSize, curveSize, 0, 90);
                    borderPath.AddArc(rect.X, rect.Bottom - curveSize, curveSize, curveSize, 90, 90);
                    break;
            }
            borderPath.CloseFigure();

            return borderPath;
        }

    }
}
