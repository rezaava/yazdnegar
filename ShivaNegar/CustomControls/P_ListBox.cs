using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    public class P_ListBox : System.Windows.Forms.ListBox
    {
        public enum _TxtAlignment
        {
            Left,
            Center,
            Right
        };


        public enum _ItemHeightType
        {
            Automatic,
            Manual
        };

        public P_ListBox()
        {
            // Set required properties to act as an owner draw list box
            base.Size = Size.Empty; // protected override Size DefaultSize
            base.BorderStyle = BorderStyle.None;
            base.IntegralHeight = false;
            base.MultiColumn = false;
            base.DrawMode = DrawMode.OwnerDrawVariable;
            ItemHeightType = _ItemHeightType.Automatic;

            base.Padding = new Padding(1);
        }

        [Category("Dedicated")]
        public _TxtAlignment TextAlignment { get; set; }

        /// <summary>
        /// Automatic: Set by Height of Font
        /// Manual: Set by User
        /// </summary>
        [Category("Dedicated")]
        public _ItemHeightType ItemHeightType { get; set; }

        /// <summary>
        /// Gets the default size of the control.
        /// </summary>
        protected override Size DefaultSize
        {
            get { return new Size(120, 96); }
        }

        /// <summary>
        /// Gets and sets the drawing mode of the checked list box.
        /// </summary>
        public override DrawMode DrawMode
        {
            get { return DrawMode.OwnerDrawVariable; }
            set { }
        }

        /// <summary>
        /// Force the remeasure of items so they are sized correctly.
        /// </summary>
        public void RefreshItemSizes()
        {
            base.DrawMode = DrawMode.OwnerDrawFixed;
            base.DrawMode = DrawMode.OwnerDrawVariable;
        }

        /// <summary>
        /// Gets and sets the public padding space.
        /// </summary>
        [DefaultValue(typeof(Padding), "1,1,1,1")]
        [Browsable(true)]
        public new Padding Padding
        {
            get { return base.Padding; }

            set
            {
                base.Padding = value;
                this.Invalidate();
            }
        }

        #region protected override
        protected override void OnPaint(PaintEventArgs e)
        {
            e.ClipRectangle.Offset(e.ClipRectangle.X, 10);
            base.OnPaint(e);
        }
        protected override void OnDrawItem(DrawItemEventArgs e)
        {

            base.OnDrawItem(e);
            //e.DrawBackground();
            //Graphics g = e.Graphics;
            //g.FillRectangle(new SolidBrush(Color.Orange), e.Bounds);
            //g.DrawString("sd", e.Font, new SolidBrush(e.ForeColor), new PointF(e.Bounds.X, e.Bounds.Y));
            //e.DrawFocusRectangle();
            TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter;

            if (TextAlignment == _TxtAlignment.Left)
                flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter;
            else if (TextAlignment == _TxtAlignment.Center)
                flags = TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter;
            else if (TextAlignment == _TxtAlignment.Right)
                flags = TextFormatFlags.Right | TextFormatFlags.VerticalCenter;

            if (e.Index >= 0)
            {
                e.DrawBackground();

                var textRect = e.Bounds;
                string itemText = DesignMode ? "Custom ListBox" : Items[e.Index].ToString();
                TextRenderer.DrawText(e.Graphics, itemText, e.Font, textRect, e.ForeColor, flags);
                e.DrawFocusRectangle();
            }

            //const TextFormatFlags flags = TextFormatFlags.Left | TextFormatFlags.VerticalCenter;
            //
            //if (e.Index >= 0)
            //{
            //    e.DrawBackground();
            //    e.Graphics.DrawRectangle(Pens.Red, 2, e.Bounds.Y + 2, 14, 14); // Simulate an icon.
            //
            //    var textRect = e.Bounds;
            //    textRect.X += 20;
            //    textRect.Width -= 20;
            //    string itemText = DesignMode ? "Custom ListBox" : Items[e.Index].ToString();
            //    TextRenderer.DrawText(e.Graphics, itemText, e.Font, textRect, e.ForeColor, flags);
            //    e.DrawFocusRectangle();
            //}
        }
        protected override void OnMeasureItem(MeasureItemEventArgs e)
        {
            if (ItemHeightType == _ItemHeightType.Automatic)
                e.ItemHeight = this.Font.Height;
            base.OnMeasureItem(e);
        }
        #endregion
    }
}