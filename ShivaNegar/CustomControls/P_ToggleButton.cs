using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace ShivaNegar.CustomControls
{
    public class P_ToggleButton : CheckBox
    {
        //Values
        private Color onBackColor = Color.FromArgb(6, 174, 244);
        private Color onToggleColor = Color.WhiteSmoke;
        private Color offBackColor = Color.Gray;
        private Color offToggleColor = Color.WhiteSmoke;
        private Color disableBackColor = SystemColors.InactiveCaption;
        private Color disableToggleColor = Color.WhiteSmoke;
        private bool solidStyle = true;


        //Cunstructor
        public P_ToggleButton()
        {
            this.MinimumSize = new Size(40, 18);
            this.AutoSize = false;
        }
        //Properties(Getter and Setter)
        [Category("Dedicated Style")]
        public Color OnBackColor { get => onBackColor; set { onBackColor = value; this.Invalidate(); } }
        [Category("Dedicated Style")]
        public Color OnToggleColor { get => onToggleColor; set { onToggleColor = value; this.Invalidate(); } }
        [Category("Dedicated Style")]
        public Color OffBackColor { get => offBackColor; set { offBackColor = value; this.Invalidate(); } }
        [Category("Dedicated Style")]
        public Color OffToggleColor { get => offToggleColor; set { offToggleColor = value; this.Invalidate(); } }

        [Category("Dedicated Style")]
        [DefaultValue(true)]
        public bool SolidStyle { get => solidStyle; set { solidStyle = value; this.Invalidate(); } }

        public override string Text { get => base.Text; set { } }

        [Category("Dedicated Style")]
        public Color DisableBackColor { get => disableBackColor; set => disableBackColor = value; }
        [Category("Dedicated Style")]
        public Color DisableToggleColor { get => disableToggleColor; set => disableToggleColor = value; }


        //Methods
        private GraphicsPath getFigurePath()
        {
            int arcSize = this.Height - 1;
            Rectangle leftArc = new Rectangle(0, 0, arcSize, arcSize);
            Rectangle rightArc = new Rectangle(this.Width - arcSize - 2, 0, arcSize, arcSize);

            GraphicsPath path = new GraphicsPath();
            path.StartFigure();
            path.AddArc(leftArc, 90, 180);
            path.AddArc(rightArc, 270, 180);
            path.CloseFigure();

            return path;
        }

        protected override void OnPaint(PaintEventArgs pevent)
        {
            Graphics graphic = pevent.Graphics;
            int toggleSize = this.Height - 5;
            graphic.SmoothingMode = SmoothingMode.AntiAlias;

            /*            RectangleF rectRbBorder2 = new RectangleF()
                        {
                            X = 0.5F,
                            Y = (this.Height - rbBorderSize) / 2, //Center
                            Width = rbBorderSize,
                            Height = rbBorderSize
                        };
                        RectangleF rectRbCheck2 = new RectangleF()
                        {
                            X = rectRbBorder.X + ((rectRbBorder.Width - rbCheckSize) / 2), //Center
                            Y = (this.Height - rbCheckSize) / 2, //Center
                            Width = rbCheckSize,
                            Height = rbCheckSize
                        };
            */
            graphic.Clear(this.Parent.BackColor);

            using (SolidBrush brushText = new SolidBrush(this.ForeColor))
            {
                if (this.Enabled)
                {
                    if (this.Checked) // Checked
                    {
                        //Draw the control surface
                        if (solidStyle)
                            graphic.FillPath(new SolidBrush(onBackColor), getFigurePath());
                        else
                            graphic.DrawPath(new Pen(onBackColor, 2), getFigurePath());
                        //Draw the toggle
                        graphic.FillEllipse(new SolidBrush(onToggleColor), new Rectangle(this.Width - this.Height + 1, 2, toggleSize, toggleSize));
                    }
                    else // not Checked
                    {
                        //Draw the control surface
                        if (solidStyle)
                            graphic.FillPath(new SolidBrush(offBackColor), getFigurePath());
                        else
                            graphic.DrawPath(new Pen(offBackColor, 2), getFigurePath());
                        //Draw the toggle
                        graphic.FillEllipse(new SolidBrush(offToggleColor), new Rectangle(2, 2, toggleSize, toggleSize));

                    }
                }
                else
                {
                    if (this.Checked) // Checked
                    {
                        //Draw the control surface
                        if (solidStyle)
                            graphic.FillPath(new SolidBrush(DisableBackColor), getFigurePath());
                        else
                            graphic.DrawPath(new Pen(DisableBackColor, 2), getFigurePath());
                        //Draw the toggle
                        graphic.FillEllipse(new SolidBrush(DisableToggleColor), new Rectangle(this.Width - this.Height + 1, 2, toggleSize, toggleSize));
                    }
                    else // not Checked
                    {
                        //Draw the control surface
                        if (solidStyle)
                            graphic.FillPath(new SolidBrush(DisableBackColor), getFigurePath());
                        else
                            graphic.DrawPath(new Pen(DisableBackColor, 2), getFigurePath());
                        //Draw the toggle
                        graphic.FillEllipse(new SolidBrush(DisableToggleColor), new Rectangle(2, 2, toggleSize, toggleSize));
                    }
                }
                graphic.DrawString(this.Text, this.Font, new SolidBrush(this.ForeColor), toggleSize, (this.Height - TextRenderer.MeasureText(this.Text, this.Font).Height) / 2);
            }
        }
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            //this.Width = TextRenderer.MeasureText(this.Text, this.Font).Width + 30;
        }

    }
}
