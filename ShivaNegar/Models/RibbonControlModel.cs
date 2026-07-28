using System.Drawing;

namespace ShivaNegar.Models
{
    internal class RibbonControlModel
    {
        internal enum RibbonControlTypes
        {
            BuiltIn,
            Tab,
            Group,
            Button,
            BackstageButton,
            DialogBoxLauncherButton,
            Menu,
            Gallery,
            GalleryItem,
        }

        private string id;
        private RibbonControlTypes ribbonControlType;
        private Bitmap image;
        private bool enable;
        private bool visible;
        private string label;
        private string description;
        private string superTip;
        private string screenTip;
        private string keyTip;
        private object content;

        private string shortcutText;

        private string initialLabel;
        private string initialSuperTip;
        private string initialScreenTip;

        internal string Id { get => id; set => id = value; }
        internal RibbonControlTypes RibbonControlType { get => ribbonControlType; set => ribbonControlType = value; }
        internal Bitmap Image { get => image; set => image = value; }
        internal bool Enable { get => enable; set => enable = value; }
        internal bool Visible { get => visible; set => visible = value; }
        internal string Label { get => label; set => label = value; }
        internal string Description { get => description; set => description = value; }
        internal string SuperTip { get => superTip; set => superTip = value; }
        internal string ScreenTip { get => screenTip; set => screenTip = value; }
        internal string KeyTip { get => keyTip; set => keyTip = value; }
        internal object Content { get => content; set => content = value; }

        internal string ShortcutText { get => shortcutText; set => shortcutText = value; }
        internal string InitialLabel { get => initialLabel; set => initialLabel = value; }
        internal string InitialSuperTip { get => initialSuperTip; set => initialSuperTip = value; }
        internal string InitialScreenTip { get => initialScreenTip; set => initialScreenTip = value; }

        /// <summary>
        /// for built in Ribbon Controls inside Command (only enabled has accessed)
        /// </summary>
        /// <param name="id"></param>
        /// <param name="enable"></param>
        internal RibbonControlModel(string id, bool enable = true)
        {
            this.Id = id;
            this.Enable = enable;
        }

        internal RibbonControlModel(string id, RibbonControlTypes ribbonControlType, Bitmap image, bool enable, bool visible, string label, string screenTip, string superTip, string keyTip, object content = null)
        {
            this.Id = id;
            this.RibbonControlType = ribbonControlType;
            this.Image = image;
            this.Enable = enable;
            this.Visible = visible;
            this.Label = label;
            this.Description = "";
            this.SuperTip = superTip;
            this.ScreenTip = screenTip;
            this.KeyTip = keyTip;
            this.Content = content;

            this.InitialLabel = label;
            this.InitialSuperTip = superTip;
            this.InitialScreenTip = screenTip;
        }

    }
}
