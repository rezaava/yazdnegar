namespace ShivaNegar.Models
{
    internal class KeyboardRelationModel
    {
        private string callFunctionName;
        private int keyboardShortcuts;
        private string ribbonControlID;
        private string ribbonSupertipContent;
        private int ribbonItemIndex;

        public KeyboardRelationModel(string callFunctionName, int keyboardShortcuts, string ribbonControlID, string ribbonSupertipContent, int ribbonItemIndex = -1)
        {
            this.CallFunctionName = callFunctionName;
            this.KeyboardShortcut = keyboardShortcuts;
            this.RibbonControlID = ribbonControlID;
            this.RibbonScreenTipContent = ribbonSupertipContent;
            this.RibbonItemIndex = ribbonItemIndex;
        }

        public string CallFunctionName { get => callFunctionName; set => callFunctionName = value; }
        public int KeyboardShortcut { get => keyboardShortcuts; set => keyboardShortcuts = value; }
        public string RibbonControlID { get => ribbonControlID; set => ribbonControlID = value; }
        public string RibbonScreenTipContent { get => ribbonSupertipContent; set => ribbonSupertipContent = value; }
        public int RibbonItemIndex { get => ribbonItemIndex; set => ribbonItemIndex = value; }
    }
}
