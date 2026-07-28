namespace ShivaNegar.Forms.ShivaNegarManager.CreateDocument.Models
{
    internal class CreateDocumentControlModel
    {
        public enum ControlLevels
        {
            Essential,
            Optional,
        }

        private object control;
        private bool validate;
        private ControlLevels controlLevel;

        public CreateDocumentControlModel(object control, ControlLevels controlLevel)
        {
            this.Control = control;
            this.Validate = false;
            this.ControlLevel = controlLevel;
        }

        public object Control { get => control; set => control = value; }
        public bool Validate { get => validate; set => validate = value; }
        internal ControlLevels ControlLevel { get => controlLevel; set => controlLevel = value; }
    }
}
