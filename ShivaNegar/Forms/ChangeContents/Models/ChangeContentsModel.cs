using System;
namespace ShivaNegar.Forms.ChangeContents.Models
{
    public class ChangeContentsModel
    {

        private object control;

        private string contentControlName;
        private string variableName;

        private bool isOptional;

        public ChangeContentsModel(object control, string contentControlName, string variableName, bool isOptional = false)
        {
            Control = control;
            ContentControlName = contentControlName;
            VariableName = variableName;
            IsOptional = isOptional;
        }

        public bool isContentControlExist()
        {
            return ContentControlName != null;
        }
        public bool isVariableExist()
        {
            return variableName != null;
        }

        public bool useVariable()
        {
            if (isVariableExist()) return true;
            else if (isContentControlExist()) return false;
            else
                throw new ArgumentNullException();
        }

        public object Control { get => control; set => control = value; }
        public string ContentControlName { get => contentControlName; set => contentControlName = value; }
        public string VariableName { get => variableName; set => variableName = value; }
        public bool IsOptional { get => isOptional; set => isOptional = value; }

    }
}
