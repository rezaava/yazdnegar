using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Models;

namespace ShivaNegar.Forms.AddRemovePages.Models
{
    public class AddRemovePageRelationModel : TemplateRelationshipModel
    {
        private bool isPageExist;
        private object control;

        internal AddRemovePageRelationModel(Document doc, int order, string id, string fileName, string resourcePath, bool isRequired) : base(order, id, fileName, resourcePath, isRequired)
        {
            this.IsPageExist = pageExist(doc, PageID);

            if (this.SubTemplateType == SubTemplateTypes.Chapter || this.SubTemplateType == SubTemplateTypes.ChapterNotMade)
            {
                if (IsPageExist)
                    this.SubTemplateType = SubTemplateTypes.Chapter;
                else
                    this.SubTemplateType = SubTemplateTypes.ChapterNotMade;
            }
        }

        internal bool IsPageExist { get => isPageExist; set => isPageExist = value; }
        internal object Control { get => control; set => control = value; }

        internal static bool pageExist(Document doc, PageIDs pageID)
        {
            string variableValue = DedicatedFunctions.getStaticVariableValue(doc, pageID.ToString());
            if (variableValue == SettingValues.NotExist)
                return false;
            else
                return true;
        }
    }
}
