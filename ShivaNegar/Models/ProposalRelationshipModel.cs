using ShivaNegar.Constants;

namespace ShivaNegar.Models
{
    public class ProposalRelationshipModel
    {
        private string _resourcePath;
        private string _fileName;
        private AcademicDegrees _academicDegreeID;

        internal string FileName { get => _fileName; set => _fileName = value; }
        internal string ResourcePath { get => _resourcePath; set => _resourcePath = value; }
        internal AcademicDegrees AcademicDegreeID { get => _academicDegreeID; set => _academicDegreeID = value; }

        internal ProposalRelationshipModel(AcademicDegrees academicDegreeID, string fileName, string resourcePath)
        {
            FileName = fileName;
            ResourcePath = resourcePath;
            AcademicDegreeID = academicDegreeID;
        }
    }
}
