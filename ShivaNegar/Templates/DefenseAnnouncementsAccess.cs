using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml;
using ShivaNegar.Constants;

namespace ShivaNegar.Templates
{
    internal class DefenseAnnouncementsAccess
    {
        #region get template data
        internal static List<DefenseAnnouncementModel> getDefenseAnnouncementsModels()
        {
            List<DefenseAnnouncementModel> list = new List<DefenseAnnouncementModel>();

            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream daXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.DefenseAnnouncements);

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(daXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//defenseAnnouncement");
            foreach (XmlNode template in templates)
            {
                string id = template.Attributes["id"].Value;
                string fileName = template.Attributes["fileName"].Value;
                string image = template.Attributes["image"].Value;
                string path = template.Attributes["path"].Value;
                string name = template.Attributes["name"].Value;
                string sourceFrom = template.Attributes["sourceFrom"].Value;

                DefenseAnnouncementModel model = new DefenseAnnouncementModel(id, fileName, image, path, name, sourceFrom);
                list.Add(model);
            }
            return list;
        }

        internal static Stream getCoverImageStream(DefenseAnnouncementModel model)
        {
            string resourcePath = model.Path + model.Image;
            return DedicatedFunctions.getStream(resourcePath);
        }
        internal static Stream getDefenseAnnouncementFileStream(DefenseAnnouncementModel model)
        {
            string resourcePath = model.Path + model.FileName;
            return DedicatedFunctions.getStream(resourcePath);
        }
        #endregion
    }

    internal class DefenseAnnouncementModel
    {
        private string id;
        private string fileName;
        private string image;
        private string path;
        private string name;
        private string sourceFrom;

        public DefenseAnnouncementModel(string id, string fileName, string image, string path, string name, string sourceFrom)
        {
            Id = id;
            FileName = fileName;
            Image = image;
            Path = path;
            Name = name;
            SourceFrom = sourceFrom;
        }

        public string Id { get => id; set => id = value; }
        public string FileName { get => fileName; set => fileName = value; }
        public string Image { get => image; set => image = value; }
        public string Path { get => path; set => path = value; }
        public string Name { get => name; set => name = value; }
        public string SourceFrom { get => sourceFrom; set => sourceFrom = value; }
    }
}
