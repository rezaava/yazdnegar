using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml;

namespace ShivaNegar.Constants.ComboBoxData
{
    class UniversitiesData
    {
        internal static List<string> getPersianUniversities()
        {
            List<string> list = new List<string>();

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(EmbeddedResourceNames.Universities))
            {
                doc.Load(stream);
            }

            //Access to nodes and features
            XmlNodeList universities = doc.SelectNodes("//university");
            foreach (XmlNode university in universities)
            {
                list.Add(university.Attributes["persian"].Value);
            }
            return list;
        }
        internal static List<string> getEnglishUniversities()
        {
            List<string> list = new List<string>();

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(EmbeddedResourceNames.Universities))
            {
                doc.Load(stream);
            }

            //Access to nodes and features
            XmlNodeList universities = doc.SelectNodes("//university");
            foreach (XmlNode university in universities)
            {
                list.Add(university.Attributes["english"].Value);
            }
            return list;
        }

        internal static string getBranchFaOfUniversity(Universities university)
        {
            //get data from XML file
            XmlDocument doc = new XmlDocument();
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(EmbeddedResourceNames.Universities))
            {
                doc.Load(stream);
            }

            //Access to nodes and features
            XmlNodeList universities = doc.SelectNodes("//university");
            foreach (XmlNode universityNode in universities)
            {
                if (universityNode.Attributes["id"].Value == ((int)university).ToString())
                {
                    XmlAttribute branchFaAttribute = universityNode.Attributes["branchFa"];
                    return branchFaAttribute == null ? "" : "واحد " + branchFaAttribute.Value;
                }
            }
            return "";
        }

        internal static string getBranchEnOfUniversity(Universities university)
        {
            //get data from XML file
            XmlDocument doc = new XmlDocument();
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(EmbeddedResourceNames.Universities))
            {
                doc.Load(stream);
            }

            //Access to nodes and features
            XmlNodeList universities = doc.SelectNodes("//university");
            foreach (XmlNode universityNode in universities)
            {
                if (universityNode.Attributes["id"].Value == ((int)university).ToString())
                {
                    XmlAttribute branchEnAttribute = universityNode.Attributes["branchEn"];
                    return branchEnAttribute == null ? null : branchEnAttribute.Value + " Branch";
                }
            }
            return null;
        }
    }
}
