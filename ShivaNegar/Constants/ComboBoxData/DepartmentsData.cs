using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Xml;

namespace ShivaNegar.Constants.ComboBoxData
{
    class DepartmentsData
    {
        internal static List<string> getPersianDepartments(Universities university, bool addOther)
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
            foreach (XmlNode universityNode in universities)
            {
                if (universityNode.Attributes["id"].Value == ((int)university).ToString())
                {
                    XmlNode departmentsNode = universityNode.SelectSingleNode("departments");
                    XmlNodeList departmentNodeList = departmentsNode.SelectNodes("department");
                    foreach (XmlNode departmentNode in departmentNodeList)
                    {
                        list.Add(departmentNode.Attributes["persian"].Value);
                    }
                    break;
                }
            }
            if (addOther)
                list.Add(ComboBoxData.otherFa);

            return list;
        }
        internal static List<string> getEnglistDepartments(Universities university, bool addOther)
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
            foreach (XmlNode universityNode in universities)
            {
                if (universityNode.Attributes["id"].Value == ((int)university).ToString())
                {
                    XmlNode departmentsNode = universityNode.SelectSingleNode("departments");
                    XmlNodeList departmentNodeList = departmentsNode.SelectNodes("department");
                    foreach (XmlNode departmentNode in departmentNodeList)
                    {
                        list.Add(departmentNode.Attributes["english"].Value);
                    }
                    break;
                }
            }
            if (addOther)
                list.Add(ComboBoxData.otherEn);

            return list;
        }

        internal static List<string> getPersianGroups(Universities university, string persianDepartment)
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
            foreach (XmlNode universityNode in universities)
            {
                if (universityNode.Attributes["id"].Value == ((int)university).ToString())
                {
                    XmlNode departmentsNode = universityNode.SelectSingleNode("departments");
                    XmlNodeList departmentNodeList = departmentsNode.SelectNodes("department");
                    foreach (XmlNode departmentNode in departmentNodeList)
                    {
                        if (departmentNode.Attributes["persian"].Value == persianDepartment)
                        {
                            XmlNodeList groups = departmentNode.SelectNodes("group");
                            foreach (XmlNode group in groups)
                            {
                                list.Add(group.Attributes["persian"].Value);
                            }
                        }
                    }
                    break;
                }
            }
            if (list.Count == 0)
                list.Add(ComboBoxData.nothingFa);
            list.Add(ComboBoxData.otherFa);

            return list;
        }
        internal static List<string> getEnglishGroups(Universities university, string persianDepartment)
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
            foreach (XmlNode universityNode in universities)
            {
                if (universityNode.Attributes["id"].Value == ((int)university).ToString())
                {
                    XmlNode departmentsNode = universityNode.SelectSingleNode("departments");
                    XmlNodeList departmentNodeList = departmentsNode.SelectNodes("department");
                    foreach (XmlNode departmentNode in departmentNodeList)
                    {
                        if (departmentNode.Attributes["persian"].Value == persianDepartment)
                        {
                            XmlNodeList groups = departmentNode.SelectNodes("group");
                            foreach (XmlNode group in groups)
                            {
                                list.Add(group.Attributes["english"].Value);
                            }
                        }
                    }
                    break;
                }
            }
            if (list.Count == 0)
                list.Add(ComboBoxData.nothingEn);

            list.Add(ComboBoxData.otherEn);

            return list;
        }
    }
}
