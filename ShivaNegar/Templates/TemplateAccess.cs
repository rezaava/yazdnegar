using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml;
using Microsoft.Office.Interop.Word;
using ShivaNegar.Constants;
using ShivaNegar.Forms.AddRemovePages.Models;
using ShivaNegar.Models;
using static ShivaNegar.Models.TemplateRelationshipModel;

namespace ShivaNegar.Templates
{
    /*
	 <templates>
		<template id="" name="" fileName="" universityIcon="" path="">

			<subTemplates>
				<subTemplate order="" type="" id="" fileName="" />
				<subTemplate order="" type="" id="" fileName="" />
				<subTemplate order="" type="" id="" fileName="" />
			</subTemplates>
			
			<notIncludeSubTemplates>
				<notIncludeSubTemplate documentType="">
					<notInclude id=""/>
				</notIncludeSubTemplate>
			</notIncludeSubTemplates>

			<requiredSubTemplates>
				<defaultRequiredSubTemplates>
					<required id="" />
					<required id="" />
					<required id="" />
				</defaultRequiredSubTemplates>

				<customRequiredSubTemplate documentType="">
					<required id="" />
					<required id="" />
					<required id="" />
				</customRequiredSubTemplate>
			</requiredSubTemplates>

		</template>
	  </templates>
	 */


    internal class TemplateAccess
    {

        #region get template data
        internal static Bitmap getUniversityIcon(Universities university, TemplateTypes templateType)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    string universityTemplatePath = template.Attributes["path"].Value;
                    string iconName = template.Attributes["universityIcon"].Value;
                    string resourcePath = universityTemplatePath + iconName;
                    Stream streamIcon = DedicatedFunctions.getStream(resourcePath);
                    Bitmap bitmap = new Bitmap(streamIcon);
                    return bitmap;
                }
            }
            return null;
        }
        internal static Stream getTemplateFileStream(Universities university, TemplateTypes templateType)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    string universityTemplatePath = template.Attributes["path"].Value;
                    string templateFileName = template.Attributes["fileName"].Value;
                    string resourcePath = universityTemplatePath + templateFileName;

                    if (Directory.Exists(Properties.Settings.Default.WorkSpaceDirectory + StringConstant.DocumentsTemplateFolder))
                    {
                        if (File.Exists(Properties.Settings.Default.WorkSpaceDirectory + StringConstant.DocumentsTemplateFolder + templateFileName))
                            return new FileStream(Properties.Settings.Default.WorkSpaceDirectory + StringConstant.DocumentsTemplateFolder + templateFileName, FileMode.Open);
                    }
                    else
                    {
                        Stream streamTemplate = DedicatedFunctions.getStream(resourcePath);
                        return streamTemplate;
                    }
                }
            }
            return null;
        }
        internal static string getTemplateFileName(Universities university, TemplateTypes templateType)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    return template.Attributes["fileName"].Value;
                }
            }
            return null;
        }
        internal static string getTemplateSourcePath(Universities university, TemplateTypes templateType)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    return template.Attributes["source"].Value;
                }
            }
            return null;
        }
        #endregion

        #region get subTemplate data
        internal static string getSubTemplateFileName(Universities university, TemplateTypes templateType, SubTemplateIDs sti)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    XmlNode subTemplatesNode = template.SelectSingleNode("subTemplates");

                    XmlNodeList subTemplates = subTemplatesNode.SelectNodes("subTemplate");
                    foreach (XmlNode subTemplate in subTemplates)
                    {
                        if (subTemplate.Attributes["id"].Value == sti.ToString())
                        {
                            return subTemplate.Attributes["fileName"].Value;
                        }
                    }
                }
            }
            return null;
        }

        internal static string getCustomTitle(Universities university, TemplateTypes templateType, ContentControlNames contentControlID)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    XmlNode TitlesNode = template.SelectSingleNode("titles");

                    XmlNodeList titleNodeList = TitlesNode.SelectNodes("title");
                    foreach (XmlNode titleNode in titleNodeList)
                    {
                        if (titleNode.Attributes["id"].Value == contentControlID.ToString())
                        {
                            return titleNode.Attributes["text"].Value;
                        }
                    }
                }
            }
            return "";
        }
        #endregion

        #region get subTemplates and requiredSubTemplates List
        internal static List<ProposalRelationshipModel> getProposalModelList(Universities university, TemplateTypes templateType)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            List<ProposalRelationshipModel> proposalList = null;

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    XmlNode subProposalNode = template.SelectSingleNode("proposals");

                    if (subProposalNode == null)
                        return null;

                    XmlNodeList subTemplateNodeList = subProposalNode.SelectNodes("proposal");
                    proposalList = new List<ProposalRelationshipModel>();

                    foreach (XmlNode subTemplateNode in subTemplateNodeList)
                    {
                        int academicID = int.Parse(subTemplateNode.Attributes["academicDegree"].Value);
                        AcademicDegrees ad = (AcademicDegrees)academicID;

                        string resourcePath = template.Attributes["path"].Value + subProposalNode.Attributes["subPath"].Value;
                        string fileName = subTemplateNode.Attributes["fileName"].Value;

                        ProposalRelationshipModel model = new ProposalRelationshipModel(ad, fileName, resourcePath);
                        proposalList.Add(model);
                    }
                    break;
                }
            }
            return proposalList;
        }

        internal static List<TemplateRelationshipModel> getTemplateRelationshipModelList(Universities university, TemplateTypes templateType, DocumentTypes documentType, bool getAll)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            List<TemplateRelationshipModel> subTemplateList = new List<TemplateRelationshipModel>();
            List<string> requiredList = new List<string>();
            List<string> notIncludeList = new List<string>();

            //get data from XML file
            XmlDocument doc = new XmlDocument();
            doc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = doc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    XmlNode subTemplatesNode = template.SelectSingleNode("subTemplates");
                    XmlNode requiredSubTemplatesNode = template.SelectSingleNode("requiredSubTemplates");
                    XmlNode notIncludeSubTemplatesNode = template.SelectSingleNode("notIncludeSubTemplates");

                    if (notIncludeSubTemplatesNode != null)
                    {
                        XmlNodeList notIncludeSubTemplateNodeList = notIncludeSubTemplatesNode.SelectNodes("notIncludeSubTemplate");
                        if (notIncludeSubTemplateNodeList != null)
                        {
                            foreach (XmlNode notIncludeSubTemplateNode in notIncludeSubTemplateNodeList)
                            {
                                string documentTypeAttribute = notIncludeSubTemplateNode.Attributes["documentType"].Value;

                                if (int.Parse(documentTypeAttribute) == ((int)documentType))
                                {
                                    XmlNodeList notIncludeNodeList = notIncludeSubTemplateNode.SelectNodes("notInclude");
                                    foreach (XmlNode notIncludeNode in notIncludeNodeList)
                                    {
                                        notIncludeList.Add(notIncludeNode.Attributes["id"].Value);
                                    }
                                    break;
                                }

                            }
                        }
                    }

                    XmlNodeList customRequiredSubTemplateNodeList = requiredSubTemplatesNode.SelectNodes("customRequiredSubTemplate");
                    if (customRequiredSubTemplateNodeList != null)
                    {
                        foreach (XmlNode customRequiredSubTemplateNode in customRequiredSubTemplateNodeList)
                        {
                            string documentTypeAttribute = customRequiredSubTemplateNode.Attributes["documentType"].Value;

                            if (int.Parse(documentTypeAttribute) == ((int)documentType))
                            {
                                XmlNodeList requiredNodeList = customRequiredSubTemplateNode.SelectNodes("required");
                                foreach (XmlNode requiredNode in requiredNodeList)
                                {
                                    requiredList.Add(requiredNode.Attributes["id"].Value);
                                }
                                break;
                            }

                        }
                    }

                    if (requiredList.Count == 0)
                    {
                        XmlNode defaultRequiredSubTemplatesNode = requiredSubTemplatesNode.SelectSingleNode("defaultRequiredSubTemplates");
                        XmlNodeList requiredNodeList = defaultRequiredSubTemplatesNode.SelectNodes("required");
                        foreach (XmlNode requiredNode in requiredNodeList)
                        {
                            requiredList.Add(requiredNode.Attributes["id"].Value);
                        }
                    }

                    XmlNodeList subTemplateNodeList = subTemplatesNode.SelectNodes("subTemplate");
                    foreach (XmlNode subTemplateNode in subTemplateNodeList)
                    {
                        string id = subTemplateNode.Attributes["id"].Value;
                        bool isDisable = subTemplateNode.Attributes["isDisable"] == null ? false : true;
                        if (getAll || (!notIncludeList.Contains(id) && !isDisable))
                        {
                            string resourcePath = template.Attributes["path"].Value;
                            int order = int.Parse(subTemplateNode.Attributes["order"].Value);
                            string fileName = subTemplateNode.Attributes["fileName"].Value;
                            bool isRequired = requiredList.Contains(id);

                            TemplateRelationshipModel model = new TemplateRelationshipModel(order, id, fileName, resourcePath, isRequired);
                            subTemplateList.Add(model);
                        }
                    }
                    break;
                }
            }
            subTemplateList = subTemplateList.OrderBy(o => o.Order).ToList();
            return subTemplateList;
        }
        internal static List<AddRemovePageRelationModel> getAddRemovePageRelationModelList(Document doc, Universities university, TemplateTypes templateType, DocumentTypes documentType)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            Stream templateXMLStream = assembly.GetManifestResourceStream(EmbeddedResourceNames.Templates);

            List<AddRemovePageRelationModel> subTemplateList = new List<AddRemovePageRelationModel>();
            List<string> requiredList = new List<string>();
            List<string> notIncludeList = new List<string>();

            //get data from XML file
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(templateXMLStream);

            //Access to nodes and features
            XmlNodeList templates = xmlDoc.SelectNodes("//template");
            foreach (XmlNode template in templates)
            {
                if (template.Attributes["id"].Value == DedicatedFunctions.getTemplateID(templateType, university).ToString())
                {
                    XmlNode subTemplatesNode = template.SelectSingleNode("subTemplates");
                    XmlNode requiredSubTemplatesNode = template.SelectSingleNode("requiredSubTemplates");
                    XmlNode notIncludeSubTemplatesNode = template.SelectSingleNode("notIncludeSubTemplates");

                    if (notIncludeSubTemplatesNode != null)
                    {
                        XmlNodeList notIncludeSubTemplateNodeList = notIncludeSubTemplatesNode.SelectNodes("notIncludeSubTemplate");
                        if (notIncludeSubTemplateNodeList != null)
                        {
                            foreach (XmlNode notIncludeSubTemplateNode in notIncludeSubTemplateNodeList)
                            {
                                string documentTypeAttribute = notIncludeSubTemplateNode.Attributes["documentType"].Value;

                                if (int.Parse(documentTypeAttribute) == ((int)documentType))
                                {
                                    XmlNodeList notIncludeNodeList = notIncludeSubTemplateNode.SelectNodes("notInclude");
                                    foreach (XmlNode notIncludeNode in notIncludeNodeList)
                                    {
                                        notIncludeList.Add(notIncludeNode.Attributes["id"].Value);
                                    }
                                    break;
                                }

                            }
                        }
                    }

                    if (requiredSubTemplatesNode != null)
                    {
                        XmlNodeList customRequiredSubTemplateNodeList = requiredSubTemplatesNode.SelectNodes("customRequiredSubTemplate");
                        if (customRequiredSubTemplateNodeList != null)
                        {
                            foreach (XmlNode customRequiredSubTemplateNode in customRequiredSubTemplateNodeList)
                            {
                                string documentTypeAttribute = customRequiredSubTemplateNode.Attributes["documentType"].Value;

                                if (int.Parse(documentTypeAttribute) == ((int)documentType))
                                {
                                    XmlNodeList requiredNodeList = customRequiredSubTemplateNode.SelectNodes("required");
                                    foreach (XmlNode requiredNode in requiredNodeList)
                                    {
                                        requiredList.Add(requiredNode.Attributes["id"].Value);
                                    }
                                    break;
                                }

                            }
                        }
                    }

                    if (requiredList.Count == 0)
                    {
                        XmlNode defaultRequiredSubTemplatesNode = requiredSubTemplatesNode.SelectSingleNode("defaultRequiredSubTemplates");
                        XmlNodeList requiredNodeList = defaultRequiredSubTemplatesNode.SelectNodes("required");
                        foreach (XmlNode requiredNode in requiredNodeList)
                        {
                            requiredList.Add(requiredNode.Attributes["id"].Value);
                        }
                    }


                    XmlNodeList subTemplateNodeList = subTemplatesNode.SelectNodes("subTemplate");
                    foreach (XmlNode subTemplateNode in subTemplateNodeList)
                    {
                        string id = subTemplateNode.Attributes["id"].Value;
                        bool isDisable = subTemplateNode.Attributes["isDisable"] == null ? false : true;
                        if (!notIncludeList.Contains(id) && !isDisable)
                        {
                            string resourcePath = template.Attributes["path"].Value;
                            int order = int.Parse(subTemplateNode.Attributes["order"].Value);
                            string fileName = subTemplateNode.Attributes["fileName"].Value;
                            bool isRequired = requiredList.Contains(id);

                            subTemplateList.Add(new AddRemovePageRelationModel(doc, order, id, fileName, resourcePath, isRequired));
                        }
                    }
                    break;
                }
            }
            subTemplateList = subTemplateList.OrderBy(o => o.Order).ToList();
            return subTemplateList;
        }
        #endregion

    }
}
