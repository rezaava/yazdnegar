using ShivaNegar.Constants;

namespace ShivaNegar.Models
{
    public class TemplateRelationshipModel
    {
        public enum SubTemplateIDs
        {
            //Template,

            PersianIdentity,//شناسنامه فارسی
            InTheNameOfAllah,//بسم الله
            DefenseMeeting,//صورت جلسه دفاع یا تاییدیه هیئت داوران
            LegalRights,//تعهد رعایت حقوق معنوی دانشگاه
            ResearchEthics,//منشور اخلاقی پژوهش
            Dedication,//تقدیم
            Acknowledgment,//سپاسگزاری
            PersianAbstract,//چکیده فارسی
            Preface,//پیشگفتار
            TableOfContents,//فهرست مطالب
            TableOfFigures,//فهرست شکل ها و نمودار ها
            TableOfTables,//فهرست جدول ها
            Abbreviations,//اختصارات
            Appendices,//ضمائم
            Endnotes,//یادداشت های پایانی
            Glossary,//واژه نامه
            Sources,//منابع
            EnglishAbstract,//چکیده انگلیسی
            EnglishIdentity,//شناسنامه انگلیسی
            Chapter1,
            Chapter2,
            Chapter3,
            Chapter4,
            Chapter5,
            Chapter6,
            Chapter7,
            Chapter8,
            Chapter9,
            Chapter10,

            TableOfMaps,//جدول نقشه ها
            EvaluationForm,//فرم نمره
            OriginalityCommitment,//تعهد اصالت

            FormB,//تعهد اصالت
        }
        internal enum SubTemplateTypes
        {
            Required,
            Optional,
            Chapter,
            ChapterNotMade,
            //NotIncluded,
        }

        private int _order;
        private string _pageTitle;
        private string _resourcePath;
        private string _fileName;

        private PageIDs _pageID;
        private SubTemplateIDs _subTemplateID;
        private SubTemplateTypes _subTemplateType;

        internal int Order { get => _order; set => _order = value; }
        internal string PageTitle { get => _pageTitle; set => _pageTitle = value; }
        internal string FileName { get => _fileName; set => _fileName = value; }
        internal string ResourcePath { get => _resourcePath; set => _resourcePath = value; }

        internal PageIDs PageID { get => _pageID; set => _pageID = value; }
        public SubTemplateIDs SubTemplateID { get => _subTemplateID; set => _subTemplateID = value; }
        internal SubTemplateTypes SubTemplateType { get => _subTemplateType; set => _subTemplateType = value; }

        internal TemplateRelationshipModel(int order, string id, string fileName, string resourcePath, bool isRequired)
        {
            Order = order;
            SubTemplateType = isRequired ? SubTemplateTypes.Required : SubTemplateTypes.Optional;
            FileName = fileName;
            ResourcePath = resourcePath;

            switch (id)
            {
                case "PersianIdentity":
                    PageTitle = "شناسنامه فارسی";
                    SubTemplateID = SubTemplateIDs.PersianIdentity;
                    PageID = PageIDs._page_PersianIdentity;
                    break;
                case "InTheNameOfAllah":
                    PageTitle = "به نام خدا";
                    SubTemplateID = SubTemplateIDs.InTheNameOfAllah;
                    PageID = PageIDs._page_InTheNameOfAllah;
                    break;
                case "DefenseMeeting":
                    PageTitle = "صورت جلسه دفاع";
                    SubTemplateID = SubTemplateIDs.DefenseMeeting;
                    PageID = PageIDs._page_DefenseMeeting;
                    break;
                case "LegalRights":
                    PageTitle = "تعهد رعایت حقوق معنوی دانشگاه";
                    SubTemplateID = SubTemplateIDs.LegalRights;
                    PageID = PageIDs._page_LegalRights;
                    break;
                case "ResearchEthics":
                    PageTitle = "منشور اخلاقی پژوهش";
                    SubTemplateID = SubTemplateIDs.ResearchEthics;
                    PageID = PageIDs._page_ResearchEthics;
                    break;
                case "Dedication":
                    PageTitle = "تقدیم";
                    SubTemplateID = SubTemplateIDs.Dedication;
                    PageID = PageIDs._page_Dedication;
                    break;
                case "Acknowledgment":
                    PageTitle = "قدردانی";
                    SubTemplateID = SubTemplateIDs.Acknowledgment;
                    PageID = PageIDs._page_Acknowledgment;
                    break;
                case "PersianAbstract":
                    PageTitle = "چکیده فارسی";
                    SubTemplateID = SubTemplateIDs.PersianAbstract;
                    PageID = PageIDs._page_PersianAbstract;
                    break;
                case "Preface":
                    PageTitle = "پیشگفتار";
                    SubTemplateID = SubTemplateIDs.Preface;
                    PageID = PageIDs._page_Preface;
                    break;
                case "TableOfContents":
                    PageTitle = "فهرست مطالب";
                    SubTemplateID = SubTemplateIDs.TableOfContents;
                    PageID = PageIDs._page_TableOfContents;
                    break;
                case "TableOfFigures":
                    PageTitle = "فهرست شکل‌ها/نمودارها";
                    SubTemplateID = SubTemplateIDs.TableOfFigures;
                    PageID = PageIDs._page_TableOfFigures;
                    break;
                case "TableOfTables":
                    PageTitle = "فهرست جدول‌ها";
                    SubTemplateID = SubTemplateIDs.TableOfTables;
                    PageID = PageIDs._page_TableOfTables;
                    break;
                case "Abbreviations":
                    PageTitle = "فهرست نماد‌ها/معادله‌ها";
                    SubTemplateID = SubTemplateIDs.Abbreviations;
                    PageID = PageIDs._page_Abbreviations;
                    break;

                case "Chapter1":
                    PageTitle = "فصل اوّل";
                    SubTemplateID = SubTemplateIDs.Chapter1;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter1;
                    break;
                case "Chapter2":
                    PageTitle = "فصل دوم";
                    SubTemplateID = SubTemplateIDs.Chapter2;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter2;
                    break;
                case "Chapter3":
                    PageTitle = "فصل سوم";
                    SubTemplateID = SubTemplateIDs.Chapter3;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter3;
                    break;
                case "Chapter4":
                    PageTitle = "فصل چهارم";
                    SubTemplateID = SubTemplateIDs.Chapter4;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter4;
                    break;
                case "Chapter5":
                    PageTitle = "فصل پنجم";
                    SubTemplateID = SubTemplateIDs.Chapter5;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter5;
                    break;
                case "Chapter6":
                    PageTitle = "فصل ششم";
                    SubTemplateID = SubTemplateIDs.Chapter6;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter6;
                    break;
                case "Chapter7":
                    PageTitle = "فصل هفتم";
                    SubTemplateID = SubTemplateIDs.Chapter7;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter7;
                    break;
                case "Chapter8":
                    PageTitle = "فصل هشتم";
                    SubTemplateID = SubTemplateIDs.Chapter8;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter8;
                    break;
                case "Chapter9":
                    PageTitle = "فصل نهم";
                    SubTemplateID = SubTemplateIDs.Chapter9;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter9;
                    break;
                case "Chapter10":
                    PageTitle = "فصل دهم";
                    SubTemplateID = SubTemplateIDs.Chapter10;
                    SubTemplateType = isRequired ? SubTemplateTypes.Chapter : SubTemplateTypes.ChapterNotMade;
                    PageID = PageIDs._page_Chapter10;
                    break;

                case "Appendices":
                    PageTitle = "ضمائم";
                    SubTemplateID = SubTemplateIDs.Appendices;
                    PageID = PageIDs._page_Appendices;
                    break;
                case "Endnotes":
                    PageTitle = "پی‌نوشت‌ها";
                    SubTemplateID = SubTemplateIDs.Endnotes;
                    PageID = PageIDs._page_Endnotes;
                    break;
                case "Glossary":
                    PageTitle = "واژه نامه";
                    SubTemplateID = SubTemplateIDs.Glossary;
                    PageID = PageIDs._page_Glossary;
                    break;
                case "Sources":
                    PageTitle = "منابع و مآخذ";
                    SubTemplateID = SubTemplateIDs.Sources;
                    PageID = PageIDs._page_Sources;
                    break;
                case "EnglishAbstract":
                    PageTitle = "چکیده انگلیسی";
                    SubTemplateID = SubTemplateIDs.EnglishAbstract;
                    PageID = PageIDs._page_EnglishAbstract;
                    break;
                case "EnglishIdentity":
                    PageTitle = "شناسنامه انگلیسی";
                    SubTemplateID = SubTemplateIDs.EnglishIdentity;
                    PageID = PageIDs._page_EnglishIdentity;
                    break;


                case "TableOfMaps":
                    PageTitle = "فهرست نقشه‌ها";
                    SubTemplateID = SubTemplateIDs.TableOfMaps;
                    PageID = PageIDs._page_TableOfMaps;
                    break;
                case "EvaluationForm":
                    PageTitle = "فرم نمره";
                    SubTemplateID = SubTemplateIDs.EvaluationForm;
                    PageID = PageIDs._page_EvaluationForm;
                    break;
                case "OriginalityCommitment":
                    PageTitle = "تعهد اصالت";
                    SubTemplateID = SubTemplateIDs.OriginalityCommitment;
                    PageID = PageIDs._page_OriginalityCommitment;
                    break;

                case "FormB":
                    PageTitle = "فرم ب";
                    SubTemplateID = SubTemplateIDs.FormB;
                    PageID = PageIDs._page_FormB;
                    break;

                default:
                    throw new System.Exception("page not exist!");
            }

        }
    }
}
