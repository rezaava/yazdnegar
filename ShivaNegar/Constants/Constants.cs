using System.Drawing;
using Microsoft.Office.Interop.Word;

namespace ShivaNegar.Constants
{
    public enum DocumentTypes : int
    {
        Nothing = -1,
        Project = 0,
        Thesis = 1,
        Dissertation = 2,
        SchoolResearch = 3,
    }

    public enum RoleTypes : int
    {
        CurrentUser_Chat = -1,
        Creator = 0,

        Supervisor1 = 1,
        Supervisor2 = 2,
        Advisor1 = 3,
        Advisor2 = 4,
        Analyst = 5,
    }

    public struct RoleNames
    {
        internal const string Supervisor1 = "استاد راهنما اول";
        internal const string Supervisor2 = "استاد راهنما دوم";
        internal const string Advisor1 = "استاد مشاور اول";
        internal const string Advisor2 = "استاد مشاور دوم";
        internal const string Analyst = "همکار";

        internal const string Creator = "سازنده";
    }

    public enum TemplateTypes : int
    {
        ShivaNegarTemplate = -3,
        ParsaTemplate = -2,
        Nothing = -1,
        UniversityTemplate = 1,
        StudentTemplate = 2,
        OfficeTemplate = 3,

    }
    public enum Universities : int
    {
        Nothing = -1,
        YazdUniversity = 0,
        AzadUniversityYazd = 1,
        AzadUniversityCentral = 2,
        ScienceAndTechnologyUni = 3,
    }
    public enum ReportTypes : int
    {
        UserReport = 1,
        ApplicationReport = 2,
    }

    public enum PoemTypes
    {
        CreatePoemUsingTextColumns,
        CreatePoemUsingTable,
    }

    public struct EmbeddedResourceNames
    {
        internal const string ShivaNegarShortcut = "ShivaNegar.Constants.ShivaNegarShortcut.dotm";

        internal const string HalfSpace = "ShivaNegar.Constants.VirastarFiles.HalfSpace.csv";
        internal const string Standard = "ShivaNegar.Constants.VirastarFiles.Standard.csv";
        internal const string Tanvin = "ShivaNegar.Constants.VirastarFiles.Tanvin.csv";
        internal const string Signs = "ShivaNegar.Constants.VirastarFiles.Signs.csv";
        internal const string Tashdid = "ShivaNegar.Constants.VirastarFiles.Tashdid.csv";
        internal const string SpellingCorrection = "ShivaNegar.Constants.VirastarFiles.SpellingCorrection.csv";

        internal const string Khotbeh = "ShivaNegar.Constants.NahjBalagheFiles.Khotbeh.csv";
        internal const string Nameh = "ShivaNegar.Constants.NahjBalagheFiles.Nameh.csv";
        internal const string Hekmat = "ShivaNegar.Constants.NahjBalagheFiles.Hekmat.csv";

        internal const string Ayeh = "ShivaNegar.Constants.QuranFiles.Ayeh.csv";
        internal const string Surah = "ShivaNegar.Constants.QuranFiles.Surah.csv";

        internal const string Dedicate = "ShivaNegar.Constants.OtherFiles.Dedicate.csv";

        internal const string DefenseAnnouncements = "ShivaNegar.Templates.DefenseAnnouncements.xml";
        internal const string Templates = "ShivaNegar.Templates.Templates.xml";

        internal const string Universities = "ShivaNegar.Constants.Universities.xml";

        internal const string B_Yagut = "ShivaNegar.Resources.B Yagut.ttf";
        internal const string B_Yagut_Bold = "ShivaNegar.Resources.B Yagut Bold.ttf";

        internal const string B_Badr_Bold = "ShivaNegar.Resources.B Badr Bold.ttf";
        internal const string B_Badr = "ShivaNegar.Resources.B Badr.ttf";
        internal const string B_Lotus_Bold = "ShivaNegar.Resources.B Lotus Bold.ttf";
        internal const string B_Lotus = "ShivaNegar.Resources.B Lotus.ttf";
        internal const string B_Nazanin_Bold = "ShivaNegar.Resources.B Nazanin Bold.ttf";
        internal const string B_Nazanin = "ShivaNegar.Resources.B Nazanin.ttf";
        internal const string B_Titr_Bold = "ShivaNegar.Resources.B Titr Bold.ttf";
        internal const string B_Zar_Bold = "ShivaNegar.Resources.B Zar Bold.ttf";
        internal const string B_Zar = "ShivaNegar.Resources.B Zar.ttf";
        internal const string Besmellah_1 = "ShivaNegar.Resources.Besmellah_1.ttf";
        internal const string Besmellah_2 = "ShivaNegar.Resources.Besmellah_2.ttf";
        internal const string Besmellah_3 = "ShivaNegar.Resources.Besmellah_3.ttf";
        internal const string Besmellah_4 = "ShivaNegar.Resources.Besmellah_4.ttf";
        internal const string IranNastaliq = "ShivaNegar.Resources.IranNastaliq.ttf";
        internal const string times = "ShivaNegar.Resources.TIMES.ttf";
        internal const string times_BD = "ShivaNegar.Resources.TIMESBD.ttf";
        internal const string times_BI = "ShivaNegar.Resources.TIMESBI.ttf";
        internal const string times_I = "ShivaNegar.Resources.TIMESI.ttf";
        internal const string Vazir = "ShivaNegar.Resources.Vazir.ttf";
        internal const string Vazirmatn_Bold = "ShivaNegar.Resources.Vazirmatn-Bold.ttf";
        internal const string Vazirmatn_Regular = "ShivaNegar.Resources.Vazirmatn-Regular.ttf";
    }

    struct ParagraphAndTextWrapMarks
    {
        internal const string ParagraphMarkInText = "\r";
        internal const char ParagraphMarkInTextChar = '\r';
        internal const string ParagraphMarkInFind = "^p";

        internal const string TextWrapMarkInText = "\v";//manual line break
        internal const string TextWrapMarkInFind = "^l";//manual line break
    }

    struct StringConstant
    {
        internal static string[] AdministratorAccounts = { "09355374226", "09132233916" , "09912071445" };

        internal const string PrimaryServerApiBaseAddress = PrimaryServerBaseAddress + "api/";
        internal const string PrimaryServerBaseAddress = "http://shivanegar.ir/";

        internal const string NameOfProject = "پارسانگار";
        internal const string SupportEmail = "nazariansani@gmail.com";
        internal const string SupportMobile = "09039402316";

        internal const string GUID = "e3ec9f6c-c8f5-4172-ae4b-125fac41e118";
        internal const string AppTempFolder = "c6f9eafb-8787-4e75-9a9a-088b74035e2c\\";
        internal const string VBAPassword = "YTDgzvk93GJZKYeU";
        internal const string DocumentPassword = "Qje3qMyTrK3pr5A";
        //internal const string DocumentPassword2 = "J9t1yiZVo494xuk";
        internal const string DocumentProtectionPassword = "DE1xkD62VPp8eaVW";

        internal const string DefaultWorkspaceName = "ShivaNegarWorkspace";

        internal const string ArchiveFolder = "بایگانی\\";
        internal const string VirastarFolder = "ویراستار\\";
        internal const string DocumentsTemplateFolder = "اسناد الگو\\";
        
        internal const string TashdidFile = "Tashdid.csv";
        internal const string TanvinFile = "Tanvin.csv";
        internal const string SignFile = "Signs.csv";
        internal const string HalfSpaceFile = "HalfSpace.csv";
        internal const string StandardCorrectionFile = "Standard.csv";
        internal const string spellingCorrectionFile = "SpellingCorrection.csv";


        // CSV Code

        public const string ButtonCodeSigns = "1000";
        public const string ButtonCodeSpelling = "1001";
        public const string ButtonCodeHalfSpace = "1002";
        public const string ButtonCodeStandard = "1004";
    }

    struct DialogBoxMessages
    {
        internal const string RequiredDedicatedDocument = "در حال حاضر این سند به عنوان سند جاری انتخاب نشده است.";

    }

    //Microsoft.Office.Core.MsoTriState.msoTrue
    struct ColorsApp
    {
        internal static Color PrimaryColor = Color.FromArgb(255, 6, 174, 244);
    }
    struct SpecialTablesMessage
    {
        internal const string replace1 = "replace1";
        internal const string replace2 = "replace2";

        internal const string OldBibliographyNoItemMessage = "There are no sources in the current document.";
        internal const string NewBibliographyNoItemMessage = "هیچ منبعی در متن ارجاع داده نشده است. تنها منابعی در لیست ذکر می‌شوند که در متن به آنها ارجاع شده باشد";

        internal const string OldTableOfFiguresNoItemMessage = "No table of figures entries found.";
        internal const string NewTableOfFiguresNoItemMessage = "هیچ آیتمی از " + replace1 + " وجود ندارد، برای نمایش فهرست، عنوانی را اضافه نمایید";

    }

    internal enum TextDirection
    {
        RTL,
        LTR
    }

    internal enum KeyboardLanguage : int
    {
        Persian = WdLanguageID.wdPersian,
        English = WdLanguageID.wdEnglishUS,
    }


    struct CaptionLabels
    {
        internal const string captionTable = "جدول";
        internal const string captionFigure = "شکل";
        internal const string captionFormula = "معادله";//فرمول
        internal const string captionMap = "نقشه";
        internal const string Nothing = "هیچکدام";//for TableOfContents

        //internal const string captionChart = "نمودار";
        //internal const string captionAttach= "پیوست";

    }

    struct TaskPaneTitles
    {
        internal const string CreateCD = "تنظیمات لوح فشرده";
        internal const string CrossReference = "ارجاع دادن";
        internal const string InsertCitation = "درج منبع";
        internal const string InsertQuran = "درج آیه قرآن";
        internal const string InsertNahjBalaghe = "درج نهج البلاغه";
        internal const string ChatBoxNetworking = "محیط تبادل نظر";

        internal const string InsertDedicate = "تغییر/درج تقدیم";
        internal const string DefenseAnnouncements = "اطلاعیه دفاع";
    }

    internal enum CustomProperties
    {
        Custom_PageSetup,

        Section_Break_OddPage,
        Section_Break_EvenPage,
    }

    struct SettingValues
    {
        internal const string NullInHeaderFooterTemplates = "[null]";

        internal const string True = "True";
        internal const string False = "False";
        internal const string NotExist = "[-1]";
    }

    #region Encryption
    struct EncryptionKeys
    {

        internal static int keySize = 256;
        internal static int blockSize = 128;
        internal static int feedbackSize = 128;

        internal static System.Security.Cryptography.CipherMode cipherMode = System.Security.Cryptography.CipherMode.CBC;
        internal static System.Security.Cryptography.PaddingMode paddingMode = System.Security.Cryptography.PaddingMode.PKCS7;

        internal static byte[] IV = { 7, 152, 116, 244, 199, 145, 11, 247, 97, 27, 210, 192, 3, 37, 26, 86 };

        internal static byte[] Key = { 207, 86, 231, 21, 14, 225, 143, 36, 129, 5, 253, 208, 237, 77, 222, 225, 235, 87, 188, 2, 78, 144, 119, 220, 96, 188, 112, 110, 70, 215, 20, 76 };
    }
    #endregion

    #region Type And AcademicDegree Values
    struct DocumentTypeValues
    {
        internal const string DocumentType_NothingFa = "صفحه خالی";
        internal const string DocumentType_NothingEn = "Nothing";

        internal const string DocumentType_ProjectFa = "پروژه";
        //internal const string DocumentType_SchoolResearchFa = "تحقیق درسی، موضوع ویژه و گزارش کارورزی";
        internal const string DocumentType_SchoolResearchFa = "تحقیق درسی";
        internal const string DocumentType_ThesisFa = "پایان نامه";
        internal const string DocumentType_DissertationFa = "رساله";
        //internal const string DocumentType_SkillTrainingFa = "مهارت آموزی";
        //internal const string DocumentType_InternshipFa = "کارورزی";
        //internal const string DocumentType_SeminarFa = "موضوع ویژه(سمینار)";

        internal const string DocumentType_ProjectEn = "Project";
        internal const string DocumentType_SchoolResearchEn = "School Research";
        internal const string DocumentType_ThesisEn = "Thesis";
        internal const string DocumentType_DissertationEn = "Dissertation";
        //internal const string DocumentType_SkillTrainingEn = "Skill Training";
        //internal const string DocumentType_InternshipEn = "Internship";
        //internal const string DocumentType_SeminarEn = "Seminar";
    }
    struct TemplateTypeValues
    {
        internal const string TemplateType_Parsa = "قالب پارسا (وزارت علوم و تحقیقات و فناوری)";
        internal const string TemplateType_University = "قالب دانشگاهی";
        internal const string TemplateType_ShivaNegar = "قالب پارسانگار";
        internal const string TemplateType_Student = "قالب دانش آموزی (به زودی)";
        internal const string TemplateType_Office = "قالب گزارش اداری (به زودی)";
        internal const string TemplateType_Nothing = "صفحه خالی";
        internal const string TemplateType_ParsaEn = "Parsa Template";
        internal const string TemplateType_UniversityEn = "University Template";
        internal const string TemplateType_ShivaNegarEn = "Shivanegar Template";
        internal const string TemplateType_NothingEn = "Nothing";
    }

    enum AcademicDegrees : int
    {
        Nothing = 0,

        Doctoral = 1,
        MasterOfScience = 2,
        BachelorOfScience = 3,
        AssociateOfScience = 4,
    }

    struct AcademicDegreeValues
    {
        internal const string AcademicDegree_AssociateOfScienceFa = "کاردانی";
        internal const string AcademicDegree_BachelorOfScienceFa = "کارشناسی";
        internal const string AcademicDegree_PartTimeBachelorOfScienceFa = "کارشناسی ناپیوسته";
        internal const string AcademicDegree_MasterOfScienceFa = "کارشناسی ارشد";
        internal const string AcademicDegree_DoctoralFa = "دکتری";

        //S : Science | A: Arts
        //Ph.D. : Doctor of Philosophy | Ed.D. : Doctor of Education | D.Sc. : Doctor of Science
        internal const string AcademicDegree_AssociateOfScienceEn = "A.Sc"; // "A.Sc." "A.A" "A.D"
        internal const string AcademicDegree_BachelorOfScienceEn = "B.Sc"; // "B.Sc." "B.A"
        internal const string AcademicDegree_PartTimeBachelorOfScienceEn = "B.A";
        internal const string AcademicDegree_MasterOfScienceEn = "M.Sc";//"M.Sc." "M.A"
        internal const string AcademicDegree_DoctoralEn = "Ph.D";//"Ed.D" "D.Sc." 
    }
    #endregion


    #region Error Codes And Messages
    struct ErrorMessages
    {
        internal const string ErrorServiceUnavailable = "در برقراری ارتباط با سرور مشکلی وجود دارد" + "\nلطفا اتصال اینترنت خود را بررسی کنید.";
    }

    internal enum ErrorCodes : int
    {
        Nothing = -1,

        UnexpectedServerResponse = 600,

        InternalStartup = 19,
        StartupProblem = 20,

        CursorInLockedSpace = 10,


        ActionsPaneInitials = 50,
        ActionsPaneCDEnabling = 51,
        ActionsPaneCitationEnabling = 52,
        ActionsPaneSettingsEnabling = 53,
        ActionsPaneCrossReferenceEnabling = 54,

        ActionsPaneCDDisabling = 70,
        ActionsPaneCitationDisabling = 71,
        ActionsPaneSettingsDisabling = 72,
        ActionsPaneCrossReferenceDisabling = 73,


        UnprotectVBA = 130,
        UnprotectVBA2 = 131,
        UnprotectVBA3 = 132,
        SetKeyboardShortcut = 133,
        SetKeyboardShortcut2 = 134,
        CopyShortcutTemplate = 135,


        addVariable = 160,
        setStaticVariableValue = 162,
        isStaticVariableExist = 163,


        BookmarkNotExist = 200,


        CD_ShowActionPane = 250,
        CD_Exporting = 251,


        ContentControlProtect = 512,

    }
    #endregion

    #region Shape,Table ID
    internal enum ShapeIDs
    {
        _Pilot_CreateCD_,
        _CD_UniversityIcon_,
        _CD_CirculeInline_,
        _CD_CirculeOutline_,
        _CD_Title_,
    }

    internal struct TableIDs
    {
        internal const string tableOneColumnPoem = "{tableOneColumnPoem}";
        internal const string tableTwoColumnPoem = "{tableTwoColumnPoem}";
    }
    #endregion

    #region Type Enums

    internal enum TableTypes
    {
        TableOfContents,
        TableOfFigures,
    }
    internal enum TextTypes
    {
        RTL,
        LTR,
        BOTH
    }

    #endregion

    #region Font And Style
    internal struct FontNames
    {
        internal const string fontBNazanin = "B Nazanin";
        internal const string fontBZar = "B Zar";
        internal const string fontBLotus = "B Lotus";
        internal const string fontBYagut = "B Yagut";
        internal const string fontBBadr = "B Badr";
        internal const string fontIranNastaliq = "IranNastaliq";
        internal const string fontBTitr = "B Titr";
        internal const string fontTimesNewRoman = "Times New Roman";

        //internal const string fontTahoma = "Tahoma";
        //internal const string fontArial = "Arial";

        //internal const string fontBMitra = "B Mitra";
        //internal const string fontBLotus = "B Lotus";
        //internal const string fontBTraffic = "B Traffic";
        //internal const string fontBKoodak = "B Koodak";
        //internal const string fontBHoma = "B Homa";
        //internal const string fontIranNastaliq = "IranNastaliq";

        internal const string fontBesmellah1 = "Besmellah 1";
        internal const string fontBesmellah2 = "Besmellah 2";
        internal const string fontBesmellah3 = "Besmellah 3";
        internal const string fontBesmellah4 = "Besmellah 4";
        internal const string fontVazirmatn = "Vazirmatn";
    }
    internal struct StyleNames
    {
        internal const string CustomSectionBreakStyle = "Protected Section Break";


        internal const WdBuiltinStyle BodyText = WdBuiltinStyle.wdStyleBodyText;
        internal const WdBuiltinStyle styleNormal = WdBuiltinStyle.wdStyleNormal;

        internal const WdBuiltinStyle styleHeading1 = WdBuiltinStyle.wdStyleHeading1;
        internal const WdBuiltinStyle styleHeading2 = WdBuiltinStyle.wdStyleHeading2;
        internal const WdBuiltinStyle styleHeading3 = WdBuiltinStyle.wdStyleHeading3;
        internal const WdBuiltinStyle styleHeading4 = WdBuiltinStyle.wdStyleHeading4;
        internal const WdBuiltinStyle styleHeading5 = WdBuiltinStyle.wdStyleHeading5;
        internal const WdBuiltinStyle styleHeading6 = WdBuiltinStyle.wdStyleHeading6;
        internal const WdBuiltinStyle styleHeading7 = WdBuiltinStyle.wdStyleHeading7;
        internal const WdBuiltinStyle styleHeading8 = WdBuiltinStyle.wdStyleHeading8;
        internal const WdBuiltinStyle styleHeading9 = WdBuiltinStyle.wdStyleHeading9;

        internal const WdBuiltinStyle styleTitle = WdBuiltinStyle.wdStyleTitle;
        internal const WdBuiltinStyle styleSubtitle = WdBuiltinStyle.wdStyleSubtitle;
        internal const WdBuiltinStyle styleEmphasis = WdBuiltinStyle.wdStyleEmphasis;
        internal const WdBuiltinStyle styleStrong = WdBuiltinStyle.wdStyleStrong;
        internal const WdBuiltinStyle styleIntenseQuote = WdBuiltinStyle.wdStyleIntenseQuote;
        internal const WdBuiltinStyle styleQuote = WdBuiltinStyle.wdStyleQuote;
        internal const WdBuiltinStyle styleListParagraph = WdBuiltinStyle.wdStyleListParagraph;
        internal const string styleNoSpacing = "No Spacing";
        internal const string styleBlockText = "Block Text";

        internal const WdBuiltinStyle styleHeader = WdBuiltinStyle.wdStyleHeader;
        internal const WdBuiltinStyle styleFooter = WdBuiltinStyle.wdStyleFooter;
        internal const WdBuiltinStyle styleFootnoteText = WdBuiltinStyle.wdStyleFootnoteText;
        internal const WdBuiltinStyle styleEndnoteText = WdBuiltinStyle.wdStyleEndnoteText;
        internal const WdBuiltinStyle styleCaption = WdBuiltinStyle.wdStyleCaption;
        //internal const WdBuiltinStyle styleBibliography = WdBuiltinStyle.wdStyleBibliography;
        //internal const WdBuiltinStyle styleTableOfFigures = WdBuiltinStyle.wdStyleTableOfFigures;

        internal const WdBuiltinStyle styleHyperlink = WdBuiltinStyle.wdStyleHyperlink;
        internal const WdBuiltinStyle styleLineNumber = WdBuiltinStyle.wdStyleLineNumber;
        internal const WdBuiltinStyle stylePageNumber = WdBuiltinStyle.wdStylePageNumber;



        //internal const WdBuiltinStyle styleTOCHeading = WdBuiltinStyle.wdStyleTocHeading;
        internal const WdBuiltinStyle styleTOC1 = WdBuiltinStyle.wdStyleTOC1;
        internal const WdBuiltinStyle styleTOC2 = WdBuiltinStyle.wdStyleTOC2;
        internal const WdBuiltinStyle styleTOC3 = WdBuiltinStyle.wdStyleTOC3;
        internal const WdBuiltinStyle styleTOC4 = WdBuiltinStyle.wdStyleTOC4;
        internal const WdBuiltinStyle styleTOC5 = WdBuiltinStyle.wdStyleTOC5;
        internal const WdBuiltinStyle styleTOC6 = WdBuiltinStyle.wdStyleTOC6;
        internal const WdBuiltinStyle styleTOC7 = WdBuiltinStyle.wdStyleTOC7;
        internal const WdBuiltinStyle styleTOC8 = WdBuiltinStyle.wdStyleTOC8;
        internal const WdBuiltinStyle styleTOC9 = WdBuiltinStyle.wdStyleTOC9;
    }

    #endregion

    #region Variables
    internal struct InitialVariables
    {
        internal const string initialPageID = "_page_";
        internal const string initialVariableNames = "_variable_";

        internal const string initialVariableFieldNames = "_variable_field_";
        internal const string initialVariableServerNames = "_variable_server_";
        internal const string initialVariableTypeNames = "_variable_type_";
        internal const string initialVariableIDNames = "_variable_id_";
        internal const string initialVariableVersionNames = "_variable_version_";
        internal const string initialVariableOptionNames = "_variable_option_";
        internal const string initialVariableDocumentNames = "_variable_document_";

        internal const string initialPageChapters = "_page_Chapter";
    }

    internal enum VariableTypes
    {
        Server,
        Page,
        Field,
        Type,
        ID,
        Version,
        Option,
        Document,
    }

    internal enum VariableServerIDs
    {
        _variable_server_VersionNumber,
        _variable_server_UserToken,
        _variable_server_DocumentID,
        _variable_server_UpdatedAt,
        _variable_server_UpdatedFile,
        _variable_server_UpdatedConfig,
    }


    /// <summary>
    /// It is common in the tag name in ContentControl as well as the variable name in word
    ///In ContentControl, only the name is saved so that the index can be accessed from SectionBreak later
    ///The ContentControl identifier is also stored in the variable, which is unique, and if the user copies the relevant ContentControl and creates several of these items, he can identify which one it is.
    ///(one more will not be made)
    /// </summary>
    public enum PageIDs
    {
        //_variable_page_Independent,//TODO: split created section by user and mark as Independent

        _page_PersianIdentity,
        _page_InTheNameOfAllah,
        _page_DefenseMeeting,
        _page_LegalRights,
        _page_ResearchEthics,
        _page_Dedication,
        _page_Acknowledgment,
        _page_PersianAbstract,
        _page_TableOfContents,
        _page_TableOfFigures,
        _page_TableOfTables,
        _page_Abbreviations,
        _page_Preface,
        _page_Appendices,
        _page_Endnotes,
        _page_Glossary,
        _page_Sources,
        _page_EnglishAbstract,
        _page_EnglishIdentity,

        _page_Chapter1,
        _page_Chapter2,
        _page_Chapter3,
        _page_Chapter4,
        _page_Chapter5,
        _page_Chapter6,
        _page_Chapter7,
        _page_Chapter8,
        _page_Chapter9,
        _page_Chapter10,

        _page_TableOfMaps,
        _page_EvaluationForm,
        _page_OriginalityCommitment,

        _page_FormB,
    }
    internal enum VariableFieldIDs
    {
        _variable_field_NameOfCourse_Fa,//for School Research

        _variable_field_Abstract_Fa,

        _variable_field_University_Fa,
        _variable_field_University_En,

        _variable_field_Branch_Fa,//مثلا واحد یزد
        _variable_field_Branch_En,

        //_variable_field_Faculty_Fa,
        //_variable_field_Faculty_En,
        _variable_field_Department_Fa,
        _variable_field_Department_En,

        _variable_field_Group_Fa,
        _variable_field_Group_En,

        _variable_field_FieldOfStudy_Fa,
        _variable_field_FieldOfStudy_En,

        _variable_field_AreaOfStudy_Fa,
        _variable_field_AreaOfStudy_En,

        _variable_field_Title_Fa,
        _variable_field_Title_En,

        _variable_field_Supervisor_Fa,
        _variable_field_Supervisor_En,

        _variable_field_Advisor_Fa,
        _variable_field_Advisor_En,

        _variable_field_Author_Fa,
        _variable_field_Author_En,

        _variable_field_DefenseDate_Fa,
        _variable_field_DefenseDate_En,

        _variable_field_AcademicDegree_Fa,
        _variable_field_AcademicDegree_En,
    }
    internal enum VariableTypeIDs
    {
        _variable_type_Document,
        _variable_type_Template,
    }
    internal enum VariableIdentifierIDs
    {
        _variable_id_Hardware,
        _variable_id_Document,
        _variable_id_Template,
        _variable_id_University,
        _variable_id_AcademicDegree,
        _variable_id_GUID,
    }
    internal enum VariableVersionIDs
    {
        _variable_version_AddIn,
        _variable_version_Template,
    }
    internal enum VariableDocumentIDs
    {
        _variable_document_PositionCurrentPage,
        _variable_document_DefaultPersianFont,
        _variable_document_DefaultEnglishFont,
        _variable_document_SharedDocument,
    }
    internal enum VariableOptionIDs
    {
        _variable_option_BibliographyStyle,//APA,IEEE
        _variable_option_FootnoteSeparatorType,
        _variable_option_FootnoteStyle,
        _variable_option_FootnoteNumberStyle,
        _variable_option_TableOfContentsStyle,
        //_variable_other_VirastarSettings_CreatePoemType,
    }
    #endregion

    #region Bookmark And ContentControl
    //internal struct InitialBookmarks
    //{
    //	internal const string initialBookmarkNames = "bookmark_";

    //	internal const string initialBookmarkIcon = initialBookmarkNames + "Icon";

    //	internal const string initialBookmarkUniversityIcon = initialBookmarkNames + "Icon_University";

    //	internal const string initialBookmarkType = initialBookmarkNames + "Type";

    //}
    internal struct InitialContentControls
    {
        internal const string initialContentControlNames = "_field_";

        internal const string initialContentControlIcon = initialContentControlNames + "Icon";

        internal const string initialContentControlUniversityIcon = initialContentControlNames + "Icon_University";

        internal const string initialContentControlType = initialContentControlNames + "Type";
        internal const string initialPageID = "_page_";

        internal const string initialHidden = "_field_Hidden_";
        internal const string initialHiddenChapter = "_field_Hidden_Chapter";
        internal const string initialChapter = "_field_Chapter";

    }

    internal class Dictionaries
    {
        internal static string iconMark = "{Icon}";
    }
    internal enum UniversityContentControlNames
    {
        _university_DocumentID,
        _university_DocumentEditing
    }
    internal enum ContentControlNames
    {
        _field_Type_Fa,
        _field_Type_En,

        _field_Icon_University,
        _field_University_Fa,
        _field_Branch_Fa,
        //_field_Faculty_Fa,
        _field_Department_Fa,
        _field_Group_Fa,
        _field_FieldOfStudy_Fa,
        _field_AreaOfStudy_Fa,
        _field_Title_Fa,

        _field_Advisor_Title_Fa,
        _field_AreaOfStudy_Title_Fa,
        _field_Author_Title_Fa,
        _field_Supervisor_Title_Fa,
        _field_Title_Title_Fa,

        _field_Advisor_Title_En,
        _field_AreaOfStudy_Title_En,
        _field_Author_Title_En,
        _field_Supervisor_Title_En,
        _field_Title_Title_En,


        _field_Supervisor_Fa,
        _field_Advisor_Fa,
        _field_Author_Fa,
        _field_DefenseDate_Fa,
        _field_AcademicDegree_Fa,
        _field_NameOfCourse_Fa,//for School Research

        _field_InTheNameOfAllah,//TODO:Need?

        _field_Abstract_Fa,
        _field_Keywords_Fa,

        _field_Abstract_En,
        _field_Keywords_En,

        _field_Dedication_Fa,
        _field_Acknowledgment_Fa,

        _field_University_En,
        _field_Branch_En,
        //_field_Faculty_En,
        _field_Department_En,
        _field_Group_En,
        _field_FieldOfStudy_En,
        _field_AreaOfStudy_En,
        _field_Title_En,
        _field_Supervisor_En,
        _field_Advisor_En,
        _field_Author_En,
        _field_DefenseDate_En,
        _field_AcademicDegree_En,

        _field_DefenseLocation_Fa,
        _field_Examiner_Fa,
        _field_Examiner_Title_Fa,

        _field_Chapter1_Title,
        _field_Chapter2_Title,
        _field_Chapter3_Title,
        _field_Chapter4_Title,
        _field_Chapter5_Title,
        _field_Chapter6_Title,
        _field_Chapter7_Title,
        _field_Chapter8_Title,
        _field_Chapter9_Title,
        _field_Chapter10_Title,

        _field_Hidden_Chapter1_Title,
        _field_Hidden_Chapter2_Title,
        _field_Hidden_Chapter3_Title,
        _field_Hidden_Chapter4_Title,
        _field_Hidden_Chapter5_Title,
        _field_Hidden_Chapter6_Title,
        _field_Hidden_Chapter7_Title,
        _field_Hidden_Chapter8_Title,
        _field_Hidden_Chapter9_Title,
        _field_Hidden_Chapter10_Title,
    }
    #endregion

    #region Keyboard
    internal struct KeyboardShortcuts
    {
        internal const int documentsManager = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyF1;

        internal const int chatBoxNetworking = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyPeriod;

        internal const int insertDedicate = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyD;

        internal const int insertQuran = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyQ;
        internal const int insertNahjBalaghe = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyA;
        internal const int besmellahPageForm = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyZ;
        internal const int createProposal = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyL;

        internal const int changeContent = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyF2;
        internal const int addRemovePages = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyF3;
        internal const int setNormalStyle = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyN;
        internal const int setHeading2Style = (int)WdKey.wdKeyAlt + (int)WdKey.wdKey2;
        internal const int setHeading3Style = (int)WdKey.wdKeyAlt + (int)WdKey.wdKey3;
        internal const int setHeading4Style = (int)WdKey.wdKeyAlt + (int)WdKey.wdKey4;
        internal const int setHeading5Style = (int)WdKey.wdKeyAlt + (int)WdKey.wdKey5;
        internal const int insertShapeCaption = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyP;
        internal const int insertTableCaption = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyT;
        internal const int insertFormulaCaption = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyF;

        internal const int crossReferenceMenu = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyR;

        internal const int insertPersianFootnote = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyEquals;
        internal const int insertEnglishFootnote = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyHyphen;
        internal const int updateDocument = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyR;
        internal const int uploadDocument = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyU;
        internal const int insertHalfSpace = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyH;
        internal const int halfSpaceCorrection = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyV;
        internal const int neshanehGozariCorrection = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyB;
        internal const int standardCorrection = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeySemiColon;
        internal const int spellingCorrection = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyE;
        internal const int convertNumberToPersian = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeySlash;
        internal const int convertNumberToEnglish = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyPeriod;
        internal const int poemModeOneColumn = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyQ;
        internal const int poemModeTwoColumn = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyZ;

        internal const int changeBesmellahDialog = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyF2;

        internal const int insertCitationMenu = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyI;
        internal const int sourceManagementDialog = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyS;
        internal const int insertSources = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyM;
        internal const int importSourcesDialog = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyL;
        internal const int goToBibliography = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyG;

        internal const int exportToWord = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyW;
        internal const int exportAsGrayscale = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyG;
        internal const int exportToIdentification = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyI;
        internal const int exportToPDF = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyP;
        internal const int exportCDMenu = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyC;

        internal const int defenseAnnouncements = (int)WdKey.wdKeyAlt + (int)WdKey.wdKeyShift + (int)WdKey.wdKeyB;
    }
    #endregion
}
