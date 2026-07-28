using ShivaNegar.Constants;
using ShivaNegar.Models;
using Microsoft.Office.Interop.Word;
using ShivaNegar;
using ShivaNegar.Constants;
using ShivaNegar.Models;
using stdole;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using static ShivaNegar.DedicatedFunctions;
using ContentControl = Microsoft.Office.Interop.Word.ContentControl;


namespace ShivaNegar
{

    [ComVisible(true)]
    public class Ribbon : Microsoft.Office.Core.IRibbonExtensibility
    {
        internal List<RibbonControlModel> ribbonComponents;

        internal List<RibbonControlModel> ribbonControlledInCommands;

        public Microsoft.Office.Core.IRibbonUI ribbon;
        private Document doc;
        DedicatedFunctions.AccessType accessType;

        #region IRibbonExtensibility Members
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ShivaNegar.Ribbon.xml");
        }
        #endregion

        public Ribbon()
        {
            ribbonControlledInCommands = new List<RibbonControlModel>()
            {
                new RibbonControlModel(RibbonControlNames.builtInTabSave),
                new RibbonControlModel(RibbonControlNames.builtInTabPublish),
                new RibbonControlModel(RibbonControlNames.builtInTabShare),
                new RibbonControlModel(RibbonControlNames.builtInTabHelp),

                new RibbonControlModel(RibbonControlNames.builtInGroupPermissions),
                new RibbonControlModel(RibbonControlNames.builtInGroupPrepareForSharing),
                new RibbonControlModel(RibbonControlNames.builtInGroupVersions),

                new RibbonControlModel(RibbonControlNames.builtInGroupProtectedView),
                new RibbonControlModel(RibbonControlNames.builtInFilePermissionRestrictAs),
                new RibbonControlModel(RibbonControlNames.builtInFilePermission),
                new RibbonControlModel(RibbonControlNames.builtInFilePermissionView),
                new RibbonControlModel(RibbonControlNames.builtInFilePermissionRestrictMenu),
                new RibbonControlModel(RibbonControlNames.builtInFilePermissionUnrestricted),
                new RibbonControlModel(RibbonControlNames.builtInFilePermissionDoNotDistribute),
                new RibbonControlModel(RibbonControlNames.builtInProtectOrUnprotectDocument),
                new RibbonControlModel(RibbonControlNames.builtInProtectDocument),
                new RibbonControlModel(RibbonControlNames.builtInReviewProtectDocumentMenu),
                new RibbonControlModel(RibbonControlNames.builtInGroupProtect),
                new RibbonControlModel(RibbonControlNames.builtInReviewRestrictFormatting),

                new RibbonControlModel(RibbonControlNames.builtInGroupCode),
                new RibbonControlModel(RibbonControlNames.builtInVisualBasic),
                new RibbonControlModel(RibbonControlNames.builtInMacroPlay),
                new RibbonControlModel(RibbonControlNames.builtInMacroRecordOrStop),
                new RibbonControlModel(RibbonControlNames.builtInMacroRecorderPause),
                new RibbonControlModel(RibbonControlNames.builtInMacroSecurity),
                new RibbonControlModel(RibbonControlNames.builtInControlProperties),
                new RibbonControlModel(RibbonControlNames.builtInDesignMode),
            };


            RibbonControlModel[] galleryDocumentManagerItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id60",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ShivaNegarDocuments,true,true,"مدیریت اسناد","مدیریت اسناد","مدیریت اسناد","D"),
                new RibbonControlModel("__id61",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ShivaNegarDocuments,true,true,"اسناد بایگانی","اسناد بایگانی","اسناد بایگانی","D"),
                new RibbonControlModel("__id62",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ShivaNegarDocuments,true,true,"ایجاد سند جدید","ایجاد سند جدید","ایجاد سند جدید","D"),
            };
            RibbonControlModel[] galleryFootnoteItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id27",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.InsertFootnotePersian,true,true,"فارسی","درج پانویس فارسی","درج پانویس به صورت فارسی",null),
                new RibbonControlModel("__id28",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.InsertFootnoteEnglish,true,true,"انگلیسی","درج پانویس انگلیسی","درج پانویس به صورت انگلیسی",null),
            };
            RibbonControlModel[] gallerySimpleVirastarItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id38",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.SimpleVirastar,true,true,"درج نیم‌فاصله","یا (Ctrl+Shift+2) درج نیم‌فاصله","با استفاده از کلید های میانبری که تعریف شده اند شناسه نیم‌فاصله را به صورت استاندارد در متن استفاده کنید تا در محیط برنامه های دیگر تحت ویندوز نیز نیم‌فاصله ها به درستی نمایش داده شوند.",null),
                new RibbonControlModel("__id39",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.SimpleVirastar,true,true,"نیم‌فاصله گذاری","نیم‌فاصله گذاری","فاصله بین کلمات پرکاربرد، به عنوان مثال، افعالی که با \"می\" شروع می شوند و یا به \"اند\" ختم می شوند را در متن سند یا بخشی از متن که کاربر انتخاب کرده است به نیم فاصله تبدیل می کند.",null),
                new RibbonControlModel("__id40",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.SimpleVirastar,true,true,"نشانه گذاری (سجاوندی)","نشانه گذاری (سجاوندی)","کلمات تنوین دار نوشته شده در متن را ویرایش می کند. به عنوان مثال کلمات حتما یا حتمن را به حتماً تبدیل می کند.",null),
                new RibbonControlModel("__id41",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.SimpleVirastar,true,true,"غلط املایی","غلط املایی","غلط های املایی پیشبینی شده را تصحیح میکند",null),
            };
            RibbonControlModel[] galleryConvertNumberItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id29",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ConvertNumberPersian,true,true,"به فارسی","به فارسی","اعداد انگلیسی داخل سند یا محدوده انتخاب شده را به فارسی تبدیل می کند.",null),
                new RibbonControlModel("__id30",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ConvertNumberEnglish,true,true,"به انگلیسی","به انگلیسی","اعداد فارسی داخل سند یا محدوده انتخاب شده را به فارسی تبدیل می کند.",null),
            };
            RibbonControlModel[] galleryPoemModeItems = new RibbonControlModel[]
            {
                new RibbonControlModel("PoemMode1",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.OneColumnPoemMode,true,true,"تک ستونه",null,"",null),
                new RibbonControlModel("PoemMode2",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.TwoColumnPoemMode,true,true,"دو ستونه",null,"",null),
            };
            RibbonControlModel[] galleryInsertImportSourcesItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id31",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.InsertSources,true,true,"افزودن دستی","افزودن دستی","وارد کردن مشخصات منبع استفاده شده به صورت دستی",null),
                new RibbonControlModel("__id32",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.importSources,true,true,"ورود خودکار","ورود خودکار","وارد کردن مشخصات منبع استفاده شده، در قالب فایلی با فرمت لاتکس، که به عنوان مثال از گوگل اسکالر دانلود کرده اید.",null),
            };
            RibbonControlModel[] galleryCitationTypeItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id48",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ManageSources,true,true,"APA","APA","تغییر سبک نمایش منابع به APA",null),
                new RibbonControlModel("__id49",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ManageSources,true,true,"IEEE","IEEE","تغییر سبک نمایش منابع به IEEE",null),
            };
            RibbonControlModel[] galleryExportItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id34",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ExportGrayscale_2,true,true,"خروجی سیاه و سفید","خروجی سیاه و سفید","در این خروجی کلیه شکل های سند به صورت سیاه و سفید خواهند بود. این خروجی برای چاپ کاغذی سند به صورت سیاه و سفید مناسب است.",null),
                new RibbonControlModel("__id36",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ExportPDF,true,true,"PDF خروجی","PDF خروجی","سند را به قالبی تبدیل می کند که معمولا برای داوران محترم ارسال می شود. پس از اتمام نگارش و نهایی شدن سند از این خروجی استفاده کنید.",null),
                new RibbonControlModel("__id33",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ExportDocx,false,false,"Word خروجی","Word خروجی","نسخه ای از سند را در قالب ورد تهیه می کند که بدون نیاز به افزونه پارسانگار قابل استفاده است. در این نسخه دیگر امکان استفاده از قابلت های پارسانگار میسر نیست و برای ارائه به استاد راهنما، استاد مشاور، دفتر آموزش دانشگاه و سامانه ایرانداک مناسب است.",null),
            };

            RibbonControlModel[] galleryToolsItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id43",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ExportDocx,true,true,"صفحه بسم الله","صفحه بسم الله","",null),
                new RibbonControlModel("__id50",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.Dedication,true,true,"ویرایش/درج تقدیم","ویرایش/درج تقدیم","",null),
                new RibbonControlModel("__id44",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ExportCD,true,true,"برچسب لوح فشرده","برچسب لوح فشرده","PDF ایجاد طرح لوح فشرده و ذخیره در قالب",null),
                new RibbonControlModel("__id45",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ExportDocx,true,true,"اطلاعیه دفاع","اطلاعیه دفاع","",null),
                //new RibbonControlModel("__id46",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.ExportDocx,true,true,"سند پروپوزال","","",null),
            };

            RibbonControlModel[] galleryBooksItems = new RibbonControlModel[]
            {
                new RibbonControlModel("__id46",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.quran,true,true,"قرآن مجید","قرآن مجید","",null),
                new RibbonControlModel("__id47",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.PoemModeMenu,true,true,"نهج البلاغه","نهج البلاغه","",null),
            };

            RibbonControlModel[] galleryAboutItems = new RibbonControlModel[]
           {
                new RibbonControlModel("__id51",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.Tutorial,true,true,"آرشیو فیلم های راهنما","آرشیو فیلم های راهنما","",null),
                new RibbonControlModel("__id52",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.Web,true,true,"سامانه شیوانگار","سامانه شیوانگار","",null),
                new RibbonControlModel("__id53",RibbonControlModel.RibbonControlTypes.GalleryItem,Properties.ResourceRibbonIcons.Version,true,true,"درباره پارسانگار","درباره پارسانگار",null,null),

           };


            ribbonComponents = new List<RibbonControlModel>()
            {
                new RibbonControlModel(RibbonControlNames.tabJameNeger,RibbonControlModel.RibbonControlTypes.Tab,null,true,true,"پارسانگار",null,null,null),


                new RibbonControlModel(RibbonControlNames.btnLogin, RibbonControlModel.RibbonControlTypes.Button, Properties.ResourceRibbonIcons.EditFile, true, true, "ورود به حساب کاربری", "ورود به حساب کاربری", "برای دسترسی به امکانات افزونه وارد حساب کاربری خود شوید", null),


                new RibbonControlModel(RibbonControlNames.grpDocumentManager,RibbonControlModel.RibbonControlTypes.Group,null,true,true,"پارسانگار",null,null,null,null),
                new RibbonControlModel(RibbonControlNames.galleryDocumentsManager,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.ShivaNegarDocuments,true,true,"مدیریت اسناد","مدیریت اسناد",null,null,galleryDocumentManagerItems),
                new RibbonControlModel(RibbonControlNames.btnChatBoxNetworking,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.Messages,true,true,"تبادل نظر","تبادل نظر","گفگتو و تبادل نظر بین افراد مشترک مرتبط با سند","C"),
                new RibbonControlModel(RibbonControlNames.btnChangeContents,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.EditFile,true,true,"ویرایش اطلاعات","ویرایش اطلاعات","ویرایش برخی از اطلاعات شناسنامه ای سند جاری",null),
                new RibbonControlModel(RibbonControlNames.btnAddRemovePages,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.AddFile,true,true,"صفحه‌بندی","صفحه‌بندی و فصل‌بندی","حذف و اضافه صفحات اختیاری و مدیریت فصل های سند جاری",null),



                new RibbonControlModel(RibbonControlNames.grpFormat,RibbonControlModel.RibbonControlTypes.Group,null,true,true,"مدیریت متن",null,null,null,null),
                new RibbonControlModel(RibbonControlNames.btnTextFormating,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.DocumentText,true,true,"قالب بندی متن","قالب بندی متن","مقادیری مانند نوع قلم، اندازه قلم و تو رفتگی بند را براساس شیوه نامه تنظیم می کند.",null),

                new RibbonControlModel("menu1",RibbonControlModel.RibbonControlTypes.Menu,Properties.ResourceRibbonIcons.HeadingMenu,true,true,"عنوان بندی","عنوان بندی","مدیریت عنوان های سند؛ عنوان های مورد نظر کاربر را ایجاد کرده و سطح بندی آنها را تنظیم و مدیریت می کند. این افزونه قابلیت مدیریت و شماره بندی عنوان ها تا 5 سطح را دارد. شایان ذکر است عنوان سطح اول همان شماره فصل است.",null),
                new RibbonControlModel(RibbonControlNames.btnHeading2,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.Heading2,true,true,"عنوان سطح دو",null,null,null),
                new RibbonControlModel(RibbonControlNames.btnHeading3,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.Heading3,true,true,"عنوان سطح سه",null,null,null),
                new RibbonControlModel(RibbonControlNames.btnHeading4,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.Heading4,true,true,"عنوان سطح چهار",null,null,null),
                new RibbonControlModel(RibbonControlNames.btnHeading5,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.Heading5,true,true,"عنوان سطح پنج",null,null,null),


                new RibbonControlModel(RibbonControlNames.grpCaptionsAndRefer,RibbonControlModel.RibbonControlTypes.Group,null,true,true,"عنوان ها و ارجاع",null,null,null),
                new RibbonControlModel(RibbonControlNames.btnRefer,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.Refer,true,true,"ارجاع","ارجاع","این گزینه امکان ارجاع دادن به شکل، جدول و معادله را در متن سند فراهم می کند. لازم است قبل از ارجاع دادن عنوان شکل یا جدول نوشته شده باشد. برای معادله نیز به محض درج معادله امکان ارجاع دادن فراهم می شود.",null),

                new RibbonControlModel(RibbonControlNames.btnInsertWriteCaptionForShape,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.InsertCaptionImage,true,true,"عنوان شکل","عنوان شکل","ضمن مدیریت و شماره گذاری شکل ها، قالب بندی لازم برای نوشتن عنوان شکل را فراهم می کند تا بتوان به صورت خودکار فهرست شکل های داخل سند را تهیه کرد.",null),
                new RibbonControlModel(RibbonControlNames.btnInsertWriteCaptionForTable,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.InsertCaptionTable,true,true,"عنوان جدول","عنوان جدول","ضمن مدیریت و شماره گذاری جدول ها، قالب بندی لازم برای نوشتن عنوان جدول را فراهم می کند تا بتوان به صورت خودکار فهرست جدول های داخل سند را تهیه کرد.",null),
                new RibbonControlModel(RibbonControlNames.btnInsertCaptionForFormula,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.InsertCaptionFormula,true,true,"معادله نویسی","معادله نویسی","امکان نوشتن معادله و شماره گذاری آن را، به منظور ارجاع دادن به معادله در متن سند، فراهم می کند.",null),
                new RibbonControlModel(RibbonControlNames.grpCaptionsAndRefer__btn,RibbonControlModel.RibbonControlTypes.DialogBoxLauncherButton,null,true,true,null,null,null,null),
                new RibbonControlModel(RibbonControlNames.grpFormat__btn,RibbonControlModel.RibbonControlTypes.DialogBoxLauncherButton,null,true,true,null,null,null,null),



                new RibbonControlModel(RibbonControlNames.grpFootnotes,RibbonControlModel.RibbonControlTypes.Group,null,true,true,"پانویس",null,null,null),
                new RibbonControlModel(RibbonControlNames.galleryInsertFootnote,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.InsertFootnoteMenu,true,true,"درج پانویس (پاورقی)","درج پانویس (پاورقی)","امکان نوشتن پانویس و نوشتن توضیحات لازم برای کلمه مورد نظر را به دو صورت فارسی و انگلیسی فراهم می کند. شیوه نمایش پانویس را می توانید در کادر محاوره تنظیماتی که در همین بخش ارائه شده است تغییر دهید.",null,galleryFootnoteItems),
                new RibbonControlModel(RibbonControlNames.grpFootnotes__btn,RibbonControlModel.RibbonControlTypes.DialogBoxLauncherButton,null,true,true,null,null,null,null),


                new RibbonControlModel(RibbonControlNames.grpVirastar,RibbonControlModel.RibbonControlTypes.Group,null,true,true,"ویراستاری",null,null,null),
                new RibbonControlModel(RibbonControlNames.gallerySimpleVirastar,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.SimpleVirastar,true,true,"ویراستاری مقدماتی","ویراستاری مقدماتی","ویراستاری مقدماتی برخی از نکات ویرایشی مانند نیم فاصله گذاری، نشانه گذاری (سجاوندی) و حذف فضاهای خالی اضافه در متن سند را انجام می دهد. بدیهی است این بخش فرآیندهای ذکر شده را به صورت کامل انجام نمی دهد ولی بیشتر از 70 درصد از اصلاحات ذکر شده را در متن انجام خواهد داد.",null,gallerySimpleVirastarItems),

                new RibbonControlModel(RibbonControlNames.btnStandardCorrection,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.SimpleVirastar,true,true,"زبان معیار","زبان معیار","در متن سند واژه های بیگانه را با واژه فارسی مصوب فرهنگستان زبان و ادب فارسی جایگزین می‌کند.",null),
                new RibbonControlModel(RibbonControlNames.galleryConvertNumber,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.ConvertNumberMenu,true,true,"تبدیل اعداد","تبدیل اعداد","فرآیند تبدیل اعداد، به خصوص اعداد اعشاری، را از فارسی به انگلیسی یا بالعکس در تمام متن یا بخشی از متن که انتخاب شده است انجام می دهد.",null,galleryConvertNumberItems),
                new RibbonControlModel(RibbonControlNames.galleryPoemMode,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.PoemModeMenu,true,true,"قالب شعر","قالب شعر","اشعار فارسی را به صورت یک ستونه یا دو ستونه می نویسد. برای نوشتن شعر به صورت زیر عمل کنید:\r\n1-  مصرع اول شعر خود را به عنوان یک پاراگراف جدید بنویسید. \r\n2-بعد از پایان هر مصرع و برای رفتن به مصرع بعد از کلید اینتر استفاده کنید.\r\n3-پس از پایان شعر تمام مصرع های نوشته شده را انتخاب کنید.\r\n4-با توجه به حالتی که می خواهید شعر نوشته شود گزینه یک ستونه یا دو ستونه را انتخاب کنید.",null,galleryPoemModeItems),
                new RibbonControlModel(RibbonControlNames.grpVirastar__btn,RibbonControlModel.RibbonControlTypes.DialogBoxLauncherButton,null,true,true,null,null,null,null),



                new RibbonControlModel(RibbonControlNames.grpCitationManager,RibbonControlModel.RibbonControlTypes.Group,null,true,true,"مدیریت استنادهای علمی",null,null,null),
                new RibbonControlModel(RibbonControlNames.btnInsertCitation,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.InsertCitation,true,true,"استناد دهی","استناد دهی","استناد دادن به منبع علمی استفاده شده در متن سند" ,null),
                new RibbonControlModel(RibbonControlNames.galleryInsertImportSources,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.AddSources,true,true,"افزودن منبع","افزودن منبع","اطلاعات منبع جدیدی که در تحقیق استفاده شده است را به صورت دستی یا توسط فایلی با فرمت لاتکس به لیست منابع اضافه می‌کند",null,galleryInsertImportSourcesItems),
                new RibbonControlModel(RibbonControlNames.galleryCitationType,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.ManageSources,true,true,"شیوه","سبک نمایش منابع","",null,galleryCitationTypeItems),
                new RibbonControlModel(RibbonControlNames.btnManageResources,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.ManageSources,true,true,"مدیریت منابع","مدیریت منابع","منابع استفاده شده در سند را مدیریت کرده و آنها را ویرایش کرده و یا اینکه به لیست اضافه یا حذف می کند. ",null),
                new RibbonControlModel(RibbonControlNames.btnGoToBibliography,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.GoToBibliography,true,true,"فهرست منابع","فهرست منابع","فهرست منابع ثبت شده در انتهای سند را نمایش می دهد.",null),
                //new RibbonControlModel(RibbonControlNames.grpCitationManager__btn,RibbonControlModel.RibbonControlTypes.DialogBoxLauncherButton,null,true,true,null,null,null,null),



                new RibbonControlModel(RibbonControlNames.grpOtherOptions,RibbonControlModel.RibbonControlTypes.Group,null,true,true,"مدیریت سند",null,null,null),
                new RibbonControlModel(RibbonControlNames.btnUpdateDocument,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.UpdateDocument,true,true,"بروزرسانی سند","بروزرسانی سند","فهرست ها، ارجاع ها و استنادهای داخل سند را به روز رسانی می کند.",null),
                new RibbonControlModel(RibbonControlNames.btnUploadDocument,RibbonControlModel.RibbonControlTypes.Button,Properties.ResourceRibbonIcons.UploadDocument,true,true,"همرسانی سند","همرسانی سند","درصورت استفاده از این گزینه برای نخستین بار، یک نسخه پشتیبان از سند جاری به حساب کاربری در فضای ابری منتقل می شود و در دفعات بعد تغییرات ایجاد شده در سند جاری با سند موجود در حساب کاربری همرسانی می گردد.",null),
                new RibbonControlModel(RibbonControlNames.galleryExport,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.ExportMenu,true,true,"خروجی سند","خروجی سند","انواع خروجی های مورد نیاز کاربر در این قسمت ارائه شده اند مانند تهیه نسخه ای از سند به صورت سیاه و سفید، تولید نسخه ای از سند برای انجام فرآیند همانندجویی و...",null,galleryExportItems),
                new RibbonControlModel(RibbonControlNames.grpOtherOptions__btn,RibbonControlModel.RibbonControlTypes.DialogBoxLauncherButton,null,true,true,null,null,null,null),

                new RibbonControlModel(RibbonControlNames.grpTools,RibbonControlModel.RibbonControlTypes.Group,null,true,true,"ابزار ها و کتاب ها",null,null,null),
                new RibbonControlModel(RibbonControlNames.galleryTools,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.tools,true,true,"ابزارهای کاربردی","ابزارهای کاربردی",null,null,galleryToolsItems),
                new RibbonControlModel(RibbonControlNames.galleryBooks,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.quran,true,true,"کتاب های مرجع","کتاب های مرجع",null,null,galleryBooksItems),
                new RibbonControlModel(RibbonControlNames.galleryAbout,RibbonControlModel.RibbonControlTypes.Gallery,Properties.ResourceRibbonIcons.About,true,true,"درباره","درباره","درباره.",null,galleryAboutItems),

            };
        }


        #region BuiltIn Ribbon Commands
        public bool getEnabledCommands(Microsoft.Office.Core.IRibbonControl control)
        {
            if (Globals.ThisAddIn.Application.Documents.Count != 0)
            {
                Document doc;
                try
                {
                    doc = Globals.ThisAddIn.Application.ActiveDocument;
                }
                catch (Exception)
                {
                    return false;
                }

                if (StringConstant.AdministratorAccounts.Contains(Properties.Settings.Default.Mobile))
                    return true;
                else if (DedicatedFunctions.hasAccess(doc) == DedicatedFunctions.AccessType.AccessGranted)
                    return false;
                else
                    return true;

                //bool status = ribbonControlledInCommands.Where(p => p.Id == control.Id).First().Enable;
                //return status;
            }
            else
                return true;
        }
        #endregion

        #region Get Callbacks
        public IPictureDisp getImage(Microsoft.Office.Core.IRibbonControl control)
        {
            Bitmap bitmap = ribbonComponents.Where(p => p.Id == control.Id).First().Image;
            if (bitmap != null)
                return PictureDispConverter.ToIPictureDisp(bitmap);
            else
            {
                return PictureDispConverter.ToIPictureDisp(Properties.ResourceRibbonIcons.ShivaNegarDocuments);
            }
        }
        public bool getEnabled(Microsoft.Office.Core.IRibbonControl control)
        {
            if (control.Id == "btnLogin")
            {
                // دکمه لاگین همیشه فعال باشد
                return true;
            }
            if (control.Id == "galleryDocumentsManager")
            {
                return true;
            }
            else
            {
                bool status = ribbonComponents.Where(p => p.Id == control.Id).First().Enable;
                return status;
            }
            //return true;
        }

        public bool getVisible(Microsoft.Office.Core.IRibbonControl control)
        {
            bool isLoggedIn = !string.IsNullOrEmpty(Properties.Settings.Default.UserToken);

            // کنترل‌های خاص
            if (control.Id == RibbonControlNames.btnLogin)
            {
                // دکمه لاگین: فقط در صورت لاگین نبودن
                return !isLoggedIn;
            }
            if (control.Id == RibbonControlNames.galleryDocumentsManager)
            {
                // گالری: فقط در صورت لاگین بودن
                return isLoggedIn;
            }

            // بقیه کنترلها از تنظیمات اصلی
            bool status = ribbonComponents.Where(p => p.Id == control.Id).First().Visible;
            return status;
        }

        public string getLabel(Microsoft.Office.Core.IRibbonControl control)
        {
            string text = ribbonComponents.Where(p => p.Id == control.Id).First().Label;
            return text;
        }
        public string getDescription(Microsoft.Office.Core.IRibbonControl control)
        {
            string text = ribbonComponents.Where(p => p.Id == control.Id).First().Description;
            return text;
        }
        public string getScreentip(Microsoft.Office.Core.IRibbonControl control)
        {
            string text = ribbonComponents.Where(p => p.Id == control.Id).First().ScreenTip;
            return text;
        }
        public string getSupertip(Microsoft.Office.Core.IRibbonControl control)
        {
            string text = ribbonComponents.Where(p => p.Id == control.Id).First().SuperTip;
            return text;
        }
        public string getKeytip(Microsoft.Office.Core.IRibbonControl control)
        {
            string text = ribbonComponents.Where(p => p.Id == control.Id).First().KeyTip;
            return text;
        }

        public int getItemCount(Microsoft.Office.Core.IRibbonControl control)
        {
            object content = ribbonComponents.Where(p => p.Id == control.Id).First().Content;
            if (content == null)
                return 0;

            RibbonControlModel[] ribbonControlModels = (RibbonControlModel[])content;
            if (control.Id == RibbonControlNames.galleryExport)
            {
                if (StringConstant.AdministratorAccounts.Contains(Properties.Settings.Default.Mobile))
                    return ribbonControlModels.Length;
                else
                    return ribbonControlModels.Length - 1;
            }
            else
            {
                int count = ribbonControlModels.Length;
                return count;
            }
        }
        public string getItemLabel(Microsoft.Office.Core.IRibbonControl control, int index)
        {
            object content = ribbonComponents.Where(p => p.Id == control.Id).First().Content;
            if (content == null)
                return "";

            string text = ((RibbonControlModel[])content)[index].Label;
            return text;
        }
        public string getItemSupertip(Microsoft.Office.Core.IRibbonControl control, int index)
        {
            object content = ribbonComponents.Where(p => p.Id == control.Id).First().Content;
            if (content == null)
                return "";

            string text = ((RibbonControlModel[])content)[index].SuperTip;
            return text;
        }
        public string getItemScreentip(Microsoft.Office.Core.IRibbonControl control, int index)
        {
            object content = ribbonComponents.Where(p => p.Id == control.Id).First().Content;
            if (content == null)
                return "";

            string text = ((RibbonControlModel[])content)[index].ScreenTip;
            return text;
        }
        public IPictureDisp getItemImage(Microsoft.Office.Core.IRibbonControl control, int index)
        {
            object content = ribbonComponents.Where(p => p.Id == control.Id).First().Content;

            Bitmap bitmap = ((RibbonControlModel[])content)[index].Image;
            if (bitmap != null)
            {
                return PictureDispConverter.ToIPictureDisp(bitmap);
            }
            return null;
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Microsoft.Office.Core.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public void btnLogin_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.documentsManager();
        }
        public void btnShivaNegar_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Ribbon.setTabProperties("", true);
        }
        //public void btnDocumentsManager_Click(Microsoft.Office.Core.IRibbonControl control)
        //{
        //    Globals.ThisAddIn.documentsManager();
        //}
        public void btnChatBoxNetworking_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.chatBoxNetworking();
        }
        public void btnChangeContents_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.changeContent();
        }
        public void btnAddRemovePages_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.addRemovePages();
        }

        public void galleryDocumentManager_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {

            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.documentsManager();
            }
            else if (selectedIndex == 1)
            {


                Globals.ThisAddIn.documentsManagerArchive();


            }
            else if (selectedIndex == 2)
            {


                Globals.ThisAddIn.documentsManagerCreate();

            }

        }

        #region DialogLauncher 
        public void grpFormat_DialogLauncherClick(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.formatSettings();
        }
        public void grpCaptionsAndRefer_DialogLauncherClick(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.captionSettings();
        }
        public void grpFootnotes_DialogLauncherClick(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.footnoteSettings();
        }
        public void grpVirastar_DialogLauncherClick(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.virastarSettings();
        }
        public void grpOtherOptions_DialogLauncherClick(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.listsManagerSettings();
        }
        public void grpCitationManager_DialogLauncherClick(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.citationSettings();
        }


        #endregion

        #region Change Style Selection
        public void btnTextFormating_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.setNormalStyle();
        }
        public void btnHeading2_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.setHeadingStyle(2);
        }
        public void btnHeading3_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.setHeadingStyle(3);
        }
        public void btnHeading4_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.setHeadingStyle(4);
        }
        public void btnHeading5_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.setHeadingStyle(5);
        }
        #endregion

        #region Caption
        public void btnInsertCaptionForShape_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.InsertCaption(0);
        }
        public void btnInsertCaptionForTable_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.InsertCaption(1);
        }
        public void btnInsertCaptionForFormula_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.InsertCaption(2);
        }
        #endregion

        #region Refer
        public void btnRefer_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.ShowCrossReferenceMenu();
        }
        #endregion

        #region Table And Lists
        public void btnUpdateDocument_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.updateDocument();
        }
        public void btnUploadDocument_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.uploadDocument();
        }
        #endregion

        #region References And Resources
        public void galleryInsertImportSources_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.insertSources();
            }
            else if (selectedIndex == 1)
            {
                Globals.ThisAddIn.importSourcesDialog();
            }
        }
        public void galleryCitationType_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            //Document doc;
            //DedicatedFunctions.AccessType accessType;
            if (selectedIndex == 0)
            {
                string value = "APA";
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر سبک نمایش منابع به " + value);

                DedicatedFunctions.setORAddStaticVariableValue(doc, VariableOptionIDs._variable_option_BibliographyStyle.ToString(), value);
                Globals.ThisAddIn.Application.Options.BibliographyStyle = value;
                doc.Bibliography.BibliographyStyle = value;

                DedicatedFunctions.updateBibliography(doc, doc.ActiveWindow.Selection, accessType);

                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
            else if (selectedIndex == 1)
            {
                string value = "IEEE";
                Globals.ThisAddIn.Application.UndoRecord.StartCustomRecord("تغییر سبک نمایش منابع به " + value);

                DedicatedFunctions.setORAddStaticVariableValue(doc, VariableOptionIDs._variable_option_BibliographyStyle.ToString(), value);
                Globals.ThisAddIn.Application.Options.BibliographyStyle = value;
                doc.Bibliography.BibliographyStyle = value;

                DedicatedFunctions.updateBibliography(doc, doc.ActiveWindow.Selection, accessType);

                Globals.ThisAddIn.Application.UndoRecord.EndCustomRecord();
            }
        }

        public void btnInsertCitation_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.insertCitationMenu();
        }
        public void btnManageResources_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.sourceManagementDialog();
        }
        public void btnGoToBibliography_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.goToBibliography();
        }

        #endregion

        #region Virastar
        public void galleryConvertNumber_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.convertNumber(true);
            }
            else if (selectedIndex == 1)
            {
                Globals.ThisAddIn.convertNumber(false);
            }
        }

        public void gallerySimpleVirastar_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.insertHalfSpace();
            }
            else if (selectedIndex == 1)
            {
                Globals.ThisAddIn.halfSpaceCorrection();
            }
            else if (selectedIndex == 2)
            {
                Globals.ThisAddIn.neshanehGozariCorrection();
            }
            else if (selectedIndex == 3)
            {
                Globals.ThisAddIn.spellingCorrection();
            }
        }
        public void btnStandardCorrection_Click(Microsoft.Office.Core.IRibbonControl control)
        {
            Globals.ThisAddIn.standardCorrection();
        }

        public void galleryPoemMode_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.poemMode(1);
            }
            else if (selectedIndex == 1)
            {
                Globals.ThisAddIn.poemMode(2);
            }
        }

        #endregion

        #region Export
        public void galleryExport_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.exportAsGrayscale();
            }
            else if (selectedIndex == 1)
            {
                Globals.ThisAddIn.exportToIdentification();
            }
            else if (selectedIndex == 2)
            {
                Globals.ThisAddIn.exportToPDF();
            }
            else if (selectedIndex == 3)
            {
                Globals.ThisAddIn.exportToLatex();
            }
            else if (selectedIndex == 4)
            {
                if (StringConstant.AdministratorAccounts.Contains(Properties.Settings.Default.Mobile))
                    Globals.ThisAddIn.exportToWord();
            }
        }
        #endregion
        public void galleryTools_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.besmellahPageForm();
            }
            else if (selectedIndex == 1)
            {
                Globals.ThisAddIn.insertDedicate();
            }
            else if (selectedIndex == 2)
            {
                Globals.ThisAddIn.exportCDMenu();
            }
            else if (selectedIndex == 3)
            {
                Globals.ThisAddIn.defenseAnnouncements();
            }
            //else if (selectedIndex == 4)
            //{
            //    Globals.ThisAddIn.createProposal();
            //}
        }
        public void galleryBooks_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.insertQuran();
            }
            else if (selectedIndex == 1)
            {
                Globals.ThisAddIn.insertNahjBalaghe();
            }
        }

        public void galleryAbout_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                try
                {
                    System.Diagnostics.Process.Start("https://shivanegar.ir/videos/4");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show($"خطا در باز کردن وبسایت: {ex.Message}");
                }
            }
            else if (selectedIndex == 1)
            {
                try
                {
                    System.Diagnostics.Process.Start("https://shivanegar.ir/");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show($"خطا در باز کردن وبسایت: {ex.Message}");
                }
            }
            else if (selectedIndex == 2)
            {
                Globals.ThisAddIn.ShowAbout();
            }
        }

        #region Footnote
        public void galleryInsertFootnote_Click(Microsoft.Office.Core.IRibbonControl control, string selectedId, int selectedIndex)
        {
            if (selectedIndex == 0)
            {
                Globals.ThisAddIn.insertFootnote(true);
            }
            else if (selectedIndex == 1)
            {
                Globals.ThisAddIn.insertFootnote(false);
            }
        }
        #endregion

        #endregion

        #region Helpers
        public static void setTabProperties(string label, bool activate)
        {
            string id = RibbonControlNames.tabJameNeger;

            RibbonControlModel rp = Globals.Ribbon?.ribbonComponents?.Where(p => p.Id == id).First();

            if (rp != null)
            {
                // موقتی هست، لازمه بررسی بشه که استاد راهنما یا افراد دیگه باز کردند یا خود دانشجو اینکار رو کرده
                Document doc;
                string limitUsage = "";
                try
                {
                    doc = Globals.ThisAddIn?.Application?.ActiveDocument;
                    if (doc != null)
                    {
                        if (DedicatedFunctions.getStaticVariableValue(doc, VariableDocumentIDs._variable_document_SharedDocument.ToString()) == "1")
                            limitUsage = " (دسترسی محدود)";
                    }
                }
                catch (Exception)
                {
                }


                if (label != null && label != "")
                    rp.Label = label + limitUsage;
                else
                {
                    label = StringConstant.NameOfProject + limitUsage;
                }

                //if(visible != null)
                //	rp.Visible = (bool)visible;

                Globals.Ribbon?.ribbon?.InvalidateControl(id);

                if (activate)
                    Globals.Ribbon?.ribbon?.ActivateTab(id);
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        #region Shortcut
        internal static void loadKeyboardShortcut()
        {
            if (!Globals.ThisAddIn.SetKeyBindingStatus)
            {
                string templatesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates", "ShivaNegarTemplates");

                //string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                //string templatesLocation = System.IO.Path.Combine(appDataPath , @"Microsoft\Templates\");
                String fullPath = "";
                try
                {
                    fullPath = DedicatedFunctions.copyFileToFolder(DedicatedFunctions.getStream(EmbeddedResourceNames.ShivaNegarShortcut), nameof(EmbeddedResourceNames.ShivaNegarShortcut) + ".dotm", templatesPath);
                }
                catch (Exception)
                {
                    // on using Template File
                    fullPath = templatesPath + "\\" + nameof(EmbeddedResourceNames.ShivaNegarShortcut) + ".dotm";
                }

                if (new FileInfo(fullPath).Exists)
                {
                    Globals.Shortcut = Globals.ThisAddIn.Application.AddIns.Add(fullPath, true);
                    //Globals.Shortcut.Installed = true;

                    //object cContext = Application.CustomizationContext;
                    //Application.CustomizationContext = Application.Templates[Globals.Shortcut.Path + "\\" + Globals.Shortcut.Name];

                    setKeyBinding(getKeyboardRelations());
                    //Application.CustomizationContext = cContext;
                    //Application.Documents[Globals.Shortcut.Name].Saved = true;
                }
                else
                {
                    DedicatedFunctions.ShowErrorMessage("خطایی در تنظیم کلید های میانبر به وجود آمد", (int)ErrorCodes.CopyShortcutTemplate, StringConstant.SupportEmail);
                }
            }
        }

        internal static void unloadKeyboardShortcut()
        {
            if (Globals.Shortcut != null)
            {
                try
                {
                    Globals.Shortcut.Installed = false;
                    if (File.Exists(Globals.Shortcut.Path + "\\" + Globals.Shortcut.Name))
                    {
                        foreach (Document item in Globals.ThisAddIn.Application.Documents)
                        {
                            if (item.Name == Globals.Shortcut.Name && item.Path == Globals.Shortcut.Path)
                            {
                                DedicatedFunctions.saveDocument(item);
                                item.Close(WdSaveOptions.wdDoNotSaveChanges);
                                break;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    DedicatedFunctions.removeFileFromSystem(Globals.Shortcut.Path + "\\" + Globals.Shortcut.Name);
                }
                catch (Exception)// on using Template File
                {
                }
                try
                {
                    Globals.Shortcut.Delete();
                }
                catch (Exception)// on using Template File
                {
                }

            }
            unsetKeyBinding(getKeyboardRelations());
        }
        internal static List<KeyboardRelationModel> getKeyboardRelations()
        {
            return new List<KeyboardRelationModel>
                {
                    //new KeyboardRelationModel("documentsManager",KeyboardShortcuts.documentsManager,RibbonControlNames.btnDocumentsManager,"\n(Alt+F1)"),
                    new KeyboardRelationModel("changeContent",KeyboardShortcuts.changeContent,RibbonControlNames.btnChangeContents,"\n(Alt+F2)"),
                    new KeyboardRelationModel("addRemovePages",KeyboardShortcuts.addRemovePages,RibbonControlNames.btnAddRemovePages,"\n(Alt+F3)"),

                    new KeyboardRelationModel("setNormalStyle",KeyboardShortcuts.setNormalStyle,RibbonControlNames.btnTextFormating,"\n(Alt+N)"),
                    new KeyboardRelationModel("setHeading2Style",KeyboardShortcuts.setHeading2Style,RibbonControlNames.btnHeading2,"\n(Alt+2)"),
                    new KeyboardRelationModel("setHeading3Style",KeyboardShortcuts.setHeading3Style,RibbonControlNames.btnHeading3,"\n(Alt+3)"),
                    new KeyboardRelationModel("setHeading4Style",KeyboardShortcuts.setHeading4Style,RibbonControlNames.btnHeading4,"\n(Alt+4)"),
                    new KeyboardRelationModel("setHeading5Style",KeyboardShortcuts.setHeading5Style,RibbonControlNames.btnHeading5,"\n(Alt+5)"),

                    new KeyboardRelationModel("insertShapeCaption",KeyboardShortcuts.insertShapeCaption,RibbonControlNames.btnInsertWriteCaptionForShape,"\n(Alt+P)"),
                    new KeyboardRelationModel("insertTableCaption",KeyboardShortcuts.insertTableCaption,RibbonControlNames.btnInsertWriteCaptionForTable,"\n(Alt+T)"),
                    new KeyboardRelationModel("insertFormulaCaption",KeyboardShortcuts.insertFormulaCaption,RibbonControlNames.btnInsertCaptionForFormula,"\n(Alt+F)"),

                    new KeyboardRelationModel("crossReferenceMenu",KeyboardShortcuts.crossReferenceMenu,RibbonControlNames.btnRefer,"\n(Alt+Shift+R)"),

                    new KeyboardRelationModel("insertPersianFootnote",KeyboardShortcuts.insertPersianFootnote,RibbonControlNames.galleryInsertFootnote,"\n(Alt+=)",0),
                    new KeyboardRelationModel("insertEnglishFootnote",KeyboardShortcuts.insertEnglishFootnote,RibbonControlNames.galleryInsertFootnote,"\n(Alt+-)",1),


                    new KeyboardRelationModel("updateDocument",KeyboardShortcuts.updateDocument,RibbonControlNames.btnUpdateDocument,"\n(Alt+R)"),
                    new KeyboardRelationModel("uploadDocument",KeyboardShortcuts.uploadDocument,RibbonControlNames.btnUploadDocument,"\n(Alt+U)"),

                    new KeyboardRelationModel("insertHalfSpace",KeyboardShortcuts.insertHalfSpace,RibbonControlNames.gallerySimpleVirastar,"\n(Alt+H)",0),
                    new KeyboardRelationModel("halfSpaceCorrection",KeyboardShortcuts.halfSpaceCorrection,RibbonControlNames.gallerySimpleVirastar,"\n(Alt+V)",1),
                    new KeyboardRelationModel("neshanehGozariCorrection",KeyboardShortcuts.neshanehGozariCorrection,RibbonControlNames.gallerySimpleVirastar,"\n(Alt+B)",2),
                    new KeyboardRelationModel("spellingCorrection",KeyboardShortcuts.spellingCorrection,RibbonControlNames.gallerySimpleVirastar,"\n(Alt+E)",3),

                    new KeyboardRelationModel("standardCorrection",KeyboardShortcuts.standardCorrection,RibbonControlNames.btnStandardCorrection,"\n(Alt+;)"),

                    new KeyboardRelationModel("convertNumberToPersian",KeyboardShortcuts.convertNumberToPersian,RibbonControlNames.galleryConvertNumber,"\n(Alt+/)",0),
                    new KeyboardRelationModel("convertNumberToEnglish",KeyboardShortcuts.convertNumberToEnglish,RibbonControlNames.galleryConvertNumber,"\n(Alt+.)",1),
                    new KeyboardRelationModel("poemModeOneColumn",KeyboardShortcuts.poemModeOneColumn,RibbonControlNames.galleryPoemMode,"\n(Alt+q)",0),
                    new KeyboardRelationModel("poemModeTwoColumn",KeyboardShortcuts.poemModeTwoColumn,RibbonControlNames.galleryPoemMode,"\n(Alt+z)",1),

                    new KeyboardRelationModel("insertCitationMenu",KeyboardShortcuts.insertCitationMenu,RibbonControlNames.btnInsertCitation,"\n(Alt+I)"),
                    new KeyboardRelationModel("sourceManagementDialog",KeyboardShortcuts.sourceManagementDialog,RibbonControlNames.btnManageResources,"\n(Alt+S)"),
                    new KeyboardRelationModel("insertSources",KeyboardShortcuts.insertSources,RibbonControlNames.galleryInsertImportSources,"\n(Alt+M)",0),//insert
					new KeyboardRelationModel("importSourcesDialog",KeyboardShortcuts.importSourcesDialog,RibbonControlNames.galleryInsertImportSources,"\n(Alt+L)",1),//import
					new KeyboardRelationModel("goToBibliography",KeyboardShortcuts.goToBibliography,RibbonControlNames.btnGoToBibliography,"\n(Alt+G)"),

					//new KeyboardRelations("exportToWord",KeyboardShortcuts.exportToWord,RibbonControlNames.galleryExport,"\n(Alt+Shift+W)",-1),
					new KeyboardRelationModel("exportAsGrayscale",KeyboardShortcuts.exportAsGrayscale,RibbonControlNames.galleryExport,"\n(Alt+Shift+G)",0),
                    new KeyboardRelationModel("exportToIdentification",KeyboardShortcuts.exportToIdentification,RibbonControlNames.galleryExport,"\n(Alt+Shift+I)",1),
                    new KeyboardRelationModel("exportToPDF",KeyboardShortcuts.exportToPDF,RibbonControlNames.galleryExport,"\n(Alt+Shift+P)",2),


                    new KeyboardRelationModel("chatBoxNetworking",KeyboardShortcuts.chatBoxNetworking,RibbonControlNames.btnChatBoxNetworking,"\n(Alt+Shift+.)"),

                    new KeyboardRelationModel("besmellahPageForm",KeyboardShortcuts.besmellahPageForm,RibbonControlNames.galleryTools,"\n(Alt+Shift+Z)",0),
                    new KeyboardRelationModel("insertDedicate",KeyboardShortcuts.insertDedicate,RibbonControlNames.galleryTools,"\n(Alt+Shift+D)",1),
                    new KeyboardRelationModel("exportCDMenu",KeyboardShortcuts.exportCDMenu,RibbonControlNames.galleryTools,"\n(Alt+Shift+C)",2),
                    new KeyboardRelationModel("defenseAnnouncements",KeyboardShortcuts.exportCDMenu,RibbonControlNames.galleryTools,"\n(Alt+Shift+B)",3),
                    //new KeyboardRelationModel("createProposal",KeyboardShortcuts.createProposal,RibbonControlNames.galleryTools,"\n(Alt+Shift+L)",4),

                    new KeyboardRelationModel("insertQuran",KeyboardShortcuts.insertQuran,RibbonControlNames.galleryBooks,"\n(Alt+Shift+Q)",0),
                    new KeyboardRelationModel("insertNahjBalaghe",KeyboardShortcuts.insertNahjBalaghe,RibbonControlNames.galleryBooks,"\n(Alt+Shift+A)",1),

					//new KeyboardRelations("documentManagerSettingsMenu",null,null,"\n()"),
					//new KeyboardRelations("captionSettings",null,null,"\n()"),
					//new KeyboardRelations("virastarSettings",null,null,"\n()"),
					//new KeyboardRelations("footnoteSettings",null,null,"\n()"),
					//new KeyboardRelations("listsManagerSettings",null,null,"\n()"),
					//new KeyboardRelations("citationSettings",null,null,"\n()"),
				};
        }
        internal static void setKeyBinding(List<KeyboardRelationModel> keyboardRelations)
        {
            try
            {
                foreach (KeyboardRelationModel kr in keyboardRelations)
                {
                    Globals.ThisAddIn.Application.KeyBindings.Add(WdKeyCategory.wdKeyCategoryMacro, kr.CallFunctionName, kr.KeyboardShortcut);

                    if (kr.RibbonControlID != null)
                    {
                        var ribbonProperties = Globals.Ribbon.ribbonComponents.Where(p => p.Id == kr.RibbonControlID).First();

                        if (ribbonProperties == null)
                            continue;

                        if (kr.RibbonItemIndex == -1)
                        {
                            ribbonProperties.ScreenTip = kr.RibbonScreenTipContent + " " + ribbonProperties.InitialScreenTip;
                            ribbonProperties.ShortcutText = " " + kr.RibbonScreenTipContent;
                        }
                        else
                        {
                            var content = (RibbonControlModel[])ribbonProperties.Content;
                            content[kr.RibbonItemIndex].ScreenTip = kr.RibbonScreenTipContent + " " + content[kr.RibbonItemIndex].InitialScreenTip;
                            content[kr.RibbonItemIndex].ShortcutText = " " + kr.RibbonScreenTipContent;
                        }
                    }
                }
                Globals.Ribbon.ribbon?.Invalidate();
                Globals.ThisAddIn.SetKeyBindingStatus = true;
            }
            catch (Exception e)
            {
                DedicatedFunctions.ShowErrorMessage("خطایی در تنظیم کلید های میانبر به وجود آمد" + "\n" + "پیغام خطا" + "\n" + e.Message,
                    (int)ErrorCodes.SetKeyboardShortcut, StringConstant.SupportEmail);
            }
            //Globals.ThisAddIn.Application.KeyBindings.Key
            //Globals.ThisAddIn.Application.KeyBindings.Context
            //Globals.ThisAddIn.Application.KeyBindings.ClearAll
            //Globals.ThisAddIn.Application.KeyBindings[].
        }
        internal static void unsetKeyBinding(List<KeyboardRelationModel> keyboardRelations)
        {
            try
            {
                KeyBindings kbs = Globals.ThisAddIn.Application.KeyBindings;
                for (int i = Globals.ThisAddIn.Application.KeyBindings.Count; i >= 1; i--)
                {
                    for (int x = keyboardRelations.Count - 1; x >= 0; x--)
                    {
                        if (kbs[i].KeyCode == keyboardRelations[x].KeyboardShortcut)
                        {
                            try
                            {
                                if (keyboardRelations[x].RibbonControlID != null)
                                {
                                    var ribbonProperties = Globals.Ribbon.ribbonComponents.Where(p => p.Id == keyboardRelations[x].RibbonControlID).First();

                                    if (ribbonProperties != null)
                                    {
                                        if (keyboardRelations[x].RibbonItemIndex == -1)
                                        {
                                            ribbonProperties.ScreenTip = ribbonProperties.InitialScreenTip;
                                            ribbonProperties.ShortcutText = "";
                                        }
                                        else
                                        {
                                            var content = (RibbonControlModel[])ribbonProperties.Content;
                                            content[keyboardRelations[x].RibbonItemIndex].ScreenTip = content[keyboardRelations[x].RibbonItemIndex].InitialScreenTip;
                                            content[keyboardRelations[x].RibbonItemIndex].ShortcutText = "";
                                        }
                                    }
                                }
                            }
                            catch (Exception)
                            {
                            }

                            keyboardRelations.Remove(keyboardRelations[x]);
                            kbs[i].Clear();
                            break;
                        }

                        //if(kr.RibbonControlID != null)
                        //{
                        //var ribbonProperties = Globals.Ribbon.ribbonComponents.Where(p => p.Id == kr.RibbonControlID).First();

                        //if(kr.RibbonItemIndex == -1)
                        //{
                        //	ribbonProperties.SuperTip = ribbonProperties.SuperTip + kr.RibbonScreenTipContent;
                        //}
                        //else
                        //{
                        //	var content = (RibbonProperties[])ribbonProperties.Content;
                        //	content[kr.RibbonItemIndex].SuperTip = content[kr.RibbonItemIndex].SuperTip + kr.RibbonScreenTipContent;
                        //}
                        //}

                    }
                }
                Globals.ThisAddIn.SetKeyBindingStatus = false;
            }
            catch (Exception e)
            {
                DedicatedFunctions.ShowErrorMessage("خطایی در حذف کلید های میانبر به وجود آمد" + "\n" + "پیغام خطا" + "\n" + e.Message,
                    (int)ErrorCodes.SetKeyboardShortcut, StringConstant.SupportEmail);
            }
        }

        #endregion



        internal static void InitializeRibbon(string ribbonTitle)
        {
            Ribbon.setTabProperties(ribbonTitle, true);
            RibbonControlsVisibility(true);
        }
        internal static void RibbonControlsVisibility(bool enable)
        {
            try
            {
                //Custom Ribbon Controls
                List<RibbonControlModel> ribbonProperties = Globals.Ribbon.ribbonComponents.Where(p => p.RibbonControlType != RibbonControlModel.RibbonControlTypes.Tab && p.RibbonControlType != RibbonControlModel.RibbonControlTypes.Group && p.RibbonControlType != RibbonControlModel.RibbonControlTypes.BuiltIn).ToList();

                List<string> exceptEnable = new List<string>()
                {
                    //RibbonControlNames.btnDocumentsManagerBackstage,
                    //RibbonControlNames.btnDocumentsManager,
                    //RibbonControlNames.btnLogin,
                };

                //List<string> exceptVisibility = new List<string>()
                //{ };
                //List<string> defaultHiddenVisibility = new List<string>()
                //{
                //    RibbonControlNames.btnChatBoxNetworking
                //};


                // دسترسی باز برای افراد مشترک سند
                Document doc;
                try
                {
                    doc = Globals.ThisAddIn?.Application?.ActiveDocument;
                    if (doc != null)
                    {
                        string value = DedicatedFunctions.getStaticVariableValue(doc, VariableDocumentIDs._variable_document_SharedDocument.ToString());

                        if (value == "1")
                        {
                            exceptEnable.Add(RibbonControlNames.btnChatBoxNetworking);
                            //exceptVisibility.Add(RibbonControlNames.btnChatBoxNetworking);
                            enable = false;
                        }
                    }
                }
                catch (Exception)
                {
                }


                foreach (RibbonControlModel rp in ribbonProperties)
                {
                    if (!exceptEnable.Contains(rp.Id))
                        rp.Enable = enable;
                    else
                        rp.Enable = true;

                    if (rp.Id == RibbonControlNames.galleryTools)
                    {
                        object content = rp.Content;
                        if (content != null)
                        {
                            doc = Globals.ThisAddIn?.Application?.ActiveDocument;
                            if (doc != null)
                            {
                                ContentControl[] ccs = DedicatedFunctions.getContentControls(doc, ContentControlNames._field_Dedication_Fa.ToString());

                                if (ccs != null && ccs.Length > 0)
                                {
                                    ((RibbonControlModel[])content)[1].Label = "تغییر متن تقدیم";
                                    ((RibbonControlModel[])content)[1].ScreenTip = "تغییر متن تقدیم" + ((RibbonControlModel[])content)[1].ShortcutText;
                                }
                                else
                                {
                                    ((RibbonControlModel[])content)[1].Label = "درج متن تقدیم";
                                    ((RibbonControlModel[])content)[1].ScreenTip = "درج متن تقدیم" + ((RibbonControlModel[])content)[1].ShortcutText;
                                }
                            }
                        }
                    }
                    //if (exceptVisibility.Contains(rp.Id))
                    //    rp.Visible = true;
                    //else if (defaultHiddenVisibility.Contains(rp.Id))
                    //    rp.Visible = false;
                }

                Globals.Ribbon.ribbon?.Invalidate();
            }
            catch (Exception)
            {
            }
        }
    }



    internal struct RibbonControlNames
    {
        internal const string tabJameNeger = "tabJameNeger";

        internal const string btnLogin = "btnLogin";

        internal const string galleryDocumentsManager = "galleryDocumentsManager";

        internal const string btnChatBoxNetworking = "btnChatBoxNetworking";

        internal const string grpDocumentManager = "grpDocumentManager";

        //internal const string btnDocumentsManagerBackstage = "btnDocumentsManagerBackstage";
        internal const string btnChangeContents = "btnChangeContents";
        internal const string btnAddRemovePages = "btnAddRemovePages";

        internal const string grpFormat = "grpFormat";
        internal const string btnTextFormating = "btnTextFormating";
        internal const string btnHeading2 = "btnHeading2";
        internal const string btnHeading3 = "btnHeading3";
        internal const string btnHeading4 = "btnHeading4";
        internal const string btnHeading5 = "btnHeading5";


        internal const string grpCaptionsAndRefer = "grpCaptionsAndRefer";
        internal const string btnRefer = "btnRefer";
        internal const string btnInsertWriteCaptionForShape = "btnInsertWriteCaptionForShape";
        internal const string btnInsertWriteCaptionForTable = "btnInsertWriteCaptionForTable";
        internal const string btnInsertCaptionForFormula = "btnInsertCaptionForFormula";
        internal const string grpCaptionsAndRefer__btn = "grpCaptionsAndRefer__btn";
        internal const string grpFormat__btn = "grpFormat__btn";


        internal const string grpFootnotes = "grpFootnotes";
        internal const string galleryInsertFootnote = "galleryInsertFootnote";
        internal const string grpFootnotes__btn = "grpFootnotes__btn";


        internal const string grpVirastar = "grpVirastar";
        internal const string gallerySimpleVirastar = "gallerySimpleVirastar";
        internal const string btnStandardCorrection = "btnStandardCorrection";

        internal const string galleryConvertNumber = "galleryConvertNumber";
        internal const string galleryPoemMode = "galleryPoemMode";
        internal const string grpVirastar__btn = "grpVirastar__btn";

        internal const string grpCitationManager = "grpCitationManager";
        internal const string btnInsertCitation = "btnInsertCitation";
        internal const string galleryInsertImportSources = "galleryInsertImportSources";
        internal const string galleryCitationType = "galleryCitationType";
        internal const string btnManageResources = "btnManageResources";
        internal const string btnGoToBibliography = "btnGoToBibliography";
        //internal const string grpCitationManager__btn = "grpCitationManager__btn";

        internal const string grpOtherOptions = "grpOtherOptions";
        internal const string btnUpdateDocument = "btnUpdateDocument";
        internal const string btnUploadDocument = "btnUploadDocument";
        internal const string galleryExport = "galleryExport";
        internal const string grpOtherOptions__btn = "grpOtherOptions__btn";

        internal const string grpTools = "grpTools";
        internal const string galleryBooks = "galleryBooks";
        internal const string galleryTools = "galleryTools";
        internal const string btnIntroduction = "btnIntroduction";
        internal const string galleryAbout = "galleryAbout";


        //BuiltIns

        //Backstage
        //internal const string builtInTabPrint = "TabPrint";
        internal const string builtInTabSave = "TabSave";
        internal const string builtInTabPublish = "TabPublish";
        internal const string builtInTabShare = "TabShare";
        internal const string builtInTabHelp = "TabHelp";

        //Backstage_TabInfo
        internal const string builtInGroupPermissions = "GroupPermissions";
        internal const string builtInGroupPrepareForSharing = "GroupPrepareForSharing";
        internal const string builtInGroupVersions = "GroupVersions";

        //All Disable Protect and Password Controls and Groups
        internal const string builtInGroupProtectedView = "GroupProtectedView";
        internal const string builtInFilePermissionRestrictAs = "FilePermissionRestrictAs";
        internal const string builtInFilePermission = "FilePermission";
        internal const string builtInFilePermissionView = "FilePermissionView";
        internal const string builtInFilePermissionRestrictMenu = "FilePermissionRestrictMenu";
        internal const string builtInFilePermissionUnrestricted = "FilePermissionUnrestricted";
        internal const string builtInFilePermissionDoNotDistribute = "FilePermissionDoNotDistribute";
        internal const string builtInProtectOrUnprotectDocument = "ProtectOrUnprotectDocument";
        internal const string builtInProtectDocument = "ProtectDocument";
        internal const string builtInReviewProtectDocumentMenu = "ReviewProtectDocumentMenu";
        internal const string builtInGroupProtect = "GroupProtect";
        internal const string builtInReviewRestrictFormatting = "ReviewRestrictFormatting";

        //Developer Tab
        internal const string builtInGroupCode = "GroupCode";
        internal const string builtInVisualBasic = "VisualBasic";
        internal const string builtInMacroPlay = "MacroPlay";
        internal const string builtInMacroRecordOrStop = "MacroRecordOrStop";
        internal const string builtInMacroRecorderPause = "MacroRecorderPause";
        internal const string builtInMacroSecurity = "MacroSecurity";
        internal const string builtInControlProperties = "ControlProperties";
        internal const string builtInDesignMode = "DesignMode";
    }


    public sealed class PictureDispConverter
    {
        [DllImport("OleAut32.dll", EntryPoint = "OleCreatePictureIndirect", ExactSpelling = true, PreserveSig = false)]
        private static extern stdole.IPictureDisp

            OleCreatePictureIndirect([MarshalAs(UnmanagedType.AsAny)] object picdesc, ref Guid iid, [MarshalAs(UnmanagedType.Bool)] bool fOwn);

        static Guid iPictureDispGuid = typeof(stdole.IPictureDisp).GUID;



        private static class PICTDESC

        {
            //Picture Types
            public const short PICTYPE_UNINITIALIZED = -1;
            public const short PICTYPE_NONE = 0;
            public const short PICTYPE_BITMAP = 1;
            public const short PICTYPE_METAFILE = 2;
            public const short PICTYPE_ICON = 3;
            public const short PICTYPE_ENHMETAFILE = 4;



            [StructLayout(LayoutKind.Sequential)]
            public class Icon
            {
                public int cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC.Icon));
                public int picType = PICTDESC.PICTYPE_ICON;
                public IntPtr hicon = IntPtr.Zero;
                public int unused1;
                public int unused2;

                public Icon(System.Drawing.Icon icon)
                {
                    this.hicon = icon.ToBitmap().GetHicon();
                }
            }

            [StructLayout(LayoutKind.Sequential)]
            public class Bitmap
            {
                public int cbSizeOfStruct = Marshal.SizeOf(typeof(PICTDESC.Bitmap));
                public int picType = PICTDESC.PICTYPE_BITMAP;
                public IntPtr hbitmap = IntPtr.Zero;
                public IntPtr hpal = IntPtr.Zero;
                public int unused;
                public Bitmap(System.Drawing.Bitmap bitmap)
                {
                    this.hbitmap = bitmap.GetHbitmap();
                }
            }
        }
        public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Icon icon)
        {
            PICTDESC.Icon pictIcon = new PICTDESC.Icon(icon);
            return OleCreatePictureIndirect(pictIcon, ref iPictureDispGuid, true);
        }
        public static stdole.IPictureDisp ToIPictureDisp(System.Drawing.Bitmap bmp)
        {
            PICTDESC.Bitmap pictBmp = new PICTDESC.Bitmap(bmp);
            return OleCreatePictureIndirect(pictBmp, ref iPictureDispGuid, true);
        }
    }
}
