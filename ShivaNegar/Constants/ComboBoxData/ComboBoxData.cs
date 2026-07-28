namespace ShivaNegar.Constants.ComboBoxData
{
    struct ComboBoxData
    {
        internal const string nothingFa = "ندارد";
        internal const string nothingEn = "nothing";
        internal const string otherFa = "موارد دیگر";
        internal const string otherEn = "other";

        //internal static readonly string[] InTheNameOfAllah = { "a" , "b" , "c" , "d" , "e" , "f" , "g" , "h" , "i" , "j" , "k" , "l" , "m" , "n" , "o" , "p" , "q" , "r" , "s" , "t" , "u" , "v" , "w" , "x" , "y" , "z" , "0" , "1" , "2" , "3" , "4" , "5" , "6" , "7" , "8" , "9" , "\\" , "[" , "]" , "{" , "}" , "=" , "-" , "_" , "+" , "A" , "B" , "C" , "D" , "E" , "F" , "G" , "H" , "I" , "J" , "K" , "L" , "M" , "N" , "O" , "P" , "Q" , "R" , "S" , "T" , "U" , "V" , "W" , "X" , "Y" , "Z" , };
        //internal static readonly string[] InTheNameOfAllah = { "a" , "b" , "c" , "d" , "e" , "f" , "g" , "h" , "i" , "j" , "k" , "l" , "m" , "n" , "o" , "p" , "q" , "r" , "s" , "t" , "u" , "v" , "w" , "x" , "y" , "z" , "\\" , "[" , "]" , "{" , "}" , "=" , "-" , "_" , "+" , "A" , "B" , "C" , "D" , "E" , "F" , "G" , "H" , "I" , "J" , "K" , "L" , "M" , "N" , "O" , "P" , "Q" , "R" , "S" , "T" , "U" , "V" , "W" , "X" , "Y" , "Z" , };
        //internal static readonly string[] InTheNameOfAllah = { "R" , "y" , "\\" , "[" , "{" , "=" , "Q" , "G" , "[" , "N" , "i" , "k" , "_" , "+" , "B" , "F" , "L" , "M" , "X" , "P" , "Z" , "U" , "Y" , "S" , "J" , "E" , "D" , "f" };
        internal static readonly string[] InTheNameOfAllah = { "R", "y", "Q", "G", "N", "i", "k", "B", "F", "L", "M", "X", "P", "Z", "U", "Y", "S", "J", "E", "D", "f" };
    }

    struct ComboBoxDataFaculty
    {
        //internal static readonly string[] Faculty_Fa =
        //{
        //    "پردیس علوم انسانی و اجتماعی",
        //    "پردیس علوم",
        //    "پردیس فنی و مهندسی",
        //    "پردیس آزادی",
        //    "پردیس مهریز",
        //    "دانشکده هنر و معماری",
        //    "دانشکده منابع طبیعی و کویرشناسی",
        //    "موارد دیگر"
        //};
        //internal static readonly string[] Faculty_En =
        //{
        //    "Faculty of Humanities & Social Sciences",
        //    "Faculty of Science",
        //    "Faculty of Engineering",
        //    "Faculty of Azadi",
        //    "Faculty of Mehriz",
        //    "School of Arts & Architecture",
        //    "School of Natural Resources & Desert Studies",
        //    "other"
        //};
    }

    struct ComboBoxDataAcademicDegree
    {
        internal static readonly string[] AcademicDegree_Fa =
        {
            AcademicDegreeValues.AcademicDegree_AssociateOfScienceFa,
            AcademicDegreeValues.AcademicDegree_BachelorOfScienceFa,
            AcademicDegreeValues.AcademicDegree_PartTimeBachelorOfScienceFa,
            AcademicDegreeValues.AcademicDegree_MasterOfScienceFa,
            AcademicDegreeValues.AcademicDegree_DoctoralFa,
        };
        internal static readonly string[] AcademicDegree_En =
        {
            AcademicDegreeValues.AcademicDegree_AssociateOfScienceEn,
            AcademicDegreeValues.AcademicDegree_BachelorOfScienceEn,
            AcademicDegreeValues.AcademicDegree_PartTimeBachelorOfScienceEn,
            AcademicDegreeValues.AcademicDegree_MasterOfScienceEn,
            AcademicDegreeValues.AcademicDegree_DoctoralEn,
        };

        internal static readonly string[] AcademicDegree_Project_Fa =
        {
            AcademicDegreeValues.AcademicDegree_AssociateOfScienceFa,
            AcademicDegreeValues.AcademicDegree_BachelorOfScienceFa,
        };
        internal static readonly string[] AcademicDegree_Project_En =
        {
            AcademicDegreeValues.AcademicDegree_AssociateOfScienceEn,
            AcademicDegreeValues.AcademicDegree_BachelorOfScienceEn,
        };
        internal static readonly string[] AcademicDegree_Thesis_Fa =
        {
            AcademicDegreeValues.AcademicDegree_MasterOfScienceFa,
        };
        internal static readonly string[] AcademicDegree_Thesis_En =
        {
            AcademicDegreeValues.AcademicDegree_MasterOfScienceEn,
        };
        internal static readonly string[] AcademicDegree_Dissertation_Fa =
        {
            AcademicDegreeValues.AcademicDegree_DoctoralFa,
        };
        internal static readonly string[] AcademicDegree_Dissertation_En =
        {
            AcademicDegreeValues.AcademicDegree_DoctoralEn,
        };

    }

    //struct ComboBoxDataGroup
    //{
    //	internal static readonly string[] TestGroup_Fa =
    //	{
    //		//"گروه اول",
    //		//"گروه دوم",
    //		//"گروه سوم",
    //		ComboBoxData.otherFa
    //	};
    //	internal static readonly string[] TestGroup_En =
    //	{
    //		//"group one",
    //		//"group two",
    //		//"group three",
    //		ComboBoxData.otherEn
    //	};
    //}

}
