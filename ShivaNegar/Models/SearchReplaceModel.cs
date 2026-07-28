namespace ShivaNegar.Models
{
    internal class SearchReplaceModel
    {
        public string Search { get; set; }
        public string Replace { get; set; }
        public string Wildcard { get; set; }
        public string MatchCase { get; set; }
        public string MatchWholeWord { get; set; }
        public string MatchKashida { get; set; }
        public string MatchDiacritics { get; set; }
        public string MatchAlefHamza { get; set; }

        public string CounterUp { get; set; }
    }
}
