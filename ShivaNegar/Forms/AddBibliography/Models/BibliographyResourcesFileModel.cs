namespace ShivaNegar.Forms.AddBibliography.Models
{
    public class BibliographyResourcesFileModel
    {
        private int id;
        private string fileName;
        private string fileExtension;
        private string path;
        private string content;

        public BibliographyResourcesFileModel(int id, string fileName, string fileExtension, string path, string content)
        {
            this.Id = id;
            this.FileName = fileName;
            this.FileExtension = fileExtension;
            this.Path = path;
            this.Content = content;
        }

        public int Id { get => id; set => id = value; }
        public string FileName { get => fileName; set => fileName = value; }
        public string FileExtension { get => fileExtension; set => fileExtension = value; }
        public string Path { get => path; set => path = value; }
        public string Content { get => content; set => content = value; }
    }
}
