using System;
using System.Drawing;
using ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Utilities;

namespace ShivaNegar.Forms.ShivaNegarManager.DocumentManager.Models
{
    public class FileModel : ViewModelBase
    {
        private Icon _image;
        private int _id;
        private string _versionAddin;
        private string _name;
        private string _documentURL;
        private string _type;
        private string _typeHolder;
        private DateTime createdAt;
        private DateTime updatedAt;
        private DateTime updatedFile;
        private DateTime updatedConfig;

        private string _titleSearched;

        private string _universityName;
        private string _groupOFUniversity;
        private string _departmentOfUniversity;
        private string _fieldOfStudy;
        private string _title;
        private string _supervisor;
        private string _author;

        private string _abstract;

        private bool _networkedDocument;
        private bool _sharedDocuemnt;


        //private string _path;
        //public string Path
        //{
        //	get => _path;
        //	set => RaisePropertyChanged(ref _path , value);
        //}

        public bool NetworkedDocument
        {
            get => _networkedDocument;
            set
            {
                if (value && !SharedDocuemnt)
                    Type = TypeHolder + " (اشتراک گذاری شده)";
                else if (!SharedDocuemnt)
                    Type = TypeHolder;
                else
                    Type = TypeHolder + " (سند اشتراکی دانشجو)";

                RaisePropertyChanged(ref _networkedDocument, value);
            }
        }

        public bool SharedDocuemnt
        {
            get => _sharedDocuemnt;
            set
            {
                if (NetworkedDocument && !value)
                    Type = TypeHolder + " (اشتراک گذاری شده)";
                else if (!value)
                    Type = TypeHolder;
                else
                    Type = TypeHolder + " (سند اشتراکی دانشجو)";

                RaisePropertyChanged(ref _sharedDocuemnt, value);
            }
        }

        public string Abstract
        {
            get => _abstract;
            set => RaisePropertyChanged(ref _abstract, value);
        }

        public Icon Image
        {
            get => _image;
            set => RaisePropertyChanged(ref _image, value);
        }
        public string VersionAddin
        {
            get => _versionAddin;
            set => RaisePropertyChanged(ref _versionAddin, value);
        }
        public int ID
        {
            get => _id;
            set => RaisePropertyChanged(ref _id, value);
        }
        public string Name
        {
            get => _name;
            set => RaisePropertyChanged(ref _name, value);
        }
        public string DocumentURL
        {
            get => _documentURL;
            set => RaisePropertyChanged(ref _documentURL, value);
        }
        public string Type
        {
            get => _type;
            set => RaisePropertyChanged(ref _type, value);
        }
        public string TypeHolder
        {
            get => _typeHolder;
            set => RaisePropertyChanged(ref _typeHolder, value);
        }
        public DateTime CreatedAt
        {
            get => createdAt;
            set => createdAt = value;
        }
        public DateTime UpdatedAt
        {
            get => updatedAt;
            set => updatedAt = value;
        }

        public DateTime UpdatedFile
        {
            get => updatedFile;
            set => updatedFile = value;
        }

        public DateTime UpdatedConfig
        {
            get => updatedConfig;
            set => updatedConfig = value;
        }

        public string TitleSearched
        {
            get => _titleSearched;
            set => RaisePropertyChanged(ref _titleSearched, value);
        }

        public string UniversityName
        {
            get => _universityName;
            set => RaisePropertyChanged(ref _universityName, value);
        }
        public string GroupOFUniversity
        {
            get => _groupOFUniversity;
            set => RaisePropertyChanged(ref _groupOFUniversity, value);
        }
        public string DepartmentOfUniversity
        {
            get => _departmentOfUniversity;
            set => RaisePropertyChanged(ref _departmentOfUniversity, value);
        }
        public string FieldOfStudy
        {
            get => _fieldOfStudy;
            set => RaisePropertyChanged(ref _fieldOfStudy, value);
        }
        public string Title
        {
            get => _title;
            set => RaisePropertyChanged(ref _title, value);
        }
        public string Supervisor
        {
            get => _supervisor;
            set => RaisePropertyChanged(ref _supervisor, value);
        }
        public string Author
        {
            get => _author;
            set => RaisePropertyChanged(ref _author, value);
        }
    }
}
