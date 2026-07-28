using System.Runtime.InteropServices;

namespace ShivaNegar.Interfaces
{

    [ComVisible(true)]
    public interface IKeyboardShortcuts
    {
        void documentsManager();
        void changeContent();
        void addRemovePages();

        void captionSettings();
        void virastarSettings();
        void footnoteSettings();
        void listsManagerSettings();
        void citationSettings();

        void setNormalStyle();
        void setHeadingStyle(int index);

        void InsertCaption(int index);

        void ShowCrossReferenceMenu();
        void insertCitationMenu();

        void insertFootnote(bool isPersian);

        void updateDocument();
        void uploadDocument();

        void insertHalfSpace();
        void halfSpaceCorrection();
        void neshanehGozariCorrection();
        void standardCorrection();
        void spellingCorrection();

        void convertNumber(bool toPersian);
        void poemMode(int column);

        void sourceManagementDialog();
        void insertSources();
        void importSourcesDialog();
        void goToBibliography();


        void exportToWord();
        void exportToPDF();
        void exportCDMenu();
        void exportAsGrayscale();
        void exportToIdentification();

        void chatBoxNetworking();

        void createProposal();
        void besmellahPageForm();
        void insertNahjBalaghe();
        void insertQuran();
        void insertDedicate();

        void defenseAnnouncements();
    }
}
