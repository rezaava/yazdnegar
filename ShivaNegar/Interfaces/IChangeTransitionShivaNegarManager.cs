using System;
namespace ShivaNegar.Interfaces
{
    interface IChangeTransitionDocumentManager
    {
        Action TransitionDocumentManagerRequest { get; set; }
    }
    interface IChangeTransitionCreateDocument
    {
        string Token { get; set; }
        string DocumentName { get; set; }
        Action TransitionCreateDocumentRequest { get; set; }
    }
}
