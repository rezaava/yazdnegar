using System;
namespace ShivaNegar.Interfaces
{
    interface IStatusFormRequest
    {
        Action CloseFormRequest { get; set; }
        Action AlwaysOnTopEnableRequest { get; set; }
        Action AlwaysOnTopDisableRequest { get; set; }

        Action MinimizeStateFormRequest { get; set; }
        Action NormalStateFormRequest { get; set; }
        Action MaximizeStateFormRequest { get; set; }
    }
}
