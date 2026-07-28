using System;
namespace ShivaNegar.Interfaces
{
    interface ICurrentViewChanged
    {
        Action CurrentViewChanged { get; set; }
    }
}
