using System;
namespace YazdNegar.Interfaces
{
    interface ICurrentViewChanged
    {
        Action CurrentViewChanged { get; set; }
    }
}
