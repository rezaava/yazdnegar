using System;
namespace YazdNegar.Interfaces
{
    internal interface ITransitionCommand
    {
        Action TransitionMovePreviousCommand { get; set; }
        Action TransitionMoveNextCommand { get; set; }

    }
}
