using System;
namespace ShivaNegar.Interfaces
{
    internal interface ITransitionCommand
    {
        Action TransitionMovePreviousCommand { get; set; }
        Action TransitionMoveNextCommand { get; set; }

    }
}
