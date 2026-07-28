using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ShivaNegar
{
    class KeyboardHooking
    {
        #region Events
        private static int MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            //if(nCode == HC_ACTION)
            //{
            //	var mouseHookStruct = (MouseHookStructEx)Marshal.PtrToStructure(lParam , typeof(MouseHookStructEx));
            //
            //	// handle mouse message here
            //	WindowMessages message = (WindowMessages)wParam;
            //	if(message == WindowMessages.WM_LBUTTONDOWN ||
            //		message == WindowMessages.WM_LBUTTONUP)
            //	{
            //		Debug.WriteLine("{0} event detected at position {1} - {2}" , message , mouseHookStruct.pt.X , mouseHookStruct.pt.Y);
            //	}
            //}
            return (int)CallNextHookEx(_mouseHookID, nCode, wParam, lParam);
        }

        //Note that the custom code goes in this method the rest of the class stays the same.
        //It will trap if BOTH keys are pressed down.
        private static int KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == HC_ACTION)
            {
                Keys keyData = (Keys)wParam;
                KeyEventArgs args = new KeyEventArgs(keyData);


                //Keyboard Shortcut
                // CTRL + SHIFT + 7
                if ((BindingFunctions.IsKeyDown(Keys.ControlKey) == true)
                    && (BindingFunctions.IsKeyDown(Keys.ShiftKey) == true)
                    && (BindingFunctions.IsKeyDown(keyData) == true) && (keyData == Keys.D7))
                {
                    //Debug.WriteLine("Tested CTRL + SHIFT + 7");
                }

                //detect KeyDown and KeyUp
                bool isKeyDown = ((ulong)lParam & 0x40000000) == 0;
                if (isKeyDown)
                {
                    if (!isBackKeyDowned)
                    {
                        onKeyDown(args);
                        onKeyPress(args);
                        if (args.KeyData == Keys.Back)
                            isBackKeyDowned = true;
                    }
                }
                else
                {
                    bool isLastKeyUp = ((ulong)lParam & 0x80000000) == 0x80000000;
                    if (isLastKeyUp)
                    {
                        isBackKeyDowned = false;
                        onKeyUp(args);
                    }
                    else
                        onKeyPress(args);

                }
            }
            return (int)CallNextHookEx(_keyboardHookID, nCode, wParam, lParam);
        }

        private static void onKeyDown(KeyEventArgs args)
        {
            //Debug.WriteLine("KeyDown " + args.KeyData.ToString());
        }
        private static void onKeyPress(KeyEventArgs args)
        {
            //Debug.WriteLine("KeyPress");
        }
        private static void onKeyUp(KeyEventArgs args)
        {
            //Debug.WriteLine("KeyUp " + args.KeyData.ToString());
        }

        #endregion

        #region DLL Import
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod,
            uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);
        #endregion

        #region Constant Variables and Enums
        private const int HC_ACTION = 0;
        //private const int WH_KEYBOARD_LL = 13; // keyboard

        internal enum HookType : int
        {
            WH_KEYBOARD = 2,
            WH_MOUSE = 7
        }

        internal enum WindowMessages : uint
        {
            WM_KEYDOWN = 0x0100,
            WM_KEYFIRST = 0x0100,
            WM_KEYLAST = 0x0108,
            WM_KEYUP = 0x0101,
            WM_LBUTTONDBLCLK = 0x0203,
            WM_LBUTTONDOWN = 0x0201,
            WM_LBUTTONUP = 0x0202,
            WM_MBUTTONDBLCLK = 0x0209,
            WM_MBUTTONDOWN = 0x0207,
            WM_MBUTTONUP = 0x0208,
            WM_MOUSEACTIVATE = 0x0021,
            WM_MOUSEFIRST = 0x0200,
            WM_MOUSEHOVER = 0x02A1,
            WM_MOUSELAST = 0x020D,
            WM_MOUSELEAVE = 0x02A3,
            WM_MOUSEMOVE = 0x0200,
            WM_MOUSEWHEEL = 0x020A,
            WM_MOUSEHWHEEL = 0x020E,
            WM_RBUTTONDBLCLK = 0x0206,
            WM_RBUTTONDOWN = 0x0204,
            WM_RBUTTONUP = 0x0205,
            WM_SYSDEADCHAR = 0x0107,
            WM_SYSKEYDOWN = 0x0104,
            WM_SYSKEYUP = 0x0105
        }
        [StructLayout(LayoutKind.Sequential)]
        internal struct MouseHookStructEx
        {
            internal Point pt;
            internal IntPtr hwnd;
            internal uint wHitTestCode;
            internal IntPtr dwExtraInfo;
            internal int MouseData;
        }

        //typedef struct tagKBDLLHOOKSTRUCT
        //{
        //	DWORD vkCode;
        //	DWORD scanCode;
        //	DWORD flags;
        //	DWORD time;
        //	ULONG_PTR dwExtraInfo;
        //}
        #endregion

        #region Variables
        static bool isBackKeyDowned = false;


        internal delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _keyboardProc = KeyboardHookCallback;
        private static LowLevelKeyboardProc _mouseProc = MouseHookCallback;
        private static IntPtr _keyboardHookID = IntPtr.Zero;
        private static IntPtr _mouseHookID = IntPtr.Zero;
        #endregion

        #region Functions
        internal static void SetHook()
        {
            // Ignore this compiler warning, as SetWindowsHookEx doesn't work with ManagedThreadId
#pragma warning disable 618
            _keyboardHookID = SetWindowsHookEx((int)HookType.WH_KEYBOARD, _keyboardProc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
            _mouseHookID = SetWindowsHookEx((int)HookType.WH_MOUSE, _mouseProc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());
#pragma warning restore 618

        }
        internal static void ReleaseHook()
        {
            UnhookWindowsHookEx(_keyboardHookID);
            UnhookWindowsHookEx(_mouseHookID);
        }
        #endregion
    }

    internal class BindingFunctions
    {
        [DllImport("user32.dll")]
        static extern short GetKeyState(int nVirtKey);

        internal static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }

    }
}
