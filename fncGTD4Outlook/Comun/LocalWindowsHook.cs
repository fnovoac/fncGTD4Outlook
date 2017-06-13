using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace fncGTD4Outlook.Comun
{
    public class InterceptKeys
    {
        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;

        //Declare the mouse hook constant.
        //For other hook types, you can obtain these values from Winuser.h in the Microsoft SDK.
        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;

        public static void SetHook()
        {
            //_hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());

            //ver:
            //https://stackoverflow.com/questions/25501823/how-to-get-the-current-thread-id-without-appdomain-getcurrentthreadid-so-that
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, GetCurrentThreadId());
        }

        public static void ReleaseHook()
        {
            UnhookWindowsHookEx(_hookID);
        }

        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            int PreviousStateBit = 31;
            bool KeyWasAlreadyPressed = false;

            Int64 bitmask = (Int64)Math.Pow(2, (PreviousStateBit - 1));

            try
            {
                if (nCode < 0)
                {
                    return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
                }
                else
                {
                    if (nCode == HC_ACTION)
                    {
                        Keys keyData = (Keys)wParam;
                        KeyWasAlreadyPressed = ((Int64)lParam & bitmask) > 0;

                        //if (Functions.IsKeyDown(Keys.ControlKey) && keyData == Keys.D1 && KeyWasAlreadyPressed == false)
                        //{
                        //    object missing = System.Reflection.Missing.Value;
                        //    Excel.Workbook exBook = Globals.ThisAddIn.Application.ActiveWorkbook;
                        //    Excel.Worksheet exSheet = (Excel.Worksheet)exBook.ActiveSheet;
                        //    exSheet.get_Range("A4", missing).Value2 = exSheet.get_Range("A4", missing).Value2 + "A";
                        //}

                        if (Functions.IsKeyDown(Keys.ControlKey) && Functions.IsKeyDown(Keys.ShiftKey) &&  keyData == Keys.Enter && KeyWasAlreadyPressed == false)
                        {
                            Utils.MostrarDelegarForm();
                        }
                    }
                    return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
            LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
            IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();


    }

    public class Functions
    {
        public static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }

        [DllImport("user32.dll")]
        static extern short GetKeyState(int nVirtKey);

    }

}
