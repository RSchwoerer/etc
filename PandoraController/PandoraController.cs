using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;

/// <remarks>
/// http://www.daveamenta.com/2010-06/pandora-one-media-keys-enable-them/
/// http://www.daveamenta.com/download/PandoraController.cs.txt
/// </remarks>
namespace PandoraOneMediaKeys
{
    class PandoraController
    {
        #region Win32 Platform Invoke
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsIconic(IntPtr hWnd);

        private const uint WM_LBUTTONDOWN = 0x0201;
        private const uint WM_LBUTTONUP = 0x0202;
        private const int SW_RESTORE = 9;

        private static IntPtr MakeLParam(int LoWord, int HiWord)
        {
            return (IntPtr)((HiWord << 16) | (LoWord & 0xffff));
        }

        #endregion

        public static bool Except = false;

        private static void Click(int x, int y)
        {
            IntPtr hWnd = FindWindow("ApolloRuntimeContentWindow", "Pandora");
            if (hWnd == IntPtr.Zero)
            {
                if (Except)
                {
                    throw new InvalidOperationException("Pandora One is not running");
                }
                else
                {
                    return;
                }
            }

            if (IsIconic(hWnd))
            {
                ShowWindow(hWnd, SW_RESTORE);
                int cTimeout = 0;
                do
                {
                    Thread.Sleep(100);
                    cTimeout++;
                    if (cTimeout > 10) break; // 1sec max
                } while (IsIconic(hWnd));
            }

            IntPtr p = MakeLParam(x, y);
            SendMessage(hWnd, WM_LBUTTONDOWN, IntPtr.Zero, p);
            SendMessage(hWnd, WM_LBUTTONUP, IntPtr.Zero, p);
        }

        public static void PlyaPause()
        {
            Click(142, 349);
        }

        public static void Next()
        {
            Click(181, 349);
        }

        public static void Like()
        {
            Click(98, 349);
        }

        public static void Dislike()
        {
            Click(26, 349);
        }
    }
}