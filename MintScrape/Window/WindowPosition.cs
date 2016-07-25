using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace MintScrape.Window {
    public static class Position {
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        private static extern IntPtr FindWindowByCaption(IntPtr zeroOnly, string lpWindowName);

        public static void ToTop() {
            var originalTitle = Console.Title;
            var uniqueTitle = Guid.NewGuid().ToString();
            Console.Title = uniqueTitle;
            Thread.Sleep(50);
            var handle = FindWindowByCaption(IntPtr.Zero, uniqueTitle);

            if (handle == IntPtr.Zero) {
                Console.WriteLine("Oops, cant find main window.");
                return;
            }
            Console.Title = originalTitle;
            SetForegroundWindow(handle);
        }
    }
}