using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace officePresentationSplit.Methods
{
    public static class NativeMethods
    {
        internal const int Srccopy = 0x00CC0020; // BitBlt dwRop parameter

        [DllImport("User32", CharSet = CharSet.Auto, ExactSpelling = true)]
        internal static extern IntPtr SetParent(IntPtr hWndChild, IntPtr hWndParent);

        [DllImport("user32.dll")]
        internal static extern bool MoveWindow(IntPtr hWnd, int x, int y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        internal static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("User32.dll")]
        internal static extern int SetForegroundWindow(int hWnd);

        [DllImport("user32.dll")]
        internal static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);

        [DllImport("user32.dll")]
        internal static extern bool RegisterHotKey(int hwnd, int id, uint fsModifiers, uint vk);

        [DllImport("user32.dll")]
        internal static extern bool UnregisterHotKey(int hwnd, int id);

        [DllImport("user32.dll")]
        internal static extern short GetKeyState(int keyId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern IntPtr SetWindowsHookEx(int kid, LowLevelKeyboardProc llkp, IntPtr h, uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool UnhookWindowsHookEx(IntPtr hookid);

        //Screenshot methods
        [DllImport("gdi32.dll")]
        internal static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest, int nWidth, int nHeight,
            IntPtr hObjectSource, int nXSrc, int nYSrc, int dwRop);

        [DllImport("gdi32.dll")]
        internal static extern IntPtr CreateCompatibleBitmap(IntPtr hDc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        internal static extern IntPtr CreateCompatibleDC(IntPtr hDc);

        [DllImport("gdi32.dll")]
        internal static extern bool DeleteDC(IntPtr hDc);

        [DllImport("gdi32.dll")]
        internal static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        internal static extern IntPtr SelectObject(IntPtr hDc, IntPtr hObject);

        [DllImport("wininet.dll", ExactSpelling = true, CharSet = CharSet.Ansi, SetLastError = true)]
        internal static extern bool InternetGetConnectedState(IntPtr lpSFlags, int dwReserved);

        //DirectX Native methods...

        [DllImport("user32.dll")]
        internal static extern bool GetClientRect(IntPtr hWnd, out Rect lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr hWnd, out Rect lpRect);

        internal static Rectangle GetClientRect(IntPtr hwnd)
        {
            Rect rect;
            GetClientRect(hwnd, out rect);
            return rect.AsRectangle;
        }

        internal static Rectangle GetWindowRect(IntPtr hwnd)
        {
            Rect rect;
            GetWindowRect(hwnd, out rect);
            return rect.AsRectangle;
        }

        internal static Rectangle GetAbsoluteClientRect(IntPtr hWnd)
        {
            var windowRect = GetWindowRect(hWnd);
            var clientRect = GetClientRect(hWnd);
            var chromeWidth = (windowRect.Width - clientRect.Width)/2;
            return
                new Rectangle(
                    new Point(windowRect.X + chromeWidth,
                        windowRect.Y + (windowRect.Height - clientRect.Height - chromeWidth)), clientRect.Size);
        }

        //Keyboard actions realted events
        internal delegate int LowLevelKeyboardProc(bool pressEnd, int keyData, long keyState);

        [StructLayout(LayoutKind.Sequential)]
        internal struct Rect
        {
            internal int left;
            internal int top;
            internal int right;
            internal int bottom;

            public Rect(int left, int top, int right, int bottom)
            {
                this.left = left;
                this.top = top;
                this.right = right;
                this.bottom = bottom;
            }

            public Rectangle AsRectangle
            {
                get { return new Rectangle(this.left, this.top, this.right - this.left, this.bottom - this.top); }
            }

            public static Rect FromXywh(int x, int y, int width, int height)
            {
                return new Rect(x, y, x + width, y + height);
            }

            public static Rect FromRectangle(Rectangle rect)
            {
                return new Rect(rect.Left, rect.Top, rect.Right, rect.Bottom);
            }
        }
    }
}
