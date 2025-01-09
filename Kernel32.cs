using System;
using System.Runtime.InteropServices;

namespace WinCon {

  internal static class Kernel32 {

    private const string Dll = "kernel32.dll";

    private const int STD_OUTPUT_HANDLE = -11;
    public const int STD_INPUT_HANDLE = -10;

    [DllImport(Dll)]
    internal static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport(Dll)]
    internal static extern IntPtr CreateConsoleScreenBuffer(
      uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes, uint dwFlags, IntPtr lpScreenBufferData
    );

    [DllImport(Dll)]
    internal static extern bool WriteConsole(
      IntPtr hConsoleOutput, string lpBuffer, uint nNumberOfCharsToWrite, out uint lpNumberOfCharsWritten, IntPtr lpReserved
    );

    [DllImport(Dll)]
    internal static extern bool SetConsoleActiveScreenBuffer(IntPtr hConsoleOutput);

    [DllImport(Dll)]
    internal static extern bool SetConsoleCursorPosition(IntPtr hConsoleOutput, COORD dwCursorPosition);

    [DllImport(Dll)]
    internal static extern bool SetConsoleTextAttribute(IntPtr hConsoleOutput, short wAttributes);

    private const uint GENERIC_READ =  0x80000000;
    private const uint GENERIC_WRITE = 0x40000000;

    private const uint CONSOLE_TEXTMODE_BUFFER = 1;

    internal static IntPtr GetStdOOutputHandle() {
      return GetStdHandle(STD_OUTPUT_HANDLE);
    }

    internal static IntPtr CreateScreenBufferHandle() {
      return CreateConsoleScreenBuffer(GENERIC_READ | GENERIC_WRITE, 0, IntPtr.Zero, CONSOLE_TEXTMODE_BUFFER, IntPtr.Zero);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct CONSOLE_CURSOR_INFO {
      public uint dwSize;
      public bool bVisible;
    }

    [DllImport(Dll)]
    private static extern bool GetConsoleCursorInfo(IntPtr hConsoleOutput, out CONSOLE_CURSOR_INFO lpConsoleCursorInfo);

    [DllImport(Dll)]
    internal static extern bool SetConsoleCursorInfo(IntPtr hConsoleOutput, ref CONSOLE_CURSOR_INFO lpConsoleCursorInfo);

    internal static CONSOLE_CURSOR_INFO GetCursorInfo(IntPtr hConsoleOutput) {
      GetConsoleCursorInfo(hConsoleOutput, out CONSOLE_CURSOR_INFO cursorInfo);
      return cursorInfo;
    }

    internal static IntPtr GetStdOutHandle() {
      return GetStdHandle(STD_OUTPUT_HANDLE);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct COORD {
      public short X;
      public short Y;
    }

    [DllImport(Dll)]
    internal static extern bool CloseHandle(IntPtr hObject);

    [StructLayout(LayoutKind.Sequential)]
    internal struct SMALL_RECT {
      public short Left;
      public short Top;
      public short Right;
      public short Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct CONSOLE_SCREEN_BUFFER_INFO {
      public COORD dwSize;
      public COORD dwCursorPosition;
      public short wAttributes;
      public SMALL_RECT srWindow;
      public COORD dwMaximumWindowSize;
    }

    [DllImport(Dll)]
    internal static extern bool GetConsoleScreenBufferInfo(IntPtr hConsoleOutput, out CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo);

    internal static CONSOLE_SCREEN_BUFFER_INFO GetScreenBufferInfo(IntPtr hConsoleOutput) {
      GetConsoleScreenBufferInfo(hConsoleOutput, out var info);
      return info;
    }

    [DllImport(Dll)]
    internal static extern bool SetConsoleScreenBufferInfo(IntPtr hConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo);

    internal const ushort FOCUS_EVENT = 0x0010;
    internal const ushort WINDOW_BUFFER_SIZE_EVENT = 0x0004;
    internal const ushort KEY_EVENT = 0x0001;
    internal const ushort MOUSE_EVENT = 0x0002;

    internal const uint DOUBLE_CLICK = 0x0002;
    internal const uint MOUSE_HWHEELED = 0x0008;
    internal const uint MOUSE_MOVED = 0x0001;
    internal const uint MOUSE_WHEELED = 0x0004;
    internal const uint CAPSLOCK_ON = 0x0080;
    internal const uint ENHANCED_KEY = 0x0100;
    internal const uint LEFT_ALT_PRESSED = 0x0002;
    internal const uint LEFT_CTRL_PRESSED = 0x0008;
    internal const uint NUMLOCK_ON = 0x0020;
    internal const uint RIGHT_ALT_PRESSED = 0x0001;
    internal const uint RIGHT_CTRL_PRESSED = 0x0004;
    internal const uint SCROLLLOCK_ON = 0x0040;
    internal const uint SHIFT_PRESSED = 0x0010;

    internal const uint FROM_LEFT_1ST_BUTTON_PRESSED = 0x0001;
    internal const uint FROM_LEFT_2ND_BUTTON_PRESSED = 0x0004;
    internal const uint FROM_LEFT_3RD_BUTTON_PRESSED = 0x0008;
    internal const uint FROM_LEFT_4TH_BUTTON_PRESSED = 0x0010;
    internal const uint RIGHTMOST_BUTTON_PRESSED = 0x0002;

    internal const uint ENABLE_ECHO_INPUT = 0x0004;
    internal const uint ENABLE_INSERT_MODE = 0x0020;
    internal const uint ENABLE_LINE_INPUT = 0x0002;
    internal const uint ENABLE_MOUSE_INPUT = 0x0010;
    internal const uint ENABLE_PROCESSED_INPUT = 0x0001;
    internal const uint ENABLE_QUICK_EDIT_MODE = 0x0040;
    internal const uint ENABLE_WINDOW_INPUT = 0x0008;

    [DllImport(Dll)]
    internal static extern bool GetNumberOfConsoleInputEvents(IntPtr hConsoleInput, ref uint lpcNumberOfEvents);

    [DllImport(Dll)]
    internal static extern bool PeekConsoleInput(IntPtr hConsoleInput, ref INPUT_RECORD lpBuffer, uint nLength, ref uint lpNumberOfEventsRead);

    [DllImport(Dll)]
    internal static extern bool ReadConsoleInput(IntPtr hConsoleInput, ref INPUT_RECORD lpBuffer, uint nLength, ref uint lpNumberOfEventsRead);

    [DllImport(Dll)]
    internal static extern bool FlushConsoleInputBuffer(IntPtr hConsoleInput);

    [DllImport(Dll)]
    internal static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

    [DllImport(Dll)]
    internal static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    [StructLayout(LayoutKind.Explicit)]
  internal struct INPUT_RECORD
        {
            [FieldOffset(0)]
            public ushort EventType;
            [FieldOffset(4)]
            public KEY_EVENT_RECORD KeyEvent;
            [FieldOffset(4)]
            public MOUSE_EVENT_RECORD MouseEvent;
            [FieldOffset(4)]
            public WINDOW_BUFFER_SIZE_RECORD WindowBufferSizeEvent;
            [FieldOffset(4)]
            public FOCUS_EVENT_RECORD FocusEvent;
        }

    [StructLayout(LayoutKind.Sequential)]
    internal struct KEY_EVENT_RECORD {
        public bool bKeyDown;
        public ushort wRepeatCount;
        public ushort wVirtualKeyCode;
        public ushort wVirtualScanCode;
        public char UnicodeChar;
        public uint dwControlKeyState;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MOUSE_EVENT_RECORD {
        public COORD dwMousePosition;
        public uint dwButtonState;
        public uint dwControlKeyState;
        public uint dwEventFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct WINDOW_BUFFER_SIZE_RECORD {
        public COORD dwSize;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct FOCUS_EVENT_RECORD {
        public bool bSetFocus;
    }


    internal static int HiWord(int number) {
        return (number >> 16) & 0xFFFF;
    }

  }

}