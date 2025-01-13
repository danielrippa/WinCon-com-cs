using System;
using System.Runtime.InteropServices;

namespace WinCon {

  internal static class Kernel32 {

    private const string Dll = "kernel32.dll";

    //

    private const int STD_OUTPUT_HANDLE = -11;
    public const int STD_INPUT_HANDLE = -10;

    [DllImport(Dll)]
    private static extern IntPtr GetStdHandle(int nStdHandle);

    internal static IntPtr GetStdOutHandle() {
      return GetStdHandle(STD_OUTPUT_HANDLE);
    }

    internal static IntPtr GetStdInHandle() {
      return GetStdHandle(STD_OUTPUT_HANDLE);
    }

    [DllImport(Dll)]
    private static extern IntPtr CreateConsoleScreenBuffer(
      uint dwDesiredAccess, uint dwShareMode, IntPtr lpSecurityAttributes, uint dwFlags, IntPtr lpScreenBufferData
    );

    private const uint GENERIC_READ =  0x80000000;
    private const uint GENERIC_WRITE = 0x40000000;

    private const uint CONSOLE_TEXTMODE_BUFFER = 1;

    internal static IntPtr CreateScreenBufferHandle() {
      return CreateConsoleScreenBuffer(GENERIC_READ | GENERIC_WRITE, 0, IntPtr.Zero, CONSOLE_TEXTMODE_BUFFER, IntPtr.Zero);
    }

    [DllImport(Dll)]
    internal static extern bool CloseHandle(IntPtr handle);

    //

    [StructLayout(LayoutKind.Sequential)]
    internal struct COORD {
      public short X;
      public short Y;
    }

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
    private static extern bool GetConsoleScreenBufferInfo(IntPtr hConsoleOutput, out CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo);

    internal static CONSOLE_SCREEN_BUFFER_INFO GetScreenBufferInfo(IntPtr handle) {
      GetConsoleScreenBufferInfo(handle, out var info);
      return info;
    }

    [DllImport(Dll)]
    private static extern bool SetConsoleScreenBufferInfo(IntPtr hConsoleOutput, ref CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo);

    [DllImport(Dll)]
    internal static extern bool SetConsoleTextAttribute(IntPtr hConsoleOutput, short wAttributes);

    [DllImport(Dll)]
    internal static extern bool SetConsoleActiveScreenBuffer(IntPtr hConsoleOutput);

    [DllImport(Dll)]
    internal static extern bool WriteConsole(
      IntPtr hConsoleOutput, string lpBuffer, uint nNumberOfCharsToWrite, out uint lpNumberOfCharsWritten, IntPtr lpReserved
    );

    // Cursor

    [DllImport(Dll)]
    internal static extern bool SetConsoleCursorPosition(IntPtr hConsoleOutput, COORD dwCursorPosition);

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

    // Mode

    [DllImport(Dll)]
    private static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

    [DllImport(Dll)]
    private static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);

    internal static bool GetConsoleModeBit(IntPtr handle, uint bit) {
      GetConsoleMode(handle, out uint mode);
      return (mode & bit) != 0;
    }

    internal static void SetConsoleModeBit(IntPtr handle, uint bit, bool value) {
      GetConsoleMode(handle, out uint mode);
      mode = value ? (mode | bit): (mode & ~bit);
      SetConsoleMode(handle, mode);
    }

    internal static uint ENABLE_PROCESSED_OUTPUT = 0x0001;
    internal static uint ENABLE_WRAP_AT_EOL_OUTPUT = 0x0002;
    internal static uint ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004;
    internal static uint DISABLE_NEWLINE_AUTO_RETURN = 0x0008;
    internal static uint ENABLE_LVB_GRID_WORLDWIDE = 0x0010;

  }

}