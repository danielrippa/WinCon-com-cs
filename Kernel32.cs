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

  }

}