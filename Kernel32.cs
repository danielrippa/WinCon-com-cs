using System;
using System.Runtime.InteropServices;

namespace WinCon {

  internal static class Kernel32 {

    private const string Dll = "kernel32.dll";
    private const int STD_OUTPUT_HANDLE = -11;

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

  }

}