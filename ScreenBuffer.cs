using System;
using System.Runtime.InteropServices;

namespace WinCon {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("1B775129-9A5A-4D14-9465-57B38A71BC5B")]
  [ProgId("WinCon.ScreenBuffer")]

  public class ScreenBuffer {

    private IntPtr handle = Kernel32.GetStdOutHandle();

    private Cursor cursor = new Cursor();
    private TextAttributes textAttributes = new TextAttributes();

    public IntPtr Handle {
      get => handle;
      set {
        handle = value;

        cursor.Handle = value;
        textAttributes.Handle = value;
      }
    }

    public Cursor Cursor {
      get => cursor;
    }

    public TextAttributes TextAttributes {
      get => textAttributes;
    }

    //

    public void Write(string lpBuffer) {
      Kernel32.WriteConsole(Handle, lpBuffer, (uint)lpBuffer.Length, out _, IntPtr.Zero);
    }

    //

    public static IntPtr CreateHandle() {
      return Kernel32.CreateScreenBufferHandle();
    }

    public void AssignHandle() {
      Handle = CreateHandle();
    }

    public void Close() {
      Kernel32.CloseHandle(Handle);
    }

    public void Activate() {
      Kernel32.SetConsoleActiveScreenBuffer(Handle);
    }

    // Modes

    public bool ProcessedOutput {
      get => GetModeBit(Kernel32.ENABLE_PROCESSED_OUTPUT);
      set => SetModeBit(Kernel32.ENABLE_PROCESSED_OUTPUT, value);
    }

    public bool WrapAtEOL {
      get => GetModeBit(Kernel32.ENABLE_WRAP_AT_EOL_OUTPUT);
      set => SetModeBit(Kernel32.ENABLE_WRAP_AT_EOL_OUTPUT, value);
    }

    public bool TerminalMode {
      get => GetModeBit(Kernel32.ENABLE_VIRTUAL_TERMINAL_PROCESSING);
      set => SetModeBit(Kernel32.ENABLE_VIRTUAL_TERMINAL_PROCESSING, value);
    }

    public bool BordersMode {
      get => GetModeBit(Kernel32.ENABLE_LVB_GRID_WORLDWIDE);
      set => SetModeBit(Kernel32.ENABLE_LVB_GRID_WORLDWIDE, value);
    }

    private bool GetModeBit(uint bit) {
      return Kernel32.GetConsoleModeBit(Handle, bit);
    }

    private void SetModeBit(uint bit, bool value) {
      Kernel32.SetConsoleModeBit(Handle, bit, value);
    }

  }

}