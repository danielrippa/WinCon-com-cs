using System;
using System.Runtime.InteropServices;

namespace WinCon {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("1B775129-9A5A-4D14-9465-57B38A71BC5B")]
  [ProgId("WinCon.ScreenBuffer")]

  public class ScreenBuffer {

    private IntPtr _handle;
    private Cursor _cursor;
    private TextAttributes _textAttributes;
    
    public IntPtr Handle {
      get => _handle;
      set {
        _handle = value;
        _cursor.Handle = value;
        _textAttributes.Handle = value;
      }
    }

    public ScreenBuffer() {
      _cursor = new Cursor();
      _textAttributes = new TextAttributes();
      Handle = Kernel32.GetStdOutHandle();
    }

    public Cursor Cursor {
      get => _cursor;
    }

    public TextAttributes TextAttributes {
      get => _textAttributes;
    }
    
    public void Write(string lpBuffer) {
      Kernel32.WriteConsole(Handle, lpBuffer, (uint)lpBuffer.Length, out _, IntPtr.Zero);
    }

    public void AssignHandle() {
      Handle = Kernel32.CreateScreenBufferHandle();
    }

    public void Close() {
      Kernel32.CloseHandle(Handle);
    }

    public void Activate() {
      Kernel32.SetConsoleActiveScreenBuffer(Handle);
    }

  }
}
