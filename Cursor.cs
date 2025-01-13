using System;
using System.Runtime.InteropServices;

namespace WinCon {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("AE72DC0B-E2EC-42B4-B95E-0A1171897413")]
  [ProgId("WinCon.Cursor")]

  public class Cursor {

    public IntPtr Handle = Kernel32.GetStdOutHandle();

    public void Goto(short x, short y) {
      var position = new Kernel32.COORD {
        X = x, Y = y
      };
      Kernel32.SetConsoleCursorPosition(Handle, position);
    }

    private Kernel32.COORD GetPosition() {
      return Kernel32.GetScreenBufferInfo(Handle).dwCursorPosition;
    }

    public short Row {
      get => GetPosition().Y;
      set => Goto(Column, value);
    }

    public short Column {
      get => GetPosition().X;
      set => Goto(value, Row);
    }

    private Kernel32.CONSOLE_CURSOR_INFO Info {
      get => Kernel32.GetCursorInfo(Handle);
      set => Kernel32.SetConsoleCursorInfo(Handle, ref value);
    }

    public short Size {
      get => (short)Info.dwSize;
      set {
        var info = Info;
        info.dwSize = (uint)value;
        Info = info;
      }
    }

    public bool Visible {
      get => Info.bVisible;
      set {
        var info = Info;
        info.bVisible = value;
        Info = info;
      }
    }

  }

}
