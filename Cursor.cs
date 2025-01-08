using System;
using System.Runtime.InteropServices;

namespace WinCon {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("AE72DC0B-E2EC-42B4-B95E-0A1171897413")]
  [ProgId("WinCon.Cursor")]

  public class Cursor {

    public IntPtr Handle;

    public Cursor() {
      Handle = Kernel32.GetStdOutHandle();
    }

    public void Goto(short x, short y) {
      var position = new Kernel32.COORD { X = x, Y = y };
      Kernel32.SetConsoleCursorPosition(Handle, position);
    }

    private Kernel32.COORD GetPosition() {
      var info = Kernel32.GetScreenBufferInfo(Handle);
      return info.dwCursorPosition;
    }

    public short Row {
      get => GetPosition().Y;
      set => Goto(Column, value);
    }

    public short Column {
      get => GetPosition().X;
      set => Goto(value, Row);
    }

    private Kernel32.CONSOLE_CURSOR_INFO GetCursorInfo() {
      return Kernel32.GetCursorInfo(Handle);
    }

    private void SetCursorInfo(Kernel32.CONSOLE_CURSOR_INFO info) {
      Kernel32.SetConsoleCursorInfo(Handle, ref info);
    }

    public short Size {
      get => (short)GetCursorInfo().dwSize;
      set {
        var info = GetCursorInfo();
        info.dwSize = (uint)value;
        SetCursorInfo(info);
      }

    }

    public bool Visible {
      get => GetCursorInfo().bVisible;
      set {
        var info = GetCursorInfo();
        info.bVisible = value;
        SetCursorInfo(info);
      }
    }

  }

}