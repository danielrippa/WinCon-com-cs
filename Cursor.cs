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

  }

}
