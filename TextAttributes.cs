using System;
using System.Runtime.InteropServices;

  namespace WinCon
  {

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("17982804-7649-47E9-9787-805F29744287")]
    [ProgId("WinCon.TextAttributes")]

    public class TextAttributes
    {

      public IntPtr Handle = Kernel32.GetStdOutHandle();

      public short Value
      {
        get => Kernel32.GetScreenBufferInfo(Handle).wAttributes;
        set => Kernel32.SetConsoleTextAttribute(Handle, value);
      }

      public TextColor Ink {
        get => GetTextColorFromValue(Value);
        set => Value = ApplyTextColor(Value, value);
      }

      public TextColor Paper {
        get => GetTextColorFromValue(Value, 4);
        set => Value = ApplyTextColor(Value, value, 4);
      }

      public bool Inverted {
        get => (Value & (1 << 14)) != 0;
        set => Value = ApplyBit(Value, 14, value);
      }

      public TextBorder Border {
        get => GetTextBorderFromValue(Value);
        set => Value = ApplyTextBorder(Value, value);
      }

      //

      private static TextColor GetTextColorFromValue(short value, int offset = 0) {
        return new TextColor {
          Intensity = (value & (1 << (3 + offset))) != 0,

          Red = (value & (1 << (2 + offset))) != 0,
          Green = (value & (1 << (1 + offset))) != 0,
          Blue = (value & (1 << (0 + offset))) != 0
        };
      }

      private static TextBorder GetTextBorderFromValue(short value) {
        return new TextBorder {
          Top = (value & (1 << 10)) != 0,
          Left = (value & (1 << 11)) != 0,
          Right = (value & (1 << 12)) != 0,
          Bottom = (value & (1 << 15)) != 0
        };
      }

      private static short ApplyBit(short value, int bit, bool enabled) {
        short mask = (short)(1 << bit);
        return enabled ? (short)(value | mask) : (short)(value & ~mask);
      }

      private static short ApplyTextColor(short value, TextColor color, int offset = 0) {
        if (color.Intensity.HasValue) value = ApplyBit(value, 3 + offset, color.Intensity.Value);

        if (color.Red.HasValue) value = ApplyBit(value, 2 + offset, color.Red.Value);
        if (color.Green.HasValue) value = ApplyBit(value, 1 + offset, color.Green.Value);
        if (color.Blue.HasValue) value = ApplyBit(value, 0 + offset, color.Blue.Value);

        return value;
      }

      private static short ApplyTextBorder(short value, TextBorder border) {
        if (border.Top.HasValue) value = ApplyBit(value, 10, border.Top.Value);
        if (border.Left.HasValue) value = ApplyBit(value, 11, border.Left.Value);
        if (border.Bottom.HasValue) value = ApplyBit(value, 15, border.Bottom.Value);
        if (border.Right.HasValue) value = ApplyBit(value, 12, border.Right.Value);

        return value;
      }

    }

    //

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("65F7D943-FBB2-4942-B4E8-BD355CEDE334")]
    [ProgId("WinCon.TextColor")]
    public class TextColor {

      public bool? Red { get; set; }
      public bool? Green { get; set; }
      public bool? Blue { get; set; }
      public bool? Intensity { get; set; }

    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("D4E1F3A2-5B8B-4F8A-9B6E-8C6B3A8D4F3A")]
    [ProgId("WinCon.TextColors")]
    public class TextColors {

      public TextColor Black { get; } = CreateTextColor();
      public TextColor Red { get; } = CreateTextColor(red: true);
      public TextColor Green { get; } = CreateTextColor(green: true);
      public TextColor Navy { get; } = CreateTextColor(blue: true);

      public TextColor Celestial { get; } = CreateTextColor(green: true, blue: true);
      public TextColor Curry { get; } = CreateTextColor(red: true, green: true);
      public TextColor Purple { get; } = CreateTextColor(red: true, blue: true);
      public TextColor Gray { get; } = CreateTextColor(red: true, green: true, blue: true);

      public TextColor Silver { get; } = CreateTextColor(intensity: true);
      public TextColor Vermillion { get; } = CreateTextColor(red: true, intensity: true);
      public TextColor Emerald { get; } = CreateTextColor(green: true, intensity: true);
      public TextColor Blue { get; } = CreateTextColor(blue: true, intensity: true);
      public TextColor Turquoise { get; } = CreateTextColor(green: true, blue: true, intensity: true);
      public TextColor Yellow { get; } = CreateTextColor(red: true, green: true, intensity: true);
      public TextColor Fuchsia { get; } = CreateTextColor(red: true, blue: true, intensity: true);
      public TextColor White { get; } = CreateTextColor(red: true, green: true, blue: true, intensity: true);

      private static TextColor CreateTextColor(bool red = false, bool green = false, bool blue = false, bool intensity = false) {
        return new TextColor {
          Red = red,
          Green = green,
          Blue = blue,

          Intensity = intensity
        };
      }

    }

    //

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("A1928701-43A6-4C67-8286-1E8A3023A8EB")]
    [ProgId("WinCon.TextBorder")]

    public class TextBorder {

      public bool? Top { get; set; }
      public bool? Left { get; set; }
      public bool? Bottom { get; set; }
      public bool? Right { get; set; }

    }

    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    [Guid("B1C2D3E4-F5A6-7B8C-9D0E-1F2A3B4C5D6E")]
    [ProgId("WinCon.TextBorders")]

    public class TextBorders {

      public TextBorder None { get; } = CreateTextBorder();
      public TextBorder TopBorder { get; } = CreateTextBorder(top: true);
      public TextBorder Sides { get; } = CreateTextBorder(left: true, right: true);
      public TextBorder Underlined { get; } = CreateTextBorder(bottom: true);

      public TextBorder NW { get; } = CreateTextBorder(top: true, left: true);
      public TextBorder NE { get; } = CreateTextBorder(top: true, right: true);
      public TextBorder SE { get; } = CreateTextBorder(bottom: true, right: true);
      public TextBorder SW { get; } = CreateTextBorder(bottom: true, left: true);

      public TextBorder All { get; } = CreateTextBorder(top: true, left: true, bottom: true, right: true);

      private static TextBorder CreateTextBorder(bool top = false, bool left = false, bool bottom = false, bool right = false) {
        return new TextBorder {
          Top = top,
          Left = left,
          Bottom = bottom,
          Right = right
        };
      }

    }

  }