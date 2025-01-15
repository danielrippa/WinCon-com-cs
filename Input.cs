using System;
using System.Runtime.InteropServices;
using Newtonsoft.Json;

namespace WinCon {

  [ComVisible(true)]
  [ClassInterface(ClassInterfaceType.AutoDispatch)]
  [Guid("E4B35606-B68E-4B54-A438-E2DD1B139022")]
  [ProgId("WinCon.Input")]

  public class Input {

    private IntPtr handle = Kernel32.GetStdInHandle();

    public string GetInputEvent() {

      object eventDetails = new { type = "None" };

      var inputRecord = ReadInputRecord();

      switch (inputRecord.EventType) {

        case Kernel32.FOCUS_EVENT:
        case Kernel32.WINDOW_BUFFER_SIZE_EVENT:
          eventDetails = GetWindowEvent(inputRecord);
          break;

        case Kernel32.KEY_EVENT:
          eventDetails = GetKeyEvent(inputRecord);
          break;

        case Kernel32.MOUSE_EVENT:
          eventDetails = GetMouseEvent(inputRecord);
          break;

        default:
          Console.WriteLine($"Otro event que nada que ver {inputRecord.EventType}");
          break;

      }

      return JsonConvert.SerializeObject(eventDetails);
    }

    // modes

    public bool EchoInput {
      get => GetModeBit(Kernel32.ENABLE_ECHO_INPUT);
      set => SetModeBit(Kernel32.ENABLE_ECHO_INPUT, value);
    }

    public bool QuickEdit {
      get => GetModeBit(Kernel32.ENABLE_QUICK_EDIT_MODE);
      set {
        var mode = Kernel32.ENABLE_EXTENDED_FLAGS;
        // https://stackoverflow.com/questions/76156132/non-documented-console-input-mode-flags

        if (value) {
          mode = Kernel32.ENABLE_QUICK_EDIT_MODE | mode;
        }
        SetModeBit(mode, value);
      }
    }

    public bool ProcessedInput {
      get => GetModeBit(Kernel32.ENABLE_PROCESSED_INPUT);
      set => SetModeBit(Kernel32.ENABLE_PROCESSED_INPUT, value);
    }

    public bool Insert {
      get => GetModeBit(Kernel32.ENABLE_INSERT_MODE);
      set {
        var mode = Kernel32.ENABLE_EXTENDED_FLAGS;
        // https://stackoverflow.com/questions/76156132/non-documented-console-input-mode-flags

        if (value) {
          mode = Kernel32.ENABLE_INSERT_MODE | mode;
        };
        SetModeBit(mode, value);
      }
    }

    public bool LineInput {
      get => GetModeBit(Kernel32.ENABLE_LINE_INPUT);
      set => SetModeBit(Kernel32.ENABLE_LINE_INPUT, value);
    }

    public bool MouseInput {
      get => GetModeBit(Kernel32.ENABLE_MOUSE_INPUT);
      set => SetModeBit(Kernel32.ENABLE_MOUSE_INPUT, value);
    }

    public bool WindowInput {
      get => GetModeBit(Kernel32.ENABLE_WINDOW_INPUT);
      set => SetModeBit(Kernel32.ENABLE_WINDOW_INPUT, value);
    }

    private bool GetModeBit(uint bit) {
      return Kernel32.GetConsoleModeBit(handle, bit);
    }

    private void SetModeBit(uint bit, bool value) {
      Kernel32.SetConsoleModeBit(handle, bit, value);
    }

    // events

    private uint PendingEventsCount() {
      uint events = 0;
      Kernel32.GetNumberOfConsoleInputEvents(handle, ref events);
      return events;
    }

    private Kernel32.INPUT_RECORD ReadInputRecord() {
      uint events = 0;
      var inputRecord = new Kernel32.INPUT_RECORD();
      Kernel32.ReadConsoleInput(handle, ref inputRecord, 1, ref events);
      if (events == 0) {
        inputRecord.EventType = 0; // An invalid value
      }
      return inputRecord;
    }

    private object GetWindowEvent(Kernel32.INPUT_RECORD inputRecord) {
      switch (inputRecord.EventType) {

        case Kernel32.FOCUS_EVENT:
          return new {
            type = "WindowFocus",
            focused = inputRecord.FocusEvent.bSetFocus
          };

        case Kernel32.WINDOW_BUFFER_SIZE_EVENT:
          var size = inputRecord.WindowBufferSizeEvent.dwSize;
          return new {
            type = "WindowResized",
            rows = size.X,
            columns = size.Y
          };

      }

      return new {};
    }

    private object GetControlKeys(uint controlKeyState) {
      return new {

        capsLockOn = (controlKeyState & Kernel32.CAPSLOCK_ON) != 0,
        enhancedKey = (controlKeyState & Kernel32.ENHANCED_KEY) != 0,
        leftAltPressed = (controlKeyState & Kernel32.LEFT_ALT_PRESSED) != 0,
        leftCtrlPressed = (controlKeyState & Kernel32.LEFT_CTRL_PRESSED) != 0,
        numLockOn = (controlKeyState & Kernel32.NUMLOCK_ON) != 0,
        rightAltPressed = (controlKeyState & Kernel32.RIGHT_ALT_PRESSED) != 0,
        rightCtrlPressed = (controlKeyState & Kernel32.RIGHT_CTRL_PRESSED) != 0,
        scrollLockOn = (controlKeyState & Kernel32.SCROLLLOCK_ON) != 0,
        shiftPressed = (controlKeyState & Kernel32.SHIFT_PRESSED) != 0

      };
    }

    private string GetKeyType(int keyCode) {

        if (keyCode >= 0x41 && keyCode <= 0x5A)
            return "alphabetic";
        if (keyCode >= 0x30 && keyCode <= 0x39)
            return "numeric";
        if (keyCode >= 0x70 && keyCode <= 0x87)
            return "function";
        if (keyCode >= 0x21 && keyCode <= 0x28)
            return "navigation";
        if (keyCode == 0x20)
            return "alphabetic";
        if (keyCode == 0x09)
            return "navigation";
        if (keyCode == 0x08 || keyCode == 0x2E || keyCode == 0x2D)
            return "edition";
        if (keyCode >= 0xBB && keyCode <= 0xBE)
            return "punctuation";
        if (keyCode >= 0xAD && keyCode <= 0xB3)
            return "media";

        return "unknown";
    }


    private object GetKeyEvent(Kernel32.INPUT_RECORD inputRecord) {

      var key = inputRecord.KeyEvent;

      return new {
        type = key.bKeyDown ? "KeyPressed" : "KeyReleased",

        scanCode = key.wVirtualScanCode,
        keyCode = key.wVirtualKeyCode,
        unicodeChar = key.UnicodeChar,

        keyType = GetKeyType(key.wVirtualKeyCode),

        controlKeys = GetControlKeys(key.dwControlKeyState),

        repetitions = key.wRepeatCount
      };
    }

    private object GetMouseEvent(Kernel32.INPUT_RECORD inputRecord) {
      var mouse = inputRecord.MouseEvent;

      var doubleClicked = false;
      var eventType = "MouseButtonClicked";

      var state = (short)Kernel32.HiWord((int)mouse.dwButtonState) > 0;
      var wheelDirection = "";

      if (mouse.dwEventFlags == 0) {

        eventType = "MouseButtonReleased";

      } else {

        switch (mouse.dwEventFlags) {

          case Kernel32.MOUSE_HWHEELED:
            eventType = "HorizontalWheel";
            wheelDirection = state ? "right" : "left";
            break;

          case Kernel32.MOUSE_WHEELED:
            eventType = "VerticalWheel";
            wheelDirection = state ? "forward" : "backwards";
            break;

          case Kernel32.DOUBLE_CLICK:
            doubleClicked = true;
            break;

          case Kernel32.MOUSE_MOVED:
            eventType = "MouseMoved";
            break;

        }

      }

      return new {

        type = eventType,

        doubleClick = doubleClicked,

        location = new {
          row = mouse.dwMousePosition.X,
          column = mouse.dwMousePosition.Y
        },

        buttons = GetMouseButtons(mouse.dwButtonState),
        controlKeys = GetControlKeys(mouse.dwControlKeyState),

        wheelDirection = wheelDirection

      };
    }

    private object GetMouseButtons(uint buttonsState) {
      return new {

        left1Pressed = (buttonsState & Kernel32.FROM_LEFT_1ST_BUTTON_PRESSED) != 0,
        left2Pressed = (buttonsState & Kernel32.FROM_LEFT_2ND_BUTTON_PRESSED) != 0,
        left3Pressed = (buttonsState & Kernel32.FROM_LEFT_3RD_BUTTON_PRESSED) != 0,
        left4Pressed = (buttonsState & Kernel32.FROM_LEFT_4TH_BUTTON_PRESSED) != 0,

        rightMostPressed = (buttonsState & Kernel32.RIGHTMOST_BUTTON_PRESSED) != 0

      };
    }

  }

}