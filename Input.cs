using System;
using System.Runtime.InteropServices;
using Newtonsoft.Json;

namespace WinCon {

   [ComVisible(true)]
   [ClassInterface(ClassInterfaceType.AutoDispatch)]
   [Guid("E4B35606-B68E-4B54-A438-E2DD1B139022")]
   [ProgId("WinCon.Input")]

    public class Input {

        private IntPtr inputHandle;

        public Input() {
          inputHandle = Kernel32.GetStdHandle(Kernel32.STD_INPUT_HANDLE);
        }

        public string GetInputEvent() {
          var inputEventType = GetNextInputEventType();

          object eventDetails;
          switch (inputEventType) {
            case InputEventType.WindowEvent:
              eventDetails = GetWindowEventDetails();
              break;
            case InputEventType.KeyEvent:
              eventDetails = GetKeyEventDetails();
              break;
            case InputEventType.MouseEvent:
              eventDetails = GetMouseEventDetails();
              break;
            default:
              eventDetails = new { type = "None" };
              break;
          }

          return JsonConvert.SerializeObject(eventDetails);
        }

        private bool GetConsoleMode(uint bit) {
          Kernel32.GetConsoleMode(inputHandle, out uint mode);
          return (mode & bit) != 0;
        }

        private void SetConsoleMode(uint bit, bool value) {
          Kernel32.GetConsoleMode(inputHandle, out uint mode);
          mode = value ? (mode | bit) : (mode & ~bit);
          Kernel32.SetConsoleMode(inputHandle, mode);
        }

        public bool EchoInputEnabled {
          get { return GetConsoleMode(Kernel32.ENABLE_ECHO_INPUT); }
          set { SetConsoleMode(Kernel32.ENABLE_ECHO_INPUT, value); }
        }

        public bool QuickEditModeEnabled {
          get { return GetConsoleMode(Kernel32.ENABLE_QUICK_EDIT_MODE); }
          set { SetConsoleMode(Kernel32.ENABLE_QUICK_EDIT_MODE, value); }
        }

        public bool ProcessedInputEnabled {
          get { return GetConsoleMode(Kernel32.ENABLE_PROCESSED_INPUT); }
          set { SetConsoleMode(Kernel32.ENABLE_PROCESSED_INPUT, value); }
        }

        public bool InsertModeEnabled {
          get { return GetConsoleMode(Kernel32.ENABLE_INSERT_MODE); }
          set { SetConsoleMode(Kernel32.ENABLE_INSERT_MODE, value); }
        }

        public bool LineInputEnabled {
          get { return GetConsoleMode(Kernel32.ENABLE_LINE_INPUT); }
          set { SetConsoleMode(Kernel32.ENABLE_LINE_INPUT, value); }
        }

        public bool MouseInputEnabled {
          get { return GetConsoleMode(Kernel32.ENABLE_MOUSE_INPUT); }
          set { SetConsoleMode(Kernel32.ENABLE_MOUSE_INPUT, value); }
        }

        public bool WindowInputEnabled {
          get { return GetConsoleMode(Kernel32.ENABLE_WINDOW_INPUT); }
          set { SetConsoleMode(Kernel32.ENABLE_WINDOW_INPUT, value); }
        }

        private InputEventType GetNextInputEventType() {
          uint numberOfEvents = 0;
          if (Kernel32.GetNumberOfConsoleInputEvents(inputHandle, ref numberOfEvents) && numberOfEvents > 0) {

            uint eventsRead = 0;
            Kernel32.INPUT_RECORD inputRecord = new Kernel32.INPUT_RECORD();
            if (Kernel32.PeekConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) && eventsRead > 0) {

              switch (inputRecord.EventType) {
                case Kernel32.FOCUS_EVENT:
                case Kernel32.WINDOW_BUFFER_SIZE_EVENT:
                  return InputEventType.WindowEvent;
                case Kernel32.KEY_EVENT:
                  return InputEventType.KeyEvent;
                case Kernel32.MOUSE_EVENT:
                  return InputEventType.MouseEvent;

                default:
                  return InputEventType.None;
              }

            }
          }
          return InputEventType.None;
        }

        private WindowEventType GetNextWindowEventType()
        {
          uint eventsRead = 0;
          Kernel32.INPUT_RECORD inputRecord = new Kernel32.INPUT_RECORD();
          if (Kernel32.PeekConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) && eventsRead > 0) {
            switch (inputRecord.EventType) {
              case Kernel32.FOCUS_EVENT:
                return WindowEventType.Focus;
              case Kernel32.WINDOW_BUFFER_SIZE_EVENT:
                return WindowEventType.Resized;

              default:
                return WindowEventType.None;
            }
          }
          return WindowEventType.None;
        }

        private KeyEventType GetNextKeyEventType() {
          uint eventsRead = 0;
          Kernel32.INPUT_RECORD inputRecord = new Kernel32.INPUT_RECORD();
          if (Kernel32.PeekConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) && eventsRead > 0) {
            return inputRecord.KeyEvent.bKeyDown ? KeyEventType.Pressed : KeyEventType.Released;
          }
          return KeyEventType.None;
        }

        private MouseEventType GetNextMouseEventType() {
          uint eventsRead = 0;
          Kernel32.INPUT_RECORD inputRecord = new Kernel32.INPUT_RECORD();
          if (Kernel32.PeekConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead) && eventsRead > 0) {

            switch (inputRecord.MouseEvent.dwEventFlags) {
              case Kernel32.DOUBLE_CLICK:
                return MouseEventType.DoubleClick;
              case Kernel32.MOUSE_HWHEELED:
                return MouseEventType.HorizontalWheel;
              case Kernel32.MOUSE_MOVED:
                return MouseEventType.MouseMoved;
              case Kernel32.MOUSE_WHEELED:
                return MouseEventType.VerticalWheel;

              default:
                return inputRecord.MouseEvent.dwButtonState != 0 ? MouseEventType.SingleClick : MouseEventType.ButtonReleased;
            }
          }
          return MouseEventType.None;
        }

        private KeyEventBase GetNextKeyEvent(bool isPressed) {
          KeyEventBase keyEvent;
          if (isPressed) {
            keyEvent = new PressedKeyEvent();
          } else {
            keyEvent = new ReleasedKeyEvent();
          };

          uint eventsRead = 0;
          Kernel32.INPUT_RECORD inputRecord = new Kernel32.INPUT_RECORD();
          Kernel32.ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead);

          var keyEventDetails = inputRecord.KeyEvent;
          keyEvent.ScanCode = keyEventDetails.wVirtualScanCode;
          keyEvent.KeyCode = keyEventDetails.wVirtualKeyCode;
          keyEvent.UnicodeChar = keyEventDetails.UnicodeChar;
          keyEvent.ControlKeys = GetControlKeys(keyEventDetails.dwControlKeyState);

          if (isPressed && keyEvent is PressedKeyEvent pressedKeyEvent) {
            pressedKeyEvent.Repetitions = keyEventDetails.wRepeatCount;
          }
          return keyEvent;
        }

        private object GetKeyEventDetails() {
          var keyEventType = GetNextKeyEventType();
          switch (keyEventType) {
            case KeyEventType.Pressed:
              return CreateKeyEventDetails("KeyPressed", (PressedKeyEvent)GetNextKeyEvent(true));
            case KeyEventType.Released:
              return CreateKeyEventDetails("KeyReleased", (ReleasedKeyEvent)GetNextKeyEvent(false));

            default:
              return new { type = "None" };
          }
        }

        private MouseEventBase GetNextMouseEvent(MouseEventType eventType) {
          MouseEventBase mouseEvent;
          switch (eventType) {
            case MouseEventType.DoubleClick:
            case MouseEventType.SingleClick:
              mouseEvent = new MouseButtonClickedEvent();
              break;
            case MouseEventType.HorizontalWheel:
              mouseEvent = new MouseHorizontalWheelEvent();
              break;
            case MouseEventType.VerticalWheel:
              mouseEvent = new MouseVerticalWheelEvent();
              break;
            default:
              mouseEvent = new MouseEvent();
              break;
          }

          uint eventsRead = 0;
          Kernel32.INPUT_RECORD inputRecord = new Kernel32.INPUT_RECORD();
          Kernel32.ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead);

          mouseEvent.CursorLocation = new MouseCursorLocation { Row = inputRecord.MouseEvent.dwMousePosition.Y, Column = inputRecord.MouseEvent.dwMousePosition.X };
          mouseEvent.ControlKeys = GetControlKeys(inputRecord.MouseEvent.dwControlKeyState);

          if (mouseEvent is MouseButtonClickedEvent clickedEvent) {
              clickedEvent.Buttons = GetMouseButtons(inputRecord.MouseEvent.dwButtonState);
          } else if (mouseEvent is MouseWheelEvent wheelEvent) {
              wheelEvent.WheelDirection = (short)Kernel32.HiWord((int)inputRecord.MouseEvent.dwButtonState) > 0 ?
                (mouseEvent is MouseHorizontalWheelEvent ? "right" : "forward") :
                (mouseEvent is MouseHorizontalWheelEvent ? "left" : "backwards");
          }

          return mouseEvent;
        }

        private object GetMouseEventDetails() {
            var mouseEventType = GetNextMouseEventType();
            var mouseEvent = GetNextMouseEvent(mouseEventType);

            var cursorLocation = new {
              row = mouseEvent.CursorLocation.Row,
              column = mouseEvent.CursorLocation.Column
            };

            var controlKeys = GetControlKeysDetails(mouseEvent.ControlKeys);

            switch (mouseEventType) {
              case MouseEventType.DoubleClick:
              case MouseEventType.SingleClick:
                return new {
                  type = "MouseClick",
                  cursorLocation,
                  buttons = GetMouseButtonsDetails(((MouseButtonClickedEvent)mouseEvent).Buttons),
                  controlKeys
                };
              case MouseEventType.MouseMoved:
                return new {
                  type = "MouseMoved",
                  cursorLocation,
                  controlKeys
                };
              case MouseEventType.HorizontalWheel:
                return new {
                  type = "HorizontalWheel",
                  cursorLocation,
                  direction = ((MouseHorizontalWheelEvent)mouseEvent).WheelDirection,
                  controlKeys
                };
              case MouseEventType.VerticalWheel:
                return new {
                  type = "VerticalWheel",
                  cursorLocation,
                  direction = ((MouseVerticalWheelEvent)mouseEvent).WheelDirection,
                  controlKeys
                };
              case MouseEventType.ButtonReleased:
                return new {
                  type = "MouseButtonReleased",
                  cursorLocation,
                  controlKeys
                };
              default:
                return new { type = "None" };
            }
        }

        private object GetWindowEventDetails() {
          var windowEventType = GetNextWindowEventType();
          switch (windowEventType) {
            case WindowEventType.Focus:
              return new {
                type = "WindowFocus",
                focused = GetNextWindowFocusEvent().Focused
              };
            case WindowEventType.Resized:
              return new {
                type = "WindowResized",
                rows = GetNextWindowResizedEvent().Rows,
                columns = GetNextWindowResizedEvent().Columns
              };
            default:
              return new { type = "None" };
          }
        }

        private object CreateKeyEventDetails(string eventType, KeyEventBase keyEvent) {
          return new {
            type = eventType,
            scanCode = keyEvent.ScanCode,
            keyCode = keyEvent.KeyCode,
            unicodeChar = keyEvent.UnicodeChar,
            repetitions = keyEvent is PressedKeyEvent pressedKeyEvent ? pressedKeyEvent.Repetitions : (int?)null,
            keyType = GetKeyType(keyEvent.KeyCode),
            controlKeys = GetControlKeysDetails(keyEvent.ControlKeys)
          };
        }

        private ControlKeys GetControlKeys(uint controlKeyState) {
          return new ControlKeys {
            CapsLockOn = (controlKeyState & Kernel32.CAPSLOCK_ON) != 0,
            EnhancedKey = (controlKeyState & Kernel32.ENHANCED_KEY) != 0,
            LeftAltPressed = (controlKeyState & Kernel32.LEFT_ALT_PRESSED) != 0,
            LeftCtrlPressed = (controlKeyState & Kernel32.LEFT_CTRL_PRESSED) != 0,
            NumLockOn = (controlKeyState & Kernel32.NUMLOCK_ON) != 0,
            RightAltPressed = (controlKeyState & Kernel32.RIGHT_ALT_PRESSED) != 0,
            RightCtrlPressed = (controlKeyState & Kernel32.RIGHT_CTRL_PRESSED) != 0,
            ScrollLockOn = (controlKeyState & Kernel32.SCROLLLOCK_ON) != 0,
            ShiftPressed = (controlKeyState & Kernel32.SHIFT_PRESSED) != 0
          };
        }

        private MouseButtons GetMouseButtons(uint buttonState) {
          return new MouseButtons {
            Left1Pressed = (buttonState & Kernel32.FROM_LEFT_1ST_BUTTON_PRESSED) != 0,
            Left2Pressed = (buttonState & Kernel32.FROM_LEFT_2ND_BUTTON_PRESSED) != 0,
            Left3Pressed = (buttonState & Kernel32.FROM_LEFT_3RD_BUTTON_PRESSED) != 0,
            Left4Pressed = (buttonState & Kernel32.FROM_LEFT_4TH_BUTTON_PRESSED) != 0,
            RightMostPressed = (buttonState & Kernel32.RIGHTMOST_BUTTON_PRESSED) != 0
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

            return "none";
        }

        private object GetControlKeysDetails(ControlKeys controlKeys) {
          return new {
            capsLockOn = controlKeys.CapsLockOn,
            enhancedKey = controlKeys.EnhancedKey,
            leftAltPressed = controlKeys.LeftAltPressed,
            leftCtrlPressed = controlKeys.LeftCtrlPressed,
            numLockOn = controlKeys.NumLockOn,
            rightAltPressed = controlKeys.RightAltPressed,
            rightCtrlPressed = controlKeys.RightCtrlPressed,
            scrollLockOn = controlKeys.ScrollLockOn,
            shiftPressed = controlKeys.ShiftPressed
          };
        }

        private object GetMouseButtonsDetails(MouseButtons buttons) {
          return new {
            left1Pressed = buttons.Left1Pressed,
            left2Pressed = buttons.Left2Pressed,
            left3Pressed = buttons.Left3Pressed,
            left4Pressed = buttons.Left4Pressed,
            rightMostPressed = buttons.RightMostPressed
          };
        }

        private WindowResizedEvent GetNextWindowResizedEvent() {
          uint eventsRead = 0;
          Kernel32.INPUT_RECORD inputRecord = new Kernel32.INPUT_RECORD();
          Kernel32.ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead);
          return new WindowResizedEvent {
            Rows = inputRecord.WindowBufferSizeEvent.dwSize.Y,
            Columns = inputRecord.WindowBufferSizeEvent.dwSize.X
          };
        }

        private WindowFocusEvent GetNextWindowFocusEvent() {
          uint eventsRead = 0;
          Kernel32.INPUT_RECORD inputRecord = new Kernel32.INPUT_RECORD();
          Kernel32.ReadConsoleInput(inputHandle, ref inputRecord, 1, ref eventsRead);
          return new WindowFocusEvent {
              Focused = inputRecord.FocusEvent.bSetFocus
          };
        }
    }

    internal enum InputEventType {
      None,
      WindowEvent,
      KeyEvent,
      MouseEvent
    }

    internal enum WindowEventType {
      None,
      Focus,
      Resized
    }

    internal enum KeyEventType {
      None,
      Pressed,
      Released
    }

    internal enum MouseEventType {
      None,
      MouseMoved,
      SingleClick,
      DoubleClick,
      HorizontalWheel,
      VerticalWheel,
      ButtonReleased
    }

    internal abstract class InputEventBase {
      public ControlKeys ControlKeys { get; set; }
    }

    internal abstract class KeyEventBase : InputEventBase {
      public int ScanCode { get; set; }
      public int KeyCode { get; set; }
      public char UnicodeChar { get; set; }
    }

    internal class PressedKeyEvent : KeyEventBase {
      public int Repetitions { get; set; }
    }

    internal class ReleasedKeyEvent : KeyEventBase { }

    internal abstract class MouseEventBase : InputEventBase {
      public MouseCursorLocation CursorLocation { get; set; }
    }

    internal class MouseButtonClickedEvent : MouseEventBase {
      public MouseButtons Buttons { get; set; }
    }

    internal class MouseEvent : MouseEventBase { }

    internal class MouseWheelEvent : MouseEventBase {
        public string WheelDirection { get; set; }
    }

    internal class MouseHorizontalWheelEvent : MouseWheelEvent { }

    internal class MouseVerticalWheelEvent : MouseWheelEvent { }

    internal class WindowResizedEvent {
      public int Rows { get; set; }
      public int Columns { get; set; }
    }

    internal class WindowFocusEvent {
      public bool Focused { get; set; }
    }

    internal class ControlKeys {
      public bool CapsLockOn { get; set; }
      public bool EnhancedKey { get; set; }
      public bool LeftAltPressed { get; set; }
      public bool LeftCtrlPressed { get; set; }
      public bool NumLockOn { get; set; }
      public bool RightAltPressed { get; set; }
      public bool RightCtrlPressed { get; set; }
      public bool ScrollLockOn { get; set; }
      public bool ShiftPressed { get; set; }
    }

    internal class MouseCursorLocation {
      public int Row { get; set; }
      public int Column { get; set; }
    }

    internal class MouseButtons {
      public bool Left1Pressed { get; set; }
      public bool Left2Pressed { get; set; }
      public bool Left3Pressed { get; set; }
      public bool Left4Pressed { get; set; }
      public bool RightMostPressed { get; set; }
    }
}
