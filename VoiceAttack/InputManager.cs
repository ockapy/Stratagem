using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace VoiceAttack
{
    public static class InputManager
    {
        /// <summary>
        /// Structure pour représenter une entrée clavier
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct KeyboardInput
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        /// <summary>
        /// Structure pour représenter une entrée souris
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct MouseInput
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        /// <summary>
        /// Structure pour représenter une entrée hardware
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct HardwareInput
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        /// <summary>
        /// Union des différentes entrées (clavier, souris, hardware)
        /// </summary>
        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)] public MouseInput mi;
            [FieldOffset(0)] public KeyboardInput ki;
            [FieldOffset(0)] public HardwareInput hi;
        }

        /// <summary>
        /// Structure générale d'une entrée pour SendInput
        /// </summary>
        public struct Input
        {
            public int type;
            public InputUnion u;
        }

        /// <summary>
        /// Enumération des types d'entrées pour SendInput
        /// </summary>
        public enum InputType
        {
            Mouse = 0,
            Keyboard = 1,
            Hardware = 2
        }

        /// <summary>
        /// Enum pour les événements du clavier
        /// </summary>
        [Flags]
        public enum KeyEventF
        {
            KeyDown = 0x0000,
            ExtendedKey = 0x0001,
            KeyUp = 0x0002,
            Unicode = 0x0004,
            Scancode = 0x0008
        }

        /// <summary>
        /// Enum pour les événements de la souris
        /// </summary>
        [Flags]
        public enum MouseEventF
        {
            Absolute = 0x8000,
            HWheel = 0x01000,
            Move = 0x0001,
            MoveNoCoalesce = 0x2000,
            LeftDown = 0x0002,
            LeftUp = 0x0004,
            RightDown = 0x0008,
            RightUp = 0x0010,
            MiddleDown = 0x0020,
            MiddleUp = 0x0040,
            VirtualDesk = 0x4000,
            Wheel = 0x0800,
            XDown = 0x0080,
            XUp = 0x0100
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, Input[] pInputs, int cbSize);

        [DllImport("user32.dll")]
        private static extern IntPtr GetMessageExtraInfo();

        /// <summary>
        /// Crée un tableau d'entrées clavier (appui et relâchement de touche)
        /// </summary>
        /// <param name="key">Le ScanCode de la touche</param>
        /// <returns>Un tableau d'entrées clavier avec appui et relâchement de touche</returns>
        public static Input[] CreateKeyboardInput(ushort key)
        {
            return new Input[]
            {
                new Input
                {
                    type = (int)InputType.Keyboard,
                    u = new InputUnion
                    {
                        ki = new KeyboardInput
                        {
                            wVk = 0,
                            wScan = key,
                            dwFlags = (uint)(KeyEventF.KeyDown | KeyEventF.Scancode),
                            time = 0,
                            dwExtraInfo = GetMessageExtraInfo()
                        }
                    }
                },
                new Input
                {
                    type = (int)InputType.Keyboard,
                    u = new InputUnion
                    {
                        ki = new KeyboardInput
                        {
                            wVk = 0,
                            wScan = key,
                            dwFlags = (uint)(KeyEventF.KeyUp | KeyEventF.Scancode),
                            time = 0,
                            dwExtraInfo = GetMessageExtraInfo()
                        }
                    }
                }
            };
        }

        /// <summary>
        /// Envoie les entrées via SendInput et gère les erreurs si elles surviennent
        /// </summary>
        /// <param name="inputs">Tableau d'entrées à envoyer</param>
        public static void ExecuteInput(Input[] inputs)
        {
            foreach (var input in inputs)
            {
                uint result = SendInput(1, new Input[] { input }, Marshal.SizeOf(typeof(Input)));
                if (result == 0)
                {
                    int error = Marshal.GetLastWin32Error();
                    Console.WriteLine($"SendInput failed with error code {error}");
                }
                else
                {
                    Console.WriteLine("Sent input event successfully");
                }

                Thread.Sleep(1000); // Delay for the key press
            }
        }
    }
}
