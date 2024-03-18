using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Threading;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using static VoiceAttack.Program;

namespace VoiceAttack
{
    class Program
    {
        /// <summary>
        /// Définition d'une entrée clavier
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
        /// Défintion d'un entrée souris
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
        /// Défintion d'une entrée hardware
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct HardwareInput
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }

        /// <summary>
        /// Union des différentes entrées pour réaliser des actions solicitants différent types d'entrées
        /// </summary>

        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)] public MouseInput mi;
            [FieldOffset(0)] public KeyboardInput ki;
            [FieldOffset(0)] public HardwareInput hi;
        }

        /// <summary>
        /// Définition généralisé d'une entrée
        /// </summary>
        public struct Input
        {
            public int type;
            public InputUnion u;
        }

        /// <summary>
        /// Id des différents types d'entrées
        /// </summary>
        [Flags]
        public enum InputType
        {
            Mouse = 0,
            Keyboard = 1,
            Hardware = 2
        }

        /// <summary>
        /// Enum pour faciliter l'utilisation des type d'évenements claviers
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
        /// Enum pour faciliter l'utilisation des types d'évenements souris
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
        /// Création d'un input clavier en mode Press and Release.
        /// </summary>
        /// <param name="key">Le ScanCode AZERTY de la touche clavier</param>
        /// <param name="duration">Le temps d'appui sur la touche</param>
        /// <returns>Un Input clavier</returns>
        public static Input[] CreateKeyboardInput(ushort key)
        {
            return new Input[] {
                new() {
                    type = (int)InputType.Keyboard,
                    u = new InputUnion
                    {
                        ki = new KeyboardInput
                        {
                            wVk = 0,
                            wScan = key, // Z
                            dwFlags = (uint)(KeyEventF.KeyDown | KeyEventF.Scancode),
                            dwExtraInfo = GetMessageExtraInfo(),
                            time = 0

                        }
                    },

                },
                new()
                {
                    type = (int)InputType.Keyboard,
                    u = new InputUnion
                    {
                        ki = new KeyboardInput
                        {
                            wVk = 0,
                            wScan = key, // Z
                            dwFlags = (uint)(KeyEventF.KeyUp | KeyEventF.Scancode),
                            dwExtraInfo = GetMessageExtraInfo(),
                            time = 0
                        }
                    },     
                }
            };
        }

        /// <summary>
        /// Variable pour stocker l'interpréteur vocal courant afin de pouvoir le détruire pour en construire un nouveau
        /// </summary>
        public static SpeechRecognitionEngine engine;

        /// <summary>
        /// Dictionnaire qui représente les macros crées. En clé le nom de la macro et en valeur une séquence d'Input qui sera éxecutée.
        /// </summary>
        public static Dictionary<String, Input[][]> macros = new Dictionary<string, Input[][]>();

        /// <summary>
        /// Definit l'état de l'application
        /// </summary>
        public static bool running = true;

        static void Main(string[] args)
        {
            Initialize_Inputs("Default"); // Première initialisation par défaut sans profil.
            InitialStart(); // Création de l'interpréteur d'acceuil et de selection de profil
            
            while (running)
            {
                Console.ReadLine(); // Garde la console ouverte.
            }
           

        }

        /// <summary>
        /// Initialise les macros selon le profil choisi
        /// </summary>
        /// <param name="profile">Le profil d'ont ont veux charger les macros</param>
        static void Initialize_Inputs(String profile)
        {
            macros = new();
            switch (profile)
            {
                case "Default":

                    macros.Add("test", new Input[][] { CreateKeyboardInput(0x11) });
                    break;

                case "SC":
                    macros.Add("demarage", new Input[][] { CreateKeyboardInput(0x13) }); // TODO: Ajouter un Enum pour les touches clavier
                    break;

            //AJOUTEZ DES COMMANDES ICI
            }
        }

        /// <summary>
        /// Charge le Dictionnaire vocal (grammaire) du profil par défaut.
        /// </summary>
        /// <param name="recognizer">Le moteur de reconnaissance vocal qui utilisera cette grammaire</param>
        static void SetupSpeechRecognizer(SpeechRecognitionEngine recognizer)
        {
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            Choices choices = new Choices();

            // COMMANDES PREDEFINIES

            choices.Add("Star Citizen");
            choices.Add("Hook");
            choices.Add("stop");


            grammarBuilder.Append(choices);
            Grammar grammar = new Grammar(grammarBuilder);

            recognizer.LoadGrammar(grammar);
        }

        /// <summary>
        /// Charge le Dictionnaire vocal (grammaire) du profil en paramètre.
        /// </summary>
        /// <param name="recognizer">Le moteur de reconnaissance vocal qui utilisera cette grammaire</param>
        /// <param name="profile">Le profil d'ont ont veux charger la grammaire</param>
        private static void SetupSpeechRecognizer(SpeechRecognitionEngine recognizer, string profile)
        {
            GrammarBuilder grammarBuilder = new GrammarBuilder();
            Choices choices = new Choices();

            switch (profile)
            {
                // COMMANDES PREDEFINIES SELON LE PROFIL

                case "SC":
                    choices.Add("démare le vaisseau");
                    choices.Add("pose toi");
                    break;
            }
            grammarBuilder.Append(choices);
            Grammar grammar = new Grammar(grammarBuilder);
            recognizer.LoadGrammar(grammar);
        }

        /// <summary>
        /// Execute les inputs d'une macro
        /// </summary>
        /// <param name="key">Le nom de la macro a executer</param>
        static void ExecuteCommand(string key)
        {
            if (macros.ContainsKey(key))
            {
                foreach (Input[] sequence in macros[key])
                {
                    foreach (Input input in sequence)
                    {
                        Input[] single = { input };
                        if (input.type == (int)InputType.Keyboard && input.u.ki.dwFlags == 8)
                        {
                            // Pour KeyDown
                            uint result = SendInput(1, single, Marshal.SizeOf(typeof(Input)));
                          
                            if (result == 0)
                            {
                                int error = Marshal.GetLastWin32Error();
                                Console.WriteLine($"SendInput failed with error code {error}");
                            }
                            else
                            {
                                // DElAY pour la touche
                                Thread.Sleep(1000);
                                Console.WriteLine($"Sent key-down event successfully");
                            }                           
                        }
                        else
                        {
                           // KEYUP
                            uint result = SendInput(1, single, Marshal.SizeOf(typeof(Input)));
                            if (result == 0)
                            {
                                int error = Marshal.GetLastWin32Error();
                                Console.WriteLine($"SendInput failed with error code {error}");
                            }
                            else
                            {
                                Console.WriteLine($"Sent input event successfully");
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine($"Macro '{key}' not found");
            }
        }


        /// <summary>
        /// Prend le résultat de la reconnaissance vocale et permet de choisir un profil.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void Recognizer_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            Console.WriteLine($"Recognized: {e.Result.Text}");
            switch (e.Result.Text.ToLower())
            {
                case "star citizen":
                    Console.WriteLine("Switching to Star Citizen profile");
                    Rewrite("SC");
                    break;

                case "hook":
                    Console.WriteLine("sent");
                    ExecuteCommand("test");
                    Console.WriteLine("end");
                    break;
                    // Add more commands here
            }
        }

        /// <summary>
        /// Fait fonctionner les commandes vocales pour Star Citizen
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void StarCitizen_SpeechProcessor(object sender, SpeechRecognizedEventArgs e)
        {
            Console.WriteLine($"Recognized: {e.Result.Text}");

        }

        /// <summary>
        /// Détruit l'interpréteur courant et en reconstruit un adapté au profil
        /// </summary>
        /// <param name="game">Le jeu pour lequel on souhaite charger un interpréteur vocal</param>
        private static void Rewrite(String game)
        {
            engine.Dispose();
            switch (game)
            {
                case "SC":
                    StartEngine("SC",StarCitizen_SpeechProcessor);
                    break;
            }
        }

        /// <summary>
        /// Lance un interpréteur pour le profil selectionné
        /// </summary>
        /// <param name="profile">Le profil pour lequel on charge l'interpréteur</param>
        /// <param name="speechProcessor">La méthode qui s'occupera d'analyser les résultats de la reconnaissance vocale</param>
        private static void StartEngine(String profile,  System.EventHandler<System.Speech.Recognition.SpeechRecognizedEventArgs> speechProcessor)
        {
            using (SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine())
            {
                engine = recognizer;
                SetupSpeechRecognizer(recognizer,profile);
                new System.Globalization.CultureInfo("fr-FR");

                recognizer.SpeechRecognized += speechProcessor;
                recognizer.SetInputToDefaultAudioDevice();
                recognizer.RecognizeAsync(RecognizeMode.Multiple);

                Console.WriteLine("Voice Attack activated with profile " + profile + ". Listening for commands...");

                Console.ReadLine();
            }
        }

        /// <summary>
        /// Mise en place de l'interpréteur de selection de profil
        /// </summary>
        static void InitialStart()
        {
            using (SpeechRecognitionEngine recognizer = new SpeechRecognitionEngine())
            {
                engine = recognizer;
                SetupSpeechRecognizer(recognizer);
                new System.Globalization.CultureInfo("fr-FR");

                recognizer.SpeechRecognized += Recognizer_SpeechRecognized;
                recognizer.SetInputToDefaultAudioDevice();
                recognizer.RecognizeAsync(RecognizeMode.Multiple);

                Console.WriteLine("Voice Attack activated. Listening for commands...");

                Console.ReadLine();
            }
        }

 
        
        // ENUM DES KEYCODES
        private enum VirtualKeys : int
        {
            VK_UP = 0x26,
            VK_DOWN = 0x28,
            VK_LEFT = 0x25,
            VK_RIGHT = 0x27,
            VK_CTRL = 0x11
            // ICI POUR EN AJOUTER
        }

    }
}
