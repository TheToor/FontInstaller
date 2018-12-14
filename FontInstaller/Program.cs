using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FontInstaller
{
    class Program
    {
        internal static string Folder = "Fonts";
        internal static bool TypeCheck = false;
        internal static bool CheckOnly = false;
        internal static bool IgnoreUnderlineInFilenames = false;
        internal static bool MatchFontStyleToFile = false;

        internal const int ExitOK = 0;

        internal static int ExitError = 1;
        internal static int ExitErrorFolderNotSpecified = 2;
        internal static int ExitErrorFolderNotFound = 3;

        [STAThread]
        static void Main(string[] args)
        {
            var arguments = args.ToList();

            if (arguments.Count > 0)
            {
                if (arguments.Any(a => a == "--help" || a == "-h"))
                {
                    Console.WriteLine("--help or -h | Shows this help");
                    Console.WriteLine("--folder <folder> or -f <folder> | Optional: Install fonts from the specified folder instead of the default 'Fonts' folder");
                    Console.WriteLine("--no-error or -ne | Optional: Always exits with error code 0");
                    Console.WriteLine("--type-check or -tc | Optional: Check font type. Tries to install a Type1 font if there is only a OpenType font installed");
                    Console.WriteLine("--no-install or -ni | Optional: Only shows output / runs the checks. Does not install the fonts");
                    Console.WriteLine("--ignore-underline or -iu | Optional: Ignores '_' in font names on font file checkes. Warning: Can lead to false positive detections");
                    Console.WriteLine("--match-style-to-file or -mstf | Optional: Checks on installed fonts if the file exists in the windows font folder with that style. If not tries installs the font");
                    Environment.Exit(ExitOK);
                }

                if (arguments.Any(a => a == "--no-error" || a == "-ne"))
                {
                    Console.WriteLine("Warning: Errors disabled");
                    ExitError = ExitOK;
                    ExitErrorFolderNotFound = ExitOK;
                    ExitErrorFolderNotSpecified = ExitOK;
                }

                if (arguments.Any(a => a == "--type-check" || a == "-tc"))
                    TypeCheck = true;

                if (arguments.Any(a => a == "--no-install" || a == "-ni"))
                    CheckOnly = true;

                if (arguments.Any(a => a == "--ignore-underline" || a == "-iu"))
                    IgnoreUnderlineInFilenames = true;

                if (arguments.Any(a => a == "--match-style-to-file" || a == "-mstf"))
                    MatchFontStyleToFile = true;

                if (arguments.Any(a => a == "--folder" || a == "-f"))
                {
                    var index = arguments.IndexOf("--folder");
                    if (index == -1)
                        index = arguments.IndexOf("-f");

                    if (arguments.Count >= (index + 1))
                        Environment.Exit(ExitErrorFolderNotSpecified);

                    Folder = arguments.ElementAt(index + 1);
                    if (string.IsNullOrEmpty(Folder))
                        Environment.Exit(ExitErrorFolderNotFound);

                    Folder = Path.GetFullPath(Folder);
                    Console.WriteLine($"Installing fonts from {Folder}");
                }
            }

            InstallFonts(Folder);
        }

        private static void InstallFonts(string directory)
        {
            if (!Directory.Exists(directory))
                directory = "Fonts";

            if (!Directory.Exists(directory))
                Environment.Exit(ExitError);

            var fontsToInstall = new List<Font>();
            var files = Directory.GetFiles("Fonts");
            foreach (var file in files)
            {
                var extension = file.Substring(file.Length - 3, 3).ToUpper();
                Font font = null;

                if (extension == "PFM")
                    font = new Type1Font(file);
                else if (extension == "TTF" || extension == "OTF")
                    font = new OpenTypeFont(file);

                if (font == null)
                    continue;

                if (!font.Installed())
                    fontsToInstall.Add(font);
                else
                    Console.WriteLine($"({font.FileName})\t{font.Name}-{font.Family} already installed");
            }

            if (fontsToInstall.Count > 0)
            {
                var sorted = from font in fontsToInstall
                             orderby font.Name.Length ascending
                             select font;

                foreach (var font in sorted)
                {
                    font.Install();
                    Console.WriteLine($"({font.FileName})\t{font.Name}-{font.Family} installed. Check: {font.Installed()}");
                }
            }

            Environment.Exit(ExitOK);
        }
    }
}
