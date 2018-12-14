using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace FontInstaller
{
    enum FontInstallFlags : int
    {
        // Installs font directly where it is laying currently
        LinkAndInstall,
        // Copys font to windows fonts directory then installs it
        CopyAndInstall
    }

    internal static class FontInstaller
    {
        [DllImport("fontext.dll", CharSet = CharSet.Auto)]
        internal static extern void InstallFontFile(IntPtr hwnd, string filePath, FontInstallFlags flags);
    }

    enum FontFormat
    {
        Unknown,
        Type1,
        OpenType
    }

    internal abstract class Font
    {
        internal abstract FontFormat FontFormat { get; }

        internal string Path { get; set; }
        internal string FileName { get; set; }

        internal string Family { get; set; }
        internal string Name { get; set; }
        internal string OriginalName { get; set; }

        public bool Installed()
        {
            var detectedTypes = new List<FontFormat>();
            var windowsFonts = new InstalledFontCollection();
            var families = windowsFonts.Families;
            var isInstalled = families.Any(f => f.Name == Name);

            // Perform registry checks only in this mode so we have a file-name/type
            if (Program.TypeCheck)
                isInstalled = false;

            if (!isInstalled)
            {
                var fileName = FileName;
                if (Program.IgnoreUnderlineInFilenames)
                {
                    //Just remove everything that is after the first _
                    var index = fileName.IndexOf('_');
                    if (index != -1)
                        fileName = fileName.Substring(0, index);
                }

                var registryFonts = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts");
                if (registryFonts != null)
                {
                    var fontNames = registryFonts.GetValueNames();
                    isInstalled = fontNames
                        .ToList()
                        .Any(f => f.Contains(OriginalName) || f.Contains(Name));

                    if(!isInstalled)
                    {
                        var fontValues = new List<string>();
                        fontNames.ToList().ForEach(f => fontValues.Add(registryFonts.GetValue(f) as string));

                        if (!Program.IgnoreUnderlineInFilenames)
                            isInstalled = fontValues.Any(value => value == fileName);
                        else
                            isInstalled = fontValues.Any(value => value.StartsWith(fileName));
                    }

                    if (isInstalled)
                        detectedTypes.Add(FontFormat.OpenType);
                }

                //In this mode we need to know all installed types
                if (!isInstalled || Program.TypeCheck)
                {
                    var type1Fonts = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Type 1 Installer\Type 1 Fonts");
                    if (type1Fonts != null)
                    {
                        var fontNames = type1Fonts.GetValueNames();
                        isInstalled = fontNames
                            .ToList()
                            .Any(f => f.Contains(OriginalName) || f.Contains(Name));

                        if (!isInstalled)
                        {
                            var fontValues = new List<string[]>();
                            fontNames.ToList().ForEach(f => fontValues.Add((string[])type1Fonts.GetValue(f)));
                            if (!Program.IgnoreUnderlineInFilenames)
                                isInstalled = fontValues.Any(value => value.Contains(fileName));
                            else
                                isInstalled = fontValues.Any(value => value.Any(f => f.StartsWith(fileName)));
                        }

                        if (isInstalled)
                            detectedTypes.Add(FontFormat.Type1);
                    }
                }

                // We can detect a OpenType font but due the TypeCheck we could override the 'isInstalled' variable with an invalid result
                // This happens when we only have a OpenType but no Type1 font
                if (Program.TypeCheck && !isInstalled && detectedTypes.Count > 0)
                    isInstalled = true;
            }

            if (isInstalled)
            {
                var collection = new PrivateFontCollection();
                collection.AddFontFile(Path);
                if (collection.Families.Length > 0)
                {
                    var font = collection.Families[0];

                    var fileName = FileName.ToLower();
                    foreach (FontStyle style in Enum.GetValues(typeof(FontStyle)))
                    {
                        var styleName = style.ToString().ToLower();
                        if (!fileName.Contains(styleName))
                            continue;
                        
                        var styleFound = font.IsStyleAvailable(style);
                        if (!font.IsStyleAvailable(style) || (Program.MatchFontStyleToFile && !File.Exists($"C:\\Windows\\Fonts\\{FileName}")))
                        {
                            isInstalled = false;
                            Console.WriteLine($"Info: Style not installed. Forcing installation...");
                        }
                    }
                }
                else
                    Console.WriteLine($"Warning: {FileName} could not be parsed to a FontFamily");

                if(Program.TypeCheck && !detectedTypes.Contains(FontFormat))
                {
                    Console.WriteLine($"Info: Other font type installed. Forcing installation... ({FontFormat})");
                    if (detectedTypes.Count == 1)
                        Console.WriteLine($"Other detected font format: {detectedTypes[0]}");
                    else
                        Console.WriteLine($"Other detected font formats: {detectedTypes.ConvertAll((i) => i.ToString()).Aggregate((i, j) => i + ", " + j)}");
                    isInstalled = false;
                }
            }

            return isInstalled;
        }

        public bool Install()
        {
            if (Installed())
            {
                Console.WriteLine($"Failed to install font {Name}-{Family}: Can't install a font if the font is already installed");
                return false;
            }

            try
            {
                FontInstaller.InstallFontFile(IntPtr.Zero, Path, FontInstallFlags.CopyAndInstall);
                
                // Alternative method
                //var shell = new Shell32.Shell();
                //var fontsDirectory = shell.NameSpace(0x14);
                //fontsDirectory.CopyHere(Path);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to install font {Name}-{Family}: {ex}");
                return false;
            }
        }
    }
}
