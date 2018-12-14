using System.Drawing.Text;

namespace FontInstaller
{
    internal class OpenTypeFont : Font
    {
        internal override FontFormat FontFormat { get => FontFormat.OpenType; }

        public OpenTypeFont(string path)
        {
            Path = System.IO.Path.GetFullPath(path);
            FileName = System.IO.Path.GetFileName(path);

            var fontCollection = new PrivateFontCollection();
            fontCollection.AddFontFile(Path);
            var font = fontCollection.Families[0];

            Family = font.Name;
            OriginalName = font.GetName(0);
            Name = font.GetName(0);
        }
    }
}
