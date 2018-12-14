using System.IO;
using System.Text;

namespace FontInstaller
{
    internal class Type1Font : Font
    {
        internal override FontFormat FontFormat { get => FontFormat.Type1; }

        public Type1Font(string path)
        {
            Path = System.IO.Path.GetFullPath(path);
            FileName = System.IO.Path.GetFileName(path);
             
            using (var file = File.OpenRead(path))
            {
                using (var binaryReader = new BinaryReader(file))
                {
                    binaryReader.ReadChars(210); // Skip the first 120 characters
                    Family = ReadBytes(binaryReader);
                    OriginalName = ReadBytes(binaryReader);
                    Name = OriginalName.Replace('-', ' ');
                }
            }
        }

        private string ReadBytes(BinaryReader binaryReader)
        {
            var stringBuilder = new StringBuilder();
            do
            {
                stringBuilder.Append(binaryReader.ReadChar());
            }
            while (binaryReader.PeekChar() != 0);
            binaryReader.ReadChar();

            return stringBuilder.ToString();
        }
    }
}
