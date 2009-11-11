// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CompoundFile;
using WordFileReader;

namespace WordFileDump
{
    class Program
    {
        static void Main(string[] args)
        {
            if(args.Length == 0)
            {
                Console.WriteLine("USAGE: WordFileDrop.exe <doc-file>");
                return;
            }

            string filename = args[0];
            using (CompoundFileSystem system = new CompoundFileSystem(filename))
            {
                WordDocument doc = new WordDocument();
                doc.Load(system);

                foreach (Paragraph p in doc.Paragraphs)
                {
                    string text = p.GetText();
                    Console.WriteLine("{0}/{1}, {3}: {2}",
                        p.Offset, p.Length, Escape(text),
                        p.Style.Name);
                }
                Console.WriteLine();
                foreach (StyleDefinition s in doc.StyleDefinitions)
                {
                    Console.WriteLine("Style {0}: {1}", s.Name, s.IsTextStyle);

                }
                Console.WriteLine();
                foreach (CharacterFormatting f in doc.Formattings)
                {
                    string text = f.GetText();
                    StyleDefinition style = f.Style;
                    Console.WriteLine("Format {0}/{1}, {3}: {2}",
                        f.Offset, f.Length, Escape(text),
                        style != null ? style.Name : "-");
                }
            }
        }

        static string Escape(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char ch in s)
            {
                if (ch < ' ' || ch == '\\')
                    sb.Append(@"\u").Append(((int)ch).ToString("X4"));
                else
                    sb.Append(ch);
            }
            return sb.ToString();
        }
    }

}