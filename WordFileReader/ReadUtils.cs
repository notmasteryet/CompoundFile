// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WordFileReader
{
    static class ReadUtils
    {
        internal const int ByteSize = 1;
        internal const int WordSize = 2;
        internal const int DWordSize = 4;

        internal static byte ReadByte(Stream s)
        {
            int b = s.ReadByte();
            if (b < 0) throw new WordFileReaderException("Unexpected EOF");
            return (byte)b;
        }

        internal static byte ReadByte(Stream s, ref int read)
        {
            byte b = ReadByte(s);
            ++read;
            return b;
        }

        internal static byte[] ReadExact(Stream s, int count)
        {
            byte[] data = new byte[count];
            int read = s.Read(data, 0, count);
            if (read != count) throw new WordFileReaderException("Unexpected EOF");
            return data;
        }

        internal static byte[] ReadExact(Stream s, int count, ref int read)
        {
            byte[] data = ReadExact(s, count);
            read += data.Length;
            return data;
        }

        internal static void Skip(Stream s, int count)
        {
            s.Seek(count, SeekOrigin.Current);
        }
    }
}
