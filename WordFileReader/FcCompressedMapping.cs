// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordFileReader
{
    // p.267
    class FcCompressedMapping
    {
        static Dictionary<byte, ushort> map;

        internal static char GetChar(byte code)
        {
            if (code < 0x80)
                return (char)code;
            else
            {
                ushort value;
                if (map.TryGetValue(code, out value))
                    return (char)value;
                else
                    return '\x7F';
            }
        }

        internal static char[] GetChars(byte[] data, int offset, int count)
        {
            char[] result = new char[count];
            GetChars(data, offset, count, result, 0);
            return result;
        }

        internal static void GetChars(byte[] data, int offset, int count, char[] output, int outputOffset)
        {
            for (int i = 0; i < count; i++)
            {
                output[outputOffset + i] = GetChar(data[offset + i]);
            }
        }

        static FcCompressedMapping()
        {
            map = new Dictionary<byte, ushort>();
            map.Add(0x82, 0x201A);
            map.Add(0x83, 0x0192);
            map.Add(0x84, 0x201E);
            map.Add(0x85, 0x2026);
            map.Add(0x86, 0x2020);
            map.Add(0x87, 0x2021);
            map.Add(0x88, 0x02C6);
            map.Add(0x89, 0x2030);
            map.Add(0x8A, 0x0160);
            map.Add(0x8B, 0x2039);
            map.Add(0x8C, 0x0152);
            map.Add(0x91, 0x2018);
            map.Add(0x92, 0x2019);
            map.Add(0x93, 0x201C);
            map.Add(0x94, 0x201D);
            map.Add(0x95, 0x2022);
            map.Add(0x96, 0x2013);
            map.Add(0x97, 0x2014);
            map.Add(0x98, 0x02DC);
            map.Add(0x99, 0x2122);
            map.Add(0x9A, 0x0161);
            map.Add(0x9B, 0x203A);
            map.Add(0x9C, 0x0153);
            map.Add(0x9F, 0x0178);
        }
    }
}
