// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Text;

namespace CompoundFile
{
    static class FATSectorIds
    {
        internal const uint MAXREGSECT = 0xFFFFFFFA;
        internal const uint DIFSECT = 0xFFFFFFFC;
        internal const uint FATSECT = 0xFFFFFFFD;
        internal const uint ENDOFCHAIN = 0xFFFFFFFE;
        internal const uint FREESECT = 0xFFFFFFFF;
    }

    static class DirectoryStreamIds
    {
        internal const uint MAXREGSID = 0xFFFFFFFA;
        internal const uint NOSTREAM = 0xFFFFFFFF;
    }

    static class DirectoryObjectTypes
    {
        internal const byte Unknown = 0x00;
        internal const byte Storage = 0x01;
        internal const byte Stream = 0x02;
        internal const byte RootStorage = 0x05;
    }

    static class TreeColors
    {
        internal const byte Red = 0x00;
        internal const byte Black = 0x01;
    }
}
