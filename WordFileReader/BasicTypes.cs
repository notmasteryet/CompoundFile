// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WordFileReader
{
    class Clx
    {
        //p.247
        internal Prc[] RgPrc;
        internal Pcdt Pcdt;
    }

    class Prc
    {
        internal const byte DefaultClxt = 0x01;
        //p.435
        internal byte clxt; // 1
        internal short cbGrpprl;
        internal Prl[] GrpPrl;
    }

    class Prl
    {
        //p.31
        internal Sprm sprm;
        internal byte[] operand;

        internal Prl() { }

        internal Prl(ushort sprm, byte[] operand)
            : this(new Sprm(sprm), operand) { }

        internal Prl(Sprm sprm, byte[] operand)
        {
            this.sprm = sprm;
            this.operand = operand;
        }
    }

    struct Sprm
    {
        internal ushort ispmd { get { return (ushort)(sprm & 0x1FF); } }
        internal bool fSpec { get { return (sprm & 0x200) != 0; } }
        internal byte sgc { get { return (byte)((sprm >> 10) & 0x07); } }
        internal byte spra { get { return (byte)((sprm >> 13) & 0x07); } }

        internal ushort sprm;

        internal Sprm(ushort sprm)
        {
            this.sprm = sprm;
        }

        public override string ToString()
        {
            string sprmName;
            if(SinglePropertyModifiers.map.TryGetValue(sprm, out sprmName))
                return sprmName;
            else
                return "sprm: 0x" + sprm.ToString("X4");
        }
    }

    class Pcdt
    {
        internal const byte DefaultClxt = 0x02;
        // p.415
        internal byte clxt; // 2
        internal uint lcb;
        internal PlcPcd PlcPcd;
    }

    class PlcPcd
    {
        internal static int CalcLength(uint size) { return (int)((size - 4) / (4 + Pcd.Size)); } // p.26

        internal uint[] CPs;
        internal Pcd[] Pcds;
    }

    class Pcd
    {
        internal const int Size = 8;
        // p.415
        internal bool fNoParaLast;
        internal bool fR1;
        internal bool fDirty;
        internal FcCompressed fc; 
        internal Prm prm; 
    }

    struct FcCompressed
    {
        internal int fc;
        internal bool fCompressed;
    }

    struct Prm
    {
        internal ushort prm;

        internal bool fComplex { get { return (prm & 1) != 0; } }

        //Prm0
        internal byte isprm { get { return (byte)((prm >> 1) & 0x7F); } }
        internal byte val { get { return (byte)((prm >> 8) & 0xFF); } }

        //Prm1
        internal byte igrpprl { get { return (byte)((prm >> 1) & 0x7FFF); } }
    }

    class PlcBtePapx
    {
        // p.201
        internal static int GetLength(uint size) { return (int)((size - 4) / (4 + PnFkpPapx.Size)); }

        internal uint[] aFC;
        internal PnFkpPapx[] aPnBtePapx;
    }

    struct PnFkpPapx
    {
        // p.434
        internal const int Size = 4;

        internal uint pn;
    }

    class PapxFkp
    {
        // p.413 (512 bytes)
        internal const int Size = 512;

        internal uint[] rgfc;
        internal KeyValuePair<BxPap, PapxInFkps>[] rgbx;
        internal byte cpara;
    }

    class PapxInFkps
    {
        // p.414 
        internal byte cb;
        internal GrpPrlAndIstd grpprlInPapx;
    }

    class GrpPrlAndIstd
    {
        // p.372
        internal ushort istd;
        internal Prl[] grpprl;
    }

    struct BxPap
    {
        // p.236 (13 bytes)
        internal const int Size = 13;

        internal byte bOffset;

    }

    class PlcBteChpx
    {
        //p.201 
        internal static int GetLength(uint size) { return (int)((size - 4) / (4 + PnFkpChpx.Size)); }

        internal uint[] aFC;
        internal PnFkpChpx[] aPnBteChpx;
    }

    struct PnFkpChpx
    {
        // p.433
        internal const int Size = 4;

        internal uint pn;
    }

    class ChpxFkp
    {
        // p.241
        internal const int Size = 512;

        internal uint[] rgfc;
        internal KeyValuePair<byte, Chpx>[] rgb;
        internal byte crun;
    }

    class Chpx
    {
        // p.241
        internal byte cb;
        internal Prl[] grpprl;
    }

    class STSH
    {
        // p.477
        internal Stshi stshi;
        internal STD[] rglpstd;
    }

    class Stshi
    {
        // p.478
        internal Stshif stshif;
        internal short ftcBi;
        internal StshiLsd StshiLsd;
    }

    class Stshif
    {
        // p.479
        internal const int Size = 18;

        internal ushort cstd;
        internal ushort cbSTDBaseInFile;
        internal ushort stiMaxWhenSaved;
        internal ushort istdMaxFixedWhenSaved;
        internal ushort nVerBuiltInNamesWhenSaved;
        internal short ftcAsci;
        internal short ftcFE;
        internal short ftcOther;
    }

    class StshiLsd
    {
        // p.480
        internal ushort cbLSD;
        internal LSD[] mpstiilsd;
    }

    class LSD
    {
        // p.390
        internal const int Size = 4;

        internal bool fLocked;
        internal bool fSemiHidden;
        internal bool fUnhideWhenUsed;
        internal bool fQFormat;
        internal ushort iPriority;
    }

    class STD
    {
        // LP - p.387
        //p.469
        internal Stdf stdf;
        internal Xstz xstzName;
        internal GrLPUpxSw grLPUpxSw;
    }

    class Stdf
    {
        //p.470 (10-18 bytes)
        internal StdfBase stdfBase;
        internal StdfPost2000OrNone StdfPost2000OrNone;
    }

    class StdfBase
    {
        //p.470
        internal const int Size = 10;
        internal const ushort IstdNull = 0x0FFF;

        internal ushort sti;
        internal byte stk;
        internal ushort istdBase;
        internal byte cupx;
        internal ushort istdNext;
        internal ushort bchUpe;
        internal GRFSTD grfstd;
    }

    class GRFSTD
    {
        // p.371
        internal bool fAutoRedef;
        internal bool fHidden;
        internal bool f97LidsSet;
        internal bool fCopyLang;
        internal bool fPersonalCompose;
        internal bool fPersonalReply;
        internal bool fPersonal;
        internal bool fSemiHidden;
        internal bool fLocked;
        internal bool fUnhideWhenUsed;
        internal bool fQFormat;
    }

    class StdfPost2000OrNone
    {
        internal const int Size = 8;

        internal ushort istdLink;
        internal bool fHasOriginalStyle;
        internal uint rsid;
        internal ushort iPriority;
    }

    class Xstz
    {
        //p.539
        internal Xst xst;

        public override string ToString()
        {
            return xst.ToString();
        }
    }

    class Xst
    {
        //p.539
        internal ushort cch;
        internal char[] rgtchar;

        public override string ToString()
        {
            return new String(rgtchar);
        }
    }

    abstract class GrLPUpxSw
    {
        // p.372
        internal const int StkParaGRLPUPXStkValue = 1; // paragraph
        internal const int StkCharGRLPUPXStkValue = 2; // character
        internal const int StkTableGRLPUPXStkValue = 3; // table
        internal const int StkListGRLPUPXStkValue = 4; // numbering
    }

    class StkParaGRLPUPX : GrLPUpxSw
    {
        // p.475

        internal UpxPapx upxPapx;
        internal UpxChpx upxChpx;
        internal StkParaUpxGrLPUpxRM StkParaLPUpxGrLPUpxRM;
    }

    class StkCharGRLPUPX : GrLPUpxSw
    {
        //p.473

        internal UpxChpx upxChpx;
        internal StkCharUpxGrLPUpxRM StkCharLPUpxGrLPUpxRM;
    }

    class StkTableGRLPUPX : GrLPUpxSw
    {
        //p.476
        internal UpxTapx upxTapx;
        internal UpxPapx upxPapx;
        internal UpxChpx upxChpx;
    }

    class StkListGRLPUPX : GrLPUpxSw
    {
        //p.474
        internal UpxPapx upxPapx;
    }

    class StkParaUpxGrLPUpxRM
    {
        //LP p.475
        // p.476
        internal UpxRm upxRm;
        internal UpxPapx upxPapxRM;
        internal UpxChpx upxChpxRM;
    }

    class StkCharUpxGrLPUpxRM
    {
        // p.474
        internal UpxRm upxRm;
        internal UpxChpx upxChpxRM;
    }

    class UpxChpx
    {
        // LP - p.388
        // p.529
        internal Prl[] grpprlChpx;
    }

    class UpxPapx
    {
        //LP - p.389
        // p.530
        internal ushort istd;
        internal Prl[] grpprlPapx;
    }

    class UpxRm
    {
        internal const int Size = 8;

        internal DTTM date;
        internal short ibstAuthor;
    }

    class UpxTapx
    {
        internal Prl[] grpprlTapx;
    }

    struct DTTM
    {
        internal byte mint;
        internal byte hr;
        internal byte dom;
        internal byte mon;
        internal ushort year;
        internal byte wdy;
    }
}
