// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WordFileReader
{
    static class BasicTypesReader
    {
        internal static Clx ReadClx(Stream s, int length)
        {
            int position = 0;
            List<Prc> prcs = new List<Prc>();
            if (length < 1)
                throw new WordFileReaderException("Invalid Clx: length");
            byte clxt = ReadUtils.ReadByte(s);
            position++;
            while (clxt == Prc.DefaultClxt)
            {
                int read;
                prcs.Add(ReadPrcWithoutCltx(s, length - position, out read));
                position += read;

                if (position >= length)
                    throw new WordFileReaderException("Invalid Clx: length");
                clxt = ReadUtils.ReadByte(s);
                position++;
            }
            if (clxt != Pcdt.DefaultClxt)
                throw new WordFileReaderException("Invalid Pcdt: signature");
            Pcdt pcdt = ReadPcdtWithoutClxt(s, length - position);

            Clx clx = new Clx();
            clx.RgPrc = prcs.ToArray();
            clx.Pcdt = pcdt;
            return clx;
        }

        private static Pcdt ReadPcdtWithoutClxt(Stream s, int length)
        {
            if (length < ReadUtils.DWordSize)
                throw new WordFileReaderException("Invalid Prc: length");
            int read = 0;
            Pcdt pcdt = new Pcdt();
            pcdt.clxt = Pcdt.DefaultClxt;
            pcdt.lcb = BitConverter.ToUInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize, ref read), 0);            

            if (read + (long)pcdt.lcb > length)
                throw new WordFileReaderException("Invalid Prc: length");

            pcdt.PlcPcd = ReadPlcPcd(s, pcdt.lcb);
            return pcdt;
        }

        private static PlcPcd ReadPlcPcd(Stream s, uint length)
        {
            if (length < 4)
                throw new WordFileReaderException("Invalid Prc: length (0)");

            int n = PlcPcd.CalcLength(length);
            int read = 0;
            PlcPcd plcPcd = new PlcPcd();
            plcPcd.CPs = new uint[n + 1];
            for (int i = 0; i <= n; i++)
            {
                plcPcd.CPs[i] = BitConverter.ToUInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize, ref read), 0);
            }
            plcPcd.Pcds = new Pcd[n];
            for (int i = 0; i < n; i++)
            {
                plcPcd.Pcds[i] = ReadPcd(s);
                read += Pcd.Size;
            }
            if (read != length)
                throw new WordFileReaderException("Invalid PlcPcd: length");
            return plcPcd;
        }

        private static Pcd ReadPcd(Stream s)
        {
            byte[] data = ReadUtils.ReadExact(s, Pcd.Size);
            Pcd pcd = new Pcd();
            pcd.fNoParaLast = (data[0] & 0x01) != 0;
            pcd.fR1 = (data[0] & 0x02) != 0;
            pcd.fDirty = (data[0] & 0x04) != 0;
            uint uCompressed = BitConverter.ToUInt32(data, 2);
            FcCompressed comressed = new FcCompressed();
            comressed.fc = (int)(uCompressed & 0x3FFFFFFF);
            comressed.fCompressed = (uCompressed & 0x40000000) != 0;
            pcd.fc = comressed;
            ushort uPrm = BitConverter.ToUInt16(data, 6);
            Prm prm = new Prm();
            prm.prm = uPrm;
            pcd.prm = prm;
            return pcd;
        }

        private static Prc ReadPrcWithoutCltx(Stream s, int length, out int read)
        {
            if (length < ReadUtils.WordSize)
                throw new WordFileReaderException("Invalid Prc: length");
            read = 0;

            Prc prc = new Prc();
            prc.clxt = Prc.DefaultClxt;
            prc.cbGrpprl = BitConverter.ToInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);

            int cbGrpprl = prc.cbGrpprl;
            if (length < cbGrpprl || cbGrpprl < 0)
                throw new WordFileReaderException("Invalid Prc: array length");

            prc.GrpPrl = ReadPrls(s, cbGrpprl, ref read);

            return prc;
        }

        internal static Prl[] ReadPrls(Stream s, int length, ref int read)
        {
            List<Prl> prls = new List<Prl>();
            int position = 0;
            while (position < length)
            {
                int readPortion;
                prls.Add(ReadPrl(s, length - position, out readPortion));
                position += readPortion;
            }
            read += position;
            return prls.ToArray();
        }

        private static Prl ReadPrl(Stream s, int length, out int read)
        {
            if (length < 2) throw new WordFileReaderException("Invalid Prl: length");

            read = 0;
            Prl prl = new Prl();
            prl.sprm = ReadSprm(s, ref read);

            int operandLength = GetLengthFromSprm(prl.sprm);
            if (operandLength < 0)
            {
                if (prl.sprm.sprm == SinglePropertyModifiers.sprmPChgTabs)
                {
                    if (read + ReadUtils.ByteSize > length) throw new WordFileReaderException("Invalid Prl: operand length (1)");
                    operandLength = ReadUtils.ReadByte(s, ref read);
                    if (operandLength == 255)
                    {
                        if (read + ReadUtils.ByteSize > length) throw new WordFileReaderException("Invalid Prl: operand length (2)");
                        operandLength = ReadUtils.ReadByte(s, ref read);

                        MemoryStream ms = new MemoryStream();
                        byte cTabs = ReadUtils.ReadByte(s);
                        ms.WriteByte(cTabs);
                        int delPortionSize = 1 + cTabs * 4;
                        if (read + delPortionSize > length) throw new WordFileReaderException("Invalid Prl: operand length (3)");
                        byte[] data = ReadUtils.ReadExact(s, cTabs * 4);
                        ms.Write(data, 0, data.Length);
                        if (read + 1 + delPortionSize > length) throw new WordFileReaderException("Invalid Prl: operand length (4)");
                        cTabs = ReadUtils.ReadByte(s);
                        ms.WriteByte(cTabs);

                        int addPortionSize = 1 + cTabs * 3;
                        if (read + delPortionSize + addPortionSize > length) throw new WordFileReaderException("Invalid Prl: operand length (5)");
                        data = ReadUtils.ReadExact(s, cTabs * 3);
                        ms.Write(data, 0, data.Length);
                        ms.Close();

                        data = ms.ToArray();
                        operandLength = data.Length;

                        prl.operand = data;
                        read += operandLength;
                        return prl;
                    }
                }
                else if (prl.sprm.sprm == SinglePropertyModifiers.sprmTDefTable)
                {
                    if (read + ReadUtils.WordSize > length) throw new WordFileReaderException("Invalid Prl: operand length (6)");
                    operandLength = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0) + 1;
                }
                else
                {
                    if (read + ReadUtils.ByteSize > length) throw new WordFileReaderException("Invalid Prl: operand length (0)");
                    operandLength = ReadUtils.ReadByte(s, ref read);
                }            
            }
            if (read + operandLength > length) throw new WordFileReaderException("Invalid Prl: operand length");
            prl.operand = ReadUtils.ReadExact(s, operandLength, ref read);
            return prl;
        }

        private static Sprm ReadSprm(Stream s, ref int read)
        {
            ushort data = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            return new Sprm(data);
        }

        static int GetLengthFromSprm(Sprm sprm)
        {
            switch (sprm.spra)
            {
                case 0:
                case 1:
                    return 1;
                case 2:
                case 4:
                case 5:
                    return 2;
                case 3:
                    return 4;
                case 6:
                    return -1;
                case 7:
                    return 3;
                default:
                    throw new WordFileReaderException("Invalid Sprm.spra value");
            }
        }

        internal static PlcBtePapx ReadPlcfBtePapx(Stream s, uint length)
        {
            int n = PlcBtePapx.GetLength(length);
            PlcBtePapx plcfBtePapx = new PlcBtePapx();
            plcfBtePapx.aFC = new uint[n + 1];
            for (int i = 0; i <= n; i++)
            {
                plcfBtePapx.aFC[i] = BitConverter.ToUInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize), 0);

            }
            plcfBtePapx.aPnBtePapx = new PnFkpPapx[n];
            for (int i = 0; i < n; i++)
            {
                PnFkpPapx item = new PnFkpPapx();
                item.pn = BitConverter.ToUInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize), 0);
                plcfBtePapx.aPnBtePapx[i] = item;
            }
            return plcfBtePapx;
        }

        internal static PapxFkp ReadPapxFkp(Stream s)
        {
            byte[] data = ReadUtils.ReadExact(s, PapxFkp.Size);
            MemoryStream dataStream = new MemoryStream(data, false);

            PapxFkp papxFkp = new PapxFkp();
            
            papxFkp.cpara = data[PapxFkp.Size - 1];
            int cpara = papxFkp.cpara;
            if (cpara < 1 || cpara > 0x1D)
                throw new WordFileReaderException("Invalid PapxFkp: cpara");

            papxFkp.rgfc = new uint[cpara + 1];
            for (int i = 0; i <= cpara; i++)
            {
                papxFkp.rgfc[i] = BitConverter.ToUInt32(data, i * 4);
            }
            papxFkp.rgbx = new KeyValuePair<BxPap, PapxInFkps>[cpara];
            int rgbxOffset = (cpara + 1) * 4;
            for (int i = 0; i < cpara; i++)
            {
                BxPap bxPap = new BxPap();
                bxPap.bOffset = data[rgbxOffset + i * BxPap.Size];

                int papxInFkpsOffset = 2 * bxPap.bOffset;

                PapxInFkps papxInFkps = new PapxInFkps();
                papxInFkps.cb = data[papxInFkpsOffset++];
                int grpprlInPapxLength;
                if (papxInFkps.cb == 0)
                {
                    int cb_ = data[papxInFkpsOffset++];
                    grpprlInPapxLength = 2 * cb_;
                }
                else
                {
                    grpprlInPapxLength = 2 * papxInFkps.cb - 1;
                }
                GrpPrlAndIstd grpPrlAndIstd = new GrpPrlAndIstd();
                grpPrlAndIstd.istd = BitConverter.ToUInt16(data, papxInFkpsOffset);
                dataStream.Position = papxInFkpsOffset + 2;
                int position = 2;
                grpPrlAndIstd.grpprl = ReadPrls(dataStream, grpprlInPapxLength - position, ref position);

                papxInFkps.grpprlInPapx = grpPrlAndIstd; 

                papxFkp.rgbx[i] = new KeyValuePair<BxPap,PapxInFkps>(bxPap, papxInFkps);
            }
            return papxFkp;
        }

        internal static PlcBteChpx ReadPlcBteChpx(Stream s, uint length)
        {
            int n = PlcBteChpx.GetLength(length);
            PlcBteChpx plcfBteChpx = new PlcBteChpx();
            plcfBteChpx.aFC = new uint[n + 1];
            for (int i = 0; i <= n; i++)
            {
                plcfBteChpx.aFC[i] = BitConverter.ToUInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize), 0);

            }
            plcfBteChpx.aPnBteChpx = new PnFkpChpx[n];
            for (int i = 0; i < n; i++)
            {
                PnFkpChpx item = new PnFkpChpx();
                item.pn = BitConverter.ToUInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize), 0);
                plcfBteChpx.aPnBteChpx[i] = item;
            }
            return plcfBteChpx;
        }

        internal static ChpxFkp ReadChpxFkp(Stream s)
        {
            byte[] data = ReadUtils.ReadExact(s, ChpxFkp.Size);
            MemoryStream dataStream = new MemoryStream(data, false);

            ChpxFkp chpxFkp = new ChpxFkp();

            chpxFkp.crun = data[PapxFkp.Size - 1];
            int crun = chpxFkp.crun;
            if (crun < 1 || crun > 0x65)
                throw new WordFileReaderException("Invalid ChpxFkp: cpara");

            chpxFkp.rgfc = new uint[crun + 1];
            for (int i = 0; i <= crun; i++)
            {
                chpxFkp.rgfc[i] = BitConverter.ToUInt32(data, i * ReadUtils.DWordSize);
            }
            chpxFkp.rgb = new KeyValuePair<byte, Chpx>[crun];
            int rgbOffset = (crun + 1) * ReadUtils.DWordSize;
            for (int i = 0; i < crun; i++)
            {
                byte offset = data[rgbOffset + i];

                int chpxFkpOffset = 2 * offset;

                Chpx chpx = new Chpx();
                chpx.cb = data[chpxFkpOffset++];
                int grpprlLength = chpx.cb;
                dataStream.Position = chpxFkpOffset;
                int position = 0;
                chpx.grpprl = ReadPrls(dataStream, grpprlLength - position, ref position);

                chpxFkp.rgb[i] = new KeyValuePair<byte, Chpx>(offset, chpx);
            }
            return chpxFkp;
        }

        internal static STSH ReadStsh(Stream s, uint stshLength)
        {
            int read;
            STSH stsh = new STSH();

            read = 0;
            ushort cbStshi = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);

            
            int readPortion;
            Stshi stshi = ReadStshi(s, cbStshi, out readPortion);
            read += readPortion;

            stsh.stshi = stshi;

            stsh.rglpstd = new STD[stsh.stshi.stshif.cstd];
            for (int i = 0; i < stsh.rglpstd.Length; i++)
            {
                ushort cbStd = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);

                if (cbStd > 0)
                {
                    STD std = ReadStd(s, cbStd,
                        stshi.stshif.cbSTDBaseInFile, out readPortion);
                    stsh.rglpstd[i] = std;
                    read += readPortion;
                }
            }

            return stsh;
        }

        private static STD ReadStd(Stream s, ushort length, ushort stdfLength, out int read)
        {
            STD std = new STD();
            read = 0;
            Stdf stdf = new Stdf();
            StdfBase stdfBase = new StdfBase();
            byte[] data = ReadUtils.ReadExact(s, StdfBase.Size, ref read);

            stdfBase.sti = (ushort)(BitConverter.ToUInt16(data, 0) & 0xFFF);
            stdfBase.stk = (byte)(data[2] & 0x0F);
            stdfBase.istdBase = (ushort)(BitConverter.ToUInt16(data, 2) >> 4);
            stdfBase.cupx = (byte)(data[4] & 0x0F);
            stdfBase.istdNext = (ushort)(BitConverter.ToUInt16(data, 4) >> 4);
            stdfBase.bchUpe = BitConverter.ToUInt16(data, 6);
            ushort grfstdData = BitConverter.ToUInt16(data, 8);
            GRFSTD grfstd = new GRFSTD();
            grfstd.fAutoRedef = (grfstdData & 0x0001) != 0;
            grfstd.fHidden = (grfstdData & 0x0002) != 0;
            grfstd.f97LidsSet = (grfstdData & 0x0004) != 0;
            grfstd.fCopyLang = (grfstdData & 0x0008) != 0;
            grfstd.fPersonalCompose = (grfstdData & 0x0010) != 0;
            grfstd.fPersonalReply = (grfstdData & 0x0020) != 0;
            grfstd.fPersonal = (grfstdData & 0x0040) != 0;
            grfstd.fSemiHidden = (grfstdData & 0x0100) != 0;
            grfstd.fLocked = (grfstdData & 0x0200) != 0;
            grfstd.fUnhideWhenUsed = (grfstdData & 0x0800) != 0;
            grfstd.fQFormat = (grfstdData & 0x1000) != 0;
            stdfBase.grfstd = grfstd;
            stdf.stdfBase = stdfBase;

            if (StdfBase.Size + StdfPost2000OrNone.Size <= stdfLength)
            {
                data = ReadUtils.ReadExact(s, StdfPost2000OrNone.Size, ref read);

                StdfPost2000OrNone stdfPost2000 = new StdfPost2000OrNone();
                stdfPost2000.istdLink = (ushort)(BitConverter.ToUInt16(data, 0) & 0xFFF);
                stdfPost2000.fHasOriginalStyle = (data[1] & 0x10) != 0;
                stdfPost2000.rsid = BitConverter.ToUInt32(data, 2);
                stdfPost2000.iPriority = (ushort)(BitConverter.ToUInt16(data, 6) >> 4);
                stdf.StdfPost2000OrNone = stdfPost2000;
            }
            else
                stdf.StdfPost2000OrNone = null;
            std.stdf = stdf;

            std.xstzName = ReadXstz(s, length - read, ref read);

            switch (std.stdf.stdfBase.stk)
            {
                case GrLPUpxSw.StkParaGRLPUPXStkValue:
                    std.grLPUpxSw = ReadStkParaGRLPUPX(s, length - read, ref read);
                    break;
                case GrLPUpxSw.StkCharGRLPUPXStkValue:
                    std.grLPUpxSw = ReadStkCharGRLPUPX(s, length - read, ref read);
                    break;
                case GrLPUpxSw.StkTableGRLPUPXStkValue:
                    std.grLPUpxSw = ReadStkTableGRLPUPX(s, length - read, ref read);
                    break;
                case GrLPUpxSw.StkListGRLPUPXStkValue:
                    std.grLPUpxSw = ReadStkListGRLPUPX(s, length - read, ref read);
                    break;
                default:
                    throw new WordFileReaderException("Invalid Std: stk");
            }
            return std;
        }

        private static Xstz ReadXstz(Stream s, long length, ref int read)
        {
            Xstz xstz = new Xstz();
            xstz.xst = new Xst();
            if (length < ReadUtils.WordSize)
                throw new WordFileReaderException("Invalid Xst: length (0)");
            xstz.xst.cch = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            if (length < ReadUtils.WordSize + 2 * xstz.xst.cch)
                throw new WordFileReaderException("Invalid Xst: length (1)");
            xstz.xst.rgtchar = Encoding.Unicode.GetChars(ReadUtils.ReadExact(s, 2 * xstz.xst.cch, ref read));
            if (length < ReadUtils.WordSize + 2 * xstz.xst.cch + ReadUtils.WordSize)
                throw new WordFileReaderException("Invalid Xstz: length");
            ushort zeros = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            if (zeros != 0) throw new WordFileReaderException("Invalid Xstz: no zeros");
            return xstz;
        }

        private static Stshi ReadStshi(Stream s, long length, out int read)
        {
            Stshi stshi = new Stshi();

            read = 0;
            if (length < 18)
                throw new WordFileReaderException("Invalid Stshi: length (0)");
            byte[] data = ReadUtils.ReadExact(s, 18, ref read);
            Stshif stshif = new Stshif();
            stshif.cstd = BitConverter.ToUInt16(data, 0);
            stshif.cbSTDBaseInFile = BitConverter.ToUInt16(data, 2);
            stshif.stiMaxWhenSaved = BitConverter.ToUInt16(data, 6);
            stshif.istdMaxFixedWhenSaved = BitConverter.ToUInt16(data, 8);
            stshif.nVerBuiltInNamesWhenSaved = BitConverter.ToUInt16(data, 10);
            stshif.ftcAsci = BitConverter.ToInt16(data, 12);
            stshif.ftcFE = BitConverter.ToInt16(data, 14);
            stshif.ftcOther = BitConverter.ToInt16(data, 16);
            stshi.stshif = stshif;
            

            if (read < length)
            {
                stshi.ftcBi = BitConverter.ToInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            }

            if (read < length)
            {
                StshiLsd stshiLsd = new StshiLsd();
                stshiLsd.cbLSD = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);

                stshiLsd.mpstiilsd = new LSD[stshif.stiMaxWhenSaved];
                for (int i = 0; i < stshiLsd.mpstiilsd.Length; i++)
                {
                    data = ReadUtils.ReadExact(s, stshiLsd.cbLSD, ref read);

                    LSD lsd = new LSD();
                    lsd.fLocked = (data[0] & 0x01) != 0;
                    lsd.fSemiHidden = (data[0] & 0x02) != 0;
                    lsd.fUnhideWhenUsed = (data[0] & 0x04) != 0;
                    lsd.fQFormat = (data[0] & 0x08) != 0;
                    lsd.iPriority = (ushort)((data[1] << 4) | (data[0] >> 4));
                    stshiLsd.mpstiilsd[i] = lsd;
                }
                stshi.StshiLsd = stshiLsd;
            }

            if (read < length)
            {
                int grpprlChpStandard_cbGrpprl = BitConverter.ToInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize), 0);
                ReadUtils.Skip(s, grpprlChpStandard_cbGrpprl); // ignoring
                read += 4 + grpprlChpStandard_cbGrpprl;
                int grpprlPapStandard_cbGrpprl = BitConverter.ToInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize), 0);
                ReadUtils.Skip(s, grpprlPapStandard_cbGrpprl); // ignoring
                read += 4 + grpprlPapStandard_cbGrpprl;
            }
            return stshi;
        }

        private static StkParaGRLPUPX ReadStkParaGRLPUPX(Stream s, int length, ref int read)
        {
            int position = 0;
            StkParaGRLPUPX stkParaGRLPUPX = new StkParaGRLPUPX();
            stkParaGRLPUPX.upxPapx = ReadUpxPapx(s, length - position, ref position);
            stkParaGRLPUPX.upxChpx = ReadUpxChpx(s, length - position, ref position);

            if (position < length)
            {
                StkParaUpxGrLPUpxRM stkParaUpxGrLPUpxRM = new StkParaUpxGrLPUpxRM();
                stkParaUpxGrLPUpxRM.upxRm = ReadUpxRm(s, length - position, ref position);
                stkParaUpxGrLPUpxRM.upxPapxRM = ReadUpxPapx(s, length - position, ref position);
                stkParaUpxGrLPUpxRM.upxChpxRM = ReadUpxChpx(s, length - position, ref position);
                stkParaGRLPUPX.StkParaLPUpxGrLPUpxRM = stkParaUpxGrLPUpxRM;

                ReadUtils.Skip(s, length - position);
            }

            read += length;
            return stkParaGRLPUPX;
        }

        private static UpxRm ReadUpxRm(Stream s, int length, ref int read)
        {
            if(length < UpxRm.Size)
                throw new WordFileReaderException("Invalid UpxRm: length");

            UpxRm upxRm = new UpxRm();
            upxRm.date = ReadDTTM(s, ref read);
            upxRm.ibstAuthor = BitConverter.ToInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            return upxRm;
        }

        private static UpxPapx ReadUpxPapx(Stream s, int length, ref int read)
        {
            UpxPapx upxPapx = new UpxPapx();
            ushort cbUpx = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            if (cbUpx + ReadUtils.WordSize > length)
                throw new WordFileReaderException("Invalid UpxPapx: length");
            upxPapx.istd = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            upxPapx.grpprlPapx = ReadPrls(s, cbUpx - ReadUtils.WordSize, ref read);
            if ((cbUpx & 1) != 0) ReadUtils.ReadByte(s, ref read);

            return upxPapx;
        }
        private static UpxChpx ReadUpxChpx(Stream s, int length, ref int read)
        {
            UpxChpx upxChpx = new UpxChpx();
            ushort cbUpx = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            if (cbUpx + ReadUtils.WordSize > length)
                throw new WordFileReaderException("Invalid UpxChpx: length");
            upxChpx.grpprlChpx = ReadPrls(s, cbUpx, ref read);
            if ((cbUpx & 1) != 0) ReadUtils.ReadByte(s, ref read);
            return upxChpx;
        }

        private static StkCharGRLPUPX ReadStkCharGRLPUPX(Stream s, int length, ref int read)
        {
            int position = 0;
            StkCharGRLPUPX stkCharGRLPUPX = new StkCharGRLPUPX();
            stkCharGRLPUPX.upxChpx = ReadUpxChpx(s, length - position, ref position);

            if (position < length)
            {
                StkCharUpxGrLPUpxRM stkCharUpxGrLPUpxRM = new StkCharUpxGrLPUpxRM();
                stkCharUpxGrLPUpxRM.upxRm = ReadUpxRm(s, length - position, ref position);
                stkCharUpxGrLPUpxRM.upxChpxRM = ReadUpxChpx(s, length - position, ref position);
                stkCharGRLPUPX.StkCharLPUpxGrLPUpxRM = stkCharUpxGrLPUpxRM;

                ReadUtils.Skip(s, length - position);
            }

            read += length;
            return stkCharGRLPUPX;
        }

        private static StkTableGRLPUPX ReadStkTableGRLPUPX(Stream s, int length, ref int read)
        {
            StkTableGRLPUPX stkTableGRLPUPX = new StkTableGRLPUPX();
            int position = 0;
            stkTableGRLPUPX.upxTapx = ReadUpxTapx(s, length - position, ref position);
            stkTableGRLPUPX.upxPapx = ReadUpxPapx(s, length - position, ref position);
            stkTableGRLPUPX.upxChpx = ReadUpxChpx(s, length - position, ref position);

            read += position;
            return stkTableGRLPUPX;
        }

        private static UpxTapx ReadUpxTapx(Stream s, int length, ref int read)
        {
            UpxTapx upxTapx = new UpxTapx();

            ushort cbUpx = BitConverter.ToUInt16(ReadUtils.ReadExact(s, ReadUtils.WordSize, ref read), 0);
            if (cbUpx + ReadUtils.WordSize > length)
                throw new WordFileReaderException("Invalid UpxTapx: length");
            upxTapx.grpprlTapx = ReadPrls(s, cbUpx, ref read);
            if ((cbUpx & 1) != 0) ReadUtils.ReadByte(s, ref read);
            return upxTapx;
        }

        private static StkListGRLPUPX ReadStkListGRLPUPX(Stream s, int length, ref int read)
        {
            int position = 0;
            StkListGRLPUPX stkListGRLPUPX = new StkListGRLPUPX();
            stkListGRLPUPX.upxPapx = ReadUpxPapx(s, length - position, ref position);

            read += position;
            return stkListGRLPUPX;
        }

        private static DTTM ReadDTTM(Stream s, ref int read)
        {
            uint dttmData = BitConverter.ToUInt32(ReadUtils.ReadExact(s, ReadUtils.DWordSize, ref read), 0);

            return ParseDTTM(dttmData);
        }

        internal static DTTM ParseDTTM(uint dttmData)
        {
            DTTM dttm = new DTTM();
            dttm.mint = (byte)(dttmData & 0x3F);
            dttm.hr = (byte)((dttmData >> 6) & 0x1F);
            dttm.dom = (byte)((dttmData >> 11) & 0x1F);
            dttm.mon = (byte)((dttmData >> 16) & 0x0F);
            dttm.year = (ushort)((dttmData >> 20) & 0x1FF);
            dttm.wdy = (byte)((dttmData >> 29) & 0x07);
            return dttm;
        }
    }
}
