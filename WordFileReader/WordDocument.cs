// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using CompoundFile;

namespace WordFileReader
{
    /// <summary>
    /// Word document model. 
    /// See specification for explanations.
    /// [MS-DOC]: Word Binary File Format (.doc) Structure Specification
    /// Release: Thursday, August 27, 2009
    /// </summary>
    public class WordDocument
    {
        char[] characters;
        FileCharacterPosition[] fileCharacterPositions;
        Paragraph[] paragraphs;
        CharacterFormatting[] formattings;
        Dictionary<ushort, StyleDefinition> styleDefinitionsMap;
        Prl[] defaultPrls;

        internal char[] Characters
        {
            get { return characters; }
        }

        public IEnumerable<Paragraph> Paragraphs
        {
            get { return paragraphs; }
        }

        public IEnumerable<CharacterFormatting> Formattings
        {
            get { return formattings; }
        }

        public IEnumerable<StyleDefinition> StyleDefinitions
        {
            get { return styleDefinitionsMap.Values; }
        }

        public IDictionary<ushort, StyleDefinition> StyleDefinitionsMap
        {
            get { return styleDefinitionsMap; }
        }

        public void Load(CompoundFileSystem system)
        {
            const string WordDocumentStreamName = "WordDocument";
            CompoundFileStorage wordDocumentStorage = system.GetRootStorage().FindStorage(WordDocumentStreamName);
            wordDocumentStorage.CopyToFile("wd.bin");
            using (Stream wordDocumentStream = wordDocumentStorage.CreateStream())
            {
                Fib fib;
                fib = FibStructuresReader.ReadFib(wordDocumentStream);

                string tableStreamName = GetTableStreamName(fib);
                CompoundFileStorage tableStorage = system.GetRootStorage().FindStorage(tableStreamName);
                tableStorage.CopyToFile(tableStreamName + ".bin");
                using (Stream tableStream = tableStorage.CreateStream())
                {
                    LoadContent(wordDocumentStream, tableStream, fib);
                }
            }
        }

        static string GetTableStreamName(Fib fib)
        {
            return fib.@base.fWhichTblStm ? "1Table" : "0Table";
        }

        private void LoadContent(Stream wordDocumentStream, Stream tableStream, Fib fib)
        {
            // reading text. p.35
            uint clxOffset = ((FibRbFcLcb97)fib.fibRgFcLcbBlob).fcClx;
            tableStream.Position = clxOffset;
            uint clxLength = ((FibRbFcLcb97)fib.fibRgFcLcbBlob).lcbClx;
            Clx clx = BasicTypesReader.ReadClx(tableStream, (int)clxLength);

            LoadCharacters(wordDocumentStream, clx);

            // reading paragraphs p.36
            uint plcfBtePapxOffset = ((FibRbFcLcb97)fib.fibRgFcLcbBlob).fcPlcfBtePapx;
            tableStream.Position = plcfBtePapxOffset;
            uint plcfBtePapxLength = ((FibRbFcLcb97)fib.fibRgFcLcbBlob).lcbPlcfBtePapx;
            PlcBtePapx plcfBtePapx = BasicTypesReader.ReadPlcfBtePapx(tableStream, plcfBtePapxLength);

            LoadParagraphs(wordDocumentStream, clx, plcfBtePapx);

            // reading character formattings p.43
            uint plcfBteChpxOffset = ((FibRbFcLcb97)fib.fibRgFcLcbBlob).fcPlcfBteChpx;
            tableStream.Position = plcfBteChpxOffset;
            uint plcfBteChpxLength = ((FibRbFcLcb97)fib.fibRgFcLcbBlob).lcbPlcfBteChpx;
            PlcBteChpx plcBteChpx = BasicTypesReader.ReadPlcBteChpx(tableStream, plcfBteChpxLength);
            LoadCharacterFormatting(wordDocumentStream, plcBteChpx);

            // reading styles p.46
            uint stshOffset = ((FibRbFcLcb97)fib.fibRgFcLcbBlob).fcStshf;
            tableStream.Position = stshOffset;
            uint stshLength = ((FibRbFcLcb97)fib.fibRgFcLcbBlob).lcbStshf;
            STSH stsh = BasicTypesReader.ReadStsh(tableStream, stshLength);
            LoadStsh(wordDocumentStream, stsh);
        }

        private void LoadStsh(Stream wordDocumentStream, STSH stsh)
        {
            styleDefinitionsMap = new Dictionary<ushort, StyleDefinition>();
            for (ushort istd = 0; istd < stsh.rglpstd.Length; istd++)
            {
                STD std = stsh.rglpstd[istd];
                if (std == null) continue; // skip empty

                StyleDefinition def = new StyleDefinition(this, std);
                styleDefinitionsMap.Add(istd, def);
            }

            // global/default properties
            List<Prl> defaultPrls = new List<Prl>();

            short ftcBi = stsh.stshi.ftcBi; // sprmCFtcBi
            defaultPrls.Add(new Prl(SinglePropertyModifiers.sprmCFtcBi, BitConverter.GetBytes(ftcBi)));
            short ftcAsci = stsh.stshi.stshif.ftcAsci; // sprmCRgFtc0
            defaultPrls.Add(new Prl(SinglePropertyModifiers.sprmCRgFtc0, BitConverter.GetBytes(ftcAsci)));
            short ftcFE = stsh.stshi.stshif.ftcFE; // sprmCRgFtc1
            defaultPrls.Add(new Prl(SinglePropertyModifiers.sprmCRgFtc1, BitConverter.GetBytes(ftcFE)));
            short ftcOther = stsh.stshi.stshif.ftcOther; // sprmCRgFtc2
            defaultPrls.Add(new Prl(SinglePropertyModifiers.sprmCRgFtc2, BitConverter.GetBytes(ftcOther)));
            this.defaultPrls = defaultPrls.ToArray();
        }

        private void LoadCharacterFormatting(Stream wordDocumentStream, PlcBteChpx plcfBteChpx)
        {
            List<CharacterFormatting> formattings = new List<CharacterFormatting>();
            for (int i = 0; i < plcfBteChpx.aPnBteChpx.Length; i++)
            {
                wordDocumentStream.Position = plcfBteChpx.aPnBteChpx[i].pn * 512;
                ChpxFkp papxFkp = BasicTypesReader.ReadChpxFkp(wordDocumentStream);
                for (int j = 0; j < papxFkp.rgb.Length; j++)
                {
                    int startFc = (int)papxFkp.rgfc[j];
                    int endFc = (int)papxFkp.rgfc[j + 1];
                    FileCharacterPosition fc = Array.Find(fileCharacterPositions, new Predicate<FileCharacterPosition>(delegate(FileCharacterPosition f)
                    {
                        return !(endFc <= f.Offset || f.Offset + f.BytesPerCharacter * f.Length < startFc);
                    }));
                    if (fc == null) continue; // invalid section

                    if (startFc < fc.Offset) startFc = fc.Offset;
                    if (endFc > fc.Offset + fc.BytesPerCharacter * fc.Length)
                        endFc = fc.Offset + fc.BytesPerCharacter * fc.Length;

                    CharacterFormatting para = new CharacterFormatting(this,
                        (startFc - fc.Offset) / fc.BytesPerCharacter,
                        (endFc - startFc) / fc.BytesPerCharacter,
                        fc, papxFkp.rgb[j].Value);
                    formattings.Add(para);
                }
            }
            formattings.Sort(new Comparison<CharacterFormatting>(delegate(CharacterFormatting f1, CharacterFormatting f2)
            {
                return f1.Offset - f2.Offset;
            }));
            this.formattings = formattings.ToArray();
        }

        private void LoadParagraphs(Stream wordDocumentStream, Clx clx, PlcBtePapx plcfBtePapx)
        {
            List<Paragraph> paragraphs = new List<Paragraph>();
            for (int i = 0; i < plcfBtePapx.aPnBtePapx.Length; i++)
            {
                wordDocumentStream.Position = plcfBtePapx.aPnBtePapx[i].pn * 512;
                PapxFkp papxFkp = BasicTypesReader.ReadPapxFkp(wordDocumentStream);
                for (int j = 0; j < papxFkp.rgbx.Length; j++)
                {
                    int startFc = (int)papxFkp.rgfc[j];
                    int endFc = (int)papxFkp.rgfc[j + 1];
                    FileCharacterPosition fc = Array.Find(fileCharacterPositions, new Predicate<FileCharacterPosition>(delegate(FileCharacterPosition f)
                    {
                        return !(endFc <= f.Offset || f.Offset + f.BytesPerCharacter * f.Length < startFc);
                    }));
                    if (fc == null) continue; // invalid section

                    if (startFc < fc.Offset) startFc = fc.Offset;
                    if (endFc > fc.Offset + fc.BytesPerCharacter * fc.Length)
                        endFc = fc.Offset + fc.BytesPerCharacter * fc.Length;

                    Paragraph para = new Paragraph(this, 
                        (startFc - fc.Offset) / fc.BytesPerCharacter,
                        (endFc - startFc) / fc.BytesPerCharacter,
                        fc, papxFkp.rgbx[j].Value);
                    paragraphs.Add(para);
                }
            }
            paragraphs.Sort(new Comparison<Paragraph>(delegate(Paragraph p1, Paragraph p2)
                {
                    return p1.Offset - p2.Offset;
                }));
            this.paragraphs = paragraphs.ToArray();
        }

        private void LoadCharacters(Stream wordDocumentStream, Clx clx)
        {
            PlcPcd plcPcd = clx.Pcdt.PlcPcd;
            char[] text = new char[plcPcd.CPs[plcPcd.CPs.Length - 1]];
            FileCharacterPosition[] fileCharacters = new FileCharacterPosition[plcPcd.Pcds.Length];
            int position = 0;
            for (int i = 0; i < plcPcd.Pcds.Length; i++)
            {
                int length = (int)(plcPcd.CPs[i + 1] - plcPcd.CPs[i]);
                int offset = plcPcd.Pcds[i].fc.fc;

                bool compressed = plcPcd.Pcds[i].fc.fCompressed;
                if (compressed)
                {
                    wordDocumentStream.Position = offset / 2;
                    byte[] data = ReadUtils.ReadExact(wordDocumentStream, length);
                    FcCompressedMapping.GetChars(data, 0, length, text, position);
                }
                else
                {
                    wordDocumentStream.Position = offset;
                    byte[] data = ReadUtils.ReadExact(wordDocumentStream, length * 2);
                    Encoding.Unicode.GetChars(data, 0, length, text, position);
                }

                FileCharacterPosition fc = new FileCharacterPosition();
                fc.Offset = compressed ? offset / 2 : offset;
                fc.BytesPerCharacter = compressed ? 1 : 2;
                fc.Length = length;
                fc.CharacterIndex = position;
                fc.Prls = ExpandPrm(plcPcd.Pcds[i].prm, clx);

                fileCharacters[i] = fc;

                position += length;
            }
            Array.Sort(fileCharacters, new Comparison<FileCharacterPosition>(delegate(FileCharacterPosition fc1, FileCharacterPosition fc2)
            {
                return fc1.Offset - fc2.Offset;
            }));
            this.fileCharacterPositions = fileCharacters;
            this.characters = text;
        }

        private Prl[] ExpandPrm(Prm prm, Clx clx)
        {
            if (prm.fComplex)
            {
                int index = prm.igrpprl;
                return clx.RgPrc[index].GrpPrl;
            }
            else if (prm.prm != 0x0000)
            {
                ushort sprm;
                if (!SinglePropertyModifiers.prm0Map.TryGetValue(prm.isprm, out sprm))
                    throw new WordFileReaderException("Invalid Prm: isprm");
                byte value = prm.val;

                Prl prl = new Prl();
                prl.sprm = new Sprm(sprm);
                prl.operand = new byte[] { value };
                return new Prl[] { prl };
            }
            else
                return new Prl[0];
        }


        public string GetText(int position, int length)
        {
            return new String(characters, position, length);
        }

        public StyleCollection GetDefaults()
        {
            return new StyleCollection(new Prl[][] { defaultPrls });
        }

        public StyleCollection GetStyle(int characterPosition, FormattingLevel level) 
        {
            List<Prl[]> prls = new List<Prl[]>();
            if(level >= FormattingLevel.Character)
            {
                foreach (CharacterFormatting formatting in formattings)
                {
                    if (formatting.FileCharacterPosition.Contains(characterPosition))
                    {
                        if (level >= FormattingLevel.CharacterStyle)
                        {
                            StyleDefinition definition = formatting.Style;
                            if (definition != null) definition.ExpandStyles(prls);
                        }
                    }
                }
            }
            if (level >= FormattingLevel.Paragraph)
            {
                foreach (Paragraph paragraph in paragraphs)
                {
                    if (paragraph.FileCharacterPosition.Contains(characterPosition))
                    {
                        if (level >= FormattingLevel.ParagraphStyle)
                        {
                            StyleDefinition definition = paragraph.Style;
                            definition.ExpandStyles(prls);
                        }
                    }
                }
            }

            if (level >= FormattingLevel.Part)
            {
                foreach (FileCharacterPosition fcp in fileCharacterPositions)
                {
                    if (fcp.Contains(characterPosition))
                    {
                        prls.Add(fcp.Prls);
                    }
                }
            }

            if (level >= FormattingLevel.Global)
            {
                prls.Add(defaultPrls);
            }
            return new StyleCollection(prls);
        }

        internal class FileCharacterPosition
        {
            internal int Offset;
            internal int Length;
            internal int CharacterIndex;
            internal int BytesPerCharacter;
            internal Prl[] Prls;

            internal bool Contains(int position)
            {
                return CharacterIndex <= position && position < CharacterIndex + Length;
            }
        }
    }

    public enum FormattingLevel
    {
        Character = 0,
        CharacterStyle = 1,
        Paragraph = 2,
        ParagraphStyle = 3,
        Part = 4,
        Global = 5
    }

}
