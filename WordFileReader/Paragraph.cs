// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordFileReader
{
    public class Paragraph
    {
        private WordDocument owner;
        public int Offset { get; private set; }
        public int Length { get; private set; }
        internal WordDocument.FileCharacterPosition FileCharacterPosition { get; private set; }
        internal PapxInFkps PapxInFkps { get; private set; }

        public bool IsInTable
        {
            get
            {
                byte[] data = GetProperty(SinglePropertyModifiers.sprmPFInTable);
                return data != null && data[0] != 0;
            }
        }

        public int TableDepth
        {
            get
            {
                byte[] data = GetProperty(SinglePropertyModifiers.sprmPItap);
                if (data == null)
                    return 0;
                else
                    return (int)BitConverter.ToUInt32(data, 0);
            }
        }

        public bool IsTableRowEnd
        {
            get
            {
                if (owner.Characters[Offset + Length - 1] == '\u0007')
                {
                    byte[] data = GetProperty(SinglePropertyModifiers.sprmPFTtp);
                    return data != null && data[0] != 0;
                }
                else
                {
                    byte[] data = GetProperty(SinglePropertyModifiers.sprmPFInnerTtp);
                    return data != null && data[0] != 0;
                }
            }
        }

        public bool IsTableCellEnd
        {
            get
            {
                if (owner.Characters[Offset + Length - 1] == '\u0007')
                {
                    return true;
                }
                else
                {
                    byte[] data = GetProperty(SinglePropertyModifiers.sprmPFInnerTableCell);
                    return data != null && data[0] != 0;
                }
            }
        }

        public bool IsList
        {
            get
            {
                byte[] data = GetProperty(SinglePropertyModifiers.sprmPIlfo);
                return data != null && BitConverter.ToInt16(data, 0) != 0;
            }
        }

        public int ListLevel
        {
            get
            {
                byte[] data = GetProperty(SinglePropertyModifiers.sprmPIlvl);
                return data != null ? data[0] : 0;
            }
        }

        public StyleDefinition Style
        {
            get
            {
                return owner.StyleDefinitionsMap[PapxInFkps.grpprlInPapx.istd];
            }
        }

        public string GetText()
        {
            return owner.GetText(Offset, Length);
        }

        internal Paragraph(WordDocument owner, int offset, int length,
            WordDocument.FileCharacterPosition fcp, PapxInFkps papxInFkps)
        {
            this.owner = owner;
            this.Offset = offset;
            this.Length = length;
            this.FileCharacterPosition = fcp;
            this.PapxInFkps = papxInFkps;
        }

        private byte[] GetProperty(ushort sprm)
        {
            foreach (var p in PapxInFkps.grpprlInPapx.grpprl)
            {
                if (p.sprm.sprm == sprm)
                    return p.operand;
            }
            foreach (var p in FileCharacterPosition.Prls)
            {
                if (p.sprm.sprm == sprm)
                    return p.operand;
            }
            return null;
        }

        public StyleCollection GetStyles(FormattingLevel level)
        {
            if (level != FormattingLevel.Paragraph &&
                level != FormattingLevel.ParagraphStyle) throw new ArgumentOutOfRangeException("level");

            List<Prl[]> prls = new List<Prl[]>();
            prls.Add(PapxInFkps.grpprlInPapx.grpprl);
            if (level == FormattingLevel.ParagraphStyle)
            {
                this.Style.ExpandStyles(prls);
            }
            return new StyleCollection(prls);
        }
    }

}
