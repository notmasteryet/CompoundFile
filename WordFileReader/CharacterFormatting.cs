// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordFileReader
{
    public class CharacterFormatting
    {
        private WordDocument owner;
        public int Offset { get; private set; }
        public int Length { get; private set; }
        internal WordDocument.FileCharacterPosition FileCharacterPosition { get; private set; }
        internal Chpx Chpx { get; private set; }

        public StyleDefinition Style
        {
            get
            {
                byte[] istdData = GetProperty(SinglePropertyModifiers.sprmCIstd);
                if (istdData != null)
                {
                    ushort istd = BitConverter.ToUInt16(istdData, 0);
                   
                    if(owner.StyleDefinitionsMap.ContainsKey(istd))
                        return owner.StyleDefinitionsMap[istd];                    
                }                
                return null;
            }
        }

        internal CharacterFormatting(WordDocument owner, int offset, int length,
            WordDocument.FileCharacterPosition fcp, Chpx chpx)
        {
            this.owner = owner;
            this.Offset = offset;
            this.Length = length;
            this.FileCharacterPosition = fcp;
            this.Chpx = chpx;
        }

        public string GetText()
        {
            return owner.GetText(Offset, Length);
        }

        private byte[] GetProperty(ushort sprm)
        {
            foreach (var p in Chpx.grpprl)
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

        public StyleCollection GetStyles()
        {
            return GetStyles(FormattingLevel.Character);
        }

        public StyleCollection GetStyles(FormattingLevel level)
        {
            if(level != FormattingLevel.Character &&
                level != FormattingLevel.CharacterStyle) throw new ArgumentOutOfRangeException("level");

            List<Prl[]> prls = new List<Prl[]>();
            prls.Add(Chpx.grpprl);

            if (level == FormattingLevel.CharacterStyle)
            {
                StyleDefinition def = this.Style;
                if (def != null) def.ExpandStyles(prls);
            }
            return new StyleCollection(prls);
        }
    }

}
