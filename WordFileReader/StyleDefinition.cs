// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordFileReader
{
    public class StyleDefinition
    {
        WordDocument owner;
        STD std;

        public string Name
        {
            get { return std.xstzName.ToString(); }
        }

        public bool IsTextStyle
        {
            get { return std.stdf.stdfBase.stk == GrLPUpxSw.StkCharGRLPUPXStkValue; }
        }

        internal StyleDefinition(WordDocument owner, STD std)
        {
            this.owner = owner;
            this.std = std;
        }

        public StyleCollection GetStyles()
        {
            List<Prl[]> styles = new List<Prl[]>();
            ExpandStyles(styles);
            return new StyleCollection(styles);
        }

        internal void ExpandStyles(List<Prl[]> styles)
        {
            if (std.stdf.stdfBase.istdBase != StdfBase.IstdNull)
            {
                owner.StyleDefinitionsMap[std.stdf.stdfBase.istdBase].ExpandStyles(styles);
            }

            switch (std.stdf.stdfBase.stk)
            {
                case GrLPUpxSw.StkParaGRLPUPXStkValue:
                    styles.Add(((StkParaGRLPUPX)std.grLPUpxSw).upxPapx.grpprlPapx);
                    styles.Add(((StkParaGRLPUPX)std.grLPUpxSw).upxChpx.grpprlChpx);
                    break;
                case GrLPUpxSw.StkCharGRLPUPXStkValue:
                    styles.Add(((StkCharGRLPUPX)std.grLPUpxSw).upxChpx.grpprlChpx);
                    break;
                case GrLPUpxSw.StkTableGRLPUPXStkValue:
                    styles.Add(((StkTableGRLPUPX)std.grLPUpxSw).upxTapx.grpprlTapx);
                    styles.Add(((StkTableGRLPUPX)std.grLPUpxSw).upxPapx.grpprlPapx);
                    styles.Add(((StkTableGRLPUPX)std.grLPUpxSw).upxChpx.grpprlChpx);
                    break;
                case GrLPUpxSw.StkListGRLPUPXStkValue:
                    styles.Add(((StkListGRLPUPX)std.grLPUpxSw).upxPapx.grpprlPapx);
                    break;
            }
        }
    }
}
