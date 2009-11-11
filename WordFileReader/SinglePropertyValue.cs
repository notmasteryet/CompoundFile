// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace WordFileReader
{
    static class SinglePropertyValue
    {
        internal static object ParseValue(uint sprm, byte[] value)
        {
            switch (sprm)
            {
                case SinglePropertyModifiers.sprmCFRMarkDel: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFRMarkIns: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFFldVanish: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCPicLocation: return ParseInt32(value);
                case SinglePropertyModifiers.sprmCIbstRMark: return ParseInt16(value);
                case SinglePropertyModifiers.sprmCDttmRMark: return ParseDTTM(value);
                case SinglePropertyModifiers.sprmCFData: return ParseBool(value);
                case SinglePropertyModifiers.sprmCIdslRMark: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmCSymbol: return ParseCSymbolOperand(value);
                case SinglePropertyModifiers.sprmCFOle2: return ParseBool(value);
                case SinglePropertyModifiers.sprmCHighlight: return ParseIco(value);
                case SinglePropertyModifiers.sprmCFWebHidden: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCRsidProp: return ParseInt32(value);
                case SinglePropertyModifiers.sprmCRsidText: return ParseInt32(value);
                case SinglePropertyModifiers.sprmCRsidRMDel: return ParseInt32(value);
                case SinglePropertyModifiers.sprmCFSpecVanish: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFMathPr: return ParseMathPrOperand(value);
                case SinglePropertyModifiers.sprmCIstd: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmCIstdPermute: return ParseSPPOperand(value);
                case SinglePropertyModifiers.sprmCPlain: return ParseByte(value);
                case SinglePropertyModifiers.sprmCKcd: return ParseByte(value);
                case SinglePropertyModifiers.sprmCFBold: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFItalic: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFStrike: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFOutline: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFShadow: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFSmallCaps: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFCaps: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFVanish: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCKul: return ParseKul(value);
                case SinglePropertyModifiers.sprmCDxaSpace: return ParseXAS(value);
                case SinglePropertyModifiers.sprmCIco: return ParseIco(value);
                case SinglePropertyModifiers.sprmCHps: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmCHpsPos: return ParseInt16(value);
                case SinglePropertyModifiers.sprmCMajority: return ParseCMajorityOperand(value);
                case SinglePropertyModifiers.sprmCIss: return ParseByte(value);
                case SinglePropertyModifiers.sprmCHpsKern: return ParseInt16(value);
                case SinglePropertyModifiers.sprmCHresi: return ParseHresiOperand(value);
                case SinglePropertyModifiers.sprmCRgFtc0: return ParseInt16(value);
                case SinglePropertyModifiers.sprmCRgFtc1: return ParseInt16(value);
                case SinglePropertyModifiers.sprmCRgFtc2: return ParseInt16(value);
                case SinglePropertyModifiers.sprmCCharScale: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmCFDStrike: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFImprint: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFSpec: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFObj: return ParseBool(value);
                case SinglePropertyModifiers.sprmCPropRMark90: return ParsePropRMarkOperand(value);
                case SinglePropertyModifiers.sprmCFEmboss: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCSfxText: return ParseByte(value);
                case SinglePropertyModifiers.sprmCFBiDi: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFBoldBi: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFItalicBi: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFtcBi: return ParseInt16(value);
                case SinglePropertyModifiers.sprmCLidBi: return ParseLID(value);
                case SinglePropertyModifiers.sprmCIcoBi: return ParseIco(value);
                case SinglePropertyModifiers.sprmCHpsBi: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmCDispFldRMark: return ParseDispFldRmOperand(value);
                case SinglePropertyModifiers.sprmCIbstRMarkDel: return ParseInt16(value);
                case SinglePropertyModifiers.sprmCDttmRMarkDel: return ParseDTTM(value);
                case SinglePropertyModifiers.sprmCBrc80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmCShd80: return ParseShd80(value);
                case SinglePropertyModifiers.sprmCIdslRMarkDel: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmCFUsePgsuSettings: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCRgLid0_80: return ParseLID(value);
                case SinglePropertyModifiers.sprmCRgLid1_80: return ParseLID(value);
                case SinglePropertyModifiers.sprmCIdctHint: return ParseByte(value);
                case SinglePropertyModifiers.sprmCCv: return ParseCOLORREF(value);
                case SinglePropertyModifiers.sprmCShd: return ParseSHDOperand(value);
                case SinglePropertyModifiers.sprmCBrc: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmCRgLid0: return ParseLID(value);
                case SinglePropertyModifiers.sprmCRgLid1: return ParseLID(value);
                case SinglePropertyModifiers.sprmCFNoProof: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCFitText: return ParseCFitTextOperand(value);
                case SinglePropertyModifiers.sprmCCvUl: return ParseCOLORREF(value);
                case SinglePropertyModifiers.sprmCFELayout: return ParseFarEastLayoutOperand(value);
                case SinglePropertyModifiers.sprmCLbcCRJ: return ParseLBCOperand(value);
                case SinglePropertyModifiers.sprmCFComplexScripts: return ParseToggleOperand(value);
                case SinglePropertyModifiers.sprmCWall: return ParseBool(value);
                case SinglePropertyModifiers.sprmCCnf: return ParseCNFOperand(value);
                case SinglePropertyModifiers.sprmCNeedFontFixup: return ParseFFM(value);
                case SinglePropertyModifiers.sprmCPbiIBullet: return ParseCP(value);
                case SinglePropertyModifiers.sprmCPbiGrf: return ParsePbiGrfOperand(value);
                case SinglePropertyModifiers.sprmCPropRMark: return ParsePropRMarkOperand(value);
                case SinglePropertyModifiers.sprmCFSdtVanish: return ParseBool(value);
                case SinglePropertyModifiers.sprmPIstd: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmPIstdPermute: return ParseSPPOperand(value);
                case SinglePropertyModifiers.sprmPIncLvl: return ParseSByte(value);
                case SinglePropertyModifiers.sprmPJc80: return ParseByte(value);
                case SinglePropertyModifiers.sprmPFKeep: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFKeepFollow: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFPageBreakBefore: return ParseBool(value);
                case SinglePropertyModifiers.sprmPIlvl: return ParseByte(value);
                case SinglePropertyModifiers.sprmPIlfo: return ParseInt16(value);
                case SinglePropertyModifiers.sprmPFNoLineNumb: return ParseBool(value);
                case SinglePropertyModifiers.sprmPChgTabsPapx: return ParsePChgTabsPapxOperand(value);
                case SinglePropertyModifiers.sprmPDxaRight80: return ParseXAS(value);
                case SinglePropertyModifiers.sprmPDxaLeft80: return ParseXAS(value);
                case SinglePropertyModifiers.sprmPNest80: return ParseXAS(value);
                case SinglePropertyModifiers.sprmPDxaLeft180: return ParseXAS(value);
                case SinglePropertyModifiers.sprmPDyaLine: return ParseLSPD(value);
                case SinglePropertyModifiers.sprmPDyaBefore: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmPDyaAfter: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmPChgTabs: return ParsePChgTabsOperand(value);
                case SinglePropertyModifiers.sprmPFInTable: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFTtp: return ParseBool(value);
                case SinglePropertyModifiers.sprmPDxaAbs: return ParseXAS_plusOne(value);
                case SinglePropertyModifiers.sprmPDyaAbs: return ParseYAS_plusOne(value);
                case SinglePropertyModifiers.sprmPDxaWidth: return ParseXAS_nonNeg(value);
                case SinglePropertyModifiers.sprmPPc: return ParsePositionCodeOperand(value);
                case SinglePropertyModifiers.sprmPWr: return ParseByte(value);
                case SinglePropertyModifiers.sprmPBrcTop80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPBrcLeft80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPBrcBottom80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPBrcRight80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPBrcBetween80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPBrcBar80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPFNoAutoHyph: return ParseBool(value);
                case SinglePropertyModifiers.sprmPWHeightAbs: return ParseWHeightAbs(value);
                case SinglePropertyModifiers.sprmPDcs: return ParseDCS(value);
                case SinglePropertyModifiers.sprmPShd80: return ParseShd80(value);
                case SinglePropertyModifiers.sprmPDyaFromText: return ParseYAS_nonNeg(value);
                case SinglePropertyModifiers.sprmPDxaFromText: return ParseXAS_nonNeg(value);
                case SinglePropertyModifiers.sprmPFLocked: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFWidowControl: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFKinsoku: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFWordWrap: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFOverflowPunct: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFTopLinePunct: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFAutoSpaceDE: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFAutoSpaceDN: return ParseBool(value);
                case SinglePropertyModifiers.sprmPWAlignFont: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmPFrameTextFlow: return ParseFrameTextFlowOperand(value);
                case SinglePropertyModifiers.sprmPOutLvl: return ParseByte(value);
                case SinglePropertyModifiers.sprmPFBiDi: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFNumRMIns: return ParseBool(value);
                case SinglePropertyModifiers.sprmPNumRM: return ParseNumRMOperand(value);
                case SinglePropertyModifiers.sprmPHugePapx: return ParseUInt32(value);
                case SinglePropertyModifiers.sprmPFUsePgsuSettings: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFAdjustRight: return ParseBool(value);
                case SinglePropertyModifiers.sprmPItap: return ParseInt32(value);
                case SinglePropertyModifiers.sprmPDtap: return ParseInt32(value);
                case SinglePropertyModifiers.sprmPFInnerTableCell: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFInnerTtp: return ParseBool(value);
                case SinglePropertyModifiers.sprmPShd: return ParseSHDOperand(value);
                case SinglePropertyModifiers.sprmPBrcTop: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmPBrcLeft: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmPBrcBottom: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmPBrcRight: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmPBrcBetween: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmPBrcBar: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmPDxcRight: return ParseInt16(value);
                case SinglePropertyModifiers.sprmPDxcLeft: return ParseInt16(value);
                case SinglePropertyModifiers.sprmPDxcLeft1: return ParseInt16(value);
                case SinglePropertyModifiers.sprmPDylBefore: return ParseInt16(value);
                case SinglePropertyModifiers.sprmPDylAfter: return ParseInt16(value);
                case SinglePropertyModifiers.sprmPFOpenTch: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFDyaBeforeAuto: return ParseBool(value);
                case SinglePropertyModifiers.sprmPFDyaAfterAuto: return ParseBool(value);
                case SinglePropertyModifiers.sprmPDxaRight: return ParseXAS(value);
                case SinglePropertyModifiers.sprmPDxaLeft: return ParseXAS(value);
                case SinglePropertyModifiers.sprmPNest: return ParseXAS(value);
                case SinglePropertyModifiers.sprmPDxaLeft1: return ParseXAS(value);
                case SinglePropertyModifiers.sprmPJc: return ParseByte(value);
                case SinglePropertyModifiers.sprmPFNoAllowOverlap: return ParseBool(value);
                case SinglePropertyModifiers.sprmPWall: return ParseBool(value);
                case SinglePropertyModifiers.sprmPIpgp: return ParseUInt32(value);
                case SinglePropertyModifiers.sprmPCnf: return ParseCNFOperand(value);
                case SinglePropertyModifiers.sprmPRsid: return ParseInt32(value);
                case SinglePropertyModifiers.sprmPIstdListPermute: return ParseSPPOperand(value);
                case SinglePropertyModifiers.sprmPTableProps: return ParseUInt32(value);
                case SinglePropertyModifiers.sprmPTIstdInfo: return ParsePTIstdInfoOperand(value);
                case SinglePropertyModifiers.sprmPFContextualSpacing: return ParseBool(value);
                case SinglePropertyModifiers.sprmPPropRMark: return ParsePropRMarkOperand(value);
                case SinglePropertyModifiers.sprmPFMirrorIndents: return ParseBool(value);
                case SinglePropertyModifiers.sprmPTtwo: return ParseByte(value);
                case SinglePropertyModifiers.sprmTJc90: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmTDxaLeft: return ParseXAS(value);
                case SinglePropertyModifiers.sprmTDxaGapHalf: return ParseXAS(value);
                case SinglePropertyModifiers.sprmTFCantSplit90: return ParseBool(value);
                case SinglePropertyModifiers.sprmTTableHeader: return ParseBool(value);
                case SinglePropertyModifiers.sprmTTableBorders80: return ParseTableBordersOperand80(value);
                case SinglePropertyModifiers.sprmTDyaRowHeight: return ParseYAS(value);
                case SinglePropertyModifiers.sprmTDefTable: return ParseTDefTableOperand(value);
                case SinglePropertyModifiers.sprmTDefTableShd80: return ParseDefTableShd80Operand(value);
                case SinglePropertyModifiers.sprmTTlp: return ParseTLP(value);
                case SinglePropertyModifiers.sprmTFBiDi: return ParseBool(value);
                case SinglePropertyModifiers.sprmTDefTableShd3rd: return ParseDefTableShdOperand(value);
                case SinglePropertyModifiers.sprmTPc: return ParsePositionCodeOperand(value);
                case SinglePropertyModifiers.sprmTDxaAbs: return ParseXAS_plusOne(value);
                case SinglePropertyModifiers.sprmTDyaAbs: return ParseYAS_plusOne(value);
                case SinglePropertyModifiers.sprmTDxaFromText: return ParseXAS_nonNeg(value);
                case SinglePropertyModifiers.sprmTDyaFromText: return ParseYAS_nonNeg(value);
                case SinglePropertyModifiers.sprmTDefTableShd: return ParseDefTableShdOperand(value);
                case SinglePropertyModifiers.sprmTTableBorders: return ParseTableBordersOperand(value);
                case SinglePropertyModifiers.sprmTTableWidth: return ParseFtsWWidth_Table(value);
                case SinglePropertyModifiers.sprmTFAutofit: return ParseBool(value);
                case SinglePropertyModifiers.sprmTDefTableShd2nd: return ParseDefTableShdOperand(value);
                case SinglePropertyModifiers.sprmTWidthBefore: return ParseFtsWWidth_TablePart(value);
                case SinglePropertyModifiers.sprmTWidthAfter: return ParseFtsWWidth_TablePart(value);
                case SinglePropertyModifiers.sprmTFKeepFollow: return ParseBool(value);
                case SinglePropertyModifiers.sprmTBrcTopCv: return ParseBrcCvOperand(value);
                case SinglePropertyModifiers.sprmTBrcLeftCv: return ParseBrcCvOperand(value);
                case SinglePropertyModifiers.sprmTBrcBottomCv: return ParseBrcCvOperand(value);
                case SinglePropertyModifiers.sprmTBrcRightCv: return ParseBrcCvOperand(value);
                case SinglePropertyModifiers.sprmTDxaFromTextRight: return ParseXAS_nonNeg(value);
                case SinglePropertyModifiers.sprmTDyaFromTextBottom: return ParseYAS_nonNeg(value);
                case SinglePropertyModifiers.sprmTSetBrc80: return ParseTableBrc80Operand(value);
                case SinglePropertyModifiers.sprmTInsert: return ParseTInsertOperand(value);
                case SinglePropertyModifiers.sprmTDelete: return ParseItcFirstLim(value);
                case SinglePropertyModifiers.sprmTDxaCol: return ParseTDxaColOperand(value);
                case SinglePropertyModifiers.sprmTMerge: return ParseItcFirstLim(value);
                case SinglePropertyModifiers.sprmTSplit: return ParseItcFirstLim(value);
                case SinglePropertyModifiers.sprmTTextFlow: return ParseCellRangeTextFlow(value);
                case SinglePropertyModifiers.sprmTVertMerge: return ParseVertMergeOperand(value);
                case SinglePropertyModifiers.sprmTVertAlign: return ParseCellRangeVertAlign(value);
                case SinglePropertyModifiers.sprmTSetShd: return ParseTableShadeOperand(value);
                case SinglePropertyModifiers.sprmTSetShdOdd: return ParseTableShadeOperand(value);
                case SinglePropertyModifiers.sprmTSetBrc: return ParseTableBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellPadding: return ParseCSSAOperand(value);
                case SinglePropertyModifiers.sprmTCellSpacingDefault: return ParseCSSAOperand(value);
                case SinglePropertyModifiers.sprmTCellPaddingDefault: return ParseCSSAOperand(value);
                case SinglePropertyModifiers.sprmTCellWidth: return ParseTableCellWidthOperand(value);
                case SinglePropertyModifiers.sprmTFitText: return ParseCellRangeFitText(value);
                case SinglePropertyModifiers.sprmTFCellNoWrap: return ParseCellRangeNoWrap(value);
                case SinglePropertyModifiers.sprmTIstd: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmTCellPaddingStyle: return ParseCSSAOperand(value);
                case SinglePropertyModifiers.sprmTCellFHideMark: return ParseCellHideMarkOperand(value);
                case SinglePropertyModifiers.sprmTSetShdTable: return ParseSHDOperand(value);
                case SinglePropertyModifiers.sprmTWidthIndent: return ParseFtsWWidth_Indent(value);
                case SinglePropertyModifiers.sprmTCellBrcType: return ParseTCellBrcTypeOperand(value);
                case SinglePropertyModifiers.sprmTFBiDi90: return ParseBool(value);
                case SinglePropertyModifiers.sprmTFNoAllowOverlap: return ParseBool(value);
                case SinglePropertyModifiers.sprmTFCantSplit: return ParseBool(value);
                case SinglePropertyModifiers.sprmTPropRMark: return ParsePropRMarkOperand(value);
                case SinglePropertyModifiers.sprmTWall: return ParseBool(value);
                case SinglePropertyModifiers.sprmTIpgp: return ParseUInt32(value);
                case SinglePropertyModifiers.sprmTCnf: return ParseCNFOperand(value);
                case SinglePropertyModifiers.sprmTDefTableShdRaw: return ParseDefTableShdOperand(value);
                case SinglePropertyModifiers.sprmTDefTableShdRaw2nd: return ParseDefTableShdOperand(value);
                case SinglePropertyModifiers.sprmTDefTableShdRaw3rd: return ParseDefTableShdOperand(value);
                case SinglePropertyModifiers.sprmTRsid: return ParseInt32(value);
                case SinglePropertyModifiers.sprmTCellVertAlignStyle: return ParseVerticalAlign(value);
                case SinglePropertyModifiers.sprmTCellNoWrapStyle: return ParseBool(value);
                case SinglePropertyModifiers.sprmTCellBrcTopStyle: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellBrcBottomStyle: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellBrcLeftStyle: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellBrcRightStyle: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellBrcInsideHStyle: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellBrcInsideVStyle: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellBrcTL2BRStyle: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellBrcTR2BLStyle: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmTCellShdStyle: return ParseSHDOperand(value);
                case SinglePropertyModifiers.sprmTCHorzBands: return ParseByte(value);
                case SinglePropertyModifiers.sprmTCVertBands: return ParseByte(value);
                case SinglePropertyModifiers.sprmTJc: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmScnsPgn: return ParseCNS(value);
                case SinglePropertyModifiers.sprmSiHeadingPgn: return ParseByte(value);
                case SinglePropertyModifiers.sprmSDxaColWidth: return ParseSDxaColWidthOperand(value);
                case SinglePropertyModifiers.sprmSDxaColSpacing: return ParseSDxaColSpacingOperand(value);
                case SinglePropertyModifiers.sprmSFEvenlySpaced: return ParseBool(value);
                case SinglePropertyModifiers.sprmSFProtected: return ParseBool(value);
                case SinglePropertyModifiers.sprmSDmBinFirst: return ParseSDmBinOperand(value);
                case SinglePropertyModifiers.sprmSDmBinOther: return ParseSDmBinOperand(value);
                case SinglePropertyModifiers.sprmSBkc: return ParseSBkcOperand(value);
                case SinglePropertyModifiers.sprmSFTitlePage: return ParseBool(value);
                case SinglePropertyModifiers.sprmSCcolumns: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSDxaColumns: return ParseXAS_nonNeg(value);
                case SinglePropertyModifiers.sprmSNfcPgn: return ParseMSONFC(value);
                case SinglePropertyModifiers.sprmSFPgnRestart: return ParseBool(value);
                case SinglePropertyModifiers.sprmSFEndnote: return ParseBool(value);
                case SinglePropertyModifiers.sprmSLnc: return ParseSLncOperand(value);
                case SinglePropertyModifiers.sprmSNLnnMod: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSDxaLnn: return ParseXAS_nonNeg(value);
                case SinglePropertyModifiers.sprmSDyaHdrTop: return ParseYAS_nonNeg(value);
                case SinglePropertyModifiers.sprmSDyaHdrBottom: return ParseYAS_nonNeg(value);
                case SinglePropertyModifiers.sprmSLBetween: return ParseBool(value);
                case SinglePropertyModifiers.sprmSVjc: return ParseVic(value);
                case SinglePropertyModifiers.sprmSLnnMin: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSPgnStart97: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSBOrientation: return ParseSBOrientationOperan(value);
                case SinglePropertyModifiers.sprmSXaPage: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSYaPage: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSDxaLeft: return ParseXAS_nonNeg(value);
                case SinglePropertyModifiers.sprmSDxaRight: return ParseXAS_nonNeg(value);
                case SinglePropertyModifiers.sprmSDyaTop: return ParseYAS(value);
                case SinglePropertyModifiers.sprmSDyaBottom: return ParseYAS(value);
                case SinglePropertyModifiers.sprmSDzaGutter: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSDmPaperReq: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSFBiDi: return ParseBool(value);
                case SinglePropertyModifiers.sprmSFRTLGutter: return ParseBool(value);
                case SinglePropertyModifiers.sprmSBrcTop80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmSBrcLeft80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmSBrcBottom80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmSBrcRight80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmSPgbProp: return ParseSPgbPropOperand(value);
                case SinglePropertyModifiers.sprmSDxtCharSpace: return ParseInt32(value);
                case SinglePropertyModifiers.sprmSDyaLinePitch: return ParseYAS(value);
                case SinglePropertyModifiers.sprmSClm: return ParseSClmOperand(value);
                case SinglePropertyModifiers.sprmSTextFlow: return ParseMSOTXFL(value);
                case SinglePropertyModifiers.sprmSBrcTop: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmSBrcLeft: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmSBrcBottom: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmSBrcRight: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmSWall: return ParseBool(value);
                case SinglePropertyModifiers.sprmSRsid: return ParseInt32(value);
                case SinglePropertyModifiers.sprmSFpc: return ParseSFpcOperand(value);
                case SinglePropertyModifiers.sprmSRncFtn: return ParseRnc(value);
                case SinglePropertyModifiers.sprmSRncEdn: return ParseRnc(value);
                case SinglePropertyModifiers.sprmSNFtn: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSNfcFtnRef: return ParseMSONFC(value);
                case SinglePropertyModifiers.sprmSNEdn: return ParseUInt16(value);
                case SinglePropertyModifiers.sprmSNfcEdnRef: return ParseMSONFC(value);
                case SinglePropertyModifiers.sprmSPropRMark: return ParsePropRMarkOperand(value);
                case SinglePropertyModifiers.sprmSPgnStart: return ParseUInt32(value);
                case SinglePropertyModifiers.sprmPicBrcTop80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPicBrcLeft80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPicBrcBottom80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPicBrcRight80: return ParseBrc80(value);
                case SinglePropertyModifiers.sprmPicBrcTop: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmPicBrcLeft: return ParseBrcOperand(value);
                case SinglePropertyModifiers.sprmPicBrcBottom: return ParseBrcOperand(value);
                default:
                    throw new WordFileReaderException("Unknow sprm value");
            }
        }

        const float dptConstant = 0.125f;
        const float dxaConstant = 0.125f; //???

        static object ParseBool(byte[] value) { return value[0] != 0; }

        static object ParseBrc80(byte[] value) 
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("LineWidth", value[0] * dptConstant);
            result.Add("BorderType", (int)value[1]);
            result.Add("Color", (int)value[2]);
            result.Add("Space", (value[3] & 0x1F) * dptConstant);
            result.Add("Shadow", (value[3] & 0x20) != 0);
            result.Add("Frame", (value[3] & 0x40) != 0);
            return result;
        }

        static object ParseBrcCvOperand(byte[] value)  // var
        {
            return new ColorRef[] {
                ColorRef.FromBytes(value, 0), ColorRef.FromBytes(value, 4),
                ColorRef.FromBytes(value, 8), ColorRef.FromBytes(value, 12)
            };
        }

        static object ParseBrcOperand(byte[] value) 
        {   // var
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("Color", ColorRef.FromBytes(value, 0));
            result.Add("LineWidth", value[4] * dptConstant);
            result.Add("BorderType", (int)value[5]);
            result.Add("Space", (value[6] & 0x1F) * dptConstant);
            result.Add("Shadow", (value[6] & 0x20) != 0);
            result.Add("Frame", (value[6] & 0x40) != 0);
            return result;
        }

        static object ParseByte(byte[] value) { return value[0]; }

        static object ParseCellHideMarkOperand(byte[] value) // var
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("Range", new int[] {value[0], value[1]});
            result.Add("NoHeight", value[2] != 0);
            return result;
        }

        static object ParseCellRangeFitText(byte[] value) 
        { 
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("Range", new int[] { value[0], value[1] });
            result.Add("FitText", value[2] != 0);
            return result;
        }

        static object ParseCellRangeNoWrap(byte[] value) 
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("Range", new int[] { value[0], value[1] });
            result.Add("NoWrap", value[2] != 0);
            return result;
        }

        static object ParseCellRangeTextFlow(byte[] value)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("Range", new int[] { value[0], value[1] });
            result.Add("TextFlow", BitConverter.ToUInt16(value, 2));
            return result;
        }
        static object ParseCellRangeVertAlign(byte[] value)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("Range", new int[] { value[0], value[1] });
            result.Add("VertAlign", (int)value[2]);
            return result;
        }

        static object ParseCFitTextOperand(byte[] value)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("FitText", BitConverter.ToInt32(value, 0) * dxaConstant);
            result.Add("FitTextID", BitConverter.ToInt32(value, 2));
            return result;
        }

        static object ParseCMajorityOperand(byte[] value)
        {
            MemoryStream ms = new MemoryStream(value, false);
            int read = 0;
            Prl[] prls = BasicTypesReader.ReadPrls(ms, value.Length, ref read);
            Dictionary<string, object> result = new Dictionary<string, object>();
            foreach (Prl prl in prls)
            {
                string name = SinglePropertyModifiers.map[prl.sprm.sprm];
                result.Add(name.Substring(4), ParseValue(prl.sprm.sprm, prl.operand));
            }
            return result;
        }

        static object ParseCNFOperand(byte[] value)
        {
            MemoryStream ms = new MemoryStream(value, false);
            int read = 0;
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("$FormattingCondition",
                (int)BitConverter.ToUInt16(ReadUtils.ReadExact(ms, ReadUtils.WordSize, ref read), 0));
            Prl[] prls = BasicTypesReader.ReadPrls(ms, value.Length - read, ref read);
            foreach (Prl prl in prls)
            {
                string name = SinglePropertyModifiers.map[prl.sprm.sprm];
                result.Add(name.Substring(4), ParseValue(prl.sprm.sprm, prl.operand));
            }
            return result;
        }

        static object ParseCNS(byte[] value) 
        { 
            return (int)value[0]; 
        }

        static object ParseCOLORREF(byte[] value) { return ColorRef.FromBytes(value, 0); }
        static object ParseCP(byte[] value) { return BitConverter.ToUInt32(value, 0); }

        static object ParseCSSAOperand(byte[] value)
        {
            throw new NotSupportedException();
        }

        static object ParseCSymbolOperand(byte[] value)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("FontIndex", (int)BitConverter.ToUInt16(value, 0));
            result.Add("Symbol", (char)BitConverter.ToUInt16(value, 2));
            return result;
        }

        static object ParseDCS(byte[] value)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("DropCap", (int)(value[0] & 0x1F));
            result.Add("DropLines", (int)(value[0] >> 5));
            return result;
        }

        static object ParseDefTableShd80Operand(byte[] value) 
        {
            // array of Shd80
            throw new NotSupportedException();
        }

        static object ParseDefTableShdOperand(byte[] value)
        {
            throw new NotSupportedException();
        }

        static object ParseDispFldRmOperand(byte[] value) 
        {
            throw new NotSupportedException();
        }

        static object ParseDTTM(byte[] value)
        {
            return BasicTypesReader.ParseDTTM(BitConverter.ToUInt32(value, 0));
        }

        static object ParseFarEastLayoutOperand(byte[] value)
        {
            throw new NotSupportedException();
        }

        static object ParseFFM(byte[] value)
        {
            return (int)value[0];
        }

        static object ParseFrameTextFlowOperand(byte[] value)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("Vertical", (value[0] & 0x01) != 0);
            result.Add("Backwards", (value[0] & 0x02) != 0);
            result.Add("RotateFont", (value[0] & 0x04) != 0);
            return result;
        }

        static object ParseFtsWWidth_Indent(byte[] value)
        {
            throw new NotSupportedException();
        }

        static object ParseFtsWWidth_Table(byte[] value)
        {
            throw new NotSupportedException();
        }

        static object ParseFtsWWidth_TablePart(byte[] value)
        {
            throw new NotSupportedException();
        }

        static object ParseHresiOperand(byte[] value)
        {
            Dictionary<string, object> result = new Dictionary<string, object>();
            result.Add("WordBreakingMethod", (int)value[0]);
            result.Add("Character", (char)value[1]);
            return result;
        }

        static object ParseIco(byte[] value) { return (int)value[0]; }

        static object ParseInt16(byte[] value) { return BitConverter.ToInt16(value, 0); }

        static object ParseInt32(byte[] value) { return BitConverter.ToInt32(value, 0); }

        static object ParseItcFirstLim(byte[] value) 
        { 
            return new int[] { value[0], value[1] }; 
        }

        static object ParseKul(byte[] value) { return (int)value[0]; }

        static object ParseLBCOperand(byte[] value) { return (int)value[0]; }

        static object ParseLID(byte[] value) { return (int)BitConverter.ToUInt16(value, 0); }

        static object ParseLSPD(byte[] value)
        {
            throw new NotSupportedException();
        }

        static object ParseMathPrOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseMSONFC(byte[] value) { throw new NotSupportedException(); }
        static object ParseMSOTXFL(byte[] value) { throw new NotSupportedException(); }
        static object ParseNumRMOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParsePbiGrfOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParsePChgTabsOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParsePChgTabsPapxOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParsePositionCodeOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParsePropRMarkOperand(byte[] value) { throw new NotSupportedException(); } // var
        static object ParsePTIstdInfoOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseRnc(byte[] value) { throw new NotSupportedException(); }
        static object ParseSBkcOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseSBOrientationOperan(byte[] value) { throw new NotSupportedException(); }

        static object ParseSByte(byte[] value) { return (sbyte)value[0]; }

        static object ParseSClmOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseSDmBinOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseSDxaColSpacingOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseSDxaColWidthOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseSFpcOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseShd80(byte[] value) { throw new NotSupportedException(); }
        static object ParseSHDOperand(byte[] value) { throw new NotSupportedException(); } //var
        static object ParseSLncOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseSPgbPropOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseSPPOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTableBordersOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTableBordersOperand80(byte[] value) { throw new NotSupportedException(); }
        static object ParseTableBrc80Operand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTableBrcOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTableCellWidthOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTableShadeOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTCellBrcTypeOperand(byte[] value) { throw new NotSupportedException(); } // var
        static object ParseTDefTableOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTDxaColOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTInsertOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseTLP(byte[] value) { throw new NotSupportedException(); }

        static object ParseToggleOperand(byte[] value) { return (ToggleOperand)value[0]; }

        static object ParseUInt16(byte[] value) { return BitConverter.ToUInt16(value, 0); }

        static object ParseUInt32(byte[] value) { return BitConverter.ToUInt32(value, 0); }

        static object ParseVerticalAlign(byte[] value) { throw new NotSupportedException(); }
        static object ParseVertMergeOperand(byte[] value) { throw new NotSupportedException(); }
        static object ParseVic(byte[] value) { throw new NotSupportedException(); }
        static object ParseWHeightAbs(byte[] value) { throw new NotSupportedException(); }
        static object ParseXAS(byte[] value) { throw new NotSupportedException(); }
        static object ParseXAS_nonNeg(byte[] value) { throw new NotSupportedException(); }
        static object ParseXAS_plusOne(byte[] value) { throw new NotSupportedException(); }
        static object ParseYAS(byte[] value) { throw new NotSupportedException(); }
        static object ParseYAS_nonNeg(byte[] value) { throw new NotSupportedException(); }
        static object ParseYAS_plusOne(byte[] value) { throw new NotSupportedException(); }
    }


    public class ColorRef
    {
        public bool IsAuto { get; set; }
        public byte R { get; set; }
        public byte G { get; set; }
        public byte B { get; set; }

        internal static ColorRef FromBytes(byte[] value, int offset)
        {
            ColorRef c = new ColorRef();
            c.R = value[0];
            c.G = value[1];
            c.B = value[2];
            c.IsAuto = value[3] != 0;
            return c;
        }
    }

    public class FtsWidth
    {
        public FtsWidthType Type { get; set; }
        public float Value { get; set; }

        internal static FtsWidth FromBytes(byte[] value, int offset)
        {
            throw new NotImplementedException();
        }
    }

    public enum FtsWidthType
    {
        None = 0,
        Auto = 1,
        Percent = 2,
        Dxa = 3,
        DxaSys = 19
    }

    public enum ToggleOperand
    {
        Reset = 0,
        Set = 1,
        Positive = 0x80,
        Negative = 0x81
    }
}
