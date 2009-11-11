// Author: notmasteryet; License: Ms-PL
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace WordFileReader
{
    static class SinglePropertyModifiers
    {        
        internal const ushort sprmCFRMarkDel = 0x0800; // 0x00
        internal const ushort sprmCFRMarkIns = 0x0801; // 0x01
        internal const ushort sprmCFFldVanish = 0x0802; // 0x02
        internal const ushort sprmCPicLocation = 0x6A03; // 0x03
        internal const ushort sprmCIbstRMark = 0x4804; // 0x04
        internal const ushort sprmCDttmRMark = 0x6805; // 0x05
        internal const ushort sprmCFData = 0x0806; // 0x06
        internal const ushort sprmCIdslRMark = 0x4807; // 0x07
        internal const ushort sprmCSymbol = 0x6A09; // 0x09
        internal const ushort sprmCFOle2 = 0x080A; // 0x0A
        internal const ushort sprmCHighlight = 0x2A0C; // 0x0C
        internal const ushort sprmCFWebHidden = 0x0811; // 0x11
        internal const ushort sprmCRsidProp = 0x6815; // 0x15
        internal const ushort sprmCRsidText = 0x6816; // 0x16
        internal const ushort sprmCRsidRMDel = 0x6817; // 0x17
        internal const ushort sprmCFSpecVanish = 0x0818; // 0x18
        internal const ushort sprmCFMathPr = 0xC81A; // 0x1A
        internal const ushort sprmCIstd = 0x4A30; // 0x30
        internal const ushort sprmCIstdPermute = 0xCA31; // 0x31
        internal const ushort sprmCPlain = 0x2A33; // 0x33
        internal const ushort sprmCKcd = 0x2A34; // 0x34
        internal const ushort sprmCFBold = 0x0835; // 0x35
        internal const ushort sprmCFItalic = 0x0836; // 0x36
        internal const ushort sprmCFStrike = 0x0837; // 0x37
        internal const ushort sprmCFOutline = 0x0838; // 0x38
        internal const ushort sprmCFShadow = 0x0839; // 0x39
        internal const ushort sprmCFSmallCaps = 0x083A; // 0x3A
        internal const ushort sprmCFCaps = 0x083B; // 0x3B
        internal const ushort sprmCFVanish = 0x083C; // 0x3C
        internal const ushort sprmCKul = 0x2A3E; // 0x3E
        internal const ushort sprmCDxaSpace = 0x8840; // 0x40
        internal const ushort sprmCIco = 0x2A42; // 0x42
        internal const ushort sprmCHps = 0x4A43; // 0x43
        internal const ushort sprmCHpsPos = 0x4845; // 0x45
        internal const ushort sprmCMajority = 0xCA47; // 0x47
        internal const ushort sprmCIss = 0x2A48; // 0x48
        internal const ushort sprmCHpsKern = 0x484B; // 0x4B
        internal const ushort sprmCHresi = 0x484E; // 0x4E
        internal const ushort sprmCRgFtc0 = 0x4A4F; // 0x4F
        internal const ushort sprmCRgFtc1 = 0x4A50; // 0x50
        internal const ushort sprmCRgFtc2 = 0x4A51; // 0x51
        internal const ushort sprmCCharScale = 0x4852; // 0x52
        internal const ushort sprmCFDStrike = 0x2A53; // 0x53
        internal const ushort sprmCFImprint = 0x0854; // 0x54
        internal const ushort sprmCFSpec = 0x0855; // 0x55
        internal const ushort sprmCFObj = 0x0856; // 0x56
        internal const ushort sprmCPropRMark90 = 0xCA57; // 0x57
        internal const ushort sprmCFEmboss = 0x0858; // 0x58
        internal const ushort sprmCSfxText = 0x2859; // 0x59
        internal const ushort sprmCFBiDi = 0x085A; // 0x5A
        internal const ushort sprmCFBoldBi = 0x085C; // 0x5C
        internal const ushort sprmCFItalicBi = 0x085D; // 0x5D
        internal const ushort sprmCFtcBi = 0x4A5E; // 0x5E
        internal const ushort sprmCLidBi = 0x485F; // 0x5F
        internal const ushort sprmCIcoBi = 0x4A60; // 0x60
        internal const ushort sprmCHpsBi = 0x4A61; // 0x61
        internal const ushort sprmCDispFldRMark = 0xCA62; // 0x62
        internal const ushort sprmCIbstRMarkDel = 0x4863; // 0x63
        internal const ushort sprmCDttmRMarkDel = 0x6864; // 0x64
        internal const ushort sprmCBrc80 = 0x6865; // 0x65
        internal const ushort sprmCShd80 = 0x4866; // 0x66
        internal const ushort sprmCIdslRMarkDel = 0x4867; // 0x67
        internal const ushort sprmCFUsePgsuSettings = 0x0868; // 0x68
        internal const ushort sprmCRgLid0_80 = 0x486D; // 0x6D
        internal const ushort sprmCRgLid1_80 = 0x486E; // 0x6E
        internal const ushort sprmCIdctHint = 0x286F; // 0x6F
        internal const ushort sprmCCv = 0x6870; // 0x70
        internal const ushort sprmCShd = 0xCA71; // 0x71
        internal const ushort sprmCBrc = 0xCA72; // 0x72
        internal const ushort sprmCRgLid0 = 0x4873; // 0x73
        internal const ushort sprmCRgLid1 = 0x4874; // 0x74
        internal const ushort sprmCFNoProof = 0x0875; // 0x75
        internal const ushort sprmCFitText = 0xCA76; // 0x76
        internal const ushort sprmCCvUl = 0x6877; // 0x77
        internal const ushort sprmCFELayout = 0xCA78; // 0x78
        internal const ushort sprmCLbcCRJ = 0x2879; // 0x79
        internal const ushort sprmCFComplexScripts = 0x0882; // 0x82
        internal const ushort sprmCWall = 0x2A83; // 0x83
        internal const ushort sprmCCnf = 0xCA85; // 0x85
        internal const ushort sprmCNeedFontFixup = 0x2A86; // 0x86
        internal const ushort sprmCPbiIBullet = 0x6887; // 0x87
        internal const ushort sprmCPbiGrf = 0x4888; // 0x88
        internal const ushort sprmCPropRMark = 0xCA89; // 0x89
        internal const ushort sprmCFSdtVanish = 0x2A90; // 0x90
        internal const ushort sprmPIstd = 0x4600; // 0x00
        internal const ushort sprmPIstdPermute = 0xC601; // 0x01
        internal const ushort sprmPIncLvl = 0x2602; // 0x02
        internal const ushort sprmPJc80 = 0x2403; // 0x03
        internal const ushort sprmPFKeep = 0x2405; // 0x05
        internal const ushort sprmPFKeepFollow = 0x2406; // 0x06
        internal const ushort sprmPFPageBreakBefore = 0x2407; // 0x07
        internal const ushort sprmPIlvl = 0x260A; // 0x0A
        internal const ushort sprmPIlfo = 0x460B; // 0x0B
        internal const ushort sprmPFNoLineNumb = 0x240C; // 0x0C
        internal const ushort sprmPChgTabsPapx = 0xC60D; // 0x0D
        internal const ushort sprmPDxaRight80 = 0x840E; // 0x0E
        internal const ushort sprmPDxaLeft80 = 0x840F; // 0x0F
        internal const ushort sprmPNest80 = 0x4610; // 0x10
        internal const ushort sprmPDxaLeft180 = 0x8411; // 0x11
        internal const ushort sprmPDyaLine = 0x6412; // 0x12
        internal const ushort sprmPDyaBefore = 0xA413; // 0x13
        internal const ushort sprmPDyaAfter = 0xA414; // 0x14
        internal const ushort sprmPChgTabs = 0xC615; // 0x15
        internal const ushort sprmPFInTable = 0x2416; // 0x16
        internal const ushort sprmPFTtp = 0x2417; // 0x17
        internal const ushort sprmPDxaAbs = 0x8418; // 0x18
        internal const ushort sprmPDyaAbs = 0x8419; // 0x19
        internal const ushort sprmPDxaWidth = 0x841A; // 0x1A
        internal const ushort sprmPPc = 0x261B; // 0x1B
        internal const ushort sprmPWr = 0x2423; // 0x23
        internal const ushort sprmPBrcTop80 = 0x6424; // 0x24
        internal const ushort sprmPBrcLeft80 = 0x6425; // 0x25
        internal const ushort sprmPBrcBottom80 = 0x6426; // 0x26
        internal const ushort sprmPBrcRight80 = 0x6427; // 0x27
        internal const ushort sprmPBrcBetween80 = 0x6428; // 0x28
        internal const ushort sprmPBrcBar80 = 0x6629; // 0x29
        internal const ushort sprmPFNoAutoHyph = 0x242A; // 0x2A
        internal const ushort sprmPWHeightAbs = 0x442B; // 0x2B
        internal const ushort sprmPDcs = 0x442C; // 0x2C
        internal const ushort sprmPShd80 = 0x442D; // 0x2D
        internal const ushort sprmPDyaFromText = 0x842E; // 0x2E
        internal const ushort sprmPDxaFromText = 0x842F; // 0x2F
        internal const ushort sprmPFLocked = 0x2430; // 0x30
        internal const ushort sprmPFWidowControl = 0x2431; // 0x31
        internal const ushort sprmPFKinsoku = 0x2433; // 0x33
        internal const ushort sprmPFWordWrap = 0x2434; // 0x34
        internal const ushort sprmPFOverflowPunct = 0x2435; // 0x35
        internal const ushort sprmPFTopLinePunct = 0x2436; // 0x36
        internal const ushort sprmPFAutoSpaceDE = 0x2437; // 0x37
        internal const ushort sprmPFAutoSpaceDN = 0x2438; // 0x38
        internal const ushort sprmPWAlignFont = 0x4439; // 0x39
        internal const ushort sprmPFrameTextFlow = 0x443A; // 0x3A
        internal const ushort sprmPOutLvl = 0x2640; // 0x40
        internal const ushort sprmPFBiDi = 0x2441; // 0x41
        internal const ushort sprmPFNumRMIns = 0x2443; // 0x43
        internal const ushort sprmPNumRM = 0xC645; // 0x45
        internal const ushort sprmPHugePapx = 0x6646; // 0x46
        internal const ushort sprmPFUsePgsuSettings = 0x2447; // 0x47
        internal const ushort sprmPFAdjustRight = 0x2448; // 0x48
        internal const ushort sprmPItap = 0x6649; // 0x49
        internal const ushort sprmPDtap = 0x664A; // 0x4A
        internal const ushort sprmPFInnerTableCell = 0x244B; // 0x4B
        internal const ushort sprmPFInnerTtp = 0x244C; // 0x4C
        internal const ushort sprmPShd = 0xC64D; // 0x4D
        internal const ushort sprmPBrcTop = 0xC64E; // 0x4E
        internal const ushort sprmPBrcLeft = 0xC64F; // 0x4F
        internal const ushort sprmPBrcBottom = 0xC650; // 0x50
        internal const ushort sprmPBrcRight = 0xC651; // 0x51
        internal const ushort sprmPBrcBetween = 0xC652; // 0x52
        internal const ushort sprmPBrcBar = 0xC653; // 0x53
        internal const ushort sprmPDxcRight = 0x4455; // 0x55
        internal const ushort sprmPDxcLeft = 0x4456; // 0x56
        internal const ushort sprmPDxcLeft1 = 0x4457; // 0x57
        internal const ushort sprmPDylBefore = 0x4458; // 0x58
        internal const ushort sprmPDylAfter = 0x4459; // 0x59
        internal const ushort sprmPFOpenTch = 0x245A; // 0x5A
        internal const ushort sprmPFDyaBeforeAuto = 0x245B; // 0x5B
        internal const ushort sprmPFDyaAfterAuto = 0x245C; // 0x5C
        internal const ushort sprmPDxaRight = 0x845D; // 0x5D
        internal const ushort sprmPDxaLeft = 0x845E; // 0x5E
        internal const ushort sprmPNest = 0x465F; // 0x5F
        internal const ushort sprmPDxaLeft1 = 0x8460; // 0x60
        internal const ushort sprmPJc = 0x2461; // 0x61
        internal const ushort sprmPFNoAllowOverlap = 0x2462; // 0x62
        internal const ushort sprmPWall = 0x2664; // 0x64
        internal const ushort sprmPIpgp = 0x6465; // 0x65
        internal const ushort sprmPCnf = 0xC666; // 0x66
        internal const ushort sprmPRsid = 0x6467; // 0x67
        internal const ushort sprmPIstdListPermute = 0xC669; // 0x69
        internal const ushort sprmPTableProps = 0x646B; // 0x6B
        internal const ushort sprmPTIstdInfo = 0xC66C; // 0x6C
        internal const ushort sprmPFContextualSpacing = 0x246D; // 0x6D
        internal const ushort sprmPPropRMark = 0xC66F; // 0x6F
        internal const ushort sprmPFMirrorIndents = 0x2470; // 0x70
        internal const ushort sprmPTtwo = 0x2471; // 0x71
        internal const ushort sprmTJc90 = 0x5400; // 0x00
        internal const ushort sprmTDxaLeft = 0x9601; // 0x01
        internal const ushort sprmTDxaGapHalf = 0x9602; // 0x02
        internal const ushort sprmTFCantSplit90 = 0x3403; // 0x03
        internal const ushort sprmTTableHeader = 0x3404; // 0x04
        internal const ushort sprmTTableBorders80 = 0xD605; // 0x05
        internal const ushort sprmTDyaRowHeight = 0x9407; // 0x07
        internal const ushort sprmTDefTable = 0xD608; // 0x08
        internal const ushort sprmTDefTableShd80 = 0xD609; // 0x09
        internal const ushort sprmTTlp = 0x740A; // 0x0A
        internal const ushort sprmTFBiDi = 0x560B; // 0x0B
        internal const ushort sprmTDefTableShd3rd = 0xD60C; // 0x0C
        internal const ushort sprmTPc = 0x360D; // 0x0D
        internal const ushort sprmTDxaAbs = 0x940E; // 0x0E
        internal const ushort sprmTDyaAbs = 0x940F; // 0x0F
        internal const ushort sprmTDxaFromText = 0x9410; // 0x10
        internal const ushort sprmTDyaFromText = 0x9411; // 0x11
        internal const ushort sprmTDefTableShd = 0xD612; // 0x12
        internal const ushort sprmTTableBorders = 0xD613; // 0x13
        internal const ushort sprmTTableWidth = 0xF614; // 0x14
        internal const ushort sprmTFAutofit = 0x3615; // 0x15
        internal const ushort sprmTDefTableShd2nd = 0xD616; // 0x16
        internal const ushort sprmTWidthBefore = 0xF617; // 0x17
        internal const ushort sprmTWidthAfter = 0xF618; // 0x18
        internal const ushort sprmTFKeepFollow = 0x3619; // 0x19
        internal const ushort sprmTBrcTopCv = 0xD61A; // 0x1A
        internal const ushort sprmTBrcLeftCv = 0xD61B; // 0x1B
        internal const ushort sprmTBrcBottomCv = 0xD61C; // 0x1C
        internal const ushort sprmTBrcRightCv = 0xD61D; // 0x1D
        internal const ushort sprmTDxaFromTextRight = 0x941E; // 0x1E
        internal const ushort sprmTDyaFromTextBottom = 0x941F; // 0x1F
        internal const ushort sprmTSetBrc80 = 0xD620; // 0x20
        internal const ushort sprmTInsert = 0x7621; // 0x21
        internal const ushort sprmTDelete = 0x5622; // 0x22
        internal const ushort sprmTDxaCol = 0x7623; // 0x23
        internal const ushort sprmTMerge = 0x5624; // 0x24
        internal const ushort sprmTSplit = 0x5625; // 0x25
        internal const ushort sprmTTextFlow = 0x7629; // 0x29
        internal const ushort sprmTVertMerge = 0xD62B; // 0x2B
        internal const ushort sprmTVertAlign = 0xD62C; // 0x2C
        internal const ushort sprmTSetShd = 0xD62D; // 0x2D
        internal const ushort sprmTSetShdOdd = 0xD62E; // 0x2E
        internal const ushort sprmTSetBrc = 0xD62F; // 0x2F
        internal const ushort sprmTCellPadding = 0xD632; // 0x32
        internal const ushort sprmTCellSpacingDefault = 0xD633; // 0x33
        internal const ushort sprmTCellPaddingDefault = 0xD634; // 0x34
        internal const ushort sprmTCellWidth = 0xD635; // 0x35
        internal const ushort sprmTFitText = 0xF636; // 0x36
        internal const ushort sprmTFCellNoWrap = 0xD639; // 0x39
        internal const ushort sprmTIstd = 0x563A; // 0x3A
        internal const ushort sprmTCellPaddingStyle = 0xD63E; // 0x3E
        internal const ushort sprmTCellFHideMark = 0xD642; // 0x42
        internal const ushort sprmTSetShdTable = 0xD660; // 0x60
        internal const ushort sprmTWidthIndent = 0xF661; // 0x61
        internal const ushort sprmTCellBrcType = 0xD662; // 0x62
        internal const ushort sprmTFBiDi90 = 0x5664; // 0x64
        internal const ushort sprmTFNoAllowOverlap = 0x3465; // 0x65
        internal const ushort sprmTFCantSplit = 0x3466; // 0x66
        internal const ushort sprmTPropRMark = 0xD667; // 0x67
        internal const ushort sprmTWall = 0x3668; // 0x68
        internal const ushort sprmTIpgp = 0x7469; // 0x69
        internal const ushort sprmTCnf = 0xD66A; // 0x6A
        internal const ushort sprmTDefTableShdRaw = 0xD670; // 0x70
        internal const ushort sprmTDefTableShdRaw2nd = 0xD671; // 0x71
        internal const ushort sprmTDefTableShdRaw3rd = 0xD672; // 0x72
        internal const ushort sprmTRsid = 0x7479; // 0x79
        internal const ushort sprmTCellVertAlignStyle = 0x347C; // 0x7C
        internal const ushort sprmTCellNoWrapStyle = 0x347D; // 0x7D
        internal const ushort sprmTCellBrcTopStyle = 0xD47F; // 0x7F
        internal const ushort sprmTCellBrcBottomStyle = 0xD680; // 0x80
        internal const ushort sprmTCellBrcLeftStyle = 0xD681; // 0x81
        internal const ushort sprmTCellBrcRightStyle = 0xD682; // 0x82
        internal const ushort sprmTCellBrcInsideHStyle = 0xD683; // 0x83
        internal const ushort sprmTCellBrcInsideVStyle = 0xD684; // 0x84
        internal const ushort sprmTCellBrcTL2BRStyle = 0xD685; // 0x85
        internal const ushort sprmTCellBrcTR2BLStyle = 0xD686; // 0x86
        internal const ushort sprmTCellShdStyle = 0xD687; // 0x87
        internal const ushort sprmTCHorzBands = 0x3488; // 0x88
        internal const ushort sprmTCVertBands = 0x3489; // 0x89
        internal const ushort sprmTJc = 0x548A; // 0x8A
        internal const ushort sprmScnsPgn = 0x3000; // 0x00
        internal const ushort sprmSiHeadingPgn = 0x3001; // 0x01
        internal const ushort sprmSDxaColWidth = 0xF203; // 0x03
        internal const ushort sprmSDxaColSpacing = 0xF204; // 0x04
        internal const ushort sprmSFEvenlySpaced = 0x3005; // 0x05
        internal const ushort sprmSFProtected = 0x3006; // 0x06
        internal const ushort sprmSDmBinFirst = 0x5007; // 0x07
        internal const ushort sprmSDmBinOther = 0x5008; // 0x08
        internal const ushort sprmSBkc = 0x3009; // 0x09
        internal const ushort sprmSFTitlePage = 0x300A; // 0x0A
        internal const ushort sprmSCcolumns = 0x500B; // 0x0B
        internal const ushort sprmSDxaColumns = 0x900C; // 0x0C
        internal const ushort sprmSNfcPgn = 0x300E; // 0x0E
        internal const ushort sprmSFPgnRestart = 0x3011; // 0x11
        internal const ushort sprmSFEndnote = 0x3012; // 0x12
        internal const ushort sprmSLnc = 0x3013; // 0x13
        internal const ushort sprmSNLnnMod = 0x5015; // 0x15
        internal const ushort sprmSDxaLnn = 0x9016; // 0x16
        internal const ushort sprmSDyaHdrTop = 0xB017; // 0x17
        internal const ushort sprmSDyaHdrBottom = 0xB018; // 0x18
        internal const ushort sprmSLBetween = 0x3019; // 0x19
        internal const ushort sprmSVjc = 0x301A; // 0x1A
        internal const ushort sprmSLnnMin = 0x501B; // 0x1B
        internal const ushort sprmSPgnStart97 = 0x501C; // 0x1C
        internal const ushort sprmSBOrientation = 0x301D; // 0x1D
        internal const ushort sprmSXaPage = 0xB01F; // 0x1F
        internal const ushort sprmSYaPage = 0xB020; // 0x20
        internal const ushort sprmSDxaLeft = 0xB021; // 0x21
        internal const ushort sprmSDxaRight = 0xB022; // 0x22
        internal const ushort sprmSDyaTop = 0x9023; // 0x23
        internal const ushort sprmSDyaBottom = 0x9024; // 0x24
        internal const ushort sprmSDzaGutter = 0xB025; // twips.
        internal const ushort sprmSDmPaperReq = 0x5026; // 0x26
        internal const ushort sprmSFBiDi = 0x3228; // 0x28
        internal const ushort sprmSFRTLGutter = 0x322A; // 0x2A
        internal const ushort sprmSBrcTop80 = 0x702B; // 0x2B
        internal const ushort sprmSBrcLeft80 = 0x702C; // 0x2C
        internal const ushort sprmSBrcBottom80 = 0x702D; // 0x2D
        internal const ushort sprmSBrcRight80 = 0x702E; // 0x2E
        internal const ushort sprmSPgbProp = 0x522F; // 0x2F
        internal const ushort sprmSDxtCharSpace = 0x7030; // 0x30
        internal const ushort sprmSDyaLinePitch = 0x9031; // 0x31
        internal const ushort sprmSClm = 0x5032; // 0x32
        internal const ushort sprmSTextFlow = 0x5033; // 0x33
        internal const ushort sprmSBrcTop = 0xD234; // 0x34
        internal const ushort sprmSBrcLeft = 0xD235; // 0x35
        internal const ushort sprmSBrcBottom = 0xD236; // 0x36
        internal const ushort sprmSBrcRight = 0xD237; // 0x37
        internal const ushort sprmSWall = 0x3239; // 0x39
        internal const ushort sprmSRsid = 0x703A; // 0x3A
        internal const ushort sprmSFpc = 0x303B; // 0x3B
        internal const ushort sprmSRncFtn = 0x303C; // 0x3C
        internal const ushort sprmSRncEdn = 0x303E; // 0x3E
        internal const ushort sprmSNFtn = 0x503F; // 0x3F
        internal const ushort sprmSNfcFtnRef = 0x5040; // 0x40
        internal const ushort sprmSNEdn = 0x5041; // 0x41
        internal const ushort sprmSNfcEdnRef = 0x5042; // 0x42
        internal const ushort sprmSPropRMark = 0xD243; // 0x43
        internal const ushort sprmSPgnStart = 0x7044; // 0x44
        internal const ushort sprmPicBrcTop80 = 0x6C02; // 0x02
        internal const ushort sprmPicBrcLeft80 = 0x6C03; // 0x03
        internal const ushort sprmPicBrcBottom80 = 0x6C04; // 0x04
        internal const ushort sprmPicBrcRight80 = 0x6C05; // 0x05
        internal const ushort sprmPicBrcTop = 0xCE08; // 0x08
        internal const ushort sprmPicBrcLeft = 0xCE09; // 0x09
        internal const ushort sprmPicBrcBottom = 0xCE0A; // 0x0A

        internal static Dictionary<ushort, string> map;
        internal static Dictionary<byte, ushort> prm0Map;

        static SinglePropertyModifiers()
        {
            map = new Dictionary<ushort, string>();
            foreach (FieldInfo fi in typeof(SinglePropertyModifiers).GetFields(BindingFlags.NonPublic | BindingFlags.Static))
            {
                if (fi.IsLiteral && fi.Name.StartsWith("sprm"))
                {
                    map.Add((ushort)fi.GetValue(null), fi.Name);
                }
            }

            prm0Map = new Dictionary<byte, ushort>();
            prm0Map.Add(0x00, sprmCLbcCRJ);
            prm0Map.Add(0x04, sprmPIncLvl);
            prm0Map.Add(0x05, sprmPJc);
            prm0Map.Add(0x07, sprmPFKeep);
            prm0Map.Add(0x08, sprmPFKeepFollow);
            prm0Map.Add(0x09, sprmPFPageBreakBefore);
            prm0Map.Add(0x0C, sprmPIlvl);
            prm0Map.Add(0x0D, sprmPFMirrorIndents);
            prm0Map.Add(0x0E, sprmPFNoLineNumb);
            prm0Map.Add(0x0F, sprmPTtwo);
            prm0Map.Add(0x18, sprmPFInTable);
            prm0Map.Add(0x19, sprmPFTtp);
            prm0Map.Add(0x1D, sprmPPc);
            prm0Map.Add(0x25, sprmPWr);
            prm0Map.Add(0x2C, sprmPFNoAutoHyph);
            prm0Map.Add(0x32, sprmPFLocked);
            prm0Map.Add(0x33, sprmPFWidowControl);
            prm0Map.Add(0x35, sprmPFKinsoku);
            prm0Map.Add(0x36, sprmPFWordWrap);
            prm0Map.Add(0x37, sprmPFOverflowPunct);
            prm0Map.Add(0x38, sprmPFTopLinePunct);
            prm0Map.Add(0x39, sprmPFAutoSpaceDE);
            prm0Map.Add(0x3A, sprmPFAutoSpaceDN);
            prm0Map.Add(0x41, sprmCFRMarkDel);
            prm0Map.Add(0x42, sprmCFRMarkIns);
            prm0Map.Add(0x43, sprmCFFldVanish);
            prm0Map.Add(0x47, sprmCFData);
            prm0Map.Add(0x4B, sprmCFOle2);
            prm0Map.Add(0x4D, sprmCHighlight);
            prm0Map.Add(0x4E, sprmCFEmboss);
            prm0Map.Add(0x4F, sprmCSfxText);
            prm0Map.Add(0x50, sprmCFWebHidden);
            prm0Map.Add(0x51, sprmCFSpecVanish);
            prm0Map.Add(0x53, sprmCPlain);
            prm0Map.Add(0x55, sprmCFBold);
            prm0Map.Add(0x56, sprmCFItalic);
            prm0Map.Add(0x57, sprmCFStrike);
            prm0Map.Add(0x58, sprmCFOutline);
            prm0Map.Add(0x59, sprmCFShadow);
            prm0Map.Add(0x5A, sprmCFSmallCaps);
            prm0Map.Add(0x5B, sprmCFCaps);
            prm0Map.Add(0x5C, sprmCFVanish);
            prm0Map.Add(0x5E, sprmCKul);
            prm0Map.Add(0x62, sprmCIco);
            prm0Map.Add(0x68, sprmCIss);
            prm0Map.Add(0x73, sprmCFDStrike);
            prm0Map.Add(0x74, sprmCFImprint);
            prm0Map.Add(0x75, sprmCFSpec);
            prm0Map.Add(0x76, sprmCFObj);
            prm0Map.Add(0x78, sprmPOutLvl);
            prm0Map.Add(0x7B, sprmCFSdtVanish);
            prm0Map.Add(0x7C, sprmCNeedFontFixup);
            prm0Map.Add(0x7E, sprmPFNumRMIns);
        }

        internal static ushort GetSprmByName(string name)
        {
            FieldInfo fi = typeof(SinglePropertyModifiers).GetField(name,
                BindingFlags.NonPublic | BindingFlags.Static);

            if (fi != null && fi.IsLiteral)
                return (ushort)fi.GetValue(null);
            else
                throw new WordFileReaderException("Invalid SPRM name: " + name);
         }
    }
}
