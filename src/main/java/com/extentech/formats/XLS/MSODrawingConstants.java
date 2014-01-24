/*
 * --------- BEGIN COPYRIGHT NOTICE ---------
 * Copyright 2002-2012 Extentech Inc.
 * Copyright 2013 Infoteria America Corp.
 * 
 * This file is part of OpenXLS.
 * 
 * OpenXLS is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as
 * published by the Free Software Foundation, either version 3 of
 * the License, or (at your option) any later version.
 * 
 * OpenXLS is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU Lesser General Public License for more details.
 * 
 * You should have received a copy of the GNU Lesser General Public
 * License along with OpenXLS.  If not, see
 * <http://www.gnu.org/licenses/>.
 * ---------- END COPYRIGHT NOTICE ----------
 */
package com.extentech.formats.XLS;

/**
 * Msodrawing record-related constants
 * NOT A COMPLETE LIST!
 */
public class MSODrawingConstants
{
	// 20070910 KSC: constants for MSO record + atom ids (records keep atoms and other containers; atoms contain info and are kept inside containers)
	// common record header for both:  ver, inst, fbt, len; fbt deterimes record type (0xF000 to 0xFFFF)
	// MsoDrawingGroup
	public static final int MSOFBTDGGCONTAINER = 0xF000;                // Drawing Group Container (Msodrawinggroup)
	public static final int MSOFBTDGG = 0xF006;                        // Drawing Group Record (of Msodrawinggoup)
	public static final int MSOFBTCLSID = 0xF016;                    // Clipboard format
	public static final int MSOFBTOPT = 0xF00B;                        // Property Table Record for newly created shapes, array of FOPTEs
	public static final int MSOFBTCOLORMRU = 0xF11A;
	public static final int MSOFBTSPLITMENUCOLORS = 0xF11E;
	public static final int MSOFBTBSTORECONTAINER = 0xF001;            // Stores BLIPS (=pix) in a separate container
	public static final int MSOFBTBSE = 0xF007;                    // BLIP Store Entry Record
	public static final int MSOFBTCALLOUTRULE = 0xF017;            // One callout rule per callout shape
	public static final int MSOFBTBLIP = 0xF018;
	// MsoDrawing
	public static final int MSOFBTDGCONTAINER = 0xF002;                    // Drawing Container - (Msodrawing records contained in MsoDrawinggroup)
	public static final int MSOFBTDG = 0xF008;                        // Basic drawing info
	public static final int MSOFBTREGROUPITEMS = 0xF118;            // Mappings to reconstitute groups
	public static final int MSOFBTCOLORSCHEME = 0xF120;
	public static final int MSOFBTSECONDARYOPT = 0xF121;            // secondary opt block - "the property table msofbtOpt may be split into as many as 3 blocks"
	public static final int MSOFBTTERTIARYOPT = 0xF122;        // default properties of new shapes - only those props which differ from the per-property defaults are saved
	public static final int MSOFBTSPGRCONTAINER = 0xF003;            // Patriarch shape, with all non-bg non-deleted shapes inside it
	public static final int MSOFBTSPCONTAINER = 0xF004;                // Shape Container
	public static final int MSOFBTSPGR = 0xF009;                // Group-shape-specific info  (i.e. shapes that are groups)
	public static final int MSOFBTSP = 0xF00A;                    // A shape atom rec (inst= shape type) rec= shape ID + group of flags
	public static final int MSOFBTTEXTBOX = 0xF00C;                // if the shape has text
	public static final int MSOFBTCLIENTTEXTBOX = 0xF00D;        // for clipboard stream
	public static final int MSOFBTANCHOR = 0xF00E;                // Anchor or location fo a shape (if streamed to a clipboard)
	public static final int MSOFBTCHILDANCHOR = 0xF00F;            //   " ", if shape is a child of a group shape
	public static final int MSOFBTCLIENTANCHOR = 0xF010;        //   " ", for top-level shapes
	public static final int MSOFBTCLIENTDATA = 0xF011;            // content is determined by host
	public static final int MSOFBTCONNECTORRULE = 0xF012;    // connector rule
	public static final int MSOFBTALIGNRULE = 0xF013;
	public static final int MSOFBTARCRULE = 0xF014;
	public static final int MSOFBTCLIENTRULE = 0xF015;
	public static final int MSOFBTOLEOBJECT = 0xF11F;        //
	public static final int MSOFBTDELETEDPSPL = 0xF11D;
	public static final int MSOFBTSOLVERCONTAINER = 0xF005;    // 	the rules governing shapes; count of rules
	public static final int MSOFBTSELECTION = 0xF119;
	// MSOBI == encoded BLIP types
	public static final int msobiUNKNOWN = 0;
	public static final int msobiWMF = 0x216;
	public static final int msobiEMF = 0x3D4;
	public static final int msobiPICT = 0x542;
	public static final int msobiPNG = 0x6E0;
	public static final int msobiJFIF = 0x46A;
	public static final int msobiJPEG = msobiJFIF;
	public static final int msobiDIB = 0x7A8;
	public static final int msobiCLIENT = 0x800;
	//
	public static final int msofbtBlipFirst = 0xF018;            // used to calculate fbt of BLIP in MsofbtBSE

	// pertinent Property Table/MSOFBTOPT property id's
	public static final int msooptfLockAgainstGrouping = 127;    // BOOL - Do not group this shape
	public static final int msofbtlTxid = 128;                    // LONG - id for the text, value determined by the host
	//margins relative to shape's inscribed text rectangle (in EMUs)
	public static final int dxTextLeft = 129;    // 	LONG	1/10 inch
	public static final int dyTextTop = 130; // 	LONG	1/20 inch
	public static final int dxTextRight = 131;    // 	LONG 	1/10 inch
	public static final int dyTextBottom = 132;    //	LONG 	1/20 inch
	// How to anchor the text
	public static final int anchorText = 135; // 	MSOANCHOR	def= Top
	public static final int msofbttxdir = 139;                    // MSOTXDIR - Bi-Di Text direction
	public static final int msooptfFitTextToShape = 191;        // BOOL - Size text to fit shape size
	public static final int msooptpib = 260;                    // IMsoBlip	- id of Blip to display == imageIndex
	public static final int msooptpibName = 261;                // WCHAR - Blip File Name == imageName
	public static final int msooptpibFlags = 262;                // MSOBLIPFLAGS - Blip Flags
	public static final int msooptpictureActive = 319;            // Server is active (OLE objects only)default= false
	// fill attrbutes
	public static final int msooptFillType = 384;
	public static final int msooptfillColor = 385;                // MSOCLR - foreground color
	public static final int msooptFillOpacity = 386;
	public static final int msooptfillBackColor = 387;            // MSOCLR - background color
	public static final int msooptFillBackOpacity = 388;
	public static final int msooptFillCrMod = 389;    // foreground color of the fill for black-and-white display mode.
	public static final int msooptFillBlip = 390;
	public static final int msooptFillBlipName = 391;        // fillBlipName_complex -- specifies additional data for the fillBlipName property
	public static final int msooptFillBlipFlags = 392;        // specifies how to interpret the fillBlipName_complex property
	public static final int msooptFillWidth = 393;        // the width of the fill. This property applies only to texture, picture, and pattern fills.
	// A signed integer that specifies the width of the fill in units that are specified by the fillDztype property, as defined in section 2.3.7.24. If fillDztype equals msodztypeDefault, this value MUST be ignored. The default value for this property is 0x00000000.
	public static final int msooptFillHeight = 394;        // A signed integer that specifies the height of the fill in units that are specified by the fillDztype property, as defined in section 2.3.7.24. If fillDztype equals msodztypeDefault, this value MUST be ignored. The default value for this property is 0x00000000.
	public static final int msooptFillAngle = 395;        // A value of type FixedPoint, as specified in [MS-OSHARED] section 2.2.1.6, that specifies the angle of the gradient fill. Zero degrees represents a vertical vector from bottom to top. The default value for this property is 0x00000000.
	public static final int msooptFillFocus = 396;        // specifies the relative position of the last color in the shaded fill.
	public static final int msooptFillToLeft = 397;
	public static final int msooptFillToTop = 398;
	public static final int msooptFillToRight = 399;
	public static final int msooptFillToBottom = 400;        // A signed integer that specifies the left boundary, in EMUs, of the bounding rectangle of the shaded fill. If the Fill Style BooleanfillUseRect property, as defined in section 2.3.7.43, equals 0x0, this value MUST be ignored. The default value for this property is 0x00000000.
	public static final int msooptFillRectLeft = 401;        // A signed integer that specifies the top boundary, in EMUs, of the bounding rectangle of the shaded fill. If the Fill Style BooleanfillUseRect property, as defined in section 2.3.7.43, equals 0x0, this value MUST be ignored. The default value for this property is 0x00000000.
	public static final int msooptFillRectTop = 402;
	public static final int msooptFillRectRight = 403;
	public static final int msooptFillRectBottom = 404;
	public static final int msooptFillDztype = 405;        // An MSODZTYPE enumeration value, as defined in section 2.4.12, that specifies how the fillWidth, as defined in section 2.3.7.12, and fillHeight, as defined in section 2.3.7.13, properties are interpreted. The default value for this property is msodztypeDefault.
	public static final int msooptFillShadePreset = 406;    // A signed integer that specifies the preset colors of the gradient fill. This value MUST be from 0x00000088 through 0x0000009F, inclusive. if the fillShadeColors_complex property, as defined in section 2.3.7.27, exists, this value MUST be ignored. The default value for this property is 0x00000000.
	public static final int msooptFillShadeColors = 407;    // The number of bytes of data in the fillShadeColors_complex property. If opid.fComplex equals 0x0, this value MUST be 0x00000000. The default value for this property is 0x00000000.
	public static final int msooptFillOriginX = 408;
	public static final int msooptFillOriginY = 409;
	public static final int msooptFillShapeOriginX = 410;
	public static final int msooptFillShapeOriginY = 411;
	public static final int msooptFillShadeType = 412;    // : An MSOSHADETYPE record, as defined in section 2.2.50, that specifies how the shaded fill is computed. The default value for this property is msoshadeDefault.
	public static final int msooptFFilled = 443;
	// line attributes
	public static final int msooptfNoFillHitTest = 447;            // BOOL - Hit test a shape as though filled
	public static final int msooptlineColor = 448;                // MSOCLR - Color of line
	public static final int msooptlineWidth = 459;                // LONG - 1pt= 12700 EMUs
	public static final int msooptLineMiterLimit = 460;
	public static final int msooptLineStyle = 461;
	public static final int msooptLineDashing = 462;
	public static final int msooptLineDashStyle = 463;
	public static final int msooptLineStartArrowhead = 464;
	public static final int msooptLineEndArrowhead = 465;
	public static final int msooptLineStartArrowWidth = 466;
	public static final int msooptLineStartArrowLength = 467;
	public static final int msooptLineEndArrowWidth = 468;
	public static final int msooptLineEndArrowLength = 469;
	public static final int msooptLineJoinStyle = 470;
	public static final int msooptLineEndCapStyle = 471;
	public static final int msooptFArrowheadsOK = 507;
	public static final int msooptLine = 508;
	public static final int msooptFHitTestLine = 509;
	public static final int msooptLineFillShape = 510;
	public static final int msooptfNoLineDrawDash = 511;        // BOOL - Draw a dashed line if no line
	// Shadow attributes - many to add
	public static final int msooptfshadowColor = 513;            // MSOCLR - foreground color - default = 0x808080
	public static final int msooptfShadowObscured = 575;        // BOOL - Excel5-style Shadow
	public static final int msooptfBackground = 831;            // BOOL - If true, this is the background shape
	public static final int msooptwzName = 896;                    // WCHAR - Shape Name (present only if explicitly set in Named Range box)
	public static final int msooptwzDescription = 897;            // WCHAR - Alternate text
	// group shape attributes (selected)
	public static final int msooptidRegroup = 904;                // LONG - Regroup id
	public static final int msooptMetroBlob = 937;				/*The shape‘s 2007 representation in Office Open XML format.
	 																The actual data is a package in Office XML format, which can simply be opened as a zip file.
     																This zip file contains an XML file with the root element ―sp‖. 
     																Refer to the publically available Office Open XML documentation for more information about this data. 
     																In case we lose any property when converting a 2007 Office Art shape to 2003 shape, 
     																we use this blob to retrieve the original Office Art property data when opening the file in 2007. 
     																See Appendix F for more information*/
	// misc
	public static final int msooptGroupShapeProperties = 959;    // The Group Shape Boolean Properties specify a 32-bit field of Boolean properties for either a shape or a group.
	// ...
	// right now, we support GIF, JPG and PNG
	public static final int IMAGE_TYPE_GIF = 0;
	public static final int IMAGE_TYPE_EMF = 2;
	public static final int IMAGE_TYPE_WMF = 3;
	public static final int IMAGE_TYPE_PICT = 4;
	public static final int IMAGE_TYPE_JPG = 5;
	public static final int IMAGE_TYPE_PNG = 6;
	public static final int IMAGE_TYPE_DIB = 7;

	// Shape Types
	/**
	 * Internally, a shape type is defined as a fixed set of property values,
	 * the most important being the geometry of the shape (the pVertices property, etc.).
	 * Each shape stores in itself only those properties that differ from its shape type.
	 * When a shape is asked for a property that isn't in its local table,
	 * it looks in the shape type's table.
	 * If the shape type doesn't define a value for the property,
	 * then the property's default value is used.
	 */
	public static final int msosptMin = 0;
	public static final int msosptNotPrimitive = msosptMin;
	public static final int msosptRectangle = 1;
	public static final int msosptRoundRectangle = 2;
	public static final int msosptEllipse = 3;
	public static final int msosptDiamond = 4;
	public static final int msosptIsocelesTriangle = 5;
	public static final int msosptRightTriangle = 6;
	public static final int msosptParallelogram = 7;
	public static final int msosptTrapezoid = 8;
	public static final int msosptHexagon = 9;
	public static final int msosptOctagon = 10;
	public static final int msosptPlus = 11;
	public static final int msosptStar = 12;
	public static final int msosptArrow = 13;
	public static final int msosptThickArrow = 14;
	public static final int msosptHomePlate = 15;
	public static final int msosptCube = 16;
	public static final int msosptBalloon = 17;
	public static final int msosptSeal = 18;
	public static final int msosptArc = 19;
	public static final int msosptLine = 20;
	public static final int msosptPlaque = 21;
	public static final int msosptCan = 22;
	public static final int msosptDonut = 23;
	public static final int msosptTextSimple = 24;
	public static final int msosptTextOctagon = 25;
	public static final int msosptTextHexagon = 26;
	public static final int msosptTextCurve = 27;
	public static final int msosptTextWave = 28;
	public static final int msosptTextRing = 29;
	public static final int msosptTextOnCurve = 30;
	public static final int msosptTextOnRing = 31;
	public static final int msosptStraightConnector1 = 32;
	public static final int msosptBentConnector2 = 33;
	public static final int msosptBentConnector3 = 34;
	public static final int msosptBentConnector4 = 35;
	public static final int msosptBentConnector5 = 36;
	public static final int msosptCurvedConnector2 = 37;
	public static final int msosptCurvedConnector3 = 38;
	public static final int msosptCurvedConnector4 = 39;
	public static final int msosptCurvedConnector5 = 40;
	public static final int msosptCallout1 = 41;
	public static final int msosptCallout2 = 42;
	public static final int msosptCallout3 = 43;
	public static final int msosptAccentCallout1 = 44;
	public static final int msosptAccentCallout2 = 45;
	public static final int msosptAccentCallout3 = 46;
	public static final int msosptBorderCallout1 = 47;
	public static final int msosptBorderCallout2 = 48;
	public static final int msosptBorderCallout3 = 49;
	public static final int msosptAccentBorderCallout1 = 50;
	public static final int msosptAccentBorderCallout2 = 51;
	public static final int msosptAccentBorderCallout3 = 52;
	public static final int msosptRibbon = 53;
	public static final int msosptRibbon2 = 54;
	public static final int msosptChevron = 55;
	public static final int msosptPentagon = 56;
	public static final int msosptNoSmoking = 57;
	public static final int msosptSeal8 = 58;
	public static final int msosptSeal16 = 59;
	public static final int msosptSeal32 = 60;
	public static final int msosptWedgeRectCallout = 61;
	public static final int msosptWedgeRRectCallout = 62;
	public static final int msosptWedgeEllipseCallout = 63;
	public static final int msosptWave = 64;
	public static final int msosptFoldedCorner = 65;
	public static final int msosptLeftArrow = 66;
	public static final int msosptDownArrow = 67;
	public static final int msosptUpArrow = 68;
	public static final int msosptLeftRightArrow = 69;
	public static final int msosptUpDownArrow = 70;
	public static final int msosptIrregularSeal1 = 71;
	public static final int msosptIrregularSeal2 = 72;
	public static final int msosptLightningBolt = 73;
	public static final int msosptHeart = 74;
	public static final int msosptPictureFrame = 75;
	public static final int msosptQuadArrow = 76;
	public static final int msosptLeftArrowCallout = 77;
	public static final int msosptRightArrowCallout = 78;
	public static final int msosptUpArrowCallout = 79;
	public static final int msosptDownArrowCallout = 80;
	public static final int msosptLeftRightArrowCallout = 81;
	public static final int msosptUpDownArrowCallout = 82;
	public static final int msosptQuadArrowCallout = 83;
	public static final int msosptBevel = 84;
	public static final int msosptLeftBracket = 85;
	public static final int msosptRightBracket = 86;
	public static final int msosptLeftBrace = 87;
	public static final int msosptRightBrace = 88;
	public static final int msosptLeftUpArrow = 89;
	public static final int msosptBentUpArrow = 90;
	public static final int msosptBentArrow = 91;
	public static final int msosptSeal24 = 92;
	public static final int msosptStripedRightArrow = 93;
	public static final int msosptNotchedRightArrow = 94;
	public static final int msosptBlockArc = 95;
	public static final int msosptSmileyFace = 96;
	public static final int msosptVerticalScroll = 97;
	public static final int msosptHorizontalScroll = 98;
	public static final int msosptCircularArrow = 99;
	public static final int msosptNotchedCircularArrow = 100;
	public static final int msosptUturnArrow = 101;
	public static final int msosptCurvedRightArrow = 102;
	public static final int msosptCurvedLeftArrow = 103;
	public static final int msosptCurvedUpArrow = 104;
	public static final int msosptCurvedDownArrow = 105;
	public static final int msosptCloudCallout = 106;
	public static final int msosptEllipseRibbon = 107;
	public static final int msosptEllipseRibbon2 = 108;
	public static final int msosptFlowChartProcess = 109;
	public static final int msosptFlowChartDecision = 110;
	public static final int msosptFlowChartInputOutput = 111;
	public static final int msosptFlowChartPredefinedProcess = 112;
	public static final int msosptFlowChartInternalStorage = 113;
	public static final int msosptFlowChartDocument = 114;
	public static final int msosptFlowChartMultidocument = 115;
	public static final int msosptFlowChartTerminator = 116;
	public static final int msosptFlowChartPreparation = 117;
	public static final int msosptFlowChartManualInput = 118;
	public static final int msosptFlowChartManualOperation = 119;
	public static final int msosptFlowChartConnector = 120;
	public static final int msosptFlowChartPunchedCard = 121;
	public static final int msosptFlowChartPunchedTape = 122;
	public static final int msosptFlowChartSummingJunction = 123;
	public static final int msosptFlowChartOr = 124;
	public static final int msosptFlowChartCollate = 125;
	public static final int msosptFlowChartSort = 126;
	public static final int msosptFlowChartExtract = 127;
	public static final int msosptFlowChartMerge = 128;
	public static final int msosptFlowChartOfflineStorage = 129;
	public static final int msosptFlowChartOnlineStorage = 130;
	public static final int msosptFlowChartMagneticTape = 131;
	public static final int msosptFlowChartMagneticDisk = 132;
	public static final int msosptFlowChartMagneticDrum = 133;
	public static final int msosptFlowChartDisplay = 134;
	public static final int msosptFlowChartDelay = 135;
	public static final int msosptTextPlainText = 136;
	public static final int msosptTextStop = 137;
	public static final int msosptTextTriangle = 138;
	public static final int msosptTextTriangleInverted = 139;
	public static final int msosptTextChevron = 140;
	public static final int msosptTextChevronInverted = 141;
	public static final int msosptTextRingInside = 142;
	public static final int msosptTextRingOutside = 143;
	public static final int msosptTextArchUpCurve = 144;
	public static final int msosptTextArchDownCurve = 145;
	public static final int msosptTextCircleCurve = 146;
	public static final int msosptTextButtonCurve = 147;
	public static final int msosptTextArchUpPour = 148;
	public static final int msosptTextArchDownPour = 149;
	public static final int msosptTextCirclePour = 150;
	public static final int msosptTextButtonPour = 151;
	public static final int msosptTextCurveUp = 152;
	public static final int msosptTextCurveDown = 153;
	public static final int msosptTextCascadeUp = 154;
	public static final int msosptTextCascadeDown = 155;
	public static final int msosptTextWave1 = 156;
	public static final int msosptTextWave2 = 157;
	public static final int msosptTextWave3 = 158;
	public static final int msosptTextWave4 = 159;
	public static final int msosptTextInflate = 160;
	public static final int msosptTextDeflate = 161;
	public static final int msosptTextInflateBottom = 162;
	public static final int msosptTextDeflateBottom = 163;
	public static final int msosptTextInflateTop = 164;
	public static final int msosptTextDeflateTop = 165;
	public static final int msosptTextDeflateInflate = 166;
	public static final int msosptTextDeflateInflateDeflate = 167;
	public static final int msosptTextFadeRight = 168;
	public static final int msosptTextFadeLeft = 169;
	public static final int msosptTextFadeUp = 170;
	public static final int msosptTextFadeDown = 171;
	public static final int msosptTextSlantUp = 172;
	public static final int msosptTextSlantDown = 173;
	public static final int msosptTextCanUp = 174;
	public static final int msosptTextCanDown = 175;
	public static final int msosptFlowChartAlternateProcess = 176;
	public static final int msosptFlowChartOffpageConnector = 177;
	public static final int msosptCallout90 = 178;
	public static final int msosptAccentCallout90 = 179;
	public static final int msosptBorderCallout90 = 180;
	public static final int msosptAccentBorderCallout90 = 181;
	public static final int msosptLeftRightUpArrow = 182;
	public static final int msosptSun = 183;
	public static final int msosptMoon = 184;
	public static final int msosptBracketPair = 185;
	public static final int msosptBracePair = 186;
	public static final int msosptSeal4 = 187;
	public static final int msosptDoubleWave = 188;
	public static final int msosptActionButtonBlank = 189;
	public static final int msosptActionButtonHome = 190;
	public static final int msosptActionButtonHelp = 191;
	public static final int msosptActionButtonInformation = 192;
	public static final int msosptActionButtonForwardNext = 193;
	public static final int msosptActionButtonBackPrevious = 194;
	public static final int msosptActionButtonEnd = 195;
	public static final int msosptActionButtonBeginning = 196;
	public static final int msosptActionButtonReturn = 197;
	public static final int msosptActionButtonDocument = 198;
	public static final int msosptActionButtonSound = 199;
	public static final int msosptActionButtonMovie = 200;
	public static final int msosptHostControl = 201;    //Host controls extend various user interface (UI) objects in the Word and Excel object models
			public static final int msosptTextBox = 202;
	public static final int msosptMax = 0x0FFF;
	public static final int msosptNil = 0x0FFF;
}
