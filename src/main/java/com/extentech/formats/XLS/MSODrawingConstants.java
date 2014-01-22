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
	public static final int msosptNotPrimitive = msosptMin, msosptRectangle = 1, msosptRoundRectangle = 2, msosptEllipse = 3, msosptDiamond = 4, msosptIsocelesTriangle = 5, msosptRightTriangle = 6, msosptParallelogram = 7, msosptTrapezoid = 8, msosptHexagon = 9, msosptOctagon = 10, msosptPlus = 11, msosptStar = 12, msosptArrow = 13, msosptThickArrow = 14, msosptHomePlate = 15, msosptCube = 16, msosptBalloon = 17, msosptSeal = 18, msosptArc = 19, msosptLine = 20, msosptPlaque = 21, msosptCan = 22, msosptDonut = 23, msosptTextSimple = 24, msosptTextOctagon = 25, msosptTextHexagon = 26, msosptTextCurve = 27, msosptTextWave = 28, msosptTextRing = 29, msosptTextOnCurve = 30, msosptTextOnRing = 31, msosptStraightConnector1 = 32, msosptBentConnector2 = 33, msosptBentConnector3 = 34, msosptBentConnector4 = 35, msosptBentConnector5 = 36, msosptCurvedConnector2 = 37, msosptCurvedConnector3 = 38, msosptCurvedConnector4 = 39, msosptCurvedConnector5 = 40, msosptCallout1 = 41, msosptCallout2 = 42, msosptCallout3 = 43, msosptAccentCallout1 = 44, msosptAccentCallout2 = 45, msosptAccentCallout3 = 46, msosptBorderCallout1 = 47, msosptBorderCallout2 = 48, msosptBorderCallout3 = 49, msosptAccentBorderCallout1 = 50, msosptAccentBorderCallout2 = 51, msosptAccentBorderCallout3 = 52, msosptRibbon = 53, msosptRibbon2 = 54, msosptChevron = 55, msosptPentagon = 56, msosptNoSmoking = 57, msosptSeal8 = 58, msosptSeal16 = 59, msosptSeal32 = 60, msosptWedgeRectCallout = 61, msosptWedgeRRectCallout = 62, msosptWedgeEllipseCallout = 63, msosptWave = 64, msosptFoldedCorner = 65, msosptLeftArrow = 66, msosptDownArrow = 67, msosptUpArrow = 68, msosptLeftRightArrow = 69, msosptUpDownArrow = 70, msosptIrregularSeal1 = 71, msosptIrregularSeal2 = 72, msosptLightningBolt = 73, msosptHeart = 74, msosptPictureFrame = 75, msosptQuadArrow = 76, msosptLeftArrowCallout = 77, msosptRightArrowCallout = 78, msosptUpArrowCallout = 79, msosptDownArrowCallout = 80, msosptLeftRightArrowCallout = 81, msosptUpDownArrowCallout = 82, msosptQuadArrowCallout = 83, msosptBevel = 84, msosptLeftBracket = 85, msosptRightBracket = 86, msosptLeftBrace = 87, msosptRightBrace = 88, msosptLeftUpArrow = 89, msosptBentUpArrow = 90, msosptBentArrow = 91, msosptSeal24 = 92, msosptStripedRightArrow = 93, msosptNotchedRightArrow = 94, msosptBlockArc = 95, msosptSmileyFace = 96, msosptVerticalScroll = 97, msosptHorizontalScroll = 98, msosptCircularArrow = 99, msosptNotchedCircularArrow = 100, msosptUturnArrow = 101, msosptCurvedRightArrow = 102, msosptCurvedLeftArrow = 103, msosptCurvedUpArrow = 104, msosptCurvedDownArrow = 105, msosptCloudCallout = 106, msosptEllipseRibbon = 107, msosptEllipseRibbon2 = 108, msosptFlowChartProcess = 109, msosptFlowChartDecision = 110, msosptFlowChartInputOutput = 111, msosptFlowChartPredefinedProcess = 112, msosptFlowChartInternalStorage = 113, msosptFlowChartDocument = 114, msosptFlowChartMultidocument = 115, msosptFlowChartTerminator = 116, msosptFlowChartPreparation = 117, msosptFlowChartManualInput = 118, msosptFlowChartManualOperation = 119, msosptFlowChartConnector = 120, msosptFlowChartPunchedCard = 121, msosptFlowChartPunchedTape = 122, msosptFlowChartSummingJunction = 123, msosptFlowChartOr = 124, msosptFlowChartCollate = 125, msosptFlowChartSort = 126, msosptFlowChartExtract = 127, msosptFlowChartMerge = 128, msosptFlowChartOfflineStorage = 129, msosptFlowChartOnlineStorage = 130, msosptFlowChartMagneticTape = 131, msosptFlowChartMagneticDisk = 132, msosptFlowChartMagneticDrum = 133, msosptFlowChartDisplay = 134, msosptFlowChartDelay = 135, msosptTextPlainText = 136, msosptTextStop = 137, msosptTextTriangle = 138, msosptTextTriangleInverted = 139, msosptTextChevron = 140, msosptTextChevronInverted = 141, msosptTextRingInside = 142, msosptTextRingOutside = 143, msosptTextArchUpCurve = 144, msosptTextArchDownCurve = 145, msosptTextCircleCurve = 146, msosptTextButtonCurve = 147, msosptTextArchUpPour = 148, msosptTextArchDownPour = 149, msosptTextCirclePour = 150, msosptTextButtonPour = 151, msosptTextCurveUp = 152, msosptTextCurveDown = 153, msosptTextCascadeUp = 154, msosptTextCascadeDown = 155, msosptTextWave1 = 156, msosptTextWave2 = 157, msosptTextWave3 = 158, msosptTextWave4 = 159, msosptTextInflate = 160, msosptTextDeflate = 161, msosptTextInflateBottom = 162, msosptTextDeflateBottom = 163, msosptTextInflateTop = 164, msosptTextDeflateTop = 165, msosptTextDeflateInflate = 166, msosptTextDeflateInflateDeflate = 167, msosptTextFadeRight = 168, msosptTextFadeLeft = 169, msosptTextFadeUp = 170, msosptTextFadeDown = 171, msosptTextSlantUp = 172, msosptTextSlantDown = 173, msosptTextCanUp = 174, msosptTextCanDown = 175, msosptFlowChartAlternateProcess = 176, msosptFlowChartOffpageConnector = 177, msosptCallout90 = 178, msosptAccentCallout90 = 179, msosptBorderCallout90 = 180, msosptAccentBorderCallout90 = 181, msosptLeftRightUpArrow = 182, msosptSun = 183, msosptMoon = 184, msosptBracketPair = 185, msosptBracePair = 186, msosptSeal4 = 187, msosptDoubleWave = 188, msosptActionButtonBlank = 189, msosptActionButtonHome = 190, msosptActionButtonHelp = 191, msosptActionButtonInformation = 192, msosptActionButtonForwardNext = 193, msosptActionButtonBackPrevious = 194, msosptActionButtonEnd = 195, msosptActionButtonBeginning = 196, msosptActionButtonReturn = 197, msosptActionButtonDocument = 198, msosptActionButtonSound = 199, msosptActionButtonMovie = 200, msosptHostControl = 201,    //Host controls extend various user interface (UI) objects in the Word and Excel object models
			msosptTextBox = 202, msosptMax = 0x0FFF, msosptNil = 0x0FFF;
}
