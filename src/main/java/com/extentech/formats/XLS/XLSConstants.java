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
 *



 */
public interface XLSConstants
{
	// Stream types
	static final short WK_GLOBALS = 0x5;
	static final short VB_MODULE = 0x6;
	static final short WK_WORKSHEET = 0x10;
	static final short WK_CHART = 0x20;
	static final short WK_MACROSHEET = 0x40;
	static final short WK_FILE = 0x100;

	// Cell types
	public static final int TYPE_BLANK = -1;
	public static final int TYPE_STRING = 0;
	public static final int TYPE_FP = 1;
	public static final int TYPE_INT = 2;
	public static final int TYPE_FORMULA = 3;
	public static final int TYPE_BOOLEAN = 4;
	public static final int TYPE_DOUBLE = 5;

	// Book Options and constants

	// CalculationOptions
	public static int CALCULATE_ALWAYS = 0;
	public static int CALCULATE_EXPLICIT = 1;
	public static int CALCULATE_AUTO = 2; // replacement for calc always
	public static String CALC_MODE_PROP = "com.extentech.extenxls.calcmode";
	public static String REFTRACK_PROP = "com.extentech.extenxls.trackreferences";
	public static String USETEMPFILE_PROP = "com.extentech.formats.LEO.usetempfile";
	public static String VALIDATEWORKBOOK = "com.extentech.formats.LEO.validateworkbook";

	// String table handling
	public static int STRING_ENCODING_AUTO = 0;
	public static int STRING_ENCODING_UNICODE = 1;
	public static int STRING_ENCODING_COMPRESSED = 2;
	public static int ALLOWDUPES = 0;
	public static int SHAREDUPES = 1;
	public static String DEFAULTENCODING = "ISO-8859-1"; // "UTF-8";"UTF-8";
	public static String UNICODEENCODING = "UTF-16LE"; // "UnicodeLittleUnmarked";

	// XLSRecord Opcodes
	static final short EXCEL2K = 0x1C0;
	static final short GARBAGE = 0xFFFFFFFE;
	static final short TXO = 0x1B6;
	static final short MSODRAWINGGROUP = 0xEB;
	static final short MSODRAWING = 0xEC;
	static final short MSODRAWINGSELECTION = 0xED;
	static final short PHONETIC = 0xEF;
	static final short CONTINUE = 0x3C;
	static final short COLINFO = 0x7D;
	static final short SST = 0xFC;
	static final short DSF = 0x161;
	static final short EXTSST = 0xFF;
	static final short ENDEXTSST = 0xFE;
	static final short BOF = 0x809;
	static final short FILEPASS = 0x2F;
	static final short INDEX = 0x20B;
	static final short DBCELL = 0xD7;
	static final short BOUNDSHEET = 0x85;
	static final short COUNTRY = 0x8C;    // record just after bound sheet
	static final short BOOKBOOL = 0xDA;
	static final short CALCCOUNT = 0x0C;
	static final short CALCMODE = 0x0D;
	static final short PRECISION = 0x0E;
	static final short REFMODE = 0x0F;
	static final short DELTA = 0x10;
	static final short ITERATION = 0x11;
	static final short DATE1904 = 0x22;
	static final short BACKUP = 0x40;
	static final short PRINT_ROW_HEADERS = 0x2A;
	static final short PRINT_GRIDLINES = 0x2B;
	static final short HORIZONTAL_PAGE_BREAKS = 0x1B;
	static final short HLINK = 0x1B8;
	static final short VERTICAL_PAGE_BREAKS = 0x1A;
	static final short DEFAULTROWHEIGHT = 0x225;
	static final short FONT = 0x31;
	static final short HEADERREC = 0x14;
	static final short FOOTERREC = 0x15;
	static final short LEFT_MARGIN = 0x26;
	static final short RIGHT_MARGIN = 0x27;
	static final short TOP_MARGIN = 0x28;
	static final short BOTTOM_MARGIN = 0x29;
	static final short DCON = 0x50;
	static final short DEFCOLWIDTH = 0x55;
	static final short EXTERNCOUNT = 0x16;
	static final short EXTERNSHEET = 0x17;
	static final short EXTERNNAME = 0x23;
	static final short FORMAT = 0x41E;
	static final short XF = 0xE0;
	static final short NAME = 0x18;
	static final short DIMENSIONS = 0x200;
	static final short FILE_LOCK = 0x195;
	static final short RRD_INFO = 0x196;
	static final short RRD_HEAD = 0x138;
	static final short EOF = 0x0A;
	static final short BLANK = 0x201;
	static final short MERGEDCELLS = 0xE5;
	static final short MULBLANK = 0xBE;
	static final short MULRK = 0xBD;
	static final short NOTE = 0x1C;
	static final short NUMBER = 0x203;
	static final short LABEL = 0x204;
	static final short LABELSST = 0xFD;
	static final short BOOLERR = 0x205;
	static final short FORMULA = 0x06; // 0x406;
	static final short ARRAY = 0x221;//0x21; //
	static final short SELECTION = 0x1D;
	static final short STYLE = 0x293;
	static final short ROW = 0x208;
	static final short RK = 0x27E; // this is wrong according to the documentation (0x27) ... -jm
	static final short RSTRING = 0xD6;
	static final short SHRFMLA = 0x4BC; // according to docs this is 0xBC
	static final short STRINGREC = 0x207;
	static final short TABLE = 0x236;
	static final short PANE = 0x41;
	static final short PASSWORD = 0x13;
	static final short INTERFACE_HDR = 0xE1;
	static final short USR_EXCL = 0x194;
	static final short PALETTE = 0x92;
	static final short PROTECT = 0x12;
	static final short OBJPROTECT = 0x63;
	static final short SCENPROTECT = 0xDD;
	static final short FEATHEADR = 0x867;    // extra protection settings + smarttag settings
	static final short SCL = 0xA0; // zoom
	static final short SHEETPROTECTION = 0x867;    //
	static final short SHEETLAYOUT = 0x862;
	static final short RANGEPROTECTION = 0x868;
	static final short PROT4REV = 0x1AF;
	static final short WINDOW_PROTECT = 0x19;
	static final short WINDOW1 = 0x3D;
	static final short WINDOW2 = 0x23E;
	static final short PLV = 0x88B;
	static final short RTENTEXU = 0x1B;
	static final short DV = 0x1BE;
	static final short DVAL = 0x1B2;
	static final short RTMERGECELLS = 0xE5;
	static final short SUPBOOK = 0x1AE;
	static final short USERSVIEWBEGIN = 0x1AA;
	static final short USERSVIEWEND = 0x1AB;
	static final short USERBVIEW = 0x1A9;
	static final short PLS = 0x4D;
	static final short WSBOOL = 0x81;
	static final short OBJ = 0x5D;
	static final short OBPROJ = 0xD3;
	static final short XLS_MAX_COLS = 0x100;
	static final short TABID = 0x13d;
	static final short GUTS = 0x80;
	static final short CODENAME = 0x1BA;
	static final short XCT = 0x59;    // 20080122 KSC:
	static final short CRN = 0x5A;    // ""

	// Pivot Table records
	static final short SXVIEW = 0xB0;
	static final short TABLESTYLES = 0x88E;
	static final short SXSTREAMID = 0xD5;
	static final short SXVS = 0xE3;
	static final short SXADDL = 0x864;
	static final short SXVDEX = 0x100;
	static final short SXPI = 0xB6;
	static final short SXDI = 0xC5;
	static final short SXDB = 0xC6;
	static final short SXFDB = 0xC7;
	static final short SXEX = 0xF1;
	static final short QSISXTAG = 0x802;
	static final short SXVIEWEX9 = 0x810;
	static final short DCONREF = 0x51;
	static final short DCONNAME = 0x52;
	static final short DCONBIN = 0x1B5;
	static final short SXFORMAT = 0xFB;
	static final short SXLI = 0xB5;
	static final short SXVI = 0xB2;
	static final short SXVD = 0xB1;
	static final short SXIVD = 0xB4;
	static final short SXDBEX = 0x122;
	static final short SXFDBTYPE = 0x1BB;
	static final short SXDBB = 0xC8;
	static final short SXNUM = 0xC9;
	static final short SXBOOL = 0xCA;
	static final short SXSTRING = 0xCD;

	// Printing records
	static final short SETUP = 0xA1;
	static final short HCENTER = 0x83;
	static final short VCENTER = 0x84;
	static final short LEFTMARGIN = 0x26;
	static final short RIGHTMARGIN = 0x27;
	static final short TOPMARGIN = 0x28;
	static final short BOTTOMMARGIN = 0x29;
	static final short PRINTGRID = 0x2B;
	static final short PRINTROWCOL = 0x2A;

	// Conditional Formatting
	static final short CF = 0x1B1;
	static final short CONDFMT = 0x1B0;
	// 2007 Conditional Formatting
	static final short CF12 = 0x87A;
	static final short CONDFMT12 = 0x879;

	// AutoFilter
	static final short AUTOFILTER = 0x9E;

	// Chart items
	static final short UNITS = 0x1001;
	static final short CHART = 0x1002;
	static final short SERIES = 0x1003;
	static final short DATAFORMAT = 0x1006;
	static final short LINEFORMAT = 0x1007;
	static final short MARKERFORMAT = 0x1009;
	static final short AREAFORMAT = 0x100A;
	static final short PIEFORMAT = 0x100B;
	static final short ATTACHEDLABEL = 0x100C;
	static final short SERIESTEXT = 0x100D;
	static final short CHARTFORMAT = 0x1014;
	static final short LEGEND = 0x1015;
	static final short SERIESLIST = 0x1016;
	static final short BAR = 0x1017;
	static final short LINE = 0x1018;
	static final short PIE = 0x1019;
	static final short AREA = 0x101A;
	static final short SCATTER = 0x101B;
	static final short CHARTLINE = 0x101C;
	static final short AXIS = 0x101D;
	static final short TICK = 0x101E;
	static final short VALUERANGE = 0x101F;
	static final short CATSERRANGE = 0x1020;
	static final short AXISLINEFORMAT = 0x1021;
	static final short CHARTFORMATLINK = 0x1022;
	static final short DEFAULTTEXT = 0x1024;
	static final short TEXTDISP = 0x1025;
	static final short FONTX = 0x1026;
	static final short OBJECTLINK = 0x1027;
	static final short FRAME = 0x1032;
	static final short BEGIN = 0x1033;
	static final short END = 0x1034;
	static final short PLOTAREA = 0x1035;
	static final short THREED = 0x103A;
	static final short PICF = 0x103C;
	static final short DROPBAR = 0x103D;
	static final short RADAR = 0x103E;
	static final short SURFACE = 0x103F;
	static final short RADARAREA = 0x1040;
	static final short AXISPARENT = 0x1041;
	static final short LEGENDXN = 0x1043;
	static final short SHTPROPS = 0x1044;
	static final short SERTOCRT = 0x1045;
	static final short AXESUSED = 0x1046;
	static final short SBASEREF = 0x1048;
	static final short SERPARENT = 0x104A;
	static final short SERAUXTREND = 0x104B;
	static final short IFMT = 0x104E;
	static final short POS = 0x104F;
	static final short ALRUNS = 0x1450;
	static final short AI = 0x1051;
	static final short SERAUXERRBAR = 0x105B;
	static final short SERFMT = 0x105D;
	static final short CHART3DBARSHAPE = 0x105F;
	static final short FBI = 0x1460;
	static final short BOPPOP = 0x1061;
	static final short AXCENT = 0x1062;
	static final short DAT = 0x1063;
	static final short PLOTGROWTH = 0x1064;
	static final short SIIINDEX = 0x1065;
	static final short GELFRAME = 0x1066;
	static final short BOPPOPCUSTOM = 0x1067;
	//20080703 KSC: Excel 9 Chart Records
	static final short CRTLAYOUT12 = 0x89D;
	static final short CRTLAYOUT12A = 0x08A7;
	static final short CHARTFRTINFO = 0x0850;
	static final short FRTWRAPPER = 0x0851;
	static final short STARTBLOCK = 0x0852;
	static final short ENDBLOCK = 0x0853;
	static final short STARTOBJECT = 0x0854;
	static final short ENDOBJECT = 0x0855;
	static final short CATLAB = 0x0856;
	static final short YMULT = 0x0857;
	static final short SXVIEWLINK = 0x0858;
	static final short PIVOTCHARTBITS = 0x0859;
	static final short FRTFONTLIST = 0x085A;
	static final short PIVOTCHARTLINK = 0x0861;
	static final short DATALABEXT = 0x086A;
	static final short DATALABEXTCONTENTS = 0x086B;
	static final short FONTBASIS = 0x1060;

	// max size of records -- depends on XLS version
	static final int MAXRECLEN = 8224;
	static final int MAXROWS_BIFF8 = 65536;
	static final int MAXCOLS_BIFF8 = 256;
	static final int MAXROWS = 1048576; // new Excel 2007 limits
	static final int MAXCOLS = 16384;    // new Excel 2007 limits

}