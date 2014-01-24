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

import java.awt.Color;

/**
 * /** Constants relevant to XLS Formatting
 * <p/>
 * NOTE:
 * Excel automatically assigns
 * If you type     this number format
 * -------------------------------------------
 * <p/>
 * 1.0             General
 * 1.123           General
 * 1.1%            0.00%
 * 1.1E+2          0.00E+00
 * 1 1/2           # ?/?
 * $1.11           Currency, 2 decimal places
 * 1/1/01          Date
 * 1:10            Time
 * <p/>
 * CUSTOM FORMATS:
 * <p/>
 * Each format that you create can have up to three sections for numbers and a fourth section for text.
 * <POSITIVE>;<NEGATIVE>;<ZERO>;<TEXT>
 *
 * @see CustomFormatHelper
 */
public interface FormatConstants
{

	public static final int DEFAULT_FONT_WEIGHT = 200;
	public static final int DEFAULT_FONT_SIZE = 20;
	public static final String DEFAULT_FONT_FACE = "Arial";

	// Font Weights
	public static final int PLAIN = 400;
	public static final int BOLD = 700;

	//	 Font UnderLine Style constants
	public static final byte STYLE_UNDERLINE_NONE = 0x0;
	public static final byte STYLE_UNDERLINE_SINGLE = 0x1;
	public static final byte STYLE_UNDERLINE_DOUBLE = 0x2;
	public static final byte STYLE_UNDERLINE_SINGLE_ACCTG = 0x21;
	public static final byte STYLE_UNDERLINE_DOUBLE_ACCTG = 0x22;

	// Color Constants
	public static Color Black = new Color( 0, 0, 0 );
	public static Color Brown = new Color( 153, 51, 0 );
	public static Color OliveGreen = new Color( 51, 51, 0 );
	public static Color Dark_Green = new Color( 0, 51, 0 );
	public static Color Dark_Teal = new Color( 0, 51, 102 );
	public static Color Dark_Blue = new Color( 0, 0, 128 );
	public static Color Indigo = new Color( 51, 51, 153 );
	public static Color Gray80 = new Color( 51, 51, 51 );
	// could we be wrong??  test it per claritas
// public static Color Gray50			= new Color(110,110,110);
	// Excel 07 likes real 50 percent 			= new Color(127,127,127);
	public static Color Gray50 = new Color( 136, 136, 136 );

	public static Color Dark_Red = new Color( 128, 0, 0 );
	public static Color Orange = new Color( 255, 102, 0 );
	public static Color Dark_Yellow = new Color( 128, 128, 0 );
	public static Color Green = new Color( 0, 128, 0 );
	public static Color Teal = new Color( 0, 128, 128 );
	public static Color Blue = new Color( 0, 0, 255 );
	public static Color BlueGray = new Color( 102, 102, 153 );

	public static Color Red = new Color( 255, 0, 0 );
	public static Color Light_Orange = new Color( 255, 153, 0 );
	public static Color Lime = new Color( 153, 204, 0 );
	public static Color SeaGreen = new Color( 51, 153, 102 );
	public static Color Aqua = new Color( 51, 204, 204 );
	public static Color Light_Blue = new Color( 0, 128, 255 );
	public static Color Violet = new Color( 128, 0, 128 );
	public static Color Gray40 = new Color( 150, 150, 150 );

	public static Color Pink = new Color( 255, 0, 255 );
	public static Color Gold = new Color( 255, 204, 0 );
	public static Color Yellow = new Color( 255, 255, 0 );
	public static Color BrightGreen = new Color( 0, 255, 0 );
	public static Color Turquoise = new Color( 0, 255, 255 );
	public static Color SkyBlue = new Color( 0, 204, 255 );
	public static Color Plum = new Color( 153, 51, 102 );
	public static Color Gray25 = new Color( 192, 192, 192 );

	public static Color Rose = new Color( 255, 153, 204 );
	public static Color Tan = new Color( 255, 204, 153 );
	public static Color Light_Yellow = new Color( 255, 255, 153 );
	public static Color Light_Green = new Color( 204, 255, 204 );
	public static Color Light_Turquoise = new Color( 204, 255, 255 );
	public static Color PaleBlue = new Color( 153, 204, 255 );
	public static Color Lavender = new Color( 204, 153, 255 );
	public static Color White = new Color( 255, 255, 255 );
	public static Color Gray15 = new Color( 216, 216, 216 );
	public static Color Dark_Purple = new Color( 102, 0, 102 );
	public static Color Salmon = new Color( 255, 128, 128 );
	public static Color Light_Purple = new Color( 204, 204, 255 );
	public static Color Medium_Purple = new Color( 153, 153, 255 );

	/**
	 * Excel's basic icv color table.  This may be altered by the Palette record.
	 * See WorkBook.CCLORTABLE
	 */
	static final Color[] COLORTABLE = {
			FormatConstants.Black,
			FormatConstants.White,
			FormatConstants.Red,
			FormatConstants.BrightGreen,
			/// WRONG, was duplicate, incorrect GREEN
			FormatConstants.Blue,
			FormatConstants.Yellow,
			FormatConstants.Pink,
			FormatConstants.Turquoise,
			FormatConstants.Black,
/**
 * NOTE: from documentation:
 * if palette record exists 8-63 (0x3f) gets read from it,
 * otherwise If a Palette record exists in this file, these icv values 
 * specify colors from the rgColor array in the Palette record. 			
 */
			/*
			 * Color 9 === AUTOMATIC color
			 * the opposite of what is normal; for instance, background color
			 * index 9= black, font color index 9= white. It is set here to
			 * black, and in Fonts.getColor, a special case is put for icv= 9
			 */
			FormatConstants.White,
			// 9
			FormatConstants.Red,
			// 10
			FormatConstants.BrightGreen,
			FormatConstants.Blue,
			FormatConstants.Yellow,
			FormatConstants.Pink,
			FormatConstants.Turquoise,
			// 15
			FormatConstants.Dark_Red,
// Burgundy //16
			FormatConstants.Green,
			FormatConstants.Dark_Blue,
// purple
			FormatConstants.Dark_Yellow,
			FormatConstants.Violet,
// 20 dark pink
			FormatConstants.Teal,
			FormatConstants.Gray25,
			FormatConstants.Gray50,
			FormatConstants.Medium_Purple,
			FormatConstants.Plum,
// 25
			FormatConstants.Light_Yellow,
			FormatConstants.Light_Turquoise,
			FormatConstants.Dark_Purple,
			FormatConstants.Salmon,
			FormatConstants.BlueGray,
// 30
			FormatConstants.Light_Purple,
			FormatConstants.Dark_Blue,
			FormatConstants.Pink,
			FormatConstants.Yellow,
			FormatConstants.Turquoise,
// 35
			FormatConstants.Violet,
			FormatConstants.Dark_Red,
			FormatConstants.Teal,
// ////////////WRONG - WAS DUPLICATE
			FormatConstants.Blue,
			FormatConstants.SkyBlue,
// 40
			FormatConstants.Light_Turquoise,
			FormatConstants.Light_Green,
			FormatConstants.Light_Yellow,
			FormatConstants.PaleBlue,
			FormatConstants.Rose,
// 45
			FormatConstants.Lavender,
			FormatConstants.Tan,
			FormatConstants.Dark_Yellow,
			FormatConstants.Aqua,
			FormatConstants.Lime,
// 50
			FormatConstants.Gold,
			FormatConstants.Light_Orange,
			FormatConstants.Orange,
			FormatConstants.BlueGray,
			FormatConstants.Gray40,
// 55
			FormatConstants.Dark_Teal,
			FormatConstants.SeaGreen,
			FormatConstants.Dark_Green,
			FormatConstants.OliveGreen,
			FormatConstants.Brown,
// 60
			FormatConstants.Plum,
			FormatConstants.Indigo,
			FormatConstants.Gray80,
			FormatConstants.White,
// 64
			FormatConstants.Black,
// 65 WHY DO THESE CHANGE???

	};

	// Formatting Color Constants
	public static int COLOR_BLACK = 0;
	public static int COLOR_WHITE = 1;
	public static int COLOR_RED = 2;
	public static int COLOR_BLUE = 4;
	public static int COLOR_YELLOW = 5;
	public static int COLOR_BLACK3 = 8;
	public static int COLOR_WHITE3 = 9;
	public static int COLOR_RED_CHART = 10;    // chart fills use SPECIFIC icv numbers -- incorrect # results in black fill
	public static int COLOR_BRIGHT_GREEN = 11;
	public static int COLOR_YELLOW_CHART = 13;
	public static int COLOR_DARK_RED = 16;
	public static int COLOR_GREEN = 17;
	public static int COLOR_DARK_BLUE = 18;
	public static int COLOR_OLIVE_GREEN_CHART = 19;
	public static int COLOR_DARK_YELLOW = 19;
	public static int COLOR_VIOLET = 20;
	public static int COLOR_TEAL = 21;
	public static int COLOR_GRAY25 = 22;
	public static int COLOR_GRAY50 = 23;
	public static int COLOR_MEDIUM_PURPLE = 24;
	public static int COLOR_PLUM = 25;
	public static int COLOR_LIGHT_YELLOW = 26;
	public static int COLOR_LIGHT_TURQUOISE = 27;
	public static int COLOR_DARK_PURPLE = 28;
	public static int COLOR_SALMON = 29;
	public static int COLOR_LIGHT_BLUE = 30;
	public static int COLOR_LIGHT_PURPLE = 31;
	public static int COLOR_PINK = 33;
	public static int COLOR_CYAN = 35;
	public static int COLOR_BLUE_CHART = 39;
	public static int COLOR_SKY_BLUE = 40;

	public static int COLOR_LIGHT_GREEN = 42;
	public static int COLOR_PALE_BLUE = 44;
	public static int COLOR_ROSE = 45;
	public static int COLOR_LAVENDER = 46;
	public static int COLOR_TAN = 47;
	public static int COLOR_TURQUOISE = 48;
	public static int COLOR_AQUA = 49;
	public static int COLOR_LIME = 50;
	public static int COLOR_GOLD = 51;
	public static int COLOR_LIGHT_ORANGE = 52;
	public static int COLOR_ORANGE = 53;
	public static int COLOR_BLUE_GRAY = 54;
	public static int COLOR_GRAY40 = 55;
	public static int COLOR_DARK_TEAL = 56;
	public static int COLOR_SEA_GREEN = 57;
	public static int COLOR_DARK_GREEN = 58;
	public static int COLOR_OLIVE_GREEN = 59;
	public static int COLOR_BROWN = 60;
	public static int COLOR_INDIGO = 62;
	public static int COLOR_GRAY80 = 63;
	public static int COLOR_WHITE2 = 64;
	public static int COLOR_BLACK2 = 65;

	/**
	 * given color int ala COLORTABLE, interpret Excel Color name
	 */
	public static String[] COLORNAMES = {
			"Black", "White", "Red", "BrightGreen", "Blue", "Yellow", "Pink", "Turquoise", "Black", "White", "Red",        //10
			"BrightGreen", "Blue", "Yellow", "Pink", "Turquoise",    // 15
			"Dark_Red",//Burgundy //16
			"Green", "Dark_Blue", //purple
			"Dark_Yellow", "Violet", //20 dark pink
			"Teal", "Gray25", "Gray50", "MediumPurple", "Plum", //25
			"Light_Yellow", "Light_Turquoise", "Dark_Purple", "Salmon", "BlueGray", //30
			"Light_Purple", "Dark_Blue", "Pink", "Yellow", "Turquoise",    // 35
			"Violet", "Dark_Red", "Teal", "Blue", "SkyBlue", //40
			"Light_Turquoise", "Light_Green", "Light_Yellow", "PaleBlue", "Rose", //45
			"Lavender", "Tan", "Blue", "Aqua", "Lime", // 50
			"Gold", "Light_Orange", "Orange", "BlueGray", "Gray40", //55
			"Dark_Teal", "SeaGreen", "Dark_Green", "OliveGreen", "Brown", //60
			"Plum", "Indigo", // WRONG
			"Gray80", "White", // 64
			"Black", // 65
	};

	/**
	 * given color int ala COLORTABLE, interpret HTML Color name
	 */
	public static String[] HTMLCOLORNAMES = {
			"Black", "White", "Red", "Green", "Blue", "Yellow", "Pink", "Turquoise", "Black", "White", "Red",        //10
			"Lime",            //	BrightGreen",
			"Blue", "Yellow", "Pink", "Turquoise",    // 15
			"Dark_Red", "Green", "DarkBlue", "DarkKhaki",        //Dark_Yellow",
			"Violet",    //20
			"Teal", "LightGray",        // Gray25
			"DarkGray",            // Gray50"
			"DodgerBlue", "DarkRed", //25
			"LightYellow", "PaleTurquoise",    // Light_Turquoise",
			"DarkMagenta", "Red", "CadetBlue",        // BlueGray", //30
			"PowderBlue",        // PaleBlue",
			"DarkBlue", "Pink", "Yellow", "Turquoise",    // 35
			"Plum", "Brown", "SeaGreen", "Blue", "SkyBlue", //40
			"PaleTurquoise",    //Light_Turquoise",
			"LightGreen", "LightYellow", "PowderBlue",        // PaleBlue",
			"Crimson",            // Rose", //45
			"Lavender", "Tan", "Blue", "Aqua", "Lime", // 50
			"Gold", "Orange",        //Light_Orange",
			"Orange", "CadetBlue",        // BlueGray",
			"YellowGreen",            // Gray40", //55
			"DarkSlateBlue",        //Dark_Teal",
			"SeaGreen", "DarkGreen", "OliveDrab",        //OliveGreen",
			"Brown", //60
			"Plum", "Indigo", // WRONG
			"DimGray",            // Gray80
			"White", // 64
			"Black", // 65
	};

	public static String[] SVGCOLORSTRINGS = {
			"black", "white", "red", "lime",    // = BrightGreen
			"blue", "yellow",            // 5
			"fuchsia",    // = Pink
			"aqua",        // = Turquoise
			"black", "white", "red",       // 10
			"lime", "blue", "yellow", "fuchsia", "aqua",        // 15
			"maroon",    // = Dark_Red
			"green", "navy",        // = Dark_Blue
			"olive",    // = Dark_Yellow
			"purple",    // = Violet // 20
			"teal", "silver",        // = Gray25
			"gray",        // = Gray50
			"lightblue",    //lightsteelblue",
			"palevioletred",    //mediumorchid",	// =Plum 	// 25
			"darkolivegreen", "aqua",    // = Light_Turqouise
			"purple",    // = Dark_Purple
			"salmon", "slateblue",    // = BlueGray		// 30
			"aliceblue", // = Light_Purple
			"navy",        // = Dark_Blue
			"fuchsia",        // = Pink
			"yellow", "aqua",            // = Turquoise //35
			"purple",        // = Violet
			"maroon",        // = Dark_Red
			"teal", "blue", "deepskyblue",    // = SkyBlue				// 40
			"azure",        // = Light_Turquoise
			"honeydew",    // = Light_Green
			"lightyellow", "powderblue",    // = PaleBlue
			"lightpink",    // = Rose			// 45
			"aliceblue",    // = Lavendar
			"navajowhite",    // = Tan
			"blue", "mediumturquoise",    // = Aqua
			"lawngreen",        // = Lime	// 50
			"gold", "orange", "darkorange", "slateblue",    // = BlueGray
			"darkgray",        // = Gray40		//55
			"midnightblue",    // = Dark_Teal
			"seagreen", "darkgreen", "darkolivegreen", "saddlebrown",                // 60
			"mediumvioletred", "darkslateblue",    //= Indigo
			"darkslategray",        // = Gray80
			"white", // 64
			"black", // 65

	};

	// Formatting Pattern Constants
	public static int PATTERN_NONE = 0;
	public static int PATTERN_FILLED = 1;
	public static int PATTERN_LIGHT_FILL = 2;
	public static int PATTERN_MED_FILL = 3;
	public static int PATTERN_HVY_FILL = 4;
	public static int PATTERN_HOR_STRIPES1 = 5;
	public static int PATTERN_VERT_STRIPES1 = 6;
	public static int PATTERN_DIAG_STRIPES1 = 7;
	public static int PATTERN_DIAG_STRIPES2 = 8;
	public static int PATTERN_CHECKERBOARD1 = 9;
	public static int PATTERN_CHECKERBOARD2 = 10;
	public static int PATTERN_HOR_STRIPES2 = 11;
	public static int PATTERN_VERT_STRIPES2 = 12;
	public static int PATTERN_DIAG_STRIPES3 = 13;
	public static int PATTERN_DIAG_STRIPES4 = 14;
	public static int PATTERN_GRID1 = 15;
	public static int PATTERN_CROSSPATCH1 = 16;
	public static int PATTERN_MED_DOTS = 17;
	public static int PATTERN_LIGHT_DOTS = 18;
	public static int PATTERN_HVY_DOTS = 19;
	public static int PATTERN_DIAG_STRIPES5 = 20;
	public static int PATTERN_DIAG_STRIPES6 = 21;
	public static int PATTERN_DIAG_STRIPES7 = 22;
	public static int PATTERN_DIAG_STRIPES8 = 23;
	public static int PATTERN_VERT_STRIPES3 = 24;
	public static int PATTERN_HOR_STRIPES3 = 25;
	public static int PATTERN_VERT_STRIPES4 = 26;
	public static int PATTERN_HOR_STRIPES4 = 27;
	public static int PATTERN_PATCHY1 = 28;
	public static int PATTERN_PATCHY2 = 29;
	public static int PATTERN_PATCHY3 = 30;
	public static int PATTERN_PATCHY4 = 31;

	// Border line styles

	public static short BORDER_NONE = 0x0;
	public static short BORDER_THIN = 0x1;
	public static short BORDER_MEDIUM = 0x2;
	public static short BORDER_DASHED = 0x3;
	public static short BORDER_DOTTED = 0x4;
	public static short BORDER_THICK = 0x5;
	public static short BORDER_DOUBLE = 0x6;
	public static short BORDER_HAIR = 0x7;
	public static short BORDER_MEDIUM_DASHED = 0x8;
	public static short BORDER_DASH_DOT = 0x9;
	public static short BORDER_MEDIUM_DASH_DOT = 0xA;
	public static short BORDER_DASH_DOT_DOT = 0xB;
	public static short BORDER_MEDIUM_DASH_DOT_DOT = 0xC;
	public static short BORDER_SLANTED_DASH_DOT = 0xD;

	public static String[] BORDER_NAMES = {
			"None",
			"Thin",
			"Medium",
			"Dashed",
			"Dotted",
			"Thick",
			"Double",
			"Hair",
			"Medium dashed",
			"Dash-dot",
			"Medium dash-dot",
			"Dash-dot-dot",
			"Medium dash-dot-dot",
			"Slanted dash-dot"
	};

	// interpret border sizes from specifications above
	public static String[] BORDER_SIZES_HTML = {
			"", "2px",    // thin
			"3px",    // medium
			"2px", "2px", "4px",    // thick
			"2ox", "1px",    // hair
			"3px",    // medium dashed
			"2px", "2px", "2px", "2px", "2px"
	};

	public static String[] BORDER_NAMES_HTML = {
			"", "solid",    // thin
			"solid",    // medium
			"dashed", "dotted", "solid",    // thick
			"double", "solid",    // hair
			"dashed", "dashed", "dashed", "dotted", "dotted", "dashed"
	};

	public static final String[] BORDER_STYLES_JSON = {
			"none",                    // 0x0 No border
			"thin",                    // 0x1 Thin line
			"medium",                // 0x2 Medium line
			"dashed",                // 0x3 Dashed line
			"dotted",                // 0x4 Dotted line
			"thick",                // 0x5 Thick line
			"double",                // 0x6 Double line
			"hairline",                // 0x7 Hairline
			"medium_dashed",        // 0x8 Medium dashed line
			"dash-dot",                // 0x9 Dash-dot line
			"medium_dash-dot",        // 0xA Medium dash-dot line
			"dash-dot-dot",            // 0xB Dash-dot-dot line
			"medium_dash-dot-dot",    // 0xC Medium dash-dot-dot line
			"slant_dash-dot-dot"    // 0xD Slanted dash-dot-dot line
	};

	public static int FORMAT_SUBSCRIPT = 2;
	public static int FORMAT_SUPERSCRIPT = 1;
	public static int FORMAT_NOSCRIPT = 0;

	// Decoded Built-in Format Patterns + HEX Format IDs
	// includes number, currency and date formats
	/**
	 * Please use the getBuiltinFormats in FormatConstantsImpl for locale specific formatting!
	 */
	public static String[][] BUILTIN_FORMATS = {
			{ "General", "0" },
			{ "0", "1" },
			{ "0.00", "2" },
			{ "#,##0", "3" },
			{ "#,##0.00", "4" },
			{ "$#,##0;($#,##0)", "5" },
			{ "$#,##0;[Red]($#,##0)", "6" },
			{ "$#,##0.00;($#,##0.00)", "7" },
			{ "$#,##0.00;[Red]($#,##0.00)", "8" },
			{ "0%", "9" },
			{ "0.00%", "a" },
			{ "0.00E+00", "b" },
			{ "# ?/?", "c" },
			{ "# ??/??", "d" },

			// dates
			{ "mm-dd-yy", "E" },
			// REGIONAL FORMAT!	m/d/yyyy  or m/d/yy
			{ "d-mmm-yy", "F" },
			{ "d-mmm", "10" },
			{ "mmm-yy", "11" },
			{ "h:mm AM/PM", "12" },
			{ "h:mm:ss AM/PM", "13" },
			{ "h:mm", "14" },
			{ "h:mm:ss", "15" },
			{ "m/d/yy h:mm", "16" },

			// 17h through 24h are undocumented international formats!!!
			{ "reserved", "17" },
			{ "reserved", "18" },
			{ "reserved", "19" },
			{ "reserved", "1a" },
			{ "reserved", "1b" },
			{ "reserved", "1c" },
			{ "reserved", "1d" },
			{ "reserved", "1e" },
			{ "reserved", "1f" },
			{ "reserved", "20" },
			{ "reserved", "21" },
			{ "reserved", "22" },
			{ "reserved", "23" },
			{ "reserved", "24" },

			// more currency
			{ "(#,##0_);(#,##0)", "25" },
			{ "(#,##0_);[Red](#,##0)", "26" },
			{ "(#,##0.00_);(#,##0.00)", "27" },
			{ "(#,##0.00_);[Red](#,##0.00)", "28" },
			{ "_(*#,##0_);_(*(#,##0);_(*\"-\"_);_(@_)", "29" },
			{ "_($* #,##0_);_($* (#,##0);_($*\"-\"_);_(@_)", "2a" },
			{ "_(* #,##0.00_);_(* (#,##0.00);_(*\"-\"??_);_(@_)", "2b" },
			{ "_($* #,##0.00;_($* (#,##0.00);_($* \"-\"??;_(@_)", "2c" },

			// more dates
			{ "mm:ss", "2D" },
			{ "[h]:mm:ss", "2E" },
			{ "mm:ss.0", "2F" },

			// misc
			{ "##0.0E+0", "30" },
			{ "@", "31" },
	};

	public static String[][] BUILTIN_FORMATS_JP = {
			{ "General", "0" },
			{ "0", "1" },
			{ "0.00", "2" },
			{ "#,##0", "3" },
			{ "#,##0.00", "4" },
			{ "\u00a5#,##0;\u00a5-#,##0", "5" },
			{ "\u00a5#,##0;[Red]\u00a5-#,##0", "6" },
			{ "\u00a5#,##0.00;\u00a5-#,##0.00", "7" },
			{ "\u00a5#,##0.00;[Red]\u00a5-#,##0.00", "8" },
			{ "0%", "9" },
			{ "0.00%", "a" },
			{ "0.00E+00", "b" },
			{ "# ?/?", "c" },
			{ "# ??/??", "d" },

			// dates
			{ "yyyy/m/d", "e" },
			{ "d-mmm-yy", "f" },
			{ "d-mmm", "10" },
			{ "mmm-yy", "11" },
			{ "h:mm AM/PM", "12" },
			{ "h:mm:ss AM/PM", "13" },
			{ "h:mm", "14" },
			{ "h:mm:ss", "15" },
			{ "yyyy/m/d h:mm", "16" },

			// 17h through 24h are undocumented international formats!!!
			{ "($#,##0_);($#,##0)", "17" },
			{ "($#,##0_);[Red]($#,##0)", "18" },
			{ "($#,##0_);[Red]($#,##0)", "19" },
			{ "($#,##0.00_);[Red]($#,##0.00)", "1a" },
			{ "[$-411]ge.m.d", "1b" },
			{ "[$-411]ggge\u5E74m\u6708d\u65E5", "1c" },
			{ "reserved", "1d" },
			{ "m/d/yy", "1e" },
			{ "yyyy\u5E74m\u6708d\u65E5", "1f" },
			{ "h\u6642mm\u5206", "20" },
			{ "h\u6642mm\u5206ss\u79D2", "21" },
			{ "yyyy\u5E74m\u6708", "22" },
			{ "m\u6708d\u65E5", "23" },
			{ "reserved", "24" },

			// more currency
			{ "#,##0;-#,##0", "25" },
			{ "#,##0;[Red]-#,##0", "26" },
			{ "#,##0.00;-#,##0.00", "27" },
			{ "#,##0.00;[Red]-#,##0.00", "28" },
			{ "_ * #,##0_ ;_ * -#,##0_ ;_ * -_ ;_ @_ ", "29" },
			{ "_ \u00a5* #,##0_ ;_ \u00a5* -#,##0_ ;_ \u00a5* -_ ;_ @_ ", "2a" },
			{ "_ * #,##0.00_ ;_ * -#,##0.00_ ;_ * -_ ;_ @_ ", "2b" },
			{ "_ \u00a5* #,##0.00_ ;_ \u00a5* -#,##0.00_ ;_ \u00a5* -_ ;_ @_ ", "2c" },

			// misc
			{ "mm:ss", "2d" },
			{ "[h]:mm:ss", "2e" },
			{ "mm:ss.0", "2f" },
			{ "##0.0E+0", "30" },
			{ "@", "31" },
			{ "yyyy\u5E74m\u6708", "37" },
			{ "m\u6708d\u65E5", "38" },
			{ "[$-411]ge.m.d", "39" },
			{ "[$-411]ggge\u5E74m\u6708d\u65E5", "3a" },
	};

	/* The NumberFormat element contains one of the following string constants 
	 * specifying the date or time format: 
	 * General Date, Short Date, Medium Date, Long Date, Short Time, Medium Time, or Long Time; 
	 * one of the following string constants specifying the numeric format: 
	 * Currency, Fixed, General, General Number, Percent, Scientific, or Standard (#,##0.00);
	 */
	// TODO: Add
	// TODO: These may not be correct mappings
	 /* one of the following string constants specifying the Boolean values displayed: On/Off, True/False, or Yes/No;*/
	public static String[][] BUILTIN_FORMATS_MHTML = {
			{ "General Date", "M/D/YY" },
			{ "Short Date", "D-MMM" },
			{ "Medium Date", "D-MMM-YY" },
			{ "Long Date", "MMM-YY" },
			{ "Short Time", "mm:ss" },
			{ "Medium Time", "h:mm AM/PM" },
			{ "Long Time", "h:mm:ss AM/PM" },
			{ "Currency", "\"$\"#,##0_);(\"$\"#,##0)" },
			{ "Fixed", "0.00" },
			{ "General Number", "0" },
			{ "Percent", "0%" },
			{ "Scientific", "0.00E+00" },
			{ "Standard", "(#,##0.00);" },

	};

	/**
	 * - Now this array includes the java format pattern.  If multiple patterns exist for negative
	 * numbers, that is after a semicolon.  The [Red] needs to be stripped as well and applied outside
	 * the Java formatting if desired.
	 */
	public static String[][] NUMERIC_FORMATS = {
			{ "General", "0", "" },
			{ "0", "1", "0" },
			{ "0.00", "2", "0.00" },

			// add below 2 to match excel
			{ "#,##0", "3", "#,##0" },
			{ "#,##0.00", "4", "#,##0.00" },

			{ "0%", "9", "0%" },
			// percent
			{ "0.00%", "a", "#.00%" },
			// percent
			{ "0.00E+00", "b", "0.00E00" },
			// scientific
			{ "(#,##0.00_);(#,##0.00)", "27", "#,##0.00;(#,##0.00)" },
			//?
			{ "_(*#,##0_);_(*(#,##0);_(*\"-\"_);_(@_)", "29", "#,##0;(#,##0)" },
			//?
			{ "##00E+0", "30", "#.##E00" },
			//TODO: unable to write a good generic conversion string here, returning generic exponent
			{ "@", "31", "@" },
			//?
			// User-defined (Format ID > 164, but variable.  use Format ID= -1)
			{ "0.0", "FF", "0.0" },
			{ "0.00;[Red]0.00", "FF", "#,##0.00;[Red]#,##0.00" },
			{ "0.00_);(0.00)", "FF", "#,##0.00###;(#,##0.00)" },
			{ "0.00_);[Red](0.00)", "FF", "#,##0.00###;[Red](#,##0.00)" },
	};

// Decoded format Patterns, HEX ids, and java format strings for all date formats 
	/**
	 * These formats are used in isDate handling and translating those into java format strings when possible
	 * for toFormattedString behavior
	 */
	public static String[][] DATE_FORMATS = {
			{ "mm-dd-yy", "E", "MM-dd-yy" },
			{ "d-mmm-yy", "F", "d-MMM-yy" },
			{ "d-mmm", "10", "d-MMM" },
			{ "mmm-yy", "11", "MMM-yy" },
			{ "h:mm AM/PM", "12", "h:mm aa" },
			{ "h:mm:ss AM/PM", "13", "h:mm:ss aa" },
			{ "h:mm", "14", "H:mm" },
			{ "h:mm:ss", "15", "H:mm:ss" },
			{ "m/d/yy h:mm", "16", "M/d/yy H:mm" },
			{ "mm:ss", "2D", "mm:ss" },
			{ "[h]:mm:ss", "2E" },
			{ "mm:ss.0", "2F", "mm:ss.S" },
			// user-defined (Format ID > 164, but variable) - use Format ID of -1
			{ "m/d/yyyy h:mm", "FF", "M/d/yyyy H:mm" },
			{ "m/d/yyyy;@", "FF", "M/d/yyyy" },
			{ "m/d/yy h:mm", "FF", "M/d/yy H:mm" },
			{ "hh:mm AM/PM", "FF", "hh:mm aa" },
			{ "m/d/yyyy", "FF", "M/d/yyyy" },
			{ "yyyy-mm-dd", "FF", "yyyy-MM-dd" },
			// i know, weird formatting offset, but the xf says m/d/y and excel shows m/d/yyyy
			{ "m/d/y", "FF", "M/d/yyyy" },
			{ "m/d/yy", "FF", "M/d/yy" },
			{ "m/d;@", "FF", "M/d" },
			{ "m/d/yy;@", "FF", "M/d/yy" },
			{ "mm/dd/yy;@", "FF", "MM/dd/yy" },
			{ "mm/dd/yy", "FF", "MM/dd/yy" },
			{ "mm/dd/yyyy", "FF", "MM/dd/yyyy" },
			{ "dd/mm/yy", "FF", "dd/MM/yy" },
			{ "yy/mm/dd;@", "FF", "yy/MM/dd" },
			{ "m/d/yy h:mm;@", "FF", "M/d/yy H:mm" },
			{ "hh:mm:ss.000", "FF", "HH:mm:ss.SSS" },
			{ "mmmm dd, yyyy", "FF", "MMMMM dd, yyyy" },
			{ "mmm dd, yyyy", "FF", "MMM dd, yyyy" },
			{ "m/d;@", "FF", "M/d" },
			{ "[$-409]d-mmm;@", "FF", "d-MMM" },
			{ "[$-409]d-mmm-yy;@", "FF", "d-MMM-yy" },
			{ "[$-409]d-mmm-yyyy;@", "FF", "d-MMM-yyyy" },
			{ "[$-409]dd-mmm-yy;@", "FF", "dd-Myyyy-mm-ddMM-yy" },
			{ "[$-409]dddd, mmmm dd, yyyy", "FF", "EEEE, MMMM dd, yyyy" },
			{ "[$-409]mmm-yy;@", "FF", "MMM-ss" },
			{ "[$-409]mmmm-yy;@", "FF", "MMMM-yy" },
			{ "[$-409]mmmm d, yyyy;@", "FF", "MMMM d, yyyy" },
			{ "[$-409]mmmmm;@", "FF", "MMMMM" },
			{ "[$-409]mmmmm-yy;@", "FF", "MMMMM-yy" },
			{ "[$-409]m/d/yy h:mm AM/PM;@", "FF", "M/d/yy H:mm aa" },
			{ "[$-409]mmmm d, yyyy;@", "FF", "MMMM d, yyyy" },
			{ "[$-409]h:mm:ss AM/PM", "FF", "H:mm:ss aa" },
			{ "[$-409]d-mmm;@", "FF", "d-MMM" },
			{ "[$-F800]dddd,mmmm dd,yyyy", "FF", "EEEE,MMMM dd, yyyy" },
			{ "[$-F800]dddd, mmmm dd, yyyy", "FF", "EEEE, MMMM dd, yyyy" },
			{ "dd/mm/yy hh:mm", "FF", "dd/MM/yy HH:mm" },
			{ "d-mmm-yyyy hh:mm", "FF", "d-MMM-yyyy HH:mm" },
			{ "yyyy-mm-dd hh:mm", "FF", "yyyy-MM-dd HH:mm" },
			{ "yyyy-mm-ss hh:mm:ss.000", "FF", "yyyy-MM-dd HH:mm:ss.SSS" },
			{ "yyyy-mm-dd", "FF", "yyyy-MM-dd" },
			{ "d/mm/yy", "FF", "d/MM/yy" },
			{ "d/mm/yy;@", "FF", "d/MM/yy" },
			{ "dddd, mmmm dd, yyyy", "FF", "EEEEE, MMMMM d, yyyy" },
			{ "[hhh]:mm", "FF", "h:mm" },	/* ?? */

	};
	// Decoded Format Patterns and hex Format Ids
	public static String[][] CURRENCY_FORMATS = {
// below 4 do not match excel!	   
//   	{"$#,##0;($#,##0)", "1", "$#,##0;($#,##0)"},
//		{"$#,##0;[Red]($#,##0)", "2", "$#,##0;[Red]($#,##0)"},
//		{"$#,##0.00;($#,##0.00)", "3", "$#,##0.00;($#,##0.00)"},
//		{"$#,##0.00;[Red]($#,##0.00)", "4", "$#,##0.00;[Red]($#,##0.00)"},
			{ "#,##0_;($#,##0)", "5", "#,##0;($#,##0)" },
			{ "#,##0_;[Red]($#,##0)", "6", "#,##0;[Red]($#,##0)" },
			{ "#,##0.00_;($#,##0.00)", "7", "#,##0.00###;($#,##0.00)" },
			{ "#,##0.00_;[Red]($#,##0.00)", "8", "#,##0.00;[Red]($#,##0.00)" },
			{ "_$*#,##0_;_($*($#,##0);_($*\"-\"_);_(@_)", "2a", "$#,##0;($#,##0)" },
			{ "_*#,##0_;_(*($#,##0);_(*\"-\"??_);_(@_)", "2b", "#,##0;(#,##0)" },
			{ "_$*#,##0_;_($*($#,##0);_($*\"-\"??_);_(@_)", "2c", "$#,##0;($#,##0)" },
			// User-defined (Format ID > 164, but variable.  use Format ID= -1)
			{ "_$*#,##0_;_($*($#,##0);_($*\"-\"??_);_(@_)", "FF", "$#,##0;($#,##0)" },
			{ "#,##0.00 [$-1]_);[Red](#,##0.00 [$-1])", "FF", "#,##0.00 ;[Red](#,##0.00 )" },
			{ "_ * #,##0.00_) [$-1]_ ;_ * (#,##0.00) [$-1]_ ;_ * -??_) [$-1]_ ;_ @_ ", "FF", "#,##0.00 ;(#,##0.00) " },
			{ "_ * #,##0.00_) [$€-1]_ ;_ * (#,##0.00) [$€-1]_ ;_ * -??_) [$€-1]_ ;_ @_ ", "FF", "#,##0.00 €;(#,##0.00) €" },
			{ "#,##0.00 [$€-1]", "FF", "#,##0.00 €" },
			{ "#,##0.00 [$-1]", "FF", "#,##0.00 " },
			{ "$#,##0.00", "FF", "$#,##0.00" },
			{ "_($* #,##0_);_($* (#,##0);_($* -??_);_(@_)", "FF", "$ #,##0;($* (#,##0)" },
			{ "#,##0.00 [$-1];[Red]#,##0.00 [$-1]", "FF", "#,##0.00 ;[Red]#,##0.00 " },
			{ "[$-411]#,##0.00", "FF", "#,##0.00" },
			{ "[$-411]#,##0.00;[Red][$-411]#,##0.00", "FF", "#,##0.00;[Red]#,##0.00" },
			{ "[$-411]#,##0.00;-[$-411]#,##0.00", "FF", "#,##0.00" },
			{ "[$-411]#,##0.00;[Red]-[$-411]#,##0.00", "FF", "#,##0.00;[Red]-#,##0.00" },
			{ "[$-809]#,##0.00", "FF", "#,##0.00" },
			{ "[$-809]#,##0.00;[Red][$-809]#,##0.00", "FF", "#,##0.00;[Red]#,##0.00" },
			{ "[$-809]#,##0.00;-[$-809]#,##0.00", "FF", "#,##0.00;-#,##0.00" },
			{ "[$-809]#,##0.00;[Red]-[$-809]#,##0.00", "FF", "#,##0.00;[Red]-#,##0.00" },
			{ "#,##0.0 [$$-C0C];[Red]#,##0.0 [$$-C0C]", "FF" },
			// user-defined
			{ "$#,##0;($#,##0)", "FF", "$#,##0;($#,##0)" },
			{ "$#,##0.00;($#,##0.00)", "FF", "$#,##0.00;($#,##0.00)" },
			{ "$#,##0;[Red]($#,##0)", "FF", "$#,##0;[Red]($#,##0)" },
			{ "$#,##0.00;[Red]($#,##0.00)", "FF", "$#,##0.00;[Red]($#,##0.00)" },
			{ "$#,##0.00;[Red]$#,##0.00", "FF", "$#,##0.00;[Red]$#,##0.00" },
			{ "$#,##0", "FF", "$#,##0" }
	};

	// this maps Java formats to Excel formats
	public static String[][] patternmap = {
			// currency
			{ "#,##0.00", "#,##0.00" },
			{ "$#,##0;($#,##0)", "(#,##0_);($#,##0)" },
			{ "$#,##0;[Red]($#,##0)", "$###,###;($###,###)" },
			{ "$#,##0.00;($#,##0.00)", "$###,###.##;($###,###.##)" },
			{ "$#,##0.00;[Red]($#,##0.00)", "$###,###.##;($###,###.##)" },
			{ "0%", "0%" },
			{ "0.00%", "0.00%" },
			{ "0.00E+00", "0.00E+00" },
			{ "# ?/?", "# ?/?" },
			{ "# ??/??", "" },
			{ "#,##0_;($#,##0)", "" },
			{ "#,##0_;[Red]($#,##0)", "" },
			{ "#,##0.00_;($#,##0.00)", "" },
			{ "#,##0.00_;[Red]($#,##0.00)", "" },
			{ "_*#,##0_;_(*($#,##0);_(*-_);_(@_)", "" },
			{ "_$*#,##0_;_($*($#,##0);_($*-_);_(@_)", "" },
			{ "_*#,##0_;_(*($#,##0);_(*-??_);_(@_)", "" },
			{ "_$*#,##0_;_($*($#,##0);_($*-??_);_(@_)", "" },

			// date time
			{ "mm-mmm-yy", "MM-dd-YY" },
			{ "m/d/y", "M/d/y" },
			{ "d-mmm-yy", "d-MMM-yy" },
			{ "d-mmm", "d-MMM" },
			{ "mmm-yy", "MMM-yy" },
			{ "h:mm AM/PM", "hh:mm a" },
			{ "h:mm:ss AM/PM", "hh:mm:ss a" },
			{ "h:mm", "hh:mm" },
			{ "h:mm:ss", "hh:mm:ss" },
			{ "m/d/yy h:mm", "M/d/yy hh:mm" },
	};

	// look up user-friendly format string and get Excel-qualified or encoded format string
	// USED???
	public static String[][] EXCEL_FORMAT_LOOKUP = {
			// currency
			{ "0", "" },
			{ "0.00", "" },
			{ "#,##0", "" },
			{ "#,##0.00", "" },
			{ "($#,##0);($#,##0)", "\"$\"#,##0);(\"$\"#,##0)" },
			{ "($#,##0);[Red]($#,##0)", "\"$\"#,##0);[Red](\"$\"#,##0)" },
			{ "($#,##0.00);($#,##0.00)", "\"$\"#,##0.00);(\"$\"#,##0.00)" },
			{ "($#,##0.00);[Red]($#,##0.00)", "\"$\"#,##0.00);[Red](\"$\"#,##0.00)" },
			{ "0%", "" },
			{ "0.00%", "" },
			{ "0.00E+00", "" },
			{ "# ?/?", "" },
			{ "# ??/??", "" },
			// Are these correct?  Should they be qualified???
			{ "(#,##0_);($#,##0)", "" },
			{ "(#,##0_);[Red]($#,##0)", "" },
			{ "(#,##0.00_);($#,##0.00)", "" },
			{ "(#,##0.00_);[Red]($#,##0.00)", "" },

			{ "(#,##0_);(#,##0)", "" },
			{ "(#,##0_);[Red](#,##0)", "" },
			{ "(#,##0.00_);(#,##0.00)", "" },
			{ "(#,##0.00_);[Red](#,##0.00)", "" },
			{ "_(*#,##0_);_(*($#,##0);_(*-_);_(@_)", "_(*#,##0_);_(*(\"$\"#,##0);_(*\"-\"_);_(@_)" },
			{ "_($*#,##0_);_($*($#,##0);_($*-_);_(@_)", "_(\"$\"*#,##0_);_(\"$\"*($#,##0);_(\"$\"*\"-\"_);_(@_)" },
			{ "_(*#,##0_);_(*($#,##0);_(*-??_);_(@_)", "_(*#,##0_);_(*(\"$\"#,##0);_(*\"-\"??_);_(@_)" },
			{ "_($*#,##0_);_($*($#,##0);_($*-??_);_(@_)", "_(\"$\"*#,##0_);_(\"$\"*(\"$\"#,##0);_(\"$\"*\"-\"??_);_(@_)" },

			// dates
			{ "mm-mmm-yy", "" },
			{ "m/d/y", "" },
			{ "d-mmm-yy", "d\\-mmm\\-yy" },
			{ "[$-409]mmmm d, yyyy;@", "mmmm\\d\\,\\-yyyy" },
			{ "d-mmm", "d\\-mmm" },
			{ "mmm-yy", "mmm\\-yy" },
			{ "h:mm AM/PM", "h:mm\\ AM/PM" },
			{ "h:mm:ss AM/PM", "" },
			{ "h:mm", "" },
			{ "h:mm:ss", "" },
			{ "m/d/yy h:mm", "" },
	};
	// Alignment's
	public static int ALIGN_DEFAULT = 0;
	public static int ALIGN_LEFT = 1;
	public static int ALIGN_CENTER = 2;
	public static int ALIGN_RIGHT = 3;
	public static int ALIGN_FILL = 4;
	public static int ALIGN_JUSTIFY = 5;
	public static int ALIGN_CENTER_ACROSS_SELECTION = 6;
	public static int ALIGN_VERTICAL_TOP = 0;
	public static int ALIGN_VERTICAL_CENTER = 1;
	public static int ALIGN_VERTICAL_BOTTOM = 2;
	public static int ALIGN_VERTICAL_JUSTIFY = 3;

	public static String[] VERTICAL_ALIGNMENTS = {
			"top", "middle", "bottom", "text-bottom",    // justified
			"text-bottom",    // distributed"
	};

	/**
	 * These Horizontal alignments are used to map to HTML output, so some
	 * values may not be as expected.  For instance, "Center across selection" is
	 * not something you can do in HTML.  We will use "Center" for this.
	 * <p/>
	 * Not entirely True: HTML can center across merged table cells -- this is basically
	 * the purpse AFAIK. -jm
	 */
	public static String[] HORIZONTAL_ALIGNMENTS = {
			"Default", "Left", "Center", "Right", "Center", "Center", "Center"
	};
}