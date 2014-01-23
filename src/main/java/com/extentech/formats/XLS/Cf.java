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

import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.Dxf;
import com.extentech.formats.OOXML.Fill;
import com.extentech.formats.XLS.formulas.FormulaParser;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgArray;
import com.extentech.formats.XLS.formulas.PtgRefN;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Stack;

/**
 * <b>Cf:  Conditional Formatting Conditions 0x1B1</b><br>
 * <p/>
 * This record stores a conditional formatting condition
 * <p/>
 * <p/>
 * There are some restrictions in the usage of conditional formattings:
 * In the user interface it is possible to modify the font style (boldness and posture), the text colour, the underline style
 * and the strikeout style. It is not possible to change the used font, the font height, and the escapement style, though it is
 * possible to specify a new font height and escapement style in this record which are correctly displayed
 * It is not possible to change a border line style, but to preserve the line colour, and vice versa. Diagonal lines are not
 * supported at all. The user interface only offers thin line styles, but files containing other line styles work correctly too.
 * It is not possible to set the background pattern colour to No colour (using system window background
 * <p/>
 * <p/>
 * <p/>
 * OFFSET		NAME			SIZE		CONTENTS
 * -----
 * 4				ct					1			Conditional formatting type
 * 5				cp					1			Conditional formatting operator
 * 6				cce1				2			Count of bytes in rgce1	-- size of formula data for first value or formula
 * 8				cce2				2			Count of bytes in rgce2	-- size of formula data for first value or formula: used for second part of 'between' and 'not between' comparison, else 0
 * 10									4			Option flags (see below)
 * 2			Not used
 * 16 								118 		(optional, only if font = 1, see option flags) Font formatting block, see below
 * var 								8			(optional, only if bord = 1, see option flags) Border formatting block, see below
 * var								4 			(optional, only if patt = 1, see option flags) Pattern formatting block, see below
 * var			rgbdxf			var		Conditional format to apply
 * var			rgce1			var		First formula for this condition (RPN token array without size field, ?4)
 * var			rgce2			var		Second formula for this condition (RPN token array without size field, ?4)
 * <p/>
 * Conditional formatting operator:
 * 00H = No comparison (only valid for formula type, see above)
 * 01H = Between 05H = Greater than
 * 02H = Not between 06H = Less than
 * 03H = Equal 07H = Greater or equal
 * 04H = Not equal 08H = Less or equal
 * <p/>
 * Option Flags
 * If none of the formatting attributes is set, the option flags field contains 00000000H.
 * The following table assumes that
 * the conditional formatting contains at least one modified formatting attribute
 * (it will occur at least one of the formatting
 * information blocks in the record). In difference to the first case some of the
 * bits are always set now.
 * All flags specifying that an attribute is modified are 02, if the conditional formatting
 * changes the respective attribute,
 * and 12, if the original cell formatting is preserved. The flags for modified font
 * attributes are not contained in this
 * option flags field, but in the font formatting block itself. !
 * <p/>
 * Bit Mask Contents
 * 9-0 000003FFH Always 11.1111.11112 (but not used)
 * 10 00000400H 0 = Left border style and colour modified (bord - left )
 * 11 00000800H 0 = Right border style and colour modified (bord - right )
 * 12 00001000H 0 = Top border style and colour modified (bord - top )
 * 13 00002000H 0 = Bottom border style and colour modified (bord - bot )
 * 15-14 0000C000H Always 112 (but not used)
 * 16 00010000H 0 = Pattern style modified (patt - style )
 * 17 00020000H 0 = Pattern colour modified (patt - col )
 * 18 00040000H 0 = Pattern background colour modified (patt - bgcol )
 * 21-19 00380000H Always 1112 (but not used)
 * 26 04000000H 1 = Record contains font formatting block (font)
 * 28 10000000H 1 = Record contains border formatting block (bord)
 * 29 20000000H 1 = Record contains pattern formatting block (patt)
 *
 * @see Condfmt
 */
public final class Cf extends com.extentech.formats.XLS.XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Cf.class );
	private static final long serialVersionUID = 5624169378370505532L;
	short ct = 0;        //		Conditional formatting type
	short cp = 0;        //		Conditional formatting operator
	int cce1 = 0;        //		Count of bytes in rgce1	
	int cce2 = 0;        //		Count of bytes in rgce2
	String rgbdxf = "";        //		Conditional format to apply
	String rgce1 = "";            //		First formula for this condition
	String rgce2 = "";            // 		Second formula for this condition
	int flags = 0;                    // option flags 20080303 KSC:
	boolean bHasFontBlock = false;        // ""
	boolean bHasBorderBlock = false;        // ""
	boolean bHasPatternBlock = false;    // ""

	private Ptg prefHolder = null;                // this is a placeholder ptg for the conditional format expression 
	private Stack expression1 = new Stack();    // 
	private Stack expression2 = new Stack();    // 
	private Formula formula1 = null;                // 
	private Formula formula2 = null;                //
	private String containsText = null;            // OOXML-specific if type==containsText, this will hold the actual text to test

	private Condfmt condfmt = null;
	// Offset Size Contents
	// 0 2 Fill pattern style:
	// Bit Mask Contents
	// 15-10 FC00H Fill pattern style (only if patt - style = 0, ?3.11)
	public static final int PATTERN_FILL_STYLE = 0xFC00;
	private int patternFillStyle = -1;

	// 2 2 Fill pattern colour indexes:
	// Bit Mask Contents
	// 6-0 007FH Colour index (?6.70) for pattern (only if patt - col = 0)
	// 13-7 3F80H Colour index (?6.70) for pattern background (only if patt - bgcol = 0)
	public static final int PATTERN_FILL_COLOR = 0x007F;
	public static final int PATTERN_FILL_BACK_COLOR = 0x3F80;
	private int patternFillColorsFlag = 0;
	private int patternFillColor = 0;
	private int patternFillColorBack = -1;
	private Fill fill = null;
	private Font font = null;

	// Offset Size Contents
	// 0 64 Not used
	// 64 4 Font height (in twips = 1/20 of a point); or FFFFFFFFH to preserve the cell font height
	private int fontHeight = -1;

	// 68 4 Font options:
	// Bit Mask Contents
	// 1 00000002H Posture: 0 = Normal; 1 = Italic (only if font - style = 0)
	// 7 00000080H Cancellation: 0 = Off; 1 = On (only if font - canc = 0)
	public static final int FONT_OPTIONS_POSTURE = 0x2;
	public static final int FONT_OPTIONS_CANCELLATION = 0x80;
	public static final int FONT_OPTIONS_POSTURE_NORMAL = 0;
	public static final int FONT_OPTIONS_POSTURE_ITALIC = 1;
	public static final int FONT_OPTIONS_CANCELLATION_OFF = 0;
	public static final int FONT_OPTIONS_CANCELLATION_ON = 1;
	private int fontOptsFlag = -1;

	// 72 2 Font weight (100-1000, only if font - style = 0).
	private int fontWeight = -1;

	// Standard values are 0190H (400) for normal text and 02BCH (700) for bold text.

	// 74 2 Escapement type (only if font - esc = 0):
	// 0000H = None; 0001H = Superscript; 0002H = Subscript
	public static final int FONT_ESCAPEMENT_NONE = 0x0;
	public static final int FONT_ESCAPEMENT_SUPER = 0x1;
	public static final int FONT_ESCAPEMENT_SUB = 0x2;
	private int fontEscapementFlag = -1;

	// 76 1 Underline type (only if font - underl = 0):
	public static final int FONT_UNDERLINE_NONE = 0x0;
	public static final int FONT_UNDERLINE_SINGLE = 0x1;
	public static final int FONT_UNDERLINE_DOUBLE = 0x2;
	public static final int FONT_UNDERLINE_SINGLEACCOUNTING = 0x21;
	public static final int FONT_UNDERLINE_DOUBLEACCOUNTING = 0x22;
	private int fontUnderlineStyle = -1;

	private int fontColorIndex = -1;

	public static final int FONT_MODIFIED_OPTIONS_STYLE = 0x00000002;
	public static final int FONT_MODIFIED_OPTIONS_CANCELLATIONS = 0x00000080;
	private int fontModifiedOptionsFlag = -1;
	private int fontEscapementFlagModifiedFlag = -1;
	private int fontUnderlineModifiedFlag = -1;

	/**
	 * Border Formatting Block
	 */
	//	Offset Size Contents
	//	0 2 Border line styles:
	//	Bit Mask Contents
	//	3-0 000FH Left line style (only if bord - left = 0, ?3.10)
	//	7-4 00F0H Right line style (only if bord - right = 0, ?3.10)
	//	11-8 0F00H Top line style (only if bord - top = 0, ?3.10)
	//	15-12 F000H Bottom line style (only if bord - bot = 0, ?3.10)
	public static final int BORDER_LINESTYLE_LEFT = 0x000F;
	public static final int BORDER_LINESTYLE_RIGHT = 0x00F0;
	public static final int BORDER_LINESTYLE_TOP = 0x0F00;
	public static final int BORDER_LINESTYLE_BOTTOM = 0xF000;
	// if flags & BORDER_MODIFIED_XX == 0 means that this particular border has been modified
	public static final int BORDER_MODIFIED_LEFT = 0x0400;
	public static final int BORDER_MODIFIED_RIGHT = 0x0800;
	public static final int BORDER_MODIFIED_TOP = 0x1000;
	public static final int BORDER_MODIFIED_BOTTOM = 0x2000;
	private short borderLineStylesFlag = 0;
	private int borderLineStylesLeft = -1;
	private int borderLineStylesRight = -1;
	private int borderLineStylesTop = -1;
	private int borderLineStylesBottom = -1;

	// 2 4 Border line colour indexes:
	// 	Bit Mask Contents
	//	6-0 0000007FH Colour index (?6.70) for left line (only if bord - left = 0)
	//	13-7 00003F80H Colour index (?6.70) for right line (only if bord - right = 0)
	//	22-16 007F0000H Colour index (?6.70) for top line (only if bord - top = 0)
	//	29-23 3F800000H Colour index (?6.70) for bottom line (only if bord - bot = 0)
	public static final int BORDER_LINECOLOR_LEFT = 0x0000007F;
	public static final int BORDER_LINECOLOR_RIGHT = 0x00003F80;
	public static final int BORDER_LINECOLOR_TOP = 0x007F0000;
	public static final int BORDER_LINECOLOR_BOTTOM = 0x3F800000;
	private int borderLineColorsFlag = 0;
	private int borderLineColorLeft = 0;
	private int borderLineColorRight = 0;
	private int borderLineColorTop = 0;
	private int borderLineColorBottom = 0;

	// TODO: finish Cf prototype bytes + formatting options such as font block
	//private byte[] PROTOTYPE_BYTES = {1, 5, 3, 0, 0, 0, -1, -1, 59, -92, 2, -128, 0, 0, 2, 0, 2, 0, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 4, 0, 6, -36, 0, -76, 5, 83, -17, -1, -1, -1, -1, 0, 0, 0, 0, -1, -1, -1, -1, -1, 0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 0, -97, -1, -1, -1, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, -64, 22, 30, 123, 0};
	// a very basic cf which only contains a pattern fill block and function compares <= 100 
//    private byte[] PROTOTYPE_BYTES= {1, 8, 3, 0, 0, 0, -1, -1, 59, -96, 2, -128, 0, 0, -64, 26, 30, 100, 0};	 
	private byte[] PROTOTYPE_BYTES = { 1, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };//,  -64, 26, 30, 100, 0};	 

//    0x1,0x1,0x3,0x0,0x3,0x0,0xf,0xc,0x3f,0x90,0x2,0x80,0x11,0x11,0x14,0xa,0x14,0xa,0x0,0x0,0x1e,0x7b,0x0,0x1e,0xea,0x0};    

	/**
	 * default constructor
	 */
	public Cf()
	{
	}

	/**
	 * constructor which takes the cellrange of cells
	 * and the Condfmt that reference this conditional format rule
	 *
	 * @param f the condfmt
	 * @param r the cellrange
	 */
	public Cf( Condfmt f )
	{
		this();
		setData( this.PROTOTYPE_BYTES );
		this.setCondfmt( f );
	}

	/** Pattern Formatting Block */
	/**
	 * initialize the Cf record
	 */
	@Override
	public void init()
	{
		super.init();
		data = this.getData();
		ct = this.getByteAt( 0 );
		cp = this.getByteAt( 1 );
		cce1 = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		cce2 = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		//parsing of formula refs
		flags = ByteTools.readInt( this.getByteAt( 6 ), this.getByteAt( 7 ), this.getByteAt( 8 ), this.getByteAt( 9 ) );
		// Font formatting Block 
		bHasFontBlock = ((flags & 0x04000000) == 0x04000000);
		// Border Formatting Block
		bHasBorderBlock = ((flags & 0x10000000) == 0x10000000);
		// Pattern Formating Block
		bHasPatternBlock = ((flags & 0x20000000) == 0x20000000);
		int pos = 12;

		if( bHasFontBlock )
		{ // handle Font formatting section
			pos += 64; // 1st 64 bits of font block is unused  
			// 64 4 Font height (in twips = 1/20 of a point); or FFFFFFFFH to preserve the cell font height
			fontHeight = ByteTools.readInt( getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ) );

			// 68 4 Font options:
			// Bit Mask Contents
			// 1 00000002H Posture: 0 = Normal; 1 = Italic (only if font - style = 0)
			// 7 00000080H Cancellation: 0 = Off; 1 = On (only if font - canc = 0)
			fontOptsFlag = ByteTools.readInt( getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ) );

			// 72 2 Font weight (100-1000, only if font - style = 0).
			// Standard values are 0190H (400) for normal text and 02BCH (700) for bold text.
			fontWeight = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) );

			// 74 2 Escapement type (only if font - esc = 0):
			// 0000H = None; 0001H = Superscript; 0002H = Subscript
			// FONT_ESCAPEMENT_NONE 	= 0x0;
			// FONT_ESCAPEMENT_SUPER 	= 0x1;
			// FONT_ESCAPEMENT_SUB 	= 0x2;
			fontEscapementFlag = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) );

			// 76 1 Underline type (only if font - underl = 0):
			// 00H = None
			// 01H = Single 21H = Single accounting
			// 02H = Double 22H = Double accounting
			// FONT_UNDERLINE_NONE 	= 0x0;
			// FONT_UNDERLINE_SINGLE 	= 0x1;
			// FONT_UNDERLINE_DOUBLE 	= 0x2;
			fontUnderlineStyle = getByteAt( pos++ );

			// 77 3 Not used
			pos += 3;

			// 80 4 Font colour index (?6.70); or FFFFFFFFH to preserve the cell font colour
			fontColorIndex = ByteTools.readInt( getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ) );

			// 84 4 Not used
			pos += 4;

			// 88 4 Option flags for modified font attributes:
			// Bit Mask Contents
			// 1 00000002H 0 = Font style (posture or boldness) modified (font - style )
			// 4-3 00000018H Always 112 (but not used)
			// 7 00000080H 0 = Font cancellation modified (font - canc )
			fontModifiedOptionsFlag = ByteTools.readInt( getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ) );

			// 92 4 0 = Escapement type modified (font - esc )
			fontEscapementFlagModifiedFlag = ByteTools.readInt( getByteAt( pos++ ),
			                                                    getByteAt( pos++ ),
			                                                    getByteAt( pos++ ),
			                                                    getByteAt( pos++ ) );

			// 96 4 0 = Underline type modified (font - underl )
			fontUnderlineModifiedFlag = ByteTools.readInt( getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ) );
			// 100 16 Not used
			pos += 16;
			// 116 2 0001H
			pos += 2;
			if( pos != 130 )
			{
				log.warn( "Cf font block parsing pos mismatch" + pos );
			}
			getFont();
		}
		if( bHasBorderBlock )
		{

			//	Offset Size Contents
			//	0 2 Border line styles:
			//	Bit Mask Contents
			//	3-0 000FH Left line style (only if bord - left = 0, ?3.10)
			//	7-4 00F0H Right line style (only if bord - right = 0, ?3.10)
			//	11-8 0F00H Top line style (only if bord - top = 0, ?3.10)
			//	15-12 F000H Bottom line style (only if bord - bot = 0, ?3.10)
			// BORDER_LINESTYLE_LEFT 	= 0x000F;
			// BORDER_LINESTYLE_RIGHT 	= 0x00F0;
			// BORDER_LINESTYLE_TOP 	= 0x0F00;
			// BORDER_LINESTYLE_BOTTOM = 0xF000;
			borderLineStylesFlag = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) );
			updateBorderLineStyles();

			// 2 4 Border line colour indexes:
			// 	Bit Mask Contents
			//	6-0 0000007FH Colour index (?6.70) for left line (only if bord - left = 0)
			//	13-7 00003F80H Colour index (?6.70) for right line (only if bord - right = 0)
			//	22-16 007F0000H Colour index (?6.70) for top line (only if bord - top = 0)
			//	29-23 3F800000H Colour index (?6.70) for bottom line (only if bord - bot = 0)
			// BORDER_LINECOLOR_LEFT 	= 0x0000007F;
			// BORDER_LINECOLOR_RIGHT 	= 0x00003F80;
			// BORDER_LINECOLOR_TOP 	= 0x007F0000;
			// BORDER_LINECOLOR_BOTTOM = 0x3F800000;
			borderLineColorsFlag = ByteTools.readInt( getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ), getByteAt( pos++ ) );
			updateBorderLineColors();

			//6 2 Not used
			pos += 2;
		}
		if( bHasPatternBlock )
		{
			// Offset Size Contents

			// 0 2 Fill pattern style:
			// Bit Mask Contents
			// 15-10 FC00H Fill pattern style (only if patt - style = 0, ?3.11)
			// PATTERN_FILL_STYLE = 0xFC00;
			patternFillStyle = (ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) ) >> 9);
/* very strangely, patternFillStyle appears to be 0 when it should be 1 (solid)
 * found in testing through our test suite ...
 * according to documetation this algorithm *appears* ok but the doc is really confusing on cf's
 */

			// 2 2 Fill pattern colour indexes:

			// Bit Mask Contents
			// 6-0 007FH Colour index (?6.70) for pattern (only if patt - col = 0)
			// 13-7 3F80H Colour index (?6.70) for pattern background (only if patt - bgcol = 0)
			// PATTERN_FILL_COLOR 		= 0x007F;
			// PATTERN_FILL_BACK_COLOR 	= 0x3F80;
			patternFillColorsFlag = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) );

			// update the current vals
			updatePatternFillColors();

		}

		int postest = 12;
		if( bHasFontBlock )
		{
			postest += 118;
		}
		if( bHasBorderBlock )
		{
			postest += 8;
		}
		if( bHasPatternBlock )
		{
			postest += 4;
		}

		if( postest != pos )
		{
			log.warn( "Cf bad pos offset during init()." );
			pos = postest;
		}
		// 1st formula data= pos->cce1
		byte[] function = this.getBytesAt( pos, cce1 );
		try
		{
			expression1 = ExpressionParser.parseExpression( function, this, cce1 );
		}
		catch( Exception e )
		{
			log.error( "Initializing expression1 for Cf failed: " + new String( function ) );
		}
		pos += cce1;
		// 2nd formula data= pos+cce1->cce2
		function = this.getBytesAt( pos, cce2 );
		if( cce2 > 0 )
		{
			try
			{
				expression2 = ExpressionParser.parseExpression( function, this, cce2 );
			}
			catch( Exception e )
			{
				log.error( "Initializing expression2 for Cf failed: " + new String( function ) );
			}
		}
			log.trace( "Cf record encountered." );
	}

	/**
	 * take current state of Cf and update record
	 */
	private void updateRecord()
	{
		byte[] newdata = new byte[12]; // enough room for basics; formatting blocks will be appended
		newdata[0] = (byte) ct;
		newdata[1] = (byte) cp;
		flags = 0x3FF | 0xC000 | 0x380000;    // set required and unused bits of flag
		if( bHasFontBlock )
		{
			flags = (flags | 0x04000000);
		}
		if( bHasBorderBlock )
		{
			flags = (flags | 0x10000000);
			flags |= 0x400;    // left border mod
			flags |= 0x800;    // right border mod
			flags |= 0x1000;    // top border mod
			flags |= 0x2000;    // bottom border mod
		}
		if( bHasPatternBlock )
		{
			flags = (flags | 0x20000000);
			flags |= 0x10000;    // patt style mod
			flags |= 0x20000;    // patt fg color (pattern color) mod
			flags |= 0x40000;    // pat bg color mod
		}

		byte[] b = ByteTools.cLongToLEBytes( flags );
		System.arraycopy( b, 0, newdata, 6, 4 );

		int pos = 12;
		if( bHasFontBlock )
		{ // update font section
			newdata = ByteTools.append( new byte[118], newdata );
			pos += 64; // 1st 64 bits of font block is unused
			b = ByteTools.cLongToLEBytes( fontHeight );
			System.arraycopy( b, 0, newdata, pos, 4 );
			pos += 4;
			b = ByteTools.cLongToLEBytes( fontOptsFlag );
			System.arraycopy( b, 0, newdata, pos, 4 );
			pos += 4;
			b = ByteTools.shortToLEBytes( (short) fontWeight );
			System.arraycopy( b, 0, newdata, pos, 2 );
			pos += 2;
			b = ByteTools.shortToLEBytes( (short) fontEscapementFlag );
			System.arraycopy( b, 0, newdata, pos, 2 );
			pos += 2;
			b = ByteTools.shortToLEBytes( (short) fontUnderlineStyle );
			System.arraycopy( b, 0, newdata, pos, 1 );
			pos += 1;
			// 77 3 Not used
			pos += 3;
			b = ByteTools.cLongToLEBytes( fontColorIndex );
			System.arraycopy( b, 0, newdata, pos, 4 );
			pos += 4;
			// 84 4 Not used
			pos += 4;            // 88:
			b = ByteTools.cLongToLEBytes( fontModifiedOptionsFlag );
			System.arraycopy( b, 0, newdata, pos, 4 );
			pos += 4;
			b = ByteTools.cLongToLEBytes( fontEscapementFlagModifiedFlag );
			System.arraycopy( b, 0, newdata, pos, 4 );
			pos += 4;
			b = ByteTools.cLongToLEBytes( fontUnderlineModifiedFlag );
			System.arraycopy( b, 0, newdata, pos, 4 );
			pos += 4;
			// 100 16 Not used
			pos += 16;
			// 116 2 0001H
			pos++;
			newdata[pos++] = 1;
		}
		if( bHasBorderBlock )
		{
			newdata = ByteTools.append( new byte[8], newdata );
			b = ByteTools.shortToLEBytes( borderLineStylesFlag );
			System.arraycopy( b, 0, newdata, pos, 2 );
			pos += 2;
			b = ByteTools.cLongToLEBytes( borderLineColorsFlag );
			System.arraycopy( b, 0, newdata, pos, 4 );
			pos += 4;
			//6 2 Not used
			pos += 2;
		}
		if( bHasPatternBlock )
		{
			newdata = ByteTools.append( new byte[4], newdata );
			b = ByteTools.shortToLEBytes( (short) (patternFillStyle << 9) );
			System.arraycopy( b, 0, newdata, pos, 2 );
			pos += 2;
			b = ByteTools.shortToLEBytes( (short) patternFillColorsFlag );
			System.arraycopy( b, 0, newdata, pos, 2 );
			pos += 2;
		}

		if( formula1 != null )
		{
			byte[] function = getFormulaExpression( formula1 );
			newdata = ByteTools.append( function, newdata );
			cce1 = function.length;
			b = ByteTools.shortToLEBytes( (short) cce1 );
			newdata[2] = b[0];
			newdata[3] = b[1];
		}
		if( formula2 != null )
		{
			byte[] function = getFormulaExpression( formula2 );
			newdata = ByteTools.append( function, newdata );
			cce2 = function.length;
			b = ByteTools.shortToLEBytes( (short) cce2 );
			newdata[4] = b[0];
			newdata[5] = b[1];
		}
		this.setData( newdata );
//        this.init(); DO NOT DO AS can overwrite OOXML-specifics
	}

	/**
	 * return the expression bytes of the specified formula
	 *
	 * @param f
	 * @return
	 */
	private byte[] getFormulaExpression( Formula f )
	{
		boolean hasArray = false;
		byte[] expbytes = new byte[0];
		byte[] arraybytes = null;
		Stack expression = f.getExpression();
		for( int i = 0; i < expression.size(); i++ )
		{
			Object o = expression.elementAt( i );
			Ptg ptg = (Ptg) o;
			byte[] b;
			if( o instanceof PtgArray )
			{
				PtgArray pa = (PtgArray) o;
				b = pa.getPreRecord();
				arraybytes = ByteTools.append( pa.getPostRecord(), arraybytes );
				hasArray = true;
			}
			else
			{
				b = ptg.getRecord();
			}
			expbytes = ByteTools.append( b, expbytes );
		}
		if( hasArray )
		{
			expbytes = ByteTools.append( arraybytes, expbytes );
		}
		return expbytes;
	}

	/**
	 * helper method to set the style values on this object from a style string
	 * <br>each name/value pair is delimited by ;
	 * <br>possible tokens:
	 * <br>pattern		pattern fill #
	 * <br>color     	pattern fg color
	 * <br>patterncolor     pattern bg color
	 * <br>vertical	vertical alignnment
	 * <br>horizontal	horizontal alignment
	 * <br>border		sub-tokens: border-top, border-left, border-bottom, border-top
	 *
	 * @param style
	 * @param cx
	 */
	public static void setStylePropsFromString( String style, Cf cx )
	{
		String[] toks = StringTool.getTokensUsingDelim( style, ";" );
		// iterate styles, set values
		for( String tok : toks )
		{
			int pos = tok.indexOf( ":" );
			String n = "";
// TODO: border line sizes, interpret border line styles
// TODO: handle vertical, horizontal alignment, number format ...
			if( pos > 0 )
			{
				n = tok.substring( 0, pos );
			}

			String v = tok.substring( tok.indexOf( ":" ) + 1 );
			n = StringTool.strip( n, '"' );
			v = StringTool.strip( v, '#' );
			n = StringTool.allTrim( n );
			v = StringTool.allTrim( v );
			n = n.toLowerCase();
			if( n.indexOf( "border" ) == 0 )
			{ // parse the border settings
				String[] vs = StringTool.getTokensUsingDelim( v, " " );
				int sz = -1, cz = -1;
				int stl = -1;
				String clr = null;
				try
				{
					sz = Integer.parseInt( vs[0] );    // size
					clr = vs[2];                    // color String
					// TODO: no way to set border size as of yet

					// interpret style string into #:
					for( int i = 0; i < FormatConstants.BORDER_NAMES.length; i++ )
					{
						if( FormatConstants.BORDER_NAMES[i].equals( vs[1] ) )
						{
							stl = i;
							break;
						}
					}

					// HexStringToColorInt
					// java.awt.Color c = FormatHandle.HexStringToColor(clr);
					cz = FormatHandle.HexStringToColorInt( clr, FormatHandle.colorBACKGROUND );
					//	Logger.logInfo("ix " + sz + " " + clr + " " + stl );

				}
				catch( Exception ex )
				{
					;
				}
				if( n.indexOf( "border-top" ) == 0 )
				{
					if( clr != null )
					{ // set color
						cx.setBorderLineColorTop( cz );
					}
					if( stl > -1 )
					{
						cx.setBorderLineStylesTop( stl );
					}
				}
				else if( n.indexOf( "border-left" ) == 0 )
				{
					if( clr != null )
					{ // set color
						cx.setBorderLineColorLeft( cz );
					}
					if( stl > -1 )
					{
						cx.setBorderLineStylesLeft( stl );
					}
				}
				else if( n.indexOf( "border-bottom" ) == 0 )
				{
					if( clr != null )
					{ // set color
						cx.setBorderLineColorBottom( cz );
					}
					if( stl > -1 )
					{
						cx.setBorderLineStylesBottom( stl );
					}
				}
				else if( n.indexOf( "border-right" ) == 0 )
				{
					if( clr != null )
					{ // set color
						cx.setBorderLineColorRight( cz );
					}
					if( stl > -1 )
					{
						cx.setBorderLineStylesRight( stl );
					}
				}

			}
			else if( n.equalsIgnoreCase( "text-line-through" ) )
			{
				if( v.equalsIgnoreCase( "none" ) )
				{
					cx.setFontStriken( false );
				}
				else
				{
					cx.setFontStriken( true );
				}
			}
			else if( n.equalsIgnoreCase( "fg" ) )
			{    // font color
				// set the color
				int cl = FormatHandle.HexStringToColorInt( v, FormatHandle.colorFOREGROUND );
				cx.setFontColorIndex( cl );
			}
			else if( n.equalsIgnoreCase( "pattern" ) )
			{    //fill pattern
				int vv = Integer.parseInt( v );
				cx.setPatternFillStyle( vv );
			}
			else if( n.equalsIgnoreCase( "color" ) || n.equalsIgnoreCase( "patterncolor" ) )
			{    // fill (pattern or fg) color
				int cl = FormatHandle.HexStringToColorInt( v, FormatHandle.colorFONT );    // finds best match
				cx.setPatternFillColor( cl, v );
			}
			else if( n.equalsIgnoreCase( "background" ) )
			{        // fill bg color
				int cl = FormatHandle.HexStringToColorInt( v, FormatHandle.colorBACKGROUND );    // finds best match
				cx.setPatternFillColorBack( cl );
				// ALIGNMENT
			}
			else if( n.equalsIgnoreCase( "alignment-horizontal" ) )
			{
				int s = Integer.parseInt( v );
// TODO: set alignment
			}
			else if( n.equalsIgnoreCase( "alignment-vertical" ) )
			{
				int s = Integer.parseInt( v );
// TODO: set alignment
			}
			else if( n.equalsIgnoreCase( "numberformat" ) )
			{
				int s = Integer.parseInt( v );
// TODO: set number format   
				// FONT options
			}
			else if( n.equalsIgnoreCase( "font-height" ) )
			{
				int s = Integer.parseInt( v );
				cx.setFontHeight( s );
			}
			else if( n.equalsIgnoreCase( "font-weight" ) )
			{
				int s = Integer.parseInt( v );
				cx.setFontWeight( s );
			}
			else if( n.equalsIgnoreCase( "font-EscapementFlag" ) )
			{
				int s = Integer.parseInt( v );    // 0= none, 1= superscript, 2= subscript
				cx.setFontEscapement( s );
			}
			else if( n.equalsIgnoreCase( "font-striken" ) )
			{
				boolean b = Boolean.valueOf( v );
				cx.setFontStriken( b );
			}
			else if( n.equalsIgnoreCase( "font-italic" ) )
			{
				boolean b = Boolean.valueOf( v );
				cx.setFontItalic( b );
			}
			else if( n.equalsIgnoreCase( "font-ColorIndex" ) )
			{
				int s = Integer.parseInt( v );
				cx.setFontColorIndex( s );
			}
			else if( n.equalsIgnoreCase( "font-UnderlineStyle" ) )
			{
				int s = Integer.parseInt( v );
				cx.setFontUnderlineStyle( s );
			}
/*    			TODO: handle
 				fontName?
 * 				fontOptsFlag
    			fontModifiedOptionsFlag
    			fontUnderlineModified
*/
		}
		cx.updateRecord();
	}

	/**
	 * set the conditional format format settings from the OOXML dxf, or differential xf, format settings
	 *
	 * @param dxf
	 * @param cf
	 */
	public static void setStylePropsFromDxf( Dxf dxf, Cf cf )
	{
		int[] borderStyles = dxf.getBorderStyles();
		if( borderStyles != null )
		{    // then should have every other element
			int[] borderColors = dxf.getBorderColors();
			int[] borderSizes = dxf.getBorderSizes();
			cf.setBorderLineColorTop( borderColors[0] );
			cf.setBorderLineStylesTop( borderStyles[0] );
			cf.setBorderLineColorLeft( borderColors[1] );
			cf.setBorderLineStylesLeft( borderStyles[1] );
			cf.setBorderLineColorBottom( borderColors[2] );
			cf.setBorderLineStylesBottom( borderStyles[2] );
			cf.setBorderLineColorRight( borderColors[3] );
			cf.setBorderLineStylesRight( borderStyles[3] );
// TODO border size??
		}

		// FILL ****************
		int fls = dxf.getFillPatternInt();
		if( fls >= 0 )
		{
			cf.setPatternFillStyle( fls );
			int fg = dxf.getFg();
			int bg = dxf.getBg();
			if( fg != -1 )
			{
				cf.setPatternFillColor( fg, null );
			}
			if( bg != -1 )
			{
				cf.setPatternFillColorBack( bg );
			}
		}
		cf.fill = dxf.getFill();    // save OOXML var as color settings sometimes cannot be interpreted in 2003 v.  

		int z;
		// ALIGNMENT ******************
		String s = dxf.getHorizontalAlign();
		if( s != null )
		{
			z = Integer.parseInt( s );
//TODO: set alignment
		}
		s = dxf.getVerticalAlign();
		if( s != null )
		{
			z = Integer.parseInt( s );
//TODO: set alignment
		}

		// Number Format ***************
		s = dxf.getNumberFormat();
		if( s != null )
		{
			z = Integer.parseInt( s );
//TODO: set number format
		}

		// FONT ************************
		if( dxf.getFont() != null )
		{
			cf.parseFont( dxf.getFont() );
			cf.font = dxf.getFont();
		}
		
		
/*    			TODO: handle
				fontName?
* 				fontOptsFlag
			fontModifiedOptionsFlag
			fontUnderlineModified
*/
		cf.updateRecord();
	}

	/**
	 * sets the operator for this Cf Rule from a String
	 *
	 * @param s
	 */
	public void setOperator( String s )
	{
		this.cp = getConditionFromString( s );
		if( cp == 0x0 )
		{    // no comparison
			expression1 = ExpressionParser.parseExpression( s.getBytes(), this );
		}
		else
		{
			// val1 = cfx.getFormula1().calculateFormula();
		}
	}

	/**
	 * sets the Operator of this Cf Rule
	 * <br>possible values:
	 * <p>0= no comparison
	 * <p>1= between
	 * <p>2= not between
	 * <p>3= equal
	 * <p>4= not equal
	 * <p>5= greater than
	 * <p>6= less than
	 * <p>7= greater than or equal
	 * <p>8= less than or equal
	 * ********************
	 * 2007-specific operators
	 * <p>9= begins With
	 * <p>10= ends With
	 * <p>11= contains Text
	 * <p>12= not contains
	 *
	 * @param int qualifier - constant value as above
	 */
	public void setOperator( int op )
	{
		this.cp = (short) op;
	}

	/**
	 * returns the operator or qualifier of this Cf Rue as an int value
	 * <br>possible values:
	 * <p>0= no comparison
	 * <p>1= between
	 * <p>2= not between
	 * <p>3= equal
	 * <p>4= not equal
	 * <p>5= greater than
	 * <p>6= less than
	 * <p>7= greater than or equal
	 * <p>8= less than or equal
	 * ********************
	 * 2007-specific operators
	 * <p>9= begins With
	 * <p>10= ends With
	 * <p>11= contains Text
	 * <p>12= not contains
	 *
	 * @return
	 */
	public short getOperator()
	{
		return this.cp;

	}

	/**
	 * sets the type of this Cf rule from the int value
	 * <p>1= Cell is ("Cell value is")
	 * <p>2= expression ("formula value is")
	 *
	 * @param int type as above
	 */
	protected void setType( int type )
	{
		this.ct = (short) type;
	}

	/**
	 * return the type of this Cf Rule as an int
	 * <p>1= Cell is ("Cell value is")
	 * <p>2= expression ("formula value is")
	 *
	 * @return
	 */
	public short getType()
	{
		return this.ct;
	}

	/**
	 * reutrn the string representation of this Cf Rule type
	 *
	 * @return
	 */
	public String getTypeString()
	{
		return Cf.translateType( this.ct );

	}

	/**
	 * sets the first condition from a String
	 *
	 * @param s
	 */
	public void setCondition1( String s )
	{
		// takes "123" aka Value1 and converts to formula expression =123    	
		s = Cf.unescapeFormulaString( s );
		if( s.indexOf( "=" ) != 0 )
		{
			s = "=" + s;
			try
			{
				int[] r = { 0, 0 };
				formula1 = FormulaParser.getFormulaFromString( s, this.getSheet(), r );
				formula1.setWorkBook( this.getWorkBook() );
				expression1 = formula1.getExpression();
			}
			catch( Exception e )
			{
				;
			}
		}
	}

	/**
	 * sets the second condition from a String
	 *
	 * @param s
	 */
	public void setCondition2( String s )
	{
		if( s.equals( "" ) )
		{
			return; // none
		}

		s = Cf.unescapeFormulaString( s );
		// takes "123" aka Value1 and converts to formula expression =123
		if( cp == 0x0 )
		{    // no comparison
			expression2 = ExpressionParser.parseExpression( s.getBytes(), this );
		}
		else
		{
			if( s.indexOf( "=" ) != 0 )
			{
				s = "=" + s;
				try
				{
					int[] r = { 0, 0 };
					formula2 = FormulaParser.getFormulaFromString( s, this.getSheet(), r );
					expression2 = formula1.getExpression();
				}
				catch( Exception e )
				{
					;
				}
			}
		}
		this.cce2 = s.length();
	}

	/**
	 * @return Returns the condfmt.
	 */
	public Condfmt getCondfmt()
	{
		return condfmt;
	}

	/**
	 * @param condfmt The condfmt to set.
	 */
	public void setCondfmt( Condfmt condfmt )
	{
		this.condfmt = condfmt;
	}

	/**
	 * @return Returns the fontEscapementFlag.
	 */
	public int getFontEscapement()
	{
		return fontEscapementFlag;
	}

	/**
	 * returns the pattern color, if any, as an HTML color String.  Includes custom OOXML colors.
	 *
	 * @return String HTML Color String
	 */
	public String getPatternFgColor()
	{
		if( fill != null )
		{
			return fill.getFgColorAsRGB( getWorkBook().getTheme() );
		}
		if( patternFillColor != -1 )
		{
			return FormatHandle.colorToHexString( this.getColorTable()[patternFillColor] );
		}
		return null;
	}

	/**
	 * returns the pattern background color, if any, as an HTML color String.  Includes custom OOXML colors.
	 *
	 * @return String HTML Color String
	 */
	public String getPatternBgColor()
	{
		if( fill != null )
		{
			if( patternFillStyle == 1 )
			{
				return fill.getFgColorAsRGB( getWorkBook().getTheme() );
			}
			return fill.getBgColorAsRGB( getWorkBook().getTheme() );
		}
		if( patternFillColorBack != -1 )
		{
			return FormatHandle.colorToHexString( this.getColorTable()[patternFillColorBack] );
		}
		return null;
	}

	/**
	 * parse cf font into it's member elements
	 */
	private void parseFont( Font font )
	{
		int z = font.getFontHeight();
		if( z > 0 )
		{
			setFontHeight( z );
		}
		z = font.getFontWeight();
		if( z > 0 )
		{
			setFontWeight( z );
		}
//		} else if (n.equalsIgnoreCase("font-EscapementFlag")) {
//			cx.setFontEscapement(s);    			
		setFontStriken( font.getStricken() );
		setFontItalic( font.getItalic() );
		z = font.getColor();
		if( z > -1 )
		{
			setFontColorIndex( z );
		}
		z = font.getUnderlineStyle();
		if( z > -1 )
		{
			setFontUnderlineStyle( z );
		}
	}

	/**
	 * @param fontEscapementFlag The fontEscapementFlag to set.
	 */
	public void setFontEscapement( int fontEscapementFlag )
	{
		this.fontEscapementFlag = fontEscapementFlag;
		this.fontEscapementFlagModifiedFlag = 0;    // set modified
		bHasFontBlock = true;
//		if (font!=null)
//			font.setScript(ss);
	}

	// Start Font Formatting Block

	/**
	 * updates the values for the current border colors from the flag
	 * -----------------------------------------------------------
	 */
	void updateBorderLineStyles()
	{
		if( (flags & BORDER_MODIFIED_LEFT) == 0 )
		{
			this.borderLineStylesLeft = (short) (this.borderLineStylesFlag & BORDER_LINESTYLE_LEFT);
		}
		if( (flags & BORDER_MODIFIED_RIGHT) == 0 )
		{
			this.borderLineStylesRight = (short) ((this.borderLineStylesFlag & BORDER_LINESTYLE_RIGHT) >> 4);
		}
		if( (flags & BORDER_MODIFIED_TOP) == 0 )
		{
			this.borderLineStylesTop = (short) ((this.borderLineStylesFlag & BORDER_LINESTYLE_TOP) >> 8);
		}
		if( (flags & BORDER_MODIFIED_BOTTOM) == 0 )
		{
			this.borderLineStylesBottom = (short) ((this.borderLineStylesFlag & BORDER_LINESTYLE_BOTTOM) >> 12);
		}
		bHasBorderBlock = true;
	}

	/**
	 * Apply all the borderLineStyle fields into the current border line styles flag
	 */
	public void updateBorderLineStylesFlag()
	{
		flags = flags | 0x10000000;        // set flags to denote has a border block
		bHasBorderBlock = ((flags & 0x10000000) == 0x10000000);

		borderLineStylesFlag = 0;
		if( borderLineStylesLeft >= 0 )
		{
			borderLineStylesFlag = (short) borderLineStylesLeft;
		}
		if( borderLineStylesRight >= 0 )
		{
			borderLineStylesFlag = (short) ((borderLineStylesFlag | ((borderLineStylesRight) << 4)));
		}
		if( borderLineStylesTop >= 0 )
		{
			borderLineStylesFlag = (short) ((borderLineStylesFlag | ((borderLineStylesTop) << 8)));
		}
		if( borderLineStylesBottom >= 0 )
		{
			borderLineStylesFlag = (short) ((borderLineStylesFlag | ((borderLineStylesBottom) << 12)));
		}
/*
        byte[] data = this.getData();
        byte[] updated = ByteTools.shortToLEBytes(borderLineStylesFlag);
        int pos= 12;
        if (bHasFontBlock)
        	pos+=118;	
        if (data.length< (pos+2)) {
        	byte[] tmp= new byte[pos+2];
        	System.arraycopy(data, 0, tmp, 0, data.length);
        	this.setData(tmp);
        }
        data[pos++] = updated[0];
        data[pos++] = updated[1];
*/
	}

	//6 2 Not used

	/**
	 * updates the values for the current border colors from the flag
	 */
	void updateBorderLineColors()
	{
		if( (flags & BORDER_MODIFIED_LEFT) == 0 )
		{
			this.borderLineColorLeft = (short) (this.borderLineColorsFlag & BORDER_LINECOLOR_LEFT);
		}
		if( (flags & BORDER_MODIFIED_RIGHT) == 0 )
		{
			this.borderLineColorRight = (short) ((this.borderLineColorsFlag & BORDER_LINECOLOR_RIGHT) >> 7);
		}
		if( (flags & BORDER_MODIFIED_TOP) == 0 )
		{
			this.borderLineColorTop = (short) ((this.borderLineColorsFlag & BORDER_LINECOLOR_TOP) >> 16);
		}
		if( (flags & BORDER_MODIFIED_BOTTOM) == 0 )
		{
			this.borderLineColorBottom = (short) ((this.borderLineColorsFlag & BORDER_LINECOLOR_BOTTOM) >> 23);
		}
	}

	/**
	 * Apply all the borderLineColor fields into the current border line colors flag
	 */
	public void updateBorderLineColorsFlag()
	{
		flags = flags | 0x10000000;        // set flags to denote has a border block
		bHasBorderBlock = ((flags & 0x10000000) == 0x10000000);    // = true

		borderLineColorsFlag = 0;
		borderLineColorsFlag = (borderLineColorsFlag | borderLineColorLeft);
		borderLineColorsFlag = ((borderLineColorsFlag | ((borderLineColorRight) << 7)));
		borderLineColorsFlag = ((borderLineColorsFlag | ((borderLineColorTop) << 16)));
		borderLineColorsFlag = ((borderLineColorsFlag | ((borderLineColorBottom) << 23)));
/*        
        byte[] data = this.getData();
        byte[] updated = ByteTools.cLongToLEBytes( borderLineColorsFlag);
        int pos= 12;
        if (bHasFontBlock)
        	pos+=118;
        pos+=2;	// skip border line style
        if (data.length<(pos+4)) {
        	byte[] tmp= new byte[pos+4];
        	System.arraycopy(data, 0, tmp, 0, data.length);
        	this.setData(tmp);
        }
        data[pos++] = updated[0];
        data[pos++] = updated[1];
        data[pos++] = updated[2];
        data[pos++] = updated[3];
        // no need updateBorderLineColors();
*/
	}

	/**
	 * updates the values for the current pattern fill colors from the flag
	 * <p/>
	 * // Bit Mask Contents
	 * // 15-10 FC00H Fill pattern style (only if patt - style = 0, ?3.11)
	 * // PATTERN_FILL_STYLE = 0xFC00;
	 */
	void updatePatternFillColors()
	{
/* 		below appears correct in testing		this.patternFillColor = (short)(this.patternFillColorsFlag & PATTERN_FILL_COLOR);	
		this.patternFillColorBack = (short)((this.patternFillColorsFlag & PATTERN_FILL_BACK_COLOR) >> 7);
*/
		this.patternFillColor = (short) ((this.patternFillColorsFlag & PATTERN_FILL_BACK_COLOR) >> 7);
		this.patternFillColorBack = (short) (this.patternFillColorsFlag & PATTERN_FILL_COLOR);
		bHasPatternBlock = true;
	}

	/**
	 * return the 2007v Fill element, or null if not set
	 */
	public Fill getFill()
	{
		return fill;
	}

	/**
	 * Apply all the patternFillColor fields into the current pattern colors flag
	 */
	public void updatePatternFillColorsFlag()
	{
		flags = flags | 0x20000000;        // set flags to denote has a pattern block
		bHasPatternBlock = true;
		patternFillColorsFlag = (short) patternFillColorBack;
		patternFillColorsFlag = (short) (patternFillColorsFlag | (patternFillColor << 7));
	}

	@Override
	public Font getFont()
	{
		if( !this.bHasFontBlock )
		{
			return null;
		}

		if( font != null )
		{
			return font;
		}
		int t = this.fontHeight;
		int x = this.fontWeight;
		if( t == -1 )
		{
			t = 180;
		}
		else
		{
			t *= 20;
		}
		if( x == -1 )
		{
			x = Font.PLAIN;
		}
		font = new Font( "Arial", x, t );
		if( fontColorIndex > -1 )
		{
			font.setColor( fontColorIndex );
		}
		return font;
	}

	public java.awt.Color[] getBorderColors()
	{
		if( !this.bHasBorderBlock )
		{
			return null;
		}
		java.awt.Color[] test = {
				this.getColorTable()[this.getBorderLineColorTop()],
				this.getColorTable()[this.getBorderLineColorLeft()],
				this.getColorTable()[this.getBorderLineColorBottom()],
				this.getColorTable()[this.getBorderLineColorRight()]
		};
		return test;
	}

	public int[] getAllBorderColors()
	{
		if( !this.bHasBorderBlock )
		{
			return null;
		}
		int[] test = {
				this.getBorderLineColorTop(), this.getBorderLineColorLeft(), this.getBorderLineColorBottom(), this.getBorderLineColorRight()
		};
		return test;
	}

	/**
	 * order= top, left, bottom, right
	 *
	 * @return
	 */
	public int[] getBorderStyles()
	{
		if( !this.bHasBorderBlock )
		{
			return null;
		}
		int[] test = {
				this.getBorderLineStylesTop(),
				this.getBorderLineStylesLeft(),
				this.getBorderLineStylesBottom(),
				this.getBorderLineStylesRight(),
				-1
		};    // diag
		return test;
	}

	public int[] getBorderSizes()
	{
		if( !this.bHasBorderBlock )
		{
			return null;
		}
		boolean hasTop = this.getBorderLineStylesTop() > 0;
		boolean hasLeft = this.getBorderLineStylesLeft() > 0;
		boolean hasBottom = this.getBorderLineStylesBottom() > 0;
		boolean hasRight = this.getBorderLineStylesRight() > 0;
		int[] test = { (hasTop ? 1 : 0), (hasLeft ? 1 : 0), (hasBottom ? 1 : 0), (hasRight ? 1 : 0), 0 };
		return test;
	}

	public int getForegroundColor()
	{
		if( !this.bHasPatternBlock )
		{
			return -1;
		}
		if( fill != null )
		{
			return fill.getFgColorAsInt( getWorkBook().getTheme() );
		}
		if( this.patternFillStyle == 1 )
		{
			return this.patternFillColorBack;
		}
		return this.patternFillColor;
	}

	/**
	 * @return Returns the fontOptsPosture.
	 */
	public int getFontOptsPosture()
	{
		return (short) (this.fontOptsFlag & FONT_OPTIONS_POSTURE);
	}

	/**
	 * @param fontOptsPosture The fontOptsPosture to set.
	 */
	public void setFontOptsPosture( int fontOptsPosture )
	{
		this.fontOptsFlag = (short) (this.fontOptsFlag & FONT_OPTIONS_POSTURE);
		bHasFontBlock = true;
	}

	/**
	 * @return Returns the fontOptsCancellation.
	 */
	public int getFontOptsCancellation()
	{
		return (short) ((this.fontOptsFlag & FONT_OPTIONS_CANCELLATION) >> 7);
	}

	/**
	 * @return Returns the fontOptsItalic.
	 */
	public boolean getFontItalic()
	{
		return (fontOptsFlag & 0x2) == FONT_OPTIONS_POSTURE_ITALIC;    // 1 if italic 
	}

	/**
	 * @param fontOptsItalic The fontOptsItalic to set.
	 */
	public void setFontItalic( boolean italic )
	{
		if( italic )
		{
			this.fontOptsFlag = this.fontOptsFlag | 0x2;    // set italic
			this.fontModifiedOptionsFlag = this.fontModifiedOptionsFlag | 0x2;
			bHasFontBlock = true;
		}
		else
		{    // todo: is below correct?
			this.fontOptsFlag = this.fontOptsFlag ^ 0x2;    // clear italic bit
			this.fontModifiedOptionsFlag = this.fontModifiedOptionsFlag ^ 0x2;
		}
		if( font != null )
		{
			font.setItalic( italic );
		}
	}

	public boolean getFontStriken()
	{
		return ((fontOptsFlag & 0x80) == FONT_OPTIONS_CANCELLATION_ON);
	}

	public void setFontStriken( boolean bStriken )
	{
		if( bStriken )
		{
			fontOptsFlag = (fontOptsFlag | 0x80);
			fontModifiedOptionsFlag = (fontModifiedOptionsFlag | 0x80);
			bHasFontBlock = true;
		}
		else
		{    // todo: is below correct?
			fontOptsFlag = (fontOptsFlag ^ 0x80);    // turn off
			fontModifiedOptionsFlag = (fontModifiedOptionsFlag ^ 0x80);
		}
		if( font != null )
		{
			font.setStricken( bStriken );
		}
	}

	/**
	 * @return Returns true if font escapement is superscript
	 */
	public boolean getFontEscapementSuper()
	{
		return (fontEscapementFlag == FONT_ESCAPEMENT_SUPER);
	}

	/**
	 * sets the font escapement for this conditional format to superscript
	 */
	public void setFontEscapementSuper()
	{
		this.fontEscapementFlag = FONT_ESCAPEMENT_SUPER;
		this.fontEscapementFlagModifiedFlag = 0;    // 
		this.bHasFontBlock = true;
	}

	/**
	 * @return Returns true if font escapement is subscript
	 */
	public boolean getFontEscapementSub()
	{
		return (fontEscapementFlag == FONT_ESCAPEMENT_SUB);
	}

	/**
	 * sets the font escapement for this conditional format to subscript
	 */
	public void setFontEscapementSub()
	{
		this.fontEscapementFlag = FONT_ESCAPEMENT_SUB;
		this.fontEscapementFlagModifiedFlag = 0;    // 
		this.bHasFontBlock = true;
	}

	/**
	 * @return Returns the borderLineStylesLeft.
	 */
	public int getBorderLineStylesLeft()
	{
		return borderLineStylesLeft;
	}

	/**
	 * @param borderLineStylesLeft The borderLineStylesLeft to set.
	 */
	public void setBorderLineStylesLeft( int b )
	{
		this.borderLineStylesLeft = b;
		this.updateBorderLineStylesFlag();
	}

	/**
	 * @return Returns the borderLineStylesRight.
	 */
	public int getBorderLineStylesRight()
	{
		return borderLineStylesRight;
	}

	/**
	 * @param borderLineStylesRight The borderLineStylesRight to set.
	 */
	public void setBorderLineStylesRight( int b )
	{
		this.borderLineStylesRight = b;
		this.updateBorderLineStylesFlag();
	}

	/**
	 * @return Returns the borderLineStylesTop.
	 */
	public int getBorderLineStylesTop()
	{
		return borderLineStylesTop;
	}

	/**
	 * @param borderLineStylesTop The borderLineStylesTop to set.
	 */
	public void setBorderLineStylesTop( int b )
	{
		this.borderLineStylesTop = b;
		this.updateBorderLineStylesFlag();
	}

	/**
	 * @return Returns the borderLineStylesBottom.
	 */
	public int getBorderLineStylesBottom()
	{
		return borderLineStylesBottom;
	}

	/**
	 * @param borderLineStylesBottom The borderLineStylesBottom to set.
	 */
	public void setBorderLineStylesBottom( int b )
	{
		this.borderLineStylesBottom = b;
		this.updateBorderLineStylesFlag();
	}

	/**
	 * @return Returns the borderLineColorLeft.
	 */
	public int getBorderLineColorLeft()
	{
		if( borderLineColorLeft > this.getColorTable().length )
		{
			return 0;
		}
		if( borderLineColorLeft < 0 )
		{
			return 0;
		}
		return borderLineColorLeft;
	}

	/**
	 * @param borderLineColorLeft The borderLineColorLeft to set.
	 */
	public void setBorderLineColorLeft( int borderLineColorLeft )
	{
		this.borderLineColorLeft = borderLineColorLeft;
		// 20091028 KSC: insure flag denotes borderlinecolor is modified
		flags = flags & (BORDER_MODIFIED_LEFT - 1);    // set flags to denote border top is modified (set bit=0)
		this.updateBorderLineColorsFlag();
	}

	/**
	 * @return Returns the borderLineColorRight.
	 */
	public int getBorderLineColorRight()
	{
		if( borderLineColorRight > this.getColorTable().length )
		{
			return 0;
		}
		if( borderLineColorRight < 0 )
		{
			return 0;
		}
		return borderLineColorRight;
	}

	/**
	 * @param borderLineColorRight The borderLineColorRight to set.
	 */
	public void setBorderLineColorRight( int borderLineColorRight )
	{
		this.borderLineColorRight = borderLineColorRight;
		// 20091028 KSC: insure flag denotes borderlinecolor is modified
		flags = flags & (BORDER_MODIFIED_RIGHT - 1);    // set flags to denote border top is modified (set bit=0)
		this.updateBorderLineColorsFlag();
	}

	/**
	 * @return Returns the borderLineColorTop.
	 */
	public int getBorderLineColorTop()
	{
		if( borderLineColorTop > this.getColorTable().length )
		{
			return 0;
		}
		if( borderLineColorTop < 0 )
		{
			return 0;
		}
		return borderLineColorTop;
	}

	/**
	 * @param borderLineColorTop The borderLineColorTop to set.
	 */
	public void setBorderLineColorTop( int b )
	{
		this.borderLineColorTop = b;
		// 20091028 KSC: insure flag denotes borderlinecolor is modified
		flags = flags & (BORDER_MODIFIED_TOP - 1);    // set flags to denote border top is modified (set bit=0)
		this.updateBorderLineColorsFlag();
		if( this.borderLineColorTop != b )
		{
			log.warn( "setBorderLineColorTop failed" );
		}
	}

	/**
	 * @return Returns the borderLineColorBottom.
	 */
	public int getBorderLineColorBottom()
	{
		if( borderLineColorBottom > this.getColorTable().length )
		{
			return 0;
		}
		if( borderLineColorBottom < 0 )
		{
			return 0;
		}
		return borderLineColorBottom;
	}

	/**
	 * @param borderLineColorBottom The borderLineColorBottom to set.
	 */
	public void setBorderLineColorBottom( int b )
	{
		this.borderLineColorBottom = b;
		// 20091028 KSC: insure flag denotes borderlinecolor is modified
		flags = flags & (BORDER_MODIFIED_BOTTOM - 1);    // set flags to denote border top is modified (set bit=0)
		this.updateBorderLineColorsFlag();
		if( this.borderLineColorBottom != b )
		{
			log.warn( "borderLineColorBottom failed" );
		}
	}

	/**
	 * Returns the patternFillStyle.
	 * <br>NOTE in 2003-ver patternFillStyle is valid 1->
	 *
	 * @return Returns the patternFillStyle.
	 */
	public int getPatternFillStyle()
	{
		return patternFillStyle;
	}

	/**
	 * @param patternFillStyle The patternFillStyle to set.
	 */
	public void setPatternFillStyle( int p )
	{
		this.patternFillStyle = p;
		bHasPatternBlock = true;
	}

	/**
	 * @return Returns the patternFillColor.
	 */
	public int getPatternFillColor()
	{
		if( fill != null )
		{
			return fill.getFgColorAsInt( getWorkBook().getTheme() );
		}
		return patternFillColor;
	}

	/**
	 * @param patternFillColor The patternFillColor to set.
	 */
	public void setPatternFillColor( int p, String custom )
	{
		this.patternFillColor = p;
		if( fill != null )
		{
			fill.setFgColor( p );
		}
		this.updatePatternFillColorsFlag();
	}

	/**
	 * @return Returns the patternFillColorBack.
	 */
	public int getPatternFillColorBack()
	{
		if( fill != null )
		{
			return fill.getBgColorAsInt( getWorkBook().getTheme() );
		}
		return patternFillColorBack;
	}

	/**
	 * @param patternFillColorBack The patternFillColorBack to set.
	 */
	public void setPatternFillColorBack( int p )
	{
		this.patternFillColorBack = p;
		if( fill != null )
		{
			fill.setBgColor( p );
		}
		this.updatePatternFillColorsFlag();
	}

	/**
	 * @return Returns the fontHeight.
	 */
	public int getFontHeight()
	{
		return fontHeight;
	}

	/**
	 * @param fontHeight The fontHeight to set.
	 */
	public void setFontHeight( int fontHeight )
	{
		this.fontHeight = fontHeight;
		bHasFontBlock = true;
		if( font != null )
		{
			font.setFontHeight( fontHeight );
		}
	}

	/**
	 * @return Returns the fontWeight.
	 */
	public int getFontWeight()
	{
		return fontWeight;
	}

	/**
	 * @param fontWeight The fontWeight to set.
	 */
	public void setFontWeight( int f )
	{
		this.fontWeight = f;
		this.fontOptsFlag = this.fontOptsFlag & 0xFD;    // turn off bit 1 = style bit	
		bHasFontBlock = true;
		if( font != null )
		{
			font.setFontWeight( f );
		}
	}

	/**
	 * @return Returns the fontUnderlineStyle.
	 */
	public int getFontUnderlineStyle()
	{
		return fontUnderlineStyle;
	}

	/**
	 * @param fontUnderlineStyle The fontUnderlineStyle to set.
	 */
	public void setFontUnderlineStyle( int fontUnderlineStyle )
	{
		this.fontUnderlineStyle = fontUnderlineStyle;
		this.fontUnderlineModifiedFlag = 0;    // set modified flag
		bHasFontBlock = true;
		if( font != null )
		{
			font.setUnderlineStyle( (byte) fontUnderlineStyle );
		}
	}

	/**
	 * @return Returns the fontColorIndex.
	 */
	public int getFontColorIndex()
	{
		return fontColorIndex;
	}

	/**
	 * @param fontColorIndex The fontColorIndex to set.
	 */
	public void setFontColorIndex( int fontColorIndex )
	{
		this.fontColorIndex = fontColorIndex;
		bHasFontBlock = true;
		if( font != null )
		{
			font.setColor( fontColorIndex );
		}
	}

	/**
	 * Create a Cf record & populate with prototype bytes
	 *
	 * @return TODO: NOT FINISHED
	 */
	protected static XLSRecord getPrototype()
	{
		Cf cf = new Cf();
		cf.setOpcode( CF );
		cf.setData( cf.PROTOTYPE_BYTES );
		cf.init();
		return cf;
	}

	private int refPos = -1;

	/**
	 * Reset the ptgRef to A1 in these that is replaced with
	 * current Ptg
	 */
	private void resetFormulaRef()
	{
		// TODO: test what happens when A1 is a valid part of the expression
		Stack expr = this.getFormula1().getExpression();
		Iterator itx = expr.iterator();
		if( refPos > -1 )
		{
			expr.insertElementAt( prefHolder, refPos );
		}
	}

	/**
	 * There is a ptgRef to A1 in these that is replaced with
	 * current Ptg
	 */
	private void setFormulaRef( Ptg refcell ) throws FormulaNotFoundException
	{
		// TODO: test what happens when A1 is a valid part of the expression
		Stack expr = this.getFormula1().getExpression();
		Iterator itx = expr.iterator();

		int[] rc = refcell.getIntLocation();
		if( refPos == -1 )
		{
			while( itx.hasNext() )
			{
				Ptg prex = (Ptg) itx.next();
				if( prex instanceof PtgRefN )
				{
					((PtgRefN) prex).setFormulaRow( rc[0] );
					((PtgRefN) prex).setFormulaCol( rc[1] );
				}
			}
		}
		else
		{
			expr.remove( refPos );
		}
		if( refPos > -1 )
		{
			expr.remove( prefHolder );
			expr.insertElementAt( refcell, refPos );
		}

	}

	/**
	 * pass in the referenced cell and attempt to
	 * create a valid formula for this thing and
	 * calculate whether the criteria passes
	 * <p/>
	 * returns true or false
	 *
	 * @param the reference to evaluate
	 * @return boolean passes
	 */
	public boolean evaluate( Ptg refcell )
	{
		try
		{
			Object val2 = null;
			Object val1 = null;

			if( this.cp != 0x0 ) // calcs later
			{
				val1 = this.getFormula1().calculateFormula();
			}

			if( this.cce2 > 0 )
			{
				val2 = this.getFormula2().calculateFormula();
			}

			Object valX = refcell.getValue();

			// cast to double then compare
			double d1 = 0.0d;
			double d2 = 0.0d; // second val from expression2

			double dX = 0.0d; // the reference val

			try
			{
				d1 = new Double( val1.toString() );
				dX = new Double( valX.toString() );
				if( this.cce2 > 0 )
				{ // we have a second value
					d2 = new Double( val2.toString() );
				}

			}
			catch( Exception e )
			{
				; // not numeric
			}

			// handle evaluated condition
			switch( this.cp )
			{
				case 0x0:    // No comparison (only valid for formula type, see above)
					setFormulaRef( refcell );
					val1 = this.getFormula1().calculateFormula();
					return (Boolean) val1;

				case 01:    // Between
					// expression2 for the other bounds ... 
					if( (dX >= d1) && (dX <= d2) )
					{
						return true;
					}
					return false;

				case 0x5:    // Greater than
					return dX > d1;

				case 0x2:    // Not between
					if( dX < d1 )
					{
						return false;
					}
					if( dX > d1 )  // hmmm... d2 is where? an array?
					{
						return false;
					}

				case 0x6:    // Less than
					return dX < d1;

				case 0x3:    // Equal
					return dX == d1;

				case 0x7:    // Greater or equal
					return dX >= d1;

				case 0x4:    // Not equal
					return dX != d1;

				case 0x8:    // Less or equal
					return dX <= d1;

				// 2007-specific operators
				case 0x9:    // begins With
				case 0xA:    // ends With
				case 0xB:    // contains text
				case 0xC:    // not contains
					return false;

				default:
					return false;
			}
		}
		catch( Exception ex )
		{
			// log.warn("CF condition "+this.formula1.getFormulaString()+" evaluation failed for : " + refcell.toString());
			return false;
		}
	}

	/**
	 * Returns the byte Condition type from a human-readable string
	 *
	 * @return the condition type for this rule
	 */
	public static final byte getConditionFromString( String cx )
	{
		cx = StringTool.allTrim( cx ); // trim
		cx = cx.toUpperCase();
		//    	 handle evaluated condition
		if( cx.equals( "BETWEEN" ) )
		{
			return 0x1;    // Between
		}
		if( cx.equals( "GREATER THAN" ) )
		{
			return 0x5;    // Greater than
		}
		if( cx.equals( "NOT BETWEEN" ) )
		{
			return 0x2;    // Not between
		}
		if( cx.equals( "LESS THAN" ) )
		{
			return 0x6;    // Less than
		}
		if( cx.equals( "EQUALS" ) )
		{
			return 0x3;    // Equal
		}
		if( cx.equals( "GREATER THAN OR EQUAL" ) )
		{
			return 0x7;    // Greater or equal
		}
		if( cx.equals( "NOT EQUAL" ) )
		{
			return 0x4;    // Not equal
		}
		if( cx.equals( "LESS THAN OR EQUAL" ) )
		{
			return 0x8;    // Less or equal
		}
		return 0x0;    // No comparison (only valid for formula type, see above)
	}

	/**
	 * Returns the human-readable Condition type
	 *
	 * @return the condition type for this rule
	 */
	public String getConditionString()
	{
//    	 handle evaluated condition
		switch( this.cp )
		{
			case 0x0:    // No comparison (only valid for formula type, see above)
				// okay annoying, but apparenlty there is a ptgRef to A1 in these that should
				// be replaced with our ptg... whatever!!
				return this.expression1.toString() + this.expression2.toString();

			case 01:    // Between
				// expression2 for the other bounds ... 
				return "Between";

			case 0x5:    // Greater than
				return "Greater Than";

			case 0x2:    // Not between
				return "Not Between";

			case 0x6:    // Less than
				return "Less Than";

			case 0x3:    // Equal
				return "Equals";

			case 0x7:    // Greater or equal
				return "Greater Than or Equal";

			case 0x4:    // Not equal
				return "Not Equal";

			case 0x8:    // Less or equal
				return "Less Than or Equal";

			// 2007-Specific
			case 0x9:
				return "Begins With";
			case 0xA:
				return "Ends With";
			case 0xB:
				return "Contains Text";
			case 0xC:
				return "Not Contains";
			default:
				return "Unknown";
		}
	}

	/**
	 * return the first formula referenced by the Conditional Format
	 *
	 * @return Formula
	 */
	public Formula getFormula1()
	{
		if( formula1 == null )
		{ // hasn't been set
			formula1 = new Formula();
			formula1.setWorkBook( this.getWorkBook() );
			if( this.getSheet() == null )
			{
				this.setSheet( this.condfmt.getSheet() ); // help!
			}
			formula1.setSheet( this.getSheet() );
			formula1.setExpression( expression1 );
		}
// 20101216 KSC: WHY????    	formula1.setCachedValue(null);	
		return formula1;
	}

	/**
	 * return the second formula referenced by the Conditional Format
	 *
	 * @return Formula
	 */
	public Formula getFormula2()
	{
		if( (formula2 == null) && (cce2 > 0) )
		{ // hasn't been set
			formula2 = new Formula();
			formula2.setWorkBook( this.getWorkBook() );
			if( this.getSheet() == null )
			{
				this.setSheet( this.condfmt.getSheet() ); // help!
			}
			formula2.setSheet( this.getSheet() );
			formula2.setExpression( expression2 );
		}
		if( formula2 != null )
		{
			formula2.setSheet( this.getSheet() );
// 20101216 KSC: WHY???			formula2.setCachedValue(null);
		}
		return formula2;
	}

	/**
	 * restore Formula strings from XML serialization
	 *
	 * @param fmx
	 * @return
	 */
	private static String unescapeFormulaString( String fmx )
	{
		fmx = StringTool.replaceText( fmx, "&quot;", "\"" );
		// fmx = StringTool.replaceText(fmx,"&amp;","%");
		fmx = StringTool.replaceText( fmx, "&lt;", "<" );
		fmx = StringTool.replaceText( fmx, "&gt;", ">" );
		return fmx;
	}

	/**
	 * taks a String representing the Operator for this Cf Rule
	 * and translates it to an int
	 * <br>NOTE: 2003 versions do not use types
	 * 0x9, 0xA, 0xB or 0xC
	 *
	 * @param String operator - String CfRule operator attribute
	 * @return int representing Cf operator value
	 */
	protected static int translateOperator( String operator )
	{
		if( operator == null )    // type is not cellIs
		{
			return 0;
		}
		if( operator.equals( "between" ) )
		{
			return 0x1;
		}
		if( operator.equals( "greaterThan" ) )
		{
			return 0x5;
		}
		if( operator.equals( "notBetween" ) )
		{
			return 0x2;
		}
		if( operator.equals( "lessThan" ) )
		{
			return 0x6;
		}
		if( operator.equals( "equal" ) )
		{
			return 0x3;
		}
		//noinspection ConfusingElseBranch
		if( operator.equals( "greaterThanOrEqual" ) )
		{
			return 0x7;
		}
		//noinspection ConfusingElseBranch
		if( operator.equals( "notEqual" ) )
	{
		return 0x4;
	}
	else //noinspection ConfusingElseBranch
			if( operator.equals( "lessThanOrEqual" ) )
	{
		return 0x8;
	}
	// NO EQUIVALENT IN 2003:  	beginsWith, containsText, endsWith, notContains
	else //noinspection ConfusingElseBranch
			if( operator.equals( "beginsWith" ) )
	{
		return 0x9;
	}
	else //noinspection ConfusingElseBranch
			if( operator.equals( "endsWith" ) )
	{
		return 0xA;
	}
	else //noinspection ConfusingElseBranch
				if( operator.equals( "containsText" ) )
	{
		return 0xB;
	}
	else if( operator.equals( "notContains" ) )
	{
		return 0xC;
	}
		return 0;
	}

	/**
	 * Given an int type, return it's string representation
	 * <br>NOTE: if type is between 3 and 18,
	 * it is an OOXML-specific type.
	 *
	 * @param type type integer
	 * @return String reprentation
	 * @see translateType(String)
	 */
	protected static String translateType( int type )
	{
		switch( type )
		{
			case 1:
				return "Cell Is";
			case 2:
				return "expression";
			case 3:
				return "containsText";
			case 4:
				return "aboveAverage";
			case 5:
				return "beginsWith";
			case 6:
				return "colorScale";
			case 7:
				return "containsBlanks";
			case 8:
				return "containsErrors";
			case 9:
				return "dataBar";
			case 10:
				return "duplicateValues";
			case 11:
				return "endsWith";
			case 12:
				return "iconSet";
			case 13:
				return "notContainsBlanks";
			case 14:
				return "notContainsErrors";
			case 15:
				return "notContainsText";
			case 16:
				return "timePeriod";
			case 17:
				return "top10";
			case 18:
				return "uniqueValues";
			default:
				return "Unknnown";
		}
	}

	/**
	 * takes a String representing the type attribute and translates it to
	 * the corresponding integer representation
	 * <br>IMPORTANT NOTE: OOXML-specific types are converted to an integer that is not valid in 2003 versions
	 *
	 * @param String type - OOXML CfRule type attribute
	 * @return int representing Cf type value
	 */
	protected static int translateOOXMLType( String type )
	{
		if( type.equals( "cellIs" ) )
		{
			return 1;
		}
		if( type.equals( "expression" ) )
		{
			return 2;
		}
		// no equivalent in 2003: but must track for 2007 uses
		if( type.equals( "containsText" ) )
		{
			return 3;
		}
		if( type.equals( "aboveAverage" ) )
		{
			return 4;
		}
		if( type.equals( "beginsWith" ) )
		{
			return 5;
		}
		//noinspection ConfusingElseBranch
		if( type.equals( "colorScale" ) )
		{
			return 6;
		}
		//noinspection ConfusingElseBranch
		if( type.equals( "containsBlanks" ) )
	{
		return 7;
	}
	else //noinspection ConfusingElseBranch
			if( type.equals( "containsErrors" ) )
	{
		return 8;
	}
	else //noinspection ConfusingElseBranch
				if( type.equals( "dataBar" ) )
	{
		return 9;
	}
	else //noinspection ConfusingElseBranch
					if( type.equals( "duplicateValues" ) )
	{
		return 10;
	}
	else //noinspection ConfusingElseBranch
					if( type.equals( "endsWith" ) )
	{
		return 11;
	}
	else //noinspection ConfusingElseBranch
		if( type.equals( "iconSet" ) )
	{
		return 12;
	}
	else //noinspection ConfusingElseBranch
			if( type.equals( "notContainsBlanks" ) )
	{
		return 13;
	}
	else //noinspection ConfusingElseBranch
				if( type.equals( "notContainsErrors" ) )
	{
		return 14;
	}
	else //noinspection ConfusingElseBranch
				if( type.equals( "notContainsText" ) )
	{
		return 15;
	}
	else //noinspection ConfusingElseBranch
				if( type.equals( "timePeriod" ) )
	{
		return 16;
	}
	else //noinspection ConfusingElseBranch
			if( type.equals( "top10" ) )
	{
		return 17;
	}
	else if( type.equals( "uniqueValues" ) )
	{
		return 18;
	}
		return 1;    // default to cellIs ????
	}

	/**
	 * prepare Formula strings for XML serialization
	 *
	 * @param fmx
	 * @return
	 */
	private static String escapeFormulaString( String fmx )
	{
		fmx = StringTool.replaceText( fmx, "\"", "&quot;" );
		if( fmx.indexOf( "=" ) == 0 )
		{
			fmx = fmx.substring( 1 );
		}
		// fmx = StringTool.replaceText(fmx,"%","&amp;");
		fmx = StringTool.replaceText( fmx, "<", "&lt;" );
		fmx = StringTool.replaceText( fmx, ">", "&gt;" );
		return fmx;

	}

	/**
	 * returns EXML (XMLSS) for the Conditional Format Rule
	 * <p/>
	 * <p/>
	 * <p/>
	 * <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel">
	 * <Range>R12C2:R16C2</Range>
	 * <Condition>
	 * <Qualifier>Between</Qualifier>
	 * <Value1>2</Value1>
	 * <Value2>4</Value2>
	 * <Format Style='color:#002060;font-weight:700;text-line-through:none;
	 * border:.5pt solid windowtext;background:#00B0F0'/>
	 * </Condition>
	 * </ConditionalFormatting>
	 * <p/>
	 * <ConditionalFormatting xmlns="urn:schemas-microsoft-com:office:excel">
	 * <Range>R6C2</Range>
	 * <Condition>
	 * <Value1>NOT(ISERROR(SEARCH(&quot;yes&quot;,RC)))</Value1>
	 * <Format Style='color:#9C0006;background:#FFC7CE'/>
	 * </Condition>
	 * </ConditionalFormatting>
	 *
	 * @return
	 */
	public String getXML()
	{
		return getXML( this );
	}

	// helper inner method allows it to be static and reused...
	private static final String getXML( Cf cfx )
	{

		// reset the placeholder formula reference
		if( cfx.refPos > -1 )
		{
			cfx.resetFormulaRef();
		}

		StringBuffer xml = new StringBuffer();

		xml.append( "<Range>" );
		CellRange rn = new CellRange( cfx.getCondfmt().getBoundingRange() );//getConditionalFormatRange();
		if( rn != null )
		{
			xml.append( rn.getR1C1Range() );
		}

		xml.append( "</Range>" );
		xml.append( "<Condition>" );
		if( cfx.cp != 0x0 )
		{ // calcer
			xml.append( "<Qualifier>" );
			xml.append( cfx.getConditionString() );
			xml.append( "</Qualifier>" );
		}
		Object val1, val2;
		if( cfx.cp != 0x0 )
		{ // calcer
			xml.append( "<Value1>" );
			val1 = cfx.getFormula1().calculateFormula();
			if( val1 == null ) // a new formula + calc_explicit
			{
				xml.append( "" );
			}
			else
			{
				xml.append( val1.toString() );
			}
			xml.append( "</Value1>" );
		}
		else
		{
			xml.append( "<Value1>" );
			String fmx = cfx.getFormula1().getFormulaString();
			fmx = Cf.escapeFormulaString( fmx );
			xml.append( fmx );
			xml.append( "</Value1>" );
		}

		if( cfx.cce2 > 0 )
		{
			val2 = cfx.getFormula2().calculateFormula();
			xml.append( "<Value2>" );
			if( val2 == null ) // a new formula + calc_explicit
			{
				xml.append( "" );
			}
			else
			{
				xml.append( val2.toString() );
			}
			xml.append( "</Value2>" );
		}

		xml.append( "<Format Style='" );
		// attributes
		int cfi = cfx.getPatternFillColor();
		if( cfi > -1 )
		{
			String fsi = FormatHandle.colorToHexString( FormatHandle.getColor( cfi ) );
			xml.append( "color:" + fsi + ";" );
		}

		if( cfx.getFontWeight() > -1 )
		{
			xml.append( "font-weight:" + cfx.getFontWeight() + ";" );
		}

		if( cfx.getFontOptsCancellation() > -1 )
		{
			if( cfx.getFontOptsCancellation() == 0 )
			{
				xml.append( "text-line-through:none;" );
			}
			else
			{
				xml.append( "text-line-through:" + cfx.getFontOptsCancellation() + ";" );
			}
		}
		if( cfx.getBorderSizes() != null )
		{
			try
			{
				xml.append( "border-top:" + cfx.getBorderSizes()[0] + " " + FormatHandle.BORDER_NAMES[cfx.getBorderStyles()[0] + 1] + " " + FormatHandle
						.colorToHexString( cfx.getBorderColors()[0] ) + ";" ); // .5pt solid windowtext;
				xml.append( "border-left:" + cfx.getBorderSizes()[1] + " " + FormatHandle.BORDER_NAMES[cfx.getBorderStyles()[1] + 1] + " " + FormatHandle
						.colorToHexString( cfx.getBorderColors()[1] ) + ";" ); // .5pt solid windowtext;
				xml.append( "border-bottom:" + cfx.getBorderSizes()[2] + " " + FormatHandle.BORDER_NAMES[cfx.getBorderStyles()[2] + 1] + " " + FormatHandle
						.colorToHexString( cfx.getBorderColors()[2] ) + ";" ); // .5pt solid windowtext;
				xml.append( "border-right:" + cfx.getBorderSizes()[3] + " " + FormatHandle.BORDER_NAMES[cfx.getBorderStyles()[3] + 1] + " " + FormatHandle
						.colorToHexString( cfx.getBorderColors()[3] ) + ";" ); // .5pt solid windowtext;
			}
			catch( ArrayIndexOutOfBoundsException e )
			{
			}
		}
		cfi = cfx.getPatternFillColorBack();
		if( cfi > -1 )
		{
			String fsi = FormatHandle.colorToHexString( FormatHandle.getColor( cfi ) );
			xml.append( "background:" + fsi + ";" );
		}

		xml.append( "'/>" );
		xml.append( "</Condition>" );

		return xml.toString();
	}

	/**
	 * creates a Cf record from the EXML nodes
	 * <p/>
	 * <p/>
	 * <Condition>
	 * <Qualifier>Between</Qualifier>
	 * <Value1>2</Value1>
	 * <Value2>4</Value2>
	 * <Format Style='color:#002060;font-weight:700;text-line-through:none;
	 * border:.5pt solid windowtext;background:#00B0F0'/>
	 * </Condition>
	 *
	 * @param xpp
	 * @return
	 */
	public static Cf parseXML( XmlPullParser xpp, Condfmt cfx, Boundsheet bs )
	{
		Cf oe = bs.createCf( cfx );
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "Condition" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( n.equals( "Qualifier" ) )
							{
								oe.setOperator( v );
							}
							else if( n.equals( "Value1" ) )
							{
								oe.setCondition1( v );
							}
							else if( n.equals( "Value2" ) )
							{
								oe.setCondition2( v );
							}
							else if( n.equals( "Format" ) )
							{
								setStylePropsFromString( v, oe );
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "Condition" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "Cf.parseXML: " + e.toString() );
		}
		return oe;
	}

	/**
	 * OOXML-specific, in a containsText-type condition,
	 * this field is the comparitor
	 *
	 * @param s
	 */
	public void setContainsText( String s )
	{
		containsText = s;
	}

	/**
	 * generate OOXML for this Cf (BIFF8->OOXML)
	 * attributes:   type, dxfId, priority (REQ), stopIfTrue, aboveAverage,
	 * percent, bottom, operator, text, timePeriod, rank, stdDev, equalAverage
	 * children:		SEQ: formula (0-3), colorScale, dataBar, iconSet
	 *
	 * @return
	 */
	public String getOOXML( WorkBookHandle bk, int priority, ArrayList dxfs )
	{
		StringBuffer ooxml = new StringBuffer();

		// first deal with dfx's (differential xf's) - part of styles.xml; here we need to add dxf element to dxf's plus trap dxfId    	
		Dxf dxf = new Dxf();
		if( this.bHasFontBlock )
		{
			if( font != null )
			{
				dxf.setFont( font );
			}
			else
			{
				dxf.createFont( this.fontWeight, this.getFontItalic(), this.fontUnderlineStyle, this.fontColorIndex, this.fontHeight );
			}
		}
		if( this.bHasPatternBlock )
		{
			if( fill != null )
			{
				dxf.setFill( fill );
			}
			else
			{
				dxf.createFill( this.patternFillStyle, this.patternFillColor, this.patternFillColorBack, bk );
			}
		}
		if( this.bHasBorderBlock )
		{
			dxf.createBorder( bk, this.getBorderStyles(), new int[]{
					this.getBorderLineColorTop(),
					this.getBorderLineColorLeft(),
					this.getBorderLineColorBottom(),
					this.getBorderLineColorRight()
			} );
		}
		// TODO: check if this dxf already exists ****************************************************************************
		dxfs.add( dxf );    // save newly created dxf (differential xf) to workbook store 
		int dxfId = dxfs.size() - 1;    // link this cf to it's dxf  NOTE: one of the ONLY OOXML id's that is 0-based ... 

		ooxml.append( "<cfRule dxfId=\"" + dxfId + "\"" );    	
		/*
    	* attributes:   type, dxfId, priority (REQ), stopIfTrue, aboveAverage,
    	* 				percent, bottom, operator, text, timePeriod, rank, stdDev, equalAverage
    	* children:		SEQ: formula (0-3), colorScale, dataBar, iconSet
    	*/
		// TODO: ct==0 translates to ??
		// NOTE: types 3 and above are 2007 version (OOXML)-specific
		switch( this.ct )
		{
			case 1:
				ooxml.append( " type=\"cellIs\"" );
				break;
			case 2:
				ooxml.append( " type=\"expression\"" );
				break;
			case 3:
				ooxml.append( " type=\"containsText\"" );
				break;
			case 4:
				ooxml.append( " type=\"aboveAverage\"" );
				break;
			case 5:
				ooxml.append( " type=\"beginsWith\"" );
				break;
			case 6:
				ooxml.append( " type=\"colorScale\"" );
				break;
			case 7:
				ooxml.append( " type=\"containsBlanks\"" );
				break;
			case 8:
				ooxml.append( " type=\"containsErrors\"" );
				break;
			case 9:
				ooxml.append( " type=\"dataBar\"" );
				break;
			case 10:
				ooxml.append( " type=\"duplicateValues\"" );
				break;
			case 11:
				ooxml.append( " type=\"endsWith\"" );
				break;
			case 12:
				ooxml.append( " type=\"iconSet\"" );
				break;
			case 13:
				ooxml.append( " type=\"notContainsBlanks\"" );
				break;
			case 14:
				ooxml.append( " type=\"notContainsErrors\"" );
				break;
			case 15:
				ooxml.append( " type=\"notContainsText\"" );
				break;
			case 16:
				ooxml.append( " type=\"timePeriod\"" );
				break;
			case 17:
				ooxml.append( " type=\"top10\"" );
				break;
			case 18:
				ooxml.append( " type=\"uniqueValues\"" );
				break;
		}
		if( (this.ct == 3) && (containsText != null) )// containsText	- shouldn't be null!
		{
			ooxml.append( " text=\"" + containsText + "\"" );
		}

		// operator
		switch( this.cp )
		{
			case 01:    // Between
				ooxml.append( " operator=\"between\"" );
				break;
			case 0x5:    // Greater than
				ooxml.append( " operator=\"greaterThan\"" );
				break;
			case 0x2:    // Not between
				ooxml.append( " operator=\"notBetween\"" );
				break;
			case 0x6:    // Less than
				ooxml.append( " operator=\"lessThan\"" );
				break;
			case 0x3:    // Equal
				ooxml.append( " operator=\"equal\"" );
				break;
			case 0x7:    // Greater or equal
				ooxml.append( " operator=\"greaterThanOrEqual\"" );
				break;
			case 0x4:    // Not equal
				ooxml.append( " operator=\"notEqual\"" );
				break;
			case 0x8:    // Less or equal
				ooxml.append( " operator=\"lessThanOrEqual\"" );
				break;
			// 2007-specific 
			case 0x9:    // begins With
				ooxml.append( " operator=\"beginsWith\"" );
				break;
			case 0xA:    // ends With
				ooxml.append( " operator=\"endsWith\"" );
				break;
			case 0xB:    // begins With
				ooxml.append( " operator=\"containsText\"" );
				break;
			case 0xC:    // begins With
				ooxml.append( " operator=\"notContains\"" );
				break;
		}
		// priority
		ooxml.append( " priority=\"" + priority + "\"" );
		// stopIfTrue == looks like this is set by default
		ooxml.append( " stopIfTrue=\"1\"" );
		ooxml.append( ">" );
		if( this.getFormula1() != null )
		{
			ooxml.append( "<formula>" + OOXMLAdapter.stripNonAsciiRetainQuote( this.getFormula1().getFormulaString() )
			                                        .substring( 1 ) + "</formula>" );
		}
		if( this.getFormula2() != null )
		{
			ooxml.append( "<formula>" + OOXMLAdapter.stripNonAsciiRetainQuote( this.getFormula2().getFormulaString() )
			                                        .substring( 1 ) + "</formula>" );
		}

		// TODO: finish children dataBar, colorScale, iconSet, aboveAverage, bottom, equalAverage, 
		ooxml.append( "</cfRule>" );
		return ooxml.toString();
	}

}
