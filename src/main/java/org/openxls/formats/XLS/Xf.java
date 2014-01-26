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
package org.openxls.formats.XLS;

import org.openxls.ExtenXLS.FormatHandle;
import org.openxls.formats.OOXML.Fill;
import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.Color;

/**
 * <b>XF: Extended Format (E0h)</b><br>
 * The XF record stores formatting properties.
 * <p/>
 * If fStyle bit is true, then the XF is a style XF, otherwise
 * it is a BiffRec XF.  Cells and Styles both contain ixfe pointers
 * which correspond to their associated XF record.
 * <p/>
 * <pre>
 * BiffRec XF Record
 *
 * offset  Bits   MASK     name        contents
 * ---
 * 0       15-0   0xFFFF   ifnt        Index to the FONT record.
 * 2       15-0   0xFFFF   ifmt        Index to the FORMAT record.
 * 4       0      0x0001   fLocked     =1 if the cell is locked.
 * 1      0x0002   fHidden     =1 if the cell formula is hidden (value still shown)
 * 2      0x0004   fStyle      =0 for cell XF.
 * =1 for style XF.
 *
 * ~~~ additional option flags omitted ~~~
 *
 * </pre>
 *
 * @see SST
 * @see LABELSST
 * @see EXTSST
 */

public class Xf extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Xf.class );
	private static final long serialVersionUID = -419388613530529316L;

	int tableidx = -1;
	// Shared variables.
	private short ifnt;
	private short ifmt;
	private short fLocked;
	private short fHidden;
	private short fStyle;
	// the records that are initialized to "0" are a 1/0 flag.  Just makin my life easier...
	private short f123Prefix = 0;
	private short ixfParent;
	private short alc;
	private short fWrap = 0;
	private short alcV;
	// private short fJustLast; Used in Far East version of Excel only!!
	private short cIndent;
	private short trot;
	// private short cIntednt;
	private short fShrinkToFit = 0;
	private short fMergeCell = 0;
	private short iReadingOrder;
	private short fAtrNum = 0;
	private short fAtrFnt = 0;
	private short fAtrAlc = 0;
	private short fAtrBdr = 0;
	private short fAtrPat = 0;
	private short fAtrProt = 0;
	private short dgLeft;
	private short dgRight;
	private short dgTop;
	private short dgBottom;
	private short icvLeft;
	private short icvRight;
	private short grbitDiag;
	private short icvTop;
	private short icvBottom;
	private short icvDiag;
	private short dgDiag;
	private short fls;
	private short icvFore;
	private short icvBack;
	private short fSxButton = 0;
	private short icvColorFlag = 0;
	int Iflag = 0;
	byte mystery;
	public final static int NDEFAULTXFS = 20;
	private String pat = null;
	/**
	 * OOXML fill, if any
	 */
	public Fill fill = null;    // ugly that it's public ...
	// These should only be populated for boundsheet transferral issues.
	private Font myFont;
	private Format myFormat;
	private short useCount = 0;    // KSC: added 20121003 to keep track of xf usage by biffrecs

	public Xf()
	{
		//empty constructor
	}

	/**
	 * create a new Xf with pointer to its font
	 */
	public Xf( int f )
	{
		byte[] bl = { 0, 0, 0, 0, 1, 0, 32, 0, 0, 64, 0, 0, 0, 0, 0, 0, 0, 0, -64, 32 };
		setOpcode( XF );
		setLength( (short) (bl.length) );
		setData( bl );
		setFont( f );
		init();
	}

	/**
	 * create a new Xf with pointer to its font and workbook set
	 */
	public Xf( int f, WorkBook wkbook )
	{
		byte[] bl = { 0, 0, 0, 0, 1, 0, 32, 0, 0, 64, 0, 0, 0, 0, 0, 0, 0, 0, -64, 32 };
		setOpcode( XF );
		setLength( (short) (bl.length) );
		setData( bl );
		super.setWorkBook( wkbook );    // set workbook but don't insert rec or add to xfrecs
		setFont( f );
		init();
	}

	/**
	 * constructor which takes a Font object + a workbook
	 * useful for cloning xf's from other workbooks
	 *
	 * @param f      font
	 * @param wkbook
	 */
	public Xf( Font f, WorkBook wkbook )
	{
		byte[] bl = { 0, 0, 0, 0, 1, 0, 32, 0, 0, 64, 0, 0, 0, 0, 0, 0, 0, 0, -64, 32 };
		setOpcode( XF );
		setLength( (short) (bl.length) );
		myFont = f;
		System.arraycopy( ByteTools.shortToLEBytes( (short) f.getIdx() ), 0, bl, 0, 2 );
		setData( bl );
		super.setWorkBook( wkbook );    // set workbook but don't insert rec or add to xfrecs
		init();
	}

	/**
	 * Set the workbook for this XF
	 * <p/>
	 * This can get called multiple times.  This results in a disparity within
	 * xf counting in workbook.
	 */
	@Override
	public void setWorkBook( WorkBook b )
	{
		super.setWorkBook( b );
	}

	/**
	 * Create a string representation of the Xf
	 */
	public String toString()
	{
		String f = "unknown";        //Handle missing formats
		try
		{
			f = getFormatPattern();
		}
		catch( Exception e )
		{
			;
		}
		String thisToString = " format:" + f + " fill:" + getFillPattern() +
				" fg:" + getForegroundColor() +
				" bg:" + getBackgroundColor() +
				" border:[" +
				getTopBorderLineStyle() + "-" + getTopBorderColor() + ":" +
				getLeftBorderLineStyle() + "-" + getLeftBorderColor() + ":" +
				getBottomBorderLineStyle() + "-" + getBottomBorderColor() + ":" +
				getRightBorderLineStyle() + "-" + getRightBorderColor() + "]" +
				"W:" + getWrapText() +
				"R:" + getRotation() +
				"H:" + getHorizontalAlignment() + "V:" + getVerticalAlignment() +
				"I:" + getIndent() +
				"L:" + isLocked() +
				"F:" + isFormulaHidden() +
				"D:" + getRightToLeftReadingOrder();
		return getFont().toString() + thisToString;
	}

	/**
	 * inc # records using this xf
	 */
	public void incUseCount()
	{
		useCount++;
	}

	/**
	 * dec # records using this xf
	 */
	public void decUseCoount()
	{
		useCount--;
	}

	/**
	 * return # records using this xf
	 *
	 * @return
	 */
	public short getUseCount()
	{
		return useCount;
	}

	/**
	 * Populates the myFont and myFormat variables to be held onto
	 * when the xf record is serialized for boundsheet transfer
	 */
	protected void populateForTransfer()
	{
		myFont = getFont();
		myFormat = getWorkBook().getFormat( ifmt );
		getData();
	}

	public boolean getMerged()
	{
		if( fMergeCell == 1 )
		{
			return true;
		}
		return false;
	}

	// marginal!
	public void setMerged( boolean mgd )
	{
		byte[] rkdata = getData();
		rkdata[9] = 0x78; // 0xf4 ?
			log.trace( "Xf The merge style bit is: " + fMergeCell );
	}

	/**
	 * The XF record can either be a style XF or a Cell XF.
	 */
	@Override
	public void init()
	{
		super.init();
		ifnt = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		ifnt = (short) (ifnt & 0xffff);
		ifmt = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		ifmt = (short) (ifmt & 0xffff);

		short flag = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		// is the cell locked?
		if( (flag & 0x1) == 0x1 )
		{
			fLocked = 0x1;
		}
		else
		{
			fLocked = 0;
		}
		// is the cell hidden?
		if( (flag & 0x2) == 0x2 )
		{
			fHidden = 0x1;
		}
		else
		{
			fHidden = 0;
		}

		// is it a cell rec or a style rec?
		if( (flag & 0x4) == 0x4 )
		{
			fStyle = 1;
		}
		else
		{
			fStyle = 0;
		}
		if( (flag & 0x8) == 0x0008 )
		{
			f123Prefix = 0x1;
		}
		ixfParent = (short) ((flag & 0xFFF0) >> 4);

		initXF();

		pat = null;    // ensure reset if xf has changed
			log.trace( "Xf.init() ifnt: " + ifnt + " ifmt: " + ifmt + ":" +
					           toString() + " border: " + "l:" + getLeftBorderColor() + ":" + "b:" + getBottomBorderColor() + ":" + "r:" + getRightBorderColor() + ":" + "t:" + getTopBorderColor() + ":" );
	}

	/**
	 * read and interpret bytes 6-18)
	 */
	void initXF()
	{
		short flag;

		// bytes 6, 7: alignment, rotation, text break
		flag = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		alc = (short) (flag & 0x7);
		if( (flag & 0x8) == 0x8 )
		{
			fWrap = 1;
		}
		alcV = (short) ((flag & 0x70) >> 4);
		trot = (short) ((flag & 0xFF00) >> 8);

		// byte 8: indent, reading order, shrink
		flag = getByteAt( 8 );
		cIndent = (short) (flag & 0xF);
		if( (flag & 0x10) == 0x10 )
		{
			fShrinkToFit = 1;
		}
		if( (flag & 0x20) == 0x20 )
		{
			fMergeCell = 1;
		}
			log.trace( "Xf The merge cell bit is: " + fMergeCell + " and the int is " + flag );

		iReadingOrder = (short) ((flag & 0xC0));// >> 6);	// reading order is byte 7-6 mask 0xCO
		// USED_ATTRIB:	 bits 7-2 of byte 9 
		flag = getByteAt( 9 );
		/* for all these flags, a cleared bit means use Parent Style XF attribute
		 if set, means the attributes of THIS xf is used		
		bit mask 	meaning 
		0 	01H 	Flag for number format 
		1 	02H 	Flag for font
		2 	04H 	Flag for horizontal and vertical alignment, text wrap, indentation, orientation, rotation, and
		text direction
		3 	08H 	Flag for border lines
		4 	10H 	Flag for background area style
		5 	20H 	Flag for cell protection (cell locked and formula hidden)
		 */
		if( (flag & 0x4) == 0x4 )
		{
			fAtrNum = 1;        // number format
		}
		if( (flag & 0x8) == 0x8 )
		{
			fAtrFnt = 1;        // font
		}
		if( (flag & 0x10) == 0x10 )
		{
			fAtrAlc = 1;    // alignment (h + v) text wrap rotation direction indent
		}
		if( (flag & 0x20) == 0x20 )
		{
			fAtrBdr = 1;    // border lines
		}
		if( (flag & 0x40) == 0x40 )
		{
			fAtrPat = 1;    // background format pattern
		}
		if( (flag & 0x80) == 0x80 )
		{
			fAtrProt = 1;    // cell protection
		}

		// BORDER Section
		flag = ByteTools.readShort( getByteAt( 10 ), getByteAt( 11 ) );
		dgLeft = (short) (flag & 0xF);
		dgRight = (short) ((flag & 0xF0) >> 4);
		dgTop = (short) ((flag & 0xF00) >> 8);
		dgBottom = (short) ((flag & 0xF000) >> 12);

		flag = ByteTools.readShort( getByteAt( 12 ), getByteAt( 13 ) );
		icvLeft = (short) (flag & 0x7f);
		icvRight = (short) ((flag & 0x3F80) >> 7);
		grbitDiag = (short) ((flag & 0xC000) >> 15);

		// bytes 14-17 color and fill
		Iflag = ByteTools.readInt( getByteAt( 14 ), getByteAt( 15 ), getByteAt( 16 ), getByteAt( 17 ) );
		icvTop = (short) (Iflag & 0x7F);
		icvBottom = (short) ((Iflag & 0x3F80) >> 7);
		icvDiag = (short) ((Iflag & 0x1FC000) >> 14);
		dgDiag = (short) ((Iflag & 0x1E00000) >> 21);
		mystery = (byte) ((Iflag & 0x3800000) >> 25);
		fls = (short) ((Iflag & 0xFC000000) >> 26); // fill pattern

		if( (icvTop > 0) )
		{
			log.trace( "Xf The cell outline is true" );
		}
		// bytes 18, 19: fill pattern colors
		icvColorFlag = ByteTools.readShort( getByteAt( 18 ), getByteAt( 19 ) );
		icvFore = (short) (icvColorFlag & 0x7F);                // = Pattern Color
		icvBack = (short) ((icvColorFlag & 0x3F80) >> 7);    // = Pattern Background Color
		if( (icvColorFlag & 0x4000) == 0x4000 )
		{
			fSxButton = 1;
		}

		// Logger.logInfo(org.openxls.ExtenXLS.ExcelTools.getRecordByteDef(this));
	}

	/**
	 * returns the associated  Font record for this XF
	 */
	@Override
	public Font getFont()
	{
		if( myFont != null )
		{
			return myFont;
		}
		myFont = getWorkBook().getFont( ifnt );
		return myFont;
	}

	/**
	 * returns whether this Format is a Date
	 * <p/>
	 * Needs to be revisited.  Currently I am only returning true for the standard "built in" dates
	 */
	public boolean isDatePattern()
	{

		// Check the format ID against all known date formats. Why do we do
		// this instead of letting it be caught by the string matching below?
		for( int x = 0; x < FormatConstants.DATE_FORMATS.length; x++ )
		{
			short sxt = (short) Integer.parseInt( FormatConstants.DATE_FORMATS[x][1], 16 );
			if( ifmt == sxt )
			{
				return true;
			}
		}

		Format fmt = getWorkBook().getFormat( ifmt );
		if( fmt == null )
		{
			return false;
		}

		// toLowerCase is a simplistic way to implement the case insensitivity
		// of the pattern tokens. It could cause issues with string literals.
		String myfmt = fmt.getFormat().toLowerCase();
		return isDatePattern( myfmt );
	}

	public static boolean isDatePattern( String myfmt )
	{
		// Search for the format string in the list of known date formats
		for( int x = 0; x < FormatConstants.DATE_FORMATS.length; x++ )
		{
			if( FormatConstants.DATE_FORMATS[x][0].equals( myfmt ) )
			{
				return true;
			}
		}

		// check for string patterns that only exist within date records (as far as we know, may need refining)
		if( (myfmt.indexOf( "mm" ) > -1) || (myfmt.indexOf( "yy" ) > -1) || (myfmt.indexOf( "dd" ) > -1) )
		{
			return true;
		}
		return false;
	}

	/**
	 * Parses an escaped xml format pattern (from ooxml) and returns an extenxls compatible
	 * pattern.
	 * <p/>
	 * This method certainly has weaknesses, but my intention is that if it is not a fairly standard format and/or
	 * we are not sure how to parse it we should leave the existent format intact soas to not break read/write operations
	 * <p/>
	 * Oddly enough, excel seems to be able to handle biff8 patterns, in my testing so far there has been no need
	 * to reencode, that could obviously change...
	 *
	 * @param xmlFormatPattern
	 * @return compatible biff8 formatPattern
	 */
	public static String unescapeFormatPattern( String xmlFormatPattern )
	{
		// strip escaping for currency pattern.  Probably should explore all currency types and do an iteration
		xmlFormatPattern = xmlFormatPattern.replace( "\"$\"", "$" );

		// separator between positive/negative
		xmlFormatPattern = xmlFormatPattern.replace( "_);", ";" );

		// unescape parens
		xmlFormatPattern = xmlFormatPattern.replace( "\\(", "(" );
		xmlFormatPattern = xmlFormatPattern.replace( "\\)", ")" );
		return xmlFormatPattern;
	}

	/**
	 * returns whether this Format is a Currency
	 */
	public boolean isCurrencyPattern()
	{
		if( pat == null )
		{
			setFormatPattern( getFormatPattern() );
		}
		for( int x = 0; x < FormatConstants.CURRENCY_FORMATS.length; x++ )
		{
			short cpt = (short) Integer.parseInt( FormatConstants.CURRENCY_FORMATS[x][1], 16 );
			if( ifmt == cpt )
			{
				String ptx = FormatConstants.CURRENCY_FORMATS[x][0];
				// what up with this?	General?
				if( cpt == 1 )
				{
					if( pat.equals( ptx ) )
					{
						return true;
					}
					return false;

				}
				return true;
			}
		}
		// probably a built-in format that is not a currency format
		Format fmt = getWorkBook().getFormat( ifmt );
		if( fmt == null )
		{
			return false;
		}
		String myfmt = fmt.getFormat();
		for( int x = 0; x < FormatConstants.CURRENCY_FORMATS.length; x++ )
		{
			if( FormatConstants.CURRENCY_FORMATS[x][0].equals( myfmt ) )
			{
				return true;
			}
		}
		return false;
	}

	/**
	 * get the Pattern Color index for this Cell if Solid Fill, or the Foreground color if no Solid Pattern
	 */
	public short getForegroundColor()
	{
		if( fill != null )
		{
			return (short) fill.getFgColorAsInt( getWorkBook().getTheme() );
		}
		return icvFore;
	}

	/**
	 * get the Pattern Color for this Cell if Solid Fill, or the Foreground color if no Solid Pattern, as a Hex Color String
	 *
	 * @return Hex Color String
	 */
	public String getForegroundColorHEX()
	{
		if( fill != null )
		{
			return fill.getFgColorAsRGB( getWorkBook().getTheme() );
		}
		return FormatHandle.colorToHexString( FormatHandle.getColor( icvFore ) );

	}

	/**
	 * get the background Color for this Cell as a Hex Color String
	 *
	 * @return Hex Color String
	 */
	public String getBackgroundColorHEX()
	{
		if( fill != null )
		{
			return fill.getBgColorAsRGB( getWorkBook().getTheme() );
		}
		if( icvBack == 65 ) // default background color
		{
			return "#FFFFFF";    // return white
		}
		return FormatHandle.colorToHexString( FormatHandle.getColor( icvBack ) );
	}

	/**
	 * get the background Color index for this Cell
	 */
	public short getBackgroundColor()
	{
		if( fill != null )
		{
			return (short) fill.getBgColorAsInt( getWorkBook().getTheme() );
		}
		if( icvBack == 65 ) // default background color
		{
			return 64; // return white
		}
		return icvBack;
	}

	/**
	 * get the Formatting for this BiffRec from the pattern
	 * match.
	 * <p/>
	 * case insensitive pattern match is performed...
	 */
	@Override
	public String getFormatPattern()
	{
		if( pat != null )
		{
			return pat;
		}
		String[][] fmts = FormatConstantsImpl.getBuiltinFormats();
		for( String[] fmt1 : fmts )
		{
			if( ifmt == Integer.parseInt( fmt1[1], 16 ) )
			{
				pat = fmt1[0];
				return pat;
			}
		}

		Format fmt = getWorkBook().getFormat( ifmt );
		if( fmt != null )
		{
			pat = fmt.toString();
			return fmt.getFormat();
		}
		return null;
	}

	/**
	 * Sets the number format pattern for this format.
	 */
	public void setFormatPattern( String pattern )
	{
		pat = pattern;

		if( getWorkBook() == null )
		{
			throw new IllegalStateException( "attempting to set format pattern but workbook is null" );
		}

		setFormat( addFormatPattern( getWorkBook(), pattern ) );
	}

	/**
	 * Ensures that the given format pattern exists on the given workbook.
	 *
	 * @param book    the workbook to which the pattern should belong
	 * @param pattern the number format pattern to ensure exists
	 * @return the format ID of the given format pattern
	 */
	public static short addFormatPattern( WorkBook book, String pattern )
	{
		short ifmt = -1;

		// Look up the pattern on the workbook
		ifmt = book.getFormatId( pattern );

		// If the pattern is unknown, create and add a Format record
		if( ifmt == -1 )
		{
			Format format = new Format( book, pattern );
			ifmt = format.getIfmt();
		}

		return ifmt;
	}

	/**
	 * set the pointer to the XF's Format in the WorkBook
	 */
	public void setFormat( short ifm )
	{
		ifmt = ifm;
		byte[] nef = ByteTools.shortToLEBytes( ifmt );
		getData()[2] = nef[0];
		getData()[3] = nef[1];
		pat = null;    // 20080228 KSC: flag to re-input
	}

	/**
	 * set the pointer to the XF's Font in the WorkBook
	 */
	public void setFont( int ifn )
	{
		ifnt = (short) ifn;
		byte[] nef = ByteTools.shortToLEBytes( ifnt );
		getData()[0] = nef[0];
		getData()[1] = nef[1];
		// reset the pointer for xf's and font's brought from other workbooks.
		if( getWorkBook() != null )
		{
			myFont = getWorkBook().getFont( ifn );
		}
	}

	/**
	 * Set myFont in XF to the same as Workbook's
	 */
	public void setMyFont( Font f )
	{
		myFont = f;
	}

	/**
	 * set the Fill Pattern for this Format
	 */
	public void setPattern( int t )
	{
		fls = (short) t;
		updatePattern();
		if( fill != null )
		{
			fill.setFillPattern( t );
		}
	}

	/**
	 * get the format pattern for this particular XF
	 *
	 * @return
	 */
	public int getFillPattern()
	{
		if( fill != null )
		{
			return fill.getFillPatternInt();
		}
		return fls;
	}

	// BORDER SECTION

	public short getBottomBorderLineStyle()
	{
		return dgBottom;
	}

	public short getTopBorderLineStyle()
	{
		return dgTop;
	}

	public short getLeftBorderLineStyle()
	{
		return dgLeft;
	}

	public short getRightBorderLineStyle()
	{
		return dgRight;
	}

	public short getDiagBorderLineStyle()
	{
		return dgDiag;
	}

	/**
	 * set the Top Border Color for this Format
	 */
	public void setTopBorderColor( int t )
	{
		if( t == 0 )
		{
			t = 64;    // 20080118 KSC
		}
		icvTop = (short) t;
		updateBorderColors();
	}

	public int getTopBorderColor()
	{
		// 20070205 KSC: 64 is automatic border color but should be interpreted as 65
		if( icvTop == 64 )
		{
			return 65;
		}
		return icvTop;
	}

	/**
	 * set the Bottom Border Color for this Format
	 */
	public void setBottomBorderColor( int t )
	{
		if( t == 0 )
		{
			t = 64;    // 20080118 KSC
		}
		icvBottom = (short) t;
		updateBorderColors();
	}

	public int getBottomBorderColor()
	{
		// 20070205 KSC: 64 is automatic border color but should be interpreted as 65
		if( icvBottom == 64 )
		{
			return 65;
		}
		return icvBottom;
	}

	/**
	 * set the Left Border Color for this Format
	 */
	public void setLeftBorderColor( int t )
	{
		if( t == 0 )
		{
			t = 64;    // 20080118 KSC
		}
		icvLeft = (short) t;
		updateBorderColors();
	}

	public int getLeftBorderColor()
	{
		// 20070205 KSC: 64 is automatic border color but should be interpreted as 65
		if( icvLeft == 64 )
		{
			return 65;
		}
		return icvLeft;
	}

	/**
	 * set the Right Border Color for this Format
	 */
	public void setRightBorderColor( int t )
	{
		if( t == 0 )
		{
			t = 64;    // 20080118 KSC
		}
		icvRight = (short) t;
		updateBorderColors();

	}

	public short getRightBorderColor()
	{
		// 20070205 KSC: 64 is automatic border color but should be interpreted as 65
		if( icvRight == 64 )
		{
			return 65;
		}
		return icvRight;
	}

	/**
	 * set the diagonal Border Color for this Format
	 */
	public void setDiagBorderColor( int t )
	{
		if( t == 0 )
		{
			t = 64;    // 20080118 KSC
		}
		icvDiag = (short) t;
		updateBorderColors();

	}

	/**
	 * get the diagonal border color
	 *
	 * @return
	 */
	public short getDiagBorderColor()
	{
		// 20070205 KSC: 64 is automatic border color but should be interpreted as 65
		if( icvDiag == 64 )
		{
			return 65;
		}
		return icvDiag;
	}

	/**
	 * set the Left Border Color for this Format
	 */
	public void setLeftBorderColor( short t )
	{
		if( t == 0 )
		{
			t = 64;    // 20080118 KSC
		}
		icvLeft = t;
		updateBorderColors();
	}

	/**
	 * set the diagonal border for this Format
	 */
	public void setBorderDiag( int t )
	{
		Iflag = 0;
		Iflag |= icvTop;
		Iflag |= (icvBottom << 7);
		Iflag |= ((short) t << 14);
		Iflag |= (dgDiag << 21);
		Iflag |= (mystery << 25);
		Iflag |= (fls << 26);
		updatePattern();
	}

	/**
	 * set the border line style for this Format
	 */
	public void setBorderLineStyle( short t )
	{
		dgLeft = t;
		dgRight = t;
		dgTop = t;
		dgBottom = t;
		updateBorders();
	}

	/**
	 * set border line styles via array of ints representing border styles
	 * order= top, left, bottom, right [diagonal]
	 *
	 * @param b int[]
	 */
	public void setAllBorderLineStyles( int[] b )
	{
		try
		{
			if( b[0] > -1 )
			{
				dgTop = (short) b[0];
			}
			if( b[1] > -1 )
			{
				dgLeft = (short) b[1];
			}
			if( b[2] > -1 )
			{
				dgBottom = (short) b[2];
			}
			if( b[3] > -1 )
			{
				dgRight = (short) b[3];
			}
			if( b[4] > -1 )
			{
				dgDiag = (short) b[4];
			}
		}
		catch( ArrayIndexOutOfBoundsException e )
		{
		}
		updateBorders();
	}

	/**
	 * set all border colors via an array of ints representing border color ints
	 * order= top, left, bottom, right, [diagonal]
	 *
	 * @param b int[]
	 */
	public void setAllBorderColors( int[] b )
	{
		try
		{
			if( b[0] > -1 )
			{
				icvTop = (short) b[0];
			}
			if( b[1] > -1 )
			{
				icvLeft = (short) b[1];
			}
			if( b[2] > -1 )
			{
				icvBottom = (short) b[2];
			}
			if( b[3] > -1 )
			{
				icvRight = (short) b[3];
			}
			if( b[4] > -1 )
			{
				icvDiag = (short) b[4];
			}
		}
		catch( ArrayIndexOutOfBoundsException e )
		{
		}
		updateBorderColors();
	}

	public void setTopBorderLineStyle( short t )
	{
		dgTop = t;
		updateBorders();
	}

	public void setBottomBorderLineStyle( short t )
	{
		dgBottom = t;
		updateBorders();
	}

	public void setLeftBorderLineStyle( short t )
	{
		dgLeft = t;
		updateBorders();
	}

	public void setRightBorderLineStyle( short t )
	{
		dgRight = t;
		updateBorders();
	}

	public void updateBorders()
	{
		short borderflag = 0;
		borderflag = dgLeft;
		borderflag = (short) ((borderflag | ((dgRight) << 4)));
		borderflag = (short) ((borderflag | ((dgTop) << 8)));
		borderflag = (short) ((borderflag | ((dgBottom) << 12)));
		//byte[] rkdata = this.getData();
		byte[] bords = ByteTools.shortToLEBytes( borderflag );
		getData()[10] = bords[0];
		getData()[11] = bords[1];
		setAttributeFlag();
	}

	/**
	 * removes all borders for the style
	 */
	public void removeBorders()
	{
		dgBottom = 0;
		dgTop = 0;
		dgDiag = 0;
		dgLeft = 0;
		dgRight = 0;
		dgBottom = 0;
		updateBorders();
	}

	public void updateBorderColors()
	{
		setAttributeFlag();
	}

	public void updatePattern()
	{
		byte[] rkdata = getData();
		short thisFlag = 0;
		thisFlag |= icvLeft;
		thisFlag |= (icvRight << 7);
		thisFlag |= (grbitDiag << 14);
		byte[] bytes = ByteTools.shortToLEBytes( thisFlag );
		rkdata[12] = bytes[0];
		rkdata[13] = bytes[1];

		Iflag = 0;
		Iflag |= icvTop;
		Iflag |= (icvBottom << 7);
		Iflag |= (icvDiag << 14);
		Iflag |= (dgDiag << 21);
		Iflag |= (mystery << 25);
		Iflag |= (fls << 26);
		byte[] nef = ByteTools.cLongToLEBytes( Iflag );
		rkdata[14] = nef[0];
		rkdata[15] = nef[1];
		rkdata[16] = nef[2];
		rkdata[17] = nef[3];
		// update format cache upon change
		pat = null;
		wkbook.updateFormatCache( this );
	}

	/**
	 * set the Foreground Color for this Format
	 * THIS SETS THE BACKGROUND COLOR when PATTERN (fls) = PATTERN_SOLID
	 * THIS SETS THE PATTERN COLOR when PATTERN (fls) > PATTERN_SOLID
	 * <br>"If the fill style is solid: When solid is specified, the
	 * foreground color (fgColor) is the only color rendered,
	 * even when a background color (bgColor) is also specified"
	 * icvFore==Pattern Background Color
	 *
	 * @param clr java.awt.Color or null if use standard Excel 2003 Color Table
	 * @param t   best match index into 2003-style Color tabe
	 */
	public void setForeColor( int t, Color clr )
	{
		icvColorFlag = 0;
		icvColorFlag |= ((short) t);
		icvColorFlag |= (icvBack << 7);
		if( clr != null )
		{
			if( !clr.equals( FormatHandle.COLORTABLE[t] ) )
			{ // no exact match for color
				if( fill == null )
				{
					fill = new Fill( getFillPattern(),
					                 t,
					                 FormatHandle.colorToHexString( clr ),
					                 icvBack,
					                 null, getWorkBook().getTheme() );
				}
				else
				{
					fill.setFgColor( t, FormatHandle.colorToHexString( clr ) );
				}
			}
		}
		else if( fill != null )
		{
			fill.setFgColor( t );
		}
		updateColors();
	}

	/**
	 * set the Background Color for this Format (when PATTERN - fls != PATTERN_SOLID)
	 * When PATTERN is PATTERN_SOLID, == 64
	 *
	 * @param clr java.awt.Color or null if use standard Excel 2003 Color Table
	 * @param t   best-match index into 2003-style Color table
	 */
	public void setBackColor( int t, Color clr )
	{
		icvColorFlag = 0;
		icvColorFlag |= icvFore;
		icvColorFlag |= ((short) t << 7);
		if( clr != null )
		{
			if( !clr.equals( FormatHandle.COLORTABLE[t] ) )
			{ // no exact match for color - store custom color
				if( fill == null )
				{
					fill = new Fill( getFillPattern(),
					                 icvFore,
					                 null,
					                 t,
					                 FormatHandle.colorToHexString( clr ), getWorkBook().getTheme() );
				}
				else
				{
					fill.setBgColor( t, FormatHandle.colorToHexString( clr ) );
				}
			}
		}
		else if( fill != null )
		{
			fill.setBgColor( t );
		}

		updateColors();
	}

	void updateColors()
	{
		byte[] rkdata = getData();
		byte[] nef = ByteTools.shortToLEBytes( icvColorFlag );
		rkdata[18] = nef[0];
		rkdata[19] = nef[1];
		icvFore = (short) (icvColorFlag & 0x7F);
		icvBack = (short) ((icvColorFlag & 0x3F80) >> 7);
		// update format cache upon change
		pat = null;
		wkbook.updateFormatCache( this );
	}

	/**
	 * PATTERN_SOLID is a special case where icvFore= the background color and icvBack=64.
	 * "If the fill style is solid: When solid is specified, the
	 * foreground color (fgColor) is the only color rendered,
	 * even when a background color (bgColor) is also
	 * specified"
	 */
	public static final int PATTERN_SOLID = 1;    // was set to 4 but tht's wrong!!

	/**
	 * Sets the fill pattern to solid, which renders the background to 64=="the default fg color"
	 * "If the fill style is solid: When solid is specified, the
	 * foreground color (fgColor) is the only color rendered,
	 * even when a background color (bgColor) is also
	 * specified"
	 */
	public void setBackgroundSolid()
	{
		setPattern( PATTERN_SOLID );
		setBackColor( 64, null );
		if( fill != null )
		{
			fill.setFillPattern( PATTERN_SOLID );
		}
	}

	public boolean isBackgroundSolid()
	{
		if( fill != null )
		{
			return fill.isBackgroundSolid();
		}
		byte[] rkdata = getData();
		return (rkdata[17] == (byte) PATTERN_SOLID);
	}

	/**
	 * Sets the attribute flags for this xf record.  These flags consist of
	 * // bit 8= fAtrProt
	 * //     7= fAtrPat
	 * //     6= fAtrBdr
	 * //     5= fAtrAlc (Alignment)
	 * //     4= fAtrFnt
	 * //     3= fAtrNum
	 */
	private void setAttributeFlag()
	{
		setToCellXF();
		byte[] rkdata = getData();
		byte used_attrib = rkdata[9];
		byte borderFlag = (byte) (((dgBottom > 0) || (dgTop > 0) || (dgLeft > 0) || (dgRight > 0) || (dgDiag > 0)) ? 1 : 0);    // if border is set
		if( borderFlag == 1 )
		{
			used_attrib = (byte) (used_attrib | 0x20);    // set bit # 6
		}
		else
		{
			used_attrib = (byte) (used_attrib & 0xDF);    // clear it
		}
		if( (cIndent != 0) || (iReadingOrder != 0) || (alc != 0) || (alcV != 0) || (fWrap != 0) || (trot != 0) )    // set bit # 5
		{
			used_attrib = (byte) (used_attrib | 0x10);
		}
		else
		{
			used_attrib = (byte) (used_attrib & 0xEF);  // clear it
		}

		rkdata[9] = used_attrib;
		fAtrNum = (short) (((used_attrib & 0x04) == 0x04) ? 1 : 0);
		fAtrFnt = (short) (((used_attrib & 0x08) == 0x08) ? 1 : 0);
		fAtrAlc = (short) (((used_attrib & 0x10) == 0x10) ? 1 : 0);
		fAtrBdr = (short) (((used_attrib & 0x20) == 0x20) ? 1 : 0);
		fAtrPat = (short) (((used_attrib & 0x40) == 0x40) ? 1 : 0);
		fAtrProt = (short) (((used_attrib & 0x80) == 0x80) ? 1 : 0);

		// must set color flag for borders or Excel will not like [BugTracker 2861]
		if( (dgTop > 0) && (icvTop == 0) )
		{
			icvTop = 64;
		}
		if( (dgBottom > 0) && (icvBottom == 0) )
		{
			icvBottom = 64;
		}
		if( (dgRight > 0) && (icvRight == 0) )
		{
			icvRight = 64;
		}
		if( (dgLeft > 0) && (icvLeft == 0) )
		{
			icvLeft = 64;
		}
		if( (dgDiag > 0) && (icvDiag == 0) )
		{
			icvDiag = 64;
		}
		updatePattern();
	}

	/**
	 * Switch the record to a cell XF record
	 */
	public void setToCellXF()
	{
		if( fStyle != 0 )
		{// must set to cell xf (fStyle==0) as changes will not show [BugTracker 2861]
			fStyle = 0;
			byte flag = (byte) fLocked;
			flag = (byte) ((flag | ((fHidden) << 1)));
			getData()[4] = flag;
			getData()[5] = 0;   // upper bits are style parent rec index
		}
	}

	/**
	 * @return Returns the ifnt.
	 */
	public short getIfnt()
	{
		return ifnt;
	}

	/**
	 * @param ifnt The ifnt to set.
	 */
	public void setIfnt( short ifnt )
	{
		this.ifnt = ifnt;
	}

	public short getIfmt()
	{
		return ifmt;
	}

	public void setHorizontalAlignment( int hAlign )
	{
		alc = (short) hAlign;
		updateAlignment();
		setAttributeFlag();
	}

	/**
	 * set the indent (1=3 spaces)
	 *
	 * @param indent
	 */
	public void setIndent( int indent )
	{ // indent # = 3 spaces
		cIndent = (short) indent;    // mask = 0xF, 4 bits,
		byte b = (byte) (getData()[8] & 0xF0);
		b |= (cIndent);    // 1st 4 bits
		getData()[8] = b;
		if( (alc != FormatConstants.ALIGN_LEFT) || (alc != FormatConstants.ALIGN_RIGHT) )    // indent only valid for Left and Right (apparently
		{
			setHorizontalAlignment( FormatConstants.ALIGN_LEFT );
		}
		setAttributeFlag();
	}

	/**
	 * return the indent setting (1=3 spaces)
	 *
	 * @return
	 */
	public int getIndent()
	{
		return cIndent;
	}

	/**
	 * sets the Right to Left Text Direction or reading order of this style
	 *
	 * @param rtl possible values:
	 *            <br>0=Context Dependent
	 *            <br>1=Left-to-Right
	 *            <br>2=Right-to-Let
	 */
	public void setRightToLeftReadingOrder( int rtl )
	{
		// iReadingOrder= bits 7-6
		// 00= According to Context
		// 01= Left to Right (0x40)
		// 10= Right to Left (0x80)
		if( rtl == 2 )
		{
			iReadingOrder = 0x80;
		}
		else if( rtl == 1 )
		{
			iReadingOrder = 0x40;
		}
		else
		{
			iReadingOrder = 0;
		}
		byte b = getData()[8];
		b |= (iReadingOrder);
		getData()[8] = b;
		wkbook.updateFormatCache( this );
		setAttributeFlag();
	}

	/**
	 * returns true if this style is set to Right-to-Left text direction (reading order)
	 *
	 * @return
	 */
	public int getRightToLeftReadingOrder()
	{
		return iReadingOrder >> 6;
	}

	public int getHorizontalAlignment()
	{
		return alc;
	}

	public void setWrapText( boolean wraptext )
	{
		if( wraptext )
		{
			fWrap = 1;
		}
		else
		{
			fWrap = 0;
		}
		updateAlignment();
		setAttributeFlag();
	}

	public boolean getWrapText()
	{
		return (fWrap == 1) ? true : false;
	}

	public void setVerticalAlignment( int vAlign )
	{
		alcV = (short) vAlign;
		updateAlignment();
		setAttributeFlag();
	}

	public int getVerticalAlignment()
	{
		return alcV;
	}

	public void setRotation( int rot )
	{
		trot = (short) rot;
		updateAlignment();
		setAttributeFlag();
	}

	public int getRotation()
	{
		return trot;
	}

	private void updateAlignment()
	{
		//short tempAlc = (short)(alc << 3);
		short tempfWrap = (short) (fWrap << 3);
		short tempAlcV = (short) (alcV << 4);
		short tempTrot = (short) (trot << 8);
		short res = 0x0;
		res = (short) (res | alc);
		res = (short) (res | tempfWrap);
		res = (short) (res | tempAlcV);
		res = (short) (res | tempTrot);
		byte[] rkdata = getData();
		byte[] bords = ByteTools.shortToLEBytes( res );
		rkdata[6] = bords[0];
		rkdata[7] = bords[1];
		// update format cache upon change
		wkbook.updateFormatCache( this );
	}

	/**
	 * Returns an XML fragment representing the XF backing the format Handle.  The XF record is style information
	 * associated with a cell.  Font information/lookup is not included in this output so it can be used as a comparitor
	 * style
	 */
	public String getXML()
	{
		StringBuffer sb = new StringBuffer( "<XF" );
		sb.append( ">" );
		Font myf = getFont();
		// font info...
		sb.append( "<font name=\"" + myf.getFontName() );
		sb.append( "\" size=\"" + myf.getFontHeightInPoints() );
		sb.append( "\" color=\"" + FormatHandle.colorToHexString( myf.getColorAsColor() ) );
		sb.append( "\" weight=\"" + myf.getFontWeight() );
		if( myf.getIsBold() )
		{
			sb.append( "\" bold=\"1" );
		}
		sb.append( "\" />" );
		// format info, should be expanded prolly
		sb.append( "<format id=\"" + ifmt );
		sb.append( "\" />" );
		// 20071218 KSC: Add Fill
		sb.append( "<fill id=\"" + fls );
		sb.append( "\" />" );
		// get the border..
		sb.append( "<Borders>" );
		if( getRightBorderLineStyle() != 0 )
		{
			sb.append( "<Border" );
			sb.append( " Position=\"right\"" );
			sb.append( " LineStyle=\"" + FormatHandle.BORDER_NAMES[getRightBorderLineStyle()] + "\"" );
			sb.append( " Color=\"" + FormatHandle.colorToHexString( wkbook.colorTable[getRightBorderColor()] ) + "\"" );
			sb.append( "/>" );
		}
		if( getBottomBorderLineStyle() != 0 )
		{
			sb.append( "<Border" );
			sb.append( " Position=\"bottom\"" );
			sb.append( " LineStyle=\"" + FormatHandle.BORDER_NAMES[getBottomBorderLineStyle()] + "\"" );
			sb.append( " Color=\"" + FormatHandle.colorToHexString( wkbook.colorTable[getBottomBorderColor()] ) + "\"" );
			sb.append( "/>" );
		}
		if( getLeftBorderLineStyle() != 0 )
		{
			sb.append( "<Border" );
			sb.append( " Position=\"left\"" );
			sb.append( " LineStyle=\"" + FormatHandle.BORDER_NAMES[getLeftBorderLineStyle()] + "\"" );
			sb.append( " Color=\"" + FormatHandle.colorToHexString( wkbook.colorTable[getLeftBorderColor()] ) + "\"" );
			sb.append( "/>" );
		}
		if( getTopBorderLineStyle() != 0 )
		{
			sb.append( "<Border" );
			sb.append( " Position=\"top\"" );
			sb.append( " LineStyle=\"" + FormatHandle.BORDER_NAMES[getTopBorderLineStyle()] + "\"" );
			sb.append( " Color=\"" + FormatHandle.colorToHexString( wkbook.colorTable[getTopBorderColor()] ) + "\"" );
			sb.append( "/>" );
		}
		sb.append( "</Borders>" );

		// get the alignment..
		sb.append( "<Alignment" );
		sb.append( " Horizontal=\"" + FormatHandle.HORIZONTAL_ALIGNMENTS[getHorizontalAlignment()] + "\"" );
		sb.append( " />" );

		// get the background color
		if( !wkbook.colorTable[getForegroundColor()].equals( Color.WHITE ) )
		{
			sb.append( "<Interior Color=\"" +
					           FormatHandle.colorToHexString( wkbook.colorTable[getForegroundColor()] ) +
					           "\"/>" );
		}

		sb.append( "</XF>" );
		return sb.toString();
	}

	/**
	 * @return Returns the myFormat.
	 */
	public Format getFormat()
	{
		return myFormat;
	}

	/**
	 * @param myFormat The myFormat to set.
	 */
	public void setFormat( Format myFormat )
	{
		this.myFormat = myFormat;
	}

	/**
	 * get whether this cell formula is hidden
	 *
	 * @return
	 */
	public boolean isFormulaHidden()
	{
		return (fHidden == 0x1);
	}

	/**
	 * sets the cell formula as hidden
	 *
	 * @param hd
	 */
	public void setFormulaHidden( boolean hd )
	{
		if( hd )
		{
			fHidden = 0x1;
		}
		else
		{
			fHidden = 0x0;
		}
		updateLockedHidden();
	}

	/**
	 * get whether this is a locked Cell
	 *
	 * @return
	 */
	public boolean isLocked()
	{
		return (fLocked == 0x1);
	}

	/**
	 * return whether this cell is set to "shrink to fit"
	 *
	 * @return
	 */
	public boolean isShrinkToFit()
	{
		return (fShrinkToFit == 0x1);
	}

	public void setShrinkToFit( boolean b )
	{
		if( b )
		{
			fShrinkToFit = 0x1;
			getData()[9] |= 0x10;
		}
		else
		{
			fShrinkToFit = 0x0;    // turn off bit 4
			getData()[9] &= 0xF7;    // set bit 4
		}
	}

	/**
	 * sets the cell as locked
	 *
	 * @param lk
	 */
	public void setLocked( boolean lk )
	{
		if( lk )
		{
			fLocked = 0x1;
		}
		else
		{
			fLocked = 0x0;
		}
		updateLockedHidden();
	}

	/**
	 * 2 2 XF type, cell protection, and parent style XF:
	 * Bit Mask Contents
	 * 2-0 0007H XF_TYPE_PROT – XF type, cell protection (see above)
	 * 15-4 FFF0H Index to parent style XF (always FFFH in style XFs)
	 * <p/>
	 * Bit Mask Contents
	 * 0 01H 1 = Cell is locked
	 * 1 02H 1 = Formula is hidden
	 * 2 04H 0 = Cell XF; 1 = Style XF
	 */
	private void updateLockedHidden()
	{

		short tempFL = (short) (fLocked << 0x0);
		short tempFH = (short) (fHidden << 0x1);
		short tempST = (short) (fStyle << 0x2);

		short flag = 0x0;
		flag = (short) (flag | tempFL);
		flag = (short) (flag | tempFH);
		flag = (short) (flag | tempST);

		byte[] dx = getData();
		byte[] nef = ByteTools.shortToLEBytes( flag );
		dx[4] = nef[0];
		dx[5] = nef[1];
		// update format cache upon change
		pat = null;
		wkbook.updateFormatCache( this );
	}

	/**
	 * @return
	 */
	public boolean getStricken()
	{
		if( myFont != null )
		{
			return myFont.getStricken();
		}
		return false;
	}

	public void setStricken( boolean b )
	{
		if( myFont != null )
		{
			myFont.setStricken( b );
		}
	}

	/**
	 * @return
	 */
	public boolean getItalic()
	{
		if( myFont != null )
		{
			return myFont.getItalic();
		}
		return false;
	}

	public void setItalic( boolean b )
	{
		if( myFont != null )
		{
			myFont.setItalic( b );
		}
	}

	/**
	 * @return
	 */
	public boolean getUnderlined()
	{
		if( myFont != null )
		{
			return myFont.getUnderlined();
		}
		return false;
	}

	public void setUnderlined( boolean b )
	{
		if( myFont != null )
		{
			myFont.setUnderlined( b );
		}
	}

	/**
	 * @return
	 */
	public boolean getBold()
	{
		if( myFont != null )
		{
			return myFont.getBold();
		}
		return false;
	}

	public void setBold( boolean b )
	{
		if( myFont != null )
		{
			myFont.setBold( b );
		}
	}

	public int getIdx()
	{
		return tableidx;
	}

	/**
	 * return truth of "this Xf rec is a style xf"
	 *
	 * @return
	 */
	public boolean isStyleXf()
	{
		return (fStyle == 1);
	}

	/**
	 * clone the xf and add to streamer
	 *
	 * @param xf
	 * @return
	 */
	private static Xf cloneXf( Xf xf, WorkBook wkbook )
	{
		Xf clone;
		if( xf.getIdx() > -1 )
		{    // it's in the wb already
			clone = new Xf( xf.ifnt, wkbook );
			byte[] xfbytes = xf.getBytesAt( 0, xf.getLength() - 4 );
			clone.setData( xfbytes );
			clone.init();
		}
		else
		{    // xf hasn't been added to wb yet, no need to clone
			clone = xf;
		}
		clone.fill = xf.fill;
		clone.setToCellXF();    // changes will not be seen if fstyle bit is set TODO: is this correct in all cases???
		clone.tableidx = wkbook.insertXf( clone );
		return clone;
	}

	/**
	 * if xf parameter doesn't exist, create; if it does, create a new xf based on it
	 *
	 * @param xf      original xf
	 * @param fontIdx font to link xf to
	 * @param wkbook
	 * @return new xf
	 */
	public static Xf updateXf( Xf xf, int fontIdx, WorkBook wkbook )
	{
		if( xf == null )
		{
			xf = new Xf( fontIdx, wkbook );
			xf.tableidx = wkbook.insertXf( xf );    // insert new xf into stream ...
			return xf;
		}
		xf = Xf.cloneXf( xf, wkbook );
		return xf;
	}

	/**
	 * set the OOXML fill for this xf
	 *
	 * @param f
	 */
	public void setFill( Fill f )
	{
		fill = (Fill) f.cloneElement();
		fls = (short) fill.getFillPatternInt();
		icvFore = (short) fill.getFgColorAsInt( getWorkBook().getTheme() );
		icvBack = (short) fill.getBgColorAsInt( getWorkBook().getTheme() );
	}

	/**
	 * return the OOXML fill for this xf, if any
	 */
	public Fill getFill()
	{
		return fill;
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	@Override
	public void close()
	{
		super.close();
		myFont.close();
		myFormat = null;
	}

}

