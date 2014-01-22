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
package com.extentech.ExtenXLS;

import com.extentech.toolkit.CompatibleBigDecimal;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.ResourceLoader;
import com.extentech.toolkit.StringTool;

import java.util.ArrayList;
import java.util.Date;
import java.util.IllegalFormatConversionException;
import java.util.StringTokenizer;

//import java.text.SimpleDateFormat;

/**
 * ExtenXLS helper methods. <br>
 * Contains helpful methods to ease use of the ExtenXLS toolkit. <br>
 * <p/>
 * "http://www.extentech.com">Extentech Inc.</a>
 *
 * @see ByteTools
 */

public class ExcelTools implements java.io.Serializable
{

	private static final long serialVersionUID = 7622857355626065370L;

	/**
	 * Formats a double in the standard ExtenXLS (General) format. Up to
	 * 99999999999 is expressed in standard notation. Above that is formatted in
	 * scientific notation
	 * <p/>
	 * In addition, Excel precision of 9 digits is maintained
	 * <p/>
	 * <p/>
	 * returns a number formatted in Excel's General format, (assuming a wide
	 * enough column width - see below) example:
	 * formatNumericNotation(1234567890123) returns "1.23457E+12"
	 * <p/>
	 * Information on NOTATION_STANDARD_EXCEL (i.e. Excel's General Format): //
	 * Excel will show as many decimal places that the text item has room for,
	 * it won't use a thousands separator, and if the // number can't fit, Excel
	 * uses a scientific number format. // RULES: // 1- Assuming the column is
	 * wide enough numbers will only be displayed in the scientific format when
	 * they contain more than 10 digits. // 2- If you enter a number into a cell
	 * and thre is not enough room to display all the digits // then the number
	 * will either be displayed in scientific format or will not be displayed at
	 * all, meaning that ##### will appear. // The exact precision of the
	 * scientific format will depend on the width of the actual cell.
	 *
	 * @param fpnum
	 * @return String formatted number
	 */
	public static String getNumberAsString( double fpnum )
	{
		// Ensure precision and number of digits ala Excel
		// double issues - use BigDecimal
		java.math.BigDecimal bd = new java.math.BigDecimal( fpnum );
		int scale = bd.scale();
		if( (Math.abs( fpnum ) > 0.000000001) && scale > 9 )
		{
			bd = bd.setScale( 9, java.math.RoundingMode.HALF_UP );
		}
		else if( scale > 9 )
		{
			bd = new java.math.BigDecimal( fpnum, new java.math.MathContext( 5, java.math.RoundingMode.HALF_UP ) );
		}
		bd = bd.stripTrailingZeros();
		String s = bd.toPlainString();
		int len = s.length();
		// If larger than 11 characters, truncate string
		if( len > 11 && fpnum > 0 || len > 12 )
		{ // must deal with exponents and such as well
			if( scale == 0 )
			{
				s = new java.math.BigDecimal( bd.toString(), new java.math.MathContext( 6, java.math.RoundingMode.HALF_UP ) ).toString();
			}
			else if( bd.toString().indexOf( "E" ) == -1 )
			{
				s = new java.math.BigDecimal( bd.toString(), new java.math.MathContext( 10, java.math.RoundingMode.HALF_UP ) ).toString();
				while( s.length() > 0 && s.charAt( s.length() - 1 ) == '0' )
				{
					s = s.substring( 0, s.length() - 1 );
				}
				if( s.endsWith( "." ) )
				{
					s = s.substring( 0, s.length() - 1 );
				}
			}
			else
			{ // 5 + E+XX + sign
				s = new java.math.BigDecimal( bd.toString(), new java.math.MathContext( 5, java.math.RoundingMode.HALF_UP ) ).toString();
			}
		}
		return s;
	}

	/**
	 * static version of getFormattedStringVal; given an object value, a
	 * valid Excel format pattern, return the formatted string value.
	 *
	 * @param Object  o
	 * @param String  pattern	if General or "" returns string value
	 * @param boolean isInteger	if General pattern, attempt to use integer value (rather than double)
	 */
	public static String getFormattedStringVal( Object o, String pattern/*, boolean isInteger*/ )
	{
		if( o == null )
		{
			o = "";
		}

		boolean isInteger = false;
		if( o instanceof Integer || (o instanceof Double && ((Double) o).intValue() == ((Double) o).doubleValue()) )
		{
			isInteger = true;
		}
		else
		{
			isInteger = false;
		}

		if( pattern == null || pattern.equals( "" ) || pattern.equalsIgnoreCase( "GENERAL" ) )
		{
			if( isInteger )
			{
				return String.valueOf( Double.valueOf( o.toString() ).intValue() );
			}
			else
			{    // general double numbers have default precision ...
				try
				{
					double d = new Double( o.toString() );
					return ExcelTools.getNumberAsString( Double.valueOf( o.toString() ) );    // handles default precision
				}
				catch( NumberFormatException e )
				{
				}
				return o.toString();
			}
		}
		else if( pattern.equals( "000-00-0000" ) )
		{ // special case for SSN format ... sigh ...
			try
			{
				new Double( o.toString() );    // if it can't be converted to a number, return original string (tis what excel does)
				String s = o.toString();
				while( s.length() < 9 )    // tis what excel does ...
				{
					s = '0' + s;
				}
				return s.substring( 0, 3 ) + "-" + s.substring( 3, 5 ) + "-" + s.substring( 5 );
			}
			catch( Exception e )
			{
				return o.toString();
			}
		}

		/** try to determine if the format is numeric (+currency) or date */
		boolean isNumeric = false, isDate = false, isString = false;

		/** excel formats can have up to 4 parts:  <positive>;<negative>;<zero>;<text> */
		String[] pats = pattern.split( ";" );    // assign the correct pattern according to double or string value

		String tester = StringTool.convertPatternExtractBracketedExpression( pats[0] );
		if( tester.matches( ".*(((y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}).*)+).*" ) )
		{
			isDate = true;
			pats[0] = tester; // ignore locale and other info for dates ...
		}
		if( !isDate )
		{
			int idx = pats.length - 1;    // default with string
			try
			{
				double d = new Double( o.toString() );
				isNumeric = true;
				if( d > 0 )        // 1st expression is for + numbers
				{
					idx = 0;
				}
				else if( pats.length > 1 && d < 0 )    // 2nd is for - numbers
				{
					idx = 1;
				}
				else if( pats.length > 2 && d == 0 )    // 3rd for 0
				{
					idx = 2;
				}
				pattern = StringTool.convertPatternFromExcelToStringFormatter( pats[idx],
				                                                               d < 0 );    // get correct format for String.format formatter
			}
			catch( NumberFormatException e )
			{    // 4th for text (non-numeric)
				if( pats.length > 3 )
				{
					idx = 3;
				}
				isString = true;
				pattern = StringTool.convertPatternFromExcelToStringFormatter( pats[idx],
				                                                               false );    // get correct format for String.format formatter
			}
		}
		else
		{
			pattern = pats[0];
			pattern = StringTool.convertDatePatternFromExcelToStringFormatter( pattern );    // get correct format for SimpleDateFormat
		}

		if( isString )
		{    // use string portion of format, if any
			try
			{
				return String.format( pattern, o );
			}
			catch( IllegalFormatConversionException e )
			{
				return o.toString();
			}
		}
		if( isNumeric )
		{
			try
			{
				double d = new Double( o.toString() );
				if( !Double.isNaN( d ) )
				{
					d = Math.abs( d );    // negative number intricacies have been handled in convertPattern method
					// ugly, but has to be done ...
					if( pattern.indexOf( "%%" ) != -1 ) // convert to percent
					{
						d *= 100;
					}
					// special case of "@" -- integers converted to doubles format incorrectly ...
					if( pattern.equals( "%s" ) )
					{
						return o.toString();
					}
					return String.format( pattern, d );
				}
			}
			catch( Exception e )
			{
			}
			return o.toString();
		}
		if( isDate )
		{
			try
			{
				WorkBookHandle.simpledateformat.applyPattern( pattern );
			}
			catch( Exception ex )
			{
				return o.toString();
			}
			try
			{
				return WorkBookHandle.simpledateformat.format( DateConverter.getCalendarFromNumber( o ).getTime() );
/*// KSC: TESTING				
Date d= DateConverter.getCalendarFromNumber(o).getTime();				
return WorkBookHandle.simpledateformat.format(d);*/
			}
			catch( NumberFormatException e )
			{
				try
				{
					return WorkBookHandle.simpledateformat.format( new Date( o.toString() ).getTime() );
				}
				catch( IllegalArgumentException i )
				{
					if( o instanceof Number )
					{
						Logger.logWarn( "Unable to format date in " + pattern );
					}
				}
			}
			catch( IllegalArgumentException e )
			{
				if( o instanceof Number )
				{
					Logger.logWarn( "Unable to format date in " + pattern );
				}
			}
		}
		// otherwise
		return o.toString();
	}

	/**
	 * A FAIL FAST implementation for finding whether a cell string address
	 * falls within a set of row/col range coordinates.
	 * <p/>
	 * Sep 21, 2010
	 *
	 * @param rng      the range you want to test
	 * @param rowFirst in the target range
	 * @param rowLast  in the target range
	 * @param colFirst in the target range
	 * @param colLast  in the target range
	 * @return
	 */
	public static boolean isInRange( String rng, int rowFirst, int rowLast, int colFirst, int colLast )
	{
		int[] sh = com.extentech.ExtenXLS.ExcelTools.getRowColFromString( rng );

		// the guantlet
		if( sh[1] < colFirst )
		{
			return false;
		}
		if( sh[1] > colLast )
		{
			return false;
		}
		if( sh[0] < rowFirst )
		{
			return false;
		}
		if( sh[0] > rowLast )
		{
			return false;
		}

		return true; // passes!

	}

	/**
	 * returns true if range intersects with range2
	 *
	 * @param rng
	 * @param rc
	 * @return
	 */
	public static boolean intersects( String rng, int[] rc )
	{
		int[] rc2 = ExcelTools.getRangeCoords( rng );
		if( (rc[0] >= rc2[0]) && (rc[2] <= rc2[2]) && (rc[1] >= rc2[1]) && (rc[3] <= rc2[3]) )
		{
			return true;
		}
		return false;
	}

	/**
	 * returns true if address is before the range coordinates defined by rc
	 *
	 * @param rc  row col of address
	 * @param rng int[] coordinates as: row0, col0, row1, col1
	 * @return true if address is before the range coordinates
	 */
	public static boolean isBeforeRange( int[] rc, int[] rng )
	{
		if( rc[0] < rng[0] || (rc[0] == rng[0] && rc[1] < rng[1]) )
		{
			return true;
		}
		return false;
	}

	/**
	 * returns true if address is before the range coordinates defined by rc
	 *
	 * @param rc  row col of address
	 * @param rng int[] coordinates as: row0, col0, row1, col1
	 * @return true if address is before the range coordinates
	 */
	public static boolean isAfterRange( int[] rc, int[] rng )
	{
		if( rc[0] > rng[2] || (rc[0] == rng[2] && rc[1] > rng[3]) )
		{
			return true;
		}
		return false;
	}

	/**
	 * Takes an input Object and attempts to convert to numeric Objects of the
	 * highest precision possible.
	 * <p/>
	 * This method is useful for avoiding the Excel warnings
	 * "Number Stored As Text" when storing string data that contains numbers.
	 * <p/>
	 * NOTE: this method is useful for ensuring that Formula references contain
	 * true numeric values as not all String numbers are properly interpreted in
	 * Formula engines, and can silently fail.
	 * <p/>
	 * For this reason, always use numeric, non-string values to calculated
	 * cells.
	 *
	 * @param input
	 * @return
	 */
	public static Object getObject( Object in )
	{
		// do not record -- only called from other methods
		if( !(in instanceof String) )
		{
			return in;
		}

		String input = String.valueOf( in );
		Object ret = input; // default is the original string

		try
		{
			ret = new Double( input );
			return ret;
		}
		catch( NumberFormatException ex )
		{
			try
			{
				ret = new Float( input );
				return ret;
			}
			catch( NumberFormatException ex2 )
			{
				try
				{
					ret = Integer.valueOf( input );
					return ret;
				}
				catch( NumberFormatException ex3 )
				{
					// ret Is set outside of loop incase no match is ever found.
				}
			}
		}

		// list of formatting chars to check for
		String[][] fmtlist = { { "$", "," }, { ",", "," }, { "%", "," } };

		// strip the formatting
		for( int t = 0; t < fmtlist.length; t++ )
		{
			if( input.indexOf( fmtlist[t][0] ) > -1 )
			{ // contains!
				String converted = StringTool.replaceText( input, fmtlist[t][0], "" ); // strip first token (ie: '$')
				converted = StringTool.replaceText( converted, fmtlist[t][1], "" ); // strip
				// second
				// token
				// (ie: ',')
				try
				{
					ret = new Double( converted );
					return ret;
				}
				catch( NumberFormatException ex )
				{
					try
					{
						ret = new Float( converted );
						return ret;
					}
					catch( NumberFormatException ex2 )
					{
						try
						{
							ret = Integer.valueOf( converted );
							return ret;
						}
						catch( NumberFormatException ex3 )
						{
							// ret Is set outside of loop incase no match is
							// ever found.
						}
					}
				}
			}

		}

		return ret;
	}

	/**
	 * convert twips to pixels
	 * <p/>
	 * <p/>
	 * In addition to a calculated size unit derived from the average size of
	 * the default characters 0-9, Excel uses the 'twips' measurement which is
	 * defined as:
	 * <p/>
	 * 1 twip = 1/20 point or 20 twips = 1 point 1 twip = 1/567 centimeter or
	 * 567 twips = 1 centimeter 1 twip = 1/1440 inch or 1440 twips = 1 inch
	 * <p/>
	 * <p/>
	 * 1 pixel = 0.75 points 1 pixel * 1.3333 = 1 point 1 twip * 20 = 1 point
	 *
	 * @param pixels
	 * @return twips
	 */
	public static final float getPixels( float twips )
	{
		float points = twips / 20;
		float pixels = points * 1.3333333f; // good enuff precision
		return pixels;
	}

	/**
	 * convert pixels to twips
	 * <p/>
	 * <p/>
	 * In addition to a calculated size unit derived from the average size of
	 * the default characters 0-9, Excel uses the 'twips' measurement which is
	 * defined as:
	 * <p/>
	 * 1 pixel = 0.75 points 1 pixel * 1.3333 = 1 point 1 twip * 20 = 1 point
	 *
	 * @param pixels
	 * @return twips
	 */
	public static final float getTwips( float pixels )
	{
		float points = pixels * .75f;
		float twips = points * 20;
		return twips;
	}

	/**
	 * get recordy byte def as a String
	 * <p/>
	 * public static String getRecordByteDef(XLSRecord rec){ byte[] b =
	 * rec.read(); StringBuffer sb = new StringBuffer("byte[] rbytes = {");
	 * for(int t = 0;t<b.length;t++){
	 * <p/>
	 * Byte thisb = new Byte(b[t]);
	 * <p/>
	 * sb.append(thisb.toString() + ", "); } sb.append("};"); return
	 * sb.toString(); }
	 */

	public static String getLogDate()
	{
		return String.valueOf( new Date( System.currentTimeMillis() ) );
	}

	/**
	 * tracks minimal info container for counters -> start time, last time,
	 * start mem, last mem
	 *
	 * @param info
	 * @param perfobj
	 */
	public static void benchmark( String info, Object perfobj )
	{
		Runtime rt = Runtime.getRuntime();
		long[] p = null;
		long lasttime = 0l, lastmem = 0l;
		if( System.getProperties().get( perfobj.toString() ) != null )
		{
			p = (long[]) System.getProperties().get( perfobj.toString() );
			lasttime = p[1];
			lastmem = p[3];
			p[1] = System.currentTimeMillis();
			p[3] = rt.freeMemory();
			double elapsedsec = p[1] - lasttime;
			double usedmem = lastmem - p[3]; // - lastmem;
			if( usedmem < 0 )
			{
				usedmem *= -1;
			}
			Logger.logInfo( getLogDate() + " " + info );
			Logger.logInfo( " time: " + elapsedsec + " millis" );
			Logger.logInfo( " mem: " + usedmem + " bytes." );
		}
		else
		{
			p = new long[4];
			p[0] = System.currentTimeMillis();
			p[1] = System.currentTimeMillis();
			p[2] = rt.freeMemory();
			p[3] = rt.freeMemory();
			lasttime = p[1];
			lastmem = p[3];
			System.getProperties().put( perfobj.toString(), p );
		}

	}

	/**
	 * get the bytes from a Vector of objects public byte[]
	 * getBytesFrom(CompatibleVector objs){
	 * <p/>
	 * for(int t = 0;t<objs.size();t++){
	 * <p/>
	 * }
	 * <p/>
	 * <p/>
	 * }
	 */

	static char[] alpharr = {
			'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'
	};
	/*
	 * NOT USED ANYMORE -- see getAlphaVal -- OK to remove??
	 */
	public static final String[] ALPHASDELETE = {
			"A",
			"B",
			"C",
			"D",
			"E",
			"F",
			"G",
			"H",
			"I",
			"J",
			"K",
			"L",
			"M",
			"N",
			"O",
			"P",
			"Q",
			"R",
			"S",
			"T",
			"U",
			"V",
			"W",
			"X",
			"Y",
			"Z",
			"AA",
			"AB",
			"AC",
			"AD",
			"AE",
			"AF",
			"AG",
			"AH",
			"AI",
			"AJ",
			"AK",
			"AL",
			"AM",
			"AN",
			"AO",
			"AP",
			"AQ",
			"AR",
			"AS",
			"AT",
			"AU",
			"AV",
			"AW",
			"AX",
			"AY",
			"AZ",
			"BA",
			"BB",
			"BC",
			"BD",
			"BE",
			"BF",
			"BG",
			"BH",
			"BI",
			"BJ",
			"BK",
			"BL",
			"BM",
			"BN",
			"BO",
			"BP",
			"BQ",
			"BR",
			"BS",
			"BT",
			"BU",
			"BV",
			"BW",
			"BX",
			"BY",
			"BZ",
			"CA",
			"CB",
			"CC",
			"CD",
			"CE",
			"CF",
			"CG",
			"CH",
			"CI",
			"CJ",
			"CK",
			"CL",
			"CM",
			"CN",
			"CO",
			"CP",
			"CQ",
			"CR",
			"CS",
			"CT",
			"CU",
			"CV",
			"CW",
			"CX",
			"CY",
			"CZ",
			"DA",
			"DB",
			"DC",
			"DD",
			"DE",
			"DF",
			"DG",
			"DH",
			"DI",
			"DJ",
			"DK",
			"DL",
			"DM",
			"DN",
			"DO",
			"DP",
			"DQ",
			"DR",
			"DS",
			"DT",
			"DU",
			"DV",
			"DW",
			"DX",
			"DY",
			"DZ",
			"EA",
			"EB",
			"EC",
			"ED",
			"EE",
			"EF",
			"EG",
			"EH",
			"EI",
			"EJ",
			"EK",
			"EL",
			"EM",
			"EN",
			"EO",
			"EP",
			"EQ",
			"ER",
			"ES",
			"ET",
			"EU",
			"EV",
			"EW",
			"EX",
			"EY",
			"EZ",
			"FA",
			"FB",
			"FC",
			"FD",
			"FE",
			"FF",
			"FG",
			"FH",
			"FI",
			"FJ",
			"FK",
			"FL",
			"FM",
			"FN",
			"FO",
			"FP",
			"FQ",
			"FR",
			"FS",
			"FT",
			"FU",
			"FV",
			"FW",
			"FX",
			"FY",
			"FZ",
			"GA",
			"GB",
			"GC",
			"GD",
			"GE",
			"GF",
			"GG",
			"GH",
			"GI",
			"GJ",
			"GK",
			"GL",
			"GM",
			"GN",
			"GO",
			"GP",
			"GQ",
			"GR",
			"GS",
			"GT",
			"GU",
			"GV",
			"GW",
			"GX",
			"GY",
			"GZ",
			"HA",
			"HB",
			"HC",
			"HD",
			"HE",
			"HF",
			"HG",
			"HH",
			"HI",
			"HJ",
			"HK",
			"HL",
			"HM",
			"HN",
			"HO",
			"HP",
			"HQ",
			"HR",
			"HS",
			"HT",
			"HU",
			"HV",
			"HW",
			"HX",
			"HY",
			"HZ",
			"IA",
			"IB",
			"IC",
			"ID",
			"IE",
			"IF",
			"IG",
			"IH",
			"II",
			"IJ",
			"IK",
			"IL",
			"IM",
			"IN",
			"IO",
			"IP",
			"IQ",
			"IR",
			"IS",
			"IT",
			"IU",
			"IV",
			"IW",
			"IX",
			"IY",
			"IZ",
	};

	/**
	 * get the Excel-style Column alphabetical representation of an integer
	 * (0-based).
	 * <p/>
	 * for example: 0 = A 26= AA 701= ZZ 702= AAA 16383= XFD (max)
	 */
	public static String getAlphaVal( int i )
	{
		String ret = "";
		int leftover = 0;
		if( i > 701 )
		{ // has 3rd digit
			int z = (i / 676) - 1; // -1 to account for 0-based
			if( (i % 676) < 26 )
			{ // then "leftover" is actually 2nd digit
				z--;
				leftover = 676;
			}
			ret = String.valueOf( ExcelTools.alpharr[z] );
			i = i % 676;
			i += leftover;
		}
		if( i > 25 )
		{ // has 2nd digit
			int z = (i / 26) - 1; // -1 to account for 0-based
			ret = ret + ExcelTools.alpharr[z];
			i = i % 26;
		}

		// bbennett: this raises AIOOB if i = -1. In this situation we don't
		// care about alpha because it will just be in the message of the
		// exception that flags cell as non existent.
		ret += i < 0 ? Integer.toString( i ) : String.valueOf( ExcelTools.alpharr[(i)] );

		return ret;
	}

	/**
	 * get the int value of the Excel-style Column alpha representation.
	 *
	 * @param String column name
	 * @return int the 0-based column number
	 */
	public static int getIntVal( String c )
	{
		c = c.toUpperCase();
		if( c.length() > 3 ) // max col value= XFD in Excel 2007
		{
			return -1;
		}
		int i = c.length() - 1;
		int ret = 0;
		while( i >= 0 )
		{ // process least to most-sigificant dig
			int z = 0;
			char cc = c.charAt( i );
			while( z < alpharr.length && cc != alpharr[z++] )
			{
				;
			}
			z *= Math.pow( 26, c.length() - i - 1 ); // 1-based col for computing
			ret += z;
			i--;
		}
		// make 0-based
		ret--;
		return ret;
	}

	/**
	 * Parses an Excel cell address into row and column integers.
	 *
	 * @param address the address to parse, either A1 or R1C1
	 * @return int[2]: [0] row index, [1] column index
	 * @throws IllegalArgumentException if the argument is not a valid address
	 */
	public static int[] getRowColFromString( String address )
	{

		if( address.indexOf( "$" ) > -1 )
		{
			address = StringTool.strip( address, "$" );
		}
		if( address.indexOf( "!" ) > -1 )
		{
			address = address.substring( address.indexOf( "!" ) + 1 );
		}
		if( address.indexOf( ":" ) > -1 )
		{
			return getRangeRowCol( address );
		}

		char[] adrchars = address.toCharArray();
		int row = 0, col = 0;
		int charpos = -1, numpos = -1;
		boolean r1c1 = false;
		for( int i = 0; i < adrchars.length; i++ )
		{
			if( Character.isDigit( adrchars[i] ) )
			{
				if( numpos == -1 ) // its a number
				{
					numpos = i;
				}
			}
			else if( charpos == -1 )
			{
				charpos = i;
				if( numpos >= 0 )
				{ // we have already set number and we now have
					// a nondigit - R1C1 style
					r1c1 = true;
					// break, it's all over!
					break;
				}
			}
		}

		if( r1c1 )
		{ // it's a single cell ref
			try
			{
				// there's an R and a C, not adjacent
				if( address.toUpperCase().indexOf( "R" ) == 0 )
				{ // startwith R
					String rx = address.substring( 1, address.toUpperCase().indexOf( "C" ) );
					String cx = address.substring( address.toUpperCase().indexOf( "C" ) + 1 );
					row = Integer.parseInt( rx );
					col = Integer.parseInt( cx );
				}
			}
			catch( NumberFormatException e )
			{
				throw new IllegalArgumentException( "illegal R1C1 address '" + address + "'" );
			}
		}
		else
		{
			row = 0; // -1 below
			col = -1;
			if( charpos == 0 && numpos > 0 )
			{
				String colval = address.substring( 0, numpos );
				col = getIntVal( colval );
				if( col < 0 )
				{
					throw new IllegalArgumentException( "illegal column value '" + colval + "' in address '" + address + "'" );
				}
			}
			if( numpos >= 0 )
			{
				row = Integer.parseInt( address.substring( numpos ) );
				if( row < 1 )
				{
					throw new IllegalArgumentException( "row may not be negative in address '" + address + "'" );
				}
			}
			else
			{ // it's a wholecol ref
				col = getIntVal( address );
				if( col < 0 )
				{
					throw new IllegalArgumentException( "illegal column value '" + address + "' in address '" + address + "'" );
				}
			}
		}

		int[] ret = { row - 1, col };
		return ret;
	}

	/**
	 * Parses an Excel cell range and returns the addresses as an int array. The
	 * range may not be qualified with sheet names. Strip them with
	 * {@link #stripSheetNameFromRange} before calling this method. If the
	 * argument is a single cell address it will be returned for both bounds.
	 *
	 * @param range the range to parse
	 * @return int[4]: [0] first row, [1] first column, [2] second row, [3]
	 * second column
	 * @throws IllegalArgumentException if the addresses are invalid
	 */
	public static int[] getRangeRowCol( String range )
	{
		int colon = range.indexOf( ":" );

		String firstloc;
		String lastloc;

		if( colon > -1 )
		{
			firstloc = range.substring( 0, colon );
			lastloc = range.substring( colon + 1 );
		}
		else
		{
			firstloc = range;
			lastloc = range;
		}

		int[] result = new int[4];
		int[] temp;

		temp = getRowColFromString( firstloc );
		System.arraycopy( temp, 0, result, 0, 2 );

		temp = getRowColFromString( lastloc );
		System.arraycopy( temp, 0, result, 2, 2 );

		return result;
	}

	/**
	 * Takes an int array representing a row and column and formats it as a cell
	 * address.
	 * <p/>
	 * The index is zero-based.
	 * <p/>
	 * [0][0] is "A1" [1][1] is "B2" [2][2] is "C3"
	 *
	 * @param int[] the numeric range to convert
	 * @return String the string representation of the range
	 */
	public static String formatLocation( int[] rowCol )
	{
		StringBuffer sb = new StringBuffer( getAlphaVal( rowCol[1] ) );
		sb.append( String.valueOf( rowCol[0] + 1 ) );

		// handle ranges
		if( rowCol.length > 3 )
		{
			// 20090807: KSC: only a range if 1st is != 2nd cell :)
			if( rowCol[0] == rowCol[2] && rowCol[1] == rowCol[3] ) // it's a single address
			{
				return sb.toString();
			}
			sb.append( ":" );
			sb.append( getAlphaVal( rowCol[3] ) );
			sb.append( String.valueOf( rowCol[2] + 1 ) );
		}

		return sb.toString();
	}

	/**
	 * Takes an int array representing a row and column and formats it as a cell
	 * address, taking into account relative or absolute refs
	 * <p/>
	 * The index is zero-based.
	 * <p/>
	 * [0][0] is "A1", $A1, A$1 or $A$1 depending upon bRelRow or bRelCol [1][1]
	 * is "B2", $B1, B$1 or B$1 depending upon bRelRow or bRelCol [2][2] is
	 * "C3", $C1, C$1 or $C$1 depending upon bRelRow or bRelCol
	 *
	 * @param int[]   the numeric range to convert
	 * @param bRelRow if true, no "$"s are added, relative row reference
	 * @param bRelCol if true, no "$"s are added, relative col reference
	 * @return String the string representation of the range
	 */
	public static String formatLocation( int[] s, boolean bRelRow, boolean bRelCol )
	{
		StringBuffer sb = new StringBuffer( (bRelCol ? "" : "$") );
		if( s[1] > -1 ) // account for WholeRow/WholeCol references
		{
			sb.append( getAlphaVal( s[1] ) );
		}

		if( s[0] > -1 ) // account for WholeRow/WholeCol references
		{
			sb.append( (bRelRow ? "" : "$") + String.valueOf( s[0] + 1 ) );
		}
		// 20090906 KSC: handle ranges
		if( s.length > 3 )
		{
			if( s[0] == s[2] && s[1] == s[3] ) // it's a single address
			{
				return sb.toString();
			}
			sb.append( ":" );
			sb.append( (bRelCol ? "" : "$") );
			sb.append( getAlphaVal( s[3] ) );
			sb.append( (bRelRow ? "" : "$") + String.valueOf( s[2] + 1 ) );
		}

		return sb.toString();
	}

	/**
	 * Takes an array of four shorts and formats it as a cell range.
	 * <p/>
	 * IE [0][3][1][4] would be "A2:B3"
	 *
	 * @param int[] the numeric range to convert
	 * @return String the string representation of the range
	 */
	public static String formatRange( int[] s )
	{
		if( s.length != 4 )
		{
			return "incorrect array size in ExcelTools.formatLocation";
		}
		int[] temp = new int[2];
		temp[0] = s[1];
		temp[1] = s[0];
		String firstcell = formatLocation( temp );
		temp[0] = s[3];
		temp[1] = s[2];
		String lastcell = formatLocation( temp );
		return firstcell + ":" + lastcell;
	}

	/**
	 * format a range as a string, range in format of [r][c][r1][c1]
	 *
	 * @param s
	 * @return String representation of the integers as a range, ie A1:B4
	 */
	public static String formatRangeRowCol( int[] s )
	{
		if( s.length != 4 )
		{
			return "incorrect array size in ExcelTools.formatLocation";
		}
		int[] temp = new int[2];
		temp[0] = s[0];
		temp[1] = s[1];
		String firstcell = formatLocation( temp );
		temp[0] = s[2];
		temp[1] = s[3];
		String lastcell = formatLocation( temp );
		return firstcell + ":" + lastcell;
	}

	/**
	 * format a range as a string, range in format of [r][c][r1][c1] including
	 * relative address state
	 *
	 * @param s
	 * @param bRelAddresses contains relative row and col state for each rcr1c1
	 * @return String representation of the integers as a range, ie A1:B4
	 */
	public static String formatRangeRowCol( int[] s, boolean[] bRelAddresses )
	{
		if( s.length != 4 )
		{
			return "incorrect array size in ExcelTools.formatLocation";
		}
		int[] temp = new int[2];
		temp[0] = s[0];
		temp[1] = s[1];
		String firstcell = formatLocation( temp, bRelAddresses[0], bRelAddresses[1] );
		temp[0] = s[2];
		temp[1] = s[3];
		String lastcell = formatLocation( temp, bRelAddresses[2], bRelAddresses[3] );
		// unfortunately, no formatLocation can do this. This is
		// formattingRANGERowCol
		// if (firstcell.equals(lastcell)) return firstcell; // 20090309 KSC:
		return firstcell + ":" + lastcell;
	}

	/**
	 * Transforms a string to an array of ints for evaluation purposes. For
	 * example, acdc == [0][2][3][2]
	 */
	public static int[] transformStringToIntVals( String trans )
	{
		int[] intarr = new int[trans.length()];
		for( int i = 0; i < trans.length(); i++ )
		{
			char c = trans.charAt( i );
			for( int x = 0; x < alpharr.length; x++ )
			{
				if( String.valueOf( c ).equalsIgnoreCase( String.valueOf( alpharr[x] ) ) )
				{
					intarr[i] = x;
				}
			}
		}
		return intarr;
	}

	/**
	 * Formats a string representation of a numeric value as a string in the
	 * specified notation:
	 *
	 * @param int <br>
	 *            NOTATION_STANDARD = 0, <br>
	 *            NOTATION_SCIENTIFIC = 1, <br>
	 *            NOTATION_SCIENTIFIC_EXCEL = 2, <br>
	 *            EXTENXLS_NOTATION = 3
	 *            <p/>
	 *            <p/>
	 *            example: formatNumericNotation(1.23456E5, 0) returns a "123456"
	 *            example: formatNumericNotation(123456, 1) returns "1.23456E5"
	 *            example: formatNumericNotation(123456, 2) returns "1.23456E+5"
	 *            example: formatNumericNotation(123456, 3) returns "1.23456E+5"
	 */
	public static String formatNumericNotation( String num, int notationType )
	{
		// if (notationType > 2)return null;
		boolean negative = false;
		if( num.substring( 0, 1 ).equals( "-" ) )
		{
			negative = true;
			num = num.substring( 1, num.length() );
		}
		String preString, postString, fullString = "";
		switch( notationType )
		{
			case 0: // NOTATION_STANDARD
				int i = num.indexOf( "E" );
				if( i == -1 )
				{ // just return
					if( num.substring( num.length() - 2, num.length() ).equals( ".0" ) )
					{
						num = num.substring( 0, num.length() - 2 );
					}
					if( negative )
					{
						return "-" + num;
					}
					return num;
				}
				preString = num.substring( 0, i );
				CompatibleBigDecimal outNumD = new CompatibleBigDecimal( preString );
				String exp = "";
				if( num.indexOf( "+" ) == -1 )
				{
					exp = num.substring( i + 1, num.length() );
				}
				else
				{
					exp = num.substring( i + 2, num.length() );
				}
				int expNum = Integer.valueOf( exp ).intValue();
				outNumD = new CompatibleBigDecimal( outNumD.movePointRight( expNum ) );
				// outNumD = outNumD.multiply(new CompatibleBigDecimal(Math.pow(10,
				// expNum)));
				// Logger.logInfo(String.valueOf(outNumD));
				// outNum = Math.r

				// check if we should be returning a whole number or a decimal
				int moveLen = num.indexOf( "E" ) - num.indexOf( "." ) - 1;
				if( expNum >= moveLen )
				{
					if( negative )
					{
						return "-" + String.valueOf( Math.round( outNumD.doubleValue() ) );
					}
					else
					{
						return String.valueOf( Math.round( outNumD.doubleValue() ) );
					}
				}
				Object[] args = new Object[0];
				// args[0] = outNumD;
				Object res = ResourceLoader.executeIfSupported( outNumD, args, "toPlainString" );
				if( res != null )
				{
					fullString = res.toString();
				}
				else
				{
					fullString = outNumD.toCompatibleString();
				}
				break;
			case 1: // NOTATION_SCIENTIFIC
				if( num.indexOf( "E" ) != -1 && num.indexOf( "+" ) == -1 )
				{
					fullString = num;
				}
				else if( num.indexOf( "+" ) != -1 )
				{
					preString = num.substring( 0, num.indexOf( "+" ) );
					postString = num.substring( num.indexOf( "+" ) + 1, num.length() );
					return preString + postString;
				}
				else if( num.indexOf( "." ) != -1 )
				{
					int pos = num.indexOf( "." );
					preString = num.substring( 0, 1 ) + "." + num.substring( 1, num.indexOf( "." ) );
					CompatibleBigDecimal d = new CompatibleBigDecimal( num );
					if( d.doubleValue() < 1 && d.doubleValue() != 0 )
					{
						// it is a very small value, ie 1.0E-10
						int counter = 0;
						while( d.doubleValue() < 1 )
						{
							d = new CompatibleBigDecimal( d.movePointRight( 1 ) );
							counter++;
						}
						String retStr = d.toCompatibleString() + "E-" + counter;
						return retStr;
					}
					postString = num.substring( num.indexOf( "." ) + 1, num.length() );
					fullString = preString + postString;
					fullString = fullString + "E" + (pos - 1);
				}
				else
				{
					preString = num.substring( 0, 1 ) + ".";
					if( num.length() > 1 )
					{
						preString += num.substring( 1, num.length() );
					}
					else
					{
						preString += "0";
					}
					fullString = preString + "E" + (num.length() - 1);
				}
				break;
			case 2: // NOTATION_SCIENTIFIC_EXCEL
				if( num.indexOf( "E" ) != -1 && num.indexOf( "+" ) != -1 )
				{
					fullString = num;
				}
				else if( num.indexOf( "E" ) != -1 )
				{
					preString = num.substring( 0, num.indexOf( "E" ) + 1 );
					postString = "+" + num.substring( num.indexOf( "E" ) + 1, num.length() );
					fullString = preString + postString;
				}
				else if( num.indexOf( "." ) != -1 )
				{
					int pos = num.indexOf( "." );
					CompatibleBigDecimal d = new CompatibleBigDecimal( num );
					if( d.doubleValue() < 1 && d.doubleValue() != 0 )
					{
						// it is a very small value, ie 1.0E-10
						int counter = 0;
						while( d.doubleValue() < 1 )
						{
							d = new CompatibleBigDecimal( d.movePointRight( 1 ) );
							counter++;
						}
						String retStr = d.toCompatibleString() + "E-" + counter;
						return retStr;
					}
					preString = num.substring( 0, 1 ) + "." + num.substring( 1, num.indexOf( "." ) );
					postString = num.substring( num.indexOf( "." ) + 1, num.length() );
					fullString = preString + postString;
					fullString = fullString + "E+" + (pos - 1);
				}
				else
				{
					preString = num.substring( 0, 1 ) + ".";
					if( num.length() > 1 )
					{
						preString += num.substring( 1, num.length() );
					}
					else
					{
						preString += "0";
					}
					fullString = preString + "E+" + (num.length() - 1);
				}
				break;
			default:
				return num;
		}
		if( negative )
		{
			fullString = "-" + fullString;
		}
		return fullString;
	}

	/**
	 * Return an array of cell handles specified from the string passed in.
	 * <p/>
	 * Note that a CellHandle cannot exist for an empty cell, so the cells
	 * retrieved in this manner will be blank cells, not empty cells.
	 *
	 * @param cellstr - a comma delimited String representing cells and cell ranges,
	 *                example "A1,A5,A6,B1:B5" would return cells A1, A5, A6, B1,
	 *                B2, B3, B4, B5
	 * @param sheet   the worksheet containing the cells.
	 * @return CellHandle[]
	 */
	public static CellHandle[] getCellHandlesFromSheet( String strRange, WorkSheetHandle sheet )
	{
		CellHandle[] retCells;
		StringTokenizer cellTokenizer = new StringTokenizer( strRange, "," );
		ArrayList cells = new ArrayList();
		do
		{
			String element = (String) cellTokenizer.nextElement();
			if( element.indexOf( ":" ) != -1 )
			{
				CellRange aRange = new CellRange( sheet.getSheetName() + "!" + strRange, sheet.wbh, true );
				cells.addAll( aRange.getCellList() );
			}
			else
			{
				CellHandle aCell = null;
				try
				{
					aCell = sheet.getCell( element );
				}
				catch( Exception ce )
				{
					aCell = sheet.add( null, element );
				}
				if( aCell != null )
				{
					cells.add( aCell );
				}
			}
		} while( cellTokenizer.hasMoreElements() );
		retCells = new CellHandle[cells.size()];
		retCells = (CellHandle[]) cells.toArray( retCells );
		return retCells;
	}

	/**
	 * Strip sheet name(s) from range string can be Sheet1!AB:Sheet!BC or
	 * Sheet!AB:BC or AB:BC or Sheet1:Sheet2!A1:A2
	 *
	 * @param address or range String
	 * @return 1st sheetname
	 * <p/>
	 * Ok, this is a strange method. It returns a string array of the
	 * following format 0 - sheetname1 1 - cell address or range (what
	 * if there are 2?) 2 - sheetname2 3 - external link 1 ?? some ooxml
	 * record 4 - external link 2
	 */
	public static String[] stripSheetNameFromRange( String address )
	{
		String sheetname = null, sheetname2 = null;
		int m = address.indexOf( '!' );
		if( m > -1 )
		{
			if( address.substring( 0, m ).indexOf( ":" ) == -1 )
			{
				sheetname = address.substring( 0, m );
			}
			else
			{
				int z = address.indexOf( ":" );
				sheetname = address.substring( 0, z );
				sheetname2 = address.substring( z + 1, m );
			}
		}
		address = address.substring( m + 1 );
		int n = address.indexOf( '!' ); // see if 2nd sheet name exists
		if( n > -1 && !address.equals( "#REF!" ) )
		{
			m = address.indexOf( ':' );
			sheetname2 = address.substring( m + 1, n );
			m = address.indexOf( ':' );
			address = address.substring( 0, m + 1 ) + address.substring( n + 1 );
		}
		// 20090323 KSC: handle external references (OOXML-Specific format of
		// [#]SheetName!Ref where # denotes ExternalLink workbook
		String exLink1 = null, exLink2 = null;
		if( sheetname != null && sheetname.indexOf( '[' ) >= 0 )
		{ // External
			// OOXML
			// reference
			exLink1 = sheetname.substring( sheetname.indexOf( '[' ) );
			exLink1 = exLink1.substring( 0, exLink1.indexOf( ']' ) + 1 );
			sheetname = StringTool.replaceText( sheetname, exLink1, "" );
			if( sheetname.equals( "" ) )
			{
				sheetname = null; // possible to have address in form of =
			}
			// [#]!Name or range
		}
		if( sheetname2 != null && sheetname2.indexOf( '[' ) >= 0 )
		{ // External
			// OOXML
			// reference
			exLink2 = sheetname2.substring( sheetname2.indexOf( '[' ) );
			exLink2 = exLink2.substring( 0, exLink2.indexOf( ']' ) + 1 );
			sheetname2 = StringTool.replaceText( sheetname2, exLink2, "" );
			if( sheetname2.equals( "" ) )
			{
				sheetname2 = null; // possible to have address in form of =
			}
			// [#]!Name or range
		}
		// return new String[]{sheetname, address, sheetname2};
		return new String[]{ sheetname, address, sheetname2, exLink1, exLink2 }; // 20090323
		// KSC:
		// add
		// any
		// external
		// link
		// info
	}

	/**
	 * return the first and last coords of a range in int form + the number of
	 * cells in the range range is in the format of Sheet
	 */
	public static int[] getRangeCoords( String range )
	{
		int numrows = 0;
		int numcols = 0;
		int numcells = 0;
		int[] coords = new int[5];
		String temprange = range;
		// figure out the sheet bounds using the range string
		temprange = ExcelTools.stripSheetNameFromRange( temprange )[1];
		String startcell = "", endcell = "";
		int lastcolon = temprange.lastIndexOf( ":" );
		endcell = temprange.substring( lastcolon + 1 );
		if( lastcolon == -1 ) // no range
		{
			startcell = endcell;
		}
		else
		{
			startcell = temprange.substring( 0, lastcolon );
		}
		startcell = StringTool.strip( startcell, "$" );
		endcell = StringTool.strip( endcell, "$" );

		// get the first cell's coordinates
		int charct = startcell.length();
		while( charct > 0 )
		{
			if( !Character.isDigit( startcell.charAt( --charct ) ) )
			{
				charct++;
				break;
			}
		}
		String firstcellrowstr = startcell.substring( charct );
		int firstcellrow = -1;
		try
		{
			firstcellrow = Integer.parseInt( firstcellrowstr );
		}
		catch( NumberFormatException e )
		{ // could be a whole-col-style ref
		}
		String firstcellcolstr = startcell.substring( 0, charct ).trim();
		int firstcellcol = ExcelTools.getIntVal( firstcellcolstr );
		// get the last cell's coordinates
		charct = endcell.length();
		while( charct > 0 )
		{
			if( !Character.isDigit( endcell.charAt( --charct ) ) )
			{
				charct++;
				break;
			}
		}
		String lastcellrowstr = endcell.substring( charct );
		int lastcellrow = -1;
		try
		{
			lastcellrow = Integer.parseInt( lastcellrowstr );
		}
		catch( NumberFormatException e )
		{ // could be a whole-col-style ref
		}
		String lastcellcolstr = endcell.substring( 0, charct );
		int lastcellcol = ExcelTools.getIntVal( lastcellcolstr );
		numrows = (lastcellrow - firstcellrow) + 1;
		numcols = (lastcellcol - firstcellcol) + 1;
		/*
		 * if(numrows == 0)numrows =1; if(numcols == 0)numcols =1;
		 */
		numcells = numrows * numcols;
		if( numcells < 0 )
		{
			numcells *= -1; // handle swapped cells ie: "B1:A1"
		}

		coords[0] = firstcellrow;
		coords[1] = firstcellcol;
		coords[2] = lastcellrow;
		coords[3] = lastcellcol;
		coords[4] = numcells;
		// Trap errors in range
		// if (firstcellrow < 0 || lastcellrow < 0 || firstcellcol < 0 ||
		// lastcellcol < 0)
		// Logger.logErr("ExcelTools.getRangeCoords: Error in Range " + range);
		return coords;
	}
}