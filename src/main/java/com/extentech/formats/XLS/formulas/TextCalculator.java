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
package com.extentech.formats.XLS.formulas;

import com.extentech.ExtenXLS.DateConverter;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.FormatConstants;
import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.Xf;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;

import java.io.UnsupportedEncodingException;
import java.text.DecimalFormat;
import java.text.Format;
import java.util.Date;

/**
 * TextCalculator is a collection of static methods that operate
 * as the Microsoft Excel function calls do.
 * <p/>
 * All methods are called with an array of ptg's, which are then
 * acted upon.  A Ptg of the type that makes sense (ie boolean, number)
 * is returned after each method.
 */
public class TextCalculator
{

	/**
	 * ASC function
	 * For Double-byte character set (DBCS) languages, changes full-width (double-byte) characters to half-width (single-byte) characters.
	 * ASC(text)
	 * Text   is the text or a reference to a cell that contains the text you want to change. If text does not contain any full-width letters, text is not changed.
	 * NOTE: in order to use this and other DBCS Methods in Excel,
	 * the input language must be set to a DBCS language such as Japanese
	 * Otherwise, the ASC function does nothing (apparently)
	 */
	protected static Ptg calcAsc( Ptg[] operands )
	{
		if( (operands == null) || (operands[0] == null) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		// determine if Excel's language is set up for DBCS; if not, returns normal string
		com.extentech.formats.XLS.WorkBook bk = operands[0].getParentRec().getWorkBook();
		if( bk.defaultLanguageIsDBCS() )
		{ // otherwise just returns normal string
			byte[] strbytes = getUnicodeBytesFromOp( operands[0] );
			if( strbytes == null )
			{
				strbytes = (operands[0].getValue()).toString().getBytes();
			}
			try
			{
				return new PtgStr( new String( strbytes, XLSConstants.UNICODEENCODING ) );
			}
			catch( Exception e )
			{
			}
		}
		return new PtgStr( operands[0].getValue().toString() );
	}
	
	/*BAHTTEXT function
	Converts a number to Thai text and adds a suffix of "Baht."
	 */

	/**
	 * CHAR
	 * Returns the character specified by the code number
	 */
	protected static Ptg calcChar( Ptg[] operands )
	{
		Object o = operands[0].getValue();
		Byte s = new Byte( o.toString() );
		if( (s.intValue() > 255) || (s.intValue() < 1) )
		{
			return PtgCalculator.getError();
		}
		byte[] b = new byte[1];
		b[0] = s;
		String str = "";
		try
		{
			str = new String( b, XLSConstants.DEFAULTENCODING );
		}
		catch( UnsupportedEncodingException e )
		{
		}
		;
		return new PtgStr( str );
	}

	/**
	 * CLEAN
	 * Removes all nonprintable characters from text. Use CLEAN on text
	 * imported from other applications that contains characters that may
	 * not print with your operating system. For example, you can use
	 * CLEAN to remove some low-level computer code that is frequently
	 * at the beginning and end of data files and cannot be printed.
	 * <p/>
	 * Syntax
	 * <p/>
	 * CLEAN(text)
	 * <p/>
	 * Text   is any worksheet information from which you want to remove nonprintable characters.
	 * <p/>
	 * The CLEAN function was designed to remove the first 32 nonprinting characters in the 7-bit ASCII code (values 0 through 31) from text. In the Unicode character set (Unicode: A character encoding standard developed by the Unicode Consortium. By using more than one byte to represent each character, Unicode enables almost all of the written languages in the world to be represented by using a single character set.), there are additional nonprinting characters (values 127, 129, 141, 143, 144, and 157). By itself, the CLEAN function does not remove these additional nonprinting characters.
	 */
	protected static Ptg calcClean( Ptg[] operands )
	{
		String retString = "";
		try
		{
			Object o = operands[0].getValue();
			String s = o.toString();
			for( int i = 0; i < s.length(); i++ )
			{
				int c = s.charAt( i );
				if( c >= 32 )
				{
					retString += (char) c;
				}
			}
		}
		catch( Exception e )
		{
			;
		}
		return new PtgStr( retString );
	}

	/**
	 * CODE
	 * Returns a numeric code for the first character in a text string
	 */
	protected static Ptg calcCode( Ptg[] operands )
	{
		Object o = operands[0].getValue();
		String s = o.toString();
		byte[] b = null;
		try
		{
			b = s.getBytes( XLSConstants.DEFAULTENCODING );
		}
		catch( UnsupportedEncodingException e )
		{
		}
		;
		Integer i = (int) b[0];
		return new PtgInt( i );
	}

	/**
	 * CONCATENATE
	 * Joins several text items into one text item
	 */
	protected static Ptg calcConcatenate( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		String s = "";
		for( int i = 0; i < allops.length; i++ )
		{
			s += allops[i].getValue().toString();
		}
		Ptg str = new PtgStr( s );
		str.setParentRec( operands[0].getParentRec() );
		return str;
	}

	/**
	 * DOLLAR
	 * Converts a number to text, using currency format. Can
	 * have a separate operand to determine POP.
	 */
	protected static Ptg calcDollar( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		int pop = 0;
		if( operands.length > 1 )
		{
			pop = operands[1].getIntVal();
		}
		double d = operands[0].getDoubleVal();
		d = d * Math.pow( 10, pop );
		d = Math.round( d );
		d = d / Math.pow( 10, pop );
		String res = "$" + String.valueOf( d );
		return new PtgStr( res );
	}

	/**
	 * EXACT
	 * Checks to see if two text values are identical
	 */
	protected static Ptg calcExact( Ptg[] operands )
	{
		if( operands.length != 2 )
		{
			return PtgCalculator.getError();
		}
		String s1 = operands[0].getValue().toString();
		String s2 = operands[1].getValue().toString();
		if( s1.equals( s2 ) )
		{
			return new PtgBool( true );
		}
		return new PtgBool( false );
	}

	/**
	 * FIND
	 * Finds one text value within another (case-sensitive)
	 */
	protected static Ptg calcFind( Ptg[] operands )
	{
		String instring = "";
		String wholestr = "";
		int start = 0;
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		if( operands.length == 3 )
		{
			start = operands[2].getIntVal() - 1;
		}
		Object o = operands[0].getValue();
		Object oo = operands[1].getValue();

		if( (o == null) || (oo == null) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}

		instring = o.toString();
		wholestr = oo.toString();
		// note this uses a starting position to search for the string,
		// but does not account for that starting position in respects to it's
		// result.  Pretty strange
		int i = wholestr.indexOf( instring, start );
		if( i != -1 )
		{
			i = wholestr.indexOf( instring );
			i++;
			return new PtgInt( i );
		}
		return new PtgErr( PtgErr.ERROR_VALUE );

	}

	/**
	 * FINDB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, FINDB counts each character as 1.
	 * NOTES:  search is case sensitive and doesn't allow for wildcards
	 */
	protected static Ptg calcFindB( Ptg[] operands )
	{
		if( (operands == null) || (operands.length < 2) || (operands[0] == null) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		// determine if Excel's language is set up for DBCS; if not, returns normal string
		com.extentech.formats.XLS.WorkBook bk = operands[0].getParentRec().getWorkBook();
		if( !bk.defaultLanguageIsDBCS() )
		{ // otherwise just use calcFind
			return calcFind( operands );
		}
		int startnum = 0;
		if( operands.length > 2 )
		{
			startnum = operands[2].getIntVal();
		}
		byte[] strToFind = getUnicodeBytesFromOp( operands[0] );
		byte[] str = getUnicodeBytesFromOp( operands[1] );
		int index = -1;
		if( (strToFind == null) || (strToFind.length == 0) || (str == null) || (startnum < 0) || (str.length < startnum) )
		{
			return new PtgInt( startnum );
		}
		for( int i = startnum; (i < str.length) && (index == -1); i++ )
		{
			if( strToFind[0] == str[i] )
			{
				index = i;
				for( int j = 0; (j < strToFind.length) && ((i + j) < str.length) && (index == i); j++ )
				{
					if( strToFind[j] != str[i + j] )
					{
						index = -1; // start over
						break;
					}
				}
			}
		}
		if( index == -1 )// not found
		{
			new PtgErr( PtgErr.ERROR_VALUE );
		}
		return new PtgInt( index + 1 );    // return 1-based index of found bytes
	}

	/**
	 * FIXED
	 * Formats a number as text with a fixed number of decimals
	 */
	protected static Ptg calcFixed( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		boolean nocommas = false;
		if( operands.length == 3 )
		{
			Boolean boo = (Boolean) operands[2].getValue();
			nocommas = boo;
		}
		double dub = operands[0].getDoubleVal();
		if( dub == Double.NaN )
		{
			dub = 0;
		}
		int pop = operands[1].getIntVal();
		dub = dub * Math.pow( 10, pop );
		dub = Math.round( dub );
		dub = dub / Math.pow( 10, pop );
		String res = String.valueOf( dub );
		if( pop == 0 )
		{
			if( res.indexOf( "." ) > -1 )
			{
				res = res.substring( 0, res.indexOf( "." ) );
				return new PtgStr( res );
			}
		}
		// pad w/zeros if need be.
		if( (res.indexOf( "." ) == -1) && (pop > 0) )
		{
			res = res + ".0";
		}
		String mantissa = res.substring( res.indexOf( "." ) );
		while( mantissa.length() <= pop )
		{
			res += 0;
			mantissa = res.substring( res.indexOf( "." ) );
		}
		if( nocommas || (dub < 999.99) )
		{
			return new PtgStr( res );
		}

		int e = res.indexOf( "." );
		String mant = res.substring( e );
		String begin = res.substring( 0, e );
		int counter = 0;
		int s = begin.length();
		// this adds the commas;
		for( int v = 0; v < s; )
		{
			String ch = begin.substring( (s - v) - 1, s - v );
			mant = ch + mant;
			v++;
			if( (counter == 2) && (v != s) )
			{
				mant = "," + mant;
			}
			counter++;
			if( counter == 3 )
			{
				counter = 0;
			}
		}
		return new PtgStr( mant );

	}

	/**
	 * JIS function
	 * The function described in this Help topic converts half-width (single-byte)
	 * letters within a character string to full-width (double-byte) characters.
	 * The name of the function (and the characters that it converts) depends upon your language settings.
	 * For Japanese, this function changes half-width (single-byte) English letters or
	 * katakana within a character string to full-width (double-byte) characters.
	 * JIS(text)
	 * Text   is the text or a reference to a cell that contains the text you want to change. If text does not contain any half-width English letters or katakana, text is not changed.
	 * <p/>
	 * TODO: STRING ENCODING IS NOT CORRECT **************
	 */
 /*
  * encoding info:
	Shift_JIS	DBCS		16-bit Japanese encoding (Note that you must use an underscore character (_), not a hyphen (-) in the name in CFML attributes.)
	(same as MS932)
	EUC-KR		DBCS		16-bit Korean encoding
	UCS-2		DBCS		Two-byte Unicode encoding
	UTF-8		MBCS		Multibyte Unicode encoding. ASCII is 7-bit; non-ASCII characters used in European and many Middle Eastern languages are two-byte; and most Asian characters are three-byte   
*/
	protected static Ptg calcJIS( Ptg[] operands )
	{
		if( (operands == null) || (operands[0] == null) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		// determine if Excel's language is set up for DBCS; if not, returns normal string
		com.extentech.formats.XLS.WorkBook bk = operands[0].getParentRec().getWorkBook();
		if( bk.defaultLanguageIsDBCS() )
		{ // otherwise just returns normal string
			byte[] strbytes = getUnicodeBytesFromOp( operands[0] );
			if( strbytes == null )
			{
				strbytes = (operands[0].getValue()).toString().getBytes();
			}
			try
			{
				return new PtgStr( new String( strbytes, "Shift_JIS" ) );
			}
			catch( Exception e )
			{
			}
		}
		return new PtgStr( operands[0].getValue().toString() );
	}

	/**
	 * LEFT
	 * Returns the leftmost characters from a text value
	 */
	protected static Ptg calcLeft( Ptg[] operands )
	{
		int numchars = 1;
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		if( operands[0] instanceof PtgErr )
		{
			return new PtgErr( PtgErr.ERROR_NA );    // 'tis what excel does
		}
		if( operands.length == 2 )
		{
			if( operands[1] instanceof PtgErr )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			numchars = operands[1].getIntVal();
		}
		Object o = operands[0].getValue();
		if( o == null )
		{
			return new PtgStr( "" );
		}
		String str = String.valueOf( o );
		if( (str == null) || (numchars > str.length()) )
		{
			return new PtgStr( "" );    // 20081202 KSC: Don't error out if not enough chars ala Excel
		}
		String res = str.substring( 0, numchars );
		return new PtgStr( res );
	}

	/**
	 * LEFTB counts each double-byte character as 2 when you have enabled the editing of a
	 * language that supports DBCS and then set it as the default language.
	 * Otherwise, LEFTB counts each character as 1.
	 */
	protected static Ptg calcLeftB( Ptg[] operands )
	{
		com.extentech.formats.XLS.WorkBook bk = operands[0].getParentRec().getWorkBook();
		if( bk.defaultLanguageIsDBCS() )
		{// otherwise just returns normal string
			int numchars = 1;
			if( operands.length < 1 )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			if( operands.length == 2 )
			{
				if( operands[1] instanceof PtgErr )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
				try
				{
					numchars = operands[1].getIntVal();
					byte[] b = new byte[numchars];
					System.arraycopy( getUnicodeBytesFromOp( operands[0] ), 0, b, 0, numchars );
					return new PtgStr( new String( b, XLSConstants.UNICODEENCODING ) );
				}
				catch( Exception e )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
			}
		}
		return calcLeft( operands );
	}

	/**
	 * LEN
	 * Returns the number of characters in a text string
	 */
	protected static Ptg calcLen( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getError();
		}
		String s = String.valueOf( operands[0].getValue() );
		return new PtgInt( s.length() );
	}

	/**
	 * LENB counts each double-byte character as 2 when you have enabled the editing of
	 * a language that supports DBCS and then set it as the default language.
	 * Otherwise, LENB counts each character as 1.
	 */
	protected static Ptg calcLenB( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getError();
		}
		com.extentech.formats.XLS.WorkBook bk = operands[0].getParentRec().getWorkBook();
		if( bk.defaultLanguageIsDBCS() )  // otherwise just returns normal string
		{
			return new PtgInt( getUnicodeBytesFromOp( operands[0] ).length );
		}
		String s = String.valueOf( operands[0].getValue() );
		return new PtgInt( s.length() );
	}

	/**
	 * LOWER
	 * Converts text to lowercase
	 */
	protected static Ptg calcLower( Ptg[] operands )
	{
		if( operands.length > 1 )
		{
			return PtgCalculator.getError();
		}
		String s = String.valueOf( operands[0].getValue() );
		s = s.toLowerCase();
		return new PtgStr( s );
	}

	/**
	 * MID
	 * Returns a specific number of characters from a text string starting at the position you specify
	 */
	protected static Ptg calcMid( Ptg[] operands )
	{
		String s = String.valueOf( operands[0].getValue() );
		if( (s == null) || s.equals( "" ) )
		{
			return new PtgStr( "" );    //  Don't error out if "" ala Excel
		}
		if( (operands[1] instanceof PtgErr) || (operands[2] instanceof PtgErr) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		int start = operands[1].getIntVal() - 1;
		int len = operands[2].getIntVal();
		if( len < 0 )
		{
			len = start + len;
		}
		if( s.length() < start )
		{
			return new PtgStr( "" );
		}
		if( start == -1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		s = s.substring( start );
		if( len > s.length() )
		{
			return new PtgStr( s );
		}
		s = s.substring( 0, len );
		return new PtgStr( s );
	}
 /*
  * MIDB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, MIDB counts each character as 1.
  */
 /*
  * PHONETIC function
	Extracts the phonetic (furigana) characters from a text string.
	PHONETIC(reference)
	Reference   is a text string or a reference to a single cell or a range of cells that contain a furigana text string.
  */

	/**
	 * PROPER
	 * Capitalizes the first letter in each word of a text value
	 */
	protected static Ptg calcProper( Ptg[] operands )
	{
		String s = String.valueOf( operands[0].getValue() );
		s = StringTool.proper( s );
		return new PtgStr( s );
	}

	/**
	 * REPLACE
	 * Replaces characters within text
	 */
	protected static Ptg calcReplace( Ptg[] operands )
	{
		String origstr = String.valueOf( operands[0].getValue() );
		int start = operands[1].getIntVal();
		int repamount = operands[2].getIntVal();
		String repstr = String.valueOf( operands[3].getValue() );
		String begin = origstr.substring( 0, (start - 1) );
		String end = origstr.substring( (start + repamount) - 1 );
		String returnstr = begin + repstr + end;
		return new PtgStr( returnstr );

	}
 /*
  * REPLACEB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, REPLACEB counts each character as 1.
  */

	/**
	 * REPT
	 * Repeats text a given number of times
	 */
	protected static Ptg calcRept( Ptg[] operands )
	{
		String origstr = String.valueOf( operands[0].getValue() );
		int numtimes = operands[1].getIntVal();
		String retstr = "";
		for( int i = 0; i < numtimes; i++ )
		{
			retstr += origstr;
		}
		return new PtgStr( retstr );
	}

	/**
	 * RIGHT
	 * Returns the rightmost characters from a text value
	 */
	protected static Ptg calcRight( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		String origstr = String.valueOf( operands[0].getValue() );
		if( origstr.equals( "" ) )
		{
			return new PtgStr( "" );
		}
		int numchars = operands[1].getIntVal();
		if( numchars > origstr.length() )
		{
			numchars = origstr.length();
		}
		if( numchars < 0 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		String res = origstr.substring( origstr.length() - numchars );
		return new PtgStr( res );
	}
 /*
  * RIGHTB counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS and then set it as the default language. Otherwise, RIGHTB counts each character as 1.
  */

	/**
	 * SEARCH
	 * Finds one text value within another (not case-sensitive)
	 */
	protected static Ptg calcSearch( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		int start = 0;
		if( operands.length == 3 )
		{
			start = operands[2].getIntVal() - 1;
		}
		String search = operands[0].getValue().toString().toLowerCase();
		String orig = operands[1].getValue().toString().toLowerCase();
		String tmp = orig.substring( start ).toLowerCase();
		int i = tmp.indexOf( search );
		if( i == -1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		i = orig.indexOf( search );
		i++;
		return new PtgInt( i );
	}

	/**
	 * SEARCHB counts each double-byte character as 2 when you have enabled the editing of a
	 * language that supports DBCS and then set it as the default language.
	 * Otherwise, SEARCHB counts each character as 1.
	 * <p/>
	 * TODO: THIS IS NOT COMPLETE
	 */
	protected static Ptg calcSearchB( Ptg[] operands )
	{
		if( (operands == null) || (operands.length < 2) || (operands[0] == null) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		// determine if Excel's language is set up for DBCS; if not, returns normal string
		com.extentech.formats.XLS.WorkBook bk = operands[0].getParentRec().getWorkBook();
		if( !bk.defaultLanguageIsDBCS() )
		{ // otherwise just use calcFind
			return calcSearch( operands );
		}
		int startnum = 0;
		if( operands.length > 2 )
		{
			startnum = operands[2].getIntVal();
		}
		byte[] strToFind = getUnicodeBytesFromOp( operands[0] );
		byte[] str = getUnicodeBytesFromOp( operands[1] );
		int index = -1;
		if( (strToFind == null) || (strToFind.length == 0) || (str == null) || (startnum < 0) || (str.length < startnum) )
		{
			return new PtgInt( startnum );
		}

		String search = operands[0].getValue().toString().toLowerCase();
		String orig = operands[1].getValue().toString().toLowerCase();
		String tmp = orig.substring( startnum ).toLowerCase();
		index = tmp.indexOf( search );

		if( index == -1 )// not found
		{
			new PtgErr( PtgErr.ERROR_VALUE );
		}
		else
		{
			index *= 2; // count the bytes as double
		}
		return new PtgInt( index + 1 );    // return 1-based index of found bytes
	}

	/**
	 * SUBSTITUTE
	 * Substitutes new text for old text in a text string
	 */
	protected static Ptg calcSubstitute( Ptg[] operands )
	{
		int whichreplace = 0;
		if( operands.length < 3 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		if( operands.length == 4 )
		{
			whichreplace = operands[3].getIntVal() - 1;
		}
		String origstr = operands[0].getValue().toString();
		String srchstr = operands[1].getValue().toString();
		String repstr = operands[2].getValue().toString();
		String finalstr = StringTool.replaceText( origstr, srchstr, repstr, whichreplace, true );
		return new PtgStr( finalstr );
	}

	/**
	 * T
	 * According to documentation converts its arguments to text -
	 * <p/>
	 * not really though, it just returns value if they are text
	 */
	protected static Ptg calcT( Ptg[] operands )
	{
		String res = "";
		try
		{
			res = (String) operands[0].getValue();
		}
		catch( ClassCastException e )
		{
		}
		;
		return new PtgStr( res );

	}

	/**
	 * TEXT
	 * Formats a number and converts it to text
	 * <p/>
	 * Converts a value to text in a specific number format.
	 * <p/>
	 * Syntax
	 * <p/>
	 * TEXT(value,format_text)
	 * <p/>
	 * Value   is a numeric value, a formula that evaluates to a numeric value, or a reference to a cell containing a numeric value.
	 * <p/>
	 * Format_text   is a number format in text form from in the Category box on the Number tab in the Format Cells dialog box.
	 * <p/>
	 * Remarks
	 * <p/>
	 * Format_text cannot contain an asterisk (*).
	 * <p/>
	 * Formatting a cell with an option on the Number tab (Cells command, Format menu) changes only the format, not the value. Using the TEXT function converts a value to formatted text, and the result is no longer calculated as a number.
	 * <p/>
	 * Salesperson Sales
	 * Buchanan 2800
	 * Dodsworth 40%
	 * <p/>
	 * Formula Description (Result)
	 * =A2&" sold "&TEXT(B2, "$0.00")&" worth of units." Combines contents above into a phrase (Buchanan sold $2800.00 worth of units.)
	 * =A3&" sold "&TEXT(B3,"0%")&" of the total sales." Combines contents above into a phrase (Dodsworth sold 40% of the total sales.)
	 */
	protected static Ptg calcText( Ptg[] operands )
	{
		if( operands.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		String res = "#ERR!";
		try
		{
			res = String.valueOf( operands[0].getValue() );
		}
		catch( Exception e )
		{
			res = operands[0].toString();
		}

		String fmt = operands[1].toString();
		Format fmtx = null;
		// convert a string like "0"
		// to a format pattern like: "##";
		for( int t = 0; t < FormatConstants.NUMERIC_FORMATS.length; t++ )
		{
			String fmx = FormatConstants.NUMERIC_FORMATS[t][0];
			if( fmx.equals( fmt ) )
			{
				fmt = FormatConstants.NUMERIC_FORMATS[t][2];
				fmtx = new DecimalFormat( fmt );
			}
		}
		if( fmtx == null )
		{
			for( int t = 0; t < FormatConstants.CURRENCY_FORMATS.length; t++ )
			{
				String fmx = FormatConstants.CURRENCY_FORMATS[t][0];
				if( fmx.equals( fmt ) )
				{
					fmt = FormatConstants.CURRENCY_FORMATS[t][2];
					fmtx = new DecimalFormat( fmt );
				}
			}
		}
		if( fmtx != null )
		{
			try
			{
				if( (res != null) && !res.equals( "" ) )        // 20090527 KSC: when cell=="", Excel treats as 0
				{
					return new PtgStr( fmtx.format( new Float( res ) ) );
				}
				if( res != null )
				{
					return new PtgStr( fmtx.format( 0 ) );
				}
			}
			catch( Exception e )
			{
//	            Logger.logWarn("getting formatted string value for :" + res.toString() + " failed: " + e.toString()) ;
				try
				{
					return new PtgStr( res );    // 20080211 KSC: Double.valueOf(ret.toString());
				}
				catch( NumberFormatException nbe )
				{
					; // who knew? - of course,  functions don't have to return numbers!
				}
			}
		}
		for( int x = 0; x < FormatConstants.DATE_FORMATS.length; x++ )
		{
			String fmx = FormatConstants.DATE_FORMATS[x][0];
			if( fmx.equals( fmt ) )
			{
				fmt = FormatConstants.DATE_FORMATS[x][2];

				try
				{
					Date d;
					try
					{
						d = DateConverter.getDateFromNumber( new Double( res ) );
					}
					catch( NumberFormatException e )
					{
						d = DateConverter.getDate( res );    // try to convert a string date
						if( d == null )
						// what excel does, if it's an empty date, it reverts to jan 0, 1900
						{
							d = new Date( "1/1/1990" );
						}
					}
					//SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
					try
					{
						//sdf = new SimpleDateFormat(fmt);
						WorkBookHandle.simpledateformat.applyPattern( fmt );
					}
					catch( Exception ex )
					{
						Logger.logWarn( "Simple Date Format could not parse: " + fmt + ". Returning default." ); //not a valid date format
					}
					//return new PtgStr(sdf.format(d));
					return new PtgStr( WorkBookHandle.simpledateformat.format( d ) );
				}
				catch( Exception e )
				{
					Logger.logErr( "Unable to calcText formatting correctly for a date" + e );
				}
			}
		}

		// we've been unable to format, try based on the string
		try
		{
			if( Xf.isDatePattern( fmt ) )
			{
				//fmtx = new SimpleDateFormat( fmt );
				WorkBookHandle.simpledateformat.applyPattern( fmt );
				fmtx = WorkBookHandle.simpledateformat;
			}
			else
			{
				fmtx = new DecimalFormat( fmt );
			}

			if( (res != null) && !res.equals( "" ) )
			{
				return new PtgStr( fmtx.format( new Float( res ) ) );
			}
			if( res != null )
			{
				return new PtgStr( fmtx.format( 0 ) );
			}
		}
		catch( Exception e )
		{
			//Logger.logWarn("getting formatted string value for :" + res.toString() + " failed: " + e.toString()) ;
			try
			{
				return new PtgStr( res );
			}
			catch( NumberFormatException nbe )
			{
			}
		}
		return new PtgStr( res );
	}

	/**
	 * TRIM
	 * According to documentation Trim() removes leading and trailing spaces from the cell value.
	 * <p/>
	 * Actually it removes all spaces except for single spaces
	 * between words.
	 */
	protected static Ptg calcTrim( Ptg[] operands )
	{
		Object o = operands[0].getValue();
		String res;
		if( o instanceof Double )
		{
			res = ExcelTools.getNumberAsString( (Double) o );
		}
		else
		{
			res = String.valueOf( o );
		}
		if( (res == null) || res.equals( new PtgErr( PtgErr.ERROR_NA ).toString() ) )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		// first let's remove the beginning and trailing spaces.
		if( res.length() > 0 )
		{
			while( res.substring( 0, 1 ).equals( " " ) )
			{
				res = res.substring( 1, res.length() );
			}
			while( res.substring( res.length() - 1, res.length() ).equals( " " ) )
			{
				res = res.substring( 0, res.length() - 1 );
			}
			// now we need to remove double spaces
			while( res.indexOf( "  " ) != -1 )
			{
				int i = res.indexOf( "  " );
				String prestring = res.substring( 0, i );
				String poststring = res.substring( i + 1, res.length() );
				res = prestring + poststring;
			}
		}
		return new PtgStr( res );

	}

	/**
	 * UPPER
	 * Converts text to uppercase
	 */
	protected static Ptg calcUpper( Ptg[] operands )
	{
		if( operands.length > 1 )
		{
			return PtgCalculator.getError();
		}
		String s = String.valueOf( operands[0].getValue() );
		s = s.toUpperCase();
		return new PtgStr( s );
	}

	/**
	 * VALUE
	 * Converts a text argument to a number
	 */
	protected static Ptg calcValue( Ptg[] operands )
	{
		try
		{
			String s = String.valueOf( operands[0].getValue() );
			if( s.equals( "" ) )
			{
				s = "0"; // Excel returns a zero for a blank value if VALUE is called upon it.
			}
			Double d = new Double( s );
			return new PtgNumber( d );
		}
		catch( NumberFormatException e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}

	/**
	 * helper method for all DBCS-related worksheet functions
	 *
	 * @param op
	 * @return
	 */
	private static byte[] getUnicodeBytesFromOp( Ptg op )
	{
		byte[] strbytes = null;
		if( op instanceof PtgRef )
		{
			com.extentech.formats.XLS.BiffRec rec = ((PtgRef) op).getRefCells()[0];
			if( rec instanceof com.extentech.formats.XLS.Labelsst )
			{
				strbytes = ((com.extentech.formats.XLS.Labelsst) rec).getUnsharedString().readStr();
			}
			else if( rec instanceof Formula )
			{
				strbytes = op.getValue().toString().getBytes();
			}
			else // DEBUGGING- Take out when done
			{
				Logger.logWarn( "getUnicodeBytes: Unexpected rec encountered: " + op.getClass() );
			}
		}
		else if( op instanceof PtgStr )
		{
			strbytes = new byte[((PtgStr) op).record.length - 3];
			System.arraycopy( ((PtgStr) op).record, 3, strbytes, 0, strbytes.length );
		}
		else
		{
			// DEBUGGING- Take out when done
			Logger.logWarn( "getUnicodeBytes: Unexpected operand encountered: " + op.getClass() );
		}
		return strbytes;
	}
}