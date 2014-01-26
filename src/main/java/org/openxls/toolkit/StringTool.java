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
package org.openxls.toolkit;

import org.openxls.formats.XLS.OOXMLAdapter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.PrintWriter;
import java.io.Reader;
import java.io.Serializable;
import java.io.StringWriter;
import java.nio.CharBuffer;
import java.util.StringTokenizer;

/**
 * A collection of useful methods for manipulating Strings.
 */
public class StringTool implements Serializable
{
	private static final Logger log = LoggerFactory.getLogger( StringTool.class );
	// static final long serialVersionUID = -5757918511951798619l;
	static final long serialVersionUID = -2761264230959133529l;

	/**
	 * replace illegal XML characters with their html counterparts
	 * <p/>
	 * ie: "&" is converted to "&amp;"
	 * <p/>
	 * check out <a href='http://www.w3.org/TR/REC-xml/'>the w3 list of XML
	 * characters</a>
	 *
	 * @param rep
	 * @return
	 */
	public static String convertXMLChars( String rep )
	{
		return OOXMLAdapter.stripNonAscii( rep ).toString();    // 20110815 KSC: this method is more complete
	}

	// test stuff
	public static void main( String[] args )
	{

		String majorHTML = "<html><body>Testing <b>yes</b><ul><li>item1</li><li>item2</li><li>item3</li></ul><br/>newline<br/>newline2<BR/>newline3 yes no 124<>something <bR>newline4<b>bold</b></body></html>";

		log.info( StringTool.stripHTML( majorHTML ) );

	}

	/**
	 * strip out all most HTML tags
	 *
	 * @param rep
	 * @return a string stripped of all html tags
	 */
	public static String stripHTML( String rep )
	{

		// first convert newlines
		rep = rep.replaceAll( "<[B,b][R,r]?>", "\r\n" );
		rep = rep.replaceAll( "<[B,b][R,r]?/>", "\r\n" );

		rep = rep.replaceAll( "<[L,l][I,i]?>", "\r\n\r\n" );

		StringBuffer ret = new StringBuffer();
		char[] cx = rep.toCharArray();
		boolean skipping = false;
		for( int t = 0; t < cx.length; t++ )
		{
			char tt = cx[t];
			// begin match
			if( tt == '<' )
			{
				skipping = true;
				t++;
			}
			else if( tt == '>' )
			{
				skipping = false;
			}
			if( !skipping && (tt != '>') )
			{
				ret.append( cx[t] );
			}

		}
		return ret.toString();
	}

	/**
	 * replace endoded text with normal text ie: "&amp;" is converted to "&"
	 * <p/>
	 * check out <a href='http://www.w3.org/TR/REC-xml/'>the w3 list of XML
	 * characters</a>
	 *
	 * @param rep
	 * @return
	 */
	public static String convertHTML( String rep )
	{
		// if(true)return rep;
		rep.replaceAll( "&amp;", "&" );
		rep.replaceAll( "&apos;", "'" );
		rep.replaceAll( "&quot;", "\"" );
		rep.replaceAll( "&lt;", "<" );
		rep.replaceAll( "&gt;", ">" );
		rep.replaceAll( "&copy;", "" );
		return rep;
	}

	/**
	 * If the string matches any part of the pattern, strip the pattern from the
	 * string.
	 */
	public static String stripMatch( String pattern, String matchstr )
	{
		String upat = pattern.toUpperCase();
		String umat = matchstr.toUpperCase();
		String retval = "";
		if( umat.lastIndexOf( upat ) > -1 )
		{
			int pos = umat.lastIndexOf( upat ) + upat.length();
			retval = matchstr.substring( pos );
			log.info( "foundpos: " + pos );
		}
		return retval;
	}

	/**
	 * get the variable name for a "getXXXX" a field name per JavaBean java
	 * naming conventions.
	 * <p/>
	 * ie: Converts "getFirstName" to "firstName"
	 */
	public static String getVarNameFromGetMethod( String thismethod )
	{
		int getidx = thismethod.indexOf( "get" );
		if( getidx < 0 )
		{
			return "";
		}
		String retval = thismethod.substring( getidx + 3 );
		retval = retval.substring( 0, retval.length() - 2 );
		String upcase = retval.substring( 0, 1 );
		upcase = upcase.toUpperCase();
		retval = retval.substring( 1 );
		retval = upcase + retval;
		return retval;
	}

	/**
	 * converts java member naming convention to underscored DB-style naming
	 * convention
	 * <p/>
	 * ie: take upperCamelCase and turn into upper_camel_case
	 */
	public static String convertJavaStyletoDBConvention( String name )
	{
		char[] chars = name.toCharArray();
		StringBuffer buf = new StringBuffer();
		for( int i = 0; i < chars.length; i++ )
		{
			// if there is a single upper-case letter, then it's a case-word
			if( Character.isUpperCase( chars[i] ) )
			{
				if( (i > 0) && ((i + 1) < chars.length) )
				{
					if( !Character.isUpperCase( chars[i + 1] ) )
					{
						buf.append( "_" );
					}
				}
				buf.append( chars[i] );
			}
			else
			{
				buf.append( Character.toUpperCase( chars[i] ) );
			}
		}
		return buf.toString();
	}

	/**
	 * converts java member naming convention to underscored DB-style naming
	 * convention
	 * <p/>
	 * ie: take upperCamelCase and turn into upper_camel_case
	 */
	public static String convertJavaStyletoFriendlyConvention( String name )
	{
		if( name.equals( "" ) )
		{
			return "";
		}
		StringBuffer buf = new StringBuffer();
		char[] chars = name.toCharArray();
		buf.append( String.valueOf( chars[0] ).toUpperCase() );
		for( int i = 1; i < chars.length; i++ )
		{
			if( chars[i] == '_' )
			{
				chars[i++] = Character.toUpperCase( chars[i + 1] );
				buf.append( " " );
				buf.append( chars[i] );
			}
			else
			{
				buf.append( String.valueOf( chars[i] ).toLowerCase() );
			}
		}
		return buf.toString();
	}

	/**
	 * convert an Array to a String representation of its objects
	 *
	 * @param name
	 * @return
	 */
	public static String arrayToString( Object[] objs )
	{
		StringBuffer ret = new StringBuffer( "[" );
		for( Object obj : objs )
		{
			ret.append( obj.toString() );
			ret.append( ", " );
		}
		ret.setLength( ret.length() - 1 );
		ret.append( "]" );
		return ret.toString();
	}

	/**
	 * Returns the given throwable's stack trace as a string.
	 *
	 * @param target the Throwable whose stack trace should be returned
	 * @return the stack trace of the given Throwable as a String
	 * @throws NullPointerException if <code>target</code> is null
	 */
	public static final String stackTraceToString( Throwable target )
	{
		StringWriter trace = new StringWriter();
		PrintWriter writer = new PrintWriter( trace );

		target.printStackTrace( writer );
		writer.flush();

		return trace.toString();
	}

	/**
	 * converts strings to "proper" capitalization
	 * <p/>
	 * ie: take "mr. fraNK sMITH" and turn into "Mr. Frank Smith"
	 */
	public static String proper( String name )
	{
		if( name.equals( "" ) )
		{
			return "";
		}
		StringBuffer buf = new StringBuffer();
		char[] chars = name.toCharArray();
		buf.append( String.valueOf( chars[0] ).toUpperCase() );
		for( int i = 1; i < chars.length; i++ )
		{
			if( chars[i] == ' ' )
			{
				buf.append( " " );
				i++;
				if( chars[i] != ' ' )
				{
					chars[i] = Character.toUpperCase( chars[i] );
					buf.append( chars[i] );
				}
			}
			else
			{
				buf.append( String.valueOf( chars[i] ).toLowerCase() );
			}
		}
		return buf.toString();
	}

	/**
	 * Each object in the array must support the .toString() method, as it will
	 * be used to render the object to its string representation.
	 */

	public static String makeDelimitedList( Object[] list, String delimiter )
	{
		StringBuffer listBuf = new StringBuffer();
		for( int i = 0; i < list.length; i++ )
		{
			if( i != 0 )
			{
				listBuf.append( delimiter );
			}
			listBuf.append( list[i].toString() );
		}
		return listBuf.toString();
	}

	/**
	 * Returns an array of strings from a single string, similar to
	 * String.split() in JavaScript
	 */
	public static String[] splitString( String value, String delimeter )
	{

		StringTokenizer stoken = new StringTokenizer( value, delimeter );
		String[] returnValue = new String[stoken.countTokens()];
		int i = 0;
		while( stoken.hasMoreTokens() )
		{
			returnValue[i] = stoken.nextToken();
			i++;
		}
		return returnValue;
	}

	/**
	 * get compressed UNICODE string from uncompressed string
	 */
	public static String getCompressedUnicode( byte[] input )
	{
		byte[] output = new byte[input.length / 2];
		int pos = 0;
		for( int i = 0; i < input.length; i++ )
		{
			output[pos++] = input[i];
			i++;
		}
		return new String( output );
	}

	/**
	 * generate a "setXXXX" string from a field name per Extentech java naming
	 * conventions.
	 * <p/>
	 * ie: Converts "firstName" to "setFirstName"
	 */
	public static String getSetMethodNameFromVar( String thismember )
	{
		String upcase = thismember.substring( 0, 1 );
		upcase = upcase.toUpperCase();
		thismember = thismember.substring( 1 );
		thismember = "set" + upcase + thismember;
		return thismember;
	}

	/**
	 * generate a "getXXXX" string from a field name per Extentech java naming
	 * conventions.
	 * <p/>
	 * ie: Converts "firstName" to "getFirstName"
	 */
	public static String getGetMethodNameFromVar( String thismember )
	{
		String upcase = thismember.substring( 0, 1 );
		upcase = upcase.toUpperCase();
		thismember = thismember.substring( 1 );
		thismember = "get" + upcase + thismember;
		return thismember;
	}

	/**
	 * converts underscored DB-style naming convention to java member naming
	 * convention
	 */
	public static String convertDBtoJavaStyleConvention( String name )
	{
		int scoreloc = name.indexOf( "_" );
		StringBuffer buf = new StringBuffer();
		if( scoreloc < 0 )
		{
			return name;
		}
		char[] chars = name.toCharArray();
		for( int i = 0; i < chars.length; i++ )
		{
			if( chars[i] == '_' )
			{
				chars[i + 1] = Character.toUpperCase( chars[i + 1] );
			}
			else
			{
				buf.append( chars[i] );
			}
		}
		return buf.toString();
	}

	/**
	 * converts underscored DB-style naming convention to java member naming
	 * convention
	 */
	public static String convertFilenameToJSPName( String name )
	{
		StringBuffer buf = new StringBuffer();
		char[] chars = name.toCharArray();
		for( int i = 0; i < chars.length; i++ )
		{
			if( chars[i] == '_' )
			{
				chars[i + 1] = Character.toUpperCase( chars[i + 1] );
			}
			else if( chars[i] == '-' )
			{
				chars[i + 1] = Character.toUpperCase( chars[i + 1] );
			}
			else if( chars[i] == ' ' )
			{
				chars[i + 1] = Character.toUpperCase( chars[i + 1] );
			}
			else
			{
				buf.append( chars[i] );
			}
		}
		return buf.toString();
	}

	/**
	 * Basically a String tokenizer
	 *
	 * @param instr
	 * @param token
	 * @return
	 */
	public static String[] getTokensUsingDelim( String instr, String token )
	{
		if( instr.indexOf( token ) < 0 )
		{
			String[] ret = new String[1];
			ret[0] = instr;
			return ret;
		}
		CompatibleVector output = new CompatibleVector();
		new StringBuffer();
		int lastpos = 0;
		int offset = 0;
		int toklen = token.length();
		int pos = instr.indexOf( token );
		// pos--;
		while( pos > -1 )
		{
			if( lastpos > 0 )// if the line starts with a token
			{
				offset = lastpos + toklen;
			}
			else
			{
				offset = 0;
			}
			String st = instr.substring( offset, pos );
			output.add( st );
			lastpos = pos;
			pos = instr.indexOf( token, lastpos + 1 );
		}
		if( lastpos < instr.length() )
		{
			String st = instr.substring( lastpos + toklen );
			output.add( st );
		}
		String[] retval = new String[output.size()];
		for( int i = 0; i < output.size(); i++ )
		{
			retval[i] = (String) output.get( i );
		}
		return retval;
	}

	// escaped slashes
	String oneslash = String.valueOf( (char) 0x005C );
	String twoslash = oneslash + oneslash;

	/**
	 *
	 */
	public static String dbencode( String holder )
	{
		return replaceText( holder, "'", "''", 0 );
	}

	/**
	 * Lose the whitespace at the end of strings...
	 *
	 * @param holder The String that you want stripped.
	 * @return Your stripped string.
	 */

	public static String allTrim( String holder )
	{
		holder = holder.trim();
		return rTrim( holder );
	}

	/**
	 * strip trailing spaces
	 */
	public static String stripTrailingSpaces( String s )
	{
		while( s.endsWith( " " ) )
		{
			s = s.substring( 0, s.length() - 1 );
		}
		return s;
	}

	/**
	 * Strips all occurences of a string from a given string.
	 *
	 * @param tostrip   The String that you want stripped.
	 * @param stripchar The char you want stripped from the String.
	 * @return Your stripped string.
	 */

	public static String strip( String tostrip, String stripstr )
	{
		StringBuffer stripped = new StringBuffer( tostrip.length() );
		while( tostrip.indexOf( stripstr ) > -1 )
		{
			stripped.append( tostrip.substring( 0, tostrip.indexOf( stripstr ) ) );
			tostrip = tostrip.substring( tostrip.indexOf( stripstr ) + stripstr.length() );
		}
		stripped.append( tostrip );
		return (stripped.toString());
	}

	/**
	 * Strips all occurences of a character from a given string.
	 *
	 * @param tostrip   The String that you want stripped.
	 * @param stripchar The char you want stripped from the String.
	 * @return Your stripped string.
	 */

	public static String strip( String tostrip, char stripchar )
	{
		StringBuffer stripped = new StringBuffer( tostrip.length() );
		int i = 0;
		char currentChar;
		while( i < tostrip.length() )
		{
			currentChar = tostrip.charAt( i );
			if( currentChar == stripchar )
			{
				i++;
			}
			else
			{
				stripped.append( currentChar );
				i++;
			}
		}
		return (stripped.toString());
	}

	/**
	 * Replaces an occurence of String B with String C within String A. This
	 * method is case sensitive.
	 * <p/>
	 * Example: String A = "I am a happy dog."; String A =
	 * stringtool.replaceText(A, "happy", "sad", 0);
	 * <p/>
	 * The result is A="I am a sad dog."
	 *
	 * @param originalText    Original text
	 * @param replaceText     Text to replace.
	 * @param replacementText Text to replace with.
	 * @param offset          offset of replacement within original string.
	 * @return Processed text.
	 */
	public static String replaceText( String originalText, String replaceText, String replacementText, int offset, boolean skipmatch )
	{
		if( !skipmatch )
		{
			return replaceText( originalText, replaceText, replacementText, offset );
		}

		StringBuffer sb = new StringBuffer();
		if( originalText.indexOf( replaceText ) < 0 )
		{
			return originalText;
		}
		int nextidx = 0;
		int lastidx = 0;
		int pos = 0;
		int textlen = replaceText.length();
		int stringlen = originalText.length();

		while( nextidx <= originalText.lastIndexOf( replaceText ) )
		{
			pos++;
			nextidx = originalText.indexOf( replaceText, lastidx );
			sb.append( originalText.substring( lastidx, nextidx ) );
			if( pos > offset )
			{
				sb.append( replacementText );
			}
			else
			{
				sb.append( replaceText );
			}
			nextidx += textlen;
			if( textlen == 0 )
			{
				break;// case of ""
			}
			lastidx = nextidx;
		}
		if( nextidx < stringlen )
		{
			sb.append( originalText.substring( nextidx ) );
		}
		return sb.toString();
	}

	/**
	 * Replaces an occurence of String B with String C within String A. This
	 * method is case sensitive.
	 * <p/>
	 * Example: String A = "I am a happy dog."; String A =
	 * stringtool.replaceText(A, "happy", "sad", 0);
	 * <p/>
	 * The result is A="I am a sad dog."
	 *
	 * @param originalText    Original text
	 * @param replaceText     Text to replace.
	 * @param replacementText Text to replace with.
	 * @param offset          offset of replacement within original string.
	 * @return Processed text.
	 */
	public static String replaceText( String originalText, String replaceText, String replacementText, int offset )
	{

		int newlen = (originalText.length() - replaceText.length()) + replacementText.length();

		if( newlen < 1 )
		{
			newlen = 0;
		}
		StringBuffer sb = new StringBuffer( newlen );
		if( originalText.indexOf( replaceText ) < 0 )
		{
			return originalText;
		}
		if( (replaceText != null) && replaceText.equals( replacementText ) )
		{
			return originalText; // avoid infinite loops
		}
		int nextidx = 0;
		int lastidx = 0;
		int textlen = replaceText.length();
		int stringlen = originalText.length();
		while( nextidx <= originalText.lastIndexOf( replaceText ) )
		{
			nextidx = originalText.indexOf( replaceText, lastidx );
			sb.append( originalText.substring( lastidx, nextidx + offset ) );
			sb.append( replacementText );
			nextidx += textlen;
			if( textlen == 0 )
			{
				break; // case of ""
			}
			lastidx = nextidx;
		}
		if( nextidx < stringlen )
		{
			sb.append( originalText.substring( nextidx ) );
		}
		return sb.toString();
	}

	/**
	 * Trims whitespace from the right side of strings.
	 *
	 * @param originalText Text to trim.
	 * @return Trimmed text.
	 */
	public static String rTrim( String originalText )
	{
		StringBuffer sb = new StringBuffer( originalText );
		sb.reverse();
		String rstr = sb.toString();
		rstr.trim();
		sb = new StringBuffer( rstr );
		sb.reverse();
		return sb.toString();
	}

	/**
	 * This method will retrieve the first instance of text between any two
	 * given patterns. This method is case sensitive.
	 * <p/>
	 * Example:
	 * <p/>
	 * String A = "I am a happy dog."; B = getTextBetweenDelims(A,"a",".");
	 * <p/>
	 * B is now equal to "m a happy dog". C declines comment.
	 *
	 * @param originalText Text to process.
	 * @param beginDelim   Delimeter for beginning of retrieved section.
	 * @param endDelim     Delimeter for end of retrieved section.
	 * @return Text between delims or "" if not found.
	 */
	public static String getTextBetweenNestedDelims( String originalText, String beginDelim, String endDelim )
	{
		StringBuffer sb = new StringBuffer( originalText.length() );
		// Check to see that both delimiters exist in the string
		if( (originalText.indexOf( beginDelim ) < 0) || (originalText.lastIndexOf( endDelim ) < 0) )
		{
			return "";
		}
		int begidx = originalText.indexOf( beginDelim ) + beginDelim.length();
		int endidx = originalText.lastIndexOf( endDelim );
		int holder = 0;
		if( begidx < endidx )
		{
			sb.append( originalText.substring( begidx, endidx ) );
		}
		else
		{
			while( (begidx > endidx) && (endidx > -1) )
			{
				holder = endidx;
				endidx = originalText.lastIndexOf( endDelim, holder + 1 );
			}
			if( (begidx < endidx) && (endidx > -1) )
			{
				sb.append( originalText.substring( begidx, endidx ) );
			}
		}
		return sb.toString();
	}

	/**
	 * This method will retrieve the first instance of text between any two
	 * given patterns. This method is case sensitive.
	 * <p/>
	 * Example:
	 * <p/>
	 * String A = "I am a happy dog."; B = getTextBetweenDelims(A,"a",".");
	 * <p/>
	 * B is now equal to "m a happy dog". C declines comment.
	 *
	 * @param originalText Text to process.
	 * @param beginDelim   Delimeter for beginning of retrieved section.
	 * @param endDelim     Delimeter for end of retrieved section.
	 * @return Text between delims or "" if not found.
	 */
	public static String getTextBetweenDelims( String originalText, String beginDelim, String endDelim )
	{
		StringBuffer sb = new StringBuffer( originalText.length() );
		// Check to see that both delimiters exist in the string
		if( (originalText.indexOf( beginDelim ) < 0) || (originalText.indexOf( endDelim ) < 0) )
		{
			return "";
		}
		int begidx = originalText.indexOf( beginDelim ) + beginDelim.length();
		int endidx = originalText.indexOf( endDelim );
		int holder = 0;
		if( begidx < endidx )
		{
			sb.append( originalText.substring( begidx, endidx ) );
		}
		else
		{
			while( (begidx > endidx) && (endidx > -1) )
			{
				holder = endidx;
				endidx = originalText.indexOf( endDelim, holder + 1 );
			}
			if( (begidx < endidx) && (endidx > -1) )
			{
				sb.append( originalText.substring( begidx, endidx ) );
			}
		}
		return sb.toString();
	}

	/**
	 * This method will replace any instance of given text within another
	 * string. This method is case sensitive.
	 * <p/>
	 * Example:
	 * <p/>
	 * String A = "I am a happy dog."; A =
	 * replaceSection(A,"happy","hippie cat", "dog");
	 * <p/>
	 * A is now equal to "I am a hippie cat.".
	 *
	 * @param originalText    Text to process.
	 * @param replaceBegin    Beggining pattern of replaced section.
	 * @param replacementText Text to replace with.
	 * @param replaceEnd      End pattern of replaced section.
	 * @return Processed text.
	 */

	public static String replaceSection( String originalText, String replaceBegin, String replacementText, String replaceEnd )
	{
		StringBuffer sb = new StringBuffer( originalText.length() );
		if( (originalText.indexOf( replaceBegin ) < 0) || (originalText.indexOf( replaceEnd ) < 0) )
		{
			return originalText;
		}
		int begidx = originalText.indexOf( replaceBegin );
		int endlen = replaceEnd.length();
		int endidx = originalText.indexOf( replaceEnd ) + endlen;
		int holder = 0;
		if( begidx < endidx )

		{
			sb.append( originalText.substring( 0, begidx ) );
			sb.append( replacementText );
			sb.append( originalText.substring( endidx ) );
		}
		else
		{
			while( (begidx > endidx) && (endidx > -1) )
			{
				holder = endidx;
				endidx = originalText.indexOf( replaceEnd, holder + 1 );
			}
			if( (begidx < endidx) && (endidx > -1) )
			{
				sb.append( originalText.substring( 0, begidx ) );
				sb.append( replacementText );
				sb.append( originalText.substring( endidx + endlen ) );
			}
		}
		return sb.toString();
	}

	public static String StripChars( String theFilter, String theString )
	{
		StringBuffer strOut = new StringBuffer( theString.length() );
		char curChar;
		for( int i = 0; i < theString.length(); i++ )
		{
			curChar = theString.charAt( i );
			if( theFilter.indexOf( curChar ) < 0 )
			{ // if it's not in the filter,
				// send it thru
				strOut.append( curChar );
			}
		}
		return strOut.toString();
	}

	public static String UseOnlyChars( String theFilter, String theString )
	{
		StringBuffer strOut = new StringBuffer( theString.length() );
		char curChar;
		for( int i = 0; i < theString.length(); i++ )
		{
			curChar = theString.charAt( i );
			if( theFilter.indexOf( curChar ) > -1 )
			{ // if it's in the filter,
				// send it thru
				strOut.append( curChar );
			}
		}
		return strOut.toString();
	}

	public static String replaceChars( String theFilter, String theString, String replacement )
	{
		StringBuffer strOut = new StringBuffer( theString.length() );
		char curChar;
		for( int i = 0; i < theString.length(); i++ )
		{
			curChar = theString.charAt( i );
			if( theFilter.indexOf( curChar ) < 0 )
			{ // if it's not in the filter,
				// send it thru
				strOut.append( curChar );
			}
			else
			{
				strOut.append( replacement );
			}
		}
		return strOut.toString();
	}

	/**
	 * replace a section of text based on pattern match throughout string.
	 */
	public static String replaceText( String theString, String theFilter, String replacement )
	{
		return replaceText( theString, theFilter, replacement, 0 );
	}

	public static boolean AllInRange( int x, int y, String theString )
	{
		char curChar;
		for( int i = 0; i < theString.length(); i++ )
		{
			curChar = theString.charAt( i );
			if( (curChar < x) || (curChar > y) )
			{
				return false;
			}
		}
		return true;
	}

	/**
	 * Replaces a specified token in a string with a value from the passed
	 * through array this is done in matching order from String to Arrray.
	 */
	public static String replaceTokenFromArray( String replace, String token, String[] vals )
	{
		StringBuffer sb = new StringBuffer();
		StringTokenizer toke = new StringTokenizer( replace, token, false );
		int i = 0;
		// add the first element
		while( toke.hasMoreTokens() )
		{
			sb.append( toke.nextToken() );
			if( i <= (vals.length - 1) )
			{
				sb.append( vals[i] );
				++i;
			}
		}
		return String.valueOf( sb );
	}

	/**
	 * returns file path fPath qualified with an ending slash
	 *
	 * @param fPath
	 * @return fPath with an ending slash
	 */
	public static String qualifyFilePath( String fPath )
	{
		StringTool.replaceChars( "\\", fPath, "/" );
		fPath = fPath.trim();
		if( !fPath.endsWith( "/" ) )
		{
			fPath += "/";
		}
		return fPath;
	}

	/**
	 * splits a filepath into directory and filename
	 *
	 * @param filePath
	 * @return
	 */
	public static String[] splitFilepath( String filePath )
	{
		String[] path = new String[2];
		filePath = StringTool.replaceText( filePath, "\\", "/" );
		int lastpath = filePath.lastIndexOf( "/" );
		if( lastpath > -1 )
		{ // strip path and directory
			path[0] = filePath.substring( 0, lastpath + 1 ); // get directory
			path[1] = filePath.substring( lastpath + 1 ); // strip directory from
			// filename
		}
		else
		{
			path[1] = filePath;
		}
		return path;
	}

	/**
	 * strips the path portion from a filepath and returns the filename
	 *
	 * @param filePath
	 * @return
	 */
	public static String stripPath( String filePath )
	{
		filePath = StringTool.replaceText( filePath, "\\", "/" );
		int lastpath = filePath.lastIndexOf( "/" );
		if( lastpath > -1 )
		{
			return filePath.substring( lastpath + 1 ); // strip directory from
		}
		// filename

		return filePath;
	}

	/**
	 * strips the path portion from a filepath and returns it
	 *
	 * @param filePath
	 * @return
	 */
	public static String getPath( String filePath )
	{
		filePath = StringTool.replaceText( filePath, "\\", "/" );
		int lastpath = filePath.lastIndexOf( "/" );
		return filePath.substring( 0, lastpath + 1 );
	}

	/**
	 * replaces the extension of a filepath with an new extension
	 *
	 * @param filepath Source FilePaht
	 * @param ext
	 * @return
	 */
	public static String replaceExtension( String filepath, String ext )
	{
		int i = filepath.lastIndexOf( "." );
		String f = filepath;
		if( i > 0 )
		{
			f = filepath.substring( 0, i ) + ext;
		}
		else
		{
			f = filepath + ext;
		}
		return f;
	}

	/**
	 * given a string, return the maximum of the width in pixels in the given
	 * the awt Font. <br>
	 * NOTE: this method does not account for line feeds contained within
	 * strings
	 *
	 * @param f awt Font
	 * @param s String to compute
	 * @return double approximate width in pixels
	 */
	public static double getApproximateStringWidth( java.awt.Font f, String s )
	{
		java.awt.FontMetrics fm = java.awt.Toolkit.getDefaultToolkit().getFontMetrics( f );
		/* width/height in pixels = (w/h field) * DPI of the display device / 72 */
		double conversion = java.awt.Toolkit.getDefaultToolkit().getScreenResolution() / 72.0;
		return fm.stringWidth( s ) * conversion;    // pixels * conversion
	}

	/**
	 * given a string, return the maximum of the width in pixels in the given
	 * the awt Font Observing Line Breaks. <br>
	 *
	 * @param f awt Font
	 * @param s String to compute
	 * @return double approximate width in pixels
	 */
	public static double getApproximateStringWidthLB( java.awt.Font f, String s )
	{
		java.awt.FontMetrics fm = java.awt.Toolkit.getDefaultToolkit().getFontMetrics( f );
		/* width/height in pixels = (w/h field) * DPI of the display device / 72 */
		double conversion = java.awt.Toolkit.getDefaultToolkit().getScreenResolution() / 72.0;
		String[] ss = s.split( "\n" );
		double len = 0;
		for( String st : ss )
		{
			len = Math.max( len, fm.stringWidth( st ) * conversion );
		}
		//		return fm.stringWidth(s) * conversion;
		return len;
	}

	/**
	 * return the approximate witdth in width in pixels of the given character
	 *
	 * @param f awt Font
	 * @param c character
	 * @return double approximate width in pixels
	 */
	public static double getApproximateCharWidth( java.awt.Font f, Character c )
	{
		java.awt.FontMetrics fm = java.awt.Toolkit.getDefaultToolkit().getFontMetrics( f );
		/* width/height in pixels = (w/h field) * DPI of the display device / 72 */
		double conversion = java.awt.Toolkit.getDefaultToolkit().getScreenResolution() / 72.0;
		return fm.charWidth( c ) * conversion;
	}

	/**
	 * return the approximate height it takes to display the given string in the
	 * given font in the given width
	 *
	 * @param f java.awt.Font
	 * @param s String
	 * @param w width in points
	 * @return
	 */
	public static double getApproximateHeight( java.awt.Font f, String s, double w )
	{
		double len = StringTool.getApproximateStringWidth( f, s );
		while( len > w )
		{
			int lastSpace = -1;
			int j = s.lastIndexOf( "\n" ) + 1;
			len = 0;
			while( (len < w) && (j < s.length()) )
			{
				len += StringTool.getApproximateCharWidth( f, s.charAt( j ) );
				if( s.charAt( j ) == ' ' )
				{
					lastSpace = j;
				}
				j++;
			}
			if( len < w )
			{
				break; // got it
			}
			if( lastSpace == -1 )
			{ // no spaces to break apart
				if( s.indexOf( ' ' ) == -1 )
				{
					break;
				}
				lastSpace = s.lastIndexOf( ' ' ); // break at
			}
			s = s.substring( 0, lastSpace ) + "\n" + s.substring( lastSpace + 1 );
		}
		int nl = s.split( "\n" ).length;
		java.awt.FontMetrics fm = java.awt.Toolkit.getDefaultToolkit().getFontMetrics( f );
		java.awt.font.LineMetrics lm = f.getLineMetrics( s, fm.getFontRenderContext() );
		// this calc appears to match Excel's ...
		float l = lm.getLeading();
		//float h= lm.getHeight();
		//System.out.println("Font: " + f.toString());		
		//System.out.println("l-i:" + fm.getLeading() + " l:" + l + " h-i:" + fm.getHeight() + " h:" + h + " a-i:" + fm.getAscent() + " a:" + lm.getAscent() + " d-i:" + fm.getDescent() + " d:" + lm.getDescent());
		float h = fm.getHeight();    // KSC: revert for now ... - l/3;	// i don't know why but this seems to match Excel's the closest
		return Math.ceil( h * (nl) );//+1));	// KSC: added + 1 for testing
	}

	/**
	 * converts an excel-style custom format to String.format custom format i.e.
	 * %flags-width-precision-conversion
	 * <p>NOTE: the pattern should be a single item in the excel-style format i.e.one of the terms in positive;negative;zero;text
	 * without semicolon.
	 * <br>NOTE: the java-specific pattern returned will have the negative formatting (sign, parenthesis ...) and so, when used,
	 * the double value must be it's absolute value i.e String.format(pattern, Math.abs(d))
	 * <br>NOTE: date formats are handled separately.  this only applies to number and currency formats
	 * <br>Excel-style:
	 * 0 (zero) 	Digit placeholder. This code pads the value with zeros to fill the format.
	 * # 	Digit placeholder. This code does not display extra zeros.
	 * ? 	Digit placeholder. This code leaves a space for insignificant zeros but does not display them.
	 * . (period) 	Decimal number.
	 * % 	Percentage. Microsoft Excel multiplies by 100 and adds the % character.
	 * , (comma) 	Thousands separator. A comma followed by a placeholder scales the number by a thousand.
	 * E+ E- e+ e- 	Scientific notation.
	 * Text Code 	Description
	 * $ - + / ( ) : space 	These characters are displayed in the number. To display any other character, enclose the character in quotation marks or precede it with a backslash.
	 * \character 	This code displays the character you specify.
	 * "text" 	This code displays text.
	 * This code repeats the next character in the format to fill the column width.		Note Only one asterisk per section of a format is allowed.
	 * _ (underscore) 	This code skips the width of the next character.
	 * This code is commonly used as "_)" (without the quotation marks)
	 * to leave space for a closing parenthesis in a positive number format
	 * when the negative number format includes parentheses.
	 * This allows the values to line up at the decimal point.
	 *
	 * @param pattern    String format pattern in Excel format
	 * @param isNegative true if the source is a negative number
	 * @return
	 * @ Text placeholder.
	 */
	public static String convertPatternFromExcelToStringFormatter( String pattern, boolean isNegative )
	{
		String curPattern = pattern;
		String jpattern = "";        // return pattern
		int w = 0;
		int precision = 0;
		String flags = "";
		char conversion = 'f';    // default
		boolean inConversion = false;
		boolean inPrecision = false;
		boolean removeSign = false;    // true if value is negative and pattern calls for parens or color change or ... i.e. don't display the negative sign
/*			
 * TODO:  \ uXXX is Locale-specific to display? works manually ...
 * TODO;  finish fractional formats:  ?/?
 */
		for( int i = 0; i < curPattern.length(); i++ )
		{
			int c = curPattern.charAt( i );
			switch( c )
			{
				case '0':
					w++;
					if( !inConversion )
					{
						jpattern += "%";
						inConversion = true;
					}
					if( inPrecision && (conversion != 'E') )
					{
						precision++;
					}
					break;
				case '?':        // don't really know what to do with this one!
					break;
				case '#':
					if( !inConversion )
					{
						jpattern += "%";
						inConversion = true;
					}
					// TODO: handle such as:  ###0.00#########   --- what's the format spec for that?????
					// if (inPrecision) precision++;
					break;
				case ',':
					flags += ",";
					break;
				case '.':
					inPrecision = true;
					break;
				case 'E':
				case 'e':
					if( !inConversion )
					{
						jpattern += "%";
						inConversion = true;
					}
					conversion = 'E';
					i++;   // format is e+, E+, e- or E-
					break;
				case '[':    //	either color code or local-specific formatting
					int j = ++i;
					int k = j;
					for(; i < curPattern.length(); i++ )
					{    // skip colors for now
						c = curPattern.charAt( i );
						if( c == '-' )    // got end of an extended char sequence - skip rest (Locale code ...)
						{
							k = i;
						}
						if( c == ']' )
						{
							break;
						}
					}
					if( inConversion )
					{
						inConversion = false;
						inPrecision = false;
						jpattern += flags + ((w > 0) ? w : "") + "." + precision + conversion;
					}
					if( k == j ) // then it was a color string
					{
						removeSign = true;
					}
					else        // it was a locale-specific string ...
					{
						jpattern += curPattern.substring( ++j, k );
					}
					break;
				case '"':    // start of delimited text
					if( inConversion )
					{
						inConversion = false;
						inPrecision = false;
						jpattern += flags + ((w > 0) ? w : "") + "." + precision + conversion;
					}
					for( i++; i < curPattern.length(); i++ )
					{
						c = curPattern.charAt( i );
						if( c == '"' )
						{
							break;
						}
						jpattern += (char) c;

					}
					break;
				// ignore
				case '@':    // text placeholder
					jpattern += "%s";
					break;
				case '*':    // repeats the next char to fill -- IGNORE!!!
					break;
				case '(':    // enclose negative #'s in parens
				case ')':
					//flags+="(";
					if( isNegative )
					{
						if( inConversion )
						{
							inConversion = false;
							inPrecision = false;
							jpattern += flags + ((w > 0) ? w : "") + "." + precision + conversion;
						}
						jpattern += (char) c;
						removeSign = true;
					}
					break;
				case '_':    // skips the width of the next char - usually _) - to leave space for a closing parenthesis in a positive number format when the negative number format includes parentheses. This allows the values to line up at the decimal point.
					if( inConversion )
					{
						inConversion = false;
						inPrecision = false;
						jpattern += flags + ((w > 0) ? w : "") + "." + precision + conversion;
					}
					i++;    // skip next char -- true in all cases???
					break;
				case '%':
					if( inConversion )
					{
						inConversion = false;
						inPrecision = false;
						jpattern += flags + ((w > 0) ? w : "") + "." + precision + conversion;
					}
					jpattern += "%%";
					break;
				case '\\':
					if( inConversion )
					{
						inConversion = false;
						inPrecision = false;
						jpattern += flags + ((w > 0) ? w : "") + "." + precision + conversion;
					}
					int z;
					if( ((i + 1) < curPattern.length()) && (curPattern.charAt( i + 1 ) == 'u') )
					{
						z = i + 6;
					}
					else
					{
						z = i + 1;
					}
					for(; (i < z) && (i < curPattern.length()); i++ )
					{
						jpattern += (char) curPattern.charAt( i );
					}
					break;
				default:    // %, $, -  space -- keep
					if( inConversion )
					{
						inConversion = false;
						inPrecision = false;
						jpattern += flags + ((w > 0) ? w : "") + "." + precision + conversion;
					}
					jpattern += (char) c;
					break;
			}
		}
		if( inConversion )
		{
			jpattern += flags + ((w > 0) ? w : "") + "." + precision + conversion;
		}
		if( isNegative && !removeSign )
		{
			jpattern = "-" + jpattern;
		}
//System.out.print("Original Pattern " + pattern + " new " + jpattern);						
//			patterns[z]= jpattern;
		pattern = jpattern;
		return pattern;
	}

	public static String convertDatePatternFromExcelToStringFormatter( String pattern )
	{
		String jpattern = "";        // return pattern
		String dString = "";            // d string -- ddd ==> EEE and dddd ==> EEEE
		String mString = "";            // m string -- either month (M, MM, MMM, MMMM) or minute
		int prev = 0;
		for( int i = 0; i < pattern.length(); i++ )
		{
			int c = pattern.charAt( i );
			if( (c != 'd') && !dString.equals( "" ) )
			{
				if( dString.length() <= 2 )
				{
					jpattern += dString;
				}
				else if( dString.length() == 3 )
				{
					jpattern += "EEE";
				}
				else if( dString.length() == 4 )
				{
					jpattern += "EEEE";
				}
				dString = "";
			}
			else if( (c != 'm') && !mString.equals( "" ) )
			{
				if( (c == ':') || (prev == 'h') )
				{    //it's time
					jpattern += mString;
					prev = c;
				}
				else
				{
					jpattern += mString.toUpperCase();
				}
				mString = "";
			}

			switch( c )
			{
				case 'y':
					jpattern += (char) c;
					break;
				case 'h':
					jpattern += 'H';    // h in java is 1-24 excel h= 0-23
					prev = 'h';
					break;
				case '\\':    // found case of erroneous use of backslash, as in: mm\-dd\-yy  ignore!
				case '[':    // no java equivalent of [h] [m] or [ss] == elapsed time
				case ']':
					break;
				case 's':
					jpattern += (char) c;
					break;
				case 'A':
					if( pattern.substring( i, i + 5 ).equals( "AM/PM" ) )
					{
						jpattern += "a";
						i += 5;
						for( int z = jpattern.length() - 2; z >= 0; z-- )
						{
							if( jpattern.charAt( z ) == 'H' )
							{
								jpattern = jpattern.substring( 0, z ) + 'h' + jpattern.substring( z + 1 );
							}
						}
					}
					break;
				case 'd':
					dString += (char) c;
					break;
				case 'm':
					mString += (char) c;
					break;
				default:
					if( (c != ':') && (c != 'm') )
					{
						prev = c;
					}
					jpattern += (char) c;
			}
		}
		if( !mString.equals( "" ) )
		{
			if( prev == 'h' )    //it's time
			{
				jpattern += mString;
			}
			else
			{
				jpattern += mString.toUpperCase();    // remaining month string
			}
		}
		else if( !dString.equals( "" ) )
		{
			if( dString.length() <= 2 )
			{
				jpattern += dString;
			}
			else if( dString.length() == 3 )
			{
				jpattern += "EEE";
			}
			else if( dString.length() == 4 )
			{
				jpattern += "EEEE";
			}
			dString = "";
		}
		return jpattern;
	}

	/**
	 * extract info, if any, from bracketed expressions within Excel custom number formats
	 *
	 * @param pattern String Excel number format
	 * @return String returned number format without the bracketed expression
	 */
	public static String convertPatternExtractBracketedExpression( String pattern )
	{
		String[] s = pattern.split( "\\[" );
		if( s.length > 1 )
		{
			pattern = "";
			for( String value : s )
			{
				int zz = value.indexOf( "]" );
				if( zz != -1 )
				{
					String term = "";
					if( value.charAt( 0 ) == '$' )
					{
						term = value.substring( 1, zz );    // skip first $
					}
					else
					{
						term = value.substring( 0, zz );
					}
					if( term.indexOf( "-" ) != -1 )  // extract character TODO: locale specifics
					{
						pattern += term.substring( 0, term.indexOf( "-" ) );
					}
					else
					{
						pattern += term;
					}
				}
				pattern += value.substring( zz + 1 );
			}
		}
		return pattern;
	}

	/**
	 * qualifies a pattern string to make valid for applying the pattern
	 *
	 * @param pattern
	 * @return
	 */
	public static String qualifyPatternString( String pattern )
	{
		pattern = StringTool.strip( pattern, "*" );
		pattern = StringTool.strip( pattern, "_(" );    // width placeholder
		pattern = StringTool.strip( pattern, "_)" );    // width placeholder
		pattern = StringTool.strip( pattern, "_" );
		pattern = pattern.replaceAll( "\"", "" );
		pattern = StringTool.strip( pattern, "?" );
		// there are more bracketed expressions to deal with
		// see http://office.microsoft.com/en-us/excel-help/creating-international-number-formats-HA001034635.aspx?redir=0
		//pattern = StringTool.strip(pattern, "[Red]");	// [Black]  [h] [hhh] [=1] [=2]
		//pattern = StringTool.strip(pattern, "Red]");
		// TODO: implement locale-specific entries:  [$-409] [$-404] ... ********************
//        pattern= pattern.replaceAll("\\[.+?\\]", "");
/*        if (s.length > 1) {
        	System.out.println(s[0]);
        	java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\[(.*?)\\]");
        	java.util.regex.Matcher m = p.matcher(pattern);

        	while(m.find()) {
        	    System.out.println(m.group(1));
        	}        	
        }*/
		String[] s = pattern.split( "\\[" );
		if( s.length > 1 )
		{
			pattern = "";
			for( String value : s )
			{
				int zz = value.indexOf( "]" );
				if( zz != -1 )
				{
					String term = value.substring( 1, zz );    // skip first $
					if( term.indexOf( "-" ) != -1 )
					{ // extract character TODO: locale specifics
						pattern += term.substring( 0, term.indexOf( "-" ) );
					}
				}
				pattern += value.substring( zz + 1 );
			}
		}
		return pattern;
	}

	/**
	 * Reads from a <code>Reader</code> into a <code>String</code>.
	 * Blocking reads will be issued to the reader and the results will be
	 * concatenated into a string, which will be returned once the reader
	 * reports end-of-input.
	 *
	 * @param reader the <code>Reader</code> from which to read
	 * @return a string containing all characters read from the input
	 */
	public static String readString( Reader reader ) throws IOException
	{
		StringBuilder builder = new StringBuilder();
		CharBuffer buffer = CharBuffer.allocate( 512 );

		while( -1 != reader.read( buffer ) )
		{
			buffer.flip();
			builder.append( buffer );
			buffer.clear();
		}

		return builder.toString();
	}
}