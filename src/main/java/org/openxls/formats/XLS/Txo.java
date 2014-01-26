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

import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;

/**
 * <b>Txo: Text Object (1B6h)</b><br>
 * This record stores a text object.  This record is followed
 * by two CONTINUE records which contain the text data and the
 * formatting runs respectively.
 * <p/>
 * If there is no text, the two CONTINUE records are absent.
 * <p/>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       grbit       2       Option flags.  See table.
 * 6       rot         2       Orientation of text within the object.
 * 8       Reserved    6       Must be zero.
 * 14      cchText     2       Length of text in first CONTINUE rec.
 * 16      cbRuns      2       Length of formatting runs in second CONTINUE rec.
 * 18      Reserved    4       Must be zero.
 * <p/>
 * <p/>
 * The grbit field contains the following option flags:
 * <p/>
 * bits    mask    name                contents
 * ----
 * 0       0x01    Reserved
 * 3-1     0x0e    alcH                Horizontal text alignment
 * 1 = left
 * 2 = centered
 * 3 = right
 * 4 = justified
 * 6-4     0x70    alcV                Vertical text alignment
 * 1 = top
 * 2 = center
 * 3 = bottom
 * 4 = justified
 * 8-7     0x180   Reserved
 * 9       0x200   fLockText           1 = lock text option is on
 * 15-10   0xfc00  Reserved
 * <p/>
 * <p/>
 * The first CONTINUE record contains text -- the length is cchText of this object.
 * The first byte of the CONTINUE record's body data is 0x0 = compressed unicode.  The rest is text.
 * <p/>
 * </p></pre>
 *
 * @see LABELSST
 * @see STRING
 * @see CONTINUE
 */

public final class Txo extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Txo.class );
	private static final long serialVersionUID = -7043468034346138525L;
	Continue text;
	Continue formattingruns; //20100430 KSC: garbagetxo is really a masked mso, garbagetxo;  // garbagetxo is a third continue that appears to be cropping up in infoteria files.  We are removing from the file stream currently, but may need to integrate
	int state = 0;
	short grbit = 0;
	short cchText = 0;
	short cbRuns = 0;
	short rot = 0;
	boolean compressedUnicode = false;

	@Override
	public void init()
	{
		super.init();
		int datalen = getLength(); // should be 18

		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		rot = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		cchText = ByteTools.readShort( getByteAt( 10 ), getByteAt( 11 ) );
		cbRuns = ByteTools.readShort( getByteAt( 12 ), getByteAt( 13 ) );

		setIsValueForCell( false );
		// -- not always true... does it bother anybody though?
		isString = true;
		isContinueMerged = true;
	}

	/**
	 * returns the String value of this Text Object
	 */
	@Override
	public String getStringVal()
	{
		String s = "";
		if( text == null )
		{
			return null;
		}
		byte[] barr;
		int encoding = text.getData()[0];
		barr = new byte[text.getData().length - 1];
		System.arraycopy( text.getData(), 1, barr, 0, barr.length );
		try
		{
			if( encoding == 0 )
			{// normal case (Default encoding)
				s = new String( barr, WorkBookFactory.DEFAULTENCODING );
			}
			else // encoding=1  (are there other encoding options??
			{
				s = new String( barr, WorkBookFactory.UNICODEENCODING );
			}
		}
		catch( UnsupportedEncodingException e )
		{
			log.warn( "reading Text Object failed: " + e.toString(), e );
		}
		return s;
	}

	/**
	 * sets the String value of this Text Object
	 * <br>if present, will parse embedded formats within the string as:
	 * <br>the format of the embedded information is as follows:
	 * <br>&lt;font specifics>text segment&lt;font specifics for next segment>text segment...
	 * <br>where font specifics can be one or more of:
	 * <ul>b		if present, bold
	 * <br>i		if present, italic
	 * <br>s		if present, strikethru
	 * <br>u		if present, underlined
	 * <br>f=""		font name e.g. "Arial"
	 * <br>sz=""	font size in points e.g. "10"
	 * <br>delimited by ;'s
	 * <br>For Example:
	 * <br>&lt;f="Tahoma";b;sz="16">Note: &lt;f="Tahoma";sz="12">This is an important point
	 *
	 * @throws IllegalArgumentException if String is incorrect format
	 */
	// TODO: if length of string is > 8218, must have another continues ******************
	@Override
	public void setStringVal( String v ) throws IllegalArgumentException
	{
		if( (v != null) && (v.indexOf( '<' ) >= 0) )
		{
			v = parseFormatting( v );    // extracts text string from formats and sets formatting runs
		}
		else // no formatting present:
		{
			setFormattingRuns( null );    // reset formatting runs
		}
		// get the length of the first CONTINUE
		byte[] a = v.getBytes();
		byte[] b = new byte[a.length + 1];
		System.arraycopy( a, 0, b, 1, a.length );
			log.debug( "Txo CHANGING: " + getStringVal() );
		// TODO: checked for Compressed UNICODE in text CONTINUE
		b[0] = 0x0;
		if( text != null )
		{
			text.setData( b );
		}
		else    // create new text-type continues - though should be already set (see getPrototype)
		{
			text = Continue.getTextContinues( v );
		}
		b = ByteTools.shortToLEBytes( (short) a.length );
		getData()[10] = b[0];
		getData()[11] = b[1];
		cchText = ByteTools.readShort( getByteAt( 10 ), getByteAt( 11 ) );    // reset text length
			log.debug( " TO: " + getStringVal() );
	}

	/**
	 * String s contains formatting information; extracts text string from
	 * formats and sets formatting runs
	 * <br>the format of the embedded information is as follows:
	 * <br>< font specifics>text segment< font specifics for next segment>text segment...
	 * <br>where font specifics can be one or more of:
	 * <ul>b		if present, bold
	 * <br>i		if present, italic
	 * <br>s		if present, strikethru
	 * <br>u		if present, underlined
	 * <br>f=""		font name e.g. "Arial"
	 * <br>sz=""	font size in points e.g. "10"
	 * <br>delimited by ;'s
	 * <br>For Example:
	 * <br>< f="Tahoma";b;sz="16">Note: < f="Tahoma";sz="12">This is an impotant point
	 *
	 * @param s String with formatting info
	 * @return txt- string with all formatting info removed
	 */
	private String parseFormatting( String s ) throws IllegalArgumentException
	{
		// parse string for formatting run info
		try
		{
			boolean informatting = false;
			StringBuffer txt = new StringBuffer();
			short[] frs = new short[2];
			ArrayList formattingRuns = new ArrayList();
			boolean b;
			boolean it;
			boolean st;
			boolean u;
			b = it = st = u = false;
			String font = "Arial";    // default
			int sz = 10;                // default
			for( int i = 0; i < s.length(); i++ )
			{
				char c = s.charAt( i );
				if( !informatting )
				{
					if( c != '<' )
					{
						txt.append( c );
					}
					else
					{
						informatting = true;
						// initialize formatting run
						frs = new short[2];
						frs[0] = (short) txt.length();
						frs[1] = 0;
						b = it = st = u = false;
						sz = 10;            // default
						font = "Arial";    // default
					}
				}
				else
				{
					String[] z = s.substring( i ).split( "[;>]" );
					if( (z == null) || (z.length == 0) || ((z.length == 1) && (!(z[0].endsWith( ">" ) || z[0].endsWith( ";" )))) )
					{
						txt.append( '<' );    // not a real embedded format
						i--;
						informatting = false;
						continue;
					}
					String section = z[0];
					// gather up font info
					if( section.equals( "b" ) )
					{
						b = true;
					}
					else if( section.equals( "i" ) )
					{
						it = true;
					}
					else if( section.equals( "u" ) )
					{
						u = true;
					}
					else if( section.equals( "s" ) )
					{
						st = true;
					}
					else if( section.startsWith( "f=" ) )
					{
						// font name
						font = section.substring( 3 );
						font = font.substring( 0, font.indexOf( '"' ) );
					}
					else if( section.startsWith( "sz=" ) )
					{
						String ssz = section.substring( 4 );
						ssz = ssz.substring( 0, ssz.indexOf( '"' ) );
						sz = Integer.valueOf( ssz );
					}
					i += section.length();
					if( (i < s.length()) && (s.charAt( i ) == '>') )
					{    // if got end of formatting section
						// store formatting run
						informatting = false;
						Font f = new Font( font, 400, sz * 20 );// sz must be in points
						if( b )
						{
							f.setBold( b );
						}
						if( it )
						{
							f.setItalic( it );
						}
						if( u )
						{
							f.setUnderlined( u );
						}
						if( st )
						{
							f.setStricken( st );
						}
						int fIndex = getWorkBook().getFontIdx( f );  // index for specific font formatting
						if( fIndex == -1 )  // must insert new font
						{
							fIndex = getWorkBook().insertFont( f ) + 1;
						}
						frs[1] = (short) fIndex;
						formattingRuns.add( frs );
					}
				}
			}
			if( formattingRuns.size() > 0 )
			{
				formattingRuns.add( new short[]{
						(short) txt.toString().length(), 15
				} ); // 20100430 KSC: add "extra" formatting run -- necessary for Excel 2003
				setFormattingRuns( formattingRuns );
			}
			return txt.toString();
		}
		catch( Exception e )
		{
			throw new IllegalArgumentException( "Unable to parse String Pattern: " + s );
		}
	}

	/**
	 * sets the text for this object, including formatting information
	 *
	 * @param txt
	 */
	public void setStringVal( Unicodestring txt )
	{
		setStringVal( txt.getStringVal() );
		setFormattingRuns( txt.getFormattingRuns() );
	}

	/**
	 * get the formatting runs - fonts per character index, basically, for this text object,
	 * as an arraylist of short[] {char index, font index}
	 * <p/>
	 * NOTE: formatting runs in actuality differ from doc:
	 * apparently each formatting run in the Continue record occupies 8 bytes (not 4)
	 * plus there is an additional entry appended to the end, that indicates the last char index of the
	 * string (this last entry is added by Excel 2003 in the Continues Record and is not present in Unicode Strings
	 */
	public java.util.ArrayList getFormattingRuns()
	{
		java.util.ArrayList formattingRuns = new java.util.ArrayList();
		int frcontinues = getSheet().getSheetRecs().indexOf( this ) + 2;
		Continue fr = (Continue) getSheet().getSheetRecs().get( frcontinues );
		byte[] frdata = fr.getData();
		int nFormattingRuns = (frdata.length / 8);
		if( nFormattingRuns <= 1 )
		{
			return null;    // only have the "NO FONT" entry
		}
		for( int i = 0; i < (nFormattingRuns * 8); )
		{
			short idx;
			short font;
			idx = ByteTools.readShort( frdata[i++], frdata[i++] );
			font = ByteTools.readShort( frdata[i++], frdata[i++] );
			formattingRuns.add( new short[]{ idx, font } );
			i += 4;    // skip the 4 "reserved" bytes
		}
		return formattingRuns;
	}

	/**
	 * set the formatting runs - fonts per character index, basically, for this text object
	 * as an arraylist of short[] {char index, font index}
	 *
	 * @param formattingRuns // NOTES: apparently must have a minimum of 24 bytes in the Continue formatting run to work
	 *                       // 		  even though each formatting run=4 bytes, must have 4 bytes of padding between runs...?
	 *                       <p/>
	 *                       // NOTES: apparently formatting runs are different than documentation indicates:
	 *                       // each formatting run is 8 bytes:  charIndex (2), fontIndex (2), 4 bytes (ignored)
	 *                       // at end of formatting runs, must have an extra 8 bytes:  charIndex==length/fontIndex, 4 bytes
	 *                       // for the purposes of this method, the extra entry must be already present
	 *                       <p/>
	 *                       *
	 */
	public void setFormattingRuns( ArrayList formattingRuns )
	{
		byte[] frs = new byte[4];    // minimum "null formatting run" is 4? apparently must always have a "null" formatting run
		if( formattingRuns != null )
		{
			// formatting runs (charindex, fontindex)*n after string data
			frs = new byte[formattingRuns.size() * 8];
			for( int i = 0; i < formattingRuns.size(); i++ )
			{
				short[] o = (short[]) formattingRuns.get( i );
				byte[] charIndex = ByteTools.shortToLEBytes( o[0] );
				byte[] fontIndex = ByteTools.shortToLEBytes( o[1] );
				System.arraycopy( charIndex, 0, frs, (i * 8), 2 );
				System.arraycopy( fontIndex, 0, frs, (i * 8) + 2, 2 );
			}
		}
		int frcontinues = getSheet().getSheetRecs().indexOf( this ) + 2;
		Continue fr = (Continue) getSheet().getSheetRecs().get( frcontinues );
		fr.setData( frs );
		cbRuns = (short) frs.length;
		byte[] b = ByteTools.shortToLEBytes( cbRuns );
		getData()[12] = b[0];
		getData()[13] = b[1];

	}

	/**
	 * generates a skeleton Txo with 0 for text length
	 * and the minimum length for formatting runs (=2)
	 */
	public static XLSRecord getPrototype()
	{
		Txo t = new Txo();
		t.setOpcode( TXO );
		t.setData( t.PROTOTYPE_BYTES );
		t.init();
		t.text = Continue.getTextContinues( "" );
		return t;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			18, 2,	/* grbit= 530 */
			0, 0, 0, 0, 0, 0, 0, 0, 0, 0, /* cch */
			4, 0, /* formatting run length */
			0, 0, 0, 0  /* reserved must be 0 */
	};

	public String toString()
	{
		return getStringVal();
	}

}