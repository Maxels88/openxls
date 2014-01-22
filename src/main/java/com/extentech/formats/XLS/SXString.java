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
 * SXString 0x
 * specifies a pivot cache item with a string value.
 *
 *  cch (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the segment field. If cch is 0xFFFF, segment MUST NOT exist.
 segment (variable): An XLUnicodeStringNoCch structure that specifies a segment of the string. This exists only if the value of the cch field is different than 0xFFFF.

 */

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.io.UnsupportedEncodingException;
import java.util.Arrays;

public class SXString extends XLSRecord implements XLSConstants, PivotCacheRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 9027599480633995587L;
	private short cch;    // length of segment
	private String segment;    //specifies a cache item with a string value.

	public void init()
	{
		super.init();
		cch = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		if( cch > 0 )
		{
			byte encoding = this.getByteAt( 2 );

			byte[] tmp = this.getBytesAt( 3, (cch) * (encoding + 1) );
			try
			{
				if( encoding == 0 )
				{
					segment = new String( tmp, DEFAULTENCODING );
				}
				else
				{
					segment = new String( tmp, UNICODEENCODING );
				}
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logInfo( "SXString.init: " + e );
			}
		}
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( this.toString() );
		}
	}

	public String toString()
	{
		return "SXString: " + ((segment != null) ? segment : "null") +
				Arrays.toString( this.getRecord() );
	}

	/**
	 * create a new minimum SXString
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		SXString sxstring = new SXString();
		sxstring.setOpcode( SXSTRING );
		sxstring.setData( new byte[]{ -1, -1 } );
		sxstring.init();
		return sxstring;
	}

	/**
	 * sets the value for the string cache item referenced by this SXString
	 *
	 * @param s
	 */
	public void setCacheItem( String s )
	{
		segment = s;
		byte[] strbytes = new byte[0];
		if( segment != null )
		{
			try
			{
				strbytes = segment.getBytes( DEFAULTENCODING );
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logInfo( "SxString: " + e );
			}
		}
		cch = (short) strbytes.length;
		byte[] nm = ByteTools.shortToLEBytes( cch );
		byte[] data = new byte[cch + 3];    // account for encoding bytes + cch
		System.arraycopy( nm, 0, data, 0, 2 );
		System.arraycopy( strbytes, 0, data, 3, cch );
		this.setData( data );
// 	 (byte[]) [5, 0, 0, 115, 111, 117, 116, 104]	south		   
	}

	public String getCacheItem()
	{
		return segment;
	}

	/**
	 * return the bytes describing this record, including the header
	 *
	 * @return
	 */
	public byte[] getRecord()
	{
		byte[] b = new byte[4];
		System.arraycopy( ByteTools.shortToLEBytes( this.getOpcode() ), 0, b, 0, 2 );
		System.arraycopy( ByteTools.shortToLEBytes( (short) this.getData().length ), 0, b, 2, 2 );
		return ByteTools.append( this.getData(), b );

	}
}
