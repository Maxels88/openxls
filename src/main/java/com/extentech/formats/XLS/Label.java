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

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

/**
 * <b>Label: BiffRec Value, String Constant (204h)</b><br>
 * The Label record describes a cell that contains a string.
 * The String length must be in the range 000h-00ffh (0-255).
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       col         2       Column Number of the RK record
 * 8       ixfe        2       Index to XF cell format record
 * 10      cch         2       Length of the string
 * 12      rgch        var     The String
 * </p></pre>
 *
 * @see LABELSST
 * @see STRING
 * @see RSTRING
 */

public final class Label extends XLSCellRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2921430854162954640L;
	int cch;
	String val;

	public void init()
	{
		super.init();
		short s, s1;
		// get the row, col and ixfe information
		s = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		rw = (int) s;
		s = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		col = s;
		s = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		ixfe = s;
		// get the length of the string
		s1 = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
		cch = s1;
		if( this.getByteAt( 8 ) > 1 )
		{ // TODO KSC: Is this the correct indicator to read bytes as unicode??
			byte[] namebytes = this.getBytesAt( 8, this.getLength() - 8 );
			val = new String( namebytes );
		}
		else
		{
			// 20060809 KSC: read correct bytes to interpret as unicode
			try
			{
				Unicodestring thistr = null;
				byte[] tmpBytes = this.getBytesAt( 6, cch * 2 + 4 );  // i.e. (cch * 2) - 2
				thistr = new Unicodestring();
				thistr.init( tmpBytes, false );
				val = thistr.toString();
			}
			catch( Exception e )
			{
				Logger.logWarn( "ERROR Label.init: decoding string failed: " + e );
			}
		}
		this.setIsValueForCell( true );
		this.isString = true;
	}

	public void setStringVal( String v )
	{
		val = v;
		int newstrlen = v.length();
		byte[] newbytes = new byte[newstrlen + 8];
		System.arraycopy( getData(), 0, newbytes, 0, 6 );
		// byte[] newlenbytes = bto
		byte[] blen = ByteTools.cLongToLEBytes( newstrlen );
		System.arraycopy( blen, 0, newbytes, 6, 2 );
		byte[] strbytes = v.getBytes();
		System.arraycopy( strbytes, 0, newbytes, 8, newstrlen );
		this.setData( newbytes );
		this.init();
	}

	void setStringVal( String v, boolean b )
	{
		this.setStringVal( v );
	}

	public String getStringVal()
	{
		return val;
	}
}