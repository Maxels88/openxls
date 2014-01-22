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
 * <b>Headerrec: Print Header on Each Page (14h)</b><br>
 * <p/>
 * Headerrec describes the header printed on every page
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       cch             1       Length of the Header String
 * 5       rgch            var     Header String
 * <p/>
 * </p></pre>
 */
public final class Headerrec extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -8302043395108298631L;
	int cch = -1;
	String rgch = "";
	boolean DEBUG = false;

	public void setSheet( Sheet bs )
	{
		super.setSheet( bs );
		bs.setHeader( this );
	}

	/**
	 * set the Header text
	 * The key here is that we need to construct an excel formatted unicode string,
	 * <p/>
	 * Yes, this will probably be an issue in Japan some day....
	 */
	public void setHeaderText( String t )
	{
		try
		{
			if( ByteTools.isUnicode( t ) )
			{
				byte[] bts = t.getBytes( UNICODEENCODING );
				cch = bts.length / 2;
				byte[] newbytes = new byte[cch + 3];
				byte[] cchx = com.extentech.toolkit.ByteTools.shortToLEBytes( (short) cch );
				newbytes[0] = cchx[0];
				newbytes[1] = cchx[1];
				newbytes[2] = 0x1;
				System.arraycopy( bts, 0, newbytes, 3, bts.length );
				this.setData( newbytes );
			}
			else
			{
				byte[] bts = t.getBytes( DEFAULTENCODING );
				cch = bts.length;
				byte[] newbytes = new byte[cch + 3];
				byte[] cchx = com.extentech.toolkit.ByteTools.shortToLEBytes( (short) cch );
				newbytes[0] = cchx[0];
				newbytes[1] = cchx[1];
				newbytes[2] = 0x0;
				System.arraycopy( bts, 0, newbytes, 3, bts.length );
				this.setData( newbytes );
			}
		}
		catch( Exception e )
		{
			Logger.logInfo( "setting Footer text failed: " + e );
		}
		this.rgch = t;
	}

	/**
	 * get the Header text
	 */
	public String getHeaderText()
	{
		return rgch;
	}

	public void init()
	{
		super.init();
		byte[] b = this.getData();
		if( this.getLength() > 4 )
		{
			int cch = (int) this.getByteAt( 0 );
			byte[] namebytes = this.getBytesAt( 0, this.getLength() - 4 );
			Unicodestring fstr = new Unicodestring();
			fstr.init( namebytes, false );
			rgch = fstr.toString();
			if( DEBUGLEVEL > DEBUG_LOW )
			{
				Logger.logInfo( "Header text: " + this.getHeaderText() );
			}
		}
	}

}