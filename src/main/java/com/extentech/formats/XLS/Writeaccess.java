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

import com.extentech.toolkit.Logger;

import java.io.UnsupportedEncodingException;

/**
 * <b>WRITEACCESS 0x5C: Contains name of Excel installed user.</b><br>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       stName      112     User Name as unformatted Unicodestring
 * <p/>
 * </p></pre>
 *
 * @see WorkBook
 */

public class Writeaccess extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8868603864018600260L;
	private Unicodestring strname;

	/**
	 * set the Writeaccess username
	 */
	public void setName( String str )
	{
		try
		{
			byte[] nameb = str.getBytes( DEFAULTENCODING );
			byte[] newb = new byte[112];
			int diff = 112 - nameb.length;
			if( diff < 0 )
			{
				System.arraycopy( nameb, 0, newb, 0, 112 );
			}
			else
			{
				System.arraycopy( nameb, 0, newb, 0, nameb.length );
			}
			strname.init( newb, false );
			this.setData( newb );
		}
		catch( UnsupportedEncodingException e )
		{
			Logger.logInfo( "setting name in Writeaccess record failed: " + e );
		}
	}

	public String getName()
	{
		return strname.toString();
	}

	@Override
	public void init()
	{
		super.init();
		strname = new Unicodestring();
		strname.init( getBytes(), false );
	}
}