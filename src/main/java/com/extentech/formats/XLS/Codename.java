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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.UnsupportedEncodingException;

/**
 * '
 * the CODENAME record stores tha name for a worksheet object.  It is not necessarily the same name as you see
 * in the worksheet tab, rather it is the VB identifier name!
 */
public class Codename extends com.extentech.formats.XLS.XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Codename.class );
	private static final long serialVersionUID = -8327865068784623792L;

	private String stCodeName;

	byte cch;
	byte grbitChr;

	@Override
	public void init()
	{
		super.init();
		byte[] wtf = this.getBytesAt( 0, this.getLength() );
		cch = this.getByteAt( 0 );
		grbitChr = this.getByteAt( 2 );

		byte[] namebytes = this.getBytesAt( 3, cch );

		try
		{
			if( grbitChr == 0x1 )
			{
				stCodeName = new String( namebytes, WorkBookFactory.UNICODEENCODING );
			}
			else
			{
				stCodeName = new String( namebytes, WorkBookFactory.DEFAULTENCODING );
			}
		}
		catch( UnsupportedEncodingException e )
		{
			log.warn( "UnsupportedEncodingException in setting codename: " + e, e );
		}

	}

	public void setName( String newname )
	{
		int modnamelen = 0;
		int oldnamelen = 0;
		if( grbitChr == 0x0 )
		{
			oldnamelen = stCodeName.length();
		}
		else
		{
			oldnamelen = (stCodeName.length() * 2);
		}

		cch = (byte) newname.length();
		byte[] namebytes = newname.getBytes();
		// if (!ByteTools.isUnicode(namebytes)){
		if( !ByteTools.isUnicode( newname ) )
		{
			grbitChr = 0x0;
			modnamelen = newname.length();
		}
		else
		{
			grbitChr = 0x1;
			modnamelen = (newname.length() * 2);
		}
		byte[] newdata = new byte[(this.getData().length - oldnamelen) + modnamelen];
		try
		{
			if( grbitChr == 0x1 )
			{
				namebytes = newname.getBytes( WorkBookFactory.UNICODEENCODING );
			}
			else
			{
				namebytes = newname.getBytes( WorkBookFactory.DEFAULTENCODING );
			}
		}
		catch( UnsupportedEncodingException e )
		{
			log.warn( "UnsupportedEncodingException in setting sheet name: " + e, e );
		}
		System.arraycopy( namebytes, 0, newdata, 3, namebytes.length );
		newdata[0] = cch;
		newdata[2] = grbitChr;
		this.setData( newdata );
		this.init();
	}

}
