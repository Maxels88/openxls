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

public class SXBool extends XLSRecord implements XLSConstants, PivotCacheRecord
{
	private static final Logger log = LoggerFactory.getLogger( SXBool.class );
	private static final long serialVersionUID = 9027599480633995587L;
	boolean bool;

	@Override
	public void init()
	{
		super.init();
			log.trace( "SXBool -" );
		bool = (this.getByteAt( 0 ) == 0x1);
	}

	public String toString()
	{
		return "SXBool: " + bool;
	}

	/**
	 * create a new SXBool record
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		SXBool sxbool = new SXBool();
		sxbool.setOpcode( SXBOOL );
		sxbool.setData( new byte[2] );
		return sxbool;
	}

	public void setBool( boolean b )
	{
		bool = b;
		if( bool )
		{
			this.getData()[0] = 0x1;
		}
		else
		{
			this.getData()[0] = 0;
		}
	}

	public boolean getBool()
	{
		return bool;
	}

	/**
	 * return the bytes describing this record, including the header
	 *
	 * @return
	 */
	@Override
	public byte[] getRecord()
	{
		byte[] b = new byte[4];
		System.arraycopy( ByteTools.shortToLEBytes( this.getOpcode() ), 0, b, 0, 2 );
		System.arraycopy( ByteTools.shortToLEBytes( (short) this.getData().length ), 0, b, 2, 2 );
		return ByteTools.append( this.getData(), b );

	}
}
