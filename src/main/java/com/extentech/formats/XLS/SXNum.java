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

import java.util.Arrays;

public class SXNum extends XLSRecord implements XLSConstants, PivotCacheRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 9027599480633995587L;
	double num;

	@Override
	public void init()
	{
		super.init();
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "SXNum -" );
		}
		num = ByteTools.eightBytetoLEDouble( getBytesAt( 0, 8 ) );
	}

	public String toString()
	{
		return "SXNum: " + num +
				Arrays.toString( this.getRecord() );
	}

	/**
	 * create a new SXNum record
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		SXNum sxnum = new SXNum();
		sxnum.setOpcode( SXNUM );
		sxnum.setData( new byte[8] );
		sxnum.init();
		return sxnum;
	}

	public void setNum( double n )
	{
		num = n;
		this.setData( ByteTools.doubleToLEByteArray( num ) );
	}

	public double getNum()
	{
		return num;
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
