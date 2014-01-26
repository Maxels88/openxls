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

/**
 * The SXDBEx record specifies additional PivotCache properties. 
 *
 *  numDate (8 bytes): A DateAsNum structure that specifies the date and time on which the PivotCache was created or last refreshed.
 cSxFormula (4 bytes): An unsigned integer that specifies the count of SXFormula records for this cache.

 */

import org.openxls.ExtenXLS.DateConverter;
import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Arrays;
import java.util.Locale;

public class SXDBEx extends XLSRecord implements XLSConstants, PivotCacheRecord
{
	private static final Logger log = LoggerFactory.getLogger( SXDBEx.class );
	private static final long serialVersionUID = 9027599480633995587L;
	int nformulas;
	double lastdate;

	@Override
	public void init()
	{
		super.init();
			log.debug( "SXDBEx - {}", Arrays.toString( getData() ) );
		// numDate 0-8 last refresh date
		lastdate = ByteTools.eightBytetoLEDouble( getBytesAt( 0, 8 ) );
		nformulas = ByteTools.readInt( getBytesAt( 8, 4 ) );    // # SXFormula records
	}

	public String toString()
	{
		java.util.Date ld = DateConverter.getDateFromNumber( lastdate );
		java.text.DateFormat dateFormatter = java.text.DateFormat.getDateInstance( java.text.DateFormat.DEFAULT, Locale.getDefault() );
		try
		{
			return "SXDBEx: nFormulas:" + nformulas + " last Date:" + dateFormatter.format( ld ) +
					Arrays.toString( getRecord() );
		}
		catch( Exception e )
		{
		}
		return "SXDBEx: nFormulas:" + nformulas + " last Date: undefined";

	}

	/**
	 * create a new minimum SXDBEx
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		SXDBEx sxdbex = new SXDBEx();
		sxdbex.setOpcode( SXDBEX );
		byte[] data = new byte[12];
		double d = DateConverter.getXLSDateVal( new java.util.Date() );
		System.arraycopy( ByteTools.doubleToLEByteArray( d ), 0, data, 0, 8 );
		sxdbex.setData( data );
		sxdbex.init();
		return sxdbex;
	}

	public void setnFormulas( int n )
	{
		nformulas = n;
		byte[] b = ByteTools.cLongToLEBytes( n );
		System.arraycopy( b, 0, getData(), 8, 4 );
	}

	public int getnFormulas()
	{
		return nformulas;
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
		System.arraycopy( ByteTools.shortToLEBytes( getOpcode() ), 0, b, 0, 2 );
		System.arraycopy( ByteTools.shortToLEBytes( (short) getData().length ), 0, b, 2, 2 );
		return ByteTools.append( getData(), b );

	}
}
