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

/**
 * The SXDBB record specifies the values of all the cache fields that have a
 * fAllAtoms field of the SXFDB record equal to 1 and that correspond to source
 * data entities, as specified by cache fields, for a single cache record.
 * <p/>
 * blob (var) blob (variable): An array of 1-byte and 2-byte unsigned integers
 * that specifies indexes to cache items of cache fields that correspond to
 * source data entities, as specified by cache fields, that have an fAllAtoms
 * field of the SXFDB record equal to 1. The order of the indexes specified in
 * the array corresponds to the order of the cache fields as they appear in the
 * PivotCache. Each unsigned integer specifies a zero-based index of a record in
 * the sequence of records that conforms to the SRCSXOPER rule of the associated
 * cache field. The referenced record from the SRCSXOPER rule specifies a cache
 * item that specifies a value for the associated cache field. If the
 * fShortIitms field of an SXFDB record of the cache field equals 1, the index
 * value for this cache field is stored in this field in two bytes; otherwise,
 * the index value is stored in this field in a single byte.
 */

public class SxDBB extends XLSRecord implements XLSConstants, PivotCacheRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 9027599480633995587L;
	short[] cacheitems;

	// TODO: handle > 255 cache items (see SxFDB fShortItms)
	@Override
	public void init()
	{
		super.init();
		byte[] data = this.getData();
		cacheitems = new short[data.length];
		for( int i = 0; i < data.length; i++ )
		{
			cacheitems[i] = (short) data[i];        // TODO: may also be two bytes
		}
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "SXDBB -" + Arrays.toString( cacheitems ) );
		}
	}

	public String toString()
	{
		return "SXDBB: " + Arrays.toString( cacheitems ) +
				Arrays.toString( this.getRecord() );
	}

	/**
	 * create a new minimum SXDBB
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		SxDBB sxdbb = new SxDBB();
		sxdbb.setOpcode( SXDBB );
		sxdbb.setData( new byte[]{ 0 } );    // minimum (??)
		sxdbb.init();
		return sxdbb;
	}

	/**
	 * sets the cache item indexes for this cache record (or row)
	 *
	 * @param cacheitems
	 */
	public void setCacheItemIndexes( byte[] cacheitems )
	{
	   /*If the fShortIitms field of an SXFDB record of the cache field equals 1, the index
	   * value for this cache field is stored in this field in two bytes; otherwise,
	   * the index value is stored in this field in a single byte.*/
		// fShortItems means that  > 255 cache items -- assume
		this.setData( cacheitems );
		this.init();
	}

	/**
	 * retrieve the cache item indexes for this cache record (or row)
	 *
	 * @return
	 */
	public short[] getCacheItemIndexes()
	{
		return cacheitems;
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
