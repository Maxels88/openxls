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
import java.util.Arrays;

/**
 * SXDB		0xC6
 * <p/>
 * The SXDB record specifies PivotCache properties.
 * <p/>
 * crdbdb (4 bytes): A signed integer that specifies the number of cache records for this PivotCache.
 * MUST be greater than or equal to 0. MUST be 0 for OLAP PivotCaches. MUST be ignored if fSaveData is 0.
 * <p/>
 * idstm (2 bytes): An unsigned integer that specifies the stream that contains the data for this PivotCache.
 * MUST be equal to the value of the idstm field of the SXStreamID record that specifies the PivotCache stream that contains this record.
 * <p/>
 * A - fSaveData (1 bit): A bit that specifies whether cache records exist. MUST be 0 for OLAP PivotCaches.
 * <p/>
 * B - fInvalid (1 bit): A bit that specifies whether the cache records are in the not-valid state.
 * MUST be equal to 1 if the PivotCache functionality level is greater than or equal to 3.
 * MUST be equal to 1 for OLAP PivotCaches. See cache records for more information.
 * <p/>
 * C - fRefreshOnLoad (1 bit): A bit that specifies whether the PivotCache is refreshed on load.
 * <p/>
 * D - fOptimizeCache (1 bit): A bit that specifies whether optimization is applied to the PivotCache to reduce memory usage.
 * MUST be 0 and MUST be ignored for a non-ODBC PivotCache.
 * <p/>
 * E - fBackgroundQuery (1 bit): A bit that specifies whether the query used to refresh the PivotCache is executed asynchronously.
 * MUST be ignored if vsType not equals 0x0002.
 * <p/>
 * F - fEnableRefresh (1 bit): A bit that specifies whether refresh of the PivotCache is enabled.
 * MUST be equal to 0 if the PivotCache functionality level is greater than or equal to 3.
 * MUST be equal to 0 for OLAP PivotCaches.
 * <p/>
 * unused1 (10 bits): Undefined and MUST be ignored.
 * <p/>
 * unused2 (2 bytes): Undefined and MUST be ignored.
 * <p/>
 * cfdbdb (2 bytes): A signed integer that specifies the number of cache fields that corresponds to the source data.
 * MUST be greater than or equal to 0.
 * <p/>
 * cfdbTot (2 bytes): A signed integer that specifies the number of cache fields in the PivotCache.
 * MUST be greater than or equal to 0.
 * <p/>
 * crdbUsed (2 bytes): An unsigned integer that specifies the number of records used to calculate the PivotTable report.
 * Records excluded by PivotTable view filtering are not included in this value. MUST be 0 for OLAP PivotCaches.
 * <p/>
 * vsType (2 bytes): An unsigned integer that specifies the type of source data.
 * MUST be equal to the value of the sxvs field of the SXVS record that follows the SXStreamID record that
 * specifies the PivotCache stream that contains this record.
 * <p/>
 * cchWho (2 bytes): An unsigned integer that specifies the number of characters in rgb.
 * MUST be equal to 0xFFFF, or MUST be greater than or equal to 1 and less than or equal to 0x00FF.
 * <p/>
 * rgb (variable): An optional XLUnicodeStringNoCch structure that specifies the name of the user who last refreshed the PivotCache.
 * MUST exist if and only if the value of cchWho is not equal to 0xFFFF.  If this field exists, the length MUST equal cchWho.
 * The length of this value MUST be less than 256 characters. The name is an application-specific setting that is not necessarily
 * related to the User Names StreamABNF.
 */

public class SxDB extends XLSRecord implements PivotCacheRecord
{
	private static final Logger log = LoggerFactory.getLogger( SxDB.class );
	private static final long serialVersionUID = 9027599480633995587L;
	private int crdbdb;
	private short idstm, grbit, cfdbdb, cfdbTot, crdbUsed, vsType, cchWho;
	private boolean fInvalid, fRefreshOnLoad, fEnableRefresh;    // significant bit fields
	private String rgb;

	// TODO: doesnt cfdbdb==cfdbTot always??
	// TODO: cfdbUsed == filtering ****
	// TODO: flags
	// TODO: cfWho
	@Override
	public void init()
	{
		super.init();
			log.trace( "SXDB -{}",Arrays.toString( this.getData() ) );
		crdbdb = ByteTools.readInt( this.getBytesAt( 0, 4 ) );                    // # cache records
		idstm = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );    // streamid
		grbit = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );    //
		// 8,9 = unused
		cfdbdb = ByteTools.readShort( this.getByteAt( 10 ), this.getByteAt( 11 ) );    // # cache fields
		cfdbTot = ByteTools.readShort( this.getByteAt( 12 ), this.getByteAt( 13 ) );
		crdbUsed = ByteTools.readShort( this.getByteAt( 14 ), this.getByteAt( 15 ) );    // # used - filtering
// KSC: TESTING		
//if (cfdbdb!=cfdbTot)
		//Logger.logWarn("SXDB: all cache items are not being used");
		vsType = ByteTools.readShort( this.getByteAt( 16 ), this.getByteAt( 17 ) );
		cchWho = ByteTools.readShort( this.getByteAt( 18 ), this.getByteAt( 19 ) );
		if( cchWho > 0 )
		{
			byte encoding = this.getByteAt( 20 );

			byte[] tmp = this.getBytesAt( 21, (cchWho) * (encoding + 1) );
			try
			{
				if( encoding == 0 )
				{
					rgb = new String( tmp, DEFAULTENCODING );
				}
				else
				{
					rgb = new String( tmp, UNICODEENCODING );
				}
			}
			catch( UnsupportedEncodingException e )
			{
				log.error( "SxDB.init: " + e );
			}
		}
	}

	public String toString()
	{
		return "SXDB: nCacheRecords/rows:" + crdbdb +
				" nCacheFields:" + cfdbdb +
				" cfdbTot:" + cfdbTot +
				" sid:" + idstm +
				" vsType: " + vsType +
				Arrays.toString( this.getRecord() );
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			1, 0, 0, 0,	/* n cache records - minimum=1 */
			0, 0,		/* stream id */
			33, 0,		/* flags */
			-1, 31,		/* unused */
			1, 0, 		/* cfdbdb */
			1, 0, 		/* cfdbTot */
			0, 0,		/* crdbUsed */
			1, 0,		/* vsType */
			-1, -1
	};    // cch

	/**
	 * create a new minimum SXDB
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		SxDB sxdb = new SxDB();
		sxdb.setOpcode( SXDB );
		sxdb.setData( sxdb.PROTOTYPE_BYTES );
		sxdb.init();
		return sxdb;
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

	/**
	 * sets the number of cache records (= number of non-header rows)
	 *
	 * @param n
	 */
	public void setNCacheRecords( int n )
	{
		crdbdb = (short) n;
		System.arraycopy( ByteTools.cLongToLEBytes( (crdbdb) ), 0, this.getData(), 0, 4 );
//		crdbUsed= (short)n;		// TODO: filtering affects this! 
//		System.arraycopy( ByteTools.shortToLEBytes(crdbUsed), 0, this.getData(), 14, 2);
	}

	/**
	 * returns the number of cache records (= number of non-header rows)
	 *
	 * @return
	 */
	public int getNCacheRecords()
	{
		return crdbdb;
	}

	/**
	 * returns the streamId -- index linked back to SxStreamID
	 *
	 * @return
	 */
	public short getStreamID()
	{
		return idstm;
	}

	/**
	 * sets the streamId -- index linked back to SxStreamID
	 *
	 * @param sid
	 */
	public void setStreamID( int sid )
	{
		idstm = (short) sid;
		byte[] b = ByteTools.shortToLEBytes( idstm );
		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
	}

	/**
	 * sets the number of cache fields (==columns) for the pivot cache
	 *
	 * @param n
	 */
	// TODO: doesn't cfdbTot==cfdbdb always????????????
	public void setNCacheFields( int n )
	{
		cfdbdb = (short) n;
		byte[] b = ByteTools.shortToLEBytes( cfdbdb );
		this.getData()[10] = b[0];
		this.getData()[11] = b[1];
		cfdbTot = (short) n;
		b = ByteTools.shortToLEBytes( cfdbTot );
		this.getData()[12] = b[0];
		this.getData()[13] = b[1];
	}

	/**
	 * returns the number of cache fields (==columns) for the pivot cache
	 */
	public int getNCacheFields()
	{
		return cfdbdb;
	}
}