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
 * QxiSxTag  0x802
 * <p/>
 * The QsiSXTag record specifies the name and refresh information for a query
 * table or a PivotTable view, and specifies the beginning of a collection of
 * records as defined by the Worksheet SubstreamABNF. The collection of records
 * specifies additional information for a query table or a PivotTable view.
 * <p/>
 * If fSx is 0 and stName is equal to the rgchName field of a Qsi record in this
 * worksheet substream, then this collection of records applies to the query
 * table that the Qsi record is associated with. If fSx is 1 and stName is equal
 * to the stName field of an SxView record in this worksheet substream, then
 * this collection of records applies to the PivotTable view that the SxView
 * record is associated with and its associated PivotCache. Otherwise, this
 * collection of records MUST be ignored.
 * <p/>
 * <p/>
 * frtHeaderOld (4 bytes): An FrtHeaderOld structure. The frtHeaderOld.rt field MUST be 0x0802.
 * <p/>
 * fSx (2 bytes): A Boolean (section 2.5.14) that specifies whether this record relates to a PivotTable view or a query table.
 * Value	    	Meaning
 * 0x0000		    Specifies that this record relates to a query table.
 * 0x0001		    Specifies that this record relates to a PivotTable view.
 * <p/>
 * A - fEnableRefresh (1 bit): A bit that specifies whether refresh of the PivotTable view or query table
 * is enabled. MUST be 0 if fSx is 1 and the PivotCache functionality level of the associated PivotCache is greater than or equal to 3.
 * Value    	    Value of fSx	    Meaning
 * 0				0				    Whether refresh of the query table is enabled is
 * specified by the fDisableRefresh field of the associated Qsi record.
 * 0				1				    Whether refresh of the associated PivotCache is
 * enabled is specified by the fEnableRefresh field in the SXDB record of the PivotCache.
 * 1			    0				    Specifies that refresh of the query table is enabled.
 * 1			    1				    Specifies that refresh of the associated PivotCache is enabled.
 * <p/>
 * B - fInvalid (1 bit): A bit that specifies the invalid state of the cache records of the associated PivotCache;
 * see Cache Records for more information. MUST be 1 if fSx is 1 and the PivotCache functionality level of the associated PivotCache is greater than or equal to 3.
 * <p/>
 * C - fTensorEx (1 bit): A bit that specifies whether the PivotTable view is an OLAP PivotTable view.
 * MUST be equal to 0 if fSx is 0.
 * <p/>
 * reserved1 (13 bits): MUST be zero, and MUST be ignored.
 * <p/>
 * dwQsiFuture (4 bytes): This structure specifies additional option flags for a query table or a
 * PivotTable view depending on the value of the fSx field.
 * Value of fSx Fields		Type of dwQsiFuture
 * 0x0000				    DwQsiFuture
 * 0x0001				    SXView9Save
 * <p/>
 * verSxLastUpdated (1 byte): A DataFunctionalityLevel value that specifies the data functionality level that the PivotTable view was last refreshed with.
 * MUST be 0 if this record relates to a query table.
 * <p/>
 * verSxUpdatableMin (1 byte): A DataFunctionalityLevel value that specifies the minimum version of the application that can recalculate the PivotTable view.
 * MUST be 0x00 or 0x03. MUST be 0 if this record is for a query table. MUST be 3 if the PivotCache functionality level of the associated PivotCache is 3.
 * <p/>
 * obCchName (1 byte): MUST be 0x10, and MUST be ignored.
 * <p/>
 * reserved2 (1 byte): MUST be zero, and MUST be ignored.
 * <p/>
 * stName (variable): An XLUnicodeString structure that specifies the query table or PivotTable view name.
 * <p/>
 * unused (2 bytes): Undefined and MUST be ignored.
 */
public class QsiSXTag extends XLSRecord implements XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( QsiSXTag.class );
	private static final long serialVersionUID = 2639291289806138985L;
	private short flags, ver, cchName;
	private String name;

	@Override
	public void init()
	{
		super.init();
		// QXISXTAG-[2, 8, 0, 0, 1, 0,	// required for pivot table 
		//2, 0,			// flags 
		//0, 0, 0, 0,	// dqQsiFuture 
		//3,			// verSxLastUpdated  
		//3,			// verSxUpdatableMin   
		//16,			// must be 0x10
		// 0,			// reserved 
		//11, 0,		// cchName 
		//0,			// encoding byte 
		//80, 105, 118, 111, 116, 84, 97, 98, 108, 101, 49,	// pivot table 1 
		// 0, 0]		// reserved
		flags = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
		byte verSxLastUpdated = this.getByteAt( 12 );    //
		byte verSxUpdatableMin = this.getByteAt( 13 );    // either 0 or 3
		cchName = ByteTools.readShort( this.getByteAt( 16 ), this.getByteAt( 17 ) );
		if( cchName > 0 )
		{
			byte encoding = this.getByteAt( 18 );
			byte[] tmp = this.getBytesAt( 19, (cchName) * (encoding + 1) );
			try
			{
				if( encoding == 0 )
				{
					name = new String( tmp, DEFAULTENCODING );
				}
				else
				{
					name = new String( tmp, UNICODEENCODING );
				}
			}
			catch( UnsupportedEncodingException e )
			{
				log.warn( "encoding PivotTable name in QxiSXTag: " + e, e );
			}
		}
		log.debug( "QXISXTAG-" + " flags:" + flags + " verLast:" + verSxLastUpdated + " verMin:" + verSxUpdatableMin + " name:" + name );
	}

	/**
	 * returns the name of this pivot item; if not null, is the caption
	 * for this pivot item
	 *
	 * @return
	 */
	public String getName()
	{
		return name;
	}

	/**
	 * returns the name of this pivot item; if not null, is the caption
	 * for this pivot item
	 */
	public void setName( String name )
	{
		this.name = name;
		byte[] data = new byte[18];
		System.arraycopy( this.getData(), 0, data, 0, 15 );
		if( name != null )
		{
			byte[] strbytes = null;
			try
			{
				strbytes = this.name.getBytes( DEFAULTENCODING );
			}
			catch( UnsupportedEncodingException e )
			{
				log.warn( "encoding pivot table name in SXVI: " + e, e );
			}

			//update the lengths:
			cchName = (short) strbytes.length;
			byte[] nm = ByteTools.shortToLEBytes( cchName );
			data[16] = nm[0];
			data[17] = nm[1];

			// now append variable-length string data
			byte[] newrgch = new byte[cchName + 1];    // account for encoding bytes
			System.arraycopy( strbytes, 0, newrgch, 1, cchName );

			data = ByteTools.append( newrgch, data );
			data = ByteTools.append( new byte[]{ 0, 0 }, data );
		}
		this.setData( data );
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			2, 8, 0, 0, 1, 0,	/* required for pivot table */
			1, 0,				/* flags */
			0, 0, 0, 0,			/* dqQsiFuture */
			3,					/* verSxLastUpdated */
			0,					/* verSxUpdatableMin */
			16,					/* must be 0x10 */
			0,					/* reserved */
			0, 0,				/* cchName */
	};

	public static XLSRecord getPrototype()
	{
		QsiSXTag qsi = new QsiSXTag();
		qsi.setOpcode( QSISXTAG );
		qsi.setData( qsi.PROTOTYPE_BYTES );
		qsi.init();
		return qsi;
	}

}
