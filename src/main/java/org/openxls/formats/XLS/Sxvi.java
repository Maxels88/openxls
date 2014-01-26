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

import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.UnsupportedEncodingException;

/**
 * SXVI B2h: This record stores view information about an Item.
 * itmType (2 bytes): A signed integer that specifies the pivot item type.
 * The value MUST be one of the following values:
 * Value		Name	    	Meaning
 * 0x0000		itmtypeData	    A data value
 * 0x0001	    itmtypeDEFAULT	Default subtotal for the pivot field
 * 0x0002	    itmtypeSUM	    Sum of values in the pivot field
 * 0x0003		itmtypeCOUNTA	Count of values in the pivot field
 * 0x0004	    itmtypeAVERAGE	Average of values in the pivot field
 * 0x0005	    itmtypeMAX	    Max of values in the pivot field
 * 0x0006	    itmtypeMIN	    Min of values in the pivot field
 * 0x0007	    itmtypePRODUCT  Product of values in the pivot field
 * 0x0008	    itmtypeCOUNT	Count of numbers in the pivot field
 * 0x0009	    itmtypeSTDEV	Statistical standard deviation (estimate) of the pivot field
 * 0x000A	    itmtypeSTDEVP	Statistical standard deviation (entire population) of the pivot field
 * 0x000B		itmtypeVAR	    Statistical variance (estimate) of the pivot field
 * 0x000C	    itmtypeVARP	    Statistical variance (entire population) of the pivot field
 * <p/>
 * A - fHidden (1 bit): A bit that specifies whether this pivot item is hidden.
 * MUST be zero if itmType is not itmtypeData. MUST be zero for OLAP PivotTable view.
 * <p/>
 * B - fHideDetail (1 bit): A bit that specifies whether the pivot item detail is collapsed.
 * MUST be zero for OLAP PivotTable view.
 * <p/>
 * C - reserved1 (1 bit): MUST be zero, and MUST be ignored.
 * <p/>
 * D - fFormula (1 bit): A bit that specifies whether this pivot item is a calculated item.
 * This field MUST be zero if any of the following apply:
 * itmType is not zero.
 * This item is in an OLAP PivotTable view.
 * The sxaxisPage field of sxaxis in the Sxvd record of the pivot field equals 1 (the associated Sxvd is the last Sxvd record before this record in the stream (1)).
 * The fCalculatedField field in the SXVDEx record of the pivot field equals 1.
 * There is not an associated SXFDB record in the associated PivotCache.
 * The fRangeGroup field of the SXFDB record, of the associated cache field of the pivot field, equals 1.
 * The fCalculatedField field of the SXFDB record, of the associated cache field of the pivot field, equals 1.
 * <p/>
 * E - fMissing (1 bit): A bit that specifies if this pivot item does not exist in the data source (1).
 * MUST be zero if itmType is not zero. MUST be zero for OLAP PivotTable view.
 * <p/>
 * reserved2 (11 bits): MUST be zero, and MUST be ignored.
 * <p/>
 * iCache (2 bytes): A signed integer that specifies a reference to a cache item.
 * MUST be a value from the following table:
 * Value			Meaning
 * -1			    No cache item is referenced.
 * 0+			    A cache item index in the cache field associated with the pivot field, as specified by Cache Items.
 * If itmType is not zero, a reference to a cache item is not specified and this value MUST be -1.
 * Otherwise, this value MUST be greater than or equal to 0.
 * <p/>
 * cchName (2 bytes): An unsigned integer that specifies the length of the stName string.
 * If the value is 0xFFFF then stName is NULL.
 * Otherwise, the value MUST be less than or equal to 254.
 * <p/>
 * stName (variable): An XLUnicodeStringNoCch structure that specifies the name of this pivot item.
 * If not NULL, this is used as the caption of the pivot item instead of the value in the
 * cache item specified by iCache. The length of this field is specified in cchName.
 * This field exists only if cchName is not 0xFFFF. If this is in a non-OLAP PivotTable
 * view and this string is not NULL, it MUST be unique within all SXVI records in associated
 * with the pivot field.
 */

public class Sxvi extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Sxvi.class );
	private static final long serialVersionUID = 6399665481118265257L;
	byte[] data = null;
	short itemtype = -1;
	short cchName = -1;
	String name = null;
	short iCache = -1;
	boolean fHidden;
	boolean fHideDetail;
	boolean fFormula;
	boolean fMissing;

	public static final short itmtypeData = 0x0000;        //A data value
	public static final short itmtypeDEFAULT = 0x0001;    //Default subtotal for the pivot field
	public static final short itmtypeSUM = 0x0002;        //Sum of values in the pivot field
	public static final short itmtypeCOUNTA = 0x0003;    //Count of values in the pivot field
	public static final short itmtypeAVERAGE = 0x0004;    // Average of values in the pivot field
	public static final short itmtypeMAX = 0x0005;        // Max of values in the pivot field
	public static final short itmtypeMIN = 0x0006;        //Min of values in the pivot field
	public static final short itmtypePRODUCT = 0x0007;    //Product of values in the pivot field
	public static final short itmtypeCOUNT = 0x0008;    //Count of numbers in the pivot field
	public static final short itmtypeSTDEV = 0x0009;    //Statistical standard deviation (estimate) of the pivot field
	public static final short itmtypeSTDEVP = 0x000A;    //Statistical standard deviation (entire population) of the pivot field
	public static final short itmtypeVAR = 0x000B;        //Statistical variance (estimate) of the pivot field
	public static final short itmtypeVARP = 0x000C;        //Statistical variance (entire population) of the pivot field

	@Override
	public void init()
	{
		super.init();
		itemtype = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		byte b = getByteAt( 2 );
		fHidden = (b & 0x1) == 0x1;
		fHideDetail = (b & 0x2) == 0x2;
		// bit 3- reserved
		fFormula = (b & 0x8) == 0x8;
		fMissing = (b & 0x10) == 0x10;
		iCache = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		cchName = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		if( cchName != -1 )
		{
			byte encoding = getByteAt( 10 );
			byte[] tmp = getBytesAt( 11, (cchName) * (encoding + 1) );
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
				log.warn( "encoding PivotTable caption name in Sxvd: " + e, e );
			}
		}
			log.debug( "SXVI - itemtype:" + itemtype + " iCache: " + iCache + " name:" + name );
	}

	public String toString()
	{
		return "SXVI - itemtype:" + itemtype + " iCache: " + iCache + " name:" + name;
	}

	/**
	 * returns the type of this pivot item
	 * <br>one of:
	 * <li>0x0000		itmtypeData	    A data value
	 * <li>0x0001	    itmtypeDEFAULT	Default subtotal for the pivot field
	 * <li>0x0002	    itmtypeSUM	    Sum of values in the pivot field
	 * <li>0x0003		itmtypeCOUNTA	Count of values in the pivot field
	 * <li>0x0004	    itmtypeAVERAGE	Average of values in the pivot field
	 * <li>0x0005	    itmtypeMAX	    Max of values in the pivot field
	 * <li>0x0006	    itmtypeMIN	    Min of values in the pivot field
	 * <li>0x0007	    itmtypePRODUCT  Product of values in the pivot field
	 * <li>0x0008	    itmtypeCOUNT	Count of numbers in the pivot field
	 * <li>0x0009	    itmtypeSTDEV	Statistical standard deviation (estimate) of the pivot field
	 * <li>0x000A	    itmtypeSTDEVP	Statistical standard deviation (entire population) of the pivot field
	 * <li>0x000B		itmtypeVAR	    Statistical variance (estimate) of the pivot field
	 * <li>0x000C	    itmtypeVARP	    Statistical variance (entire population) of the pivot field
	 *
	 * @return
	 */
	public short getItemType()
	{
		return itemtype;
	}

	/**
	 * sets the pivot item type:
	 * <br>one of:
	 * <li>0x0000		itmtypeData	    A data value
	 * <li>0x0001	    itmtypeDEFAULT	Default subtotal for the pivot field
	 * <li>0x0002	    itmtypeSUM	    Sum of values in the pivot field
	 * <li>0x0003		itmtypeCOUNTA	Count of values in the pivot field
	 * <li>0x0004	    itmtypeAVERAGE	Average of values in the pivot field
	 * <li>0x0005	    itmtypeMAX	    Max of values in the pivot field
	 * <li>0x0006	    itmtypeMIN	    Min of values in the pivot field
	 * <li>0x0007	    itmtypePRODUCT  Product of values in the pivot field
	 * <li>0x0008	    itmtypeCOUNT	Count of numbers in the pivot field
	 * <li>0x0009	    itmtypeSTDEV	Statistical standard deviation (estimate) of the pivot field
	 * <li>0x000A	    itmtypeSTDEVP	Statistical standard deviation (entire population) of the pivot field
	 * <li>0x000B		itmtypeVAR	    Statistical variance (estimate) of the pivot field
	 * <li>0x000C	    itmtypeVARP	    Statistical variance (entire population) of the pivot field
	 *
	 * @return
	 */
	public void setItemType( int type )
	{
		itemtype = (short) type;
		byte[] b = ByteTools.shortToLEBytes( itemtype );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	/**
	 * reference to a cache item :
	 * <br>-1			    No cache item is referenced.
	 * <br>0+			    A cache item index in the cache field associated with the pivot field, as specified by Cache Items.
	 *
	 * @param icache
	 */
	public void setCacheItem( int icache )
	{
		iCache = (short) icache;
		byte[] b = ByteTools.shortToLEBytes( iCache );
		getData()[4] = b[0];
		getData()[5] = b[1];
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
		byte[] data = new byte[8];
		System.arraycopy( getData(), 0, data, 0, 7 );
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
			data[6] = nm[0];
			data[7] = nm[1];

			// now append variable-length string data
			byte[] newrgch = new byte[cchName + 1];    // account for encoding bytes
			System.arraycopy( strbytes, 0, newrgch, 1, cchName );

			data = ByteTools.append( newrgch, data );
		}
		else
		{
			data[6] = -1;
			data[7] = -1;
		}
		setData( data );
	}

	/**
	 * returns true if this pivot item is hidden
	 *
	 * @return
	 */
	public boolean getIsHidden()
	{
		return fHidden;
	}

	/**
	 * sets the hidden state for this pivot item
	 *
	 * @param b
	 */
	public void setIsHidden( boolean b )
	{
		fHidden = b;
		byte by = getByteAt( 2 );
		if( fHidden )
		{
			getData()[2] = (byte) (by & 0x1);
		}
		else
		{
			getData()[2] = (byte) (by ^ 0x1);
		}
	}

	/**
	 * specifies whether the pivot item detail is collapsed.
	 *
	 * @return
	 */
	public boolean getIsCollapsed()
	{
		return fHideDetail;
	}

	/**
	 * specifies whether the pivot item detail is collapsed.
	 *
	 * @return
	 */
	public void setIsCollapsed( boolean b )
	{
		fHideDetail = b;
		byte by = getByteAt( 2 );
		if( fHideDetail )
		{
			getData()[2] = (byte) (by & 0x2);
		}
		else
		{
			getData()[2] = (byte) (by ^ 0x2);
		}
	}
	// TODO: fFormula, fMissing

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0, 0, 	/* itmtype */
			0, 0,	/* flags */
			0, 0,   /* icache */
			-1, -1,	/* cchName */
	};

	public static XLSRecord getPrototype()
	{
		Sxvi si = new Sxvi();
		si.setOpcode( SXVI );
		si.setData( si.PROTOTYPE_BYTES );
		si.init();
		return si;
	}
}