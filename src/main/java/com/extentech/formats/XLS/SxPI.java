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
 * SXPI 0x86
 * <p/>
 * The SXPI record specifies the pivot fields and information about filtering on
 * the page axis of a PivotTable view. MUST exist if and only if the value of
 * the cDimPg field of the SxView record of the PivotTable view is greater than
 * zero.
 * <p/>
 * rgsxpi (variable): An array of SXPI_Items that specifies the pivot fields and
 * information about filtering on the page axis of a PivotTable view. The number
 * of array elements MUST equal the value of the cDimPg field of the SxView
 * record of the PivotTable view.
 * <p/>
 * The SXPI_Item structure specifies information about a pivot field and its
 * filtering on the page axis of a PivotTable view.
 * <p/>
 * isxvd (2 bytes): A signed integer that specifies a pivot field index as
 * specified by Pivot Fields. The referenced pivot field is specified to be on
 * the page axis. MUST be greater than or equal to zero and less than the cDim
 * field of the SxView record of the PivotTable view.
 * <p/>
 * isxvi (2 bytes): A signed integer that specifies the pivot item used for the
 * page axis filtering. MUST be a value from the following table: Value Meaning
 * 0x0000 to 0x7FFC This value specifies a pivot item index that specifies a
 * pivot item in the pivot field specified by isxvd. The referenced pivot item
 * specifies the page axis filtering for the pivot field. 0x7FFD This value
 * specifies all pivot items, see page axis for filtering that applies. For a
 * non-OLAP PivotTable view the value MUST be 0x7FFD or greater than or equal to
 * zero and less than the cItm field of the Sxvd record of the pivot field.
 * Otherwise the value MUST be 0x7FFD.
 * <p/>
 * idObj (2 bytes): A signed integer that specifies the object identifier of the
 * Obj record with the page item drop-down arrow.
 */
public class SxPI extends XLSRecord implements XLSConstants
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2639291289806138985L;
	private SXPI_Item[] sxpis;

	@Override
	public void init()
	{
		super.init();
		byte[] rgsxpi = this.getData();    // each item is 6 bytes
		if( rgsxpi != null )
		{
			if( (rgsxpi.length % 6) != 0 )
			{
				Logger.logWarn( "PivotTable: Irregular SxPI structure" );
			}
			sxpis = new SXPI_Item[rgsxpi.length / 6];
			for( int j = 0; j < sxpis.length; j++ )
			{
				sxpis[j] = new SXPI_Item( rgsxpi, j * 6 );
			}
			if( DEBUGLEVEL > 3 )
			{
				Logger.logInfo( "SXPI - n: " + sxpis.length + ": " + Arrays.toString( data ) );
			}
		}
		else if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "SXPI - NULL" );
		}
	}

	public String toString()
	{
		if( sxpis != null )
		{
			return "SXPI - n: " + sxpis.length + ": " + Arrays.toString( sxpis );
		}
		else
		{
			return "SXPI - NULL";
		}
	}

	/**
	 * returns the pivot field index and item index for page axis field i
	 *
	 * @param i page axis field
	 * @return
	 */
	public int[] getPivotFieldItem( int i )
	{
		if( (i >= 0) && (i < sxpis.length) )
		{
			return new int[]{ sxpis[i].isxvd, sxpis[i].idObj };
		}
		return new int[]{ -1, -1 };
	}

	/**
	 * sets the pivot field index and item index for page axis field i
	 *
	 * @param i
	 * @param fieldindex
	 * @param itemindex
	 */
	public void setPageFieldIndex( int i, int fieldindex, int itemindex )
	{
		if( (i >= 0) && (i < sxpis.length) )
		{
			sxpis[i].isxvd = (short) fieldindex;
			sxpis[i].isxvi = (short) itemindex;
			byte[] b = ByteTools.shortToLEBytes( (short) fieldindex );
			this.getData()[(i * 6)] = b[0];
			this.getData()[(i * 6) + 1] = b[1];
			b = ByteTools.shortToLEBytes( (short) itemindex );
			this.getData()[(i * 6) + 2] = b[0];
			this.getData()[(i * 6) + 3] = b[1];
		}
	}

	public static XLSRecord getPrototype()
	{
		SxPI sp = new SxPI();
		sp.setOpcode( SXPI );
		// no data, initially???
		sp.setData( new byte[]{ } );
		sp.init();
		return sp;
	}

	/**
	 * add a pivot field to the page axis
	 *
	 * @param fieldIndex
	 * @param itemIndex
	 */
	public void addPageField( int fieldIndex, int itemIndex )
	{
		getData();
		data = ByteTools.append( new byte[6], data );
		SXPI_Item[] tmp = new SXPI_Item[sxpis.length + 1];
		System.arraycopy( sxpis, 0, tmp, 0, sxpis.length );
		sxpis = tmp;
		sxpis[sxpis.length - 1] = new SXPI_Item();

		setPageFieldIndex( sxpis.length - 1, fieldIndex, itemIndex );

	}

}

/**
 * helper class defines SxPI structure
 */
class SXPI_Item
{
	public short isxvd;    // pivot field index
	public short isxvi;    // specifies the pivot item used for the page axis filtering; if 0x7FFD This valuespecifies all pivot items, otherwise it's item index of pivot field (isxvd)
	public short idObj;

	SXPI_Item()
	{
	}

	SXPI_Item( byte[] rgsxpi, int idx )
	{
		isxvd = ByteTools.readShort( rgsxpi[idx++], rgsxpi[idx++] );
		isxvi = ByteTools.readShort( rgsxpi[idx++], rgsxpi[idx++] );
		idObj = ByteTools.readShort( rgsxpi[idx++], rgsxpi[idx++] );
	}
}