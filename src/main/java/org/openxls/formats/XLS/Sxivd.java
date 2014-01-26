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

import java.util.Arrays;

/**
 * SXIVD B4h: Array of field ID numbers for the rows and columns in a PT.
 * <p/>
 * There are at most two of these recs per Table -- one for rows, one for cols.
 * <p/>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       rgisxvd     var     Array of 2 byte field ids (row or col)
 * <p/>
 * rgSxivd (variable): An array of SxIvdRw or SxIvdCol items.
 * If this is an array of SxIvdRw,
 * the count of elements in the array MUST equal the value of the cDimRw
 * field of the SxView record.
 * <p/>
 * If this is an array of SxIvdCol,
 * the count of elements in the array MUST equal the value of the cDimCol
 * field of the SxView record.
 * <p/>
 * The SxIvdRw structure specifies a reference to a pivot field or data field on the row axis.
 * The SxIvdCol structure specifies a reference to a pivot field or data field on the column axis
 * rw (2 bytes): A signed integer that specifies a pivot field or data field for the row axis of the PivotTable view.
 * MUST be a value from the following table:
 * or
 * col (2 bytes): A signed integer that specifies a pivot field or data field for the column axis of the PivotTable view.
 * MUST be a value from the following table:
 * <p/>
 * Value	    Meaning
 * -2		    This value specifies that the data field is on the row (or col) axis.
 * The sxaxisRw (or sxaxisCol) field of sxaxis4Data of SxView record
 * MUST equal 1 and the sxaxisData field of sxaxis4Data of the SxView record
 * MUST equal zero.
 * 0+		    This value specifies a pivot field index as specified in Pivot Fields.
 * The pivot field index specifies a pivot field on the row (or col) axis
 * of the PivotTable view.
 * MUST be less than the cDim field of the SxView record of the PivotTable view.
 * If the referenced pivot field is not a hidden field in an OLAP PivotTable view
 * then the sxaxisRw (or sxaxisCol) field of SXAxis of the Sxvd record of the pivot
 * field MUST equal 1.
 * <p/>
 * A pivot field is a hidden field if an SXAddl_SXCField12_SXDVer12Info record exists
 * for the pivot field, and the fHiddenLvl field of the SXAddl_SXCField12_SXDVer12Info record
 * is 1.
 * <p/>
 * <p/>
 * </p></pre>
 */

public class Sxivd extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Sxivd.class );
	private static final long serialVersionUID = 9027599480633995587L;

	@Override
	public void init()
	{
		super.init();
			log.trace( "SXIVD - n: " + getData().length + " array:" + Arrays.toString( getData() ) );
	}

	public String toString()
	{
		return "SXIVD - n: " + getData().length + " array:" + Arrays.toString( getData() );
	}

	/**
	 * adds a field to the end of this field list
	 *
	 * @param fieldNumber
	 */
	public void addField( int fieldNumber )
	{
		data = getData();
		byte[] b = ByteTools.shortToLEBytes( (short) fieldNumber );
		data = ByteTools.append( b, data );
	}

	/**
	 * for each two-byte pair in an array of [number of items]*2 bytes,
	 * //	 * <br>specifies either a pivot field index (must less than the total number of fields as specfied by Sxview)
	 * or -2 means that the data field is on the row or col axis
	 *
	 * @param items
	 */
	public void setRowOrColItems( int[] items )
	{
		byte[] data = intToByteArray( items );
		setData( data );
	}

	public static XLSRecord getPrototype()
	{
		Sxivd si = new Sxivd();
		si.setOpcode( SXIVD );
		si.setData( new byte[]{ } );
		si.init();
		return si;
	}

	public static final byte[] intToByteArray( int[] value )
	{
		if( value == null )
		{
			return null;
		}
		byte[] b = new byte[value.length];
		int j = 0;
		for( int i : value )
		{
			b[j++] = (byte) i;
		}
		return b;
	}
}