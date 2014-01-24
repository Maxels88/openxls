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

/**
 * <b>Dimensions 0x200: Describes the max BiffRec dimensions of a Sheet.</b><br>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       rowMic      4       First defined Row
 * 8       rowMac      4       Last defined Row plus 1
 * 12      colFirst    2       First defined column
 * 14      colLast     2       Last defined column plus 1
 * 16      reserved    2       must be 0
 * <p/>
 * When a record is added to a Boundsheet, we need to check if this
 * has changed the row/col dimensions of the sheet and update this
 * record accordingly.
 * <p/>
 * </p></pre>
 *
 * @see WorkBook
 * @see Boundsheet
 * @see Index
 * @see Dbcell
 * @see Row
 * @see Cell
 * @see XLSRecord
 */

public class Dimensions extends com.extentech.formats.XLS.XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Dimensions.class );
	private static final long serialVersionUID = 7156425132146869228L;
	int rowFirst = 0;
	int rowLast = 0;
	short colFirst = 0;
	short colLast = 0;

	/**
	 * set last/first cols/rows
	 */
	public void setRowFirst( int c )
	{
		c++; // inc here instead of updateRowDimensions
		byte[] b = ByteTools.cLongToLEBytes( c );
		byte[] dt = getData();
		if( dt.length > 4 )
		{
			System.arraycopy( b, 0, dt, 0, 4 );
		}
		rowFirst = c;
	}

	public int getRowFirst()
	{
		return rowFirst;
	}

	/**
	 * set last/first cols/rows
	 */
	public void setRowLast( int c )
	{
		c++; // inc here instead of updateRowDimensions
		byte[] b = ByteTools.cLongToLEBytes( c );
		byte[] dt = getData();
		if( dt.length > 4 )
		{
			System.arraycopy( b, 0, dt, 4, 4 );
		}
		rowLast = c;
	}

	public int getRowLast()
	{
		return rowLast;
	}

	/**
	 * set last/first cols/rows
	 */
	public void setColFirst( int c )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = getData();
		if( dt.length > 4 )
		{
			System.arraycopy( b, 0, dt, 8, 2 );
		}
		colFirst = (short) c;
	}

	public int getColFirst()
	{
		return colFirst;
	}

	/**
	 * set last/first cols/rows
	 */
	public void setColLast( int c )
	{
		c++;
		if( (c >= MAXCOLS_BIFF8) && !wkbook.getIsExcel2007() )
		{
			log.warn( "Dimensions.setColLast column: " + c + " is incompatible with pre Excel2007 versions." );
		}

		if( c >= MAXCOLS )
		{
			c = MAXCOLS;// odd case, its supposed to be last defined col +1, but this breaks last col
		}
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = getData();
		if( dt.length > 4 )
		{
			System.arraycopy( b, 0, dt, 10, 2 );
		}
		colLast = (short) c;
	}

	public int getColLast()
	{
		return colLast;
	}

	@Override
	public void setSheet( Sheet bs )
	{
		super.setSheet( bs );
		bs.setDimensions( this );
	}

	@Override
	public void init()
	{
		super.init();
		rowFirst = ByteTools.readInt( getByteAt( 0 ), getByteAt( 1 ), getByteAt( 2 ), getByteAt( 3 ) );
		rowLast = ByteTools.readInt( getByteAt( 4 ), getByteAt( 5 ), getByteAt( 6 ), getByteAt( 7 ) );

		colFirst = ByteTools.readShort( getByteAt( 8 ), getByteAt( 9 ) );
		colLast = ByteTools.readShort( getByteAt( 10 ), getByteAt( 11 ) );
		getData();
	}

	/**
	 * update the min/max cols and rows
	 */
	public void updateDimensions( int row, short col )
	{
		updateRowDimensions( row );
		updateColDimension( col );
	}

	/**
	 * update the min/max cols and rows
	 */
	public void updateRowDimensions( int row )
	{
		// check row dimension
		//row++; // TODO: check why we are incrementing here... nowincremented in setRowLast
		if( row >= rowLast )
		{
			setRowLast( row ); // now incremented only in setRowXX
		}
		if( row < rowFirst )
		{
			setRowFirst( row ); // now incremented only in setRowXX
		}
	}

	/**
	 * update the min/max cols and rows
	 */
	public void updateColDimension( short col )
	{
		// check cell dimension
		if( col > colLast )
		{
			setColLast( col );
		}
		else if( (col - 1) < colFirst )
		{
			setColFirst( col );
		}
	}
}