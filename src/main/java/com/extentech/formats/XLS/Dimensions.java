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

	/**
	 * serialVersionUID
	 */
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
		byte[] dt = this.getData();
		if( dt.length > 4 )
		{
			System.arraycopy( b, 0, dt, 0, 4 );
		}
		this.rowFirst = c;
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
		byte[] dt = this.getData();
		if( dt.length > 4 )
		{
			System.arraycopy( b, 0, dt, 4, 4 );
		}
		this.rowLast = c;
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
		byte[] dt = this.getData();
		if( dt.length > 4 )
		{
			System.arraycopy( b, 0, dt, 8, 2 );
		}
		this.colFirst = (short) c;
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
		if( c >= MAXCOLS_BIFF8 && // warn about maxcols
				!this.wkbook.getIsExcel2007() )
		{
			Logger.logWarn( "Dimensions.setColLast column: " + c + " is incompatible with pre Excel2007 versions." );
		}

		if( c >= MAXCOLS )
		{
			c = MAXCOLS;// odd case, its supposed to be last defined col +1, but this breaks last col
		}
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = this.getData();
		if( dt.length > 4 )
		{
			System.arraycopy( b, 0, dt, 10, 2 );
		}
		this.colLast = (short) c;
	}

	public int getColLast()
	{
		return colLast;
	}

	public void setSheet( Sheet bs )
	{
		super.setSheet( bs );
		bs.setDimensions( this );
	}

	public void init()
	{
		super.init();
		rowFirst = ByteTools.readInt( this.getByteAt( 0 ), this.getByteAt( 1 ), this.getByteAt( 2 ), this.getByteAt( 3 ) );
		rowLast = ByteTools.readInt( this.getByteAt( 4 ), this.getByteAt( 5 ), this.getByteAt( 6 ), this.getByteAt( 7 ) );

		colFirst = ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
		colLast = ByteTools.readShort( this.getByteAt( 10 ), this.getByteAt( 11 ) );
		this.getData();
	}

	/**
	 * update the min/max cols and rows
	 */
	public void updateDimensions( int row, short col )
	{
		this.updateRowDimensions( row );
		this.updateColDimension( col );
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
			this.setRowLast( row ); // now incremented only in setRowXX
		}
		if( row < rowFirst )
		{
			this.setRowFirst( row ); // now incremented only in setRowXX
		}

		if( DEBUGLEVEL > 10 )
		{
			String shtnm = this.getSheet().getSheetName();
			Logger.logInfo( shtnm + " dimensions: " + rowFirst + ":" + colFirst + "-" + rowLast + ":" + colLast );
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
			this.setColLast( col );
		}
		else if( (col - 1) < colFirst )
		{
			this.setColFirst( col );
		}
		if( DEBUGLEVEL > 10 )
		{
			String shtnm = this.getSheet().getSheetName();
			Logger.logInfo( shtnm + " dimensions: " + rowFirst + ":" + colFirst + "-" + rowLast + ":" + colLast );
		}
	}

}