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

/**
 * <b>Blank: a blank cell value 0x201</b><br>
 * <p/>
 * Blank records define a blank cell.  The rw field defines row number (0 based)
 * and the col field defines column number.
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2           Row
 * 6       col         2           Column
 * 8       ixfe        2           Index to the XF record
 * </p></pre>
 */
public final class Blank extends XLSCellRecord
{

	private static final long serialVersionUID = -3847009755105117050L;

	// return a blank string val
	@Override
	public String getStringVal()
	{
		return "";
	}

	//protected byte[] BLANK_CELL_BYTES = { 0, 0, 0, 0, 0, 0};

	public static XLSRecord getPrototype()
	{
		Blank bl = new Blank();
		bl.setData( new byte[]{ 0, 0, 0, 0, 0, 0 } );
		return bl;
	}

	Blank()
	{
		this( new byte[]{ 0, 0, 0, 0, 0, 0 } );
		/*    setData(BLANK_CELL_BYTES);
        setOpcode(BLANK);
        setLength((short)6);
        this.init();
*/
	}

	/**
	 * Provide constructor which automatically
	 * sets the body data and header info.  This
	 * is needed by Mulblank which creates the Blanks without
	 * the benefit of WorkBookFactory.parseRecord().
	 */
	Blank( byte[] b )
	{
		setData( b );
		setOpcode( BLANK );
		setLength( (short) 6 );
		this.init();
	}

	@Override
	public void init()
	{
		super.init();
		int pos = 4;
		super.initRowCol();
		ixfe = ByteTools.readShort( this.getByteAt( pos++ ), this.getByteAt( pos++ ) );
		this.setIsValueForCell( true );
		this.isBlank = true;
	}

	public void setCol( int i )
	{
		if( this.isValueForCell )
		{
			this.getData();
			if( data == null )
			{
				setData( new byte[]{ 0, 0, 0, 0, 0, 0 } );
			}
			byte[] c = ByteTools.shortToLEBytes( (short) i );
			System.arraycopy( c, 0, this.getData(), 2, 2 );
		}
		col = (short) i;
	}

	/**
	 * set the row
	 */
	public void setRow( int i )
	{
		if( this.isValueForCell )
		{
			this.getData();
			if( data == null )
			{
				setData( new byte[]{ 0, 0, 0, 0, 0, 0 } );
			}
			byte[] r = ByteTools.shortToLEBytes( (short) i );
			System.arraycopy( r, 0, this.getData(), 0, 2 );
		}
		rw = i;
	}

}