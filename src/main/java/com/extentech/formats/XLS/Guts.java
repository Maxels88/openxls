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
 * <b>GUTS: Size of Row and Column Gutters</b><br>
 * <p/>
 * This record stores information about gutter settings
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4		dxRwGut			2		Size of the row gutter
 * 6		dyColGut		2		Size of the Col gutter
 * 8		iLevelRwMac		2		Maximum outline level for row
 * 10		iLevelColMac	2		Maximum outline level for col
 * <p/>
 * </p></pre>
 */
public final class Guts extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2815489536116897500L;
	private short dxRwGut;
	private short dyColGut;
	private short iLevelRwMac;
	private short iLevelColMac;

	public void init()
	{
		super.init();
		dxRwGut = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		dyColGut = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		iLevelRwMac = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		iLevelColMac = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );

		if( DEBUGLEVEL > 5 )
		{
			Logger.logInfo( "INFO: Guts settings: dxRwGut:" + dxRwGut + " dyColGut:" + dyColGut + " iLevelRwMac:" + iLevelRwMac + " iLevelColMac:" + iLevelColMac );
		}
	}

	public void setRowGutterSize( int i )
	{
		dxRwGut = (short) i;
		updateRecBody();
	}

	public int getRowGutterSize()
	{
		return (int) dxRwGut;
	}

	public void setColGutterSize( int i )
	{
		dyColGut = (short) i;
		updateRecBody();
	}

	public int getColGutterSize()
	{
		return (int) dyColGut;
	}

	public void setMaxRowLevel( int i )
	{
		iLevelRwMac = (short) i;
		updateRecBody();
	}

	public int getMaxRowLevel()
	{
		return (int) iLevelRwMac;
	}

	public void setMaxColLevel( int i )
	{
		iLevelColMac = (short) i;
		updateRecBody();
	}

	public int getMaxColLevel()
	{
		return (int) iLevelColMac;
	}

	private void updateRecBody()
	{
		byte[] newbytes = ByteTools.shortToLEBytes( dxRwGut );
		newbytes = ByteTools.append( ByteTools.shortToLEBytes( dyColGut ), newbytes );
		newbytes = ByteTools.append( ByteTools.shortToLEBytes( iLevelRwMac ), newbytes );
		newbytes = ByteTools.append( ByteTools.shortToLEBytes( iLevelColMac ), newbytes );
		this.setData( newbytes );

	}
}
