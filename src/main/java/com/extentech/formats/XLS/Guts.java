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
	private static final Logger log = LoggerFactory.getLogger( Guts.class );
	private static final long serialVersionUID = 2815489536116897500L;
	private short dxRwGut;
	private short dyColGut;
	private short iLevelRwMac;
	private short iLevelColMac;

	@Override
	public void init()
	{
		super.init();
		dxRwGut = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		dyColGut = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		iLevelRwMac = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		iLevelColMac = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );

			log.debug( "INFO: Guts settings: dxRwGut:" + dxRwGut + " dyColGut:" + dyColGut + " iLevelRwMac:" + iLevelRwMac + " iLevelColMac:" + iLevelColMac );
	}

	public void setRowGutterSize( int i )
	{
		dxRwGut = (short) i;
		updateRecBody();
	}

	public int getRowGutterSize()
	{
		return dxRwGut;
	}

	public void setColGutterSize( int i )
	{
		dyColGut = (short) i;
		updateRecBody();
	}

	public int getColGutterSize()
	{
		return dyColGut;
	}

	public void setMaxRowLevel( int i )
	{
		iLevelRwMac = (short) i;
		updateRecBody();
	}

	public int getMaxRowLevel()
	{
		return iLevelRwMac;
	}

	public void setMaxColLevel( int i )
	{
		iLevelColMac = (short) i;
		updateRecBody();
	}

	public int getMaxColLevel()
	{
		return iLevelColMac;
	}

	private void updateRecBody()
	{
		byte[] newbytes = ByteTools.shortToLEBytes( dxRwGut );
		newbytes = ByteTools.append( ByteTools.shortToLEBytes( dyColGut ), newbytes );
		newbytes = ByteTools.append( ByteTools.shortToLEBytes( iLevelRwMac ), newbytes );
		newbytes = ByteTools.append( ByteTools.shortToLEBytes( iLevelColMac ), newbytes );
		setData( newbytes );

	}
}
