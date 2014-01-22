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
package com.extentech.formats.escher;

import com.extentech.toolkit.ByteTools;

//0xf009

/*
 * This record is present only in group shapes (not shapes in groups, shapes that are groups).
 * The group shape record defines the coordinate system of the shape, which the anchors of the 
 * child shape are expressed in. All other information is stored in the shape records that follow. 
 */
public class MsofbtSpgr extends EscherRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5591214948365806058L;
	int left = 0, top = 0, right = 0, bottom = 0;

	public MsofbtSpgr( int fbt, int inst, int version )
	{
		super( fbt, inst, version );
	}

	@Override
	protected byte[] getData()
	{
		byte[] leftBytes, topBytes, rightBytes, bottomBytes;
		leftBytes = ByteTools.cLongToLEBytes( left );
		topBytes = ByteTools.cLongToLEBytes( top );
		rightBytes = ByteTools.cLongToLEBytes( right );
		bottomBytes = ByteTools.cLongToLEBytes( bottom );

		byte[] retBytes = new byte[16];
		System.arraycopy( leftBytes, 0, retBytes, 0, 4 );
		System.arraycopy( topBytes, 0, retBytes, 4, 4 );
		System.arraycopy( rightBytes, 0, retBytes, 8, 4 );
		System.arraycopy( bottomBytes, 0, retBytes, 12, 4 );

		this.setLength( retBytes.length );

		return retBytes;
	}

	public void setRect( int left, int top, int right, int bottom )
	{
		this.left = left;
		this.right = right;
		this.bottom = bottom;
		this.top = top;

	}

}
