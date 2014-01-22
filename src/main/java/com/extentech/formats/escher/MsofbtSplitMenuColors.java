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

public class MsofbtSplitMenuColors extends EscherRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5888748984363726576L;
	//These values are from experimental records.
	int fillColor = 0x800000D, lineColor = 0x800000C, shadowColor = 0x8000017, _3dColor = 0x100000f7;

	public MsofbtSplitMenuColors( int fbt, int inst, int version )
	{
		super( fbt, inst, version );
	}

	// @Override
	@Override
	protected byte[] getData()
	{
		byte[] fillColorBytes = ByteTools.cLongToLEBytes( fillColor );
		byte[] lineColorBytes = ByteTools.cLongToLEBytes( lineColor );
		byte[] shadowColorBytes = ByteTools.cLongToLEBytes( shadowColor );
		byte[] _3dColorBytes = ByteTools.cLongToLEBytes( _3dColor );

		byte[] totalBytes = new byte[16];

		System.arraycopy( fillColorBytes, 0, totalBytes, 0, 4 );
		System.arraycopy( lineColorBytes, 0, totalBytes, 4, 4 );
		System.arraycopy( shadowColorBytes, 0, totalBytes, 8, 4 );
		System.arraycopy( _3dColorBytes, 0, totalBytes, 12, 4 );

		this.setLength( 16 );
		this.setInst( 4 );

		return totalBytes;
	}

	public int get3dColor()
	{
		return _3dColor;
	}

	public void set3dColor( int color )
	{
		_3dColor = color;
	}

	public int getFillColor()
	{
		return fillColor;
	}

	public void setFillColor( int fillColor )
	{
		this.fillColor = fillColor;
	}

	public int getLineColor()
	{
		return lineColor;
	}

	public void setLineColor( int lineColor )
	{
		this.lineColor = lineColor;
	}

	public int getShadowColor()
	{
		return shadowColor;
	}

	public void setShadowColor( int shadowColor )
	{
		this.shadowColor = shadowColor;
	}

}
