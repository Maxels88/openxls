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

import java.io.ByteArrayOutputStream;

//0xf010
public class MsofbtClientAnchor extends EscherRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 7946989934447191433L;
	// 20070914 KSC: this record doesn't store sheet index; 1st two bytes are a flag (seems always to be 2)
	short flag = 2, leftColumnIndex, xOffsetL, topRowIndex, yOffsetT, rightColIndex, xOffsetR, bottomRowIndex, yOffsetB;

	public MsofbtClientAnchor( int fbt, int inst, int version )
	{
		super( fbt, inst, version );
	}

	public byte[] getData()
	{
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		try
		{
			bos.write( ByteTools.shortToLEBytes( flag ) );
			bos.write( ByteTools.shortToLEBytes( leftColumnIndex ) );
			bos.write( ByteTools.shortToLEBytes( xOffsetL ) );
			bos.write( ByteTools.shortToLEBytes( topRowIndex ) );
			bos.write( ByteTools.shortToLEBytes( yOffsetT ) );
			bos.write( ByteTools.shortToLEBytes( rightColIndex ) );
			bos.write( ByteTools.shortToLEBytes( xOffsetR ) );
			bos.write( ByteTools.shortToLEBytes( bottomRowIndex ) );
			bos.write( ByteTools.shortToLEBytes( yOffsetB ) );
		}
		catch( Exception e )
		{

		}
		this.setLength( bos.toByteArray().length );
		return bos.toByteArray();

	}

	public void setBounds( short[] bounds )
	{
		if( bounds == null )
		{
			return;
		}
		leftColumnIndex = bounds[0];
		xOffsetL = bounds[1];
		topRowIndex = bounds[2];
		yOffsetT = bounds[3];
		rightColIndex = bounds[4];
		xOffsetR = bounds[5];
		bottomRowIndex = bounds[6];
		yOffsetB = bounds[7];

	}

	public void setFlag( short flag )
	{
		this.flag = flag;
	}

}
