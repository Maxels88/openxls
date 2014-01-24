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

//0xF008
public class MsofbtDg extends EscherRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5218802290529676567L;
	int csp;
	int lastSPID;

	public MsofbtDg( int fbt, int inst, int version )
	{
		super( fbt, inst, version );
	}

	@Override
	protected byte[] getData()
	{
		byte[] cspBytes;
		byte[] spIdBytes;

		cspBytes = ByteTools.cLongToLEBytes( csp );        // Number of shapes
		spIdBytes = ByteTools.cLongToLEBytes( lastSPID );    // last SPID

		byte[] retBytes = new byte[8];
		System.arraycopy( cspBytes, 0, retBytes, 0, 4 );
		System.arraycopy( spIdBytes, 0, retBytes, 4, 4 );

		setLength( retBytes.length );
		return retBytes;
	}

	public void setNumShapes( int value )
	{
		csp = value;
	}

	public void setLastSPID( int value )
	{

		lastSPID = value;
	}
}
