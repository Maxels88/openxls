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

//0xF00A
public class MsofbtSp extends EscherRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -5355585244930369889L;
	int id;
	int grfPersistence;

	public MsofbtSp( int fbt, int inst, int version )
	{
		super( fbt, inst, version );
	}

	@Override
	public byte[] getData()
	{
		byte[] idBytes, flagBytes;
		idBytes = ByteTools.cLongToLEBytes( id );
		flagBytes = ByteTools.cLongToLEBytes( grfPersistence );
		byte[] retData = new byte[8];
		System.arraycopy( idBytes, 0, retData, 0, 4 );
		System.arraycopy( flagBytes, 0, retData, 4, 4 );

		this.setLength( 8 );
		return retData;

	}

	public void setId( int value )
	{

		id = value;
	}

	public void setGrfPersistence( int value )
	{
		grfPersistence = value;
	}
}
