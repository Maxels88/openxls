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
package org.openxls.formats.escher;

import org.openxls.toolkit.ByteTools;

import java.io.Serializable;

abstract class EscherRecord implements Serializable
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7987132917889379656L;
	byte[] header;
	byte[] data;
	int length;
	int inst;
	int fbt;
	int version;
	boolean isDirty = false;

	/**
	 * no param constructor for Serializable
	 */
	public void EscherRecord()
	{
	}

	EscherRecord( int fbt, int inst, int version )
	{
		this.fbt = fbt;
		this.inst = inst;
		this.version = version;
	}

	public int getFbt()
	{

		return fbt;
	}

	public int getInst()
	{

		return inst;
	}

	public int getLength()
	{
		return length;
	}

	public void setFbt( int value )
	{
		fbt = value;
	}

	public void setInst( int value )
	{
		inst = value;
	}

	public void setLength( int value )
	{
		length = value;
	}

	protected abstract byte[] getData();

	private byte[] getHeaderBytes()
	{
		//TODO: Reverse the process of header decoding here
		byte[] headerBytes = new byte[4];

		headerBytes[0] = (byte) ((0xF & version) | (0xF0 & (inst << 4)));
		headerBytes[1] = (byte) ((0x00000FF0 & inst) >> 4);
		headerBytes[2] = (byte) ((0x000000FF & fbt));
		headerBytes[3] = (byte) ((0x0000FF00 & fbt) >> 8);

		int version2 = (0x0F & headerBytes[0]);
		int inst2 = ((0xFF & headerBytes[1]) >> 4) | ((0xF0 & headerBytes[0]) >> 4);

		byte[] lenBytes = ByteTools.cLongToLEBytes( length );

		byte[] retData = new byte[8];
		System.arraycopy( headerBytes, 0, retData, 0, 4 );
		System.arraycopy( lenBytes, 0, retData, headerBytes.length, 4 );

		return retData;
	}

	public byte[] toByteArray()
	{
		byte[] dataBytes = getData();  //Have it in this sequence as some records adjust their header byte length from getData
		byte[] headerBytes = getHeaderBytes();

		byte[] retData = new byte[headerBytes.length + dataBytes.length];
		System.arraycopy( headerBytes, 0, retData, 0, headerBytes.length );
		if( dataBytes.length > 0 )
		{
			System.arraycopy( dataBytes, 0, retData, headerBytes.length, dataBytes.length );
		}
		return retData;

	}
}
