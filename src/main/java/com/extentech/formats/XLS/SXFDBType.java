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
/**
 * The SXFDBType record specifies the type of data contained in this cache field.
 *
 * wTypeSql (2 bytes): An ODBCType structure that specifies the ODBC data type as returned by the ODBC provider of the data in this cache field. 
 */

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.util.Arrays;

public class SXFDBType extends XLSRecord implements XLSConstants, PivotCacheRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 9027599480633995587L;
	private short wTypeSql;

	/**
	 * The ODBCType structure specifies an ODBC data type identifier.
	 */
	enum ODBCType
	{
		SQL_TYPE_NULL( 0 ), /**
	 * Undetermined type, data source (1) does not support typed data. Data type determined based on data content: date and time, decimal or text.
	 */
	SQL_CHAR( 1 ), /**
	 * Fixed-length string of ANSI characters
	 */
	SQL_DECIMAL( 3 ), /**
	 * Fixed-precision, Fixed-scale numbers
	 */
	SQL_INTEGER( 4 ), /**
	 * 32-bit signed integer
	 */
	SQL_SMALLINT( 5 ), /**
	 * 16-bit signed integer
	 */
	SQL_FLOAT( 6 ), /**
	 * User-specified precision floating-point
	 */
	SQL_REAL( 7 ), /**
	 * 7-digits precision floating-point
	 */
	SQL_DOUBLE( 8 ), /**
	 * 15-digits precision floating-point
	 */
	SQL_TIMESTAMP( 0xB ), /**
	 * Date and Time
	 */
	SQL_VARCHAR( 0xC ), /**
	 * Variable-length string of ANSI characters
	 */
	SQL_BIT( 0xFFF9 ), /**
	 * Bit (1 or 0)
	 */
	SQL_BINARY( 0xFFFE );
		/**
		 * Fixed-length binary data
		 */

		private final short type;

		ODBCType( int type )
		{
			this.type = (short) type;
		}

		public short type()
		{
			return type;
		}

		public static ODBCType get( int type )
		{
			for( ODBCType c : values() )
			{
				if( c.type == type )
				{
					return c;
				}
			}
			return null;
		}
	}

	public void init()
	{
		super.init();
		wTypeSql = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "SXFDBType -" + Arrays.toString( this.getData() ) );
		}
	}

	/**
	 * creates new, default SXFDBType
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		SXFDBType sf = new SXFDBType();
		sf.setOpcode( SXFDBTYPE );
		sf.setData( new byte[]{ 0, 0 } );
		sf.init();
		return sf;
	}

	/**
	 * set the type of the corresponding cache field
	 *
	 * @param type
	 * @see ODBCType
	 */
	public void setType( int type )
	{
		wTypeSql = (short) type;
		byte[] b = ByteTools.shortToLEBytes( wTypeSql );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
	}

	public String toString()
	{
		return "SXFDBType: " + wTypeSql +
				Arrays.toString( this.getRecord() );
	}

	public int getType()
	{
		return wTypeSql;
	}

	/**
	 * return the bytes describing this record, including the header
	 *
	 * @return
	 */
	public byte[] getRecord()
	{
		byte[] b = new byte[4];
		System.arraycopy( ByteTools.shortToLEBytes( this.getOpcode() ), 0, b, 0, 2 );
		System.arraycopy( ByteTools.shortToLEBytes( (short) this.getData().length ), 0, b, 2, 2 );
		return ByteTools.append( this.getData(), b );

	}
}
