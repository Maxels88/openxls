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

import com.extentech.formats.LEO.InvalidFileException;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <pre><b>Bof: Beginning of File Stream 0x809</b><br>
 * <p/>
 * Marks the beginning of an XLS file Stream including Boundsheets
 * <p/>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 0       vers        1       version:
 * 1       bof         1       0x09
 * ...
 * <p/>
 * </p></pre>
 */
public final class Bof extends UnencryptedXLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Bof.class );
	private static final long serialVersionUID = 3005631881544437570L;
	short grbit;
	String xlsver = "";
	int oldlen = -1;

	public String toString()
	{
		return super.toString() + " lbplypos: " + getLbPlyPos();
	}

	/**
	 * Set the offset for this BOF
	 */
	@Override
	public void setOffset( int s )
	{
		super.setOffset( s );
		if( worksheet != null )
		{
			if( isSheetBof() || isVBModuleBof() || (worksheet.getMyBof().equals( this )) )
			{
				worksheet.setLbPlyPos( getLbPlyPos() );
			}
		}
	}

	boolean isValidBIFF8()
	{
		return oldlen == 20;
	}

	String getXLSVersionString()
	{
		return xlsver;
	}

	/**
	 * Initialize the BOF record
	 */
	@Override
	public void init()
	{
		super.init();
		getData();
		short version = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		// dt
		grbit = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		short rupBuild = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		short rupYear = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );

		byte b12 = getByteAt( 12 );
		byte b13 = getByteAt( 13 );
		byte b14 = getByteAt( 14 );
		byte b15 = getByteAt( 15 );
		int compat = ByteTools.readInt( b12, b13, b14, b15 ); // 1996
		xlsver = compat + "";
		oldlen = getLength();
		if( oldlen < 16 )
		{
			log.error( "Not Excel '97 (BIFF8) or later version.  Unsupported file format." );
			throw new InvalidFileException( "InvalidFileException: Not Excel '97 (BIFF8) or later version.  Unsupported file format." );
		}
	}

	/**
	 * this is equal to the lbPlyPos stored in
	 * the Boundsheet associated with this Bof
	 */
	long getLbPlyPos()
	{
		if( !isValidBIFF8() )
		{
			return offset + 8;
		}
		return offset;
	}

	/**
	 * @return Returns the sheetbof.
	 */
	public boolean isSheetBof()
	{
		if( (grbit & 0x10) == 0x10 )
		{
			return true;
		}
		return false;
	}

	public boolean isVBModuleBof()
	{
		if( (grbit & 0x06) == 0x06 )
		{
			return true;
		}
		return false;

	}

	/**
	 * @return Returns the sheetbof.
	 */
	public boolean isChartBof()
	{
		if( (grbit & 0x20) == 0x20 )
		{
			return true;
		}
		return false;
	}

	/**
	 * @param sheetbof The sheetbof to set.
	 */
	public void setSheetBof()
	{
		grbit = 0x10;
	}
}