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

import java.io.UnsupportedEncodingException;

/**
 * DConName 0x52
 * <p/>
 * The DConName record specifies a named range that is a data source (1) for a
 * PivotTable or a data source (1) for the data consolidation settings of the
 * associated sheet. The range is specified as a reference to an external
 * workbook or a defined name in this workbook. If the named range is in an
 * external workbook, this record specifies the path to the external workbook.
 * If the named range has a defined name that has a sheet-level scope, this
 * record also specifies the name of the sheet that contains the range.
 * <p/>
 * stName (variable): An XLNameUnicodeString structure that specifies a defined name for the source range.
 * <p/>
 * cchFile (2 bytes): An unsigned integer that specifies the character count of stFile. MUST be 0x0000, or greater than or equal to 0x0002.
 * A value of 0x0000 specifies that the defined name in stName has a workbook scope and is contained in this file.
 * <p/>
 * stFile (variable): A DConFile structure that specifies the workbook, or workbook and sheet, that contains the range specified in stName.
 * This field exists only if the value of cchFile is greater than zero. If the defined name in stName has workbook scope,
 * this field specifies the workbook file that contains the defined name and its associated range.
 * If the defined name in stName has a sheet-level scope, this field specifies both the sheet name and the workbook that
 * contains the defined name and its associated range.
 * <p/>
 * unused (variable): An array of bytes that is unused and MUST be ignored. MUST exist if and only if cchFile is greater than 0 and
 * stFile specifies a self-reference (the value of stFile.stFile.rgb[0] is 2). If the value stFile.stFile.fHighByte is 0,
 * the size of this array is 1. If the value of stFile.stFile.fHighByte is 1, the size of this array is 2.
 */

public class DConName extends XLSRecord implements XLSConstants
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2639291289806138985L;
	private short cchFile;
	private String namedRange = null;

	@Override
	public void init()
	{
		super.init();
		int cch = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		int pos = 1;
		if( cch > 0 )
		{
			//A - fHighByte (1 bit): A bit that specifies whether the characters in rgb are double-byte characters.
			// 0x0  All the characters in the string have a high byte of 0x00 and only the low bytes are in rgb.
			// 0x1  All the characters in the string are saved as double-byte characters in rgb.
			// reserved (7 bits): MUST be zero, and MUST be ignored.

			byte encoding = this.getByteAt( ++pos );
			byte[] tmp = this.getBytesAt( ++pos, (cch) * (encoding + 1) );
			try
			{
				if( encoding == 0 )
				{
					namedRange = new String( tmp, DEFAULTENCODING );
				}
				else
				{
					namedRange = new String( tmp, UNICODEENCODING );
				}
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logInfo( "encoding PivotTable name in DCONNAME: " + e );
			}
		}
		cchFile = ByteTools.readShort( this.getByteAt( pos + cch ), this.getByteAt( pos + cch + 1 ) );
		// either 0 or >=2
		if( cchFile > 0 )
		{
			Logger.logWarn( "PivotTable: External Workbooks for Named Range Source are Unspported" );
		}
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "DCONNAME: namedRange:" + namedRange + " cchFile: " + cchFile );
		}

	}

	/**
	 * return the Named Range data source for a pivot table
	 *
	 * @return
	 */
	public String getNamedRange()
	{
		return this.namedRange;
	}

	/**
	 * sets the named range for the data source for the pivot table
	 *
	 * @param namedRange
	 */
	public void setNamedRange( String namedRange )
	{
		int cch = (short) ((short) namedRange.length() + 1);
		this.namedRange = namedRange;
		byte[] data = new byte[3];
		byte[] b = ByteTools.shortToLEBytes( cchFile );
		data[0] = b[0];
		data[1] = b[1];
		try
		{
			data = ByteTools.append( namedRange.getBytes( DEFAULTENCODING ), data );
		}
		catch( UnsupportedEncodingException e )
		{
		}
		data = ByteTools.append( new byte[]{ 0, 0 }, data );    // data in same workbook
		this.setData( data );
	}

	/**
	 * create a new DCONMANE - named range data source for a pivot table
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		DConName dn = new DConName();
		dn.setOpcode( DCONNAME );
		dn.setData( new byte[]{ 0, 0 } );
		dn.init();
		return dn;
	}
}
