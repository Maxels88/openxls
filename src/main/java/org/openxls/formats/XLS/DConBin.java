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
package org.openxls.formats.XLS;

import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * DConBin 0x1B5
 * <p/>
 * The DConBin record specifies a built-in named range that is a data source (1)
 * for a PivotTable or a data source (1) for the data consolidation settings of
 * the associated sheet.
 * <p/>
 * nBuiltin (1 byte): An unsigned integer that specifies the built-in defined
 * name for the range. MUST be a value from the following table:
 * Value 		Meaning
 * 0x00 		"Consolidate_Area"
 * 0x01 		"Auto_Open"
 * 0x02 		"Auto_Close"
 * 0x03 		"Extract"
 * 0x04 		"Database"
 * 0x05 		"Criteria"
 * 0x06 		"Print_Area"
 * 0x07 		"Print_Titles"
 * 0x08		"Recorder"
 * 0x09 		"Data_Form"
 * 0x0A 		"Auto_Activate"
 * 0x0B		"Auto_Deactivate"
 * 0x0C		"Sheet_Title"
 * 0x0D 		"_FilterDatabase"
 * <p/>
 * reserved1 (2 bytes): MUST be zero and MUST be ignored.
 * <p/>
 * reserved2 (1 byte): MUST be zero and MUST be ignored.
 * <p/>
 * cchFile (2 bytes): An unsigned integer that specifies the character count of
 * stFile. MUST be 0x0000, or greater than or equal to 0x0002. A value of 0x0000
 * specifies that the built-in defined name specified in nBuiltin has a workbook
 * scope and is contained in this file.
 * <p/>
 * stFile (variable): An DConFile structure that specifies the workbook or
 * workbook and sheet that contains the range specified in nBuiltin. This field
 * MUST exist if and only if the value of cchFile is greater than zero. If the
 * built-in defined name has workbook scope this field specifies the workbook
 * file that contains the built-in defined name and its associated range. If the
 * built-in defined name has a sheet–level scope this field specifies both the
 * sheet name and the workbook file that contains the built-in defined name and
 * its associated range.
 * <p/>
 * unused (variable): An array of bytes that is unused and MUST be ignored. MUST
 * exist if and only if cchFile is greater than 0 and stFile specifies a
 * self-reference (the value of stFile.stFile.rgb[0] is 2). If the value
 * stFile.stFile.fHighByte is 0 the size of this array is 1. If the value of
 * stFile.stFile.fHighByte is 1 the size of this array is 2.
 */
public class DConBin extends XLSRecord implements XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( DConBin.class );
	private static final long serialVersionUID = 2639291289806138985L;
	private byte nBuiltin;
	private short cchFile;

	@Override
	public void init()
	{
		super.init();
		nBuiltin = getByteAt( 0 );
		cchFile = ByteTools.readShort( getByteAt( 5 ), getByteAt( 6 ) );
		if( cchFile > 0 )
		{
			log.warn( "PivotTable: External Workbooks for Built-in Named Range Source are Unsupported" );
		}
			log.debug( "DCONBIN: nBuiltin:" + nBuiltin + " cchFile: " + cchFile );
	}

	/**
	 * returns the built-in type for the built-in named range data source for a pivot table
	 * <br>Value 		Meaning
	 * <li>0x00 		"Consolidate_Area"
	 * <li>0x01 		"Auto_Open"
	 * <li>0x02 		"Auto_Close"
	 * <li>0x03 		"Extract"
	 * <li>0x04 		"Database"
	 * <li>0x05 		"Criteria"
	 * <li>0x06 		"Print_Area"
	 * <li>0x07 		"Print_Titles"
	 * <li>0x08		"Recorder"
	 * <li>0x09 		"Data_Form"
	 * <li>0x0A 		"Auto_Activate"
	 * <li>0x0B		"Auto_Deactivate"
	 * <li>0x0C		"Sheet_Title"
	 * <li>0x0D 		"_FilterDatabase"
	 *
	 * @return
	 */
	public int getBuiltInType()
	{
		return nBuiltin;
	}

	/**
	 * sets the built-in type for the built-in named range data source for a pivot table
	 * <br>Value 		Meaning
	 * <li>0x00 		"Consolidate_Area"
	 * <li>0x01 		"Auto_Open"
	 * <li>0x02 		"Auto_Close"
	 * <li>0x03 		"Extract"
	 * <li>0x04 		"Database"
	 * <li>0x05 		"Criteria"
	 * <li>0x06 		"Print_Area"
	 * <li>0x07 		"Print_Titles"
	 * <li>0x08		"Recorder"
	 * <li>0x09 		"Data_Form"
	 * <li>0x0A 		"Auto_Activate"
	 * <li>0x0B		"Auto_Deactivate"
	 * <li>0x0C		"Sheet_Title"
	 * <li>0x0D 		"_FilterDatabase"
	 */
	public void setBuiltInType( int builtinType )
	{
		nBuiltin = (byte) builtinType;
		getData()[0] = nBuiltin;
	}

	/**
	 * create a new DCONBIN - built-in named range data source for a pivot table
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		DConBin db = new DConBin();
		db.setOpcode( DCONBIN );
		db.setData( new byte[]{ 0, 0, 0, 0, 0, 0 } );
		db.init();
		return db;
	}
}
