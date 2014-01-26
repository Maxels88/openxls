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

import org.openxls.ExtenXLS.CellRange;
import org.openxls.ExtenXLS.ExcelTools;
import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.UnsupportedEncodingException;

/**
 * DConRef	0x51
 * <p/>
 * The DConRef record specifies a range in this workbook or in an external
 * workbook that is a data source for a PivotTable or a data source for the data
 * consolidation settings of the associated sheet. If the range specified is in
 * an external workbook this record also specifies the path to the external
 * workbook.
 * <p/>
 * ref (6 bytes): A RefU structure that specifies the range. If this record is part of an SXTBL production as specified in the
 * Globals Substream ABNF and this field has a rwFirst equal to 0 and a rwLast equal to 16383, this reference specifies all rows
 * within the columns specified by colFirst and colLast.
 * <p/>
 * cchFile (2 bytes): An unsigned integer that specifies the count of characters in stFile. MUST be greater than or equal to 0x0002.
 * <p/>
 * stFile (variable): A DConFile structure that specifies the workbook and sheet that contains the range specified in the ref field.
 * <p/>
 * unused (variable): An array of bytes that is unused and MUST be ignored. MUST exist if and only if stFile specifies a self reference
 * (the value of stFile.stFile.rgb[0] is 2). If the value stFile.stFile.fHighByte is 0 the size of this array is 1.
 * If the value of stFile.stFile.fHighByte is 1 the size of this array is 2.
 * <p/>
 * <p/>
 * The RefU structure specifies a range of cells on the sheet.
 * rwFirst (2 bytes): A RwU structure that specifies the first row in the range. The value MUST be less than or equal to rwLast.
 * rwLast (2 bytes): A RwU structure that specifies the last row in the range.
 * colFirst (1 byte): A ColByteU structure that specifies the first column in the range. The value MUST be less than or equal to colLast.
 * colLast (1 byte): A ColByteU structure that specifies the last column in the range.
 * <p/>
 * the DConFile structure specifies the workbook file or workbook file and sheet that contain a data source range.
 * This structure is used by the DConBin, DConRef and DConName records.
 * stFile (variable): An XLUnicodeStringNoCch that specifies the workbook file or workbook file and sheet that contain the range specified in the DConBin, DConRef or DConName record.
 * MUST be a string that conforms to the following ABNF grammar:
 * dcon-file = external-virt-path / self-reference
 * external-virt-path = volume / unc-volume / rel-volume  / transfer-protocol / startup / alt-startup / library /  simple-file-path-dcon
 * simple-file-path-dcon = %x0001 file-path
 * self-reference = %x0002 sheet-name
 * See VirtualPath for the definition of the volume, unc-volume, rel-volume, transfer-protocol, startup, alt-startup, library, file-path and sheet-name rules used in the ABNF grammar.  Note that the volume, unc-volume, rel-volume, transfer-protocol, startup, alt-startup, library, and file-path rules specify that an optional sheet name can be included.
 * If this structure is contained in a DConName or DConBin record and the defined name has a workbook scope, then this string MUST satisfy the external-virt-path rule and MUST NOT specify a sheet name.  Otherwise a sheet name MUST be specified.
 * <p/>
 * <p/>
 * <p/>
 * "http://www.extentech.com">Extentech Inc.</a>
 */
public class DConRef extends XLSRecord implements XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( DConRef.class );
	private static final long serialVersionUID = 2639291289806138985L;
	private short rwFirst;
	private short rwLast;
	private short colFirst;
	private short colLast;
	private short cchFile;
	private String fileName = null;
	private byte refType = 0;

	/**
	 * init method
	 */
	@Override
	public void init()
	{
		super.init();
		rwFirst = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		rwLast = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		colFirst = (short) getByteAt( 4 );
		colLast = (short) getByteAt( 5 );
		cchFile = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		if( cchFile > 0 )
		{
			//A - fHighByte (1 bit): A bit that specifies whether the characters in rgb are double-byte characters.
			// 0x0  All the characters in the string have a high byte of 0x00 and only the low bytes are in rgb.
			// 0x1  All the characters in the string are saved as double-byte characters in rgb.
			// reserved (7 bits): MUST be zero, and MUST be ignored.

			byte encoding = getByteAt( 8 );
			refType = getByteAt( 9 );    // 1= simple-file-path-dcon 2= self-reference

			if( refType != 2 )    // TODO: handle external refs ...
			{
				log.warn( "PivotTable: External Data Sources are not supported" );
			}
			byte[] tmp = getBytesAt( 10, (cchFile - 1) * (encoding + 1) );
			try
			{
				if( encoding == 0 )
				{
					fileName = new String( tmp, DEFAULTENCODING );
				}
				else
				{
					fileName = new String( tmp, UNICODEENCODING );
				}
			}
			catch( UnsupportedEncodingException e )
			{
				log.warn( "encoding PivotTable name in DCONREF: " + e, e );
			}
		}
			log.debug( "DCONREF: rwFirst:" + rwFirst + " rwLast:" + rwLast + " colFirst:" + colFirst + " colLast:" + colLast + " cchFile:" + cchFile + " fileName:" + fileName );
	}

	/**
	 * returns the source range for a pivot table in r0c0r1c form
	 *
	 * @return int[]
	 */
	public int[] getRange()
	{
		return new int[]{ rwFirst, colFirst, rwLast, colLast };
	}

	/**
	 * if this is a self-referentail data source i.e. in same workbook, return souce sheet name
	 *
	 * @return
	 */
	public String getSourceSheet()
	{
		if( refType != 2 )
		{
			log.warn( "External Data Sources are not supported" );
		}
		return fileName;
	}

	/**
	 * sets the source range and sheet for the pivot table
	 *
	 * @param rc
	 */
	public void setRange( int[] rc, String sheetName )
	{
		rwFirst = (short) rc[0];
		colFirst = (short) rc[1];
		rwLast = (short) rc[2];
		colLast = (short) rc[3];
		// update record
		byte[] data = getData();
		byte[] b = ByteTools.shortToLEBytes( rwFirst );
		data[0] = b[0];
		data[1] = b[1];
		b = ByteTools.shortToLEBytes( rwLast );
		data[2] = b[0];
		data[3] = b[1];
		data[4] = (byte) colFirst;
		data[5] = (byte) colLast;
		setSourceSheet( sheetName );
	}

	/**
	 * sets the source sheet for the pivot table
	 *
	 * @param sheetName
	 */
	public void setSourceSheet( String sheetName )
	{
		cchFile = (short) ((short) sheetName.length() + 1);
		fileName = sheetName;
		byte[] data = new byte[10];
		System.arraycopy( getData(), 0, data, 0, 6 );
		byte[] b = ByteTools.shortToLEBytes( cchFile );
		data[6] = b[0];    // cch
		data[7] = b[1];
		data[8] = 0;        // encoding
		data[9] = 0x2;    // self-reference flag
		try
		{
			data = ByteTools.append( sheetName.getBytes( DEFAULTENCODING ), data );
			data = ByteTools.append( new byte[]{ 0 }, data );
		}
		catch( UnsupportedEncodingException e )
		{
		}
		setData( data );
	}

	/**
	 * create a new default DCONREF source data range record
	 *
	 * @return
	 */
	public static XLSRecord getPrototype()
	{
		DConRef dr = new DConRef();
		dr.setOpcode( DCONREF );
		dr.setData( new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0 } );
		dr.init();
		return dr;
	}

	/**
	 * returns the cell range for this pivot cache
	 *
	 * @return
	 */
	public CellRange getCellRange()
	{
		String range = fileName + "!" + ExcelTools.formatLocation( new int[]{ rwFirst, colFirst, rwLast, colLast } );
		try
		{
			return new CellRange( range, null );    //this.getWorkBook());
		}
		catch( CellNotFoundException e )
		{
		}
		return null;
	}

	/**
	 * sets the cell range for the pivot cache
	 *
	 * @param cr
	 */
	public void setCellRange( CellRange cr )
	{
		try
		{
			int[] rc = cr.getRangeCoords();
			setRange( rc, cr.getSheet().getSheetName() );
		}
		catch( CellNotFoundException e )
		{
		}
	}

	/**
	 * sets the cell range for the pivot cache
	 *
	 * @param cr
	 */
	public void setCellRange( String cr )
	{
		String sheetname;
		if( cr.indexOf( "!" ) != -1 )
		{
			sheetname = cr.split( "!" )[0];
		}
		else
		{
			sheetname = fileName;
		}
		int[] rc = ExcelTools.getRangeCoords( cr );
		setRange( rc, sheetname );
	}
}
