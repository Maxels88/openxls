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
package docs.samples.CellRanges;

import org.openxls.ExtenXLS.CellRange;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import org.openxls.formats.XLS.WorkSheetNotFoundException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * Demonstrates using CellRange to perform merging operations
 */
public class TestCellRanges
{
	private static final Logger log = LoggerFactory.getLogger( TestCellRanges.class );
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/CellRanges/";

	static public void main( String[] args )
	{
		TestCellRanges tcm = new TestCellRanges();
		tcm.mergeCells();
	}

	/**
	 * adds values to a worksheet, then
	 * creates a cell range to demonstrate various
	 * merging methods
	 */
	public void mergeCells()
	{
		WorkBookHandle book = new WorkBookHandle();
		WorkSheetHandle sheet = null;
		try
		{
			sheet = book.getWorkSheet( 0 );
		}
		catch( WorkSheetNotFoundException e )
		{
			System.err.println( "Error getting worksheet " + e );
		}

		// add values to the worksheet
		sheet.add( "Each", "A1" );
		sheet.add( "cell", "B1" );
		sheet.add( "has", "C1" );
		sheet.add( "been", "D1" );
		sheet.add( "preserved", "E1" );

		// add other values to worksheet
		sheet.add( "Only first cell is preserved", "A3" );
		sheet.add( "first", "B3" );
		sheet.add( "cell", "C3" );
		sheet.add( "is", "D3" );
		sheet.add( "preserved", "E3" );

		// create cell ranges for each, required in order to merge cells
		CellRange crx1 = null;
		CellRange crx2 = null;
		try
		{
			// defines the cell ranges
			crx1 = new CellRange( "Sheet1!A1:E1", book );
			crx2 = new CellRange( "Sheet1!A3:E3", book );

			// Merge the cells, keep the values in all cells...
			crx1.mergeCells( CellRange.RETAIN_MERGED_CELLS );

			// This time, remove the trailing Cells...
			crx2.mergeCells( CellRange.REMOVE_MERGED_CELLS );
		}
		catch( Exception e )
		{
			System.err.println( "Error setting cell ranges " + e );
		}
		testWrite( book, "TestCellRanges.xls" );

	}

	private void testWrite( WorkBookHandle b, String fout )
	{
		try
		{
			java.io.File f = new java.io.File( wd + fout );
			FileOutputStream fos = new FileOutputStream( f );
			BufferedOutputStream bbout = new BufferedOutputStream( fos );
			b.write( bbout );
			bbout.flush();
			fos.close();
		}
		catch( java.io.IOException e )
		{
			log.error( "IOException in Tester.", e );
		}
	}
}
