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
/** ------------------------------------------------------------
 * CopySheets.java
 *
 *
 * ------------------------------------------------------------
 */
package docs.samples.CopySheets;

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * ------------------------------------------------------------
 * This test demonstrates how to read in values from a source workbook and split into
 * multiple output worksheets.
 */
public class CopySheets
{
	private static final Logger log = LoggerFactory.getLogger( CopySheets.class );
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/CopySheets/";

	/**
	 * ------------------------------------------------------------
	 *
	 * @param args
	 */
	public static void main( String[] args )
	{
		log.info( "ExtenXLS Version " + WorkBookHandle.getVersion() );
		log.info( "---- Testing Copy Sheet Functionality ---- " );
		CopySheets cx = new CopySheets();
		try
		{
			cx.copy();
		}
		catch( Exception ex )
		{
			log.error( "ERROR: CopySheets Failed: " + ex.toString() );
		}
		log.info( "---- Done with Copy Sheet ---- " );

	}

	/**
	 * Copy worksheets from one book to another
	 * ------------------------------------------------------------
	 *
	 * @throws Exception
	 */
	public void copy() throws Exception
	{

		// some big input file with a lot of sheets to copy
		WorkBookHandle biginput = new WorkBookHandle( wd + "source_sheets.xls" );

		// a default, empty workbook with no worksheets
		WorkBookHandle newbook = new WorkBookHandle().getNoSheetWorkBook();

		// loop through the input sheets and copy contents
		WorkSheetHandle[] shts = biginput.getWorkSheets();
		for( int i = 0; i < shts.length; i++ )
		{
			// get the sheet to copy
			WorkSheetHandle sheet = shts[i];

			// add the sheet to the destination workbook
			newbook.addWorkSheet( sheet );
		}
		testWrite( newbook, wd + "source_sheets_out.xls" );
	}

	public static void testWrite( WorkBookHandle b, String fout )
	{
		try
		{
			java.io.File f = new java.io.File( fout );
			FileOutputStream fos = new FileOutputStream( f );
			BufferedOutputStream bbout = new BufferedOutputStream( fos );
			b.write( bbout );
			bbout.flush();
			fos.close();
		}
		catch( java.io.IOException e )
		{
			log.info( "IOException in Tester.  " + e );
		}
	}

}
