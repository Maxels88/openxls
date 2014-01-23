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
package docs.samples.SplitFreeze;

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import junit.framework.Assert;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;

/**
 * Tests the freeze pane and split-pane functionality of the WorkSheetHandle
 * ------------------------------------------------------------
 */
public class testSplitFreeze
{
	private static final Logger log = LoggerFactory.getLogger( testSplitFreeze.class );
	public static final String wd = System.getProperty( "user.dir" ) + "/docs/samples/SplitFreeze/";

	WorkBookHandle workbook;

	public static void main( String[] args )
	{
		testSplitFreeze test = new testSplitFreeze();
		try
		{
			log.info( "Begin testSplitFreeze." );
			test.setUp();
			test.tesSplitCols();
			test.tesSplitRows();
			test.testFreezePanes();
			log.info( "testSplitFreeze done." );
		}
		catch( Exception ex )
		{
			log.error( "testSplitFreeze failed", ex );
		}
	}

	protected void setUp() throws Exception
	{
		workbook = new WorkBookHandle();
		workbook.setDupeStringMode( WorkBookHandle.SHAREDUPES );
		workbook.setStringEncodingMode( WorkBookHandle.STRING_ENCODING_COMPRESSED );
	}

	/**
	 * splits cols into panes
	 * ------------------------------------------------------------
	 */
	protected void tesSplitCols()
	{
		try
		{
			WorkSheetHandle sheet = workbook.getWorkSheet( 0 );
			sheet.splitCol( 5, 5000 ); // split at col 5 (F), set divider at 1000 col units
		}
		catch( Exception e )
		{
			System.out.println( "Error setting Split panes: " + e.getMessage() );
		}
		writeFile( workbook, "testSplitPanesCol.xls" );
	}

	/**
	 * splits rows into panes
	 * ------------------------------------------------------------
	 */
	protected void tesSplitRows()
	{
		try
		{
			WorkSheetHandle sheet = workbook.getWorkSheet( 0 );
			sheet.splitRow( 10, 5000 ); // split at row 10, set divider at 5000 twips
		}
		catch( Exception e )
		{
			System.out.println( "Error setting Split panes: " + e.getMessage() );
		}
		writeFile( workbook, "testSplitPanesRow.xls" );
	}

	/**
	 * freezes panes
	 * ------------------------------------------------------------
	 */
	protected void testFreezePanes()
	{
		String fileName = "testFreezePanes.xls";
		try
		{
			WorkSheetHandle sheet = workbook.getWorkSheet( 0 );
			// try freezing
			sheet.freezeRow( 9 );
			sheet.freezeCol( 17 );
		}
		catch( Exception e )
		{
			System.out.println( "Error setting Freeze panes: " + e.getMessage() );
		}
		writeFile( workbook, fileName );
	}

	/**
	 * write the file to disk
	 * ------------------------------------------------------------
	 *
	 * @param workBookHandle
	 * @param excelFileName
	 */
	private static void writeFile( WorkBookHandle workBookHandle, String excelFileName )
	{
		try
		{
			File outputFile = new File( wd + excelFileName );
			FileOutputStream fileOutputStream = new FileOutputStream( outputFile );
			BufferedOutputStream bufferedOutputStream = new BufferedOutputStream( fileOutputStream );

			workBookHandle.write( bufferedOutputStream );

			bufferedOutputStream.flush();
			fileOutputStream.close();
		}
		catch( java.io.IOException e )
		{
			Assert.fail( "Exception thrown when trying to write the file: " + e );
		}
	}

}
