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
package docs.samples.AddingCells;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * This Class Demonstrates the basic functionality of of ExtenXLS.
 */
public class TestCellAdding
{

	public static void main( String[] args )
	{
		testadds t = new testadds();
		String s = "Test Successful.";

		t.testit( s );
	}
}

/**
 * Test the creation of a new Workbook with 3 worksheets.
 */
class testadds
{
	private static final Logger log = LoggerFactory.getLogger( testadds.class );
	private String wdir = System.getProperty( "user.dir" ) + "/docs/samples/AddingCells/";
	int ROWHEIGHT = 1000;
	int NUMADDS = 11000;

	public void testit( String argstr )
	{
		WorkBookHandle book = new WorkBookHandle();

//      IMPORTANT PERFORMANCE SETTINGS!!!
		book.setDupeStringMode( WorkBookHandle.SHAREDUPES );
		book.setStringEncodingMode( WorkBookHandle.STRING_ENCODING_COMPRESSED );

		WorkSheetHandle sheet = null;
		try
		{
			sheet = book.getWorkSheet( "Sheet1" );
		}
		catch( WorkSheetNotFoundException e )
		{
			System.out.println( "couldn't find worksheet" + e );
		}
		System.out.println( "Beginning Cell Adds" );
		String addr = "";

		// add a Double check that it was set
		sheet.add( new Double( 22250.321 ), "A1" );
		try
		{
			CellHandle cellA3 = sheet.getCell( "A1" );
			System.out.println( cellA3.getStringVal() );
		}
		catch( CellNotFoundException e )
		{
			;
		}
		long ltimr = System.currentTimeMillis();

		for( int i = 1; i < NUMADDS; i++ )
		{
			addr = "E" + String.valueOf( i );
			sheet.add( new Double( 1297.2753 * i ), addr );
			// try{sheet.getCell(addr).getRow().setHeight(2000);}catch(CellNotFoundException e){;}
		}
		System.out.print( "Adding " + NUMADDS );
		System.out.println( " Double values took: " + ((System.currentTimeMillis() - ltimr)) + " milliseconds." );

		String teststr = "ExtenXLS is used around the world by Global 1000 and Fortune 500 companies to provide dynamic Excel reporting in their Java web applications.";
		teststr += "Written entirely in Java, ExtenXLS frees you from platform dependencies, allowing you to give your users the information they need, in the world's most popular Spreadsheet format.";

		try
		{
			sheet = book.getWorkSheet( "Sheet2" );
		}
		catch( WorkSheetNotFoundException e )
		{
			System.out.println( "couldn't find worksheet" + e );
		}
		String t = "";

		// IMPORTANT PERFORMANCE SETTING!!!
		sheet.setFastCellAdds( true );

		ltimr = System.currentTimeMillis();
		for( int i = 1; i < NUMADDS; i++ )
		{
			addr = "B" + String.valueOf( i );
			t = teststr + String.valueOf( i );
			sheet.add( t, addr );
			//try{sheet.getCell(addr).getRow().setHeight(ROWHEIGHT);}catch(CellNotFoundException e){;}
		}

		System.out.print( "Adding " + NUMADDS );
		System.out.println( " Strings took: " + ((System.currentTimeMillis() - ltimr)) + " milliseconds." );

		ltimr = System.currentTimeMillis();
		this.testWrite( book );
		System.out.println( "Done." );
		System.out.print( "Writing " + book );
		System.out.println( " took: " + ((System.currentTimeMillis() - ltimr)) + " milliseconds." );

	}

	public void testWrite( WorkBookHandle book )
	{
		try
		{
			java.io.File f = new java.io.File( wdir + "testAddOutput.xls" );
			FileOutputStream fos = new FileOutputStream( f );
			BufferedOutputStream bbout = new BufferedOutputStream( fos );
			book.write( fos );
	        /* WRITE_XLS // we can write out XLS explicitly if we like
            	BufferedOutputStream bbout = new BufferedOutputStream(fos);
            	book.write(bbout);
            	bbout.flush();
                bbout.close();
                fos.flush();
                fos.close();
            
            WRITE_XLSX // we can write out XLSX explicitly if we like
            	int format= WorkBookHandle.FORMAT_XLSX;
            	if (foutpath.toLowerCase().endsWith(".xlsm"))
            		format= WorkBookHandle.FORMAT_XLSM;
            	BufferedOutputStream bbout = new BufferedOutputStream(fos);
            	book.write(bbout, format);
            	bbout.flush();
                bbout.close();
                fos.flush();
                fos.close();
			*/
		}
		catch( java.io.IOException e )
		{
			log.info( "IOException in Tester.  " + e );
		}
	}

}