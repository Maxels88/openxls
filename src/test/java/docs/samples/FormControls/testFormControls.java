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
package docs.samples.FormControls;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.NameHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.toolkit.Logger;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * Test manipulation of Excel Form objects using ExtenXLS and NamedRanges
 * ------------------------------------------------------------
 */
public class testFormControls
{

	// change to the folder containing the image files and spreadsheet templates
	String workingdir = System.getProperty( "user.dir" ) + "/docs/samples/FormControls/";
	int DEBUGLEVEL = 0;
	WorkSheetHandle sheet = null;
	String sheetname = "Sheet1";
	String finpath = "";

	public static void main( String[] args )
	{
		testFormControls ti = new testFormControls( "Test Forms", 0 );
		ti.test();

	}

	public testFormControls( String name, int lev )
	{
		DEBUGLEVEL = lev;
	}

	/**
	 * read in an XLS file, output BiffRec vals, write it back to a new file
	 */
	public void test()
	{
		// working files
		String[] testfiles = {
				"testFormControls.xls", "Sheet1",
		};

		for( int i = 0; i < testfiles.length; i++ )
		{
			finpath = testfiles[i];
			System.out.println( "====================== TEST: " + String.valueOf( i ) + " ======================" );
			sheetname = testfiles[++i];
			System.out.println( "=====================> " + finpath + ":" + sheetname );

			doit( finpath, sheetname );

			System.out.println( "DONE" );
		}
	}

	void doit( String finpath, String sheetname )
	{
		System.out.println( "Begin parsing: " + workingdir + finpath );
		WorkBookHandle tbo = new WorkBookHandle( workingdir + finpath );

		try
		{
			sheet = tbo.getWorkSheet( sheetname );

			// get ahold of the 3 named ranges in the template which are bound to form controls
			// (a checkbox and a list)

			NameHandle checkbox = tbo.getNamedRange( "checkbox" );
			NameHandle radio = tbo.getNamedRange( "radio" );
			NameHandle list = tbo.getNamedRange( "list" );
			NameHandle selected = tbo.getNamedRange( "selected" );

			// confirm we got
			System.out.println( "Got 4 named ranges: " + checkbox + ", " + radio + ", " + list + ", " + selected );

			// change checkbox to selected
			CellHandle[] checkedcell = checkbox.getCells();
			checkedcell[0].setVal( true );

			// change radiobutton to selected
			CellHandle[] radiocell = radio.getCells();
			radiocell[0].setVal( 2 );

			// change size and content of list named range
			CellHandle newcell = tbo.getWorkSheet( "Sheet2" ).add( "twenty nine", "I14" );
			list.addCell( newcell );

			CellHandle[] listcells = list.getCells();

//			 set currently selected list item
			CellHandle[] selectedcells = selected.getCells();

			selectedcells[0].setVal( 5 );

		}
		catch( Exception e )
		{
			System.err.println( "testForms failed: " + e.toString() );
		}
		testWrite( tbo, workingdir + "testFormsOut.xls" );
		// write out the results
		WorkBookHandle newbook = new WorkBookHandle( workingdir + "testFormsOut.xls", 0 );
		System.out.println( "Successfully read: " + newbook );
	}

	public void testWrite( WorkBookHandle b, String fout )
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
			Logger.logInfo( "IOException in Tester.  " + e );
		}
	}

}