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
package docs.samples.Validations;

import org.openxls.ExtenXLS.CellHandle;
import org.openxls.ExtenXLS.ValidationHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * This Class Demonstrates the basic functionality of of ExtenXLS.
 */
public class TestValidations
{
	private static final Logger log = LoggerFactory.getLogger( TestValidations.class );
	public static final String wd = System.getProperty( "user.dir" ) + "/docs/samples/Validations/";

	public static void main( String[] args )
	{
		TestValidations t = new TestValidations();
		String s = "testValidation.xls";

		t.testExistingValidation( s, "Sheet1" );
		t.testCreateValidation();
	}

	/**
	 * Tests creating a validaiton (and a dvrec)
	 * ------------------------------------------------------------
	 */
	public void testCreateValidation()
	{
		WorkBookHandle book = new WorkBookHandle();
		try
		{
			WorkSheetHandle sheet = book.getWorkSheet( 0 );
			sheet.add( Integer.valueOf( 33 ), "A1" );
			ValidationHandle vh = sheet.createValidationHandle( "A1:A2",
			                                                    ValidationHandle.VALUE_INTEGER,
			                                                    ValidationHandle.CONDITION_BETWEEN,
			                                                    "errorText",
			                                                    "errorBoxTitle",
			                                                    "promptBoxText",
			                                                    "promptBoxTitle",
			                                                    "12",
			                                                    "44" );
			vh.addRange( "C1:C2" );
			testWrite( book, "new_validation_output" );
		}
		catch( Exception e )
		{
			log.error( "",e );
		}
	}

	/**
	 * Tests opening a file with validation set up, and attempts to get a handle to the existing
	 * validation.  Updates the validation, saves the file, and reopens to verify the validation exists
	 * through a save/open cycle
	 *
	 * @param file
	 * @param sheetname
	 */
	void testExistingValidation( String file, String sheetname )
	{
		WorkBookHandle book = new WorkBookHandle( wd + file );
		try
		{
			WorkSheetHandle sheet = book.getWorkSheet( sheetname );
			// this is a range validation:
			CellHandle rd = sheet.getCell( "B9" );
			ValidationHandle rvh = rd.getValidationHandle();

			rvh.setErrorBoxText( "ExtenXLS Says: enter a number bigger than 1000!" );

			rvh.setPromptBoxTitle( "ExtenXLS Controlling Your World" );
			rvh.setPromptBoxText( "ExtenXLS Says: Please enter a value bigger than a breadbox" );

			CellHandle ad = sheet.add( "Testing Validation Handle", "A1" );
			// first let's try to get a hold of a one cell validation.  D15
			CellHandle c = sheet.getCell( "D15" );
			ValidationHandle vh = c.getValidationHandle();

			vh.setFirstCondition( "D15>1000" );
			vh.setErrorBoxText( "ExtenXLS Says: enter a number bigger than 1000!" );

			vh.setPromptBoxTitle( "ExtenXLS Controlling Your World" );
			vh.setPromptBoxText( "ExtenXLS Says: Please enter a value bigger than a breadbox" );

			// entering incorrect value in D15 will show the error box text above
			testWrite( book, "testValidationOutput" );

		}
		catch( Exception e )
		{
			log.error( "Failure in testValidationHandle.testExistingValidation: " + e.toString() );
		}

	}

	public static void testWrite( WorkBookHandle b, String file )
	{
		try
		{
			java.io.File f = new java.io.File( wd + file + ".xls" );
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