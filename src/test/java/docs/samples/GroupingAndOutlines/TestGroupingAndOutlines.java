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
package docs.samples.GroupingAndOutlines;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.toolkit.Logger;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * Demonstrates the operation of Row and Column grouping and outlining
 */
public class TestGroupingAndOutlines
{

	boolean CHANGEDATA = true;
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/GroupingAndOutlines/";

	public static void main( String[] args )
	{
		TestGroupingAndOutlines test = new TestGroupingAndOutlines();
		test.doit();
	}

	public void doit()
	{
		WorkBookHandle book = new WorkBookHandle();
		try
		{

			WorkSheetHandle sheet = book.getWorkSheet( "Sheet1" );

			// adding a new cells
			CellHandle ch1 = sheet.add( Integer.valueOf( 100 ), "A1" );
			CellHandle ch2 = sheet.add( Integer.valueOf( 200 ), "A2" );
			CellHandle ch3 = sheet.add( Integer.valueOf( 300 ), "A3" );
			CellHandle ch4 = sheet.add( Integer.valueOf( 100 ), "A4" );
			CellHandle ch5 = sheet.add( Integer.valueOf( 200 ), "A5" );
			CellHandle ch6 = sheet.add( Integer.valueOf( 300 ), "A6" );
			CellHandle ch7 = sheet.add( Integer.valueOf( 100 ), "B1" );
			CellHandle ch8 = sheet.add( Integer.valueOf( 200 ), "B2" );
			CellHandle ch9 = sheet.add( Integer.valueOf( 300 ), "B3" );
			CellHandle ch10 = sheet.add( Integer.valueOf( 100 ), "B4" );
			CellHandle ch11 = sheet.add( Integer.valueOf( 200 ), "B5" );
			CellHandle ch12 = sheet.add( Integer.valueOf( 300 ), "B6" );
			CellHandle ch13 = sheet.add( Integer.valueOf( 100 ), "C1" );
			CellHandle ch14 = sheet.add( Integer.valueOf( 200 ), "C2" );
			CellHandle ch15 = sheet.add( Integer.valueOf( 300 ), "C3" );
			CellHandle ch16 = sheet.add( Integer.valueOf( 100 ), "C4" );
			CellHandle ch17 = sheet.add( Integer.valueOf( 200 ), "C5" );
			CellHandle ch18 = sheet.add( Integer.valueOf( 300 ), "C6" );
			CellHandle ch19 = sheet.add( Integer.valueOf( 300 ), "D6" );

			// okay we want to 'hide' rows 4-7 and group them
			ch3.getRow().setOutlineLevel( 1 );
			ch4.getRow().setOutlineLevel( 1 );
			ch5.getRow().setOutlineLevel( 1 );
			ch6.getRow().setOutlineLevel( 1 );
			ch5.getRow().setCollapsed( true );

			ch6.getCol().setOutlineLevel( 1 );
			ch12.getCol().setOutlineLevel( 1 );

			ch18.getCol().setOutlineLevel( 1 );
			ch18.getCol().setCollapsed( true );

			ch19.getCol().setOutlineLevel( 1 );

		}
		catch( Exception ex )
		{
			// ex.printStackTrace();
			Logger.logErr( "ERROR: Group/Outline test Logger.logErred: " + ex.getMessage() );
		}
		try
		{
			WorkSheetHandle sheet = book.getWorkSheet( "Sheet1" );

			// retrieving the new cells
			CellHandle ch1 = sheet.getCell( "A1" );
			CellHandle ch2 = sheet.getCell( "A2" );
			CellHandle ch3 = sheet.getCell( "A3" );
			CellHandle ch4 = sheet.getCell( "A4" );
			CellHandle ch5 = sheet.getCell( "A5" );
			CellHandle ch6 = sheet.getCell( "A6" );
			CellHandle ch7 = sheet.getCell( "B1" );
			CellHandle ch8 = sheet.getCell( "B2" );
			CellHandle ch9 = sheet.getCell( "B3" );
			CellHandle ch10 = sheet.getCell( "B4" );
			CellHandle ch11 = sheet.getCell( "B5" );
			CellHandle ch12 = sheet.getCell( "B6" );
			CellHandle ch13 = sheet.getCell( "C1" );
			CellHandle ch14 = sheet.getCell( "C2" );
			CellHandle ch15 = sheet.getCell( "C3" );
			CellHandle ch16 = sheet.getCell( "C4" );
			CellHandle ch17 = sheet.getCell( "C5" );
			CellHandle ch18 = sheet.getCell( "C6" );
			CellHandle ch19 = sheet.getCell( "D6" );

			// see that we 'hid' rows 4-7 and grouped them
			Logger.logInfo( String.valueOf( ch3.getRow().getOutlineLevel() ) );
			Logger.logInfo( String.valueOf( ch4.getRow().getOutlineLevel() ) );
			Logger.logInfo( String.valueOf( ch5.getRow().getOutlineLevel() ) );
			Logger.logInfo( String.valueOf( ch6.getRow().getOutlineLevel() ) );
			Logger.logInfo( String.valueOf( ch5.getRow().isCollapsed() ) );

			Logger.logInfo( String.valueOf( ch6.getCol().getOutlineLevel() ) );
			Logger.logInfo( String.valueOf( ch12.getCol().getOutlineLevel() ) );
			Logger.logInfo( String.valueOf( ch18.getCol().getOutlineLevel() ) );
			Logger.logInfo( String.valueOf( ch19.getCol().getOutlineLevel() ) );
			Logger.logInfo( String.valueOf( ch18.getCol().isCollapsed() ) );

			testWrite( book );

		}
		catch( Exception e )
		{
			Logger.logErr( "ERROR: Could not verify proper Group/Outline in output file: " + e );
		}

	}

	public void testWrite( WorkBookHandle b )
	{
		try
		{
			java.io.File f = new java.io.File( wd + "TestGroupingAndOutlines_out.xls" );
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
