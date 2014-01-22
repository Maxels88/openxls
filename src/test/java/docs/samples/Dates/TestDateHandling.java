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
package docs.samples.Dates;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.DateConverter;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.toolkit.Logger;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.Calendar;

/**
 * This Class Demonstrates the basic functionality of of ExtenXLS Date Handling
 */
public class TestDateHandling
{

	public static void main( String[] args )
	{
		testDates t = new testDates();
		t.testit();
	}
}

/**
 * Test the creation of a new Workbook with 3 worksheets.
 */
class testDates
{
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/Dates/";
	String file = wd + "testdates.xls", sheetname = "Sheet1";
	int ROWHEIGHT = 100;
	int NUMADDS = 100;
	WorkBookHandle book = null;
	WorkSheetHandle sheet = null;

	void testit()
	{
		SimpleDateFormat sdf = new SimpleDateFormat( "yyyy-MM-dd hh:mm:ss" );
		book = new WorkBookHandle( file );
		try
		{
			sheet = book.getWorkSheet( sheetname );
		}
		catch( WorkSheetNotFoundException e )
		{
			Logger.logErr( e );
		}

		// get a date-time
		CellHandle a1 = null, a2 = null, a3 = null;
		int GMT_FORMAT = -1;
		try
		{
			a1 = sheet.getCell( "a1" );
			GMT_FORMAT = a1.getFormatId();
			a2 = sheet.getCell( "a2" );
			a3 = sheet.getCell( "a3" );
		}
		catch( CellNotFoundException e )
		{
			Logger.logErr( e );
		}

		Logger.logInfo( sdf.format( DateConverter.getDateFromCell( a1 ) ) );
		Logger.logInfo( sdf.format( DateConverter.getDateFromCell( a2 ) ) );
		Logger.logInfo( sdf.format( DateConverter.getDateFromCell( a3 ) ) );

		// add new date to cells
		CellHandle a6 = sheet.add( new Date( System.currentTimeMillis() ), "a6" );
		Logger.logInfo( a6.getFormattedStringVal() );

		CellHandle a7 = sheet.add( new Date( System.currentTimeMillis() + 5555555 ), "a7", "m/d/yy h:mm" );
		Logger.logInfo( a7.getFormattedStringVal() );

		// Excel Divides the day into 100 units, so .25 is 9am, .5 is 12 noon, .75 is 9pm, and 00 is midnight
		try
		{
			sheet.add( new Float( 2000.75 ), "a5" ); // add a new cell, set date format
			CellHandle a5 = sheet.getCell( "a5" );
			a5.setFormatId( GMT_FORMAT );
			Logger.logInfo( sdf.format( DateConverter.getDateFromCell( a5 ) ) );
		}
		catch( Exception e )
		{
			Logger.logErr( "setting date failed.", e );
		}

		try
		{
			// set today's date in cell A10
			Date dt = new Date( System.currentTimeMillis() );
			Calendar gc = new java.util.GregorianCalendar();
			gc.setTime( dt );
			double dd = DateConverter.getXLSDateVal( gc );
			CellHandle a10 = sheet.add( new Double( dd ), "A10" ); // add a new cell, set date format
			a10.setFormatId( GMT_FORMAT ); // set date format on cell
			Logger.logInfo( sdf.format( DateConverter.getDateFromCell( a10 ) ) );
		}
		catch( Exception e )
		{
			Logger.logErr( "setting date failed.", e );
		}
		testWrite( book );
	}

	public void testWrite( WorkBookHandle b )
	{
		try
		{
			java.io.File f = new java.io.File( wd + "testDates_out.xls" );
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