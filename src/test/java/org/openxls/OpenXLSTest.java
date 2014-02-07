package org.openxls;

import org.junit.Test;
import org.openxls.ExtenXLS.CellHandle;
import org.openxls.ExtenXLS.ColHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;

import java.io.InputStream;
import java.util.Date;

import static org.junit.Assert.assertEquals;

/**
 * User: npratt
 * Date: 1/10/14
 * Time: 10:43
 */
public class OpenXLSTest
{
	@Test
	public void testCreateWorksheetAndSetValue() throws Exception
	{
		try( InputStream inp = OpenXLSTest.class.getResourceAsStream( "/Test.xls" ) )
		{
			WorkBookHandle wbh = new WorkBookHandle( inp );
			WorkSheetHandle worksheetHandle = wbh.getWorkSheet( 0 );
			CellHandle cellHandle = worksheetHandle.getCell( "B4" );
			ColHandle colHandle = worksheetHandle.getCol( 0 );

			CellHandle[] cells = colHandle.getCells();
			assertEquals( 6, cells.length );

			cellHandle.setVal( "Test" );

			// Why does this trigger an ArrayIndexOutOfBoundsException ?
			// 20140207 - It doesn't any more since it was fixed :-)
			colHandle.getCells();
		}
	}

	@Test
	public void testTimeNowIsUpdatedOnWorkbookOpen() throws Exception
	{
		try( InputStream inp = OpenXLSTest.class.getResourceAsStream( "/TimeNow.xls" ) )
		{
			WorkBookHandle wb = new WorkBookHandle( inp );
			WorkSheetHandle ws = wb.getWorkSheet( 0 );

			Date timeNow = new Date();
			int mins = timeNow.getMinutes() % 60;
			int hours = timeNow.getHours() % 12;
			// Dont check the days - not really worth it
//			int dayOfMonth = timeNow.getDay();
			int month = (timeNow.getMonth() + 1) % 12;
			int year = timeNow.getYear() + 1900;

			int wsMins = ws.getCell( "C2" ).getIntVal() % 60;
			int wsHours = ws.getCell( "D2" ).getIntVal() % 12;
//			int wsDay = ws.getCell( "E2" ).getIntVal();
			int wsMonth = ws.getCell( "F2" ).getIntVal() % 12;
			int wsYear = ws.getCell( "G2" ).getIntVal();

			// We need to account for the clock possibly ticking over during this test execution, so values from the sheet could be
			// '1' tick behind the real value
			boolean minsAreSame = mins == wsMins || mins == (wsMins + 1);
			assertEquals( "wsMins: " + wsMins + ", mins: " + mins, true, minsAreSame );

			boolean hoursAreSame = hours == wsHours || hours == (wsHours + 1);
			assertEquals( "wsHours: " + wsHours + ", hours: " + hours, true, hoursAreSame );

			boolean monthsAreSame = month == wsMonth || month == (wsMonth + 1);
			assertEquals( "wsMonth: " + wsMonth + ", month: " + month, true, monthsAreSame );

			boolean yearsAreSame = year == wsYear || year == (wsYear + 1);
			assertEquals( "wsYear: " + wsYear + ", year: " + year, true, yearsAreSame );
		}
	}
}


