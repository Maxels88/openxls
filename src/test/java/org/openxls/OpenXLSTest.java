package org.openxls;

import org.junit.Before;
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
	private WorkSheetHandle ws;

	@Before
	public void setUp() throws Exception
	{
		WorkBookHandle wb = new WorkBookHandle();
		ws = wb.getWorkSheet( 0 );
	}

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

	@Test
	public void testSumProductOfSingleValueRef() throws Exception
	{
		ws.add( 5, "A1" );
		CellHandle cell = ws.add( "=SUMPRODUCT(A1)", "A2" );

		assertEquals( 5.0, cell.getVal() );
	}

	@Test
	public void testSumProductOfSingleIntValue() throws Exception
	{
		CellHandle cell = ws.add( "=SUMPRODUCT(5)", "A2" );

		assertEquals( 5.0, cell.getVal() );
	}

	@Test
	public void testSumProductOfSingleDoubleValue() throws Exception
	{
		CellHandle cell = ws.add( "=SUMPRODUCT(5.5)", "A2" );

		assertEquals( 5.5, cell.getVal() );
	}

	@Test
	public void testSumProductOfTwoColRange() throws Exception
	{
		CellHandle cell = ws.add( "=SUMPRODUCT(G7:H8)", "A2" );

		assertEquals( 0.0, cell.getVal() );
	}

	@Test
	public void testSumOfSingleNumber() throws Exception
	{
		CellHandle cell = ws.add( "=SUM(6)", "A5" );

		assertEquals( 6.0, cell.getVal() );
	}

	@Test
	public void testSumOfMultipleNumbers() throws Exception
	{
		CellHandle cell = ws.add( "=SUM(6,7,8)", "A5" );

		assertEquals( 21.0, cell.getVal() );
	}

	@Test
	public void testSumOfRange() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		ws.add( "7", "A3" );
		CellHandle cell = ws.add( "=SUM(A1:A3)", "A5" );

		assertEquals( 22.0, cell.getVal() );
	}

	@Test
	public void testAdditionOf2Cells() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		CellHandle cell = ws.add( "=A1+A2", "A5" );

		assertEquals( 15.0, cell.getVal() );
	}

	@Test
	public void testAdditionOf2CellsIsAssociative() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		CellHandle cell = ws.add( "=A2+A1", "A5" );

		assertEquals( 15.0, cell.getVal() );
	}

	@Test
	public void testAdditionOfCellAndConstant() throws Exception
	{
		ws.add( "5", "A1" );
		CellHandle cell = ws.add( "=A1+8", "A5" );

		assertEquals( 13.0, cell.getVal() );
	}

	@Test
	public void testAdditionOfCellAndConstantIsAssociative() throws Exception
	{
		ws.add( "5", "A1" );
		CellHandle cell = ws.add( "=8+A1", "A5" );

		assertEquals( 13.0, cell.getVal() );
	}

	@Test
	public void testSumProductOfLogicalAdditionWithUnaryMinusCoercion() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		ws.add( "7", "A3" );
		CellHandle cell = ws.add( "=SUMPRODUCT(--(A1:A3>0))", "A5" );

		assertEquals( 3.0, cell.getVal() );
	}

	@Test
	public void testSumProductOfLogicalAdditionOutsideOfArrayFormula() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		ws.add( "7", "A3" );
		CellHandle cell = ws.add( "=SUMPRODUCT((A1:A3>0))", "A5" );

		assertEquals( 0.0, cell.getVal() );
	}

	@Test
	public void testSumProductWithAdditiveCoercion() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		ws.add( "7", "A3" );
		CellHandle cell = ws.add( "=SUMPRODUCT(0+(A1:A3>0))", "A5" );

		assertEquals( 3.0, cell.getVal() );
	}

	@Test
	public void testSumProductWithMultiplierCoercion() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		ws.add( "7", "A3" );
		CellHandle cell = ws.add( "=SUMPRODUCT(2+(A1:A3>0))", "A5" );

		assertEquals( 9.0, cell.getVal() );
	}

	@Test
	public void testSumProductWithMultiplierCoercionWhenSomeValsAreFalse() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		ws.add( "7", "A3" );
		CellHandle cell = ws.add( "=SUMPRODUCT(2+(A1:A3>6))", "A5" );

		assertEquals( 8.0, cell.getVal() );
	}

	@Test
	public void testSumProductWithAdditiveCoercionWithArrayFirst() throws Exception
	{
		ws.add( "5", "A1" );
		ws.add( "10", "A2" );
		ws.add( "7", "A3" );
		CellHandle cell = ws.add( "=SUMPRODUCT((A1:A3>0)+0)", "A5" );

		assertEquals( 3.0, cell.getVal() );
	}

	@Test
	public void testAdditionOfNumberToTrueBoolean() throws Exception
	{
		ws.add( true, "A1" );
		ws.add( 5, "A2" );
		CellHandle cell = ws.add( "=A1+A2", "A3" );

		assertEquals( 6.0, cell.getVal() );
	}

	@Test
	public void testAdditionOfNumberToFalseBoolean() throws Exception
	{
		ws.add( false, "A1" );
		ws.add( 5, "A2" );
		CellHandle cell = ws.add( "=A1+A2", "A3" );

		assertEquals( 5.0, cell.getVal() );
	}

	@Test
	public void testSumProductCalcs() throws Exception
	{
		try( InputStream inp = OpenXLSTest.class.getResourceAsStream( "/SumProductTest.xls" ) )
		{
			WorkBookHandle wb = new WorkBookHandle( inp );
			WorkSheetHandle ws = wb.getWorkSheet( 0 );

			// Value is coerced, array formula
			assertEquals( 5.0, ws.getCell( "B8" ).getVal() );
			// Value is coerced, regular formula
			assertEquals( 5.0, ws.getCell( "B9" ).getVal() );
			// Value is not coerced, array formula
			assertEquals( 0.0, ws.getCell( "B10" ).getVal() );
			// Value is coerced, regular formula
			assertEquals( 5.0, ws.getCell( "B11" ).getVal() );
			assertEquals( -5.0, ws.getCell( "B12" ).getVal() );
			assertEquals( 150.0, ws.getCell( "B13" ).getVal() );
			assertEquals( 0.0, ws.getCell( "B14" ).getVal() );
		}
	}
}


