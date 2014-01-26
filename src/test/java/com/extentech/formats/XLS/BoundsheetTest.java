package com.extentech.formats.XLS;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.ColHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import org.junit.Before;
import org.junit.Test;

import java.io.InputStream;
import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * User: npratt
 * Date: 1/23/14
 * Time: 13:17
 */
public class BoundsheetTest
{
	private Boundsheet bs;

	@Before
	public void setUp() throws Exception
	{
		bs = new Boundsheet();
	}

	@Test
	public void testBoundsheet() throws Exception
	{
		bs.addCell( createBlank( 0, 0 ) );
		bs.addCell( createBlank( 1, 0 ) );
		bs.addCell( createBlank( 2, 0 ) );
		bs.addCell( createMulblank( 0, 1, 2 ) );
		bs.addCell( createMulblank( 1, 1, 2 ) );
		bs.addCell( createMulblank( 2, 1, 2 ) );

		List<? extends BiffRec> cellsByCol = bs.getCellsByCol( 0 );

		assertEquals( 3, cellsByCol.size() );
	}

	private static CellRec createMulblank( int row, int colFirst, int colLast )
	{
		Mulblank mulblank = new Mulblank(row, colFirst, colLast);
		return mulblank;
	}

	@Test
	public void testGetCellsForCoWithJustTextLabelsl() throws Exception
	{
		try( InputStream inp = getClass().getResourceAsStream( "/Mulblank1.xls" ) )
		{
			WorkBookHandle wbh = new WorkBookHandle( inp );
			WorkSheetHandle worksheetHandle = wbh.getWorkSheet( 0 );
			ColHandle colHandle = worksheetHandle.getCol( 0 );

			CellHandle[] cells = colHandle.getCells();
			assertEquals( 3, cells.length );
		}
	}

	@Test
	public void testGetCellsForCoWithJustBlanks() throws Exception
	{
		try( InputStream inp = getClass().getResourceAsStream( "/Mulblank1.xls" ) )
		{
			WorkBookHandle wbh = new WorkBookHandle( inp );
			WorkSheetHandle worksheetHandle = wbh.getWorkSheet( 0 );
			ColHandle colHandle = worksheetHandle.getCol( 1 );

			CellHandle[] cells = colHandle.getCells();
			assertEquals( 0, cells.length );
		}
	}

	@Test
	public void testGetCellsForCoWithJustBlanksButWithOpenTopTwoRows() throws Exception
	{
		try( InputStream inp = getClass().getResourceAsStream( "/Mulblank2.xls" ) )
		{
			WorkBookHandle wbh = new WorkBookHandle( inp );
			WorkSheetHandle worksheetHandle = wbh.getWorkSheet( 0 );
			ColHandle colHandle = worksheetHandle.getCol( 0 );

			CellHandle[] cells = colHandle.getCells();
			assertEquals( 6, cells.length );
		}
	}

	@Test
	public void testGetCellsForCoWithJustTextLabelsAfterCallToGetCell() throws Exception
	{
		try( InputStream inp = getClass().getResourceAsStream( "/Mulblank2.xls" ) )
		{
			WorkBookHandle wbh = new WorkBookHandle( inp );
			WorkSheetHandle worksheetHandle = wbh.getWorkSheet( 0 );
			CellHandle cellHandle = worksheetHandle.getCell( "B4" );
			ColHandle colHandle = worksheetHandle.getCol( 0 );

			CellHandle[] cells = colHandle.getCells();
			assertEquals( 6, cells.length );
		}
	}

	static Blank createBlank( int row, int col )
	{
		Blank blank = new Blank( );
		blank.setRow( row );
		blank.setCol( col );
		return blank;
	}

	static Labelsst createLabel( int row, int col, String val )
	{
		Labelsst cell = new Labelsst(row, col);
		cell.setStringVal( val );
		return cell;
	}
}
