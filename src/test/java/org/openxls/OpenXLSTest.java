package org.openxls;

import org.openxls.ExtenXLS.CellHandle;
import org.openxls.ExtenXLS.ColHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import org.junit.Test;

import java.io.InputStream;

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
			colHandle.getCells();
		}
	}
}


