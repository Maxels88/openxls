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
package docs.samples.Formats;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.XLS.FormatConstants;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * Demonstrates creating borders on Cell Ranges
 */
public class testBorders
{
	private static final Logger log = LoggerFactory.getLogger( testBorders.class );
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/Formats/";

	/**
	 * Demonstrates creating borders on Cell Ranges
	 * Jan 19, 2010
	 *
	 * @param args
	 */
	public static void main( String[] args )
	{
		testBorders tb = new testBorders();
		tb.testit();
	}

	/**
	 * tests various border types and cell sides
	 * ------------------------------------------------------------
	 */
	public void testit()
	{
		log.info( "====================== TESTING borders on cells ======================" );
		WorkBookHandle tbo = new WorkBookHandle();
		try
		{
			WorkSheetHandle sheet = tbo.getWorkSheet( 0 );

			int[] coords = { 12, 1, 12, 5 };
			CellRange range = new CellRange( sheet, coords, true );

			CellRange range2 = new CellRange( "Sheet1!B2:Sheet1!C10", tbo, true );
			range2.setBorder( 2, FormatConstants.BORDER_THIN, Color.blue );

			// set top and bottom
			FormatHandle myfmthandle = new FormatHandle( tbo );
			myfmthandle.addCellRange( range );
			myfmthandle.setTopBorderLineStyle( FormatHandle.BORDER_DOUBLE );
			myfmthandle.setBottomBorderLineStyle( FormatHandle.BORDER_THICK );

			// set sides
			int[] coords2 = { 5, 4, 5, 8 };
			CellRange range3 = new CellRange( sheet, coords2, true );
			FormatHandle myfmthandle2 = new FormatHandle( tbo );
			myfmthandle2.addCellRange( range3 );
			myfmthandle2.setBorderLeftColor( Color.red );
			myfmthandle2.setLeftBorderLineStyle( FormatHandle.BORDER_DASH_DOT_DOT );

			myfmthandle2.setBorderRightColor( Color.blue );
			myfmthandle2.setRightBorderLineStyle( FormatHandle.BORDER_DOUBLE );

			// ok, test not clobbering
			CellRange range4 = new CellRange( sheet, coords2, true );

			CellHandle cell0 = range4.getCells()[0];

			FormatHandle clobberfmt = cell0.getFormatHandle();
			clobberfmt.setCellBackgroundColor( Color.lightGray );
			clobberfmt.setUnderlined( true );

			cell0.setVal( "hello world!" );

		}
		catch( Exception ex )
		{
			log.error( "testCellBorder failed:" + ex.toString() );
		}
		testWrite( tbo );

		log.info( "====================== DONE TESTING borders on cells ======================" );

	}

	public void testWrite( WorkBookHandle book )
	{
		try
		{
			java.io.File f = new java.io.File( wd + "testBorders_out.xls" );
			FileOutputStream fos = new FileOutputStream( f );
			BufferedOutputStream bbout = new BufferedOutputStream( fos );
			book.write( bbout );
			bbout.flush();
			fos.close();
		}
		catch( java.io.IOException e )
		{
			log.info( "IOException in Tester.  " + e );
		}
	}
}