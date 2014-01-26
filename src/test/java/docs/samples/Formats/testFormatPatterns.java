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

import org.openxls.ExtenXLS.CellHandle;
import org.openxls.ExtenXLS.FormatHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;

/**
 * Demonstration of using Numeric Format Patterns
 * ------------------------------------------------------------
 *
 * @see FormatHandle
 */
public class testFormatPatterns
{

	/**
	 * test setting of numeric format patterns (aka: $, %, etc.)
	 */
	public testFormatPatterns()
	{
		super();
	}

	public static void main( String[] args )
	{
		String wd = System.getProperty( "user.dir" ) + "/docs/samples/Formats/";
		try
		{

			WorkBookHandle bookx = new WorkBookHandle( wd + "testFormatPatterns.xls" );

			CellHandle cx[] = bookx.getCells();
			for( int t = 0; t < cx.length; t++ )
			{
				CellHandle cell = cx[t];
				System.out.println( cell.getCellAddress() + ":" + cell.getFormatPattern() );
			}

			WorkBookHandle book = new WorkBookHandle();
			WorkSheetHandle sheet = book.getWorkSheet( "Sheet1" );
			for( int i = 0; i < 12; i++ )
			{
				sheet.add( new Double( 1.23456 ), "A" + (i + 1) );
			}
			sheet.getCell( "A1" ).getFormatHandle().setFormatPattern( "#,##0" );
			sheet.getCell( "A2" ).getFormatHandle().setFormatPattern( "#,##0.0" );
			sheet.getCell( "A3" ).getFormatHandle().setFormatPattern( "#,##0.00" );
			sheet.getCell( "A4" ).getFormatHandle().setFormatPattern( "#,##0.000" );
			sheet.getCell( "A5" ).getFormatHandle().setFormatPattern( "#,##0.0000" );
			sheet.getCell( "A6" ).getFormatHandle().setFormatPattern( "#,##0.00000" );

			sheet.getCell( "A7" ).getFormatHandle().setFormatPattern( "0%" );
			sheet.getCell( "A8" ).getFormatHandle().setFormatPattern( "0.0%" );
			sheet.getCell( "A9" ).getFormatHandle().setFormatPattern( "0.00%" );
			sheet.getCell( "A10" ).getFormatHandle().setFormatPattern( "0.000%" );
			sheet.getCell( "A11" ).getFormatHandle().setFormatPattern( "0.0000%" );
			sheet.getCell( "A12" ).getFormatHandle().setFormatPattern( "0.00000%" );

			File file = new File( wd + "testFormatPatterns_out.xls" );

			BufferedOutputStream stream = new BufferedOutputStream( new FileOutputStream( file ) );
			book.write( stream );
			stream.flush();
			stream.close();

		}
		catch( Exception e )
		{
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * @param precision
	 * @return
	 */
	private static String getFormatString( int precision )
	{
		StringBuffer format = new StringBuffer( "#,##0" );
		if( precision > 0 )
		{
			format.append( "." );
		}
		for( int i = 0; i < precision; i++ )
		{
			format.append( "0" );
		}
		return format.toString();
	}

}
