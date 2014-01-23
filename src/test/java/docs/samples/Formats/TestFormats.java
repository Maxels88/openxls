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
import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * This Class Demonstrates the basic functionality of of ExtenXLS Format Handling.
 */
public class TestFormats
{
	private static final Logger log = LoggerFactory.getLogger( TestFormats.class );
	String wd = System.getProperty( "user.dir" ) + "/docs/samples/Formats/";

	private String finpath = "", sheetname = "";

	/**
	 * read in an XLS file, output Cell vals, write it back to a new file
	 */
	public void testFormats()
	{
		String[] testfiles = {
				"testFormats.xls", "Sheet1",
		};

		for( int i = 0; i < testfiles.length; i++ )
		{
			finpath = testfiles[i];
			System.out.println( "====================== TEST: " + String.valueOf( i ) + " ======================" );
			sheetname = testfiles[++i];
			System.out.println( "=====================> " + finpath + ":" + sheetname );
			testBorders( finpath, sheetname );

			System.out.println( "DONE" );
		}

	}

	WorkSheetHandle sheet = null;

	void testBorders( String finpath, String sheetname )
	{
		WorkBookHandle tbo = new WorkBookHandle( finpath );
		try
		{
			sheet = tbo.getWorkSheet( sheetname );
			FormatHandle fmt1 = new FormatHandle( tbo );
			fmt1.setFont( "Arial", Font.PLAIN, 24 );
			fmt1.setForegroundColor( FormatHandle.COLOR_LIGHT_BLUE );
			fmt1.setFontColor( FormatHandle.COLOR_YELLOW );
			fmt1.setBackgroundPattern( FormatHandle.PATTERN_HOR_STRIPES3 );

			fmt1.setBorderLineStyle( FormatHandle.BORDER_MEDIUM_DASH_DOT_DOT );
			sheet.add( "NEW CELL!", "C5" );
			CellHandle bordercell = sheet.getCell( "C3" );

			int i = bordercell.getFormatId();
			fmt1.addCell( bordercell );
			i = bordercell.getFormatId();
			System.out.println( i );
		}
		catch( Exception e )
		{
			System.out.println( e );
		}
		testWrite( tbo, "testFormats_out.xls" );
	}

	void testFormats( String finpath, String sheetname )
	{
		WorkBookHandle tbo = new WorkBookHandle( finpath );

		int color = 5, patern = 25;
		FormatHandle fmt1 = new FormatHandle( tbo );
		fmt1.setFont( "Arial", Font.PLAIN, 10 );
		fmt1.setForegroundColor( FormatHandle.COLOR_LIGHT_BLUE );
		fmt1.setFontColor( FormatHandle.COLOR_YELLOW );
		fmt1.setBackgroundPattern( FormatHandle.PATTERN_HOR_STRIPES3 );
		FormatHandle fmt2 = new FormatHandle( tbo );
		try
		{
			sheet = tbo.getWorkSheet( sheetname );

			// TEST BORDERS
			CellHandle bordercell = sheet.getCell( "C3" );
			int borderfmt = bordercell.getFormatId();

			sheet.add( "MODDY BORDER", "A10" );
			bordercell = sheet.getCell( "A10" );
			FormatHandle fmx = new FormatHandle( tbo );
			fmx.setBorderLineStyle( 3 );
			fmx.setBorderTopColor( FormatHandle.COLOR_LIGHT_BLUE );
			fmx.setBorderBottomColor( FormatHandle.COLOR_GREEN );
			fmx.setBorderLeftColor( FormatHandle.COLOR_YELLOW );
			fmx.setBorderRightColor( FormatHandle.COLOR_BLACK );
			fmx.addCell( sheet.add( "Great new cell!!", "A1" ) );
			fmx.setFont( "Courier", Font.BOLD, 12 );
			fmx.setForegroundColor( FormatHandle.COLOR_BRIGHT_GREEN );
			fmx.setBackgroundColor( FormatHandle.COLOR_BLUE );
			fmx.setFontColor( FormatHandle.COLOR_BLUE );
			fmx.setBackgroundPattern( FormatHandle.PATTERN_HOR_STRIPES3 );

			CellHandle cell1 = null;
			sheet.setHeaderText( "Extentech Inc." );
			sheet.setFooterText( "Created by ExtenXLS:" + WorkBookHandle.getVersion() );
			String addr = "";
			for( int i = 1; i < 50; i++ )
			{
				addr = "h" + String.valueOf( i );
				sheet.add( new Long( (long) 3324.234 * i ), addr );
				cell1 = sheet.getCell( addr );
				cell1.setFormatHandle( fmx );

				addr = "i" + String.valueOf( i );
				sheet.add( "Hello World " + i, addr );
				cell1 = sheet.getCell( addr );
				cell1.setFormatHandle( fmt1 );
			}

			color = FormatHandle.COLOR_LIGHT_BLUE;
			patern = FormatHandle.PATTERN_LIGHT_DOTS;
			cell1 = sheet.getCell( "H14" );
			cell1.setFormatHandle( fmt2 );
			cell1.setVal( 44444 );
			cell1.setForegroundColor( FormatHandle.COLOR_DARK_RED );
			cell1.setBackgroundColor( FormatHandle.COLOR_DARK_YELLOW );
			cell1.setBackgroundPattern( patern );
			cell1.setFontColor( color );
			cell1.setVal( 555 );
			System.out.println( "CELL FCOLOR: " + cell1.getForegroundColor() );
			System.out.println( "CELL FCOLOR: " + cell1.getBackgroundColor() );
		}
		catch( Exception e )
		{
			System.out.println( e );
		}
		testWrite( tbo, "testFormats1_out.xls" );

		// FORMAT ID Sharing
		// the fastest, no fuss way to use formats
		// on multiple cells

		WorkSheetHandle sheet1 = null;
		try
		{
			sheet1 = tbo.getWorkSheet( "Sheet1" );
		}
		catch( Exception e )
		{
			System.out.println( e );
		}
		sheet1.add( "Eurostile Template Cell", "A1" );
		int SHAREDFORMAT = 0;
		CellHandle b = null;
		CellHandle a = null;
		try
		{
			b = sheet1.getCell( "A1" );
			b.setFont( "Eurostile", Font.BOLD, 14 );
			// set format options
			b.setForegroundColor( 30 );
			b.setBackgroundColor( 54 );
			b.setBackgroundPattern( 3 );

			SHAREDFORMAT = b.getFormatId();
			for( int t = 1; t <= 10; t++ )
			{
				sheet1.add( new Float( t * 67.5 ), "E" + t );
				a = sheet1.getCell( "E" + t );
				a.setFormatId( SHAREDFORMAT );
			}

			a.setFont( "Tango", Font.BOLD, 26 );
			a.setFontColor( 10 );
			a.setFormatPattern( "[h]:mm:ss" );
			a.setURL( "http://www.extentech.com/" );
			a.setScript( 2 );
			sheet1.moveCell( a, "A10" );
			a.getCol().setWidth( 8000 );
			tbo.copyWorkSheet( sheetname, sheetname + " Copy" );
		}
		catch( Exception e )
		{
			System.out.println( e );
		}

		// optimize String table by sharing dupe entries
		tbo.setDupeStringMode( WorkBookHandle.SHAREDUPES );
		tbo.setStringEncodingMode( WorkBookHandle.STRING_ENCODING_COMPRESSED );

		//iterate backgrounds
		for( int x = 0; x < 32; x++ )
		{
			sheet1.add( "Pattern# " + x, "F" + (x + 1) );
			sheet1.add( "Text", "G" + (x + 1) );
			try
			{
				CellHandle c = sheet1.getCell( "G" + (x + 1) );
				c.setBackgroundPattern( x );
				c.setFont( "Tango", Font.BOLD, 26 );
				c.setFontColor( 10 );
			}
			catch( Exception e )
			{
				System.out.println( e );
			}
		}

		//iterate colors
		for( int x = 0; x < 64; x++ )
		{
			try
			{
				sheet1.add( "Color# " + x, "C" + (x + 1) );
				sheet1.add( " ", "D" + (x + 1) );
				CellHandle c = sheet1.getCell( "D" + (x + 1) );
				c.setBackgroundPattern( x );
				c.setBackgroundColor( x );
				c.setForegroundColor( x );
			}
			catch( Exception e )
			{
				System.out.println( e );
			}
		}

		// remove a col and a row
		try
		{
			sheet1.removeCols( 4, 1 );
			sheet1.removeRow( 2 );
		}
		catch( Exception e )
		{
			System.out.println( e );
		}
		testWrite( tbo, "testFormats2_out.xls" );
	}

	public void testWrite( WorkBookHandle b, String fout )
	{
		try
		{
			java.io.File f = new java.io.File( wd + fout );
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