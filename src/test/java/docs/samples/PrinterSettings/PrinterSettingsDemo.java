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
package docs.samples.PrinterSettings;

import com.extentech.ExtenXLS.PrinterSettingsHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * Demonstrates the operation of the PrinterSettingsHandle.
 * <br/><br/>
 * The PrinterSettingsHandle provides fine-grained control over printing settings
 * in Excel.
 * <br/><br/>
 * NOTE: you can only view the effects of these methods in an open Excel file when
 * you use the "Print Setup" command.
 * <br/><br/>
 * ExtenXLS does not currently support directly sending
 * spreadsheet data to a printer.
 */
public class PrinterSettingsDemo
{
	private static final Logger log = LoggerFactory.getLogger( PrinterSettingsDemo.class );
	String outputdir = System.getProperty( "user.dir" ) + "/docs/samples/PrinterSettings/";

	/**
	 * Test functionality of setting Printer Settings file for a Spreadsheet
	 *
	 * @param args
	 */
	public static void main( String[] args )
	{
		try
		{
			log.info( "Begin Demo of Printer Settings..." );
			// load the current file and output it to same directory
			WorkBookHandle bk = new WorkBookHandle( System.getProperty( "user.dir" ) + "/docs/samples/PrinterSettings/InvoiceTemplate.xls" );

			// execute the printer settings methods on the book
			PrinterSettingsDemo demo = new PrinterSettingsDemo();
			demo.testPrinterSettings( bk );

			log.info( "Successfully Set Printer Settings for Spreadsheet: " + bk.toString() );
		}
		catch( Exception e )
		{
			log.error( "Spreadsheet Printer Settings Demo failed.", e );
		}

	}

	/**
	 * test the various printer settings
	 * ------------------------------------------------------------
	 */
	public void testPrinterSettings( WorkBookHandle book )
	{
		WorkSheetHandle sheet = null;
		PrinterSettingsHandle printersetup = null;
		try
		{
			sheet = book.getWorkSheet( 0 );
			for( int x = 0; x < 10; x++ )
			{
				for( int t = 0; t < 10; t++ )
				{
					sheet.add( "Hello World " + t, t, x );
				}
			}
			printersetup = sheet.getPrinterSettings();
		}
		catch( Exception e )
		{
			log.error( "testPrinterSettings failed: " + e.toString() );
		}

		// fit width
		printersetup.setFitWidth( 3 );
		// fit height
		printersetup.setFitHeight( 5 );
		// header margin
		printersetup.setHeaderMargin( 1.025 );
		// footer margin
		printersetup.setFooterMargin( 1.025 );
		// number of copies
		printersetup.setCopies( 10 );
		// Paper Size
		printersetup.setPaperSize( PrinterSettingsHandle.PAPER_SIZE_LEDGER_17x11 );
		// Scaling
		printersetup.setScale( 125 );
		//	resolution
		printersetup.setResolution( 300 );

		// GRBIT settings:
		// left to right printing
		printersetup.setLeftToRight( true );
		// print as draft quality
		printersetup.setDraft( true );
		// black and white
		printersetup.setNoColor( true );
		// landscape / portrait
		printersetup.setLandscape( true );

		// write it out
		testWrite( book, "PrinterSettings_out.xls" );

		// read it in
		book = new WorkBookHandle( this.outputdir + "PrinterSettings_out.xls" );
		try
		{
			sheet = book.getWorkSheet( 0 );
			printersetup = sheet.getPrinterSettings();
		}
		catch( Exception e )
		{
			log.error( "testPrinterSettings failed: " + e.toString() );
		}

		// header margin
		log.info( "Header Margin: " + printersetup.getHeaderMargin() );
		// footer margin
		log.info( "Header Margin: " + printersetup.getFooterMargin() );

		// assertions

		// fit width
		assertEquals( (short) 3, printersetup.getFitWidth() );
		// fit height
		assertEquals( (short) 5, printersetup.getFitHeight() );
		// number of copies
		assertEquals( (short) 10, printersetup.getCopies() );
		// Paper Size
		assertEquals( (short) PrinterSettingsHandle.PAPER_SIZE_LEDGER_17x11, printersetup.getPaperSize() ); // TODO: find out what these are
		// Scaling
		assertEquals( (short) 125, printersetup.getScale() );
		//	resolution
		assertEquals( (short) 300, printersetup.getResolution() );

		// left to right printing
		assertEquals( true, printersetup.getLeftToRight() );
		// print as draft quality
		assertEquals( true, printersetup.getDraft() );
		// No color
		assertEquals( true, printersetup.getNoColor() );
		// landscape / portrait
		assertEquals( true, printersetup.getLandscape() );

	}

	/**
	 * asserts that the two objects are equal
	 * Jan 27, 2010
	 *
	 * @param o1
	 * @param o2
	 */
	private void assertEquals( int o1, int o2 )
	{
		if( o1 != o2 )
		{
			log.warn( "Values not equal:" + o1 + "!=" + o2 );
		}
	}

	/**
	 * asserts that the two objects are equal
	 * Jan 27, 2010
	 *
	 * @param o1
	 * @param o2
	 */
	private void assertEquals( boolean o1, boolean o2 )
	{
		if( o1 != o2 )
		{
			log.warn( "Values not equal:" + o1 + "!=" + o2 );
		}
	}

	public void testWrite( WorkBookHandle b, String fout )
	{
		try
		{
			java.io.File f = new java.io.File( outputdir + fout );
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
