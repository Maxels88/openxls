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
package docs.samples.Excel2007;

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;

/**
 * demonstration of usage with Excel 2007 / 2010 file types
 * <br>
 * NOTE: support for XLSX further requires:
 * <installdir>/lib/jaxen-1.1-beta-5.jar
 * <installdir>/lib/xpp3-1.1.3.4.O.jar
 */
public class Excel2007Test
{
	private static final Logger log = LoggerFactory.getLogger( Excel2007Test.class );
	public WorkBookHandle book;
	public WorkSheetHandle sheet;
	private static boolean WRITE_AUTO = true;
	private static boolean WRITE_XLSX = false;
	private static boolean WRITE_XLS = false;

	String wd = System.getProperty( "user.dir" ) + "/docs/samples/Excel2007/";

	public String finpath = "testXLSX_template.xlsx";
	public String foutpath = "testXLSX_template_out.xlsx";

	/**
	 * Test ExtenXLS handling of Excel 2007 (XLSX) files
	 * ------------------------------------------------------------
	 *
	 * @param args
	 */
	public static void main( String[] args )
	{
		Excel2007Test test = new Excel2007Test();
		test.testRead();
		log.info( "Read: " + test.book.toString() + " SUCCESS!" );
		test.testWrite();
		log.info( "Write: " + test.book.toString() + " SUCCESS!" );

	}

	/**
	 * Reads in the default file
	 * ------------------------------------------------------------
	 */
	public void testRead()
	{
		book = new WorkBookHandle( wd + finpath );
	}

	/**
	 * test write outputs XLS files from test
	 * ------------------------------------------------------------
	 */
	public void testWrite()
	{

		try
		{
			java.io.File f = new java.io.File( wd + foutpath );

			// create directories
			try
			{
				if( !f.exists() )
				{
					f.mkdirs();
					f.delete();
				}
			}
			catch( Exception e )
			{
				;
			}

			FileOutputStream fos = new FileOutputStream( f );

			// if the file is Excel 2007 format, we use "write" method
			// which auto-detects spreadsheet format
			if( book.getIsExcel2007() )
			{
				log.info( book.toString() + " is an Excel 2007/2010 XLSX spreadsheet." );
			}

			if( WRITE_AUTO )
			{ // to auto-detect file type, use WorkBookHandle.write(OutputStream out) method
				book.write( fos );
			}
			else if( WRITE_XLS )
			{ // we can write out XLS explicitly if we like
				BufferedOutputStream bbout = new BufferedOutputStream( fos );
				book.write( bbout );
				bbout.flush();
				bbout.close();
				fos.flush();
				fos.close();
			}
			else if( WRITE_XLSX )
			{ // we can write out XLSX explicitly if we like
				int format = WorkBookHandle.FORMAT_XLSX;
				if( foutpath.toLowerCase().endsWith( ".xlsm" ) )
				{
					format = WorkBookHandle.FORMAT_XLSM;
				}
				BufferedOutputStream bbout = new BufferedOutputStream( fos );
				book.write( bbout, format );
				bbout.flush();
				bbout.close();
				fos.flush();
				fos.close();

			}

			book = new WorkBookHandle( wd + foutpath );

		}
		catch( Exception e )
		{
			log.error( "Exception in Excel2007Test.", e );
		}
	}

}