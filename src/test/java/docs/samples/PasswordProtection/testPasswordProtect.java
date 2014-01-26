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
package docs.samples.PasswordProtection;

import org.openxls.ExtenXLS.CellHandle;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import junit.framework.Assert;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;

/**
 * Tests the password protection functionality of the WorkSheetHandle
 * ------------------------------------------------------------
 */
public class testPasswordProtect
{
	private static final Logger log = LoggerFactory.getLogger( testPasswordProtect.class );
	static String wd = System.getProperty( "user.dir" ) + "/docs/samples/PasswordProtection/";

	public static void main( String[] args )
	{
		WorkBookHandle book = new WorkBookHandle();
		try
		{
			WorkSheetHandle sheet = book.getWorkSheet( "Sheet1" );
			CellHandle a1 = sheet.add( "hello world", "a1" );
			sheet.setProtected( true );
			writeFile( book, "testPasswordProtect.xls" );
			book.close();
			book = new WorkBookHandle( wd + "testPasswordProtect.xls" );
			sheet = book.getWorkSheet( "Sheet1" );
			if( !sheet.getProtected() )
			{
				log.error( "Set Password Protection Failed!" );
			}

			// set password/get password
			sheet.setProtectionPassword( "g0away" );

		}
		catch( Exception ex )
		{
			log.error( "error opening password protected file " + ex.toString() );
		}
	}

	/**
	 * write the file to disk
	 * ------------------------------------------------------------
	 *
	 * @param workBookHandle
	 * @param excelFileName
	 */
	private static void writeFile( WorkBookHandle workBookHandle, String excelFileName )
	{
		try
		{
			File outputFile = new File( wd + excelFileName );
			log.info( "Begin TestPasswordProtect." );
			FileOutputStream fileOutputStream = new FileOutputStream( outputFile );
			BufferedOutputStream bufferedOutputStream = new BufferedOutputStream( fileOutputStream );

			workBookHandle.write( bufferedOutputStream );

			bufferedOutputStream.flush();
			fileOutputStream.close();
			log.info( "TestPasswordProtect done." );
		}
		catch( java.io.IOException e )
		{
			Assert.fail( "Exception thrown when trying to write the file: " + e );
		}
	}

}
