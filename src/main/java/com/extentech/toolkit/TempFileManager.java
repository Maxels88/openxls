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
/**
 * TempFileManager.java
 *
 *
 * Feb 27, 2012
 *
 *
 */
package com.extentech.toolkit;

import com.extentech.ExtenXLS.DocumentHandle;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

/**
 * The TempFileManager allows for consolidated handling of all TempFiles used by ExtenXLS.
 * <p/>
 * TempFileManager is pluggable and allows you to implement a custom TempFileGenerator and install using
 * System properties.
 * <p/>
 * ie:
 * <p/>
 * System.setProperty(TempFileManager.TEMPFILE_MANAGER_CLASSNAME, "com.acme.CustomTempFileGenerator");
 * WorkBookHandle bkx = new WorkBookHandle("test.xlsx"); // use custom tempfile generator
 */
public class TempFileManager
{

	public static String TEMPFILE_MANAGER_CLASSNAME = "com.extentech.extenxls.tempfilemanager";

	public static File createTempFile( String prefix, String extension ) throws IOException
	{
		String tmpfu = System.getProperty( TEMPFILE_MANAGER_CLASSNAME );
		if( tmpfu != null )
		{
			try
			{
				TempFileGenerator tgen = (TempFileGenerator) Class.forName( tmpfu ).newInstance();
				return tgen.createTempFile( prefix, extension );
			}
			catch( Exception e )
			{
				Logger.logErr( "Could not load custom TempFileGenerator: " + tmpfu + ". Falling back to default TempFileGenerator." );
			}
		}
		return new DefaultTempFileGeneratorImpl().createTempFile( prefix, extension );
	}

	/**
	 * Write an InputStream to disk, and return as a file handle.
	 *
	 * @param input
	 * @param prefix
	 * @param extension
	 * @return
	 */
	public static File createTempFile( InputStream input, String prefix, String extension ) throws IOException
	{
		File tmpfile = TempFileManager.createTempFile( prefix, extension );
		JFileWriter.writeToFile( input, tmpfile );
		return tmpfile;
	}

	/**
	 * Feb 27, 2012
	 *
	 * @param string
	 * @param string2
	 * @param dir
	 * @return
	 * @throws IOException
	 */
	public static File createTempFile( String prefix, String extension, File dir ) throws IOException
	{
		prefix = dir.getAbsolutePath() + prefix;
		return createTempFile( prefix, extension );
	}

	public static File writeToTempFile( String prefix, String extension, DocumentHandle doc ) throws IOException
	{
		File file = createTempFile( prefix, extension );

		BufferedOutputStream stream = new BufferedOutputStream( new FileOutputStream( file ) );
		doc.write( stream, DocumentHandle.FORMAT_NATIVE );

		stream.flush();
		stream.close();

		return file;
	}
}
