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
package com.extentech.toolkit;

import java.io.BufferedInputStream;
import java.io.DataInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

public class InFile extends DataInputStream
{

	StringBuffer sb = new StringBuffer();

	public InFile( String filename ) throws FileNotFoundException
	{
		super( new BufferedInputStream( new FileInputStream( new File( filename ) ) ) );
	}

    /* public InFile(String filename)
    throws FileNotFoundException {
        super(new BufferedInputStream(new FileInputStream(new File(filename))));
    }
    */

	public InFile( File file ) throws FileNotFoundException
	{
		this( file.getPath() );
	}

	/**
	 * Reads File from Disk
	 *
	 * @param fname path to file
	 */
	public String readFile()
	{
		try
		{
			while( this.available() != 0 )
			{
				sb.append( this.readLine() );
			}
		}
		catch( FileNotFoundException e )
		{
			Logger.logInfo( "FNF Exception in InFile: " + e );
		}
		catch( IOException e )
		{
			Logger.logInfo( "IO Exception in InFile: " + e );
		}
		return sb.toString();
	}

	/**
	 * Gets a byte arrray from a file
	 *
	 * @param file File the file to get bytes from
	 * @return byte[] Returns byte[] array file contents
	 */
	public static byte[] getBytesFromFile( File file ) throws IOException
	{
		InputStream fis = new FileInputStream( file );
		long length = file.length();
		byte[] ret = new byte[(int) length];
		int offset = 0;
		int numRead = 0;
		while( (offset < ret.length) && ((numRead = fis.read( ret, offset, ret.length - offset )) >= 0) )
		{
			offset += numRead;
		}
		if( offset < ret.length )
		{
			throw new IOException( "Read file failed -- all bytes not retreived. " + file.getName() );
		}
		fis.close();
		return ret;

	}

}

