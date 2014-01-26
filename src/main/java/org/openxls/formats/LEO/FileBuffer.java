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
/*
 * Created on Apr 25, 2005
 *
 * TODO To change the template for this generated file go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
package org.openxls.formats.LEO;

import org.openxls.toolkit.TempFileManager;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.channels.FileChannel;

/**
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
public class FileBuffer
{

	private transient ByteBuffer buffer = null;
	private File tempfile = null;
	FileChannel channel = null;
	FileInputStream input = null;

	public void close() throws IOException
	{
		input.close();
		channel.close();
		if( tempfile != null )
		{
			tempfile.deleteOnExit();
			tempfile.delete();
		}
		tempfile = null;
	}

	/**
	 *
	 */
	public FileBuffer()
	{
		super();
		// TODO Auto-generated constructor stub
	}

	public static FileBuffer readFile( String fpath )
	{
		try
		{
			File fx0 = new File( fpath );
			return readFile( fx0 );
		}
		catch( Throwable e )
		{
			throw new InvalidFileException( "LEO FileBuffer.readFile() failed: " + e.toString() );
		}
	}

	public static FileBuffer readFile( File fx0 )
	{
		try
		{
			FileBuffer fb = new FileBuffer();
			fb.input = new FileInputStream( fx0 );
			fb.channel = fb.input.getChannel();
			int fileLength = (int) fb.channel.size();
			// MappedByteBuffer 
			fb.buffer = fb.channel.map( FileChannel.MapMode.READ_ONLY, 0, fileLength );
			fb.buffer.order( ByteOrder.LITTLE_ENDIAN );
			return fb;
		}
		catch( Throwable e )
		{
			throw new InvalidFileException( "LEO FileBuffer.readFile() failed: " + e.toString() );
		}
	}

	//TODO: reimplement temp files and deal with cleanup -jm
	public static FileBuffer readFileUsingTemp( String fpath )
	{
		return readFileUsingTemp( new File( fpath ) );
	}

	public static FileBuffer readFileUsingTemp( File fx0 )
	{
		try
		{
			FileBuffer fb = new FileBuffer();
			// create Temp file and populate
			fb.tempfile = TempFileManager.createTempFile( "LEOFile_", ".tmp" );
			fb.tempfile.delete();

			FileInputStream input0 = new FileInputStream( fx0 );
			FileChannel channel0 = input0.getChannel();
			FileOutputStream output0 = new FileOutputStream( fb.tempfile );
			FileChannel channel1 = output0.getChannel();
			channel0.transferTo( 0, fx0.length(), channel1 );

			channel0.close();
			channel1.close();
			input0.close();
			output0.close();

			fb.input = new FileInputStream( fb.tempfile );
			fb.channel = fb.input.getChannel();
			int fileLength = (int) fb.channel.size();
			// MappedByteBuffer 
			fb.buffer = fb.channel.map( FileChannel.MapMode.READ_ONLY, 0, fileLength );
			fb.buffer.order( ByteOrder.LITTLE_ENDIAN );
			return fb;
		}
		catch( Throwable e )
		{
			throw new InvalidFileException( "LEO FileBuffer.readFile() failed: " + e.toString() );
		}
	}

	/**
	 * @return Returns the buffer.
	 */
	public ByteBuffer getBuffer()
	{
		return buffer;
	}

	/**
	 * @param buffer The buffer to set.
	 */
	public void setBuffer( ByteBuffer b )
	{
		buffer = b;
	}
}
