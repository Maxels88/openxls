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
package org.openxls.toolkit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.RandomAccessFile;
import java.io.StringBufferInputStream;
import java.io.StringReader;
import java.nio.channels.FileChannel;

/**
 * File utilities.
 */
public class JFileWriter
{
	private static final Logger log = LoggerFactory.getLogger( JFileWriter.class );

	java.lang.String path = "";
	java.lang.String filename = "";
	java.lang.String data = "";
	byte newLine = Character.LINE_SEPARATOR;

	public void setPath( String p )
	{
		path = p;
	}

	public void setFileName( String f )
	{
		filename = f;
	}

	public void setData( String d )
	{
		data = d;
	}

	static void printErr( String err )
	{
		log.error( "Error in JFileWriter: " + err );
	}

	/**
	 * append text to the end of a text file
	 */
	public static final synchronized void appendToFile( String pth, String text )
	{
		try
		{
			byte[] bbuf = text.getBytes( "UTF-8" );
			File outp = new File( pth );

			if( !outp.exists() )
			{
				outp.mkdirs();
				outp.delete();
				outp = new java.io.File( pth );
			}

			RandomAccessFile outputFile = new RandomAccessFile( outp, "rw" );
			outputFile.skipBytes( (int) outputFile.length() );
			int strt = 0;
			if( outp.exists() )
			{
				strt = (int) outputFile.length();
			}
			outputFile.write( bbuf, 0, bbuf.length );
			outputFile.close();
		}
		catch( Exception e )
		{
			log.error( "JFileWriter.appendToFile() IO Error : " + e.toString(), e );
		}
	}

	/**
	 * write the inputstream contents to file
	 *
	 * @param is
	 * @param file
	 * @throws IOException
	 */
	public static void writeToFile( InputStream is, File file ) throws IOException
	{
		DataOutputStream out = new DataOutputStream( new BufferedOutputStream( new FileOutputStream( file ) ) );

		// Transfer bytes from in to out
		byte[] buf = new byte[1024];
		int len;
		while( (len = is.read( buf )) > 0 )
		{
			out.write( buf, 0, len );
		}
		is.close();
		out.close();
	}

	public boolean writeIt()
	{
		try
		{
			path += filename;
			StringReader SR = new StringReader( data );
			File outputFile = new File( path );
			FileWriter out = new FileWriter( outputFile );
			int c;
			if( outputFile.length() > 0 )
			{
				return false;
			}
			while( (c = SR.read()) != -1 )
			{
				out.write( c );
			}
			out.flush();
			out.close();
		}
		catch( IOException e )
		{
			log.error( "JFileWriter IO Error : " + e.toString(), e );
		}
		return true;
	}

	public static boolean writeIt( String data, String filename, String path )
	{
		try
		{
			path += filename;
			StringReader SR = new StringReader( data );
			File outputFile = new File( path );
			FileWriter out = new FileWriter( outputFile );
			int c;
			if( outputFile.length() > 0 )
			{
				return false;
			}
			while( (c = SR.read()) != -1 )
			{
				out.write( c );
			}
			out.flush();
			out.close();
		}
		catch( IOException e )
		{
			log.error( "JFileWriter IO Error : " + e.toString(), e );
		}
		return true;
	}

	public String readFile( String fname )
	{
		StringBuffer addTxt = new StringBuffer();
		try
		{
			BufferedReader d = new BufferedReader( new FileReader( fname ) );

			while( d.ready() )
			{
				addTxt.append( d.readLine() );
			}
			d.close();
		}
		catch( Exception e )
		{
			printErr( "problem reading file: " + e );
		}
		return addTxt.toString();
	}

	public static void copyFile( String infile, String outfile ) throws FileNotFoundException, IOException
	{
		File fx = new File( infile );
		copyFile( fx, outfile );

		// this.writeLine(outfile, this.readFile(infile));
	}

	/**
	 * Copy method, using FileChannel#transferTo
	 * NOTE:  will overwrite existing files
	 *
	 * @param File source
	 * @param File target
	 * @throws IOException
	 */
	public static void copyFile( File source, String target ) throws FileNotFoundException, IOException
	{
		File fout = new File( target );
		fout.mkdirs();
		fout.delete();
		fout = new File( target );
		FileChannel in = new FileInputStream( source ).getChannel();
		FileChannel out = new FileOutputStream( target ).getChannel();
		in.transferTo( 0, in.size(), out );
		in.close();
		out.close();
	}

	public void writeLine( String file, String line )
	{
		String s;
		try
		{
			File f = new File( file );
			// f.mkdirs();
			FileWriter out = new FileWriter( f );
			DataInputStream inStream = new DataInputStream( new StringBufferInputStream( line ) );
			while( (s = inStream.readLine()) != null )
			{
				out.write( s );
				out.write( newLine );
			}
			out.close();
		}
		catch( FileNotFoundException e )
		{
			printErr( e.toString() );
		}
		catch( Exception e )
		{
			printErr( e.toString() );
		}
	}

	public void writeLogToFile( String fname, javax.swing.JTextArea jta )
	{
		try
		{
			OutFile n = new OutFile( fname );
			String logText = jta.getText();
			n.writeBytes( logText );
			jta.setText( "" );
			n.close();
		}
		catch( FileNotFoundException e )
		{
			printErr( e.toString() );
		}
		catch( IOException e )
		{
			printErr( e.toString() );
		}
	}

	public String readLog( String logFname )
	{
		String addTxt = "";
		try
		{
			InFile n = new InFile( logFname );
			while( n.available() != 0 )
			{
				addTxt += n.readLine();
			}
		}
		catch( FileNotFoundException e )
		{
			printErr( e.toString() );
		}
		catch( IOException e )
		{
			printErr( e.toString() );
		}
		return addTxt += "\r\n";
	}
}