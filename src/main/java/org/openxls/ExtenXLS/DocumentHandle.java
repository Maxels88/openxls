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
package org.openxls.ExtenXLS;

import org.openxls.formats.LEO.LEOFile;
import org.openxls.formats.XLS.OOXMLAdapter;
import org.openxls.toolkit.TempFileManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.Closeable;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URL;
import java.net.URLConnection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

/**
 * Functionality common to all document types.
 */
public abstract class DocumentHandle implements Document, Handle, Closeable
{
	private static final Logger log = LoggerFactory.getLogger( DocumentHandle.class );
	/**
	 * Format constant for the most appropriate format for this document.
	 * If the document was read in from a file, this is usually the format that
	 * was read in.
	 */
	public static final int FORMAT_NATIVE = 0;

	/**
	 * The user-visible display name or title of this document.
	 */
	protected String name = null;

	/**
	 * The file associated with this document.
	 * This will generally be the file the document was parsed from, if any.
	 */
	protected File file;

	/**
	 * Store for workbook properties.
	 */
	private Map<String, Object> props = new HashMap<>();

	/**
	 * Handling for a streaming worksheet based workbook *
	 */
	private boolean streamingSheets = false;

	/**
	 * default constructor
	 */
	public DocumentHandle()
	{
		super();

	}

	/**
	 * Apr 5, 2011
	 *
	 * @param urlx
	 */
	public DocumentHandle( InputStream urlx )
	{
		super();
		// TODO Auto-generated constructor stub
		log.error( "DocumentHandle InputStream Constructor Not Implemented" );
	}

	/**
	 * Retrieves a property in the workbook property store.
	 * This is not an Excel-compatible feature.
	 *
	 * @param name the name of the property to retrieve
	 * @return the value of the requested property or null if it doesn't exist
	 */
	@Override
	public Object getProperty( String name )
	{
		return props.get( name );
	}

	/**
	 * Sets the value of a property in the workbook property store.
	 * This is not an Excel-compatible feature.
	 *
	 * @param name  the name of the property which should be updated
	 * @param value the value to which the property should be set
	 */
	@Override
	public void addProperty( String name, Object val )
	{
		props.put( name, val );
	}

	/**
	 * Retrieves a Map containing the workbook properties store.
	 * This is not an Excel-compatible feature.
	 *
	 * @return an immutable Map containing the current workbook properties
	 */
	public Map<String, Object> getProperties()
	{
		return Collections.unmodifiableMap( props );
	}

	/**
	 * Replaces the workbook properties with the values in a given Map.
	 * This is not an Excel-compatible feature.
	 *
	 * @param properties the values that will replace the existing properties
	 */
	public void setProperties( Map<String, Object> properties )
	{
		props = new HashMap<>();
		props.putAll( properties );
	}

	/**
	 * Gets the ExtenXLS version number.
	 */
	public static String getVersion()
	{
		return GetInfo.getVersion();
	}

	/**
	 * Sets the user-visible descriptive name or title of this document.
	 * Some formats will persist this setting in the document itself.
	 */
	@Override
	public void setName( String nm )
	{
		name = nm;
	}

	/**
	 * Handling for streaming sheets.  Currently this is in development and unsupported
	 *
	 * @param streamSheets
	 */
	public void setStreamingSheets( boolean streamSheets )
	{
		streamingSheets = streamSheets;
	}

	/**
	 * Sets the file name associated with this document.
	 *
	 * @deprecated Use {@link #setFile(File)} instead.
	 */
	public void setFileName( String name )
	{
		file = new File( name ).getAbsoluteFile();
	}

	/**
	 * Sets the file associated with this document.
	 */
	public void setFile( File file )
	{
		this.file = file;
	}

	/**
	 * Gets the file associated with this document.
	 * For documents read in from a file, this defaults to that file. If no
	 * file is associated with this document, for example if the document was
	 * parsed from a stream, this may return <code>null</code>.
	 */
	public File getFile()
	{
		return file;
	}

	/**
	 * Looks for magic numbers in the given input data and attempts to parse
	 * it with an appropriate <code>DocumentHandle</code> subclass. Detection
	 * is performed on a best-effort basis and is not guaranteed to be accurate.
	 *
	 * @throws IOException       if an error occurs while reading from the stream
	 * @throws WorkBookException if parsing fails
	 */
	public static DocumentHandle getInstance( InputStream input ) throws IOException
	{
		BufferedInputStream bufferedStream = new BufferedInputStream( input );
		// read in that start of the file for checking magic numbers
		byte[] headerBytes;
		int count;
		// make sure the file is long enough to get magic numbers
		bufferedStream.mark( 1028 );
		headerBytes = new byte[512];
		count = bufferedStream.read( headerBytes );
		bufferedStream.reset();

		// if it starts with the LEO magic number check the header
		if( LEOFile.checkIsLEO( headerBytes, count ) )
		{
			LEOFile leo = new LEOFile( bufferedStream );

			if( leo.hasWorkBook() )
			{
				return new WorkBookHandle( leo );
			}
			throw new WorkBookException( "input is LEO but no supported format detected", -1 );
		}

		String headerString;
		try
		{
			headerString = new String( headerBytes, 0, count, "UTF-8" );
		}
		catch( UnsupportedEncodingException e )
		{
			// UTF-8 support is required by the JLS
			throw new Error( "the JVM does not support UTF-8", e );
		}

		// if it's a ZIP archive, try parsing as OOXML
		if( headerString.startsWith( "PK" ) )
		{
			return new WorkBookHandle( bufferedStream );
		}

		if( (headerString.indexOf( "," ) > -1) && (headerString.indexOf( "," ) > -1) )
		{
			// init a blank workbook
			WorkBookHandle book = new WorkBookHandle();

			// map CSV into workbook
			try
			{
				WorkSheetHandle sheet = book.getWorkSheet( 0 );
				sheet.readCSV( new BufferedReader( new InputStreamReader( bufferedStream ) ) );
				return book;
			}
			catch( Exception e )
			{
				throw new WorkBookException( "Error encountered importing CSV: " + e.toString(), WorkBookException.ILLEGAL_INIT_ERROR );
			}
		}
		throw new WorkBookException( "unknown file format", -1 );
	}

	/**
	 * Gets the file name associated with this document.
	 * For documents read in from a file, this defaults to that file. If no
	 * file is associated with this document, for example if the document was
	 * parsed from a stream, this may return <code>null</code>.
	 *
	 * @deprecated Use {@link #getFile()} instead.
	 */
	public String getFileName()
	{
		return (file != null) ? file.getPath() : "New Document.doc";
	}

	/**
	 * Downloads the resource at the given URL to a temporary file.
	 *
	 * @param u the URL representing the resource to be downloaded
	 * @return the path to a temporary file containing the downloaded resource
	 * or <code>null</code> if an error occurred
	 * @deprecated The download should be handled outside ExtenXLS.
	 * There is no specific replacement for this method.
	 */
	@Deprecated
	protected static File getFileFromURL( URL u )
	{
		try
		{
			File fx = TempFileManager.createTempFile( "upload-" + System.currentTimeMillis(), ".tmp" );

			URLConnection uc = u.openConnection();
			String contentType = uc.getContentType();
			int contentLength = uc.getContentLength();
			if( contentType.startsWith( "text/" ) || (contentLength == -1) )
			{
				throw new IOException( "This is not a binary file." );
			}
			InputStream raw = uc.getInputStream();
			InputStream in = new BufferedInputStream( raw );
			byte[] data = new byte[contentLength];
			int bytesRead = 0;
			int offset = 0;
			while( offset < contentLength )
			{
				bytesRead = in.read( data, offset, data.length - offset );
				if( bytesRead == -1 )
				{
					break;
				}
				offset += bytesRead;
			}
			in.close();

			if( offset != contentLength )
			{
				throw new IOException( "Only read " + offset + " bytes; Expected " + contentLength + " bytes" );
			}

			// String filename = u.getFile().substring(filename.lastIndexOf('/') + 1);
			FileOutputStream out = new FileOutputStream( fx );
			out.write( data );
			out.flush();
			out.close();
			return fx;
		}
		catch( Exception e )
		{
			log.error( "Could not load WorkBook from URL: " + e.toString() );
			return null;
		}
	}

	/**
	 * Gets the user-visible descriptive name or title of this document.
	 */
	@Override
	public String getName()
	{
		if( name != null )
		{
			return name;
		}
		return "Untitled Document";
	}

	/**
	 * Resets the document state to what it was when it was loaded.
	 *
	 * @throws UnsupportedOperationException if there is not sufficient data
	 *                                       available to perform the reversion
	 */
	@Override
	public abstract void reset();

	/**
	 * Gets the constant representing this document's native format.
	 */
	public abstract int getFormat();

	/**
	 * Gets the file name extension for this document's native format.
	 */
	public abstract String getFileExtension();

	/**
	 * Writes the document to the given stream in the requested format.
	 *
	 * @param dest   the stream to which the document should be written
	 * @param format the constant representing the desired output format
	 * @throws IllegalArgumentException if the given type code is invalid
	 * @throws IOException              if an error occurs while writing to the stream
	 */
	@Override
	public abstract void write( OutputStream dest, int format ) throws IOException;

	/**
	 * Writes the document to the given stream in its native format.
	 *
	 * @param dest the stream to which the document should be written
	 * @throws IOException if an error occurs while writing to the stream
	 */
	@Override
	public void write( OutputStream dest ) throws IOException
	{
		write( dest, FORMAT_NATIVE );
	}

	/**
	 * Writes the document to the given file in the requested format.
	 *
	 * @param file   the path to which the document should be written
	 * @param format the constant representing the desired output format
	 * @throws IllegalArgumentException if the given type code is invalid
	 * @throws IOException              if an error occurs while writing to the file
	 */
	@Override
	public void write( File file, int format ) throws IOException
	{
		if( (format > WorkBookHandle.FORMAT_XLS) && (this.file != null) )
		{
			OOXMLAdapter.refreshPassThroughFiles( (WorkBookHandle) this );
		}

		if( file.exists() )
		{
			boolean deleted = file.delete();// try this
			if(!deleted)
			{
			log.warn( "Attempt to delete file {} failed.", file.getAbsolutePath() );
			}
		}
		OutputStream stream = new BufferedOutputStream( new FileOutputStream( file ) );
		write( stream, format );
		this.file = file;    // necesary for OOXML re-write ...
		stream.flush();
		stream.close();
	}

	/**
	 * Writes the document to the given file in its native format.
	 *
	 * @param file the path to which the document should be written
	 * @throws IOException if an error occurs while writing to the stream
	 */
	@Override
	public void write( File file ) throws IOException
	{
		write( file, FORMAT_NATIVE );
	}

	/**
	 * Returns a string representation of the object.
	 * This is currently equivalent to {@link #getName()}.
	 */
	public String toString()
	{
		return getName();
	}
}