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
 * Logger.java
 *
 *
 *
 */
package com.extentech.toolkit;

import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.PrintStream;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.nio.charset.Charset;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * System-wide Logging facility
 * <p/>
 * <br>
 * Logger can be used to output standardized messages to System.out and System.err, as well as
 * to pluggable LogOutputter implementations.
 * <br><br>
 * To install a custom LogOutputter implementation, instantiate a class that implements LogOutputter, then
 * set the system property: "com.extentech.toolkit.logger"
 * <br><br>
 * For example:
 * <pre>
 *  CustomLog mylogr = new CustomLog();
 *  Properties props = System.getProperties();
 * props.put("com.extentech.toolkit.logger", mylogr );
 *  </pre>
 * <br>
 * The default Logger settings can be controlled using System properties.
 * <pre>
 * 	props.put("com.extentech.toolkit.logger.dateformat", "MMM yyyy mm:ss" );
 * 	props.put("com.extentech.toolkit.logger.dateformat", "none" );
 *
 * </pre>
 */
public class Logger extends PrintStream implements LogOutputter
{
	/**
	 * @deprecated Just use <code>this</code>.
	 */
	protected PrintStream ous = this;

	/**
	 * Copy of <code>line.separator</code> system property to save lookups.
	 */
	private static final String endl = System.getProperty( "line.separator" );

	private LogOutputter targetLogger;
	private BufferedWriter targetWriter;
	private StringBuffer lineBuffer = new StringBuffer();
	private boolean autoFlush;

	public Logger( LogOutputter target )
	{
		this();
		this.init( target );
	}

	public Logger( LogOutputter target, String charset ) throws UnsupportedEncodingException
	{
		this( charset );
		this.init( target );
	}

	public Logger( OutputStream target )
	{
		this( target, false );
	}

	public Logger( OutputStream target, boolean autoFlush )
	{
		this( new OutputStreamWriter( target ), autoFlush );
	}

	public Logger( OutputStream target, String charset, boolean autoFlush ) throws UnsupportedEncodingException
	{
		this( new OutputStreamWriter( target, charset ), charset, autoFlush );
	}

	public Logger( Writer target )
	{
		this( target, false );
	}

	public Logger( Writer target, boolean autoFlush )
	{
		this();
		this.init( target, autoFlush );
	}

	public Logger( Writer target, String charset, boolean autoFlush ) throws UnsupportedEncodingException
	{
		this( charset );
		this.init( target, autoFlush );
	}

	private Logger()
	{
		super( new IndirectOutputStream(), true );
		((IndirectOutputStream) out).setSink( new WriterOutputStream( this, Charset.defaultCharset() ) );
	}

	private Logger( String charset ) throws UnsupportedEncodingException
	{
		super( new IndirectOutputStream(), true, charset );
		((IndirectOutputStream) out).setSink( new WriterOutputStream( this, charset ) );
	}

	private void init( LogOutputter target )
	{
		targetLogger = target;
		targetWriter = null;
		autoFlush = false; // has no meaning for LogOutputter target
	}

	private void init( Writer target, boolean autoFlush )
	{
		targetLogger = null;
		targetWriter = new BufferedWriter( target );
		this.autoFlush = autoFlush;
	}

	/**
	 * Installs this logger as the default logger and replaces the standard
	 * output and error streams.
	 */
	public void install()
	{
		setLogger( this );
		System.setOut( this );
		System.setErr( this );
	}
	
	/* ---------- LogOutputter methods ---------- */

	@Override
	public void log( String message )
	{
		if( null != targetLogger )
		{
			targetLogger.log( message );
		}
		else
		{
			synchronized(targetWriter)
			{
				try
				{
					targetWriter.write( getLogDate() );
					targetWriter.write( " " );
					targetWriter.write( message );
					targetWriter.newLine();
					if( autoFlush )
					{
						targetWriter.flush();
					}
				}
				catch( IOException ex )
				{
					// we're the logger, so we can't exactly log about it
					// the interface doesn't support exceptions so just drop it
				}
			}
		}
	}

	@Override
	public void log( String message, Exception ex, boolean trace )
	{
		if( null != targetLogger )
		{
			targetLogger.log( message, ex, trace );
		}
		else
		{
			this.log( formatThrowable( message, ex, trace ) );
		}
	}

	@Override
	public void log( String message, Exception ex )
	{
		if( null != targetLogger )
		{
			targetLogger.log( message, ex );
		}
		else
		{
			this.log( formatThrowable( message, ex, false ) );
		}
	}
	
	/* ---------- PrintStream methods ---------- */

	public void logLine()
	{
		synchronized(lineBuffer)
		{
			// if the line buffer ends with a newline, strip it
			int length = lineBuffer.length();
			if( length >= endl.length() && endl.equals( lineBuffer.substring( length - endl.length(), length ) ) )
			{
				lineBuffer.setLength( length - endl.length() );
			}

			// log and reset the line buffer but don't log empty lines
			if( lineBuffer.length() > 0 )
			{
				this.log( lineBuffer.toString() );
				lineBuffer.setLength( 0 );
			}
		}
	}

	@Override
	public Logger append( char value )
	{
		lineBuffer.append( value );
		return this;
	}

	@Override
	public Logger append( CharSequence value )
	{
		lineBuffer.append( value );
		return this;
	}

	@Override
	public Logger append( CharSequence value, int start, int end )
	{
		lineBuffer.append( value, start, end );
		return this;
	}

	@Override
	public void print( boolean b )
	{
		lineBuffer.append( b );
	}

	@Override
	public void print( char c )
	{
		lineBuffer.append( c );
	}

	@Override
	public void print( int i )
	{
		lineBuffer.append( i );
	}

	@Override
	public void print( long l )
	{
		lineBuffer.append( l );
	}

	@Override
	public void print( float f )
	{
		lineBuffer.append( f );
	}

	@Override
	public void print( double d )
	{
		lineBuffer.append( d );
	}

	@Override
	public void print( char[] s )
	{
		lineBuffer.append( s );
	}

	@Override
	public void print( String s )
	{
		synchronized(lineBuffer)
		{
			lineBuffer.append( s );
			if( s.endsWith( endl ) )
			{
				this.println();
			}
		}
	}

	@Override
	public void print( Object obj )
	{
		lineBuffer.append( obj );
	}

	@Override
	public void println( boolean x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println( char x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println( int x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println( long x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println( float x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println( double x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println( char[] x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println( String x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println( Object x )
	{
		lineBuffer.append( x );
		this.println();
	}

	@Override
	public void println()
	{
		synchronized(lineBuffer)
		{
			// flush the input stream into the line buffer
			super.flush();

			// log the current line
			this.logLine();
		}
	}
	
	/* ---------- static convenience methods for logging ---------- */

	public static final String INFO_STRING = "";
	public static final String WARN_STRING = "WARNING: ";
	public static final String ERROR_STRING = "ERROR: ";

	/**
	 * Gets the current system logger.
	 */
	public static LogOutputter getLogger()
	{
		LogOutputter logger;

		try
		{
			logger = (LogOutputter) System.getProperties().get( "com.extentech.toolkit.logger" );
		}
		catch( Exception ex )
		{
			logger = null;
		}

		if( null == logger )
		{
			if( System.err instanceof Logger )
			{
				logger = (Logger) System.err;
			}
			else
			{
				logger = new Logger( System.err, true );
			}
			setLogger( logger );
		}

		return logger;
	}

	/**
	 * Replaces the system logger.
	 */
	public static void setLogger( LogOutputter logger )
	{
		System.getProperties().put( "com.extentech.toolkit.logger", logger );
	}

	public static String formatThrowable( String message, Throwable ex, boolean trace )
	{
		StringWriter writer = new StringWriter();
		writer.write( message );

		if( trace )
		{
			writer.write( endl );
			writer.write( endl );

			PrintWriter printer = new PrintWriter( writer );
			ex.printStackTrace( printer );
			printer.flush();
		}
		else
		{
			writer.write( ": " );
			writer.write( ex.toString() );
		}

		return writer.toString();
	}

	/**
	 * Logs a fatal error message to the system logger.
	 */
	public static void logErr( String message, Exception ex )
	{
		getLogger().log( ERROR_STRING + message, ex );
	}

	/**
	 * Logs a fatal error message to the system logger.
	 */
	public static void logErr( String message, Throwable ex )
	{
		getLogger().log( formatThrowable( ERROR_STRING + message, ex, false ) );
	}

	/**
	 * Logs a fatal error message to the system logger.
	 */
	public static void logErr( String message )
	{
		getLogger().log( ERROR_STRING + message );
	}

	/**
	 * Logs a fatal error message to the system logger.
	 */
	public static void logErr( String message, Exception ex, boolean trace )
	{
		getLogger().log( ERROR_STRING + message, ex, trace );
	}

	/**
	 * Logs the string conversion of an object to the system logger.
	 */
	public static void log( Object object )
	{
		logInfo( object.toString() );
	}

	/**
	 * Logs a non-fatal warning to the system logger.
	 */
	public static void logWarn( String message )
	{
		getLogger().log( WARN_STRING + message );
	}

	/**
	 * Logs the string conversion of an exception to the system logger as a
	 * fatal error message.
	 */
	public static void logErr( Exception ex )
	{
		logErr( ex.toString() );
	}

	/**
	 * Logs an informational message to the system logger.
	 */
	public static void logInfo( String message )
	{
		getLogger().log( INFO_STRING + message );
	}

	/**
	 * Attempts to replace the standard output stream with a
	 * <code>Logger</code> instance that writes to the named file. If the
	 * operation fails a message will be logged to the system logger and the
	 * method will return without throwing an exception.
	 */
	public static void setOut( String filename )
	{
		try
		{
			java.io.File logfile = new java.io.File( filename );
			FileOutputStream sysout = new FileOutputStream( logfile );
			System.setOut( new Logger( sysout ) );
		}
		catch( Exception e )
		{
			Logger.logErr( "Setting System Output Stream in Logger failed: ", e );
		}
	}

	/**
	 * Attempts to replace the standard error stream with a
	 * <code>Logger</code> instance that writes to the named file. If the
	 * operation fails a message will be logged to the system logger and the
	 * method will return without throwing an exception.
	 */
	public static void setErr( String filename )
	{
		try
		{
			java.io.File logfile = new java.io.File( filename );
			FileOutputStream sysout = new FileOutputStream( logfile );
			System.setErr( new Logger( sysout ) );
		}
		catch( Exception e )
		{
			Logger.logErr( "Setting System Error Stream in Logger failed: ", e );
		}
	}

	/**
	 * The default time stamp format for {@link #getLogDate()}.
	 */
	public static final String DATE_FORMAT = "yyyy-MM-dd HH:mm:ss:SSSS";

	private static SimpleDateFormat dateFormat = new SimpleDateFormat( DATE_FORMAT );
	private static String dateSpec = DATE_FORMAT;

	/**
	 * Returns the current time in a configurable format.
	 * If the system property <code>com.extentech.toolkit.logger.dateformat</code>
	 * exists and is a valid date format pattern it will be used. Otherwise the
	 * {@linkplain #DATE_FORMAT default format pattern} will be used.
	 */
	public static String getLogDate()
	{
		String spec = System.getProperty( "com.extentech.toolkit.logger.dateformat" );
		if( null != spec )
		{
			if( "none".equalsIgnoreCase( spec ) )
			{
				return "";
			}

			if( !dateSpec.equals( spec ) )
			{
				try
				{
					dateFormat.applyPattern( spec );
				}
				catch( IllegalArgumentException e )
				{
					dateFormat.applyPattern( DATE_FORMAT );
				}

				dateSpec = spec;
			}
		}

		return dateFormat.format( new Date() );
	}

}
