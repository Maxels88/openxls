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
package com.extentech.ExtenXLS;

import com.extentech.formats.XLS.XLSConstants;

import java.io.File;
import java.io.IOException;
import java.io.OutputStream;

/**
 * An interface representing an ExtenXLS document.
 *
 * @deprecated This interface provides no functionality beyond the abstract
 * {@link DocumentHandle} class. Use that type instead.
 */
@Deprecated
public interface Document
{
	public static final int DEBUG_LOW = XLSConstants.DEBUG_LOW;
	public static final int DEBUG_MEDIUM = XLSConstants.DEBUG_MEDIUM;
	public static final int DEBUG_HIGH = XLSConstants.DEBUG_HIGH;

	/**
	 * get a non-Excel property
	 *
	 * @return Returns the properties.
	 */
	public Object getProperty( String name );

	/**
	 * add non-Excel property
	 *
	 * @param properties The properties to set.
	 */
	public void addProperty( String name, Object val );

	/** The Session for the WorkBook instance
	 *
	 * @return public BookSession getSession();
	 */

	/**
	 * Sets the internal name of this WorkBookHandle.
	 * <p/>
	 * Overrides the default for 'getName()' which returns
	 * the file name source of this WorkBook by default.
	 *
	 * @param WorkBook Name
	 */
	public abstract void setName( String nm );

	/**
	 * Set the Debugging level.  Higher values output more
	 * debugging info during execution.
	 *
	 * @parameter int Debug level.  higher=more verbose
	 */
	public abstract void setDebugLevel( int l );

	/**
	 * Returns the name of this WorkBook
	 *
	 * @return String name of WorkBook
	 */
	public abstract String getName();

	/**
	 * Clears all values in a template WorkBook.
	 * <p/>
	 * Use this method to 'reset' the values of your
	 * WorkBook in memory to defaults.
	 * <p/>
	 * For example, if you load a Servlet with a
	 * single WorkBookHandle instance, then modify
	 * values and stream to a Client system, yo
	 * should call 'clearAll()' when the request
	 * is completed to remove the modified values
	 * and set them back to a default.
	 */
	public abstract void reset();

	/**
	 * Writes the document to the given stream in the requested format.
	 *
	 * @param dest   the stream to which the document should be written
	 * @param format the constant representing the desired output format
	 * @throws IllegalArgumentException if the given type code is invalid
	 * @throws IOException              if an error occurs while writing to the stream
	 */
	public void write( OutputStream dest, int format ) throws IOException;

	/**
	 * Writes the document to the given stream in its native format.
	 *
	 * @param dest the stream to which the document should be written
	 * @throws IOException if an error occurs while writing to the stream
	 */
	public void write( OutputStream dest ) throws IOException;

	/**
	 * Writes the document to the given file in the requested format.
	 *
	 * @param file   the path to which the document should be written
	 * @param format the constant representing the desired output format
	 * @throws IllegalArgumentException if the given type code is invalid
	 * @throws IOException              if an error occurs while writing to the file
	 */
	public void write( File file, int format ) throws IOException;

	/**
	 * Writes the document to the given file in its native format.
	 *
	 * @param file the path to which the document should be written
	 * @throws IOException if an error occurs while writing to the stream
	 */
	public void write( File file ) throws IOException;
}