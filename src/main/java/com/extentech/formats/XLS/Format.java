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
package com.extentech.formats.XLS;

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.StringTool;

/**
 * Stores a custom number format pattern.
 */
public final class Format extends XLSRecord
{
	private static final long serialVersionUID = 1199947552103220748L;

	private short ifmt = -1;
	private String pattern;
	private Unicodestring ustring;

	public Format()
	{
		setOpcode( XLSConstants.FORMAT );
	}

	/**
	 * Makes a new Format record for the given pattern.
	 *
	 * @param book    the workbook to which the pattern should belong
	 * @param pattern the number format pattern to ensure exists
	 */
	public Format( WorkBook book, String pattern )
	{
		this( book, (short) -1, pattern );
	}

	/**
	 * Makes a new Format record with the given ID and pattern.
	 * This should only be used when parsing non-BIFF8 files. BIFF8 parsing
	 * will use the normal XLSRecord init sequence. For programmatic creation
	 * of custom formats use {@link #Format(WorkBook, String)} instead.
	 *
	 * @param book    the workbook to which the pattern should belong
	 * @param id      the format ID to use or -1 to generate one
	 * @param pattern the number format pattern to ensure exists
	 */
	public Format( WorkBook book, short id, String pattern )
	{
		this();
		setWorkBook( book );

		this.pattern = pattern;
		this.ustring = Sst.createUnicodeString( pattern, null, WorkBook.STRING_ENCODING_AUTO );

		byte[] idbytes;
		if( id > 0 )
		{
			this.ifmt = id;
			idbytes = ByteTools.shortToLEBytes( id );
		}
		else
		{
			// WorkBook.insertFormat will call setIfmt
			idbytes = new byte[2];
		}

		setData( ByteTools.append( ustring.read(), idbytes ) );

		book.insertFormat( this );
	}

	/**
	 * Initializes the record from bytes.
	 * This method should only be called as part of the normal XLSRecord
	 * init sequence when parsing from bytes.
	 *
	 * @throws IllegalStateException if the record has already been parsed
	 */
	public void init()
	{
		if( pattern != null )
		{
			throw new IllegalStateException( "the record has already been parsed" );
		}

		super.init();

		ifmt = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );

		ustring = new Unicodestring();
		ustring.init( this.getBytesAt( 2, this.getLength() - 2 ), false );

		pattern = ustring.toString();
		
		/* Strip double quotes and backslashes from the format string.
		 * The quoting characters are an integral part of the format string,
		 * so this is almost certainly wrong. However, it's what the previous
		 * implementation did and I'm trying to preserve behavior.
		 * TODO: revisit stripping of quotes when writing number format parser
		 * - Sam
		 */
		pattern = StringTool.replaceText( pattern, "\"", "", 0 );
		pattern = StringTool.replaceText( pattern, "\\", "", 0 );

		this.getWorkBook().addFormat( this );
	}

	public void setWorkBook( WorkBook book )
	{
		super.setWorkBook( book );

        /* Not sure why this is here, but I think it has something to do with
         * worksheet cloning. It's harmless, so might as well leave it alone.
         * - Sam
         */
		if( ifmt != -1 )
		{
			book.addFormat( this );
		}
	}

	/**
	 * Sets the format ID of this Format record.
	 */
	public void setIfmt( short id )
	{
		ifmt = id;

		// Update the record bytes
		System.arraycopy( ByteTools.shortToLEBytes( ifmt ), 0, this.getData(), 0, 2 );
	}

	/**
	 * Gets the format pattern string represented by this Format.
	 */
	public String getFormat()
	{
		return pattern;
	}

	/**
	 * Gets the format index of this Format in its workbook.
	 */
	public short getIfmt()
	{
		return ifmt;
	}

	public String toString()
	{
		return getFormat();
	}
}