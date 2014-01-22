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
 * BOMInputstream.java
 *
 *
 * Oct 14, 2010
 *
 *
 */
package com.extentech.formats.XML;

/**
 Original pseudocode   : Thomas Weidenfeller
 Implementation tweaked: Aki Nieminen

 http://www.unicode.org/unicode/faq/utf_bom.html
 BOMs:
 00 00 FE FF    = UTF-32, big-endian
 FF FE 00 00    = UTF-32, little-endian
 FE FF          = UTF-16, big-endian
 FF FE          = UTF-16, little-endian
 EF BB BF       = UTF-8

 Win2k Notepad:
 Unicode format = UTF-16LE
 ***/

import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;

/**
 * This inputstream will recognize unicode BOM marks
 * and will skip bytes if getEncoding() method is called
 * before any of the read(...) methods.
 * <p/>
 * Usage pattern:
 * String enc = "ISO-8859-1"; // or NULL to use
 * systemdefault
 * FileInputStream fis = new FileInputStream(file);
 * UnicodeInputStream uin = new UnicodeInputStream(fis,
 * enc);
 * enc = uin.getEncoding(); // check for BOM mark and skip
 * bytes
 * InputStreamReader in;
 * if (enc == null) in = new InputStreamReader(uin);
 * else in = new InputStreamReader(uin, enc);
 */
public class UnicodeInputStream extends InputStream
{
	PushbackInputStream internalIn;
	boolean isInited = false;
	String defaultEnc;
	String encoding;

	private static final int BOM_SIZE = 4;

	public UnicodeInputStream( InputStream in, String defaultEnc )
	{
		internalIn = new PushbackInputStream( in, BOM_SIZE );
		this.defaultEnc = defaultEnc;
	}

	public String getDefaultEncoding()
	{
		return defaultEnc;
	}

	public String getEncoding()
	{
		if( !isInited )
		{
			try
			{
				init();
			}
			catch( IOException ex )
			{
				throw new IllegalStateException( "Init method failed." + ex );
//              (Throwable)ex);
			}
		}
		return encoding;
	}

	/**
	 * Read-ahead four bytes and check for BOM marks. Extra
	 * bytes are
	 * unread back to the stream, only BOM bytes are skipped.
	 */
	protected void init() throws IOException
	{
		if( isInited )
		{
			return;
		}

		byte bom[] = new byte[BOM_SIZE];
		int n, unread;
		n = internalIn.read( bom, 0, bom.length );

		if( (bom[0] == (byte) 0xEF) && (bom[1] == (byte) 0xBB) &&
				(bom[2] == (byte) 0xBF) )
		{
			encoding = "UTF-8";
			unread = n - 3;
		}
		else if( (bom[0] == (byte) 0xFE) && (bom[1] == (byte) 0xFF) )
		{
			encoding = "UTF-16BE";
			unread = n - 2;
		}
		else if( (bom[0] == (byte) 0xFF) && (bom[1] == (byte) 0xFE) )
		{
			encoding = "UTF-16LE";
			unread = n - 2;
		}
		else if( (bom[0] == (byte) 0x00) && (bom[1] == (byte) 0x00) &&
				(bom[2] == (byte) 0xFE) && (bom[3] == (byte) 0xFF) )
		{
			encoding = "UTF-32BE";
			unread = n - 4;
		}
		else if( (bom[0] == (byte) 0xFF) && (bom[1] == (byte) 0xFE) &&
				(bom[2] == (byte) 0x00) && (bom[3] == (byte) 0x00) )
		{
			encoding = "UTF-32LE";
			unread = n - 4;
		}
		else
		{
			// Unicode BOM mark not found, unread all bytes
			encoding = defaultEnc;
			unread = n;
		}
		// System.out.println("read=" + n + ", unread=" + unread);

		if( unread > 0 )
		{
			internalIn.unread( bom, (n - unread), unread );
		}

		isInited = true;
	}

	@Override
	public void close() throws IOException
	{
		//init();
		isInited = true;
		internalIn.close();
	}

	@Override
	public int read() throws IOException
	{
		//init();
		isInited = true;
		return internalIn.read();
	}
}
