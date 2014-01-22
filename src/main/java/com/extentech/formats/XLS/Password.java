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

import java.io.UnsupportedEncodingException;

/**
 * Password - specifies the password verifier for the sheet or workbook.
 * This record is required in the Globals sub-stream. If the workbook has no
 * password the value written is zero. If a sheet has a password this record
 * appears in its sub-stream. In a sheet sub-stream the value may not be zero.
 *
 * @see MS-XLS §2.4.191
 */
public final class Password extends XLSRecord
{
	private static final long serialVersionUID = 380140909635538745L;

	private short hash;

	public Password()
	{
		super();
		setOpcode( PASSWORD );
		setLength( (short) 2 );
		this.originalsize = 2;
	}

	/**
	 * Checks whether the given password matches the stored one.
	 * This hashes the given guess and then compares it against the stored
	 * value. Note that the hash is only two bytes so collisions are likely.
	 *
	 * @param guess the clear text password guess to be verified
	 * @return whether the hash of the guess matches the stored value
	 */
	public boolean validatePassword( String guess )
	{
		return hashPassword( guess ) == getPasswordHash();
	}

	/**
	 * Returns the password verifier value.
	 * This is the password after being passed through a rather strange hash
	 * function that reduces it to a two byte integer.
	 */
	public short getPasswordHash()
	{
		return hash;
	}

	/**
	 * Gets the password verifier as a hexadecimal string.
	 *
	 * @return the verifier as four upper-case hexits (0123456789ABCDEF)
	 * @see #getPasswordHash()
	 */
	public String getPasswordHashString()
	{
		String raw = Integer.toHexString( hash & 0xFFFF ).toUpperCase();
		return "0000".substring( 0, 4 - raw.length() ).concat( raw );
	}

	/**
	 * Hashes the given password with the Excel password verifier algorithm.
	 * <p>The algorithm is specified by MS-XLS §2.2.9. It defers to ECMA-376-1
	 * part 4 §3.2.29 for character encoding and MS-OFFCRYPTO §2.3.7.1 for the
	 * hash function itself. That section of ECMA-376-1 also specifies the hash
	 * function and is currently technically compatible with MS-OFFCRYPTO.
	 * MS-XLS appears to indicate that MS-OFFCRYPTO is the normative spec.</p>
	 */
	protected static short hashPassword( String password )
	{
		byte[] strBytes;
		int hash;

        /* Encode the password string as CP1252.
         * According to ECMA-376-1 16-bit Unicode characters should be encoded
         * in the best fit Windows "ANSI" code page. Java does not have the
         * best fit code page selection algorithm built in. CP1252 is the
         * correct choice for most users; other code pages are used for
         * non-Latin scripts, primarily Asian languages. We'll just use it for
         * everything until someone complains.
         */
		try
		{
			strBytes = password.getBytes( "windows-1252" );
		}
		catch( UnsupportedEncodingException e )
		{
			// The JVM doesn't support CP1252, ISO-8859-1 is almost identical
			try
			{
				strBytes = password.getBytes( "ISO-8859-1" );
			}
			catch( UnsupportedEncodingException ex )
			{
				// ISO-8859-1 is one of the mandatory character sets.
				throw new Error( "ISO-8859-1 charset is missing" );
			}
		}

		// start with a hash value of zero
		hash = 0;

		// iterate backwards over the password bytes starting with the last one
		for( int idx = strBytes.length - 1; idx >= 0; idx-- )
		{
			// bitwise XOR the hash with the current password byte
			hash ^= strBytes[idx];

			// rotate the 15 lowest-order bits (mask 0x7FFF) left by one place
			hash = (hash << 1) & 0x7FFF | ((hash >>> 14) & 0x0001);
		}

		// bitwise XOR the hash with the length of the password
		hash ^= strBytes.length & 0xFFFF;

		// ECMA-376 specifies this as (0x8000 | ('N' << 8) | 'K'), which
		// always evaluates to 0xCE4B. MS-OFFCRYPTO uses 0xCE4B directly.
		hash ^= 0xCE4B;

		return (short) hash;
	}

	/**
	 * Sets the stored password.
	 *
	 * @param password the clear text of the password to be applied
	 *                 or null to remove the existing password
	 */
	public void setPassword( String password )
	{
		hash = hashPassword( password );
		updateRecord();
	}

	/**
	 * Sets the stored password.
	 *
	 * @param hash the hash of the password to be applied
	 *             or zero to remove the existing password
	 */
	public void setHashedPassword( short hash )
	{
		this.hash = hash;
		updateRecord();
	}

	/**
	 * Sets the stored password.
	 *
	 * @param hash the four-digit hex hash of the password to be applied
	 *             or "0000" to remove the existing password
	 */
	public void setHashedPassword( String hash )
	{
		this.hash = (short) Integer.parseInt( hash, 16 );
		updateRecord();
	}

	private void updateRecord()
	{
		setData( ByteTools.shortToLEBytes( hash ) );
	}

	public void init()
	{
		super.init();
		hash = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );

		Boundsheet sheet = this.getSheet();
		if( sheet != null )
		{
			sheet.getProtectionManager().addRecord( this );
		}
	}
}