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

/**
 * Coordinates the various records involved in protection.
 */
public abstract class ProtectionManager
{
	protected Protect protect;
	protected FeatHeadr enhancedProtection;
	protected Password password;

	/**
	 * Adds a protection-related record to be managed.
	 * This method will be called automatically by the <code>init()</code>
	 * method of the record where appropriate. It should probably never be
	 * called anywhere else.
	 *
	 * @param record the record to be managed
	 */
	public void addRecord( BiffRec record )
	{
		if( record instanceof Protect )
		{
			protect = (Protect) record;
		}
		else if( record instanceof Password )
		{
			password = (Password) record;
		}
	}

	/**
	 * Returns whether the entity is protected.
	 */
	public boolean getProtected()
	{
		return protect != null && protect.getIsLocked();
	}

	/**
	 * Sets whether the sheet is protected.
	 */
	public void setProtected( boolean value )
	{
		protect.setLocked( value );
	}

	/**
	 * Returns the entity's password verifier.
	 *
	 * @return the password verifier as four upper-case hexadecimal digits
	 * or "0000" if no password is set on the sheet
	 */
	public String getPassword()
	{
		if( password == null )
		{
			return "0000";
		}
		return password.getPasswordHashString();
	}

	/**
	 * Sets or removes the entity's protection password.
	 *
	 * @param pass the string password to set or null to remove the password
	 */
	public void setPassword( String pass )
	{
		password.setPassword( pass );
	}

	/**
	 * Sets or removes the entity's protection password.
	 *
	 * @param pass the pre-hashed string password to set
	 *             or null to remove the password
	 */
	public void setPasswordHashed( String pass )
	{
		password.setHashedPassword( pass );
	}

	/**
	 * Checks whether the given password matches the protection password.
	 *
	 * @param guess the password to be checked against the stored hash
	 * @return whether the given password matches the stored hash
	 */
	public boolean checkPassword( String guess )
	{
		if( password == null )
		{
			return guess == null || guess.equals( "" );
		}
		return password.validatePassword( guess );
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		if( enhancedProtection != null )
		{
			enhancedProtection.close();
		}
		if( password != null )
		{
			password.close();
		}
		if( protect != null )
		{
			protect.close();
		}
	}
}
