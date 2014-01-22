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
 * <b>thrown when trying to access an Image and it is not found.</b>
 *
 * @see ImageHandle
 * @see WorkBook
 */

public final class ImageNotFoundException extends java.lang.Exception
{

	/**
	 *
	 *
	 */
	private static final long serialVersionUID = -5031711537364049574L;
	String description = "";

	public ImageNotFoundException( String n )
	{
		super();
		description = n;
	}

	@Override
	public String getMessage()
	{
		return this.toString();
	}

	public String toString()
	{
		return description;
	}

}