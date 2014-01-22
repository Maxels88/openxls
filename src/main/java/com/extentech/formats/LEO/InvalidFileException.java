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
package com.extentech.formats.LEO;

/**
 * This Exception is thrown whenever a LEOFile cannot be parsed or contains invalid data.
 */
public class InvalidFileException extends RuntimeException
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2547604121903676738L;
	private String err = "Not a valid OLE File.";

	public InvalidFileException( String er )
	{
		this.err = er;
	}

	public String toString()
	{
		return "InvalidFileException: " + err;
	}

}
