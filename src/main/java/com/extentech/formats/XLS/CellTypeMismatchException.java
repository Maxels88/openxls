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
 * Thrown when trying to change the value of a BiffRec to an incompatible
 * data type.
 * <p/>
 * For example: changing a float BiffRec value to a String containing non-numeric
 * characters would throw one of these.
 */
public class CellTypeMismatchException extends java.lang.NumberFormatException
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8664358251873265167L;
	String cellname = "";

	public CellTypeMismatchException( String n )
	{
		super();
		this.cellname = n;
	}

	public String toString()
	{
		return "Cell Type Mismatch Exception trying to set value of Cell: " + cellname;
	}

	public String getMessage()
	{
		return toString();
	}
}
