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
package org.openxls.formats.XLS;

/**
 * Validation Exceptions are thrown when a cell value is set to a value that does not pass
 * validity  of an Excel validation record affecting said cell
 */
public class ValidationException extends Exception
{
	private static final long serialVersionUID = -6448974788123912538L;

	private String errorTitle = "";
	private String errorText = "";

	public ValidationException( String eTitle, String eText )
	{
		super();
		errorTitle = eTitle;
		errorText = eText;
	}

	/**
	 * Returns the title of the validation error dialog.
	 */
	public String getTitle()
	{
		return errorTitle;
	}

	/**
	 * Returns the body of the validation error dialog.
	 */
	public String getText()
	{
		return errorText;
	}

	@Override
	public String getMessage()
	{
		return toString();
	}

	public String toString()
	{
		return errorTitle + ": " + errorText;
	}

}
