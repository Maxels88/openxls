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

import java.io.Serializable;

/**
 * Coordinates the various records involved in book-level protection.
 */
public class BookProtectionManager extends ProtectionManager implements Serializable
{
	/**
	 * The workbook whose protection state this instance manages.
	 */
	private WorkBook book;

	public BookProtectionManager( WorkBook target )
	{
		book = target;
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	@Override
	public void close()
	{
		book = null;
	}
}
