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

import java.io.Serializable;
import java.util.Comparator;

/**
 * A <code>Comparator</code> that sorts cell records by address.
 * As <code>XLSRecord</code> currently uses default equality the
 * imposed ordering is inconsistent with equals.
 */
public class CellAddressComparator implements Comparator, Serializable
{
	private static final long serialVersionUID = 686639297047358268L;

	/**
	 * Compares its two arguments for order.
	 * The arguments must be <code>XLSRecord</code>s that represent cells
	 * (<code>isValueForCell()</code> must return true).
	 *
	 * @param o1 the first cell record to be compared
	 * @param o2 the second cell record to be compared
	 * @throws ClassCastException if either argument is not a cell record
	 */
	@Override
	public int compare( Object o1, Object o2 )
	{
		if( o1 == null || !(o1 instanceof XLSRecord) || o2 == null || !(o2 instanceof XLSRecord) )
		{
			throw new ClassCastException();
		}

		XLSRecord rec1 = (XLSRecord) o1;
		XLSRecord rec2 = (XLSRecord) o2;

		if( !rec1.isValueForCell() || !rec2.isValueForCell() )
		{
			throw new ClassCastException();
		}

		int diff = rec1.getRowNumber() - rec2.getRowNumber();
		if( diff != 0 )
		{
			return diff;
		}

		return rec1.getColNumber() - rec2.getColNumber();
	}

	/**
	 * Indicates whether another <code>Comparator</code> imposes the same
	 * ordering as this one. As all instances of this class impose the same
	 * order, this method returns true if and only if the given object is
	 * an instance of this class.
	 */
	public boolean equals( Object obj )
	{
		return obj != null && this.getClass().equals( obj.getClass() );
	}

	/**
	 * Returns a hash code value for this object.
	 * As all instances of this class are equivalent this returns a
	 * constant value.
	 */
	public int hashCode()
	{
		// This is just a random number
		return 973216835;
	}
}