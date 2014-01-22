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

public interface ColumnRange
{
	/**
	 * Gets the first column in the range.
	 */
	public int getColFirst();

	/**
	 * Gets the last column in the range.
	 */
	public int getColLast();

	/**
	 * determines if its a single column *
	 */
	public boolean isSingleCol();

	/**
	 * A lightweight immutable <code>ColumnRange</code> useful as a reference
	 * element for searching collections.
	 */
	public static final class Reference implements ColumnRange, Serializable
	{
		private static final long serialVersionUID = -2240322394559418980L;
		private final int first, last;

		public Reference( int first, int last )
		{
			this.first = first;
			this.last = last;
		}

		public int getColFirst()
		{
			return first;
		}

		public int getColLast()
		{
			return last;
		}

		public boolean isSingleCol()
		{
			return (first == last);
		}
	}

	public static final class Comparator implements java.util.Comparator<ColumnRange>, Serializable
	{
		private static final long serialVersionUID = -4506187924019516336L;

		/**
		 * This comparator will return equal if one of the column
		 * ranges passed in is a single column reference and lies within
		 * the bounds of the second column reference.  This is required for equals
		 * within the cell collection.  If it turns out to be an issue (colinfos?) we should
		 * separate this out.
		 */
		public int compare( ColumnRange r1, ColumnRange r2 )
		{
			boolean single1 = r1.isSingleCol();
			boolean single2 = r2.isSingleCol();

			// if we're comparing a single column to a range
			if( (single1 || single2) && single1 != single2 )
			{ // XOR
				ColumnRange range = (single1 ? r2 : r1);
				ColumnRange single = (single1 ? r1 : r2);

				// and the single column falls within the range
				if( range.getColFirst() <= single.getColFirst() && range.getColLast() >= single.getColLast() )
				{

					// if it's a range boundary, it chooses what it is
					if( single instanceof CellAddressible.RangeBoundary )
					{
						int value = ((CellAddressible.RangeBoundary) single).comareToRange();

						// it needs to be reversed if the range was first
						return (single2 ? -value : value);
					}

					// otherwise it's equal
					return 0;
				}
			}

			int first = r1.getColFirst() - r2.getColFirst();
			if( 0 != first )
			{
				return first;
			}
			return r2.getColLast() - r1.getColLast();
		}
	}
}
