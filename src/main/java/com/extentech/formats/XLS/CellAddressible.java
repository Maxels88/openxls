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

/**
 * Something with a cell address.
 */
public interface CellAddressible extends ColumnRange
{

	/**
	 * Returns the row the cell resides in.
	 *
	 * @return the zero-based index of the cell's parent row
	 */
	public int getRowNumber();

	/**
	 * A simple immutable cell reference useful as a map key.
	 */
	public static class Reference implements CellAddressible, Serializable
	{
		private static final long serialVersionUID = -9071483662123798966L;
		private final int row, colFirst, colLast;

		public Reference( int row, int colFirst, int colLast )
		{
			this.row = row;
			this.colFirst = colFirst;
			this.colLast = colLast;
		}

		public Reference( int row, int col )
		{
			this( row, col, col );
		}

		public int getColFirst()
		{
			return this.colFirst;
		}

		public int getColLast()
		{
			return this.colLast;
		}

		public int getRowNumber()
		{
			return this.row;
		}

		public boolean isSingleCol()
		{
			return (this.getColFirst() == this.getColLast());
		}
	}

	/**
	 * Cell reference for use as the boundary of a column range.
	 */
	public static class RangeBoundary extends Reference
	{
		private static final long serialVersionUID = -1357617242449928095L;
		private final boolean before;

		/**
		 * @param before whether the boundary should sort before (true) or
		 *               after (false) a range which includes it
		 */
		public RangeBoundary( int row, int col, boolean before )
		{
			super( row, col, col );
			this.before = before;
		}

		public int comareToRange()
		{
			return before ? -1 : 1;
		}
	}

	/**
	 * {@link Comparator} that sorts cells in ascending row major order.
	 */
	public static final class RowMajorComparator implements java.util.Comparator<CellAddressible>, Serializable
	{
		private static final long serialVersionUID = 5477030152120715766L;
		private static final ColumnRange.Comparator colComp = new ColumnRange.Comparator();

		public int compare( CellAddressible cell1, CellAddressible cell2 )
		{
			int rows = cell1.getRowNumber() - cell2.getRowNumber();
			if( 0 != rows )
			{
				return rows;
			}

			return colComp.compare( cell1, cell2 );
		}
	}

	/**
	 * {@link Comparator} that sorts cells in ascending column major order.
	 */
	public static final class ColumnMajorComparator implements java.util.Comparator<CellAddressible>, Serializable
	{
		private static final long serialVersionUID = -1193867650674693873L;
		private static final ColumnRange.Comparator colComp = new ColumnRange.Comparator();

		public int compare( CellAddressible cell1, CellAddressible cell2 )
		{
			int cols = colComp.compare( cell1, cell2 );
			if( 0 != cols )
			{
				return cols;
			}

			return cell1.getRowNumber() - cell2.getRowNumber();
		}
	}
}
