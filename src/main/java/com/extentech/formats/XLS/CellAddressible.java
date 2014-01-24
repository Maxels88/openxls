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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
		private final int row;
		private final int colFirst;
		private final int colLast;

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

		@Override
		public int getColFirst()
		{
			return colFirst;
		}

		@Override
		public int getColLast()
		{
			return colLast;
		}

		@Override
		public int getRowNumber()
		{
			return row;
		}

		@Override
		public boolean isSingleCol()
		{
			return (getColFirst() == getColLast());
		}

		@Override
		public String toString()
		{
			return "Reference{" +
					"row=" + row +
					", cols=" + colFirst +
					"-" + colLast +
					'}';
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
		 * @param before whether the boundary should sort before (true) or after (false) a range which includes it
		 */
		public RangeBoundary( int row, int col, boolean before )
		{
			super( row, col, col );
			this.before = before;
		}

		public int compareToRange()
		{
			return before ? -1 : 1;
		}

		@Override
		public String toString()
		{
			return "RangeBoundary{" +
					"before=" + before +
					"} " + super.toString();
		}
	}

	/**
	 * {@link Comparator} that sorts cells in ascending row major order.
	 */
	public static final class RowMajorComparator implements java.util.Comparator<CellAddressible>, Serializable
	{
		private static final long serialVersionUID = 5477030152120715766L;
		private static final ColumnRange.Comparator colComp = new ColumnRange.Comparator();

		@Override
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
		private static final Logger log = LoggerFactory.getLogger( ColumnMajorComparator.class );
		private static final long serialVersionUID = -1193867650674693873L;
		private static final ColumnRange.Comparator colComp = new ColumnRange.Comparator();

		@Override
		public int compare( CellAddressible cell1, CellAddressible cell2 )
		{
			// Primary sort is on column number...
			int cols = colComp.compare( cell1, cell2 );
			if( 0 != cols )
			{
				log.trace( "Comparing: {} to {} - ColCompare: {}", cell1, cell2, cols );
				return cols;
			}

			// If the column comparison is equal, then we do a secondary comparison on row number...
			int rowNum1 = cell1.getRowNumber();
			int rowNum2 = cell2.getRowNumber();

			int rowCompare = Integer.compare( rowNum1, rowNum2 );

			log.trace( "Comparing: {} to {} - RowCompare: {}", cell1, cell2, rowCompare );
			return rowCompare;
		}
	}
}
