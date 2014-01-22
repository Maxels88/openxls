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
package com.extentech.ExtenXLS;

import com.extentech.formats.XLS.WorkSheetNotFoundException;

/**
 * Represents a reference to a 3D range of cells.
 * This class is not currently part of the public API. No input validation
 * whatsoever is performed. Any combination of values may be set, whether
 * it makes any sense or not.
 */
public class CellRangeRef implements Cloneable
{
	private int first_col, first_row, last_col, last_row;
	private String first_sheet_name, last_sheet_name;
	private WorkSheetHandle first_sheet, last_sheet;

	/**
	 * Private nullary constructor for use by static pseudo-constructors.
	 */
	private CellRangeRef()
	{
	}

	public CellRangeRef( int first_row, int first_col, int last_row, int last_col )
	{
		this.first_row = first_row;
		this.first_col = first_col;
		this.last_row = last_row;
		this.last_col = last_col;
	}

	/**
	 * return the number of cells in this rangeref
	 *
	 * @return number of cells in ref
	 */
	public int numCells()
	{
		int ret = -1;
		int numrows = this.last_row - this.first_row;
		numrows++;
		int numcols = this.last_col - this.first_col;
		numcols++;
		ret = numrows * numcols;
		return ret;
	}

	public CellRangeRef( int first_row, int first_col, int last_row, int last_col, String first_sheet, String last_sheet )
	{
		this( first_row, first_col, last_row, last_col );
		this.first_sheet_name = first_sheet;
		this.last_sheet_name = last_sheet;
	}

	public CellRangeRef( int first_row, int first_col, int last_row, int last_col, WorkSheetHandle first_sheet, WorkSheetHandle last_sheet )
	{
		this( first_row, first_col, last_row, last_col );
		this.first_sheet = first_sheet;
		this.last_sheet = last_sheet;
	}

	/**
	 * Parses a range in A1 notation and returns the equivalent CellRangeRef.
	 */
	public static CellRangeRef fromA1( String reference )
	{
		CellRangeRef ret = new CellRangeRef();
		String range;

		{
			String[] parts = ExcelTools.stripSheetNameFromRange( reference );
			ret.first_sheet_name = parts[0];
			range = parts[1];
			ret.last_sheet_name = parts[2];
		}

		if( range == null )
		{
			throw new IllegalArgumentException( "missing range component" );
		}

		{
			int[] parts = ExcelTools.getRangeRowCol( range );
			ret.first_row = parts[0];
			ret.first_col = parts[1];
			ret.last_row = parts[2];
			ret.last_col = parts[3];
		}

		return ret;
	}

	/**
	 * Convenience method combining {@link #fromA1(String)} and
	 * {@link #resolve(WorkBookHandle)}.
	 */
	public static CellRangeRef fromA1( String reference, WorkBookHandle book ) throws WorkSheetNotFoundException
	{
		CellRangeRef ret = fromA1( reference );
		ret.resolve( book );
		return ret;
	}

	/**
	 * Resolves sheet names into sheet handles against the given book.
	 *
	 * @param book the book against which the sheet names should be resolved
	 * @throws WorkSheetNotFoundException if either of the sheets does not
	 *                                    exist in the given book
	 */
	public void resolve( WorkBookHandle book ) throws WorkSheetNotFoundException
	{
		if( first_sheet_name != null )
		{
			first_sheet = book.getWorkSheet( first_sheet_name );
		}
		if( last_sheet_name != null )
		{
			last_sheet = book.getWorkSheet( last_sheet_name );
		}
	}

	/**
	 * Returns the lowest-indexed row in this range.
	 *
	 * @return the row index or null if this is a column range
	 */
	public int getFirstRow()
	{
		return first_row;
	}

	/**
	 * Returns the lowest-indexed column in this range.
	 *
	 * @return the column index or null if this is a row range
	 */
	public int getFirstColumn()
	{
		return first_col;
	}

	/**
	 * Returns the highest-indexed row in this range.
	 *
	 * @return the row index or null if this is a column range
	 */
	public int getLastRow()
	{
		return last_row;
	}

	/**
	 * Returns the highest-indexed column in this range.
	 *
	 * @return the column index or null if this is a row range
	 */
	public int getLastColumn()
	{
		return last_col;
	}

	/**
	 * Returns the name of the first sheet in this range.
	 *
	 * @return the name of the sheet or null if this range is not qualified
	 * with a sheet
	 */
	public String getFirstSheetName()
	{
		if( first_sheet != null )
		{
			return first_sheet.getSheetName();
		}
		else
		{
			return first_sheet_name;
		}
	}

	/**
	 * Returns the first sheet in this range.
	 *
	 * @return the WorkSheetHandle or null if this range is not qualified
	 * with a sheet or the sheet names have not been resolved
	 */
	public WorkSheetHandle getFirstSheet()
	{
		return first_sheet;
	}

	/**
	 * Returns the name of the last sheet in this range.
	 *
	 * @return the name of the sheet or null if this range is not qualified
	 * with a sheet
	 */
	public String getLastSheetName()
	{
		if( last_sheet != null )
		{
			return last_sheet.getSheetName();
		}
		else
		{
			return last_sheet_name;
		}
	}

	/**
	 * Returns the last sheet in this range.
	 *
	 * @return the WorkSheetHandle or null if this range is not qualified
	 * with a sheet or the sheet names have not been resolved
	 */
	public WorkSheetHandle getLastSheet()
	{
		return last_sheet;
	}

	/**
	 * Determines whether this range is qualified with a sheet.
	 */
	public boolean hasSheet()
	{
		return first_sheet != null || first_sheet_name != null;
	}

	/**
	 * Determines whether this range spans multiple sheets.
	 */
	public boolean isMultiSheet()
	{
		return (first_sheet != null && last_sheet != null && first_sheet != last_sheet) || (first_sheet_name != null && last_sheet_name != null && first_sheet_name != last_sheet_name);
	}

	/**
	 * Sets the first row in this range.
	 *
	 * @param value the row index to set
	 */
	public void setFirstRow( int value )
	{
		first_row = value;
	}

	/**
	 * Sets the first column in this range.
	 *
	 * @param value the column index to set
	 */
	public void setFirstColumn( int value )
	{
		first_col = value;
	}

	/**
	 * Sets the first sheet in this range.
	 */
	public void setFirstSheet( WorkSheetHandle sheet )
	{
		first_sheet_name = null;
		first_sheet = sheet;
	}

	/**
	 * Sets the last row in this range.
	 *
	 * @param value the row index to set
	 */
	public void setLastRow( int value )
	{
		last_row = value;
	}

	/**
	 * Sets the last column in this range.
	 *
	 * @param value the column index to set
	 */
	public void setLastColumn( int value )
	{
		last_col = value;
	}

	/**
	 * Sets the last sheet in this range.
	 */
	public void setLastSheet( WorkSheetHandle sheet )
	{
		last_sheet_name = null;
		last_sheet = sheet;
	}

	/**
	 * Returns whether this range entirely contains the given range.
	 * This ignores the sheets, if any, and compares only the cell ranges.
	 */
	public boolean contains( CellRangeRef range )
	{
		return this.first_row <= range.first_row && this.last_row >= range.last_row && this.first_col <= range.first_col && this.last_col >= range.last_col;
	}

	/**
	 * Compares this range to the specified object.
	 * The result is <code>true</code> if and only if the argument is not
	 * <code>null</code> and is a <code>CellRangeRef</code> object that
	 * represents the same range as this object.
	 */
	public boolean equals( Object other )
	{
		// if it's null or not a CellRangeRef it can't be equal
		if( other == null || !(other instanceof CellRangeRef) )
		{
			return false;
		}

		return this.toString().equals( other.toString() );
	}

	/**
	 * Creates and returns a copy of this range.
	 */
	@Override
	public Object clone()
	{
		try
		{
			return super.clone();
		}
		catch( CloneNotSupportedException e )
		{
			// This can't happen (we're Cloneable) but we have to catch it
			throw new Error( "Object.clone() threw CNSE but we're Cloneable" );
		}
	}

	/**
	 * Gets this range in A1 notation.
	 */
	public String toString()
	{
		String sheet1 = getFirstSheetName();
		String sheet2 = getLastSheetName();
		return (sheet1 != null ? sheet1 + (sheet2 != null && sheet2 != sheet1 ? ":" + sheet2 : "") + "!" : "") + ExcelTools.formatRange( new int[]{
				first_col, first_row, last_col, last_row
		} );
	}
}
