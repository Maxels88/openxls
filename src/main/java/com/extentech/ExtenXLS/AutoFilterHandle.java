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

import com.extentech.formats.XLS.AutoFilter;

/**
 * AutoFilterHandle allows for manipulation of the AutoFilters within the Spreadsheet
 * <p/>
 * AutoFilters allow for
 */
public class AutoFilterHandle implements Handle
{
	private AutoFilter af = null;

	/**
	 * For internal use only.  Creates an AutoFilter Handle based on the AutoFilter passed in
	 *
	 * @param AutoFilter af
	 */
	protected AutoFilterHandle( AutoFilter af )
	{
		this.af = af;
	}

	/**
	 * returns the string representation of this AutoFilter
	 *
	 * @return the string representation of this AutoFilter
	 */
	public String toString()
	{
		if( af != null )
		{
			return af.toString();
		}
		return "No AutoFilter";
	}

	/**
	 * returns the column this AutoFilter is applied to
	 * <br>NOTE: this may not be 100% exact
	 *
	 * @return in column number
	 */
	public int getCol()
	{
		if( af != null )
		{
			return af.getCol();
		}
		return -1;
	}

	/**
	 * Sets the custom comparison of this AutoFilter via a String operator and an Object value
	 * <br>Only those rows that meet the equation: (column value) OP value will be shown
	 * <br>e.g show all rows where column value >= 2.99
	 * <p>Object value can be of type
	 * <p>String
	 * <br>Boolean
	 * <br>Error
	 * <br>a Number type object
	 * <p>String operator may be one of: "=", ">", ">=", "<>", "<", "<="
	 *
	 * @param Object val - value to set
	 * @param String op - operator
	 * @see setVal2
	 */
	public void setVal( Object val, String op )
	{
		if( af != null )
		{
			af.setVal( val, op );
		}
	}

	/**
	 * Sets the custom comparison of the second condition of this AutoFilter via a String operator and an Object value
	 * <br>This method sets the second condition of a two-condition filter
	 * <p>Only those rows  that meet the equation:
	 * <br> first condition AND/OR (column value) OP value will be shown
	 * <br>e.g show all rows where (column value) <= 1.99 AND (column value) >= 2.99
	 * <p>Object value can be of type
	 * <p>String
	 * <br>Boolean
	 * <br>Error
	 * <br>a Number type object
	 * <p>String operator may be one of: "=", ">", ">=", "<>", "<", "<="
	 *
	 * @param Object  val - value to set
	 * @param String  op - operator
	 * @param boolean AND - true if two conditions should be AND'ed, false if OR'd
	 * @see setVal2
	 */
	public void setVal2( Object val, String op, boolean AND )
	{
		if( af != null )
		{
			af.setVal2( val, op, AND );
		}
	}

	/**
	 * returns the String representation of the comparison value for this AutoFilter, if any
	 * <p/>
	 * <br>This will return the comparison value of the first condition for those AutoFilters containing two conditions
	 *
	 * @return String comparison value for the second condition or null if none exists
	 * @see getVal2
	 */
	public String getVal()
	{
		if( af != null )
		{
			return (String) af.getVal();
		}
		return null;
	}

	/**
	 * returns the String representation of the second comparison value for this AutoFilter, if any
	 * <p/>
	 * <br>This will return the comparison value of the second condition for those AutoFilters containing two conditions
	 *
	 * @return String comparison value for the second condition or null if none exists
	 * @see getVal
	 */
	public String getVal2()
	{
		if( af != null )
		{
			return (String) af.getVal2();
		}
		return null;
	}

	/**
	 * get the operator associated with this AutoFilter
	 * <br>NOTE: this will return the operator in the first condition if this AutoFilter contains two conditions
	 * <br>Use getOp2 to retrieve the second condition operator
	 *
	 * @return String operator
	 */
	public String getOp()
	{
		if( af != null )
		{
			return af.getOp();
		}
		return null;
	}

	/**
	 * Sets this AutoFilter to be a Top-n or Bottom-n type of filter
	 * <br>Top-n filters only show the Top n values or percent in the column
	 * <br>Bottom-n filters only show the bottom n values or percent in the column
	 * <br>n can be from 1-500, or 0 to turn off Top 10 filtering
	 *
	 * @param int     n - 0-500
	 * @param boolean percent - true if show Top-n percent; false to show Top-n items
	 * @param boolean top10 - true if show Top-n (items or percent), false to show Bottom-n (items or percent)
	 */
	public void setTop10( int n, boolean percent, boolean top10 )
	{
		if( af != null )
		{
			af.setTop10( n, percent, top10 );
		}
	}

	/**
	 * returns true if this AutoFilter is set to Top-10
	 * <br>Top-n filters only show the Top n values or percent in the column
	 *
	 * @return
	 */
	public boolean isTop10()
	{
		if( af != null )
		{
			return af.isTop10();
		}
		return false;
	}

	/**
	 * sets this AutoFilter to filter all blank rows
	 */
	public void setFilterBlanks()
	{
		if( af != null )
		{
			af.setFilterBlanks();
		}
	}

	/**
	 * sets this AutoFilter to filter all non-blank rows
	 */
	public void setFilterNonBlanks()
	{
		if( af != null )
		{
			af.setFilterNonBlanks();
		}
	}

	/**
	 * returns true if this AutoFitler is set to filter all blank rows
	 *
	 * @return true if filter blanks, false otherwise
	 */
	public boolean isFilterBlanks()
	{
		if( af != null )
		{
			return af.isFilterBlanks();
		}
		return false;
	}

	/**
	 * returns true if this AutoFitler is set to filter all non-blank rows
	 *
	 * @return true if filter non-blanks, false otherwise
	 */
	public boolean isFilterNonBlanks()
	{
		if( af != null )
		{
			return af.isFilterNonBlanks();
		}
		return false;
	}
}
