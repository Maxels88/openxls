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

import com.extentech.formats.XLS.FormulaNotFoundException;

import java.util.Comparator;

/**
 * CellComparitor is a implementation of a comparitor for
 * sorting cell values.  In this instance numeric values (whether from a cell or a formula result)
 * are sorted before any string values.
 * <p/>
 * Date values are sorted according to their internal date representation.  Note that currently this means
 * Dates will always sort above strings due to them storing their value as a long.
 */
public class CellComparator implements Comparator<CellHandle>
{

	/**
	 * Compare 2 cellHandle classes
	 * <p/>
	 * This method handles comparisons of cells, note that formula results are used rather
	 * than formula strings, Numbers are sorted ahead of string values.  Dates are stored
	 * as numbers internally in excel so are sorted against numbers
	 */
	@Override
	public int compare( CellHandle cell1, CellHandle cell2 )
	{
		int cellType1 = cell1.getCellType();
		int cellType2 = cell1.getCellType();

		// numerics - possibly break out more to see if values are floating point or not. 
		// would help for equality, which is likely never reached here.
		if( cell1.isNumber() && cell2.isNumber() )
		{
			if( cell1.getDoubleVal() > cell2.getDoubleVal() )
			{
				return 1;
			}
			if( cell1.getDoubleVal() < cell2.getDoubleVal() )
			{
				return -1;
			}
			return 0;
		}
		if( cell1.isNumber() )
		{// get formula value if exists and is a numeric value
			if( cellType2 == CellHandle.TYPE_FORMULA )
			{
				FormulaHandle f = null;
				try
				{
					f = cell2.getFormulaHandle();
				}
				catch( FormulaNotFoundException e )
				{
				}
				Double d = f.getDoubleVal();
				if( !(d == Double.NaN) )
				{
					if( cell1.getDoubleVal() > d )
					{
						return 1;
					}
					if( cell1.getDoubleVal() < d )
					{
						return -1;
					}
					return 0;
				}
			}
			else
			{
				return 1;
			}
		}
		else if( cell2.isNumber() )
		{
			if( cellType1 == CellHandle.TYPE_FORMULA )
			{
				FormulaHandle f = null;
				try
				{
					f = cell1.getFormulaHandle();
				}
				catch( FormulaNotFoundException e )
				{
				}
				Double d = f.getDoubleVal();
				if( !(d == Double.NaN) )
				{
					if( cell2.getDoubleVal() < d )
					{
						return 1;
					}
					if( cell2.getDoubleVal() > d )
					{
						return -1;
					}
					return 0;
				}
			}
			else
			{
				return -1;
			}
		}

		//Two formulas;
		if( (cellType1 == CellHandle.TYPE_FORMULA) && (cellType2 == CellHandle.TYPE_FORMULA) )
		{
			try
			{
				FormulaHandle f1 = cell1.getFormulaHandle();
				FormulaHandle f2 = cell2.getFormulaHandle();
				double d1 = f1.getDoubleVal();
				double d2 = f2.getDoubleVal();
				if( !(d1 == Double.NaN) && !(d2 == Double.NaN) )
				{
					if( d1 > d2 )
					{
						return 1;
					}
					if( d1 < d2 )
					{
						return -1;
					}
					return 0;
				}
				if( !(d1 == Double.NaN) )
				{
					return 1;
				}
				if( !(d2 == Double.NaN) )
				{
					return -1;
				}
			}
			catch( FormulaNotFoundException e )
			{
			}
		}

		// Strings, the last choice
		String val1 = cell1.getStringVal();
		String val2 = cell2.getStringVal();
		return val1.compareTo( val2 );
	}
}
