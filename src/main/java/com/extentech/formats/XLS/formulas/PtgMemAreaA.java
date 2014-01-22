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
package com.extentech.formats.XLS.formulas;

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.toolkit.ByteTools;

/**
 * PtgMemArea is an optimization of referenced areas.  Sweet!
 * <p/>
 * Like most optimizations it really sucks.  It is also one of the few Ptg's that
 * has a variable length.
 * <p/>
 * Format of length section
 * <pre>
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           (reserved)     4       Whatever it may be
 * 2           cce			   2	   length of the reference subexpression
 *
 * Format of reference Subexpression
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0			cref		2			The number of rectangles to follow
 * 2			rgref		var			An Array of rectangles
 *
 * Format of Rectangles
 * Offset      Name        Size    Contents
 * ----------------------------------------------------
 * 0           rwFirst     2       The First row of the reference
 * 2           rwLast     2       The Last row of the reference
 * 4           ColFirst    1       (see following table)
 * 6           ColLast    1       (see following table)
 * </pre>
 *
 * @see Ptg
 * @see Formula
 */
public class PtgMemAreaA extends PtgMemArea
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5528547215693511069L;
	int cce = 0;
	int cref = 0;
	MemArea[] areas;

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		this.populateVals();
	}

	@Override
	void populateVals()
	{
		cce = ByteTools.readInt( record[6], record[5] );
		cref = ByteTools.readInt( record[8], record[7] );
		areas = new MemArea[cref];
		int holder = 9;
		for( int i = 0; i < cref; i++ )
		{
			byte[] arr = new byte[6];
			System.arraycopy( record, holder, arr, 0, 6 );
			areas[i].init( arr );
			holder += 6;
		}

	}

	@Override
	public int getLength()
	{
		return -1;
	}

	/*
	 *  return a string representation of all of the ranges, seperated by comma.
	 */
	@Override
	public Object getValue()
	{
		String res = "";
		for( int i = 0; i < areas.length; i++ )
		{
			res += areas[i].getString();
			if( i != (areas.length - 1) )
			{
				res += ",";
			}
		}
		Object o = res;
		return o;
	}

	/*
	 * Describes a representation of an excel Reference.  This is not an actual
	 * reference to the cells, just the description!
	 *
	 */
	private class MemArea
	{
		int rwFirst;
		int rwLast;
		int colFirst;
		int colLast;

		void init( byte[] b )
		{
			rwFirst = ByteTools.readInt( b[0], b[1] );
			rwLast = ByteTools.readInt( b[2], b[3] );
			colFirst = (int) b[4];
			colLast = (int) b[5];
		}

		/*
		 * returns a string representation of the area, or cell if only one.
		 */
		String getString()
		{
			if( (rwFirst == rwLast) && (colFirst == colLast) )
			{
				// it is a single cell amoeba
				String retstr = ExcelTools.getAlphaVal( colLast );
				retstr = retstr + (rwLast + 1);
				return retstr;
			}
			int[] arr = { rwFirst, colFirst, rwLast, colLast };
			return ExcelTools.formatRangeRowCol( arr );
		}

	}

}