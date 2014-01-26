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
package org.openxls.formats.cellformat;

import org.openxls.ExtenXLS.Cell;
import org.openxls.ExtenXLS.ExcelTools;

import java.text.FieldPosition;
import java.text.Format;
import java.text.ParsePosition;

public class GeneralCellFormat extends Format implements CellFormat
{
	private static final long serialVersionUID = -3530672760714160988L;

	// make the constructor package-private
	GeneralCellFormat()
	{
	}

	@Override
	public StringBuffer format( Object obj, StringBuffer buffer, FieldPosition pos )
	{
		// try to parse strings as numbers
		if( obj instanceof String )
		{
			try
			{
				obj = Double.valueOf( (String) obj );
			}
			catch( NumberFormatException ex )
			{
				// this is OK, it just wasn't a number
			}
		}

		if( obj instanceof Number )
		{
			Number num = (Number) obj;
			if( num.longValue() == num.doubleValue() )
			{
				// it's an integer
				return buffer.append( String.valueOf( num.longValue() ) );
			}
			// it's floating-point
			return buffer.append( ExcelTools.getNumberAsString( num.doubleValue() ) );
		}

		return buffer.append( obj.toString() );
	}

	@Override
	public Object parseObject( String source, ParsePosition pos )
	{
		throw new UnsupportedOperationException();
	}

	@Override
	public String format( Cell cell )
	{
		return format( cell.getVal() );
	}

	public static String format( String val )
	{
		return format( val );
	}
}
