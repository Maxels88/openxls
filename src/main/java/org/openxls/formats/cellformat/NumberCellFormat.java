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

import java.text.FieldPosition;
import java.text.NumberFormat;
import java.text.ParsePosition;

public class NumberCellFormat extends NumberFormat implements CellFormat
{
	private static final long serialVersionUID = -7191923168789058338L;

	private final String positive;
	private final String negative;
	private final String zero;
	private final String string;

	NumberCellFormat( String positive, String negative, String zero, String string )
	{
		this.positive = positive;
		this.negative = negative;
		this.zero = zero;
		this.string = string;
	}

	@Override
	public StringBuffer format( Object input, StringBuffer buffer, FieldPosition pos )
	{
		if( input instanceof String )
		{
			// hack to make useless @ pattern work
			if( "%s".equals( positive ) )
			{
				return buffer.append( String.valueOf( input ) );
			}
			try
			{
				Double d = new Double( input.toString() );
				input = d;
			}
			catch( NumberFormatException e )
			{
				return buffer.append( String.format( string, (String) input ) );
			}
		}
		if( input instanceof Number )
		{
			String format;
			double value = ((Number) input).doubleValue();

			if( value > 0 )
			{
				format = positive;
			}
			else if( value < 0 )
			{
				format = negative;
				value = Math.abs( value );
			}
			else
			{
				format = zero;
			}

			// hack to make percentage formats work
			if( format.contains( "%%" ) )
			{
				value *= 100;
			}

			// hack to make useless @ pattern work
			if( "%s".equals( format ) )
			{
				return buffer.append( String.valueOf( input ) );
			}

			return buffer.append( String.format( format, value ) );
		}
		throw new IllegalArgumentException( "unsupported input type" );
	}

	@Override
	public StringBuffer format( double number, StringBuffer buffer, FieldPosition pos )
	{
		return buffer.append( format( number ) );
	}

	@Override
	public StringBuffer format( long number, StringBuffer buffer, FieldPosition pos )
	{
		return buffer.append( format( number ) );
	}

	@Override
	public Number parse( String source, ParsePosition parsePosition )
	{
		throw new UnsupportedOperationException();
	}

	@Override
	public String format( Cell cell )
	{
		return format( cell.getVal() );
	}

}
