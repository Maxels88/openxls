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

import com.extentech.ExtenXLS.DateConverter;
import com.extentech.formats.XLS.BiffRec;

public class Calculator
{
	/**
	 * given a BiffRec Cell Record, an Object and an operator to compare to,
	 * return true if the comparison passes, false otherwise
	 *
	 * @param c  BiffRec cell record
	 * @param o  Object value - one of Double, String or Boolean
	 * @param op - String operator - one of "=", ">", ">=", "<", "<=" or "<>
	 * @return true if comparison of operator and value with cell value passes, false otherwise
	 */
	public static boolean compareCellValue( BiffRec c, Object o, String op )
	{
		// doper types:  numeric:  ieee, rk
		//               string:   string doper
		// 				 boolean
		//				 error
		int compare;
		try
		{
			if( o instanceof Boolean )        // TODO: 1.5 use Boolean.compareTo
			{
				compare = ((Boolean) o).toString().compareTo( Boolean.valueOf( c.getBooleanVal() ).toString() );
			}
			else if( o instanceof String )
			{
				// use "matches" to handle wildcards
				if( ((String) o).toUpperCase().matches( c.getStringVal().toUpperCase() ) )
				{
					compare = 0;    // equal or matches
				}
				else
				{
					compare = -1;    // doesn't equal
				}
				//compare= ((String) o).toUpperCase().compareTo(c.getStringVal().toUpperCase());
			}
			else // it's a Double
			{
				compare = ((Double) o).compareTo( c.getDblVal() );
			}
		}
		catch( Exception e )
		{
			; // report error?
			return false;
		}
		if( op.equals( "=" ) )
		{
			return (compare == 0);
		}
		if( op.equals( "<" ) )
		{
			return (compare > 0);
		}
		if( op.equals( "<=" ) )
		{
			return (compare >= 0);
		}
		if( op.equals( ">" ) )
		{
			return (compare < 0);
		}
		//noinspection ConfusingElseBranch
		if( op.equals( ">=" ) )
		{
			return (compare <= 0);
		}
		if( op.equals( "<>" ) )
		{
			return (compare != 0);
		}
		return false;
	}

	public static boolean compareCellValue( Object val, String compareval, String op )
	{
		// doper types:  numeric:  ieee, rk
		//               string:   string doper
		// 				 boolean
		//				 error
		int compare = -1;
		try
		{
			if( val instanceof Boolean )    // TODO: 1.5 use Boolean.compareTo
			{
				compare = (Boolean.valueOf( compareval ).toString()).compareTo( ((Boolean) val).toString() );
			}
			else if( val instanceof String )
			{
				if( (compareval.indexOf( '?' ) == -1) && (compareval.indexOf( '*' ) == -1) )    // if no wildcards
				{
					compare = (compareval).compareTo( ((String) val).toUpperCase() );
				}
				else
				{    // use "matches" to handle wildcards
					if( ((String) val).toUpperCase().matches( compareval ) )
					{
						compare = 0;    // equal or matches
					}
					else
					{
						compare = -1;    // doesn't equal
					}
				}
			}
			else if( val instanceof Number )    // assume it's a number
			{
				compare = (new Double( compareval )).compareTo( ((Number) val).doubleValue() );
			}
			else
			{
				return false;
			}
		}
		catch( Exception e )
		{
			try
			{    // try date compare
				double dt = DateConverter.getXLSDateVal( new java.util.Date( compareval ) );
				compare = (new Double( dt )).compareTo( ((Number) val).doubleValue() );
			}
			catch( Exception ex )
			{    // just try string compare
				compare = compareval.compareTo( val.toString() );
			}
		}
		if( op.equals( "=" ) )
		{
			return (compare == 0);
		}
		if( op.equals( "<" ) )
		{
			return (compare > 0);
		}
		if( op.equals( "<=" ) )
		{
			return (compare >= 0);
		}
		if( op.equals( ">" ) )
		{
			return (compare < 0);
		}
		//noinspection ConfusingElseBranch
		if( op.equals( ">=" ) )
		{
			return (compare <= 0);
		}
		if( op.equals( "<>" ) )
		{
			return (compare != 0);
		}
		return false;
	}

/*    } else if (op.equals("<")) {
    	passes= (compare < 0);
	} else if (op.equals("<=")) {
		passes= (compare <= 0);
	} else if (op.equals(">")) {
		passes= (compare > 0);
	} else if (op.equals(">=")) {
		passes= (compare >= 0);
	}
  */

	/**
	 * translate Excel-style wildcards into Java wildcards in criteria string
	 * plus handle qualified wildcard characters + percentages ...
	 *
	 * @param sCriteria criteria string
	 * @return tranformed criteria string
	 */
	public static String translateWildcardsInCriteria( String sCriteria )
	{
		String s = "";    // handle wildcards
		boolean qualified = false;
		boolean isalldigits = true;
		for( int i = 0; i < sCriteria.length(); i++ )
		{
			char c = sCriteria.charAt( i );
			if( c == '~' )
			{
				qualified = true;    // don't add tilde unless certain it's not qualifying a * or ?
			}
			else if( c == '*' )
			{
				if( !qualified )
				{
					s += ".";
				}
				s += c;
			}
			else if( c == '?' )
			{
				if( !qualified )
				{
					s += ".";
				}
				s += c;
			}
			else if( c == '%' )
			{ // translate percentage into decimals
				if( isalldigits )
				{
					s = "0" + s;
					s = s.substring( s.length() - 2, 2 );
					s = "." + s;
				}
			}
			else
			{
				if( qualified ) // really add the tilde
				{
					s += '~';
				}
				s += c;
				qualified = false;
				if( !Character.isDigit( c ) )
				{
					isalldigits = false;
				}
			}
		}
		sCriteria = s.toUpperCase();    // matching is case-insensitive
		return sCriteria;
	}

	/**
	 * given a criteria string that starts with an operator,
	 * parse and return the index that the operator ends and the crtieria starts
	 *
	 * @param criteria
	 * @return int i	position in criteria which actual criteria starts
	 */
	public static int splitOperator( String criteria )
	{
		int i = 0;
		for(; i < criteria.length(); i++ )
		{
			char c = criteria.charAt( i );
			if( Character.isJavaIdentifierPart( c ) )
			{
				break;
			}
			if( (c == '*') || (c == '?') )
			{
				break;
			}
		}
		return i;
	}

	/**
	 * takes a Reference Type Ptg and deferences and PtgNames, etc.
	 * to return a PtgArea
	 *
	 * @param p
	 * @return
	 */
	public static PtgArea getRange( Ptg p ) throws IllegalArgumentException
	{
		if( p instanceof PtgArea )
		{
			return (PtgArea) p;
		}
		if( p instanceof PtgName )
		{    // get source range
			Ptg[] pr = null;
			try
			{
				pr = ((PtgName) p).getName().getCellRangePtgs();
				return (PtgArea) pr[0];
			}
			catch( Exception e )
			{
				try
				{    // if it's a PtgRef, convert to a PtgArea
					if( !(pr[0] instanceof PtgArea) && (pr[0] instanceof PtgRef) )
					{
						PtgArea pa = new PtgArea();
						pa.setParentRec( pr[0].getParentRec() );
						pa.setLocation( pr[0].getLocation() );
						return pa;
					}
					throw new IllegalArgumentException( "Expected a reference-type operand" );
				}
				catch( Exception ex )
				{
					throw new IllegalArgumentException( "Expected a reference-type operand" );
				}
			}
		}
		return null;
	}
}
