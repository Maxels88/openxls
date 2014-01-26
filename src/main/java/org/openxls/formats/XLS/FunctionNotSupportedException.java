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
package org.openxls.formats.XLS;

import org.openxls.formats.XLS.formulas.FunctionConstants;

import java.util.Locale;

/**
 * <b>Formula function is not supported for calculation.</b>
 *
 * @see Formula
 */

public final class FunctionNotSupportedException extends java.lang.RuntimeException
{

	private static final long serialVersionUID = 3569219212252117988L;
	String functionName = "";

	public FunctionNotSupportedException( String n )
	{
		super();
		functionName = n;
	}

	@Override
	public String getMessage()
	{
		// This method is derived from class java.lang.Throwable
		// to do: code goes here
		return toString();
	}

	public String toString()
	{
		// if it is a formula calc that is failing
//		if (functionName.length() > 4){
		if( true )
		{ // 20081203 KSC: functionName is offending function
			return "Function Not Supported: " + functionName;
		}
		int fID = 0;
		String f = "Unknown Formula";
		try
		{
			fID = Integer.parseInt( functionName, 16 );    // hex
			if( Locale.JAPAN.equals( Locale.getDefault() ) )
			{
				f = FunctionConstants.getJFunctionString( (short) fID );
			}
			if( f.equals( "Unknown Formula" ) )
			{
				f = FunctionConstants.getFunctionString( (short) fID );
			}
			if( f.length() == 0 )
			{
				if( fID == FunctionConstants.xlfADDIN )
				{
					f = "AddIn Formula";
				}
				else
				{
					f = "Unknown Formula";
				}
			}
			else
			{
				f += ")";    // add ending paren
			}
		}
		catch( Exception e )
		{
		}

		return "Function: " + f + " " + functionName + " is not implemented.";
	}
}