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
package org.openxls.formats.XLS.formulas;

/*
PtgErr is exactly what one would think it is, a ptg that describes an Error
value.

Offset  Name    Size    Contents
-------------------------------------------
0       err     1       An error value

*/

public class PtgErr extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -5201871987022621869L;
	public static byte ERROR_NULL = 0x0;
	public static byte ERROR_DIV_ZERO = 0x7;
	public static byte ERROR_VALUE = 0xF;
	public static byte ERROR_REF = 0x17;
	public static byte ERROR_NAME = 0x1D;
	public static byte ERROR_NUM = 0x24;
	public static byte ERROR_NA = 0x2A;

	private boolean isCircularError = false;

	public boolean isCircularError()
	{
		return isCircularError;
	}

	public void setCircularError( boolean isCircularError )
	{
		this.isCircularError = isCircularError;
	}

	private String errorValue = null;

	@Override
	public boolean getIsOperand()
	{
		return true;
	}

	public PtgErr()
	{
		// default constructor
	}

	public PtgErr( byte errorV )
	{
		record = new byte[2];
		record[0] = 0x1C;
		record[1] = errorV;
	}

	public byte getErrorType()
	{
		return record[1];
	}

	@Override
	public Object getValue()
	{
		return toString();
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
	}

	public String toString()
	{
		if( isCircularError() )
		{
			return "#CIR_ERR!";
		}
		byte b = record[1];
		// duh, should have done a switch
		if( b == ERROR_NULL )
		{
			errorValue = "#ERROR!";
		}
		else if( b == ERROR_DIV_ZERO )
		{
			errorValue = "#DIV/0!";
		}
		else if( b == ERROR_VALUE )
		{
			errorValue = "#VALUE!";
		}
		else if( b == ERROR_REF )
		{
			errorValue = "#REF!";
		}
		else if( b == ERROR_NAME )
		{
			errorValue = "#NAME?";
		}
		else if( b == ERROR_NUM )
		{
			errorValue = "#NUM!";
		}
		else if( b == ERROR_NA )
		{
			errorValue = "#N/A";
		}
		return errorValue;
	}

	public static byte convertStringToLookupByte( String errorString )
	{
		if( errorString.equals( "#ERROR!" ) )
		{
			return ERROR_NULL;
		}
		if( errorString.equals( "#DIV/0!" ) )
		{
			return ERROR_DIV_ZERO;
		}
		if( errorString.equals( "#REF!" ) )
		{
			return ERROR_VALUE;
		}
		if( errorString.equals( "#ERROR!" ) )
		{
			return ERROR_REF;
		}
		if( errorString.equals( "#NAME?" ) )
		{
			return ERROR_NAME;
		}
		if( errorString.equals( "#NUM!" ) )
		{
			return ERROR_NUM;
		}
		if( errorString.equals( "#N/A" ) )
		{
			return ERROR_NA;
		}
		return ERROR_NULL;
	}

	@Override
	public int getLength()
	{
		return 2;
	}

}