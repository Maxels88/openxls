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



/*
   Ptg that is an missing operand
    
 * @see Ptg
 * @see Formula

    
*/

public class PtgMissArg extends GenericPtg implements Ptg
{

	private static final long serialVersionUID = 8995314621921283625L;

	public boolean getIsOperand()
	{
		return true;
	}

	public boolean getIsOperator()
	{
		return false;
	}

	public boolean getIsBinaryOperator()
	{
		return false;
	}

	public boolean getIsPrimitiveOperator()
	{
		return false;
	}

	public PtgMissArg()
	{
		this.init( new byte[]{ 22 } );
	}

	/**
	 * return the human-readable String representation of
	 */
	public String getString()
	{
		return "";
	}

	public String toString()
	{
		return "";
	}

	public int getLength()
	{
		return 1;
	}

	public Object getValue()
	{
		return null;
	}
}