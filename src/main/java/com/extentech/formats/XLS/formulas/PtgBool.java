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

import com.extentech.formats.XLS.ExpressionParser;

/*
    A parse thing that represents a boolean value.  This is made up of two bytes,
    the PtgID (0x1D) and a byte representing the boolean value (0 or 1);
*/

public class PtgBool extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7270271326251770439L;

	@Override
	public boolean getIsOperator()
	{
		return false;
	}

	@Override
	public boolean getIsOperand()
	{
		return true;
	}

	boolean val = false;

	/**
	 * return the human-readable String representation of the ptg
	 */
	@Override
	public String getString()
	{
		return String.valueOf( val );
	}

	public String toString()
	{
		return String.valueOf( val );
	}

	@Override
	public Object getValue()
	{
		Boolean b = val;
		return b;
	}

	public void setVal( boolean boo )
	{
		val = boo;
		this.updateRecord();
	}

	@Override
	public void init( byte[] rec )
	{
		this.record = rec;
		ptgId = rec[0];
		if( rec[1] == 0 )
		{
			val = false;
		}
		else
		{
			val = true;
		}
	}

	public PtgBool()
	{
		record = new byte[2];
		ptgId = ExpressionParser.ptgBool;
		record[0] = ptgId;
	}

	public PtgBool( boolean b )
	{
		ptgId = ExpressionParser.ptgBool;
		val = b;
		this.updateRecord();
	}

	public boolean getBooleanValue()
	{
		return val;
	}

	@Override
	public void updateRecord()
	{
		record = new byte[2];
		record[0] = ptgId;
		if( val )
		{
			record[1] = 1;
		}
		else
		{
			record[1] = 0;
		}
	}

	@Override
	public int getLength()
	{
		return PTG_BOOL_LENGTH;
	}

}