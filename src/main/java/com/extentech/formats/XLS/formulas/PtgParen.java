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
   Indicate the placing of operatands in parenthesis.
   
   Does not affect calculation. 
   
   =1+(2)
   
   ptgInt 1
   ptgInt 2
   ptgParen
   ptgAdd
   
   =(1+2)
   
   ptgInt 1
   ptgInt 2
   ptgAdd
   ptgParen   
   
    
 * @see Ptg
 * @see Formula

  .   
*/
public class PtgParen extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -8081397558698615537L;

	@Override
	public boolean getIsControl()
	{
		return true;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		// TODO: add logic to return proper paren 12/02 -jm

		return "(";
	}

	/**
	 * return the human-readable String representation of
	 * the "closing" portion of this Ptg
	 * such as a closing parenthesis.
	 */
	@Override
	public String getString2()
	{
		return ")";
	}

	/**
	 * Pass in the last 3 ptgs to evaluate
	 * where to place the String parens.
	 */
	@Override
	public Object evaluate( Object[] b )
	{
		return null;
	}

	@Override
	public int getLength()
	{
		return PTG_PAREN_LENGTH;
	}

	// KSC: added
	//default constructor
	public PtgParen()
	{
		ptgId = 0x15;
		record = new byte[1];
		record[0] = ptgId;
	}

	public String toString()
	{
		return ")";
	}
}