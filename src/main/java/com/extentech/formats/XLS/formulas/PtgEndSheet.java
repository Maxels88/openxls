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
       
    Undocumented PTG, the only thing we know is "ptg DELETED" from MEFF pg 446.  I am going to treat
    as a PtgRefErr for now.

 * @see Ptg
 * @see Formula

*/
public class PtgEndSheet extends GenericPtg implements Ptg
{
	// Excel can handle PtgRefErrors within formulas, as long as they are not the result so...

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2395432053363123361L;

	public boolean getIsOperand()
	{
		return true;
	}

	/**
	 * return the human-readable String representation of
	 */
	public String getString()
	{
		return "End Sheet Error";
	}

	public Object getValue()
	{
		return "#REF!";
	}

	public int getLength()
	{
		return PTG_ENDSHEET_LENGTH;
	}

}