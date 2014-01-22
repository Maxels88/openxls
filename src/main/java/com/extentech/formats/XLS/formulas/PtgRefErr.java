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
       
    Specifies a cell reference that was changed to #REF! due to worksheet editing

 * @see Ptg
 * @see Formula

*/
public class PtgRefErr extends PtgRef implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2553420345077869256L;

	public boolean getIsRefErr()
	{
		return true;
	}

	public void init( byte[] b )
	{
		record = b;
	}

	// Excel can handle PtgRefErrors within formulas, as long as they are not the result so...
	public boolean getIsOperand()
	{
		return true;
	}

	/**
	 * return the human-readable String representation of
	 */
	public String getString()
	{
		return "#REF!";    //Invalid Reference Error";
	}

	public Object getValue()
	{
		return "#REF!";
	}

	public int getLength()
	{
		return PTG_REFERR_LENGTH;
	}

	public int[] getRowCol()
	{
		return new int[]{ -1, -1 };
	}

	public String getLocation()
	{
		return "#REF!";
	}

	public void setLocation( String[] s )
	{
	}

}