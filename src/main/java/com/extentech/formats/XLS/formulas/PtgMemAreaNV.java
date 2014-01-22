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
    Currently just a placeholder class
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgMemAreaNV extends PtgMemAreaA implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7912318247202256986L;

	/**
	 * return the human-readable String representation of
	 * <p/>
	 * public String getString(){
	 * return "MEMAREANV";
	 * }
	 */

	@Override
	public int getLength()
	{
		return PTG_MEM_AREA_NV_LENGTH;
	}

}