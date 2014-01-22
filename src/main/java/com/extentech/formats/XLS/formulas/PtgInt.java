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

import com.extentech.toolkit.ByteTools;

/*
    Ptg that stores an integer value
    
    Offset  Name       Size     Contents
    ------------------------------------
    0       w           2       An unsigned integer value
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgInt extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2129624418815329359L;

	@Override
	public boolean getIsOperand()
	{
		return true;
	}

	int val;

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return String.valueOf( val );
	}

	public PtgInt()
	{
	}

	/*
	 * constructer to create ptgint's on the fly, from formulas
	 */
	public PtgInt( int i )
	{
		val = i;
		this.updateRecord();
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		this.populateVals();
	}

	// 0 to 65535 - outside of these bounds must be a PtgNumber
	private void populateVals()
	{
		byte b = 0;
		int s = ByteTools.readInt( record[1], record[2], b, b );
		val = s;
	}

	public int getVal()
	{
		return val;
	}

	@Override
	public int getIntVal()
	{
		return val;
	}

	@Override
	public Object getValue()
	{
		Integer i = val;
		return i;
	}

	public void setVal( int i )
	{
		val = i;
		this.updateRecord();
	}

	public boolean getBooleanVal()
	{
		if( val == 1 )
		{
			return true;
		}
		return false;
	}

	@Override
	public void updateRecord()
	{
		byte[] tmp = new byte[1];
		tmp[0] = PTG_INT;
		byte[] brow = ByteTools.shortToLEBytes( (short) val );
		tmp = ByteTools.append( brow, tmp );
		record = tmp;
	}

	@Override
	public int getLength()
	{
		return PTG_INT_LENGTH;
	}

	public String toString()
	{
		return String.valueOf( this.getVal() );
	}

}