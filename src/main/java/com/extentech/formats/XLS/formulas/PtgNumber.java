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
    Ptg that stores an IEEE value
    
    Offset  Name       Size     Contents
    ------------------------------------
    0       num          8      An IEEE floating point nubmer
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgNumber extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -1650136303920724485L;

	@Override
	public boolean getIsOperand()
	{
		return true;
	}

	double val;
	boolean percentage = false;    // 20081208 KSC: so can handle percentage values in String formulas

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		if( !percentage )
		{
			return String.valueOf( val );
		}
		return String.valueOf( val * 100 ) + "%";
	}

	@Override
	public Object getValue()
	{
		Double d = new Double( val );
		return d;
	}

	public PtgNumber()
	{
		ptgId = 0x1F;
		val = 0;
		this.updateRecord();
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		this.populateVals();
	}

	/**
	 * Constructer to create these on the fly, this is needed
	 * for value storage in calculations of formulas.
	 */
	public PtgNumber( double d )
	{
		ptgId = 0x1F;
		val = d;
		this.updateRecord();
	}

	private void populateVals()
	{
		byte[] barr = new byte[8];
		System.arraycopy( record, 1, barr, 0, 8 );
		val = ByteTools.eightBytetoLEDouble( barr );
	}

	public double getVal()
	{
		return val;
	}

	/**
	 * override of GenericPtg.getDoubleVal();
	 */
	@Override
	public double getDoubleVal()
	{
		return val;
	}

	public void setVal( double i )
	{
		val = i;
		this.updateRecord();
	}

	// 20081208 KSC: handle percentage values
	public void setVal( String s )
	{
		s = s.trim();
		if( s.indexOf( "%" ) == s.length() - 1 )
		{
			percentage = true;
			s = s.substring( 0, s.indexOf( "%" ) );
			val = new Double( s ).doubleValue() / 100;
		}
		else
		{
			val = new Double( s ).doubleValue();
		}
	}

	@Override
	public void updateRecord()
	{
		byte[] tmp = new byte[1];
		tmp[0] = ptgId;
		byte[] brow = ByteTools.toBEByteArray( val );
		tmp = ByteTools.append( brow, tmp );
		record = tmp;
	}

	@Override
	public int getLength()
	{
		return PTG_NUM_LENGTH;
	}

	public String toString()
	{
		return getString();
	}

}