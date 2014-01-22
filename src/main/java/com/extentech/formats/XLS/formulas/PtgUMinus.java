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
/*
   Ptg that indicates Unary minus, negates the operand on top of the stack
    
 * @see Ptg
 * @see Formula

    
*/
package com.extentech.formats.XLS.formulas;

import com.extentech.toolkit.Logger;

/**
 *
 */
public class PtgUMinus extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8448419489380791823L;

	public PtgUMinus()
	{    // 20060504 KSC: Added to fill record bytes upon creation
		ptgId = 0x13;
		record = new byte[1];
		record[0] = 0x13;
	}

	public boolean getIsOperator()
	{
		return true;
	}

	public boolean getIsPrimitiveOperator()
	{
		return true;
	}

	public boolean getIsUnaryOperator()
	{
		return true;
	}

	/**
	 * return the human-readable String representation of
	 */
	public int getLength()
	{
		return PTG_UMINUS_LENGTH;
	}

	/**
	 * Operator specific calculate method, this one returns a single value sent to it.
	 */
	public Ptg calculatePtg( Ptg[] form )
	{
		if( form.length != 1 )
		{
			Logger.logWarn( "PtgMinus calculating formula failed, wrong number of values." );
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		try
		{
			Ptg p = form[0];
			Ptg ret = null;
			if( p instanceof PtgInt )
			{
				int val = p.getIntVal();
				val *= -1;
				ret = new PtgInt( val );
			}
			else
			{
				double val = p.getDoubleVal();
				val *= -1;
				ret = new PtgNumber( val );
			}
			return ret;
		}
		catch( Exception e )
		{
			Logger.logWarn( "PtgMinus calculating formula failed, could not negate operand " + form[0].toString() + " : " + e.toString() );
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}

	public String getString()
	{
		return "-";
	}

	public String toString()
	{
		return "u-";
	}

}
