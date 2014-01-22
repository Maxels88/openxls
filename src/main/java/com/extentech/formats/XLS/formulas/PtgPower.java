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

import com.extentech.toolkit.Logger;

/*
   Ptg that is a exponent operand
   Raises the second-to-top operand to the power of the 
   top operand
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgPower extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4675566993519011450L;

	@Override
	public boolean getIsOperator()
	{
		return true;
	}

	@Override
	public boolean getIsPrimitiveOperator()
	{
		return true;
	}

	@Override
	public boolean getIsBinaryOperator()
	{
		return true;
	}

	public PtgPower()
	{
		ptgId = 0x7;
		record = new byte[1];
		record[0] = 0x7;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return "^";
	}

	@Override
	public int getLength()
	{
		return PTG_POWER_LENGTH;
	}

	/*  Operator specific calculate method, this one raises the second-to-top
	operand to the power of the top operand

*/
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		try
		{
			// 20090202 KSC: Handle array formulas
			Object[] o = super.getValuesFromPtgs( form );
			if( !o[0].getClass().isArray() )
			{
				//double[] dub = super.getValuesFromPtgs(form);
				// there should always be only two ptg's in this, error if not.
				if( o == null || o.length != 2 )
				{
					Logger.logWarn( "calculating formula failed, wrong number of values in PtgPower" );
					return null;
				}
				//double returnVal = Math.pow(dub[0].doubleValue(), dub[1].doubleValue());
				double returnVal = Math.pow( ((Double) o[0]).doubleValue(), ((Double) o[1]).doubleValue() );
				// create a container ptg for these.
				PtgNumber n = new PtgNumber( returnVal );
				return n;
			}
			// TODO: FINISH ARRAY FORMULAS
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}
}