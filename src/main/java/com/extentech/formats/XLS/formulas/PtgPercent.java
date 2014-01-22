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
   Ptg that indicates percentage, divides top operand by 100
    
 * @see Ptg
 * @see Formula

    
*/
package com.extentech.formats.XLS.formulas;

import com.extentech.toolkit.Logger;

/**
 *
 */
public class PtgPercent extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -8559541841405018157L;

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
	public boolean getIsUnaryOperator()
	{
		return true;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return "%";
	}

	@Override
	public int getLength()
	{
		return PTG_PERCENT_LENGTH;
	}

	/*  Operator specific calculate method, this one returns a single value sent to it.

   */
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		// 20090202 KSC: Handle array formulas
		Object[] o = super.getValuesFromPtgs( form );
		if( !o[0].getClass().isArray() )
		{
			//double[] dub = super.getValuesFromPtgs(form);
			// there should always be only two ptg's in this, error if not.
			if( (o == null) || (o.length != 1) )
			{
				// there should always be only one ptg in this, error if not.
				//if (form.length != 1){
				Logger.logWarn( "calculating formula failed, wrong number of values in PtgPercent" );
				return null;
			}
		}
		// TODO: finish for Array formulas
		double res = (((Double) o[0]).doubleValue()) / 100;
		return new PtgNumber( res );
	}

}
