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
   Ptg that indicates Unary plus, has no effect on operand.  Very useful PTG...lol
    
 * @see Ptg
 * @see Formula

    
*/
package com.extentech.formats.XLS.formulas;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 *
 */
public class PtgUPlus extends GenericPtg implements Ptg
{
	private static final Logger log = LoggerFactory.getLogger( PtgUPlus.class );
	private static final long serialVersionUID = -3514760881731524419L;

	public PtgUPlus()
	{    // 20060504 KSC: Added to fill record bytes upon creation
		ptgId = 0x12;
		record = new byte[1];
		record[0] = 0x12;
	}

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
		return "+";
	}

	@Override
	public int getLength()
	{
		return PTG_UPLUS_LENGTH;
	}

	/*  Operator specific calculate method, this one returns a single value sent to it.

   */
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		// there should always be only one ptg in this, error if not.
		if( form.length != 1 )
		{
			log.warn( "calculating formula failed, wrong number of values in PtgUPlus" );
			return null;
		}
		return form[0];
	}

}
