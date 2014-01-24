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
   Ptg that is a Concatenation Operand.  
   Appends the top operand to the second-to-top Operand
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgConcat extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6671404163121438253L;

	@Override
	public boolean getIsOperator()
	{
		return true;
	}

	@Override
	public boolean getIsBinaryOperator()
	{
		return true;
	}

	@Override
	public boolean getIsPrimitiveOperator()
	{
		return true;
	}    // 20060512 KSC: added

	public PtgConcat()
	{
		ptgId = 0x8;
		record = new byte[1];
		record[0] = 0x8;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
//        return "CONCAT(";	// 20060512 KSC: mod 
		return "&";
	}

	@Override
	public String getString2()
	{
//        return ")";
		return "";
	}

	@Override
	public int getLength()
	{
		return PTG_CONCAT_LENGTH;
	}

	public String toString()
	{    // KSC added
		return getString();
	}

	/**
	 * Operator specific calculate method, this Concatenates two values
	 */
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		try
		{
			Object[] o = super.getStringValuesFromPtgs( form );
			// there should always be only two ptg's in this, error if not.
			if( (o == null) || (o.length != 2) )
			{
				//if (o!=null)
				//	Logger.logWarn("calculating formula failed, wrong number of values in PtgConcat");
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			if( !o[0].getClass().isArray() )
			{
				String[] s = new String[2];
				try
				{    // 20090216 KSC: try to convert numbers to ints when converting to string as otherwise all numbers come out as x.0
					s[0] = String.valueOf( ((Double) o[0]).intValue() );
				}
				catch( Exception e )
				{
					s[0] = o[0].toString();
				}
				try
				{    // 20090216 KSC: try to convert numbers to ints when converting to string as otherwise all numbers come out as x.0
					s[1] = String.valueOf( ((Double) o[1]).intValue() );
				}
				catch( Exception e )
				{
					s[1] = o[1].toString();
				}

				String returnVal = s[0] + s[1];
				PtgStr pstr = new PtgStr( returnVal );
				return pstr;
			}
			return null;
		}
		catch( Exception e )
		{    // handle error ala Excel
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}

	}
}