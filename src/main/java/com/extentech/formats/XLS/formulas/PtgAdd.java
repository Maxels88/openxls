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

import java.lang.reflect.Array;

/*
   Ptg that is an addition operand
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgAdd extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -964400139336259946L;

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
	}

	public PtgAdd()
	{
		ptgId = 0x3;
		record = new byte[1];
		record[0] = 0x3;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return "+";
	}

	public String toString()
	{
		return this.getString();
	}

	@Override
	public int getLength()
	{
		return PTG_ADD_LENGTH;
	}

	/*  Operator specific calculate method, this one adds two values.

	*/
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		try
		{
			Object[] o = super.getValuesFromPtgs( form );
			if( o == null )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			if( !o[0].getClass().isArray() )
			{
				if( o.length != 2 )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
				double o0 = 0, o1 = 0;
				try
				{
					o0 = getDoubleValue( o[0], this.parent_rec );
					o1 = getDoubleValue( o[1], this.parent_rec );
				}
				catch( NumberFormatException e )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
				double returnVal = o0 + o1;
				PtgNumber n = new PtgNumber( returnVal );
				return n;
			}
			else
			{    // handle array fomulas
				String retArry = "";
				int nArrays = java.lang.reflect.Array.getLength( o );
				if( nArrays != 2 )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
				int nVals = java.lang.reflect.Array.getLength( o[0] );    // use first array element to determine length of values as subsequent vals might not be arrays
				for( int i = 0; i < (nArrays - 1); i += 2 )
				{
					Object secondOp = null;
					boolean comparitorIsArray = o[i + 1].getClass().isArray();
					if( !comparitorIsArray )
					{
						secondOp = o[i + 1];
					}
					for( int j = 0; j < nVals; j++ )
					{
						Object firstOp = Array.get( o[i], j );    // first array index j
						if( comparitorIsArray )
						{
							secondOp = Array.get( o[i + 1], j );    // second array index j
						}
						double o0 = 0, o1 = 0;
						try
						{
							o0 = getDoubleValue( firstOp, this.parent_rec );
							o1 = getDoubleValue( secondOp, this.parent_rec );
						}
						catch( NumberFormatException e )
						{
							retArry = retArry + "#VALUE!" + ",";
							continue;
						}
						double retVal = o0 + o1;
						retArry = retArry + retVal + ",";
					}
				}
				retArry = "{" + retArry.substring( 0, retArry.length() - 1 ) + "}";
				PtgArray pa = new PtgArray();
				pa.setVal( retArry );
				return pa;
			}

		}
		catch( NumberFormatException e )
		{
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
		catch( Exception e )
		{    // handle error ala Excel
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
	}

}