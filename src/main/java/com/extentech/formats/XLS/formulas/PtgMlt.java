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

import java.lang.reflect.Array;

/*
   Ptg that is a multiplier operand
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgMlt extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2670754297349356254L;

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

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return "*";
	}

	public String toString()
	{
		return this.getString();
	}

	public PtgMlt()
	{
		ptgId = 0x5;
		record = new byte[1];
		record[0] = 0x5;
	}

	@Override
	public int getLength()
	{
		return PTG_MLT_LENGTH;
	}

	/*  Operator specific calculate method, this one multiplies two values.

	*/
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		// handle ref errs
		if( (form[0] instanceof PtgErr) || (form[0] instanceof PtgRefErr) || (form[0] instanceof PtgRefErr3d) || (form[0] instanceof PtgAreaErr3d) )
		{
			return form[0];
		}
		if( (form[1] instanceof PtgErr) || (form[1] instanceof PtgRefErr) || (form[1] instanceof PtgRefErr3d) || (form[1] instanceof PtgAreaErr3d) )
		{
			return form[1];
		}

		try
		{
			// 20090202 KSC: Handle array formulas
			Object[] o = super.getValuesFromPtgs( form );
			if( o == null )
			{
				return new PtgErr( PtgErr.ERROR_VALUE ); // some error in value(s)
			}
			if( !o[0].getClass().isArray() )
			{
				//double[] dub = super.getValuesFromPtgs(form);
				// there should always be only two ptg's in this, error if not.
				if( o.length != 2 )
				{
					Logger.logWarn( "calculating formula failed, wrong number of values in PtgMlt" );
					return new PtgErr( PtgErr.ERROR_VALUE );    // 20081203 KSC: handle error's ala Excel return null;
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
				double returnVal = o0 * o1;
				// create a container ptg for these.
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
						double retVal = o0 * o1;
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
		{    // 20081125 KSC: handle error ala Excel
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
	}

}