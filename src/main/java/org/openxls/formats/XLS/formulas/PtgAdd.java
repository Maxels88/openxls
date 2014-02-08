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
package org.openxls.formats.XLS.formulas;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Array;

/*
   Ptg that is an addition operand
    
 * @see Ptg
 * @see Formula

    
*/
public class PtgAdd extends GenericPtg implements Ptg
{
	private static final Logger log = LoggerFactory.getLogger( PtgAdd.class );
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
		return getString();
	}

	@Override
	public int getLength()
	{
		return PTG_ADD_LENGTH;
	}

	/**
	 * Operator specific calculate method, this one adds two values.
	 * @param form
	 * @return
	 */
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		try
		{
			Object[] args = getValuesFromPtgs( form );
			if( args == null )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}

			int nArrays = java.lang.reflect.Array.getLength( args );
			if( nArrays != 2 )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			nArrays = 2;

			// FIXME THIS NEEDS TO CHECK IF EITHER ARG IS AN ARRAY, NOT JUST THE FIRST ONE! IT NEEDS TO BE ASSOCIATIVE...

			if( !args[0].getClass().isArray() && !args[1].getClass().isArray() )
			{
				double o0;
				double o1;
				try
				{
					o0 = getDoubleValue( args[0], parent_rec );
					o1 = getDoubleValue( args[1], parent_rec );
				}
				catch( NumberFormatException e )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
				double returnVal = o0 + o1;
				PtgNumber n = new PtgNumber( returnVal );
				return n;
			}

			//
			// handle array formulas
			//
			String retArry = "";
			Object arrArg;
			Object arg2;

			if( args[0].getClass().isArray() )
			{
				arrArg = args[0];
				arg2 = args[1];
			}
			else
			{
				arrArg = args[1];
				arg2 = args[0];
			}

			// use first array element to determine length of values as subsequent vals might not be arrays
			int nVals = java.lang.reflect.Array.getLength( arrArg );

			boolean arg2IsArray = arg2.getClass().isArray();

			for( int j = 0; j < nVals; j++ )
			{
				Object firstOp = Array.get( arrArg, j );    // first array index j
				Object secondOp;

				if( arg2IsArray )
				{
					secondOp = Array.get( arg2, j );    // second array index j
				}
				else
				{
					secondOp = arg2;
				}

				double o0;
				double o1;

				try
				{
					o0 = getDoubleValue( firstOp, parent_rec );
					o1 = getDoubleValue( secondOp, parent_rec );
				}
				catch( NumberFormatException e )
				{
					retArry = retArry + "#VALUE!" + ",";
					continue;
				}
				double retVal = o0 + o1;
				retArry = retArry + retVal + ",";
			}

			retArry = "{" + retArry.substring( 0, retArry.length() - 1 ) + "}";
			PtgArray pa = new PtgArray();
			pa.setVal( retArry );

			return pa;
		}
		catch( NumberFormatException e )
		{
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
		catch( Exception e )
		{
			// At least log the error so the devs have a chance to see it and fix it...
			log.error( "Error during addition", e );
			// handle error ala Excel
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
	}
}