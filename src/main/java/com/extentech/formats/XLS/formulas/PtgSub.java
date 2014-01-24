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

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.lang.reflect.Array;

/**
 * Ptg that indicates substitution (ie minus)
 *
 * @see Ptg
 * @see Formula
 */
public class PtgSub extends GenericPtg implements Ptg
{
	private static final Logger log = LoggerFactory.getLogger( PtgSub.class );
	private static final long serialVersionUID = -3252464873846778499L;

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

	public PtgSub()
	{
		ptgId = 0x4;
		record = new byte[1];
		record[0] = 0x4;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return "-";
	}

	public String toString()
	{
		return getString();
	}

	@Override
	public int getLength()
	{
		return PTG_SUB_LENGTH;
	}

	/**
	 * Operator specific calculate method, this one subtracts one value from another
	 */
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		try
		{
			Object[] o = super.getValuesFromPtgs( form );
			if( o == null )
			{
				return new PtgErr( PtgErr.ERROR_VALUE ); // some error in value(s)
			}
			if( !o[0].getClass().isArray() )
			{
				if( o.length != 2 )
				{
					log.warn( "calculating formula failed, wrong number of values in PtgSub" );
					return new PtgErr( PtgErr.ERROR_VALUE );    // 20081203 KSC: handle error's ala Excel return null;
				}
				// blank handling:
				if( form[0].isBlank() )
				{
					o[0] = (double) 0;
				}
				if( form[1].isBlank() )
				{
					o[1] = (double) 0;
				}
				// the following should only return #VALUE! if ???
				if( !((o[0] instanceof Double) && (o[1] instanceof Double)) )
				{
					if( parent_rec == null )
					{
						return new PtgErr( PtgErr.ERROR_VALUE );
					}
					if( parent_rec.getSheet().getWindow2().getShowZeroValues() )
					{
						return new PtgInt( 0 );
					}
					return new PtgStr( "" );
				}
				double returnVal = ((Double) o[0] - (Double) o[1]);
				PtgNumber n = new PtgNumber( returnVal );
				return n;
			}    // handle array fomulas
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
					if( !((firstOp instanceof Double) && (secondOp instanceof Double)) )
					{
						if( parent_rec == null )
						{
							return new PtgErr( PtgErr.ERROR_VALUE );
						}
						if( parent_rec.getSheet().getWindow2().getShowZeroValues() )
						{
							return new PtgInt( 0 );
						}
						return new PtgStr( "" );
					}    // 20081203 KSC: handle error's ala Excel
					double retVal = (Double) firstOp - (Double) secondOp;
					retArry = retArry + retVal + ",";
				}
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
		{    // 20081125 KSC: handle error ala Excel
			// log.error("PtgSub failed:" + e);
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
	}
}