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
   Equals operand
   
   Evaluates to true if the top two operands are equal, otherwise FALSE
   
   
 * @see Ptg
 * @see Formula

    
*/
public class PtgEQ extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5446048862531696036L;

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

	public PtgEQ()
	{
		ptgId = 0xB;
		record = new byte[1];
		record[0] = 0xB;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return "=";
	}

	public String toString()
	{
		return this.getString();
	}

	@Override
	public int getLength()
	{
		return PTG_EQ_LENGTH;
	}

	/*  Operator specific calculate method, this one determines if the second-to-top
		operand is equal to the top operand;  Returns a PtgBool

	*/
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		boolean res = false;
		// there should always be only two ptg's in this, error if not.
		if( form.length != 2 )
		{
			Logger.logInfo( "calculating formula, wrong number of values in PtgEQ" );
			return new PtgErr( PtgErr.ERROR_VALUE );    // 20081203 KSC: handle error's ala Excel
		}
		// check for null referenced values, a null reference is equal to the string "";
		Object[] o = super.getValuesFromPtgs( form );
		if( o == null )
		{
			return new PtgErr( PtgErr.ERROR_VALUE ); // some error in value(s)
		}
		if( o[1].getClass().isArray() && !o[0].getClass().isArray() )
		{
			Object tmp = o[0];
			o[0] = o[1];
			o[1] = tmp;
		}
		if( !o[0].getClass().isArray() )
		{
			if( o.length != 2 )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );    // 20081203 KSC: handle error's ala Excel return null;
			}
			// blank handling:
			// determine if any of the operands are double - if true,
			// then blank comparisons will be treated as 0's
			boolean isDouble = false;
			for( int i = 0; i < 2 && !isDouble; i++ )
			{
				//if (!form[i].isBlank())
				isDouble = ((o[i] instanceof Double));
			}
			for( int i = 0; i < 2; i++ )
			{
				//if (form[i].isBlank()) {
				if( o[i] != null && o[i].toString().equals( "" ) )
				{
					if( isDouble )
					{
						o[i] = new Double( 0.0 );
					}
					else
					{
						o[i] = ""; // in this case, empty cells are handled as blank, not zero
					}
				}
			}
			if( o[0] == o[1] )
			{
				res = true;
			}
			else if( o[0] == null || o[1] == null )
			{
				res = false;
			}
			else if( o[0] instanceof Double && o[1] instanceof Double )
			{
				res = (Math.abs( (((Double) o[0]).doubleValue()) - ((Double) o[1]).doubleValue() )) < doublePrecision;    // compare equality to certain precision
			}
			else if( o[0].toString().equalsIgnoreCase( o[1].toString() ) )
			{
				res = true;
			}
			// handle empty cell references vs string case, 0.0 does not match
		}
		else
		{    // handle array formulas
			String retArry = "";
			int nArrays = java.lang.reflect.Array.getLength( o );
			if( nArrays != 2 )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			int nVals = java.lang.reflect.Array.getLength( o[0] );    // use first array element to determine length of values as subsequent vals might not be arrays
			if( nVals == 0 )
			{
				retArry = "{false}";
				PtgArray pa = new PtgArray();
				pa.setVal( retArry );
				return pa;
			}
			for( int i = 0; i < nArrays - 1; i += 2 )
			{
				res = false;
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
					double fd = 0, sd = 0;
					try
					{
						fd = new Double( firstOp.toString() ).doubleValue();
						sd = new Double( secondOp.toString() ).doubleValue();
						res = ((Math.abs( fd - sd )) <= doublePrecision);    // compare to certain precision instead of equality

					}
					catch( Exception e )
					{
						//if (firstOp instanceof Double && secondOp instanceof Double)
						//res= (Math.abs((((Double)firstOp).doubleValue())-((Double)secondOp).doubleValue()))<=doublePrecision;	// compare to certain precision instead of equality 
						//else
						res = firstOp.toString().equalsIgnoreCase( secondOp.toString() );
					}
					retArry = retArry + res + ",";
				}
			}
			retArry = "{" + retArry.substring( 0, retArry.length() - 1 ) + "}";
			PtgArray pa = new PtgArray();
			pa.setVal( retArry );
			return pa;
		}
		PtgBool pboo = new PtgBool( res );
		return pboo;
	}

}