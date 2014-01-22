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
  Ptg that is a not equal to operand
   
   Evaluates to TRUE if the two top operands are not equal,
   otherwise evaluates as FALSE;
   
   
 * @see Ptg
 * @see Formula

    
*/
public class PtgNE extends GenericPtg implements Ptg
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6901661166166179786L;

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

	public PtgNE()
	{
		ptgId = 0xE;
		record = new byte[1];
		record[0] = 0xE;
	}

	/**
	 * return the human-readable String representation of
	 */
	@Override
	public String getString()
	{
		return "<>";
	}

	@Override
	public int getLength()
	{
		return PTG_NE_LENGTH;
	}

	/*  Operator specific calculate method, this one determines if the second-to-top
		operand is less than the top operand;  Returns a PtgBool

	*/
	@Override
	public Ptg calculatePtg( Ptg[] form )
	{
		try
		{
			// 20090202 KSC: Handle array formulas
			Object[] o = super.getValuesFromPtgs( form );
			boolean res;
			if( !o[0].getClass().isArray() )
			{
				//double[] dub = super.getValuesFromPtgs(form);
				// there should always be only two ptg's in this, error if not.
				if( (o == null) || (o.length != 2) )
				{
					Logger.logWarn( "calculating formula failed, wrong number of values in PtgNE" );
					return null;
				}
				// blank handling:
				// determine if any of the operands are double - if true,
				// then blank comparisons will be treated as 0's
				boolean isDouble = false;
				for( int i = 0; (i < 2) && !isDouble; i++ )
				{
					//if (!form[i].isBlank())
					isDouble = ((o[i] instanceof Double));
				}
				for( int i = 0; i < 2; i++ )
				{
					//if (form[i].isBlank()) {
					if( (o[i] != null) && o[i].toString().equals( "" ) )
					{
						if( isDouble )
						{
							o[i] = 0.0;
						}
						else
						{
							o[i] = ""; // in this case, empty cells are handled as blank, not zero
						}
					}
				}
				if( (o[0] instanceof Double) && (o[1] instanceof Double) )
				{
					res = (Math.abs( ((Double) o[0]) - (Double) o[1] )) > doublePrecision;    // compare equality to certain precision
				}
				else if( !o[0].toString().equalsIgnoreCase( o[1].toString() ) )
				{
					res = true;
				}
				else
				{
					res = false;
				}
				PtgBool pboo = new PtgBool( res );
				return pboo;
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

					if( (firstOp instanceof Double) && (secondOp instanceof Double) )
					{
						res = (Math.abs( ((Double) firstOp) - (Double) secondOp )) > doublePrecision;    // compare to certain precision instead of equality
					}
					else
					{
						res = firstOp.toString().equalsIgnoreCase( secondOp.toString() );
					}
					retArry = retArry + res + ",";
				}
			}
			retArry = "{" + retArry.substring( 0, retArry.length() - 1 ) + "}";
			PtgArray pa = new PtgArray();
			pa.setVal( retArry );
			return pa;
			/*}catch(NumberFormatException e){	// shouldn't get here!! see new code above
    		String[] s = getStringValuesFromPtgs(form);
    		if (s[0].equalsIgnoreCase(s[1]))return new PtgBool(false);
    		return new PtgBool(true);
    	}*/
		}
		catch( Exception e )
		{    // 20090212 KSC: handle error ala Excel
			PtgErr perr = new PtgErr( PtgErr.ERROR_VALUE );
			return perr;
		}
	}

}