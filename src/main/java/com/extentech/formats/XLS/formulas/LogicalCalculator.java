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
    LogicalCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/
public class LogicalCalculator
{
	/**
	 * AND
	 * <p/>
	 * Returns TRUE if all its arguments are TRUE;
	 * returns FALSE if one or more arguments is FALSE.
	 * <p/>
	 * Syntax
	 * AND(logical1,logical2, ...)
	 * <p/>
	 * Logical1, logical2, ...   are 1 to 30 conditions you want to test
	 * that can be either TRUE or FALSE.
	 * <p/>
	 * The arguments must evaluate to logical values such as TRUE or FALSE,
	 * or the arguments must be arrays or references that contain logical values.
	 * If an array or reference argument contains text or empty cells,
	 * those values are ignored.
	 * If the specified range contains no logical values,
	 * AND returns the #VALUE! error value.
	 */
	protected static Ptg calcAnd( Ptg[] operands )
	{
		boolean b = true;
		Ptg[] alloperands = PtgCalculator.getAllComponents( operands );
		for( int i = 0; i < alloperands.length; i++ )
		{
			if( alloperands[i] instanceof PtgBool )
			{
				PtgBool bo = (PtgBool) alloperands[i];
				Boolean bool = (Boolean) bo.getValue();
				if( bool.booleanValue() == false )
				{
					return new PtgBool( false );
				}
			}
			else
			{
				// probably a ref, hopefully to a bool
				String s = String.valueOf( alloperands[i].getValue() );
				if( s.equalsIgnoreCase( "false" ) )
				{
					return new PtgBool( false );
				}
			}
		}
		return new PtgBool( true );
	}

	/**
	 * IF
	 * Returns one value if a condition you specify evaluates to
	 * TRUE and another value if it evaluates to FALSE.
	 * <p/>
	 * Use IF to conduct conditional tests on values and formulas.
	 * <p/>
	 * Syntax 1
	 * <p/>
	 * IF(logical_test,value_if_true,value_if_false)
	 * <p/>
	 * Logical_test   is any value or expression that can be evaluated to TRUE or FALSE.
	 * <p/>
	 * Value_if_true   is the value that is returned if logical_test is TRUE. If logical_test is TRUE and value_if_true is omitted, TRUE is returned. Value_if_true can be another formula.
	 * <p/>
	 * Value_if_false   is the value that is returned if logical_test is FALSE. If logical_test is FALSE and value_if_false is omitted, FALSE is returned. Value_if_false can be another formula.
	 * <p/>
	 * Remarks
	 * <p/>
	 * Up to seven IF functions can be nested as value_if_true and value_if_false
	 * arguments to construct more elaborate tests. See the following last example.
	 * When the value_if_true and value_if_false arguments are evaluated,
	 * IF returns the value returned by those statements.
	 * If any of the arguments to IF are arrays, every element of the array
	 * is evaluated when the IF statement is carried out. If some of the
	 * value_if_true and value_if_false arguments are action-taking functions,
	 * all of the actions are taken.
	 */
	protected static Ptg calcIf( Ptg[] operands )
	{
		// lets assume for now there are always 3 operands.. NOPE!  sometimes the missarg gets
		// lost for some reason, so we need to treat it like that.
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg determine = operands[0];
		Ptg iftrue;
		if( operands.length > 1 )
		{
			iftrue = operands[1];    // 20070212 KSC: if this is blank, return 0 (according to help!)
		}
		else
		{
			// return strings
			if( operands[0] instanceof PtgStr )
			{
				return operands[0];
			}
			return (new PtgInt( 0 ));
		}
		Ptg iffalse;
		if( operands.length > 2 )
		{
			iffalse = operands[2];
		}
		else
		{
			iffalse = new PtgMissArg();
		}
		if( iftrue instanceof PtgMissArg )
		{
			iftrue = new PtgNumber( 0 );
		}
		if( iffalse instanceof PtgMissArg )
		{
			iffalse = new PtgNumber( 0 );
		}
		if( !(determine instanceof PtgArray) )
		{
			String strval = null;
			if( !(determine instanceof PtgRef) )
			{
				strval = determine.toString();
			}
			else
			{
				try
				{
					strval = determine.getValue().toString();
				}
				catch( Exception e )
				{    // could be a formula not found error, etc. don't ignore
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
			}
			if( strval.equalsIgnoreCase( "true" ) )
			{
				return iftrue;
			}
			else
			{
				return iffalse;
			}
		}
		else
		{
			try
			{    // see what type of operands iftrue and iffalse arrays are
				String retArry = "";
				Ptg[] p = ((PtgArray) determine).getComponents();
//			boolean trueIsArray= iftrue instanceof
				boolean res = true;
				boolean trueValueIsArray = (iftrue instanceof PtgArray);
				boolean falseValueIsArray = (iffalse instanceof PtgArray);
				for( int i = 0; i < p.length; i++ )
				{
					res = (p[i].toString().equalsIgnoreCase( "true" ));
					if( res )
					{
						if( trueValueIsArray )
						{
							retArry = retArry + ((PtgArray) iftrue).arrVals.get( i ).toString() + ",";
						}
						else
						{
							retArry = retArry + iftrue + ",";
						}
					}
					else
					{ // false
						if( falseValueIsArray )
						{
							retArry = retArry + ((PtgArray) iffalse).arrVals.get( i ).toString() + ",";
						}
						else
						{
							retArry = retArry + iffalse + ",";
						}
					}
				}
				retArry = "{" + retArry.substring( 0, retArry.length() - 1 ) + "}";
				PtgArray pa = new PtgArray();
				pa.setVal( retArry );
				return pa;
			}
			catch( Exception e )
			{    // this should hit if iftrue and iffalse are array types .... TODO: handle!
				return iffalse;
			}
		}
	}

	/**
	 * Returns the logical function False
	 */
	protected static Ptg calcFalse( Ptg[] operands )
	{
		return new PtgBool( false );
	}

	/**
	 * Returns the logical function true
	 */
	protected static Ptg calcTrue( Ptg[] operands )
	{
		return new PtgBool( true );
	}

	/**
	 * Returns the opposite boolean
	 */
	protected static Ptg calcNot( Ptg[] operands )
	{
		if( String.valueOf( operands[0].getValue() ).equalsIgnoreCase( "false" ) )
		{
			return new PtgBool( true );
		}
		else
		{
			return new PtgBool( false );
		}
	}

	/**
	 * Returns the opposite boolean
	 */
	protected static Ptg calcOr( Ptg[] operands )
	{
		Ptg[] alloperands = PtgCalculator.getAllComponents( operands );
// KSC: TESTING	
//System.out.print("\tOR " + operands[0].toString() + " " + operands[1].toString() + "? ");
		for( int i = 0; i < alloperands.length; i++ )
		{
			if( String.valueOf( alloperands[i].getValue() ).equalsIgnoreCase( "true" ) )
			{
				return new PtgBool( true );
			}
		}
		return new PtgBool( false );
	}

	/**
	 * IFERROR function
	 * Returns a value you specify if a formula evaluates to an error; otherwise, returns the result of the formula.
	 * Use the IFERROR function to trap and handle errors in a formula (formula: A sequence of values, cell references, names, functions, or operators in a cell that together produce a new value. A formula always begins with an equal sign (=).).
	 * <p/>
	 * Value   is the argument that is checked for an error.
	 * <p/>
	 * Value_if_error   is the value to return if the formula evaluates to an error. The following error types are evaluated: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!.
	 * <p/>
	 * Remarks
	 * <p/>
	 * If value or value_if_error is an empty cell, IFERROR treats it as an empty string value ("").
	 * If value is an array formula, IFERROR returns an array of results for each cell in the range specified in value. See the second example below.
	 * Example: Trapping division errors by using a regular formula
	 */
	protected static Ptg calcIferror( Ptg[] operands )
	{
		if( (operands == null) || (operands.length != 2) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		if( !(operands[0] instanceof PtgArray) )
		{
			PtgBool ret = (PtgBool) InformationCalculator.calcIserror( operands );
			if( ret.getBooleanValue() )    // it's an error
			{
				return operands[1];
			}
			else                        // it's not; return calculated results of 1st operand
			{
				return operands[0];
			}
		}
		else
		{
			Ptg[] components = operands[0].getComponents();
			String retArray = "";
			for( int i = 0; i < components.length; i++ )
			{
				Ptg[] test = new Ptg[1];
				test[0] = components[i];
				PtgBool ret = (PtgBool) InformationCalculator.calcIserror( test );
				if( ret.getBooleanValue() )    // it's an error
				{
					retArray = retArray + operands[1] + ",";
				}
				else                        // it's not; return calculated results of 1st operand
				{
					retArray = retArray + test[0] + ",";
				}
			}
			retArray = "{" + retArray.substring( 0, retArray.length() - 1 ) + "}";
			PtgArray pa = new PtgArray();
			pa.setVal( retArray );
			return pa;
		}
	}

}