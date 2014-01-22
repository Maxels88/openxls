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

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Random;


/*
    MathFunctionCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/

public class MathFunctionCalculator
{

	/**
	 * SUM
	 * Adds all the numbers in a range of cells.
	 * Ignores non-number fields
	 * <p/>
	 * Usage@ SUM(number1,number2, ...)
	 * Return@ PtgNumber
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcSum( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double result = 0;
		double[] dub = PtgCalculator.getDoubleValueArray( operands );
		if( dub == null )
		{
			return PtgCalculator.getNAError();
		}
		for( int i = 0; i < dub.length; i++ )
		{
			result += dub[i];
		}
		return new PtgNumber( result );
	}

	/**
	 * ABS
	 * Returns the absolute value of a number
	 */
	protected static Ptg calcAbs( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		dd = Math.abs( dd );
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}
		PtgNumber ptnum = new PtgNumber( dd );
		return ptnum;

	}

	/**
	 * ACOS
	 * Returns the arccosine of a number
	 */
	protected static Ptg calcAcos( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		dd = Math.acos( dd );
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}
		PtgNumber ptnum = new PtgNumber( dd );
		return ptnum;
	}

	/**
	 * ACOSH
	 * Returns the inverse hyperbolic cosine of a number
	 */
	protected static Ptg calcAcosh( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double x = 0.0;
		try
		{
			x = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		double dd = Math.log( x + ((1.0 + x) * Math.sqrt( (x - 1.0) / (x + 1.0) )) );
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}
		PtgNumber ptnum = new PtgNumber( dd );
		return ptnum;
	}

	/**
	 * ASIN
	 * Returns the arcsine of a number
	 */
	protected static Ptg calcAsin( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		dd = Math.asin( dd );
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}
		PtgNumber ptnum = new PtgNumber( dd );
		return ptnum;
	}

	/**
	 * ASINH
	 * Returns the inverse hyperbolic sine of a number
	 */
	protected static Ptg calcAsinh( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double x = 0.0;
		try
		{
			x = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( x ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

// KSC: TESTING: I BELIEVE THE CALCULATION IS NOT CORRECT!    
		BigDecimal bd = new BigDecimal( ((x > 0.0) ? 1.0 : -1.0) * getAcosh( Math.sqrt( 1.0 + (x * x) ) ) );
		bd.setScale( 15, BigDecimal.ROUND_HALF_UP );
//	PtgNumber ptnum = new PtgNumber(dd);
		PtgNumber ptnum = new PtgNumber( bd.doubleValue() );
		return ptnum;
	}

	/**
	 * ATAN
	 * Returns the arctangent of a number
	 */
	protected static Ptg calcAtan( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		dd = Math.atan( dd );
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}
		PtgNumber ptnum = new PtgNumber( dd );
		return ptnum;
	}

	/**
	 * ATAN2
	 * Returns the arctangent from x- and y- coordinates
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcAtan2( Ptg[] operands ) throws CalculationException
	{
		if( operands.length != 2 )
		{
			return PtgCalculator.getNAError();
		}
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		double res = Math.atan2( dd[0], dd[1] );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * ATANH
	 * Returns the inverse hyperbolic tangent of a number
	 */
	protected static Ptg calcAtanh( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( (dd > 1) || (dd < -1) )
		{
			return PtgCalculator.getError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		double res = (0.5 * Math.log( (1.0 + dd) / (1.0 - dd) ));
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * CEILING
	 * Rounds a number to the nearest multiple of significance;
	 * This takes 2 values, first the number to round, next the value of signifigance.
	 * Ick.  This is pretty intensive, so maybe a better way?
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcCeiling( Ptg[] operands ) throws CalculationException
	{
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		if( dd.length != 2 )
		{
			return PtgCalculator.getNAError();
		}
		double num = dd[0];
		double multiple = dd[1];
		double res = 0;
		while( res < num )
		{
			res += multiple;
		}
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * COMBIN
	 * Returns the number of combinations for a given number of objects
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcCombin( Ptg[] operands ) throws CalculationException
	{
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		if( dd.length != 2 )
		{
			return PtgCalculator.getNAError();
		}
		long num1 = Math.round( dd[0] );
		long num2 = Math.round( dd[1] );
		if( num1 < num2 )
		{
			return PtgCalculator.getError();
		}
		long res1 = stepFactorial( num1, (int) num2 );
		long res2 = factorial( num2 );
		double res = res1 / res2;
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * COS
	 * Returns the cosine of a number
	 */
	protected static Ptg calcCos( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		double res = Math.cos( dd );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * COSH
	 * Returns the hyperbolic cosine of a number
	 */
	protected static Ptg calcCosh( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		double res = 0.5 * (Math.exp( dd ) + Math.exp( -dd ));
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * DEGREES
	 * Converts radians to degrees
	 */
	protected static Ptg calcDegrees( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		double res = Math.toDegrees( dd );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * EVEN
	 * Rounds a number up to the nearest even integer
	 */
	protected static Ptg calcEven( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		long resnum = Math.round( dd );
		double remainder = dd % 2;
		if( remainder != 0 )
		{
			resnum++;
		}
		PtgInt pint = new PtgInt( (int) resnum );
		return pint;
	}

	/**
	 * EXP
	 * Returns e raised to the power of number.
	 * The constant e equals 2.71828182845904,
	 * the base of the natural logarithm.
	 * <p/>
	 * Example
	 * EXP(2) equals e2, or 7.389056
	 */
	protected static Ptg calcExp( Ptg[] operands )
	{
		if( (operands.length > 1) || (operands[0].getComponents() != null) )
		{ // not supported by function 
			return new PtgErr( PtgErr.ERROR_NULL );
		}
		double d = 2.718281828459045;
		Ptg p = operands[0];
		try
		{
			Double dub = new Double( String.valueOf( p.getValue() ) );
//	    double result = Math.pow(d, dub.doubleValue());
			BigDecimal result = new BigDecimal( Math.pow( d, dub.doubleValue() ) );
			result.setScale( 15, BigDecimal.ROUND_HALF_UP );
//	    PtgNumber pnum = new PtgNumber(result);
			PtgNumber pnum = new PtgNumber( result.doubleValue() );
			return pnum;
		}
		catch( NumberFormatException e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}

	/**
	 * FACT
	 * Returns the factorial of a number
	 */
	protected static Ptg calcFact( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		long dd = 0;
		try
		{
			dd = PtgCalculator.getLongValue( operands );
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		long res = MathFunctionCalculator.factorial( dd );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * FACTDOUBLE
	 * Returns the double factorial of a number
	 * Example
	 * !!6 = 6*4*2
	 * !!7 = 7*5*3*1;
	 */
	protected static Ptg calcFactDouble( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		long n = PtgCalculator.getLongValue( operands );
		if( new Double( n ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		if( n < 0 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		double res = 1;
		long endPoint = (((n % 2) == 0) ? 2 : 1);
		for( long i = n; i >= endPoint; i -= 2 )
		{
			res *= i;
		}
		if( n == 0 )
		{
			res = -1;    // by convention ...
		}
// KSC: not quite right!!!	replaced with above 
		//long res =  MathFunctionCalculator.doubleFactorial(n);
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * FLOOR
	 * Rounds a number down, toward zero.  Works just like Celing with two operands
	 * See comment above for celing for more info.
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcFloor( Ptg[] operands ) throws CalculationException
	{
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		if( dd.length != 2 )
		{
			return PtgCalculator.getError();
		}
		double num = dd[0];
		double multiple = dd[1];
		double res = 0;
		while( res < num )
		{
			res += multiple;
		}
		// drop one from the celing code to get the floor...
		res -= multiple;
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * GCD
	 * Returns the greatest common divisor
	 * TODO: Finish!
	 */
	protected static Ptg calcGCD( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		Ptg[] numbers = PtgCalculator.getAllComponents( operands );
		long gcd = 0;
		try
		{
			long n1;
			if( (n1 = PtgCalculator.getLongValue( numbers[0] )) < 0 )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}

			gcd = n1;
			for( int i = 1; i < numbers.length; i++ )
			{
				long n2;
				if( (n2 = PtgCalculator.getLongValue( numbers[i] )) < 0 )
				{
					return new PtgErr( PtgErr.ERROR_NUM );
				}

				long bigger, smaller, r;
				bigger = Math.max( n2, n1 );
				smaller = Math.min( n2, n1 );
				r = bigger % smaller;
				while( r != 0 )
				{
					bigger = smaller;
					smaller = r;
					r = bigger % smaller;
				}
				gcd = Math.min( gcd, smaller );
			}
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}

		PtgNumber pnum = new PtgNumber( gcd );
		return pnum;
	}

	/**
	 * INT
	 * Rounds a number down to the nearest integer
	 */
	protected static Ptg calcInt( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		long res = Math.round( dd );
		if( res > dd )
		{
			res--;
		}
		PtgInt pint = new PtgInt( (int) res );
		return pint;
	}

	/**
	 * LCM
	 * Returns the least common multiple
	 */
	protected static Ptg calcLCM( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		Ptg[] numbers = PtgCalculator.getAllComponents( operands );
		// algorithm:
		// LCM(a, b)= (a*b)/GCD(a, b)
		long lcm = 0;
		try
		{
			Ptg[] ops = new Ptg[2];
			if( (lcm = PtgCalculator.getLongValue( numbers[0] )) < 0 )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}

			for( int i = 1; i < numbers.length; i++ )
			{
				long n2;
				if( (n2 = PtgCalculator.getLongValue( numbers[i] )) < 0 )
				{
					return new PtgErr( PtgErr.ERROR_NUM );
				}

				ops[0] = new PtgNumber( lcm );
				ops[1] = new PtgNumber( n2 );
				long gcd = PtgCalculator.getLongValue( calcGCD( ops ) );
				lcm = (lcm * n2) / gcd;
			}
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}

		PtgNumber pnum = new PtgNumber( lcm );
		return pnum;
	}

	/**
	 * LN
	 * Returns the natural logarithm of a number
	 */
	protected static Ptg calcLn( Ptg[] operands )
	{
		if( operands.length > 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() )
		{
			return PtgCalculator.getError(); // Not a Num -- possibly PtgErr
		}
		double res = Math.log( dd );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * LOG
	 * Returns the logarithm of a number to a specified base
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcLog( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		if( dd.length > 2 )
		{
			return PtgCalculator.getError();
		}
		double num1;
		double num2;
		if( dd.length == 1 )
		{
			num1 = dd[0];
			num2 = 10;
		}
		else
		{
			num1 = dd[0];
			num2 = dd[1];
		}
		double res = Math.log( num1 ) / Math.log( num2 );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * LOG10
	 * Returns the base-10 logarithm of a number
	 */
	protected static Ptg calcLog10( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

//	double res = Math.log(dd)/Math.log(10);

		BigDecimal res = new BigDecimal( Math.log10( dd ) );
		res.setScale( 15, BigDecimal.ROUND_HALF_UP );
		PtgNumber ptnum = new PtgNumber( res.doubleValue() );
		return ptnum;
	}
/*
MDETERM
Returns the matrix determinant of an array

MINVERSE
Returns the matrix inverse of an array
*/

	/**
	 * MMULT
	 * Returns the matrix product of two arrays:
	 * The number of columns in array1 must be the same as the number of rows in array2,
	 * and both arrays must contain only numbers.
	 * MMULT returns the #VALUE! error when:
	 * Any cells are empty or contain text.
	 * The number of columns in array1 is different from the number of rows in array2.
	 */
	protected static Ptg calcMMult( Ptg[] operands )
	{
		if( operands.length != 2 )
		{
			return PtgCalculator.getNAError();
		}
		try
		{
			// error trap params:  must be numeric arrays with no empty spaces	
			double[][] a1 = PtgCalculator.getArray( operands[0] );
			double[][] a2 = PtgCalculator.getArray( operands[1] );
			if( a1[0].length != a2.length )
			{
				PtgCalculator.getValueError();
			}
			double sum = 0.0;
			for( int i = 0; i < a1[0].length; i++ )
			{
				sum += a1[0][i] * a2[i][0];
			}
			return new PtgNumber( sum );
		}
		catch( Exception e )
		{
//		Logger.logErr("MMULT: error in operands " + e.toString());
		}
		return PtgCalculator.getValueError();
	}

	/**
	 * MOD
	 * Returns the remainder from division
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcMod( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		if( dd.length != 2 )
		{
			return PtgCalculator.getError();
		}
		double num1 = dd[0];
		double num2 = dd[1];
		double res = num1 % num2;
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * MROUND
	 * Returns a number rounded to the desired multiple
	 */
	protected static Ptg calcMRound( Ptg[] operands )
	{
		if( operands.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		double m = 0.0, n = 0.0;
		try
		{
			n = operands[0].getDoubleVal();
			m = operands[1].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( ((n < 0) && (m > 0)) || ((n > 0) && (m < 0)) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		double result = Math.round( n / m ) * m;
		PtgNumber pnum = new PtgNumber( result );
		return pnum;
	}

	/**
	 * MULTINOMIAL
	 * Returns the multinomial of a set of numbers
	 */
	protected static Ptg calcMultinomial( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		Ptg[] numbers = PtgCalculator.getAllComponents( operands );
		long sum = 0;
		double facts = 1;
		for( int i = 0; i < numbers.length; i++ )
		{
			long n = PtgCalculator.getLongValue( operands[i] );
			if( n < 1 )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			sum += n;
			facts *= factorial( n );
		}
		double result = factorial( sum ) / facts;
		PtgNumber pnum = new PtgNumber( result );
		return pnum;
	}

	/**
	 * ODD
	 * Rounds a number up to the nearest odd integer
	 */
	protected static Ptg calcOdd( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		long resnum = Math.round( dd );
		// always round up!!
		if( resnum < dd )
		{
			resnum++;
		}
		double remainder = resnum % 2;
		if( remainder == 0 )
		{
			resnum++;
		}
		PtgInt pint = new PtgInt( (int) resnum );
		return pint;
	}

	/**
	 * PI
	 * Returns the value of Pi
	 */
	protected static Ptg calcPi( Ptg[] operands )
	{
		double pi = Math.PI;
		PtgNumber pnum = new PtgNumber( pi );
		return pnum;
	}

	/**
	 * POWER
	 * Returns the result of a number raised to a power
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcPower( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		if( dd.length != 2 )
		{
			return PtgCalculator.getError();
		}
//	double num1 = dd[0];
//	double num2 = dd[1];
		BigDecimal num1 = new BigDecimal( dd[0] );
		num1.setScale( 15, BigDecimal.ROUND_HALF_UP );
		BigDecimal num2 = new BigDecimal( dd[1] );
		num2.setScale( 15, BigDecimal.ROUND_HALF_UP );
		BigDecimal res = new BigDecimal( Math.pow( num1.doubleValue(), num2.doubleValue() ) );
		res.setScale( 15, BigDecimal.ROUND_HALF_UP );
		PtgNumber ptnum = new PtgNumber( res.doubleValue() );
		return ptnum;
	}

	/**
	 * PRODUCT
	 * Multiplies its arguments
	 * NOTE:  we gotta deal with ranges/refs/numbers here
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcProduct( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//propagate error
		}
		double result = dd[0];
		for( int i = 1; i < dd.length; i++ )
		{
			result = result * dd[i];
		}
		PtgNumber ptnum = new PtgNumber( result );
		return ptnum;
	}

	/**
	 * QUOTIENT
	 * Returns the integer portion of a division
	 */
	protected static Ptg calcQuotient( Ptg[] operands )
	{
		if( operands.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		double numerator = 0.0, denominator = 0.0;
		try
		{
			numerator = operands[0].getDoubleVal();
			denominator = operands[1].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		long result = new Double( numerator / denominator ).longValue();
		PtgNumber pnum = new PtgNumber( result );
		return pnum;
	}

	/**
	 * RADIANS
	 * Converts degrees to radians
	 */
	protected static Ptg calcRadians( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
		{
			return PtgCalculator.getError();
		}

		double res = Math.toRadians( dd );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * RAND
	 * Returns a random number between 0 and 1
	 */
	protected static Ptg calcRand( Ptg[] operands )
	{
		double dd = Math.random();
		PtgNumber pnum = new PtgNumber( dd );
		return pnum;
	}

	/**
	 * RANDBETWEEN
	 * Returns a random number between the numbers you specify
	 */
	protected static Ptg calcRandBetween( Ptg[] operands )
	{
		if( operands.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		int lower = 0, upper = 0;
		try
		{
			lower = operands[0].getIntVal();
			upper = operands[1].getIntVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		Random r = new Random();
		double result = r.nextInt( (upper - lower) + 1 ) + lower;
		PtgNumber pnum = new PtgNumber( result );
		return pnum;
	}

	/**
	 * ROMAN
	 * Converts an Arabic numeral to Roman, as text
	 * This one is a trip!
	 */
	protected static Ptg calcRoman( Ptg[] operands )
	{
		int[] numbers = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
		String[] letters = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
		String roman = "";
		if( operands.length == 1 )
		{
			double dd = operands[0].getDoubleVal();
			if( new Double( dd ).isNaN() ) // Not a Num -- possibly PtgErr
			{
				return PtgCalculator.getError();
			}

			int i = (int) dd;
			if( (i < 0) || (i > 3999) )
			{
				return PtgCalculator.getError(); // can't write nums that high!
			}
			for( int z = 0; z < numbers.length; z++ )
			{
				while( i >= numbers[z] )
				{
					roman += letters[z];
					i -= numbers[z];
				}
			}
		}
		PtgStr pstr = new PtgStr( roman );
		return pstr;
	}

	/**
	 * ROUND
	 * Rounds a number to a specified number of digits
	 * This one is kind of nasty.  3 cases of rounding.  If the rounding is a positive
	 * integer this is the number of digits to round to.  If negative, round up past the
	 * decimal, if 0, give an integer
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcRound( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		//if (dd.length != 2) return PtgCalculator.getError();
		// we need to handle arrays sent in, just use the first and last elements.
		double num = dd[0];
		double round = dd[dd.length - 1];
		double res = 0;
		if( round == 0 )
		{//return an int
			res = Math.round( num );
		}
		else if( round > 0 )
		{ //round the decimal that number of spaces
			double tempnum = num * Math.pow( 10, round );
			tempnum = Math.round( tempnum );
			res = tempnum / Math.pow( 10, round );
		}
		else
		{ //round up the decimal that numbe of places
			round = round * -1;
			double tempnum = num / Math.pow( 10, round );
			tempnum = Math.round( tempnum );
			res = tempnum * Math.pow( 10, round );
		}
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * ROUNDDOWN
	 * Rounds a number down, toward zero.  Acts much like round above
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcRoundDown( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		//if (dd.length != 2) return PtgCalculator.getError();
		double num = dd[0];
		double round = dd[dd.length - 1];
		double res = 0;
		if( round == 0 )
		{//return an int
			res = Math.round( num );
			if( res > num )
			{
				res--;
			}
		}
		else if( round > 0 )
		{ //round the decimal that number of spaces
			double tempnum = num * Math.pow( 10, round );
			double tempnum2 = Math.round( tempnum );
			if( tempnum2 > tempnum )
			{
				tempnum2--;
			}
			res = tempnum2 / Math.pow( 10, round );
		}
		else
		{ //round up the decimal that numbe of places
			round = round * -1;
			double tempnum = num / Math.pow( 10, round );
			double tempnum2 = Math.round( tempnum );
			if( tempnum2 > tempnum )
			{
				tempnum2--;
			}
			res = tempnum2 * Math.pow( 10, round );
		}
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * ROUNDUP
	 * Rounds a number up, away from zero
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcRoundUp( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double[] dd = PtgCalculator.getDoubleValueArray( operands );
		if( dd == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		//if (dd.length != 2) return PtgCalculator.getError();
		double num = dd[0];
		double round = dd[dd.length - 1];
		double res = 0;
		if( round == 0 )
		{//return an int
			res = Math.round( num );
			if( res < num )
			{
				res++;
			}
		}
		else if( round > 0 )
		{ //round the decimal that number of spaces
			double tempnum = num * Math.pow( 10, round );
			double tempnum2 = Math.round( tempnum );
			if( tempnum2 < tempnum )
			{
				tempnum2++;
			}
			res = tempnum2 / Math.pow( 10, round );
		}
		else
		{ //round up the decimal that numbe of places
			round = round * -1;
			double tempnum = num / Math.pow( 10, round );
			double tempnum2 = Math.round( tempnum );
			if( tempnum2 < tempnum )
			{
				tempnum2++;
			}
			res = tempnum2 * Math.pow( 10, round );
		}
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}
/*
SERIESSUM
Returns the sum of a power series based on the formula
TODO: requires pack to run
*/

	/**
	 * SIGN
	 * Returns the sign of a number
	 * return 1 if positive, -1 if negative, or 0 if 0;
	 */
	protected static Ptg calcSign( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		int res = 0;
		if( new Double( dd ).isNaN() )
		{
			return PtgCalculator.getError(); // Not a Num -- possibly PtgErr
		}
		if( dd == 0 )
		{
			res = 0;
		}
		if( dd > 0 )
		{
			res = 1;
		}
		if( dd < 0 )
		{
			res = -1;
		}
		PtgInt pint = new PtgInt( res );
		return pint;
	}

	/**
	 * SIN
	 * Returns the sine of the given angle
	 */
	protected static Ptg calcSin( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() )
		{
			return PtgCalculator.getError(); // Not a Num -- possibly PtgErr
		}
		double res = Math.sin( dd );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * SINH
	 * Returns the hyperbolic sine of a number
	 */
	protected static Ptg calcSinh( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() )
		{
			return PtgCalculator.getError(); // Not a Num -- possibly PtgErr
		}
		double res = 0.5 * (Math.exp( dd ) - Math.exp( -dd ));
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * SQRT
	 * Returns a positive square root
	 */
	protected static Ptg calcSqrt( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() )
		{
			return PtgCalculator.getError(); // Not a Num -- possibly PtgErr
		}
		double res = Math.sqrt( dd );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * SQRTPI
	 * Returns the square root of (number * PI)
	 */
	protected static Ptg calcSqrtPi( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() )
		{
			return PtgCalculator.getError(); // Not a Num -- possibly PtgErr
		}
		double res = Math.sqrt( dd * Math.PI );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;

	}
/*
SUBTOTAL
Returns a subtotal in a list or database
*/

	/**
	 * SUMIF
	 * Adds the cells specified by a given criteria
	 * <p/>
	 * You use the SUMIF function to sum the values in a range that meet criteria that you specify.
	 * For example, suppose that in a column that contains numbers, you want to sum only the values that are larger than 5. You can use the following formula:
	 * <p/>
	 * =SUMIF(B2:B25,">5")
	 * <p/>
	 * you can apply the criteria to one range and sum the corresponding values in a different range.
	 * For example, the formula =SUMIF(B2:B5, "John", C2:C5) sums only the values in the range C2:C5, where the corresponding cells in the range B2:B5 equal "John."
	 * <p/>
	 * range  Required. The range of cells that you want evaluated by criteria. Cells in each range must be numbers or names, arrays, or references that contain numbers. Blank and text values are ignored.
	 * criteria  Required. The criteria in the form of a number, expression, a cell reference, text, or a function that defines which cells will be added. For example, criteria can be expressed as 32, ">32", B5, 32, "32", "apples", or TODAY().
	 * Important  Any text criteria or any criteria that includes logical or mathematical symbols must be enclosed in double quotation marks ("). If the criteria is numeric, double quotation marks are not required.
	 * sum_range  Optional. The actual cells to add, if you want to add cells other than those specified in the range argument. If the sum_range argument is omitted, Excel adds the cells that are specified in the range argument
	 * (the same cells to which the criteria is applied).
	 * <p/>
	 * The sum_range argument does not have to be the same size and shape as the range argument.
	 * The actual cells that are added are determined by using theupper leftmost cell in the sum_range argument as the beginning cell,
	 * and then including cells that correspond in size and shape to the range argument.
	 * For example:If range is And sum_range is Then the actual cells are
	 * A1:A5 B1:B5 B1:B5
	 * A1:A5 B1:B3 B1:B5
	 * A1:B4 C1:D4 C1:D4
	 * A1:B4 C1:C2 C1:D4
	 * <p/>
	 * You can use the wildcard characters — the question mark (?) and asterisk (*) — as the criteria argument.
	 * A question mark matches any single character; an asterisk matches any sequence of characters.
	 * If you want to find an actual question mark or asterisk, type a tilde (~) preceding the character.
	 */
	protected static Ptg calcSumif( Ptg[] operands )
	{
		try
		{
			PtgArea range = Calculator.getRange( operands[0] );
			PtgArea sum_range = null;

			try
			{
				Ptg criteria = operands[1];
				if( operands.length > 2 )
				{    // see if has a sum_range; if not, source range is used for values as well as test
					sum_range = Calculator.getRange( operands[2] );
				}
				// OK at this point should have criteria, range and, if necessary, sum_range
				// algorithm:  for each entry that meets the criterium in range, get the cell; 
				// if there is a sum_range, sum the values of those cells in the sum_range that correspond to the range

				// Parse the criteria string: can be a double, a comparison, a string with wildcards ...
				// more info: You can use the wildcard characters — the question mark (?) and asterisk (*) — as the criteria argument. 
				// A question mark matches any single character; an asterisk matches any sequence of characters. 
				// If you want to find an actual question mark or asterisk, type a tilde (~) preceding the character. 
				String op = "=";    // operator, default is =
				String sCriteria = criteria.getString();    // criteria in string form
				// strip operator, if any, and parse criteria
				int j = Calculator.splitOperator( sCriteria );
				if( j > 0 )
				{
					op = sCriteria.substring( 0, j );    // extract operator, if any
				}
				sCriteria = sCriteria.substring( j );
				sCriteria = Calculator.translateWildcardsInCriteria( sCriteria );

				// stores the cells that pass the criteria expression and therefore will be summed up
				ArrayList passesList = new ArrayList();

				// test criteria for all cells in range, storing those cells (or sum_range cells) 
				// that pass in passesList
				Ptg[] cells = range.getComponents();
				Ptg[] sumrangecells = null;
				if( sum_range != null )
				{
					sumrangecells = sum_range.getComponents();
				}
				for( int i = 0; i < cells.length; i++ )
				{
					boolean passes = false;
					try
					{
						Object v = cells[i].getValue();
						passes = Calculator.compareCellValue( v, sCriteria, op );
					}
					catch( Exception e )
					{    // don't report error
						// Logger.logErr("MathFunctionCalculator.calcSumif:  error parsing " + e.toString());	// debugging only; take out when fully tested
					}
					if( passes )
					{
						if( sumrangecells != null )
						{
							passesList.add( sumrangecells[i] );
						}
						else
						{
							passesList.add( cells[i] );
						}
					}
				}

				// At this point we have a collection of all the cells that pass (or their corresponding cell in sum_range);
				// Now we sum up the values of these cells and return
				double ret = 0.0;
				for( int i = 0; i < passesList.size(); i++ )
				{
					Ptg cell = (Ptg) passesList.get( i );
					try
					{
						ret += cell.getDoubleVal();
					}
					catch( Exception e )
					{
						Logger.logErr( "MathFunctionCalculator.calcSumif:  error obtaining cell value: " + e.toString() );    // debugging only; take out when fully tested
						; // keep going
					}
				}
				return new PtgNumber( ret );
			}
			catch( Exception e )
			{
				Logger.logWarn( "could not calculate SUMIF function: " + e.toString() );
			}
		}
		catch( Exception e )
		{
			;
		}
		return new PtgErr( PtgErr.ERROR_NULL );
	}

	/**
	 * SUMIFS
	 * Adds the cells in a range (range: Two or more cells on a sheet.
	 * The cells in a range can be adjacent or nonadjacent.) that meet multiple criteria
	 * <p/>
	 * sum_range  Required. One or more cells to sum, including numbers or names, ranges, or cell references (cell reference: The set of coordinates that a cell occupies on a worksheet. For example, the reference of the cell that appears at the intersection of column B and row 3 is B3.) that contain numbers. Blank and text values are ignored.
	 * criteria_range1  Required. The first range in which to evaluate the associated criteria.
	 * criteria1  Required. The criteria in the form of a number, expression, cell reference, or text that define which cells in the criteria_range1 argument will be added. For example, criteria can be expressed as 32, ">32", B4, "apples", or "32."
	 * criteria_range2, criteria2, …  Optional. Additional ranges and their associated criteria. Up to 127 range/criteria pairs are allowed.
	 *
	 * @param operands
	 * @return
	 */
	protected static Ptg calcSumIfS( Ptg[] operands )
	{
		try
		{
			PtgArea sum_range = Calculator.getRange( operands[0] );
			Ptg[] sumrangecells = sum_range.getComponents();
			String[] ops = new String[(operands.length - 1) / 2];
			String[] criteria = new String[(operands.length - 1) / 2];
			Ptg[][] criteria_cells = new Ptg[(operands.length - 1) / 2][];
			int j = 0;
			for( int i = 1; (i + 1) < operands.length; i += 2 )
			{
				//criteria range - parse and get comprising cells
				PtgArea cr = Calculator.getRange( operands[i] );
				criteria_cells[j] = cr.getComponents();
				// each criteria_range must contain the same number of rows and columns as the sum_range 
				if( criteria_cells[j].length != sumrangecells.length )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
				// criteria for comparison, including operator
				criteria[j] = operands[i + 1].toString();
				// strip operator, if any, and parse criteria
				ops[j] = "=";    // operator, default is =
				int k = Calculator.splitOperator( criteria[j] );
				if( k > 0 )
				{
					ops[j] = criteria[j].substring( 0, k );    // extract operator, if any
				}
				criteria[j] = criteria[j].substring( k );
				criteria[j] = Calculator.translateWildcardsInCriteria( criteria[j] );
				j++;
			}

			// test criteria for all cells in range, storing those corresponding sum_range cells 
			// that pass in passesList
			// stores the cells that pass the criteria expression and therefore will be summed up
			ArrayList passesList = new ArrayList();
			// for each set of criteria, test all cells in range and evaluate
			// NOTE:  this is an implicit AND evaluation
			for( int i = 0; i < sumrangecells.length; i++ )
			{
				boolean passes = true;
				for( int k = 0; k < criteria.length; k++ )
				{
					try
					{
						Object v = criteria_cells[k][i].getValue();
						passes = Calculator.compareCellValue( v, criteria[k], ops[k] ) && passes;
						if( !passes )
						{
							break;    // no need to continue
						}
					}
					catch( Exception e )
					{    // don't report error
					}
				}
				if( passes )
				{
					passesList.add( sumrangecells[i] );
				}
			}
			// At this point we have a collection of all the cells that pass (or their corresponding cell in sum_range);
			// Now we sum up the values of these cells and return
			double ret = 0.0;
			for( int i = 0; i < passesList.size(); i++ )
			{
				Ptg cell = (Ptg) passesList.get( i );
				try
				{
					ret += cell.getDoubleVal();
				}
				catch( Exception e )
				{
					Logger.logErr( "MathFunctionCalculator.calcSumif:  error obtaining cell value: " + e.toString() );    // debugging only; take out when fully tested
					; // keep going
				}
			}

			return new PtgNumber( ret );

		}
		catch( Exception e )
		{
			;
		}
		return new PtgErr( PtgErr.ERROR_NULL );
	}

	/**
	 * SUMPRODUCT
	 * Returns the sum of the products of corresponding array components
	 */
	protected static Ptg calcSumproduct( Ptg[] operands )
	{
		double res = 0;
		int dim = 0;    // all arrays must have same dimension see below
		ArrayList arrays = new ArrayList();
		for( int i = 0; i < operands.length; i++ )
		{
			if( operands[i] instanceof PtgErr )
			{
				return new PtgErr( PtgErr.ERROR_NA );    // it's what excel does
			}
			Ptg[] a = operands[i].getComponents();
			if( a == null )
			{
				arrays.add( operands[i] );
				if( dim == 0 )
				{
					dim = 1;
					continue;
				}
				else
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
			}
			if( dim == 0 )
			{
				dim = a.length;
			}
			else if( dim != a.length )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			arrays.add( a );
		}
		for( int j = 0; j < dim; j++ )
		{
			double d = 1;
			for( int i = 0; i < arrays.size(); i++ )
			{
				Object o = ((Ptg[]) arrays.get( i ))[j].getValue();
				if( o instanceof Double )
				{
					d = d * ((Double) o).doubleValue();
				}
				else if( o instanceof Integer )
				{
					d = d * ((Integer) o).intValue();
				}
				else if( o instanceof Float )
				{
					d = d * ((Float) o).floatValue();
				}
				else
				{
					d = 0;    // non-numeric values are treated as 0's
				}
			}
			res += d;
		}
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

/*
SUMSQ
Returns the sum of the squares of the arguments

SUMX2MY2
Returns the sum of the difference of squares of corresponding values in two arrays

SUMX2PY2
Returns the sum of the sum of squares of corresponding values in two arrays

SUMXMY2
Returns the sum of squares of differences of corresponding values in two arrays
*/

	/**
	 * TAN
	 * Returns the tangent of a number
	 */
	protected static Ptg calcTan( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() )
		{
			return PtgCalculator.getError(); // Not a Num -- possibly PtgErr
		}
		double res = Math.tan( dd );
		PtgNumber ptnum = new PtgNumber( res );
		return ptnum;
	}

	/**
	 * TANH
	 * Returns the hyperbolic tangent of a number
	 */
	protected static Ptg calcTanh( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double x = 0.0;
		try
		{
			x = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		double A = Math.pow( Math.E, x );
		double B = Math.pow( Math.E, -x );
		double result = (A - B) / (A + B);
		PtgNumber pnum = new PtgNumber( result );
		return pnum;
	}

	/**
	 * TRUNC
	 * Truncates a number to an integer
	 */
	protected static Ptg calcTrunc( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getNAError();
		}
		double dd = 0.0;
		try
		{
			dd = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			return PtgCalculator.getValueError();
		}
		if( new Double( dd ).isNaN() )
		{
			return PtgCalculator.getError(); // Not a Num -- possibly PtgErr
		}
		int res = (int) dd;
		PtgInt pint = new PtgInt( res );
		return pint;
	}

	/*
 * 
 * These are some helper methods for the more brutal of the math functions
 */
//helper for asinh
	private static double getAcosh( double x )
	{
		return Math.log( x + ((1.0 + x) * Math.sqrt( (x - 1.0) / (x + 1.0) )) );
	}

	// factorial helper
	public static long factorial( long n )
	{
		long result;
		if( n <= 1 )
		{
			result = 1; // 1! is 1
		}
		// The recursive part
		else
		{
			result = n;
			long partial = factorial( n - 1 );
			result = result * partial;
		}
		return result;
	}

	/*
 * Step factoral calculates a factorial of a number
 * the number of steps specified
 * 
 */
//Combin helper, steps factorials
	private static long stepFactorial( long n, int numsteps )
	{
		long result = n;
		if( n < numsteps )
		{
			return -1;
		}
		while( numsteps > 1 )
		{
			long partial = n - 1;
			result = partial * result;
			n--;
			numsteps--;

		}
		return result;
	}

	// double factorial helper
	private static long doubleFactorial( long n )
	{
		long result;
		if( n <= 1 )
		{
			result = 1; // 1! is 1
		}
		else if( n == 2 )
		{
			result = 2;
		}
		// The recursive part
		else
		{
			result = n;
			long partial = factorial( n - 2 );
			result = result * partial;
		}
		return result;
	}

}
