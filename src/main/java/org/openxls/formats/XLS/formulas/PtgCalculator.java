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

import org.openxls.ExtenXLS.ExcelTools;
import org.openxls.toolkit.CompatibleVector;
import org.openxls.toolkit.FastAddVector;
import org.openxls.formats.XLS.WorkBook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Enumeration;

/**
 * PtgCalculator handles some of the standard calls that all of the
 * calculator classes need.
 */

public class PtgCalculator
{
	private static final Logger log = LoggerFactory.getLogger( PtgCalculator.class );
	/**
	 * getLongValue is for single-operand functions.
	 * It returns NaN for calculations that have too many operands.
	 *
	 * @param operands
	 * @return
	 */
	protected static long getLongValue( Ptg[] operands )
	{
		Ptg[] components = operands[0].getComponents();
		if( components != null )
		{ // check if too many operands TODO: check that ONE is ok??
			if( components.length > 1 )
			{
				return (long) Double.NaN;
			}
		}
		if( operands.length > 1 )
		{ // not supported by function
			log.warn( "PtgCalculator getting Long Value for operand failed: - UNSUPPORTED BY FUNCTION" );
			return (long) Double.NaN;
		}
		Double d = null;
		try
		{
			d = operands[0].getDoubleVal();
		}
		catch( NumberFormatException e )
		{
			log.warn( "PtgCalculator getting Long Value for operand failed: " + e );
			return (long) Double.NaN;
		}
		return d.longValue();
	}

	/*
	 * See getLongValue(operand[])
	 * Does the same thing with a single operand
	 */
	protected static long getLongValue( Ptg operand )
	{
		Ptg[] ptgArr = new Ptg[1];
		ptgArr[0] = operand;
		return getLongValue( ptgArr );
	}

	// returns an array of longs from an array of ptg's	
	protected static long[] getLongValueArray( Ptg[] operands )
	{
		Ptg[] alloperands = getAllComponents( operands );
		long[] l = new long[alloperands.length];
		for( int i = 0; i < alloperands.length; i++ )
		{
			try
			{
				Double dd = operands[i].getDoubleVal();
				l[i] = dd.longValue();
			}
			catch( NumberFormatException e )
			{
				log.warn( "PtgCalculator getting Long value array failed: " + e );
				l[i] = (long) Double.NaN;
			}
		}
		return l;
	}

	/**
	 * getDoubleValue is for multi-operand functions.  It returns NaN
	 * for calculations that have to many operands.
	 *
	 * @param operands
	 * @return
	 * @throws CircularReferenceException TODO
	 */
	protected static double[] getDoubleValueArray( Ptg[] operands ) throws CalculationException
	{

		Double d = null;
		// we don't know the size ahead of time, so use a vector for now.
		CompatibleVector cv = new CompatibleVector();

		for( Ptg operand : operands )
		{
			// is it multidimensional?
			Ptg[] pthings = operand.getComponents(); // optimized -- do it once!  -jm
			if( pthings != null )
			{
				for( Ptg pthing : pthings )
				{
					cv.add( pthing );
				}
			}
			else
			{
				cv.add( operand );
			}
		}

		double[] darr = new double[cv.size()];
		int i = 0;
		Enumeration en = cv.elements();
		while( en.hasMoreElements() )
		{
			d = null;//new Double(0.0);		// 20081229 KSC: reset
			Ptg pthing = (Ptg) en.nextElement();
			Object ob = pthing.getValue();
			if( (ob == null) || ob.toString().trim().equals( "" ) )
			{    // 20060802 KSC: added trim
				darr[i] = 0;
			}
			else if( ob.toString().equals( "#CIR_ERR!" ) )
			{
				throw new CircularReferenceException( CalculationException.VALUE );
			}
			else
			{

				try
				{
					if( ob instanceof Double )
					{
						d = (Double) ob;
					}
					else
					{
						String s = ob.toString();
						d = new Double( s );
					}
				}
				catch( NumberFormatException e )
				{
					try
					{
						String s = ob.toString();
						if( s.equals( "#N/A" ) )
						{    // 20090130 KSC: if error value, propagate error (ala Excel) -- null caught in calling method propagates "#N/A"
							return null;
						}
					}
					catch( Exception ee )
					{
						log.warn( "PtgCalculator getting Double value array failed: " + ee );
						d = Double.NaN;
					}
				}
				if( d != null )
				{
					darr[i] = d;
				}
			}
			i++;
		}
		return darr;
	}

	protected static double[] getDoubleValueArray( Ptg operands ) throws CalculationException
	{
		Ptg[] ptgarr = new Ptg[1];
		ptgarr[0] = operands;
		return getDoubleValueArray( ptgarr );
	}

	/**
	 * return a 2-dimenional array of double values
	 * i.e. keep array structure of reference and array type parameters
	 * <br> ASSUMPTIONS:
	 * <br> 1- accepts only 2-d references i.e. ranges are on 1 sheet
	 * <br> 2- assumes that range reference are in proper notation i.e A1:B6, NOT B6:A1
	 *
	 * @param operand Ptg
	 * @return double[][]
	 */
	protected static double[][] getArray( Ptg operand ) throws Exception
	{
		int nrows;
		int ncols;
		double[][] arr = null;

		if( operand instanceof PtgRef )
		{
			int[] rc = ((PtgRef) operand).getIntLocation();
			String sheet = ((PtgRef) operand).getSheetName();
			WorkBook bk = operand.getParentRec().getWorkBook();
			nrows = rc[2] - rc[0] + 1;
			ncols = rc[3] - rc[1] + 1;
			arr = new double[nrows][ncols];
			for( int j = rc[1]; j <= rc[3]; j++ )
			{
				for( int i = rc[0]; i <= rc[2]; i++ )
				{
					String cell = ExcelTools.formatLocation( new int[]{ i, j } );
					arr[i - rc[0]][j - rc[1]] = bk.getCell( sheet, cell ).getDblVal();
				}
			}

		}
		else
		{ // should be an array
			String arrStr = operand.toString().substring( 1 );
			arrStr = arrStr.substring( 0, arrStr.length() - 1 );
			String[] rows = arrStr.split( ";" );
			arr = new double[rows.length][];
			for( int i = 0; i < rows.length; i++ )
			{
				String[] s = rows[i].split( ",", -1 );    // include empty strings
				arr[i] = new double[s.length];
				for( int j = 0; j < s.length; j++ )
				{
					arr[i][j] = new Double( s[j] );
				}
			}
		}
		return arr;
	}

	/*
	 * creates a generic error ptg
	 */
	protected static Ptg getError()
	{
		PtgErr perr = new PtgErr( PtgErr.ERROR_NULL );
		return perr;
	}

	/**
	 * returns an #VALUE! error ptg
	 *
	 * @return
	 */
	protected static Ptg getValueError()
	{
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * returns an #NA! error ptg
	 *
	 * @return
	 */
	protected static Ptg getNAError()
	{
		return new PtgErr( PtgErr.ERROR_NA );
	}

	/*
	 *  Get all components recurses through the ptg's and returns an array
	 * of single ptg's for all the ptgs in the operands array.  This means it
	 * converts arrays to ref's, etc.
	 */
	protected static Ptg[] getAllComponents( Ptg[] operands )
	{

		if( operands.length == 1 )
		{
			Ptg[] ret = operands[0].getComponents();
			if( ret == null )
			{
				return operands;
			}
		}

		FastAddVector v = new FastAddVector();
		for( Ptg operand : operands )
		{
			Ptg[] pthings = operand.getComponents(); // optimized -- do it once!  -jm
			if( pthings != null )
			{
				for( Ptg pthing : pthings )
				{
					v.add( pthing );
				}
			}
			else
			{
				v.add( operand );
			}
		}
		Ptg[] res = new Ptg[v.size()];
		res = (Ptg[]) v.toArray( res );
		return res;
	}

	/*
		*  Get all components recurses through the ptg's and returns an array 
		* of single ptg's for all the ptgs in the operands array.  This means it
		* converts arrays to ref's, etc.
		*/
	protected static Ptg[] getAllComponents( Ptg operand )
	{
		Ptg[] ptgArr = new Ptg[1];
		ptgArr[0] = operand;
		return getAllComponents( ptgArr );
	}

	/*
		* Returns the boolean value of a PTG.  If no boolean available then it returns false
		* Add more types here if they are available/needed
		*/
	protected static boolean getBooleanValue( Ptg operand )
	{
		if( operand instanceof PtgBool )
		{
			PtgBool b = (PtgBool) operand;
			return b.getBooleanValue();
		}
		if( operand instanceof PtgInt )
		{
			PtgInt i = (PtgInt) operand;
			return i.getBooleanVal();
		}
		return false;
	}

}