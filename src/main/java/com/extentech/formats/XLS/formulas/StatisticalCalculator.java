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

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.Logger;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Vector;

/*
    StatisticalCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/
public class StatisticalCalculator
{

	/**
	 * AVERAGE
	 * Returns the average (arithmetic mean) of the arguments.
	 * Ignores non-numbers
	 * This cannot recurse, due to averaging needs.
	 * <p/>
	 * Usage@ AVERAGE(number1,number2, ...)
	 * Returns@ PtgNumber
	 */
	protected static Ptg calcAverage( Ptg[] operands )
	{
		Vector vect = new Vector();

		for( int i = 0; i < operands.length; i++ )
		{
			Ptg[] pthings = operands[i].getComponents(); // optimized -- do it once!! -jm
			if( pthings != null )
			{
				for( int z = 0; z < pthings.length; z++ )
				{
					vect.add( pthings[z] );
				}
			}
			else
			{
				Ptg p = operands[i];
				vect.add( p );
			}
		}
		int count = 0;
//        double total = 0;
		BigDecimal bd = new BigDecimal( 0 );
		for( int i = 0; i < vect.size(); i++ )
		{
			Ptg p = (Ptg) vect.elementAt( i );
			try
			{
				if( p.isBlank() )
				{
					continue;
				}
				Object ov = p.getValue();
				if( ov != null )
				{
//                    total += Double.parseDouble(String.valueOf(ov));
					bd = bd.add( new BigDecimal( Double.parseDouble( String.valueOf( ov ) ) ) );
					count++;
				}
			}
			catch( NumberFormatException e )
			{
			}
			;
		}
		bd = bd.setScale( 15, java.math.RoundingMode.HALF_UP );
		double total = bd.doubleValue();
		if( count == 0 )
		{
			return new PtgErr( PtgErr.ERROR_DIV_ZERO );
		}
		double result = total / count;
		PtgNumber pnum = new PtgNumber( result );
		return pnum;
	}

	/**
	 * AVERAGEIF function
	 * Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria.
	 * <p/>
	 * AVERAGEIF(range,criteria,average_range)
	 * <p/>
	 * Range  is one or more cells to average, including numbers or names, arrays, or references that contain numbers.
	 * Criteria  is the criteria in the form of a number, expression, cell reference, or text that defines which cells are averaged. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.
	 * Average_range  is the actual set of cells to average. If omitted, range is used.
	 * Average_range does not have to be the same size and shape as range. The actual cells that are averaged are determined by using the top, left cell in average_range as the beginning cell, and then including cells that correspond in size and shape to range.
	 *
	 * @param operands
	 * @return
	 */
	protected static Ptg calcAverageIf( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_DIV_ZERO );
		}
		// range used to test criteria 
		Ptg[] range = operands[0].getComponents();
		// TODO: if range is blank  or a text value returns ERROR_DIV_ZERO
		String criteria = operands[1].getString().trim();
		// Parse criteria into op  + criteria 
		int i = Calculator.splitOperator( criteria );
		String op = criteria.substring( 0, i );    // extract operand
		criteria = criteria.substring( i );
		criteria = Calculator.translateWildcardsInCriteria( criteria );
		// Average_range, if present, is used for return values if range passes criteria 
		Ptg[] average_range = null;
		boolean varyRow = false;
		if( operands.length > 2 )
		{
			//The actual cells that are averaged are determined by using the top, left cell in average_range as the beginning cell, 
			// and then including cells that correspond in size and shape to range.
			int rc[] = null;
			average_range = new Ptg[range.length];
			average_range[0] = operands[2].getComponents()[0];    // start with top left of average range
			String sheet = "";
			try
			{
				rc = average_range[0].getIntLocation();
				if( range[0].getIntLocation()[0] != range[range.length - 1].getIntLocation()[0] )    // determine if range is varied across row or column
				{
					varyRow = true;
				}
				sheet = ((PtgRef) average_range[0]).getSheetName() + "!";
			}
			catch( Exception e )
			{
				;
			}
			for( int j = 1; j < average_range.length; j++ )
			{
				if( varyRow )
				{
					rc[0]++;
				}
				else
				{
					rc[1]++;
				}
				average_range[j] = new PtgRef3d();
				average_range[j].setParentRec( range[0].getParentRec() );
				average_range[j].setLocation( sheet + ExcelTools.formatLocation( rc ) );
			}
		}
		int nresults = 0;
		double result = 0.0;
		for( int j = 0; j < range.length; j++ )
		{
			Object val = range[j].getValue();
			// TODO: TRUE and FALSE values are ignored
			// TODO: blank cells are treated as 0's
			if( Calculator.compareCellValue( val, criteria, op ) )
			{
				try
				{
					if( average_range != null )
					{
						val = average_range[j].getValue();
						if( val == null )    // if a cell is empty it's ignored --
						{
							continue;
						}
					}
					result += ((Number) val).doubleValue();
				}
				catch( ClassCastException e )
				{
					;
				}
				nresults++;
			}
		}

		// If no cells in the range meet the criteria, AVERAGEIF returns the #DIV/0! error value
		if( nresults == 0 )
		{
			return new PtgErr( PtgErr.ERROR_DIV_ZERO );
		}
		// otherwise, average
		return new PtgNumber( result / nresults );
	}

	/**
	 * AVERAGEIFS
	 * Returns the average (arithmetic mean) of all cells that meet multiple criteria.
	 * AVERAGEIFS(average_range,criteria_range1,criteria1,criteria_range2,criteria2…)
	 * Average_range   is one or more cells to average, including numbers or names, arrays, or references that contain numbers.
	 * Criteria_range1, criteria_range2, …   are 1 to 127 ranges in which to evaluate the associated criteria.
	 * Criteria1, criteria2, …   are 1 to 127 criteria in the form of a number, expression, cell reference, or text that define which cells will be averaged. For example, criteria can be expressed as 32, "32", ">32", "apples", or B4.
	 *
	 * @param operands
	 * @return
	 */
	protected static Ptg calcAverageIfS( Ptg[] operands )
	{
		try
		{
			PtgArea average_range = Calculator.getRange( operands[0] );
			Ptg[] averagerangecells = average_range.getComponents();
			if( averagerangecells.length == 0 )
			{
				return new PtgErr( PtgErr.ERROR_DIV_ZERO );
			}
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
				if( criteria_cells[j].length != averagerangecells.length )
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

			// test criteria for all cells in range, storing those corresponding average_range cells 
			// that pass in passesList
			// stores the cells that pass the criteria expression and therefore will be averaged
			ArrayList passesList = new ArrayList();
			// for each set of criteria, test all cells in range and evaluate
			// NOTE:  this is an implicit AND evaluation
			for( int i = 0; i < averagerangecells.length; i++ )
			{
				boolean passes = true;
				for( int k = 0; k < criteria.length; k++ )
				{
					try
					{
						Object v = criteria_cells[k][i].getValue();
						// If cells in average_range cannot be translated into numbers, AVERAGEIFS returns the #DIV0! error value. 
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
					passesList.add( averagerangecells[i] );
				}
			}
			// If no cells in the range meet the criteria, AVERAGEIF returns the #DIV/0! error value
			if( passesList.size() == 0 )
			{
				return new PtgErr( PtgErr.ERROR_DIV_ZERO );
			}

			// At this point we have a collection of all the cells that pass (or their corresponding cell in sum_range);
			// Now we sum up the values of these cells and return
			double result = 0.0;
			for( int i = 0; i < passesList.size(); i++ )
			{
				Ptg cell = (Ptg) passesList.get( i );
				try
				{
					result += cell.getDoubleVal();
				}
				catch( Exception e )
				{
					Logger.logErr( "MathFunctionCalculator.calcAverageIfS:  error obtaining cell value: " + e.toString() );    // debugging only; take out when fully tested
					; // keep going
				}
			}

			// otherwise, average
			return new PtgNumber( result / passesList.size() );

		}
		catch( Exception e )
		{
			;
		}
		return new PtgErr( PtgErr.ERROR_NULL );
	}

	/**
	 * AVEDEV(number1,number2, ...)
	 * Number1, number2, ...   are 1 to 30 arguments for which you want
	 * the average of the absolute deviations. You can also use a
	 * single array or a reference to an array instead of arguments
	 * separated by commas.
	 * <p/>
	 * The arguments must be either numbers or names,
	 * arrays, or references that contain numbers.
	 * <p/>
	 * If an array or reference argument contains text,
	 * logical values, or empty cells, those values are ignored;
	 * however, cells with the value zero are included.
	 */
	protected static Ptg calcAveDev( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		// Get the average for the mean		
		PtgNumber av = (PtgNumber) StatisticalCalculator.calcAverage( operands );
		double average = -0.001;
		try
		{
			Double dd = new Double( String.valueOf( av.getValue() ) );
			average = dd;
		}
		catch( NumberFormatException e )
		{
		}
		;
		if( average == -0.001 )
		{
			return PtgCalculator.getError();
		}

		// work out the total deviation
		double total = 0;
		int count = 0;
		Double d;
		Ptg[] alloperands = PtgCalculator.getAllComponents( operands );
		for( int i = 0; i < alloperands.length; i++ )
		{
			Ptg resPtg = alloperands[i];
			try
			{  // some fields may be text, so handle gracefully
				if( resPtg.getValue() != null )
				{
					d = new Double( String.valueOf( resPtg.getValue() ) );
					double dub = d;
					dub = average - dub;
					dub = Math.abs( dub );
					total += dub;
					count++;
				}
			}
			catch( NumberFormatException e )
			{
			}
			;
		}
		// work out the mean deviation
		double mean = total / count;
		PtgNumber pnum = new PtgNumber( mean );
		return pnum;
	}

	/**
	 * AVERAGEA
	 * Returns the average of its arguments, including numbers, text,
	 * and logical values
	 * <p/>
	 * The arguments must be numbers, names, arrays, or references.
	 * Array or reference arguments that contain text evaluate as 0 (zero).
	 * Empty text ("") evaluates as 0 (zero).
	 * Arguments that contain TRUE evaluate as 1;
	 * arguments that contain FALSE evaluate as 0 (zero).
	 */
	protected static Ptg calcAverageA( Ptg[] operands )
	{
		Ptg[] alloperands = PtgCalculator.getAllComponents( operands );
		double total = 0;
		for( int i = 0; i < alloperands.length; i++ )
		{
			Ptg p = alloperands[i];
			try
			{
				Object ov = p.getValue();
				if( ov != null )
				{
					if( String.valueOf( ov ) == "true" )
					{
						total++;
					}
					else
					{
						total += Double.parseDouble( String.valueOf( ov ) );
					}
				}
			}
			catch( NumberFormatException e )
			{
			}
			;

		}
		double result = total / alloperands.length;
		PtgNumber pnum = new PtgNumber( result );
		return pnum;
	}
 /*
BETADIST
 Returns the cumulative beta probability density function
 
BETAINV
 Returns the inverse of the cumulative beta probability density function
 
BINOMDIST
 Returns the individual term binomial distribution probability
 
CHIDIST
 Returns the one-tailed probability of the chi-squared distribution
 
CHIINV
 Returns the inverse of the one-tailed probability of the chi-squared distribution
 
CHITEST
 Returns the test for independence
 
CONFIDENCE
 Returns the confidence interval for a population mean
 */

	/**
	 * CORREL
	 * Returns the correlation coefficient between two data sets
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcCorrel( Ptg[] operands ) throws CalculationException
	{
		// get the covariance
		PtgNumber pnum = (PtgNumber) calcCovar( operands );
		double covar = pnum.getVal();
		Ptg[] xPtg = new Ptg[1];
		xPtg[0] = operands[0];
		Ptg[] yPtg = new Ptg[1];
		yPtg[0] = operands[1];
		pnum = (PtgNumber) calcAverage( xPtg );
		double xMean = pnum.getVal();
		pnum = (PtgNumber) calcAverage( yPtg );
		double yMean = pnum.getVal();
		double[] xVals = PtgCalculator.getDoubleValueArray( xPtg );
		double[] yVals = PtgCalculator.getDoubleValueArray( yPtg );
		if( (xVals == null) || (yVals == null) )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		double xstat = 0;
		for( int i = 0; i < xVals.length; i++ )
		{
			xstat += Math.pow( (xVals[i] - xMean), 2 );
		}
		xstat = xstat / xVals.length;
		xstat = Math.sqrt( xstat );
		double ystat = 0;
		for( int i = 0; i < yVals.length; i++ )
		{
			ystat += Math.pow( (yVals[i] - yMean), 2 );
		}
		ystat = ystat / yVals.length;
		ystat = Math.sqrt( ystat );
		double retval = covar / (ystat * xstat);
		return new PtgNumber( retval );
	}

	/**
	 COUNT
	 Counts how many numbers are in the list of arguments
	 */
	/**
	 * Counts the number of cells that contain numbers
	 * and or dates.
	 * Use COUNT to get the number of entries in a number
	 * field in a range or array of numbers.
	 * <p/>
	 * Usage@ COUNT(A1:A5,A9)
	 * Return@ PtgInt
	 * TODO: implement counting of dates!
	 */
	protected static Ptg calcCount( Ptg[] operands )
	{
		int count = 0;
		for( int i = 0; i < operands.length; i++ )
		{
			Ptg[] pref = operands[i].getComponents(); // optimized -- do it once!! -jm
			if( pref != null )
			{ // it is some sort of range  
				for( int z = 0; z < pref.length; z++ )
				{
					Object o = pref[z].getValue();
					if( o != null )
					{
						try
						{
							Double n = new Double( String.valueOf( o ) );
							count++;
						}
						catch( NumberFormatException e )
						{
						}
					}
				}
			}
			else
			{  // it's a single ptgref
				Object o = operands[i].getValue();
				if( o != null )
				{
					try
					{
						Double n = new Double( String.valueOf( o ) );
						count++;
					}
					catch( NumberFormatException e )
					{
					}
					;
				}
			}
		}
		PtgInt pint = new PtgInt( count );
		return pint;

	}

	/**
	 * COUNTA
	 * Counts the number of non-blank cells within a range
	 */
	protected static Ptg calcCountA( Ptg[] operands )
	{
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		int count = 0;
		for( int i = 0; i < allops.length; i++ )
		{
		   /* 20081120 KSC: blnaks are handled differently as Excel counts blank cells as 0's
		   Object o = allops[i].getValue();
		   if (o != null) count++;
		   */
			if( !allops[i].isBlank() )
			{
				count++;
			}
		}
		return new PtgInt( count );
	}

	/**
	 * COUNTBLANK
	 * Counts the number of blank cells within a range
	 */
	protected static Ptg calcCountBlank( Ptg[] operands )
	{
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		int count = 0;
		for( int i = 0; i < allops.length; i++ )
		{
			if( allops[i].isBlank() )    // 20081112 KSC: was Object o = getValue(); if (o==null) count++;
			{
				count++;
			}
		}
		return new PtgInt( count );
	}

	/**
	 * COUNTIF
	 * Counts the number of non-blank cells within a range which meet the given criteria
	 * BRUTAL!!
	 */
	protected static Ptg calcCountif( Ptg[] operands ) throws FunctionNotSupportedException
	{
		if( operands.length != 2 )
		{
			return PtgCalculator.getError();
		}
		String matchStr = String.valueOf( operands[1].getValue() );
		boolean donumber = true;
		double matchDub = 0;
		try
		{    // this method matches strings or numbers, here is where we differentiate
			Double d = new Double( matchStr );
			matchDub = d;
		}
		catch( Exception e )
		{
			donumber = false;
		}
		double count = 0;
		Ptg[] pref = (Ptg[]) operands[0].getComponents(); // optimize by doing it one time!!! this thing gets slow....-jm
		if( pref != null )
		{ // it is some sort of range  
			for( int z = 0; z < pref.length; z++ )
			{
				Object o = pref[z].getValue();
				if( o != null )
				{
					String match2 = o.toString();
					if( donumber )
					{
						try
						{
							Double d = new Double( match2 );
							double matchDub2 = d;
							if( matchDub == matchDub2 )
							{
								count++;
							}
						}
						catch( NumberFormatException e )
						{
						}
						;
					}
					else
					{
						if( matchStr.equalsIgnoreCase( match2 ) )
						{
							count++;
						}
					}
				}
			}
		}
		else
		{  // it's a single ptgref
			Object o = operands[0].getValue();
			if( o != null )
			{
				if( o != null )
				{
					String match2 = o.toString();
					if( donumber )
					{
						try
						{
							Double d = new Double( match2 );
							double matchDub2 = d;
							if( matchDub == matchDub2 )
							{
								count++;
							}
						}
						catch( NumberFormatException e )
						{
						}
						;
					}
					else
					{
						if( matchStr.equalsIgnoreCase( match2 ) )
						{
							count++;
						}
					}
				}

			}
		}
		PtgNumber pnum = new PtgNumber( count );
		return pnum;
	}

	/**
	 * COUNTIFS
	 * criteria_range1  Required. The first range in which to evaluate the associated criteria.
	 * criteria1  Required. The criteria in the form of a number, expression, cell reference, or text that define which cells will be counted. For example, criteria can be expressed as 32, ">32", B4, "apples", or "32".
	 * criteria_range2, criteria2, ...  Optional. Additional ranges and their associated criteria. Up to 127 range/criteria pairs are allowed.
	 *
	 * @param operands
	 * @return
	 */
	protected static Ptg calcCountIfS( Ptg[] operands )
	{
		try
		{
			String[] ops = new String[operands.length / 2];
			String[] criteria = new String[operands.length / 2];
			Ptg[][] criteria_cells = new Ptg[operands.length / 2][];
			for( int i = 0; (i + 1) < operands.length; i += 2 )
			{
				//criteria range - parse and get comprising cells
				PtgArea cr = Calculator.getRange( operands[i] );
				criteria_cells[i / 2] = cr.getComponents();
				// each criteria_range must contain the same number of rows and columns as the criteriarange 
				if( (i > 0) && (criteria_cells[i / 2].length != criteria_cells[0].length) )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
				// criteria for comparison, including operator
				criteria[i / 2] = operands[i + 1].toString();
				// strip operator, if any, and parse criteria
				ops[i / 2] = "=";    // operator, default is =
				int k = Calculator.splitOperator( criteria[i / 2] );
				if( k > 0 )
				{
					ops[i / 2] = criteria[i / 2].substring( 0, k );    // extract operator, if any
				}
				criteria[i / 2] = criteria[i / 2].substring( k );
				criteria[i / 2] = Calculator.translateWildcardsInCriteria( criteria[i / 2] );
			}

			// test criteria for all cells in range, counting each cell that passes 
			// for each set of criteria, test all cells in range and evaluate
			int count = 0;
			for( int i = 0; i < criteria_cells[0].length; i++ )
			{
				boolean passes = true;
				for( int k = 0; k < criteria.length; k++ )
				{
					try
					{
						Object v = criteria_cells[k][i].getValue();
						//  the criteria argument is a reference to an empty cell, the COUNTIFS function treats the empty cell as a 0 value.
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
					count++;
				}
			}

			return new PtgNumber( count );

		}
		catch( Exception e )
		{
			;
		}
		return new PtgErr( PtgErr.ERROR_NULL );
	}

	/**
	 * COVAR
	 * Returns covariance, the average of the products of paired deviations
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcCovar( Ptg[] operands ) throws CalculationException
	{
		Ptg[] xMeanPtg = new Ptg[1];
		xMeanPtg[0] = operands[0];
		Ptg[] yMeanPtg = new Ptg[1];
		yMeanPtg[0] = operands[1];
		PtgNumber pnum = (PtgNumber) calcAverage( xMeanPtg );
		double xMean = pnum.getVal();
		pnum = (PtgNumber) calcAverage( yMeanPtg );
		double yMean = pnum.getVal();
		double[] xVals = PtgCalculator.getDoubleValueArray( xMeanPtg );
		double[] yVals = PtgCalculator.getDoubleValueArray( yMeanPtg );
		if( (xVals == null) || (yVals == null) )
		{
			return new PtgErr( PtgErr.ERROR_NA );//propagate error
		}
		double xyMean = 0;
		if( xVals.length == yVals.length )
		{
			int addvals = 0;
			for( int i = 0; i < xVals.length; i++ )
			{
				addvals += (xVals[i] * yVals[i]);
			}
			xyMean = addvals / xVals.length;
		}
		else
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		double retval = xyMean - (xMean * yMean);
		return new PtgNumber( retval );
	}
 /*
CRITBINOM
 Returns the smallest value for which the cumulative binomial distribution is less than or equal to a criterion value
 
DEVSQ
 Returns the sum of squares of deviations
 
EXPONDIST
 Returns the exponential distribution
 
FDIST
 Returns the F probability distribution
 
FINV
 Returns the inverse of the F probability distribution
 
FISHER
 Returns the Fisher transformation
 
FISHERINV
 Returns the inverse of the Fisher transformation
 */

	/**
	 * FORECAST
	 * Returns a value along a linear trend
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcForecast( Ptg[] operands ) throws CalculationException
	{
		if( operands.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg[] p = new Ptg[2];
		p[0] = operands[0];
		p[1] = operands[1];
		PtgNumber icept = (PtgNumber) calcIntercept( p );
		double intercept = icept.getVal();
		PtgNumber slp = (PtgNumber) calcSlope( p );
		double slope = slp.getVal();
		Ptg px = operands[0];
		double knownX = new Double( String.valueOf( px.getValue() ) );
		double retval = (slope * knownX) + intercept;
		return new PtgNumber( retval );
	}

	/**
	 * FREQUENCY
	 * Returns a frequency distribution as a vertical array
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcFrequency( Ptg[] operands ) throws CalculationException
	{
		Ptg[] firstArr = PtgCalculator.getAllComponents( operands[0] );
		Ptg[] secondArr = PtgCalculator.getAllComponents( operands[1] );
		CompatibleVector t = new CompatibleVector();
		for( int i = 0; i < secondArr.length; i++ )
		{
			Ptg p = secondArr[i];
			try
			{
				Double d = new Double( String.valueOf( p.getValue() ) );
				t.addOrderedDouble( d );
			}
			catch( NumberFormatException e )
			{
			}
			;
		}
		Double[] binsArr = new Double[t.size()];
		double[] dataArr;
		dataArr = PtgCalculator.getDoubleValueArray( firstArr );
		if( dataArr == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		t.toArray( binsArr );
		int[] retvals = new int[secondArr.length + 1];
		for( int i = 0; i < dataArr.length; i++ )
		{
			for( int x = 0; x < binsArr.length; x++ )
			{
				if( dataArr[i] <= binsArr[x] )
				{
					retvals[x]++;
					x = binsArr.length;
				}
				else if( dataArr[i] > binsArr[binsArr.length - 1] )
				{
					retvals[binsArr.length]++;
					x = binsArr.length;
				}
			}
		}
		// keep the original locations, so we can put the end result array in the correct order.
		// not used!	double[] originalLocs = PtgCalculator.getDoubleValueArray(secondArr);
		String ret = "{";
		for( int i = 0; i < retvals.length; i++ )
		{
			ret += retvals[i] + ",";
		}
		ret = ret.substring( 0, ret.length() - 1 ); // get rid of final comma
		ret += "}";

		PtgArray returnArr = new PtgArray();
		returnArr.setVal( ret );
		return returnArr;
	}
 /*
FTEST
 Returns the result of an F-test
 
GAMMADIST
 Returns the gamma distribution
 
GAMMAINV
 Returns the inverse of the gamma cumulative distribution
 
GAMMALN
 Returns the natural logarithm of the gamma function, G(x)
 
GEOMEAN
 Returns the geometric mean
 
GROWTH
 Returns values along an exponential trend
 
HARMEAN
 Returns the harmonic mean
 
HYPGEOMDIST
 Returns the hypergeometric distribution
 */

	/**
	 * INTERCEPT
	 * Returns the intercept of the linear regression line
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcIntercept( Ptg[] operands ) throws CalculationException
	{
		double[] yvals;
		yvals = PtgCalculator.getDoubleValueArray( operands[0] );

		double[] xvals = PtgCalculator.getDoubleValueArray( operands[1] );
		if( (xvals == null) || (yvals == null) )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		double sumXVals = 0;
		for( int i = 0; i < xvals.length; i++ )
		{
			sumXVals += xvals[i];
		}
		double sumYVals = 0;
		for( int i = 0; i < yvals.length; i++ )
		{
			sumYVals += yvals[i];
		}
		double sumXYVals = 0;
		for( int i = 0; i < yvals.length; i++ )
		{
			sumXYVals += xvals[i] * yvals[i];
		}
		double sqrXVals = 0;
		for( int i = 0; i < xvals.length; i++ )
		{
			sqrXVals += xvals[i] * xvals[i];
		}
		double toparg = (sumXVals * sumXYVals) - (sumYVals * sqrXVals);
		double bottomarg = (sumXVals * sumXVals) - (sqrXVals * xvals.length);
		double res = toparg / bottomarg;
		return new PtgNumber( res );
	}
 /*
KURT
 Returns the kurtosis of a data set
 */

	/**
	 * LARGE
	 * <p/>
	 * Returns the k-th largest value in a data set. You can use this function to select a value based on its relative standing.
	 * For example, you can use LARGE to return the highest, runner-up, or third-place score.
	 * <p/>
	 * LARGE(array,k)
	 * <p/>
	 * Array   is the array or range of data for which you want to determine the k-th largest value.
	 * K   is the position (from the largest) in the array or cell range of data to return.
	 * <p/>
	 * If array is empty, LARGE returns the #NUM! error value.
	 * If k ≤ 0 or if k is greater than the number of data points, LARGE returns the #NUM! error value.
	 * If n is the number of data points in a range, then LARGE(array,1) returns the largest value, and LARGE(array,n) returns the smallest value.
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcLarge( Ptg[] operands ) throws CalculationException
	{
		if( operands.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg rng = operands[0];
		Ptg[] array = PtgCalculator.getAllComponents( rng );
		if( array.length == 0 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int k = new Double( PtgCalculator.getDoubleValueArray( operands[1] )[0] ).intValue();
		if( (k <= 0) || (k > array.length) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		CompatibleVector sortedValues = new CompatibleVector();
		for( int i = 0; i < array.length; i++ )
		{
			Ptg p = array[i];
			try
			{
				Double d = new Double( String.valueOf( p.getValue() ) );
				sortedValues.addOrderedDouble( d );
			}
			catch( NumberFormatException e )
			{
			}
			;
		}
		// reverse array
		Double[] dubRefs = new Double[sortedValues.size()];
		for( int i = 0; i < dubRefs.length; i++ )
		{
			dubRefs[i] = (Double) sortedValues.last();
			sortedValues.remove( sortedValues.size() - 1 );
		}

		return new PtgNumber( dubRefs[k - 1] );
/*	 
	 
	 try {
		 Ptg[] parray= PtgCalculator.getAllComponents(operands[0]);
		 Object[] array= new Object[parray.length];
		 for (int i= 0; i < array.length; i++) {
			 array[i]= parray[i].getValue();
			 if (array[i] instanceof Integer) {	// convert all to double if possible for sort below (cannot have mixed array for Arrays.sort)
				 try {	
					 array[i]= new Double(((Integer)array[i]).intValue());
				 } catch (Exception e) {				 
				 }
		 	}
		 }
		 // now sort		 
		 java.util.Arrays.sort(array);
		 int position= (int) PtgCalculator.getLongValue(operands[1]);
		 // now return the nth item in the sorted (asc) array 
		 if (position >=0 && position <=array.length) {
			 Object ret= array[array.length-position];
			 if (ret instanceof Double)
				 return new PtgNumber(((Double)ret).doubleValue());
			 else if (ret instanceof Boolean)
				 return new PtgBool(((Boolean)ret).booleanValue());
			 else if (ret instanceof String)
				 return new PtgStr((String)ret);
		 }
	 } catch (Exception e) {
		 
	 }
	 return new PtgErr(PtgErr.ERROR_NUM);
*/
	}

	/**
	 * LINEST
	 * Returns the parameters of a linear trend
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcLineSt( Ptg[] operands ) throws CalculationException
	{
		double[] ys = PtgCalculator.getDoubleValueArray( operands[0] );
		if( ys == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		double[] xs;
		if( (operands.length == 1) || (operands[1] instanceof PtgMissArg) )
		{
			// create a 1,2,3 array
			xs = new double[ys.length];
			for( int i = 0; i < ys.length; i++ )
			{
				xs[i] = i;
			}
		}
		else
		{
			xs = PtgCalculator.getDoubleValueArray( operands[1] );
			if( xs == null )
			{
				return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
			}
		}

		boolean getIntercept = false;
		if( operands.length > 2 )
		{
			if( !(operands[2] instanceof PtgMissArg) )
			{
				getIntercept = PtgCalculator.getBooleanValue( operands[2] );
			}
		}
		boolean statistics = false;
		if( operands.length > 3 )
		{
			if( !(operands[3] instanceof PtgMissArg) )
			{
				statistics = PtgCalculator.getBooleanValue( operands[3] );
			}
		}

		Ptg ps = calcSlope( operands );
		if( ps instanceof PtgErr )
		{
			return ps;
		}
		PtgNumber Pslope = (PtgNumber) calcSlope( operands );
		double slope = Pslope.getVal();  // a1 val
		PtgNumber Pintercept = (PtgNumber) calcIntercept( operands );
		double intercept = Pintercept.getVal(); // b1 val

		if( (operands.length > 3) && ((operands[3] instanceof PtgBool) || (operands[3] instanceof PtgInt)) )
		{
			boolean b = PtgCalculator.getBooleanValue( operands[3] );
			if( !b )
			{
				String retstr = "{" + slope + "," + intercept + "},";
				retstr += "{" + slope + "," + intercept + "},";
				retstr += "{" + slope + "," + intercept + "},";
				retstr += "{" + slope + "," + intercept + "},";
				retstr += "{" + slope + "," + intercept + "}";
				PtgArray para = new PtgArray();
				para.setVal( retstr );
				return para;
			}
		}
		Ptg[] p = new Ptg[1];

		// figure out the stdev of the slope
		PtgNumber Psteyx = (PtgNumber) calcSteyx( operands );
		double steyx = Psteyx.getVal(); // b3 val

		// calc the y error percentage
		double yError = steyx * steyx;
		p[0] = operands[1];
		PtgNumber vp = (PtgNumber) calcVarp( p );
		double Sxx = vp.getVal() * ys.length;
		yError = yError / Sxx;
		yError = Math.sqrt( yError ); // A2 val

		// calculate degrees of freedom 
		int degFreedom = ys.length - 2; // b4 val

		// calculate standard error of intercept
		double sumXsquared = 0;
		double sumSquaredX = 0;
		double sumXYsquared = 0;
		for( int i = 0; i < xs.length; i++ )
		{
			sumSquaredX += (xs[i] * xs[i]);
			sumXsquared += xs[i];
			sumXYsquared += (xs[i] * ys[i]);
		}
		sumXsquared *= sumXsquared;
		sumXYsquared *= sumXYsquared;
		double interceptError = 1 / (xs.length - (sumXsquared / sumSquaredX));
		interceptError = Math.sqrt( interceptError );
		interceptError *= steyx; //b2val

		// calculate residual SS
		// first create array of predicted values for the linear array
		double[] predicted = new double[xs.length];
		double residualSS = 0; // b5value
		for( int i = 0; i < xs.length; i++ )
		{
			predicted[i] = intercept + (xs[i] * slope);
			double d = (predicted[i] - ys[i]);
			residualSS += (d * d);
		}

		// calculate regression SS
		p[0] = operands[0];
		PtgNumber pnum = (PtgNumber) calcAverage( p );
		double average = pnum.getVal();
		double regressionSS = 0;
		for( int i = 0; i < xs.length; i++ )
		{
			double d = (predicted[i] - average);
			regressionSS += (d * d);//A5 value
		}
		p = new Ptg[2];
		p[0] = operands[0];
		p[1] = operands[1];
		pnum = (PtgNumber) calcRsq( p );
		double r2 = pnum.getVal();    // A3

		// calculate the F value
		double F = (regressionSS / 1) / (residualSS / degFreedom); // A4

		// construct the string for creating ptgarray
		String retstr = "{" + slope + "," + intercept + "},";
		retstr += "{" + yError + "," + interceptError + "},";
		retstr += "{" + r2 + "," + steyx + "},";
		retstr += "{" + F + "," + degFreedom + "},";
		retstr += "{" + regressionSS + "," + residualSS + "}";

		PtgArray parr = new PtgArray();
		parr.setVal( retstr );

		return parr;
	}
 /*
LOGEST
 Returns the parameters of an exponential trend
 
LOGINV
 Returns the inverse of the lognormal distribution
 
LOGNORMDIST
 Returns the cumulative lognormal distribution
 */

	/**
	 * MAX
	 * Returns the largest value in a set of values.
	 * Ignores non-number fields
	 * Recursively calls for ranges.
	 * <p/>
	 * Usage@ MAX(number1,number2, ...)
	 * returns@ PtgNumber
	 */
	//untested
	protected static Ptg calcMax( Ptg[] operands )
	{
		double result = java.lang.Double.MIN_VALUE;        // 20090129 KSC -1;
		Double d = null;
		for( int i = 0; i < operands.length; i++ )
		{
			Ptg[] pthings = operands[i].getComponents(); // optimized -- do it once!! -jm
			if( pthings != null )
			{
				Ptg resPtg = StatisticalCalculator.calcMax( pthings );
				try
				{  // some fields may be text, so handle gracefully
					if( resPtg.getValue() != null )
					{
						d = new Double( String.valueOf( resPtg.getValue() ) );
					}
					if( d > result )
					{
						result = d;
					}
				}
				catch( NumberFormatException e )
				{
				}
				;
			}
			else
			{
				Ptg p = operands[i];
				try
				{
					Object ov = p.getValue();
					if( ov != null )
					{
						d = new Double( String.valueOf( ov ) );
					}
					if( d > result )
					{
						result = d;
					}
				}
				catch( NumberFormatException e )
				{
				}
				catch( NullPointerException e )
				{
				}
			}
		}
		if( result == java.lang.Double.MIN_VALUE )    // 20090129 KSC:
		{
			result = 0;        //appears to be default in error situations
		}
		PtgNumber pnum = new PtgNumber( result );
		return pnum;
	}

	/**
	 * MAXA
	 * Returns the maximum value in a list of arguments, including numbers, text, and logical values
	 * <p/>
	 * Arguments can be the following: numbers; names, arrays, or references that contain numbers; text representations of numbers; or logical values, such as TRUE and FALSE, in a reference.
	 * Logical values and text representations of numbers that you type directly into the list of arguments are counted.
	 */
	protected static Ptg calcMaxA( Ptg[] operands )
	{
		Ptg[] alloperands = PtgCalculator.getAllComponents( operands );
		if( alloperands.length == 0 )
		{
			return new PtgNumber( 0 );
		}
		double max = Double.MIN_VALUE;
		for( int i = 0; i < alloperands.length; i++ )
		{
			Object o = alloperands[i].getValue();
			try
			{
				double d = Double.MIN_VALUE;
				if( o instanceof Number )
				{
					d = ((Number) o).doubleValue();
				}
				else if( o instanceof Boolean )
				{
					d = ((Boolean) o ? 1 : 0);
				}
				else
				{
					d = new Double( o.toString() );
				}
				max = Math.max( max, d );
			}
			catch( NumberFormatException e )
			{
				// Arguments that are error values or text that cannot be translated into numbers cause errors. 
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			;
		}
		return new PtgNumber( max );
	}

	/**
	 * MEDIAN
	 * Returns the median of the given numbers
	 */
	protected static Ptg calcMedian( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg[] alloperands = PtgCalculator.getAllComponents( operands );
		CompatibleVector t = new CompatibleVector();
		double retval = 0;
		for( int i = 0; i < alloperands.length; i++ )
		{
			Ptg p = alloperands[i];
			try
			{
				Double d = new Double( String.valueOf( p.getValue() ) );
				t.addOrderedDouble( d );
			}
			catch( NumberFormatException e )
			{
			}
			;
		}

		try
		{
			Double[] dub = new Double[t.size()];
			t.toArray( dub );
			double dd = (double) t.size() % 2;
			if( ((double) t.size() % 2) == 0 )
			{
				int firstValLoc = ((t.size()) / 2) - 1;
				int lastValLoc = firstValLoc + 1;
				double firstVal = dub[firstValLoc];
				double lastVal = dub[lastValLoc];
				retval = (firstVal + lastVal) / 2;
			}
			else
			{
				// it's odd
				int firstValLoc = ((t.size() - 1) / 2);
				double firstVal = dub[firstValLoc];
				retval = firstVal;
			}
			PtgNumber pnum = new PtgNumber( retval );
			return pnum;
		}
		catch( ArrayIndexOutOfBoundsException e )
		{    // 20090701 KSC: catch exception
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}

	/**
	 * MIN
	 * Returns the smallest number in a set of values.
	 * Ignores non-number fields.  Note that it also recursivly calls itself
	 * for things like PtgRange.
	 * <p/>
	 * Usage@ MIN(number1,number2, ...)
	 * returns PtgNumber
	 */
	protected static Ptg calcMin( Ptg[] operands )
	{
		double result = java.lang.Double.MAX_VALUE;
		Double d = null;
		for( int i = 0; i < operands.length; i++ )
		{
			Ptg[] pthings = operands[i].getComponents(); // optimized -- do it once!! -jm
			if( pthings != null )
			{
				Ptg resPtg = StatisticalCalculator.calcMin( pthings );
				try
				{  // some fields may be text, so handle gracefully
					if( resPtg instanceof PtgErr )
					{
						return resPtg;    // 20090205 KSC: propagate error
					}
					if( resPtg.getValue() != null )
					{
						d = new Double( String.valueOf( resPtg.getValue() ) );
						// 20090129 KSC; if (d.doubleValue() < result || result == -1){result = d.doubleValue();} // 20070215 KSC: only access d if not null!
						if( d < result )
						{
							result = d;
						} // 20070215 KSC: only access d if not null!
					}
				}
				catch( NumberFormatException e )
				{
				}
				catch( NullPointerException e )
				{
				}// 20070209 KSC
			}
			else
			{
				Ptg p = operands[i];
				try
				{
					Object ov = p.getValue();
					if( ov != null )
					{
						if( ov.toString().equals( new PtgErr( PtgErr.ERROR_NA ).toString() ) ) // 20090205 KSC: propagate error value
						{
							return new PtgErr( PtgErr.ERROR_NA );
						}
						d = new Double( String.valueOf( ov ) );
						// 20090129 KSC; result is defaulted to max
						if( d < result )
						{
							result = d;
						} // 20070215 KSC: only access d if not null!
					}
				}
				catch( NumberFormatException e )
				{
				}
				catch( NullPointerException e )
				{
				}
			}
		}
		if( result == java.lang.Double.MAX_VALUE )    // 20090129 KSC:
		{
			result = 0;        //appears to be default in error situations
		}
		return new PtgNumber( result );
// return pnum;  
	}

	/**
	 * MINA
	 * Returns the smallest value in a list of arguments, including numbers, text, and logical values
	 * <p/>
	 * Arguments can be the following: numbers; names, arrays, or references that contain numbers; text representations of numbers; or logical values, such as TRUE and FALSE, in a reference.
	 * Logical values and text representations of numbers that you type directly into the list of arguments are counted.
	 */
	protected static Ptg calcMinA( Ptg[] operands )
	{
		Ptg[] alloperands = PtgCalculator.getAllComponents( operands );
		if( alloperands.length == 0 )
		{
			return new PtgNumber( 0 );
		}
		double min = Double.MAX_VALUE;
		for( int i = 0; i < alloperands.length; i++ )
		{
			Object o = alloperands[i].getValue();
			try
			{
				double d = Double.MAX_VALUE;
				if( o instanceof Number )
				{
					d = ((Number) o).doubleValue();
				}
				else if( o instanceof Boolean )
				{
					d = ((Boolean) o ? 1 : 0);
				}
				else
				{
					d = new Double( o.toString() );
				}
				min = Math.min( min, d );
			}
			catch( NumberFormatException e )
			{
				// Arguments that are error values or text that cannot be translated into numbers cause errors. 
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			;
		}
		return new PtgNumber( min );
	}

	/**
	 * MODE
	 * Returns the most common value in a data set
	 */
	protected static Ptg calcMode( Ptg[] operands )
	{
		Ptg[] alloperands = PtgCalculator.getAllComponents( operands );
		Vector vals = new Vector();
		Vector occurences = new Vector();
		double retval = 0;
		for( int i = 0; i < alloperands.length; i++ )
		{
			Ptg p = alloperands[i];
			try
			{
				Double d = new Double( String.valueOf( p.getValue() ) );
				if( vals.contains( d ) )
				{
					int loc = vals.indexOf( d );
					Double nums = (Double) occurences.get( loc );
					Double newnum = nums + 1;
					occurences.setElementAt( newnum, loc );
				}
				else
				{
					vals.add( d );
					occurences.add( (double) 1 );
				}
			}
			catch( NumberFormatException e )
			{
			}
			;
		}
		double biggest = 0;
		double numvalues = 0;
		for( int i = 0; i < vals.size(); i++ )
		{
			Double size = (Double) occurences.elementAt( i );
			if( size > biggest )
			{
				biggest = size;
				Double newhigh = (Double) vals.elementAt( i );
				retval = newhigh;
			}
		}
		PtgNumber pnum = new PtgNumber( retval );
		return pnum;
	}
 
 /*
NEGBINOMDIST
 Returns the negative binomial distribution
 */

	/**
	 * NORMDIST
	 * Returns the normal cumulative distribution
	 * NORMDIST(x,mean,standard_dev,cumulative)
	 * X     is the value for which you want the distribution.
	 * Mean     is the arithmetic mean of the distribution.
	 * Standard_dev     is the standard deviation of the distribution.
	 * Cumulative     is a logical value that determines the form of the function.
	 * If cumulative is TRUE, NORMDIST returns the cumulative distribution function;
	 * if FALSE, it returns the probability mass function.
	 * <p/>
	 * ********************************************************************************
	 * IMPORTANT NOTE: when Cumulative=TRUE the results are not accurate to 9 siginfiicant digits in all cases
	 * (When cumulative = TRUE, the formula is the integral from negative infinity to x of the given formula)
	 * ********************************************************************************
	 */
	protected static Ptg calcNormdist( Ptg[] operands )
	{
		if( operands.length < 4 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		try
		{
			//If mean or standard_dev is nonnumeric, NORMDIST returns the #VALUE! error value.
			double x = operands[0].getDoubleVal();
			double mean = operands[1].getDoubleVal();
			// if standard_dev ≤ 0, NORMDIST returns the #NUM! error value.
			double stddev = operands[2].getDoubleVal();
			if( stddev <= 0 )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			boolean cumulative = PtgCalculator.getBooleanValue( operands[3] );
			// If mean = 0, standard_dev = 1, and cumulative = TRUE, NORMDIST returns the standard normal distribution, NORMSDIST.
			if( (mean == 0) && (stddev == 1.0) && cumulative )
			{
				return calcNormsdist( operands );
			}

			if( !cumulative )
			{    // return the probability mass function. *** definite excel algorithm
				double a = Math.sqrt( 2 * Math.PI * Math.pow( stddev, 2 ) );
				a = 1.0 / a;
				double exp = Math.pow( x - mean, 2 );
				exp = exp / (2 * Math.pow( stddev, 2 ));
				double b = Math.exp( -exp );
				return new PtgNumber( a * b );
			}
			// When cumulative = TRUE, the formula is the integral from negative infinity to x of the given formula.
			// = the cumulative distribution function
			Ptg[] o = { new PtgNumber( (x - mean) / (stddev * Math.sqrt( 2 )) ) };
			Ptg erf = EngineeringCalculator.calcErf( o );
			double cdf = 0.5 * (1 + erf.getDoubleVal());
			return new PtgNumber( cdf );
/*			 // try this:
			 Ptg[] o= { new PtgNumber((x-mean)/(stddev))};
			 return calcNormsdist(o);
	*/
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}

	/**
	 The NORMSDIST function returns the result of the standard normal cumulative distribution function 
	 for a particular value of the random variable X. The Excel function adheres to the 
	 following mathematical approximation, P(x), of the following 
	 standard normal cumulative distribution function (CDF):

	 * P(x) = 1 -Z(x)*(b1*t+b2*t^2+b3t^3+b4t^4+b5t^5)+error(x), where

	 Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))
	 t = 1/(1+px)
	 p = 0.2316419
	 b1 = 0.319381530
	 b2 = -0.356563782
	 b3 = 1.781477937
	 b4 = -1.821255978
	 b5 = 1.330274429

	 with the following parameters:

	 abs(error(x))<7.5 * 10^-8

	 The NORMSDIST function returns the result of the standard normal 
	 CDF for a standard normal random variable Z with a mean of 0 (zero) 
	 and a standard deviation of 1. The CDF is found by taking the integral 
	 of the following standard normal probability density function

	 Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))

	 from negative infinity to the value (z) of the random variable in question. 
	 The result of the integral gives the probability that Z will occur between the 
	 values of negative infinity and z.       

	 NORMSDIST(z) must be evaluated by using an approximation procedure. 
	 Earlier versions of Excel used the same procedure for all values of z. 
	 For Excel 2003, two different approximations are used: 
	 one for |z| less than or equal to five, and a second for |z| greater than five. 
	 The two new procedures are each more accurate than the previous procedure 
	 over the range that they are applied. In earlier versions of Excel, 
	 accuracy deteriorates in the tails of the distribution yielding three significant 
	 digits for z = 4 as reported in Knusel's paper. Also, in the neighborhood of z = 1.2, 
	 NORMSDIST yields only six significant digits. However, in practice, this is likely 
	 to be sufficient for most users.

	 INFO ATP DEFINITION NORMDIST NOVEMBER 2006:
	 The NORMSDIST function returns the result of the standard normal cumulative distribution function for a particular value of the random variable X. 
	 The Microsoft Excel function adheres to the following mathematical approximation, P(x), of the standard normal CDF

	 P(x) = 1 -Z(x)*(b1*t+b2*t^2+b3t^3+b4t^4+b5t^5)+error(x), where

	 Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))
	 t = 1/(1+px)
	 p = 0.2316419
	 b1 = 0.319381530
	 b2 = -0.356563782
	 b3 = 1.781477937
	 b4 = -1.821255978
	 b5 = 1.330274429


	 with these parameters, abs(error(x))<7.5 * 10^-8.


	 In summary, if you use Excel 2002 and earlier, you should be satisfied with NORMSDIST. 
	 However, if you must have highly accurate NORMSDIST(z) values for z far from 0 
	 (such as |z| greater than or equal to four), Excel 2003 might be required. 
	 NORMSDIST(-4) = 0.0000316712; earlier versions would be accurate only as far as 0.0000317.

	 from a forum:
	 Take into consideration that Z is related to x, xm(mean) and s(std.dev.)  
	 through the expression Z = (x - xm) / s. 
	 This means that as soon as you get Z, you can proceed and calculate the integral of the 
	 CDF by using P(x) = 1 -Z(x)*(b1*t+b2*t^2+b3t^3+b4t^4+b5t^5)+error(x).

	 I wish you good code.	 

	 Some other identities that express NORMSDIST in terms of other functions
	 that have no closed form are
	 NormSDist(x) = ErfC(-x/Sqrt(2))/2 = (1-Erf(-x/Sqrt(2)))/2 for x<=0
	 NormSDist(x) = 1-ErfC(x/Sqrt(2))/2 = (1+Erf(x/Sqrt(2)))/2 for x>=0
	 NormSDist(x) = (1â€“GammaDist(x^2/2,1/2,1,TRUE))/2 for x<=0
	 NormSDist(x) = (1+GammaDist(x^2/2,1/2,1,TRUE))/2 for x>=0
	 NormSDist(x) = ChiDist(x^2,1)/2 for x<=0
	 NormSDist(x) = 1-ChiDist(x^2,1)/2 for x>=0

	 // for 2002:
	 The NORMSDIST function returns the result of the standard normal cumulative distribution 
	 function for a particular value of the random variable X. The Excel function adheres to the 
	 following mathematical approximation, P(x), of the following standard normal cumulative 
	 distribution function (CDF)

	 P(x) = 1 -Z(x)*(b1*t+b2*t^2+b3t^3+b4t^4+b5t^5)+error(x), where

	 Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))
	 t = 1/(1+px)
	 p = 0.2316419
	 b1 = 0.319381530
	 b2 = -0.356563782
	 b3 = 1.781477937
	 b4 = -1.821255978
	 b5 = 1.330274429


	 with the following parameters:
	 abs(error(x))<7.5 * 10^-8

	 The NORMSDIST function returns the result of the standard normal CDF for a standard 
	 normal random variable Z with a mean of 0 (zero) and a standard deviation of 1. The CDF 
	 is found by taking the integral of the following standard normal probability density 
	 function

	 Z(x) = (1/(sqrt(2*pi()))*exp(-x^2/2))

	 from negative infinity to the value (z) of the random variable in question. The result 
	 of the integral gives the probability that Z will occur between the values of negative 
	 infinity and z. 	  *

	 from openoffice:
	 The wrong results in NORMSDIST are due to cancellation for small negative
	 values, where gauss() is near -0.5
	 The problem can be solved in two ways:
	 (1) Use NORMSDIST(x)= 0.5*ERFC(-x/SQRT(2)). Unfortunaly ERFC is only an addin
	 function, see my issue 97091.
	 (2) Use NORMSDIST(x) 
	 = 0.5+0.5*GetLowRegIGamma(0.5,0.5*x*x) for x>=0
	 = 0.5*GetUpRegIGamma(0.5,0.5*x*x)      for x<0


	 From a forum:
	 For z less than 2, ERF = 2/SQRT (pi) * e^(-z^2) * z (1+ (2z^2)/3 + ((2z^2)^2)/15 + â€¦
	 For z greater than 2, ERF = 1- (e^(-z^2))/(SQRT(pi)) * (1/z - 1/(2z^3) + 3/(4z^5) -â€¦.)        
	 */
	/**
	 * NORMSDIST
	 * Returns the standard normal cumulative distribution
	 * <p/>
	 * NORMSDIST(z) returns the probability that the observed value of a
	 * standard normal random variable will be less than or equal to z.
	 * A standard normal random variable has mean 0 and standard deviation 1
	 * (and also variance 1 because variance = standard deviation squared).
	 * <p/>
	 * NOTE: THIS FUNCTION IS ACCURATE AS COMPARED TO EXCEL VALUES ONLY UP TO 7 SIGNIFICANT DIGITS
	 */
	protected static Ptg calcNormsdist( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		try
		{
			double result;
			double x = operands[0].getDoubleVal();
			final double b1 = 0.319381530;
			final double b2 = -0.356563782;
			final double b3 = 1.781477937;
			final double b4 = -1.821255978;
			final double b5 = 1.330274429;
			final double p = 0.2316419;
			final double c = 0.39894228;

			// below is consistently correct to at least 7 decimals using a range of test values
			if( x >= 0.0 )
			{
				double t = 1.0 / (1.0 + (p * x));
				result = (1.0 - (c * Math.exp( (-x * x) / 2.0 ) * t * ((t * ((t * ((t * ((t * b5) + b4)) + b3)) + b2)) + b1)));
			}
			else
			{
				double t = 1.0 / (1.0 - (p * x));
				result = (c * Math.exp( (-x * x) / 2.0 ) * t * ((t * ((t * ((t * ((t * b5) + b4)) + b3)) + b2)) + b1));
			}
/*	 		
			// try this one:
		 double z= (1/(Math.sqrt(2*Math.PI))*Math.exp(-Math.pow(x, 2)/2.0));
		 double t = 1/(1+p*x);
	     double e= EngineeringCalculator.calcErf(operands).getDoubleVal();
		 result = 1 -z*(b1*t+b2*Math.pow(t, 2)+b3*Math.pow(t, 3)+b4*Math.pow(t, 4)+b5*Math.pow(t, 5))+e;
*/
			BigDecimal bd = new BigDecimal( result );
			bd.setScale( 15, java.math.RoundingMode.HALF_UP );
			return new PtgNumber( bd.doubleValue() );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}

	/**
	 * NORMSINV
	 * Returns the inverse of the standard normal cumulative distribution. The distribution has a mean of zero and a standard deviation of one.
	 * <p/>
	 * Syntax
	 * NORMSINV(probability)
	 * Probability   is a probability corresponding to the normal distribution.
	 * <p/>
	 * If probability is nonnumeric, NORMSINV returns the #VALUE! error value.
	 * If probability < 0 or if probability > 1, NORMSINV returns the #NUM! error value.
	 * <p/>
	 * Because the calculation of the NORMSINV function uses a systematic search
	 * over the returned values of the NORMSDIST function, the accuracy of the
	 * NORMSDIST function is critical.
	 * <p/>
	 * Also, the search must be sufficiently refined that it "homes in" on an appropriate
	 * answer. To use the textbook Normal probability distribution table as an analogy,
	 * entries in the table must be accurate. Also, the table must contain so many entries
	 * that you can find the appropriate row of the table that yields a probability that is
	 * correct to a specific number of decimal places.	Instead, individual entries are computed
	 * on demand as the search through the "table"
	 * <p/>
	 * However, the table must be accurate and the search must continue far enough
	 * that it does not stop prematurely at an answer that has a corresponding probability
	 * (or row of the table) that is too far from the value of p that you use in the call to
	 * NORMSINV(p). Therefore, the NORMSINV function has been improved in the following ways:
	 * <p/>
	 * - The accuracy of the NORMSDIST function has been improved.
	 * - The search process has been improved to increase refinement.
	 * <p/>
	 * The NORMSDIST function has been improved in Excel 2003 and in later versions of Excel.
	 * Typically, inaccuracies in earlier versions of Excel occur for extremely small or extremely
	 * large values of p in NORMSINV(p). The values in Excel 2003 and in later versions of Excel
	 * are much more accurate.
	 * <p/>
	 * Accuracy of NORMSDIST has been improved in Excel 2003 and in later versions of Excel.
	 * In earlier versions of Excel, a single computational procedure was used for all values
	 * of z. Results were essentially accurate to 7 decimal places, more than sufficient for
	 * most practical examples.
	 * <p/>
	 * Results in earlier versions of Excel
	 * **	The accuracy of the NORMSINV function depends on two factors. Because the calculation of the NORMSINV
	 * function uses a systematic search over the returned values of the NORMSDIST function, the accuracy of
	 * the NORMSDIST function is critical.
	 * <p/>
	 * Also, the search must be sufficiently refined that it "homes in" on an appropriate answer.
	 * To use the textbook Normal probability distribution table as an analogy, entries in the table
	 * must be accurate. Also, the table 	must contain so many entries that you can find the
	 * appropriate row of the table that yields a probability that is correct to a specific number
	 * of decimal places.
	 * <p/>
	 * ' This function is a replacement for the Microsoft Excel Worksheet function NORMSINV.
	 * ' It uses the algorithm of Peter J. Acklam to compute the inverse normal cumulative
	 * ' distribution. Refer to http://home.online.no/~pjacklam/notes/invnorm/index.html for
	 * ' a description of the algorithm.
	 * ' Adapted to VB by Christian d'Heureuse, http://www.source-code.biz.
	 * Public Function NormSInv(ByVal p As Double) As Double
	 * Const a1 = -39.6968302866538, a2 = 220.946098424521, a3 = -275.928510446969
	 * Const a4 = 138.357751867269, a5 = -30.6647980661472, a6 = 2.50662827745924
	 * Const b1 = -54.4760987982241, b2 = 161.585836858041, b3 = -155.698979859887
	 * Const b4 = 66.8013118877197, b5 = -13.2806815528857, c1 = -7.78489400243029E-03
	 * Const c2 = -0.322396458041136, c3 = -2.40075827716184, c4 = -2.54973253934373
	 * Const c5 = 4.37466414146497, c6 = 2.93816398269878, d1 = 7.78469570904146E-03
	 * Const d2 = 0.32246712907004, d3 = 2.445134137143, d4 = 3.75440866190742
	 * Const p_low = 0.02425, p_high = 1 - p_low
	 * Dim q As Double, r As Double
	 * If p < 0 Or p > 1 Then
	 * Err.Raise vbObjectError, , "NormSInv: Argument out of range."
	 * ElseIf p < p_low Then
	 * q = Sqr(-2 * Log(p))
	 * NormSInv = (((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
	 * ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
	 * ElseIf p <= p_high Then
	 * q = p - 0.5: r = q * q
	 * NormSInv = (((((a1 * r + a2) * r + a3) * r + a4) * r + a5) * r + a6) * q / _
	 * (((((b1 * r + b2) * r + b3) * r + b4) * r + b5) * r + 1)
	 * Else
	 * q = Sqr(-2 * Log(1 - p))
	 * NormSInv = -(((((c1 * q + c2) * q + c3) * q + c4) * q + c5) * q + c6) / _
	 * ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
	 * End If
	 * End Function
	 * <p/>
	 * NORMSINV= NORMINV(p; 0; 1)
	 */
	private static double expm1( double x )
	{
		final double DBL_EPSILON = 0.0000001;
		double y, a = Math.abs( x );

		if( a < DBL_EPSILON )
		{
			return x;
		}
		if( a > 0.697 )
		{
			return Math.exp( x ) - 1;  /* negligible cancellation */
		}

		if( a > 1e-8 )
		{
			y = Math.exp( x ) - 1;
		}
		else /* Taylor expansion, more accurate in this range */
		{
			y = ((x / 2) + 1) * x;
		}

	    /* Newton step for solving   log(1 + y) = x   for y : */
	    /* WARNING: does not work for y ~ -1: bug in 1.5.0 -- fixed??*/
		y -= (1 + y) * (Math.log( 1 + y ) - x);
		return y;
	}

	private int R_Q_P01_check( int p, boolean log_p )
	{
		if( (log_p && (p > 0)) || (!log_p && ((p < 0) || (p > 1))) )
		{
			return 0;
		}
		return 1;
	}

	private static double quartile( double p, double mu, double sigma )
	{
		boolean lower_tail = true;
		boolean log_p = false;
		double R_D__0 = 0;
		double R_D__1 = 1;
		double R_DT_0 = 0;    //((lower_tail) ? R_D__0 : R_D__1);      /* 0 */
		double R_DT_1 = 1;    // ((lower_tail) ? R_D__1 : R_D__0)      /* 1 */

		double p_, q, r, val;
		if( p == R_DT_0 )
		{
			return Double.NEGATIVE_INFINITY;
		}
		if( p == R_DT_1 )
		{
			return Double.POSITIVE_INFINITY;
		}
		//R_Q_P01_check(p);

		if( sigma < 0 )
		{
			return 0;
		}
		if( sigma == 0 )
		{
			return mu;
		}

		p = (log_p ? (lower_tail ? Math.exp( p ) : -expm1( p )) : (lower_tail ? (p) : (1 - (p))));
		p_ = (log_p ? (lower_tail ? Math.exp( p ) : -expm1( p )) : (lower_tail ? (p) : (1 - (p))));/* real lower_tail prob. p */
		q = p_ - 0.5;

		if( Math.abs( q ) <= .425 )
		{/* 0.075 <= p <= 0.925 */
			r = .180625 - q * q;
			val = q * ((((((((((((((r * 2509.0809287301226727) + 33430.575583588128105) * r) + 67265.770927008700853) * r) + 45921.953931549871457) * r) + 13731.693765509461125) * r) + 1971.5909503065514427) * r) + 133.14166789178437745) * r) + 3.387132872796366608) / ((((((((((((((r * 5226.495278852854561) + 28729.085735721942674) * r) + 39307.89580009271061) * r) + 21213.794301586595867) * r) + 5394.1960214247511077) * r) + 687.1870074920579083) * r) + 42.313330701600911252) * r) + 1.);
		}
		else
		{ /* closer than 0.075 from {0,1} boundary */

		     /* r = min(p, 1-p) < 0.075 */
			if( q > 0 )
			{
				r = (log_p ? (lower_tail ? -expm1( p ) : Math.exp( p )) : (lower_tail ? (1 - (p)) : (p)));/* 1-p */
			}
			else
			{
				r = p_;/* = R_DT_Iv(p) ^=  p */
			}

			r = Math.sqrt( -((log_p && ((lower_tail && (q <= 0)) || (!lower_tail && (q > 0)))) ? p : /* else */ Math.log( r )) );
			    /* r = sqrt(-log(r))  <==>  min(p, 1-p) = exp( - r^2 ) */

			if( r <= 5. )
			{ /* <==> min(p,1-p) >= exp(-25) ~= 1.3888e-11 */
				r += -1.6;
				val = ((((((((((((((r * 7.7454501427834140764e-4) + .0227238449892691845833) * r) + .24178072517745061177) * r) + 1.27045825245236838258) * r) + 3.64784832476320460504) * r) + 5.7694972214606914055) * r) + 4.6303378461565452959) * r) + 1.42343711074968357734) / ((((((((((((((r * 1.05075007164441684324e-9) + 5.475938084995344946e-4) * r) + .0151986665636164571966) * r) + .14810397642748007459) * r) + .68976733498510000455) * r) + 1.6763848301838038494) * r) + 2.05319162663775882187) * r) + 1.);
			}
			else
			{ /* very close to  0 or 1 */
				r += -5.;
				val = ((((((((((((((r * 2.01033439929228813265e-7) + 2.71155556874348757815e-5) * r) + .0012426609473880784386) * r) + .026532189526576123093) * r) + .29656057182850489123) * r) + 1.7848265399172913358) * r) + 5.4637849111641143699) * r) + 6.6579046435011037772) / ((((((((((((((r * 2.04426310338993978564e-15) + 1.4215117583164458887e-7) * r) + 1.8463183175100546818e-5) * r) + 7.868691311456132591e-4) * r) + .0148753612908506148525) * r) + .13692988092273580531) * r) + .59983220655588793769) * r) + 1.);
			}

			if( q < 0.0 )
			{
				val = -val;
			}
			    /* return (q >= 0.)? r : -r ;*/
		}
		return mu + (sigma * val);
	}

	public static Ptg calcNormInv( Ptg[] operands )
	{
		try
		{

			double p = operands[0].getDoubleVal();
			if( (p < 0) || (p > 1) )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			double mean = operands[1].getDoubleVal();
			double stddev = operands[2].getDoubleVal();
			if( stddev <= 0 )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
			// If mean = 0 and standard_dev = 1, NORMINV uses the standard normal inverse (see NORMSINV).
			double result = quartile( p, mean, stddev );
			return new PtgNumber( result );
		}
		catch( Exception e )
		{
			;
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	public static Ptg calcNormsInv( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getValueError();
		}
		try
		{
			double x = operands[0].getDoubleVal();
			if( (x < 0) || (x > 1) )
			{
				return new PtgErr( PtgErr.ERROR_NUM );
			}
		 /*
		  * the algorithm is supposed to iterate over NORMSDIST values using Newton-Raphson's approximation
		  * Newton-Raphson uses an iterative process to approach one root of a function (i.e the zero of the function or 
		  * where the function = 0) 
		  * Newton-Raphson is in the form of:
		  * 
		  * Xn+1= Xn- (f(xn)/f'(xn)
		  * where Xn is the current known X value, f(xn) is the value of the function at X, f'(Xn) is the derivative or slope
		  * at X, Xn+1 is the next X value. Essentially, f'(xn) represents (f(x)/delta x) so f(xn)/f'(xn)== delta x.
		  * the more iterations we run over, the closer delta x will be to 0.
		  * 
		  * The Newton-Raphson method does not always work, however. It runs into problems in several places. 
			What would happen if we chose an initial x-value of x=0? We would have a "division by zero" error, and would not be able to proceed. 
			You may also consider operating the process on the function f(x) = x1/3, using an inital x-value of x=1. 
			Do the x-values converge? Does the delta-x decrease toward zero (0)?
			
			 is the derivative of the standard normal distribution = the standard probability density function???
		  */
/* below is not accurate enough - but N-R approximation is impossible ((:*/
			// Coefficients in rational approximations
			double[] a = new double[]{
					-3.969683028665376e+01,
					2.209460984245205e+02,
					-2.759285104469687e+02,
					1.383577518672690e+02,
					-3.066479806614716e+01,
					2.506628277459239e+00
			};

			double[] b = new double[]{
					-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02, 6.680131188771972e+01, -1.328068155288572e+01
			};

			double[] c = new double[]{
					-7.784894002430293e-03,
					-3.223964580411365e-01,
					-2.400758277161838e+00,
					-2.549732539343734e+00,
					4.374664141464968e+00,
					2.938163982698783e+00
			};

			double[] d = new double[]{
					7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00, 3.754408661907416e+00
			};

			// Define break-points.
			double plow = 0.02425;
			double phigh = 1 - plow;
			double result;
			// Rational approximation for lower region:
			if( x < plow )
			{
				double q = Math.sqrt( -2 * Math.log( x ) );
				BigDecimal r = new BigDecimal( ((((((((((c[0] * q) + c[1]) * q) + c[2]) * q) + c[3]) * q) + c[4]) * q) + c[5]) / ((((((((d[0] * q) + d[1]) * q) + d[2]) * q) + d[3]) * q) + 1) );
				r.setScale( 15, java.math.RoundingMode.HALF_UP );
				return new PtgNumber( r.doubleValue() );
			}

			// Rational approximation for upper region:
			if( phigh < x )
			{
				double q = Math.sqrt( -2 * Math.log( 1 - x ) );
				BigDecimal r = new BigDecimal( -((((((((((c[0] * q) + c[1]) * q) + c[2]) * q) + c[3]) * q) + c[4]) * q) + c[5]) / ((((((((d[0] * q) + d[1]) * q) + d[2]) * q) + d[3]) * q) + 1) );
				r.setScale( 15, java.math.RoundingMode.HALF_UP );
				return new PtgNumber( r.doubleValue() );
			}

			// Rational approximation for central region:
			double q = x - 0.5;
			double r = q * q;
			BigDecimal rr = new BigDecimal( (((((((((((a[0] * r) + a[1]) * r) + a[2]) * r) + a[3]) * r) + a[4]) * r) + a[5]) * q) / ((((((((((b[0] * r) + b[1]) * r) + b[2]) * r) + b[3]) * r) + b[4]) * r) + 1) );

			rr.setScale( 15, java.math.RoundingMode.HALF_UP );
			return new PtgNumber( rr.doubleValue() );
		}
		catch( Exception e )
		{
			return PtgCalculator.getValueError();
		}
	}

	/**
	 * PEARSON
	 * Returns the Pearson product moment correlation coefficient
	 *
	 * @throws CalculationException
	 */
	public static Ptg calcPearson( Ptg[] operands ) throws CalculationException
	{
		return calcCorrel( operands );
	}
 /*
PERCENTILE
 Returns the k-th percentile of values in a range
 
PERCENTRANK
 Returns the percentage rank of a value in a data set
 
PERMUT
 Returns the number of permutations for a given number of objects
 
POISSON
 Returns the Poisson distribution
 
PROB
 Returns the probability that values in a range are between two limits
 */

	/**
	 * QUARTILE
	 * Returns the quartile of a data set
	 */
	protected static Ptg calcQuartile( Ptg[] operands )
	{
		Ptg[] aveoperands = new Ptg[1];
		aveoperands[0] = operands[0];
		Ptg[] allVals = PtgCalculator.getAllComponents( aveoperands );
		CompatibleVector t = new CompatibleVector();
		double retval = 0;
		for( int i = 0; i < allVals.length; i++ )
		{
			Ptg p = allVals[i];
			try
			{
				Double d = new Double( String.valueOf( p.getValue() ) );
				t.addOrderedDouble( d );
			}
			catch( NumberFormatException e )
			{
				Logger.logErr( e );
			}
			;
		}

		Double[] dub = new Double[t.size()];
		t.toArray( dub );
		Integer quart;
		Object o = operands[1].getValue();
		if( o instanceof Integer )
		{
			quart = (Integer) operands[1].getValue();
		}
		else
		{
			quart = ((Double) operands[1].getValue()).intValue();
		}

		float quartile = quart.floatValue();
		if( quart == 0 )
		{    // return minimum value
			return new PtgNumber( dub[0] );
		}
		if( quart == 4 )
		{    // return maximum value
			return new PtgNumber( dub[t.size() - 1] );
		}
		if( (quart > 4) || (quart < 0) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		// find the kth smallest
		float kk = (float) (quartile / 4);
		kk = (dub.length - 1) * kk;
		kk++;
		// truncate k, but keep the remainder.
		int k = -1;
		float remainder = 0;
		if( (kk % 1) != 0 )
		{
			remainder = kk % 1;
			String s = String.valueOf( kk );
			String ss = s.substring( s.indexOf( "." ), s.length() );
			ss = "0" + ss;
			remainder = new Float( ss );
			s = s.substring( 0, s.indexOf( "." ) );
			k = Integer.valueOf( String.valueOf( s ) );
		}
		else
		{
			k = (int) kk / 1;
		}
		if( k >= dub.length )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		double firstVal = dub[k - 1];
		double secondVal = dub[k];
		double output = firstVal + (remainder * (secondVal - firstVal));
		PtgNumber pn = new PtgNumber( output );
		return pn;
	}

	/**
	 * RANK
	 * Returns the rank of a number in a list of numbers
	 * <p/>
	 * RANK(number,ref,order)
	 * Number   is the number whose rank you want to find.
	 * Ref   is an array of, or a reference to, a list of numbers. Nonnumeric values in ref are ignored.
	 * Order   is a number specifying how to rank number.
	 * <p/>
	 * If order is 0 (zero) or omitted, Microsoft Excel ranks number as if ref were a list sorted in descending order.
	 * If order is any nonzero value, Microsoft Excel ranks number as if ref were a list sorted in ascending order.
	 */
	protected static Ptg calcRank( Ptg[] operands )
	{
		// the number
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg num = operands[0];
		Double theNum = null;
		try
		{
			Object o = num.getValue();
			if( o.equals( "" ) )
			{
				theNum = 0.0d;
			}
			else
			{
				theNum = new Double( o.toString() );
			}
		}
		catch( NumberFormatException nfm )
		{
			return new PtgErr();
		}

		//ascending or decending?
		boolean ascending = true;
		if( operands.length < 3 )
		{
			ascending = false;
		}
		else if( operands[2] instanceof PtgMissArg )
		{
			ascending = false;
		}
		else
		{
			PtgInt order = (PtgInt) operands[2];
			int i = order.getVal();
			if( i == 0 )
			{
				ascending = false;
			}
		}
		Ptg[] aveoperands = new Ptg[1];
		aveoperands[0] = operands[1];
		Ptg[] refs = PtgCalculator.getAllComponents( aveoperands );
		CompatibleVector retList = new CompatibleVector();
		double retval = 0;
		for( int i = 0; i < refs.length; i++ )
		{
			Ptg p = refs[i];
			try
			{
				Double d = new Double( String.valueOf( p.getValue() ) );
				retList.addOrderedDouble( d );
			}
			catch( NumberFormatException e )
			{
			}
			;
		}
		Double[] dubRefs = new Double[retList.size()];
		if( ascending )
		{
			retList.toArray( dubRefs );
		}
		else
		{
			for( int i = 0; i < dubRefs.length; i++ )
			{
				dubRefs[i] = (Double) retList.last();
				retList.remove( retList.size() - 1 );
			}
		}
		int res = -1;
		for( int i = 0; i < dubRefs.length; i++ )
		{
			if( dubRefs[i].toString().equalsIgnoreCase( theNum.toString() ) )
			{
				res = i + 1;
				i = dubRefs.length;
			}
		}
		if( res == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		return new PtgInt( res );

	}

	/**
	 * RSQ
	 * Returns the square of the Pearson product moment correlatin coefficient
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcRsq( Ptg[] operands ) throws CalculationException
	{
		PtgNumber p = (PtgNumber) calcPearson( operands );
		double d = p.getVal();
		d = (d * d);
		return new PtgNumber( d );
	}
 /*
SKEW
 Returns the skewness of a distribution
 */

	/**
	 * SLOPE
	 * Returns the slope of the linear regression line
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcSlope( Ptg[] operands ) throws CalculationException
	{
		if( operands.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		double[] yvals = PtgCalculator.getDoubleValueArray( operands[0] );
		double[] xvals = PtgCalculator.getDoubleValueArray( operands[1] );
		if( (xvals == null) || (yvals == null) )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		double sumXVals = 0;
		for( int i = 0; i < xvals.length; i++ )
		{
			sumXVals += xvals[i];
		}
		double sumYVals = 0;
		for( int i = 0; i < yvals.length; i++ )
		{
			sumYVals += yvals[i];
		}
		double sumXYVals = 0;
		for( int i = 0; i < yvals.length; i++ )
		{
			sumXYVals += xvals[i] * yvals[i];
		}
		double sqrXVals = 0;
		for( int i = 0; i < xvals.length; i++ )
		{
			sqrXVals += xvals[i] * xvals[i];
		}
		double toparg = (sumXVals * sumYVals) - (sumXYVals * yvals.length);
		double bottomarg = (sumXVals * sumXVals) - (sqrXVals * xvals.length);
		double res = toparg / bottomarg;
		return new PtgNumber( res );
	}

	/**
	 * SMALL
	 * Returns the k-th smallest value in a data set
	 * <p/>
	 * SMALL(array,k)
	 * Array   is an array or range of numerical data for which you want to determine the k-th smallest value.
	 * K   is the position (from the smallest) in the array or range of data to return.
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcSmall( Ptg[] operands ) throws CalculationException
	{
		if( operands.length != 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg rng = operands[0];
		Ptg[] array = PtgCalculator.getAllComponents( rng );
		if( array.length == 0 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int k = new Double( PtgCalculator.getDoubleValueArray( operands[1] )[0] ).intValue();
		if( (k <= 0) || (k > array.length) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		CompatibleVector sortedValues = new CompatibleVector();
		for( int i = 0; i < array.length; i++ )
		{
			Ptg p = array[i];
			try
			{
				Double d = new Double( String.valueOf( p.getValue() ) );
				sortedValues.addOrderedDouble( d );
			}
			catch( NumberFormatException e )
			{
			}
			;
		}
		try
		{
			return new PtgNumber( (Double) sortedValues.get( k - 1 ) );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}
 	
 /*
STANDARDIZE
 Returns a normalized value
 */

	/**
	 * STDEV(number1,number2, ...)
	 * <p/>
	 * Number1,number2, ...   are 1 to 255 number arguments corresponding to a sample of a population.
	 * You can also use a single array or a reference to an array instead of arguments separated by commas.
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcStdev( Ptg[] operands ) throws CalculationException
	{
		double[] allVals = PtgCalculator.getDoubleValueArray( operands );
		if( allVals == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		double sqrDev = 0;
		for( int i = 0; i < allVals.length; i++ )
		{
			PtgNumber p = (PtgNumber) calcAverage( operands );
			double ave = p.getVal();
			sqrDev += Math.pow( (allVals[i] - ave), 2 );
		}
		double retval = Math.sqrt( sqrDev / (allVals.length - 1) );
		return new PtgNumber( retval );
	}
 /*
STDEVA
 Estimates standard deviation based on a sample, including numbers, text, and logical values
 
STDEVP
 Calculates standard deviation based on the entire population
 
STDEVPA
 Calculates standard deviation based on the entire population, including numbers, text, and logical values
 */

	/**
	 * STEYX
	 * Returns the standard error of the predicted y-value for each x in the regression
	 *
	 * @throws CalculationException
	 */
	public static Ptg calcSteyx( Ptg[] operands ) throws CalculationException
	{
		Ptg[] arr = new Ptg[1];
		arr[0] = operands[0];
		PtgNumber pn = (PtgNumber) calcVarp( arr );
		double yVarp = pn.getVal();
		arr[0] = operands[1];
		pn = (PtgNumber) calcVarp( arr );
		double xVarp = pn.getVal();
		double[] y = PtgCalculator.getDoubleValueArray( operands[0] );
		if( y == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		yVarp *= y.length;
		xVarp *= y.length;
		pn = (PtgNumber) calcSlope( operands );
		double slope = pn.getVal();
		double retval = yVarp - ((slope * slope) * xVarp);
		retval = retval / (y.length - 2);
		retval = Math.sqrt( retval );
		return new PtgNumber( retval );
	}
 /*
TDIST
 Returns the Student's t-distribution
 
TINV
 Returns the inverse of the Student's t-distribution
 */

	/**
	 * TREND
	 * Returns values along a linear trend
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcTrend( Ptg[] operands ) throws CalculationException
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
// KSC: THIS FUNCTION DOES NOT WORK AS EXPECTED: TODO: FIX!

		if( true )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
// 	Ptg[] forecast = new Ptg[3];
		Ptg[] forecast = new Ptg[2];
		forecast[0] = operands[0];
		if( operands.length > 1 )
		{
			forecast[1] = operands[1];
		}
// TODO: 	
// 	else // If known_x's is omitted, it is assumed to be the array {1,2,3,...} that is the same size as known_y's.
		Ptg[] newXs;
		if( operands.length > 2 )
		{
			newXs = PtgCalculator.getAllComponents( operands[2] );
		}
		else
		{
			newXs = PtgCalculator.getAllComponents( operands[1] );
		}

		String retval = "";
		for( int i = 0; i < newXs.length; i++ )
		{
			//forecast[0] = newXs[i];
			PtgNumber p = (PtgNumber) calcForecast( forecast );
			double forcst = p.getVal();
			retval += "{" + String.valueOf( forcst ) + "},";
		}
		// get rid of trailing comma
		retval = retval.substring( 0, retval.length() - 1 );
		PtgArray pa = new PtgArray();
		pa.setVal( retval );
		return pa;
	}
 
 /*
TRIMMEAN
 Returns the mean of the interior of a data set
 
TTEST
 Returns the probability associated with a Student's t-Test
 */

	/**
	 * VAR
	 * Estimates variance based on a sample
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcVar( Ptg[] operands ) throws CalculationException
	{
		double[] allVals = PtgCalculator.getDoubleValueArray( operands );
		if( allVals == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		double sqrDev = 0;
		for( int i = 0; i < allVals.length; i++ )
		{
			PtgNumber p = (PtgNumber) calcAverage( operands );
			double ave = p.getVal();
			sqrDev += Math.pow( (allVals[i] - ave), 2 );
		}
		double retval = (sqrDev / (allVals.length - 1));
		return new PtgNumber( retval );
	}
 /*
VARA
 Estimates variance based on a sample, including numbers, text, and logical values
 */

	/**
	 * VARp
	 * Estimates variance based on a full population
	 *
	 * @throws CalculationException
	 */
	protected static Ptg calcVarp( Ptg[] operands ) throws CalculationException
	{
		double[] allVals = PtgCalculator.getDoubleValueArray( operands );
		if( allVals == null )
		{
			return new PtgErr( PtgErr.ERROR_NA );//20090130 KSC: propagate error
		}
		double sqrDev = 0;
		for( int i = 0; i < allVals.length; i++ )
		{
			PtgNumber p = (PtgNumber) calcAverage( operands );
			double ave = p.getVal();
			sqrDev += Math.pow( (allVals[i] - ave), 2 );
		}
		double retval = (sqrDev / allVals.length);
		return new PtgNumber( retval );
	}
 /*
VARPA
 Calculates variance based on the entire population, including numbers, text, and logical values
 
WEIBULL
 Returns the Weibull distribution
 
ZTEST
 Returns the two-tailed P-value of a z-test
 
*/

}