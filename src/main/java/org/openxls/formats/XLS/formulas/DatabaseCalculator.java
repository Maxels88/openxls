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

import org.openxls.formats.XLS.ReferenceTracker;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Vector;

/**
 * DatabaseCalculator is a collection of static methods that operate
 * as the Microsoft Excel function calls do.
 * <p/>
 * All methods are called with an array of ptg's, which are then
 * acted upon.  A Ptg of the type that makes sense (ie boolean, number)
 * is returned after each method.
 * <p/>
 * <p/>
 * Database and List Management functions
 * Microsoft Excel includes worksheet functions that analyze data stored in lists or databases.
 * Each of these functions, referred to collectively as the Dfunctions, uses three arguments: database, field, and criteria.
 * <p/>
 * These arguments refer to the worksheet ranges that are used by the function.
 * <p/>
 * DAVERAGE   Returns the average of selected database entries
 * <p/>
 * DCOUNT   Counts the cells that contain numbers in a database
 * <p/>
 * DCOUNTA   Counts nonblank cells in a database
 * <p/>
 * DGET   Extracts from a database a single record that matches the specified criteria
 * <p/>
 * DMAX   Returns the maximum value from selected database entries
 * <p/>
 * DMIN   Returns the minimum value from selected database entries
 * <p/>
 * DPRODUCT   Multiplies the values in a particular field of records that match the criteria in a database
 * <p/>
 * DSTDEV   Estimates the standard deviation based on a sample of selected database entries
 * <p/>
 * DSTDEVP   Calculates the standard deviation based on the entire population of selected database entries
 * <p/>
 * DSUM   Adds the numbers in the field column of records in the database that match the criteria
 * <p/>
 * DVAR   Estimates variance based on a sample from selected database entries
 * <p/>
 * DVARP   Calculates variance based on the entire population of selected database entries
 * <p/>
 * GETPIVOTDATA   Returns data stored in a PivotTable
 * <p/>
 * <p/>
 * ABOUT DB
 * <p/>
 * All Database Formulas take 3 arguments:
 * Database is the range of cells that makes up the list or database.
 * <p/>
 * A database is a list of related data in which rows of related information
 * are records, and columns of data are fields.
 * <p/>
 * The first row of the list contains labels for each column.
 * Field   indicates which column is used in the function.
 * Field can be given as text with the column label
 * enclosed between double quotation marks, such as "Age" or "Yield,"
 * or as a number that represents the position of the column within the
 * list: 1 for the first column, 2 for the second column, and so on.
 * <p/>
 * Criteria   is the range of cells that contains the conditions you specify.
 * <p/>
 * You can use any range for the criteria argument,  * 	as long as it
 * includes at least one column label and at least one cell below the column
 * label for specifying a condition for the column.
 * <p/>
 * Make sure the criteria range does not overlap the list.
 * <p/>
 * <p/>
 * <p/>
 * TODO To change the template for this generated type comment go to
 * Window - Preferences - Java - Code Style - Code Templates
 */
public class DatabaseCalculator
{
	/**
	 * Fetch a DB from the cache or create a new one
	 * <p/>
	 * <p/>
	 * Dbs store Cell refs...
	 *
	 * @param operands
	 * @return
	 */
	private static DB getDb( Ptg operands )
	{
		ReferenceTracker DBcache = operands.getParentRec().getWorkBook().getRefTracker();

		//gonna try never caching this... painful, but if we dont'
		if( DBcache.getListDBs().get( operands.toString() ) != null )
		{
			//Logger.logInfo("getDB: " + operands.toString()+ "using cache.");
			return (DB) DBcache.getListDBs().get( operands.toString() );
		}
		//}
		// create new
		//log.error("getDB: " + operands.toString()+ "NOT cached.");
		Ptg[] dbrange = PtgCalculator.getAllComponents( operands );
		DB ret = DB.parseList( dbrange );
		DBcache.getListDBs().put( operands.toString(), ret );
		return ret;
	}

	private static Criteria getCriteria( Ptg operands )
	{
		ReferenceTracker DBcache = operands.getParentRec().getWorkBook().getRefTracker();

		// test without cache
		if( DBcache.getCriteriaDBs().get( operands.toString() ) != null )
		{
			//Logger.logInfo("getCriteria: " + operands.toString()+ "using cache.");
			return (Criteria) DBcache.getCriteriaDBs().get( operands.toString() );
		}
		//log.error("getCriteria: " + operands.toString()+ "NOT cached.");
		Ptg[] criteria = PtgCalculator.getAllComponents( operands );
		Criteria ret = Criteria.parseCriteria( criteria );
		DBcache.getCriteriaDBs().put( operands.toString(), ret );
		return ret;
	}

	/**
	 * DAVERAGE   Returns the average of selected database entries
	 */
	protected static Ptg calcDAverage( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		double average = 0;
		int count = 0;

		// this is the colname to match
		String colname = operands[1].getValue().toString();
		for( int i = 0; i < db.rows.length; i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			Ptg[] rwz = db.getRow( i ); // slight optimization one less call to getRow -jm
			if( crit.passes( colname, rwz, db ) )
			{
				// passes; now do action
				Ptg vx = rwz[fNum];
				if( vx != null )
				{
					try
					{
						average += Double.parseDouble( vx.getValue().toString() );
						count++;    // if it can be parsed into a number, increment count
					}
					catch( NumberFormatException exp )
					{
					}
				}
			}
		}
		if( count > 0 )
		{
			average = average / count;
		}
		return new PtgNumber( average );
	}

	/**
	 * DCOUNT   Counts the cells that contain numbers in a database
	 */
	protected static Ptg calcDCount( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		int count = 0;
		int nrow = db.getNRows();
		// this is the colname to match
		String colname = operands[1].getValue().toString();
		for( int i = 0; i < nrow; i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			try
			{
				Ptg[] rr = db.getRow( i );
				
				/* "passes" means that there is a matching
				 * cell in the row of the data db cells
				 * 
				 */
				if( crit.passes( colname, rr, db ) )
				{
					// passes; now do action
					Ptg cx = db.getCell( i, fNum );
					String vtx = cx.getValue().toString();
					if( vtx != null )
					{
						Double.parseDouble( vtx );
						count++;    // if it can be parsed into a number, increment count
					}
				}
			}
			catch( NumberFormatException e )
			{
			}
		}
		return new PtgNumber( count );
	}

	/**
	 * DCOUNTA   Counts nonblank cells in a database
	 */
	protected static Ptg calcDCountA( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int count = 0;

		// this is the colname to match
		String colname = operands[1].getValue().toString();
		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			if( crit.passes( colname, db.getRow( i ), db ) )
			{
				// passes; now do action
				String s = db.getCell( i, fNum ).getValue().toString();
				if( (s != null) && !s.trim().equals( "" ) )
				{
					count++;    // if field is not blank, increment count
				}
			}
		}

		return new PtgNumber( count );
	}

	/**
	 * DGET   Extracts from a database a single record that matches the specified criteria
	 */
	protected static Ptg calcDGet( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NULL );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		String val = "";
		int count = 0;

		// this is the colname to match
		String colname = operands[1].getValue().toString();
		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			if( crit.passes( colname, db.getRow( i ), db ) )
			{
				// passes; now do action
				val = db.getCell( i, fNum ).getValue().toString();
				count++;
			}
		}
		if( count == 0 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );    // no recs match
		}
		if( count > 1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM ); // if more than one record matches criteria
		}
		return new PtgStr( val );
	}

	/**
	 * DMAX   Returns the maximum value from selected database entries
	 */
	protected static Ptg calcDMax( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		double max = Double.MIN_VALUE;

		// this is the colname to match
		String colname = operands[1].getValue().toString();
		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			if( crit.passes( colname, db.getRow( i ), db ) )
			{
				// passes; now do action
				String vtx = db.getCell( i, fNum ).getValue().toString();
				if( vtx != null )
				{
					try
					{
						if( vtx.length() > 0 )
						{
							max = Math.max( max, Double.parseDouble( vtx ) );
						}
					}
					catch( NumberFormatException exp )
					{
					}
				}
			}
		}

		if( max == Double.MIN_VALUE )
		{
			max = 0;
		}
		return new PtgNumber( max );
	}

	/**
	 * DMIN   Returns the minimum value from selected database entries
	 */
	protected static Ptg calcDMin( Ptg[] operands )
	{
		if( operands.length != 3 )  // sanity checks
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		double min = Double.MAX_VALUE;
		// this is the colname to match
		String colnamx = operands[1].getValue().toString();

		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			try
			{
				Ptg[] rwz = db.getRow( i );

				if( crit.passes( colnamx, rwz, db ) )
				{
					// passes; now do action
					try
					{
						Ptg dbx = db.getCell( i, fNum );

						if( dbx != null )
						{
							String dnb = dbx.getValue().toString();
							if( dnb != null )
							{
								if( dnb.length() > 0 )
								{
									min = Math.min( min, Double.parseDouble( dnb ) );
								}
							}
						}
					}
					catch( Exception ex )
					{
						; // normal blanks etc.
					}
				}
			}
			catch( NumberFormatException e )
			{
			}
		}
		if( min == Double.MAX_VALUE )
		{
			min = 0;
		}
		return new PtgNumber( min );
	}

	/**
	 * DPRODUCT   Multiplies the values in a particular field of records that match the criteria in a database
	 */
	protected static Ptg calcDProduct( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		double product = 1;
		// this is the colname to match
		String colname = operands[1].getValue().toString();

		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			try
			{
				if( crit.passes( colname, db.getRow( i ), db ) )
				{
					// passes; now do action
					String fnx = db.getCell( i, fNum ).getValue().toString();
					if( fnx != null )
					{
						product *= Double.parseDouble( fnx );
					}
				}
			}
			catch( NumberFormatException e )
			{
			}
		}
		return new PtgNumber( product );
	}

	/**
	 * DSTDEV   Estimates the standard deviation based on a sample of selected database entries
	 */
	protected static Ptg calcDStdDev( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		java.util.ArrayList vals = new java.util.ArrayList();
		double sum = 0;
		int count = 0;
		// this is the colname to match
		String colname = operands[1].getValue().toString();

		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			if( crit.passes( colname, db.getRow( i ), db ) )
			{
				// passes; now do action
				try
				{
					String fnx = db.getCell( i, fNum ).getValue().toString();
					if( fnx != null )
					{
						double x = Double.parseDouble( fnx );
						sum += x;
						count++;
						vals.add( Double.toString( x ) );
					}
				}
				catch( NumberFormatException e )
				{
				}
			}
		}
		double stdev = 0;
		if( count > 0 )
		{
			double average = sum / count;
			// now have all values in vals
			for( int i = 0; i < count; i++ )
			{
				double x = Double.parseDouble( (String) vals.get( i ) );
				stdev += Math.pow( (x - average), 2 );
			}
			if( count > 1 )
			{
				count--;
			}
			stdev = Math.sqrt( stdev / count );
		}
		return new PtgNumber( stdev );
	}

	/**
	 * DSTDEVP   Calculates the standard deviation based on the entire population of selected database entries
	 */
	protected static Ptg calcDStdDevP( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		java.util.ArrayList vals = new java.util.ArrayList();
		double sum = 0;
		int count = 0;
		// this is the colname to match
		String colname = operands[1].getValue().toString();

		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			if( crit.passes( colname, db.getRow( i ), db ) )
			{
				// passes; now do action
				try
				{
					String fnx = db.getCell( i, fNum ).getValue().toString();
					if( fnx != null )
					{
						double x = Double.parseDouble( fnx );
						sum += x;
						count++;
						vals.add( Double.toString( x ) );
					}
				}
				catch( NumberFormatException e )
				{
				}
			}
		}
		double stdevp = 0;
		if( count > 0 )
		{
			double average = sum / count;
			// now have all values in vals
			for( int i = 0; i < count; i++ )
			{
				double x = Double.parseDouble( (String) vals.get( i ) );
				stdevp += Math.pow( (x - average), 2 );
			}
			stdevp = Math.sqrt( stdevp / count );
		}
		return new PtgNumber( stdevp );
	}

	/**
	 * DSUM   Adds the numbers in the field column of records in the database that match the criteria
	 */
	protected static Ptg calcDSum( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}

		int count = 0;
		double sum = 0.0d;
		int nrow = db.getNRows();
		// this is the colname to match
		String colname = operands[1].getValue().toString();
		for( int i = 0; i < nrow; i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			try
			{
				Ptg[] rr = db.getRow( i );
				
				/* "passes" means that there is a matching
				 * cell in the row of the data db cells
				 * 
				 */
				if( crit.passes( colname, rr, db ) )
				{
					// passes; now do action
					try
					{
						String fnx = db.getCell( i, fNum ).getValue().toString();
						if( fnx != null )
						{
							sum += Double.parseDouble( fnx );
						}
					}
					catch( NumberFormatException e )
					{
					}
				}
			}
			catch( NumberFormatException e )
			{
			}
		}
		return new PtgNumber( sum );
	}

	/**
	 * DVAR   Estimates variance based on a sample from selected database entries
	 */
	protected static Ptg calcDVar( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		java.util.ArrayList vals = new java.util.ArrayList();
		double sum = 0;
		int count = 0;

		// this is the colname to match
		String colname = operands[1].getValue().toString();
		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			if( crit.passes( colname, db.getRow( i ), db ) )
			{
				// passes; now do action
				try
				{
					String fnx = db.getCell( i, fNum ).getValue().toString();
					if( fnx != null )
					{
						double x = Double.parseDouble( db.getCell( i, fNum ).toString() );
						sum += x;
						count++;
						vals.add( Double.toString( x ) );
					}
				}
				catch( NumberFormatException e )
				{
				}
			}
		}
		double variance = 0;
		if( count > 0 )
		{
			double average = sum / count;
			// now have all values in vals
			for( int i = 0; i < count; i++ )
			{
				double x = Double.parseDouble( (String) vals.get( i ) );
				variance += Math.pow( (x - average), 2 );
			}
			if( count > 1 )
			{
				count--;
			}
			variance = variance / count;
		}
		return new PtgNumber( variance );
	}

	/* DVARP   Calculates variance based on the entire population of selected database entries
	 */
	protected static Ptg calcDVarP( Ptg[] operands )
	{
		if( operands.length != 3 )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}

		DB db = getDb( operands[0] );
		Criteria crit = getCriteria( operands[2] );
		if( (db == null) || (crit == null) )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		int fNum = db.findCol( operands[1].getString().trim() );
		if( fNum == -1 )
		{
			return new PtgErr( PtgErr.ERROR_NUM );
		}
		java.util.ArrayList vals = new java.util.ArrayList();
		double sum = 0;
		int count = 0;

		// this is the colname to match
		String colname = operands[1].getValue().toString();
		for( int i = 0; i < db.getNRows(); i++ )
		{    // loop thru all db rows
			// check if current row passes criteria requirements
			if( crit.passes( colname, db.getRow( i ), db ) )
			{
				// passes; now do action
				try
				{
					String fnx = db.getCell( i, fNum ).getValue().toString();
					if( fnx != null )
					{
						double x = Double.parseDouble( fnx );
						sum += x;
						count++;
						vals.add( Double.toString( x ) );
					}
				}
				catch( NumberFormatException e )
				{
				}
			}
		}
		double varP = 0;
		if( count > 0 )
		{
			double average = sum / count;
			// now have all values in vals
			for( int i = 0; i < count; i++ )
			{
				double x = Double.parseDouble( (String) vals.get( i ) );
				varP += Math.pow( (x - average), 2 );
			}
			varP = varP / count;
		}
		return new PtgNumber( varP );
	}
/* GETPIVOTDATA   Returns data stored in a PivotTable
 *
 *
 */
}

/**
 * EXPLANATION of Database Formulas
 * <p/>
 * Database   is the range of cells that makes up the list or database.
 * A database is a list of related data in which rows of related
 * information are records, and columns of data are fields.
 * The first row of the list contains labels for each column.
 * <p/>
 * Field   indicates which column is used in the function.
 * Field can be given as text with the column label enclosed between
 * double quotation marks, such as "Age" or "Yield," or as a number
 * that represents the position of the column within the list:
 * <p/>
 * 1 for the first column,
 * 2 for the second column, and so on.
 * <p/>
 * NOTE: quotes around field text is optional:
 * first row of columns and rows are field labels.
 * <p/>
 * Criteria   is the range of cells that contains the conditions
 * you specify. You can use any range for the criteria argument,
 * as long as it includes at least one column label and at least one
 * cell below the column label for specifying a condition for the column.
 * <p/>
 * Example
 */
class DB
{
	protected String[] colHeaders;
	protected Ptg[][] rows;

	public DB( int nCols, int nRows )
	{
		colHeaders = new String[nCols];
		// TODO: replace with PtgRef array!
		rows = new Ptg[nRows][nCols];
	}

	public int getNCols()
	{
		return colHeaders.length;
	}

	public int getNRows()
	{
		return rows.length;
	}

	/**
	 * return the index of the col in the DB
	 *
	 * @param cname
	 * @return
	 */
	public int getCol( String cname )
	{
		for( int t = 0; t < colHeaders.length; t++ )
		{
			if( colHeaders[t].equalsIgnoreCase( cname ) )
			{
				return t;
			}
		}
		return -1;
	}

	/**
	 * return a row of DB ptgs
	 *
	 * @param i
	 * @return
	 */
	public Ptg[] getRow( int i )
	{
		if( (i > -1) && (i < rows.length) )
		{
			return rows[i];
		}
		return null;
	}

	/**
	 * return a col of ??
	 *
	 * @param i
	 * @return
	 */
	public String getCol( int i )
	{
		if( (i > -1) && (i < colHeaders.length) )
		{
			return colHeaders[i];
		}
		return null;
	}

	public Ptg getCell( int row, int col )
	{
		try
		{
			return rows[row][col];
		}
		catch( Exception e )
		{
			return null;
		}
	}

	public int findCol( String f )
	{
		for( int i = 0; i < colHeaders.length; i++ )
		{
			if( colHeaders[i].trim().equalsIgnoreCase( f ) )
			{
				return i;
			}
		}
		try
		{
			int j = Integer.parseInt( f );    // one-based index into columns
			return j - 1;
		}
		catch( Exception e )
		{
			return -1;
		}
	}

	/**
	 * Write some documentation here please... thanks! -jm
	 *
	 * @param dbrange
	 * @return
	 */
	public static DB parseList( Ptg[] dbrange )
	{
		int prevCol = -1;
		int nCols = 0;
		int nRows = 0;
		int maxRows = 0;

		// allocate the empty table for the dbrange
		for( Ptg aDbrange1 : dbrange )
		{
			if( aDbrange1 instanceof PtgRef )
			{
				PtgRef pref = (PtgRef) aDbrange1;
				int[] loc = pref.getIntLocation();
//				 TODO: check rc sanity here
				if( loc[1] != prevCol )
				{ // count # cols
					prevCol = loc[1];
					nCols++;
					nRows = 0;
				}
				else
				{ // count # rows
					nRows++;
					maxRows = Math.max( nRows, maxRows );
				}
			}
			else
			{
				return null;
			}
		}

		// now populate the table
		DB dblist = new DB( nCols, maxRows );
		prevCol = -1;
		nCols = -1;
		nRows = -1;
		for( Ptg aDbrange : dbrange )
		{
			PtgRef db1 = (PtgRef) aDbrange;
			int[] loc = db1.getIntLocation();
			Object vs = null;    // 20081120 KSC: Must distinguish between blanks and 0's
			if( !db1.isBlank() )
			{
				vs = db1.getValue();
			}

			// column headers
			if( loc[1] != prevCol )
			{
				if( vs != null )
				{
					dblist.colHeaders[++nCols] = vs.toString();
				}
				prevCol = loc[1];
				nRows = 0;
			}
			else
			{ // get value Ptgs
				try
				{
					dblist.rows[nRows++][nCols] = db1;
				}
				catch( ArrayIndexOutOfBoundsException e )
				{
					;    // possible nCols==-1
				}

			}
		}
		return dblist;
	}
}

class Criteria extends DB
{
	private static final Logger log = LoggerFactory.getLogger( Criteria.class );
	public Criteria( int nCols, int nRows )
	{
		super( nCols, nRows );
	}

	public static Criteria parseCriteria( Ptg[] criteria )
	{
		DB dblist = DB.parseList( criteria );
		if( dblist == null )
		{
			return null;
		}
		Criteria crit = new Criteria( dblist.getNCols(), dblist.getNRows() );
		crit.colHeaders = dblist.colHeaders;
		crit.rows = dblist.rows;
		if( log.isDebugEnabled() )
		{
			log.debug( "\nCriteria:" );
			for( int i = 0; i < crit.getNCols(); i++ )
			{
				log.debug( "\t" + crit.getCol( i ) );
			}
			for( int j = 0; j < crit.getNCols(); j++ )
			{
				for( int i = 0; i < crit.getNRows(); i++ )
				{
					log.debug( "\t" + crit.getCell( i, j ) );
				}
			}
		}
		return crit;
	}

	// TODO: Handle formula criteria!
	// TODO: To perform an operation on an entire column in a database, enter a blank line below the column labels in the criteria range
	// TODO: Handle various EQUALS:  currency, number ...
	public static boolean matches( String v, Object cx )
	{
		boolean bMatches = false;
		String c = "";

		if( cx instanceof Ptg )
		{
			c = ((Ptg) cx).getValue().toString();
		}
		else
		{
			c = cx.toString();
		}

		if( (c == null) || (c.length() == 0) )
		{
			return false;
		}
		if( v == null )
		{
			return false;    // 20070208 KSC: null means no match!
		}

		// TODO: handle this using calc methods

		// relational
		if( c.substring( 0, 1 ).equals( ">" ) )
		{
			try
			{
				if( (c.length() > 1) && c.substring( 0, 2 ).equals( ">=" ) )
				{
					c = c.substring( 2 );
					bMatches = (Double.parseDouble( v ) >= Double.parseDouble( c ));
				}
				else
				{
					c = c.substring( 1 );
					bMatches = (Double.parseDouble( v ) > Double.parseDouble( c ));
				}
			}
			catch( NumberFormatException e )
			{
			}
		}
		else if( c.substring( 0, 1 ).equals( "<" ) )
		{
			try
			{
				if( (c.length() > 1) && c.substring( 0, 2 ).equals( "<=" ) )
				{
					c = c.substring( 2 );
					bMatches = (Double.parseDouble( v ) <= Double.parseDouble( c ));
				}
				else
				{
					c = c.substring( 1 );
					bMatches = (Double.parseDouble( v ) < Double.parseDouble( c ));
				}
			}
			catch( NumberFormatException e )
			{
			}
		}
		else
		{// Equals
			bMatches = (v.equalsIgnoreCase( c ));
		}
		return bMatches;
	}

	/**
	 * "passes" means that there is a matching
	 * and valid (criteria passing) cell in the
	 * row of the data db cells
	 *
	 * @param field  - the field or column that is being compared
	 * @param curRow - check the current row/col
	 * @param db     - the db of vals
	 * @return
	 */
	public boolean passes( String field, Ptg[] curRow, DB db )
	{
		int nrows = getNRows();
		int ncols = getNCols();

		int crit_format = -1; // determine format of criteria table
		// multiple rows of criteria for one column are OR searches
		// on that column (col=x OR col=y OR col=z)
		if( (nrows > 1) && (ncols == 1) )
		{
			crit_format = 0;
		}

		// more than one column are AND searches (cola=x AND colb=y)
		// ONE row of criteria is applied to N rows of data
		if( (nrows == 1) && (ncols > 1) )
		{
			crit_format = 1;
		}

		// more than one row and more than one column:  
		// OR's:  (cola=x AND colb=y) || (cola=z AND colb=w)
		// applied to each row of data
		if( (nrows > 1) && (ncols > 1) )
		{
			crit_format = 2;
		}

		if( (nrows == 1) && (ncols == 1) )
		{
			crit_format = 0;
		}

		switch( crit_format )
		{

			case 0:
				return criteriaCheck1( curRow, db );

			case 1:
				return criteriaCheck2( field, curRow, db );

			case 2:
				return criteriaCheck3( field, curRow, db );

			default:
				return criteriaCheck3( field, curRow, db );
		}
	}

	/**
	 * for the current db row/column, see if all criteria for that column matches
	 * (criteria may run over many criteria rows & signifies an OR search)
	 * find db column that matches criteria col
	 * <p/>
	 * multiple rows of data one row of criteria
	 * <p/>
	 * age	height
	 * <29	>99
	 * <p/>
	 * age	height
	 * 23	88
	 * 21	99
	 * 43	56
	 * 23	56
	 * 44	76
	 *
	 * @param curRow
	 * @param db
	 * @return
	 */
	private boolean criteriaCheck1( Ptg[] curRow, DB db )
	{
		boolean bColOK = false;
		boolean bRowOK = true;
		int nrows = getNRows();
		int ncols = getNCols();
		int ndbrows = db.getNRows();

		String cl = getCol( 0 );
		int j = db.findCol( cl );
		if( j >= 0 )
		{
			bColOK = false;    // need one bColOK= true for it to pass
			for( int k = 0; (k < nrows) && !bColOK; k++ )
			{    //
				try
				{
					String v = curRow[j].getValue().toString();
					Ptg r = rows[k][0];
					String rv = r.getValue().toString();
					bColOK = matches( v, r );

					// fast succeed
					if( bColOK )
					{
						return true;
					}

				}
				catch( Exception ex )
				{
					// Logger.logInfo("DBCalc"); // TODO: check that this is OK
				}
			}
			if( !bColOK )
			{
				bRowOK = false;    // if no criteria passes, row doesn't pass
			}
		}
		return bRowOK;
	}

	/**
	 * There is only one row of criteria,
	 * but may be multiple criteria per field aka:
	 * <p/>
	 * type	age	age	 height
	 * blue >21	<50	 44
	 * <p/>
	 * <p/>
	 * // [v1	v2	v3]
	 * // [v2	v3	v4]
	 * // ...
	 * // [crit1	crit2	crit3] <- val must pass this
	 *
	 * @param field
	 * @param curRow
	 * @param db
	 * @return
	 */
	private boolean criteriaCheck2( String field, Ptg[] curRow, DB db )
	{
		boolean pass = true;
		for( int t = 0; t < curRow.length; t++ )
		{
			String valcheck = curRow[t].getValue().toString();
			// for each value check all the criteria
			for( Ptg[] row : rows )
			{
				List r = getCriteria( db.colHeaders[t] );
				Iterator tx = r.iterator();
				while( tx.hasNext() )
				{
					Ptg cv = ((Ptg) tx.next());
					String vc = cv.getValue().toString();
					if( !vc.equals( "" ) )
					{
						pass = matches( valcheck, vc );
						// fast fail is OK
						if( !pass )
						{
							return false;
						}
					}
				}
			}
		}
		return pass;
	}

	/**
	 * othertimes we have a 2D criteria
	 * <p/>
	 * type	age	age	 height
	 * blue >21	>50	 44
	 * blue <30 <100
	 * red  >30 >40  99
	 * <p/>
	 * <p/>
	 * AND criteria across criteria rows
	 * OR  criteria down cols
	 *
	 * @param field
	 * @param curRow
	 * @param db
	 * @return
	 */
	private boolean criteriaCheck3( String field, Ptg[] curRow, DB db )
	{
		boolean critRowMatch = false;
		// for each value check all the criteria in a row
		// multiple rows of criteria are combined 
		for( int x = 0; x < rows.length; x++ )
		{
			critRowMatch = true; // reset
			// for each row of criteria, iterate criteria cols
			for( String critField : colHeaders )
			{
				List r = getCriteria( x, critField );
				Iterator tx = r.iterator();
				int dv = db.getCol( critField );
				String valcheck = curRow[dv].getValue().toString();
				// this criteria row may pass/fail subsequent rows may pass/fail
				// only one has to pass OK to return true all crit in this row must pass
				while( tx.hasNext() && critRowMatch )
				{ // stop if row failure
					Ptg cv = ((Ptg) tx.next());
					String vc = cv.getValue().toString();

					if( !vc.equals( "" ) )
					{
						critRowMatch = matches( valcheck, vc );
						/* fast fail is not OK because we 
						 * may pass one row OR another
						 * AND criteria across criteria rows
						 * OR  criteria down cols
						 */
					}
				}
			}
			if( critRowMatch ) // fast succeed here, a row passed
			{
				return true;
			}
		}
		return critRowMatch;
	}

	/**
	 * equal number of data rows/cols and criteria rows/cols
	 * <p/>
	 * compare each criteria cell with each corresponding db cell
	 *
	 * @param field
	 * @param curRow
	 * @param db
	 * @return
	 */
	private boolean criteriaCheck4( String field, Ptg[] curRow, DB db )
	{
		boolean bColOK = false;
		boolean bRowOK = false;
		int nrows = getNRows();
		int ncols = getNCols();

		for( int k = 0; (k < nrows) && !bRowOK; k++ )
		{    // for each row of criteria
			// for each col check for valid criteria
			for( int i = 0; i < ncols; i++ )
			{
				// find db column that matches criteria column
				// String coli = getCol(i); // get the col

				int coln = db.findCol( field ); // which db col is this?

				if( coln >= 0 )
				{        // matching col in dblist
					// the current criteria matches on colname
					Ptg curcrit = null;
					if( rows[k].length > ncols )
					{ // matched criteria cols and dbcols
						curcrit = rows[k][coln];
					}
					else
					{
						log.warn( "DatabaseCalculator.Criteria.criteriaCheck4: wrong criteria count for db value count" );
						return false;
					}
					String rt = curcrit.getValue().toString();

					// the current val is in the matching column
					Ptg rx = curRow[coln];
					String mv = rx.getValue().toString();

					// check field if this is a matching lookup
					if( (i == 0) )
					{
						if( !mv.equals( rt ) )
						{
							return false;
						}
					}
					else
					{ // check criteria, criteria is fast fail
						if( rt.equals( "" ) )
						{ // empty criteria is not false

						}
						else
						{    // for the current db row/column, see if all criteria for that column matches
							//(criteria may run over many criteria rows & signifies an OR search)
							bColOK = matches( rt, rx );

							if( !bColOK && (nrows == 1) ) // fast fail
							{
								return false;
							}

							if( bColOK ) // fast succeed
							{
								return true;
							}
						}
					}
				}
				else
				{ // see if it's a formula - defined to NOT match db column
					// Logger.logWarn("DataBaseCalculator.Criteria.passes: no matching col in dblist.");
				}
			}
			if( bColOK )
			{
				return true; // does this cut the fat? bRowOK= true;	// if any row's column criteria passes, row passes
			}
		}
		return bRowOK;
	}

	private Map criteriaCache = null;

	/**
	 * return cached array of criteria Ptgs for a given field
	 * <p/>
	 * TODO: cache reset
	 *
	 * @param field
	 * @return
	 */
	private List getCriteria( String field )
	{
		if( true ) //criteriaCache==null)
		{
			criteriaCache = new Hashtable(); // map of vector criteria
		}
		else if( criteriaCache.get( field ) != null )
		{
			return (List) criteriaCache.get( field );
		}

		List crits = new Vector();
		for( Ptg[] row : rows )
		{
			for( int t = 0; t < colHeaders.length; t++ )
			{
				if( colHeaders[t].equals( field ) )
				{
					crits.add( row[t] );
				}
			}
		}
		criteriaCache.put( field, crits );
		return crits;
	}

	/**
	 * return cached array of criteria Ptgs for a given field
	 * <p/>
	 * TODO: cache reset
	 *
	 * @param field
	 * @return
	 */
	private List getCriteria( int critRow, String field )
	{
		List crits = new Vector();
		if( (critRow != -1) )
		{ // option to return criteria per row
			for( int t = 0; t < colHeaders.length; t++ )
			{
				if( colHeaders[t].equals( field ) )
				{
					crits.add( rows[critRow][t] );
				}
			}
		}
		return crits;
	}
}