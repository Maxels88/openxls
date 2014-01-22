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
package com.extentech.formats.XLS;

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.util.ArrayList;

/**
 * Auto filter controls the auto-filter capabilities of Excel.  It has numerous sub-records and references
 * <p/>
 * p><pre>
 * offset  name        size    contents
 * ---
 * 4       iEntry          2       Index of the active autofilter
 * 6       grbit           2       Option flags
 * 8       doper1          10      DOPER struct for the first filter condition
 * 18      doper2          10      DOPER struct for the second filter condition
 * 28      rgch            var     String storage for vtString DOPER
 * <p/>
 * </p></pre>
 */
/* TODO: ************************************************************************************************************
 * Creation of new AutoFilters, removal of existing (see Boundsheet) 
 *
 * fix:  iEntry is not always the index of the column 
 * APPARENTLY if there are more than two autofilters, iEntry is the index of the column
 * if there is only 1 autofilter, iEntry is 0 --- dependent upon Obj record????
 * 
 * Finish:  setTop10, setVal, setVal2: verify all is correctly done ....
 * *******************************************************************************************************************
 */
public final class AutoFilter extends com.extentech.formats.XLS.XLSRecord
{

	private static final long serialVersionUID = -5228830347211523997L;

	short iEntry;    // *** this is the column number ***  oops!  not always!!!
	Doper doper1, doper2;

	// booleans used here for memory space/grbit fields, these are really 1/0 values.  whas the diff?
	boolean wJoin;// true if custom filter conditions are ORed
	boolean fSimple1;// true if the first condition is a simple equality;,
	boolean fSimple2; // trueif the second condition is a simple equality;
	boolean fTop10; // true if the condition is a Top10 autofilter
	boolean fTop; // true if the top 10 AutoFilter shows the top itemsl 0 if it shows the bottom items
	boolean fPercent; // true if the Top 10 AutoFilter shows percentage, 0 if it shows items
	short wTop10; //The number of items to show (from 1-500)

	private byte[] PROTOTYPE_BYTES = {
			00, 00,	/* iEntry */
			00, 00,	/* grbit */
			00, 00, 00, 00, 00, 00, 00, 00, 00, 00, /* unused Doper1 */
			00, 00, 00, 00, 00, 00, 00, 00, 00, 00, /* unused Doper2 */
	};

	/**
	 * creates a new AutoFilter record
	 */
	public static XLSRecord getPrototype()
	{
		AutoFilter af = new AutoFilter();
		af.setOpcode( AUTOFILTER );
		af.setData( af.PROTOTYPE_BYTES );    // 20090630 KSC: don't use static PROTOTYPE_BYTES as changes are propogated
		af.init();
		return af;
	}

	/**
	 * initialize the AutoFilter record
	 */
	@Override
	public void init()
	{
		super.init();
		iEntry = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		//parse out grbit flags
		short grbit = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		this.decodeGrbit( grbit );
		//parse both DOPERs
		doper1 = parseDoper( this.getBytesAt( 4, 10 ) );
		doper2 = parseDoper( this.getBytesAt( 14, 10 ) );
		// parse string data, if any, appended to end of data record
		if( this.getData().length > 25 )
		{
			byte[] rgch = this.getBytesAt( 25, this.getData().length - 25 );
			int pos = 0;
			if( doper1 instanceof StringDoper )
			{
				((StringDoper) doper1).setString( rgch, pos );
				pos += ((StringDoper) doper1).getCch();    // set position pointer to next string data, if any
				pos++;
			}
			if( doper2 instanceof StringDoper )
			{
				((StringDoper) doper2).setString( rgch, pos );
			}
		}
	}

	/**
	 * Parse out the grbit
	 *
	 * @param grbt
	 */
	private void decodeGrbit( short grbit )
	{    // top 500: grbit= -1488
		wJoin = ((grbit & 0x3) == 0x3);    // 0= AND, 3= OR
		fSimple1 = ((grbit & 0x4) == 0x4);
		fSimple2 = ((grbit & 0x8) == 0x8);
		fTop10 = ((grbit & 0x10) == 0x10);
		fTop = ((grbit & 0x20) == 0x20);
		fPercent = ((grbit & 0x40) == 0x40);
		wTop10 = (short) ((grbit & 0xFF80) >> 7);
	}

	/**
	 * encode the grbit from the source flags
	 *
	 * @return short encoded grbit
	 */
	private short encodeGrbit()
	{
		short grbit = 0;
		if( wJoin )
		{
			grbit = 3;    // Or	0= AND
		}
		grbit = ByteTools.updateGrBit( grbit, fSimple1, 2 );
		grbit = ByteTools.updateGrBit( grbit, fSimple2, 3 );
		grbit = ByteTools.updateGrBit( grbit, fTop10, 4 );
		grbit = ByteTools.updateGrBit( grbit, fTop, 5 );
		grbit = ByteTools.updateGrBit( grbit, fPercent, 6 );
		grbit = (short) ((wTop10 << 7) | grbit);
		return grbit;
	}

	/**
	 * Take an array of 10 bytes and parse out the doper
	 *
	 * @param doperBytes
	 * @return
	 */
	private Doper parseDoper( byte[] doperBytes )
	{
		Doper retDoper = null;
		switch( doperBytes[0] )
		{
			case 0x0:
				retDoper = (Doper) new UnusedDoper( doperBytes );
				break;
			case 0x2:
				retDoper = (Doper) new RKDoper( doperBytes );
				break;
			case 0x4:
				retDoper = (Doper) new IEEEDoper( doperBytes );
				break;
			case 0x6:
				retDoper = (Doper) new StringDoper( doperBytes );
				break;
			case 0x8:
				retDoper = (Doper) new ErrorDoper( doperBytes );
				break;
			case 0xC:
				retDoper = (Doper) new AllBlanksDoper( doperBytes );
				break;
			case 0xE:
				retDoper = (Doper) new NoBlanksDoper( doperBytes );
				break;
		}
		return retDoper;
	}

	/**
	 * commit all the flags, dopers, etc to the byte rec array
	 */
	public void update()
	{
		byte[] data = new byte[25];
		byte[] b = ByteTools.shortToLEBytes( iEntry );
		System.arraycopy( b, 0, data, 0, 2 );
		b = ByteTools.shortToLEBytes( encodeGrbit() );
		System.arraycopy( b, 0, data, 2, 2 );
		System.arraycopy( doper1.getRecord(), 0, data, 4, 10 );
		System.arraycopy( doper2.getRecord(), 0, data, 14, 10 );
		if( doper1 instanceof StringDoper )  // append rgch bytes
		{
			data = ByteTools.append( ((StringDoper) doper1).getStrBytes(), data );
		}
		if( doper2 instanceof StringDoper ) // append rgch bytes
		{
			data = ByteTools.append( ((StringDoper) doper2).getStrBytes(), data );
		}
		setData( data );
	}

	/**
	 * Evaluates this AutoFilter's condition for each cell in the indicated column
	 * over all rows in the sheet; if the condition is not met, the row is set to hidden
	 * <br>NOTE: since there may be other conditions (other AutoFilters, for instance)
	 * setting the row to it's hidden state, this method <b>will not</b>
	 * set the row to unhidden if the condition passes.
	 */
	public void evaluate()
	{
		Object val1, val2 = null;
		val1 = getVal( doper1 );    // get the doper value from the 1st doper/comparison, if any
		boolean hasDoper2 = !(doper2 instanceof UnusedDoper);
		if( hasDoper2 )
		{
			val2 = getVal( doper2 );
		}

		String op1 = "=", op2 = "";
		boolean passes = true;
		if( !fSimple1 )
		{
			op1 = doper1.getComparisonOperator();
		}
		if( !fSimple2 && hasDoper2 )
		{
			op2 = doper2.getComparisonOperator();
		}

		// TODO:
		// above/below average?  ooxml only?
		// date/time comparisons????
		// begins with, ends with ...
		if( fTop10 )
		{
			if( fTop )
			{ // ascending
				evaluateTopN();
			}
			else        // descending
			{
				evaluateBottomN();
			}
		}
		else if( (doper1 instanceof AllBlanksDoper) || (doper1 instanceof NoBlanksDoper) )
		{
			boolean filterBlanks = (doper1 instanceof NoBlanksDoper);
			int n = this.getSheet().getNumRows();
			for( int i = 0; i < n; i++ )
			{
				Row r = this.getSheet().getRowByNumber( i );
				if( r == null )
				{// it's blank
					// create a blank cell and then set to hidden?
					if( filterBlanks )
					{
						this.getSheet().addValue( "", ExcelTools.formatLocation( new int[]{ i, iEntry } ) );
						r = this.getSheet().getRowByNumber( i );
						r.setHidden( true );
					}
				}
				else
				{    //row is not blank, check to see if cell is blank
					try
					{
						BiffRec c = r.getCell( iEntry );
						if( (c instanceof Blank) && filterBlanks )
						{
							r.setHidden( true );
						}
						else if( !filterBlanks )
						{
							r.setHidden( true );
						}
					}
					catch( NullPointerException e )
					{
						// NPE= blank
						if( filterBlanks )
						{
							r.setHidden( true );
						}
					}
					catch( CellNotFoundException e )
					{

					}
				}
			}
			Row[] rows = this.getSheet().getRows();
			if( !filterBlanks )
			{ // easy; everything that is NOT BLANK is hidden
				for( int i = 1; i < rows.length; i++ )
				{
					rows[i].setHidden( true );
				}
			}
			else
			{ // everything that is blank is hidden

			}
		}
		else
		{ // all other criteria are based upon operator comparisons ...
			Row[] rows = this.getSheet().getRows();
			for( int i = 1; i < rows.length; i++ )
			{
				try
				{
					BiffRec c = rows[i].getCell( iEntry );
					passes = com.extentech.formats.XLS.formulas.Calculator.compareCellValue( c, val1, op1 );
					if( hasDoper2 && (wJoin || (!wJoin && passes)) )
					{
						passes = com.extentech.formats.XLS.formulas.Calculator.compareCellValue( c, val2, op2 );
					}
					if( !passes )
					{
						rows[i].setHidden( true );
					}
				}
				catch( Exception e )
				{
					;
				} // just keep evaluation
			}
		}
	}

	/**
	 * gets the Object Value of the Doper
	 * either a Double, String or Boolean type
	 *
	 * @param d - Doper
	 * @return Object value of the Doper
	 */
	private Object getVal( Doper d )
	{
		Object val = null;
		if( d instanceof ErrorDoper )
		{
			if( ((ErrorDoper) d).isBooleanVal() )
			{
				val = ((ErrorDoper) d).getBooleanVal();
			}
			else
			{
				val = ((ErrorDoper) d).getErrVal();    // error doper
			}
		}
		else if( d instanceof StringDoper )  // string comparison
		{
			val = d.toString();
		}
		else if( (d instanceof RKDoper) )
		{
			val = ((RKDoper) d).getVal();
		}
		else if( (d instanceof IEEEDoper) )
		{
			val = ((IEEEDoper) d).getVal();
		}
		return val;
	}

	/**
	 * evaluates a Top-N condition, hiding rows that are not in the top N values of the column
	 *
	 * @see evaluateBottomN    for descending or Bottom-N evaluation
	 */
	private void evaluateTopN()
	{
		// must go thru 1+ times as must gather up values then go back and set hidden ...
		// identifies top n values then displays ALL rows that contain those values
		ArrayList top10 = new ArrayList();
		int n = ((!fPercent) ? wTop10 : (this.getSheet().getNumRows() / wTop10));
		double[] maxVals = new double[n];
		for( int i = 0; i < n; i++ )
		{
			maxVals[i] = Double.NEGATIVE_INFINITY;
		}
		double curmin = Double.NEGATIVE_INFINITY;
		Row[] rows = this.getSheet().getRows();
		for( int i = 1; i < rows.length; i++ )
		{
			try
			{
				BiffRec c = rows[i].getCell( iEntry );
				double val = c.getDblVal();
				int insertionpoint = -1;
				if( val >= curmin )
				{
					for( int j = 0; j < n; j++ )
					{    // see where new value falls
						if( val == maxVals[j] )
						{ // then no need to move values around; just add index to set of indexes
							String idxs = (String) top10.get( j );
							if( idxs == null )
							{
								idxs = i + ", ";
							}
							else
							{
								idxs += i + ", ";
							}
							top10.set( j, idxs );
							insertionpoint = -1;    // so don't add below
							break;
						}
						if( val > maxVals[j] )
						{
							if( (insertionpoint == -1) || (maxVals[j] < maxVals[insertionpoint]) )
							{
								insertionpoint = j;    // overwrite point
							}
						}
					}
					if( insertionpoint >= 0 )
					{
						if( top10.size() > insertionpoint )
						{
							top10.remove( insertionpoint );    // replace below
						}
						String idxs = i + ", ";
						top10.add( insertionpoint, idxs );
						maxVals[insertionpoint] = val;
						// reestablish curmin
						curmin = Double.MAX_VALUE;
						for( int j = 0; j < n; j++ )
						{
							curmin = Math.min( maxVals[j], curmin );
						}
					}
				} // othwerwise doesn't meet criteia

			}
			catch( Exception e )
			{
				;
			} // just keep evaluation
		}
		// get master list of non-hidden rows
		String nonhiddenIdxs = "";
		for( int i = 0; i < top10.size(); i++ )
		{
			nonhiddenIdxs += (String) top10.get( i );
		}
		// now that have list of rows which SHOULDN'T be hidden, set all else to hidden
		for( int j = 1; j < rows.length; j++ )
		{
			String idx = j + ", ";
			if( nonhiddenIdxs.indexOf( idx ) == -1 )    // not found!
			{
				rows[j].setHidden( true );
			}
		}
	}

	/**
	 * evaluates a Top-N condition in descending order, in other words, a Bottom-N evaluation,
	 * hiding rows that are not in the bottom N values of the column
	 *
	 * @see evaluateTopN    for ascending or Top-N evaluation
	 */
	private void evaluateBottomN()
	{
		// must go thru 1+ times as must gather up values then go back and set hidden ...
		// identifies bottom n values then displays ALL rows that contain those values
		ArrayList bottomN = new ArrayList();
		int n = ((!fPercent) ? wTop10 : (this.getSheet().getNumRows() / wTop10));
		double[] minVals = new double[n];
		for( int i = 0; i < n; i++ )
		{
			minVals[i] = Double.POSITIVE_INFINITY;
		}
		double curmax = Double.POSITIVE_INFINITY;
		Row[] rows = this.getSheet().getRows();
		for( int i = 1; i < rows.length; i++ )
		{
			try
			{
				BiffRec c = rows[i].getCell( iEntry );
				double val = c.getDblVal();
				int insertionpoint = -1;
				if( val <= curmax )
				{
					for( int j = 0; j < n; j++ )
					{    // see where new value falls
						if( val == minVals[j] )
						{ // then no need to move values around; just add index to set of indexes
							String idxs = (String) bottomN.get( j );
							if( idxs == null )
							{
								idxs = i + ", ";
							}
							else
							{
								idxs += i + ", ";
							}
							bottomN.set( j, idxs );
							insertionpoint = -1;    // so don't add below
							break;
						}
						if( val < minVals[j] )
						{
							if( (insertionpoint == -1) || (minVals[j] > minVals[insertionpoint]) )
							{
								insertionpoint = j;    // overwrite point
							}
						}
					}
					if( insertionpoint >= 0 )
					{
						if( bottomN.size() > insertionpoint )
						{
							bottomN.remove( insertionpoint );    // replace below
						}
						String idxs = i + ", ";
						bottomN.add( insertionpoint, idxs );
						minVals[insertionpoint] = val;
						// reestablish curmax
						curmax = Double.NEGATIVE_INFINITY;
						for( int j = 0; j < n; j++ )
						{
							curmax = Math.max( minVals[j], curmax );
						}
					}
				} // othwerwise doesn't meet criteia

			}
			catch( Exception e )
			{
				;
			} // just keep evaluation
		}
		// get master list of non-hidden rows
		String nonhiddenIdxs = "";
		for( int i = 0; i < bottomN.size(); i++ )
		{
			nonhiddenIdxs += (String) bottomN.get( i );
		}
		// now that have list of rows which SHOULDN'T be hidden, set all else to hidden
		for( int j = 1; j < rows.length; j++ )
		{
			String idx = j + ", ";
			if( nonhiddenIdxs.indexOf( idx ) == -1 )    // not found!
			{
				rows[j].setHidden( true );
			}
		}
	}

	/**
	 * return a string representation of this autofilter
	 */
	public String toString()
	{
		String op1 = "=", op2 = "";
		boolean hasDoper2 = ((doper2 != null) && !(doper2 instanceof UnusedDoper));
		if( fTop10 )
		{
			if( fTop )
			{
				if( fPercent )
				{
					return "Top " + wTop10 + "%";
				}
				return "Top " + wTop10 + " Items";
			}
			if( fPercent )
			{
				return "Bottom " + wTop10 + "%";
			}
			return "Bottom " + wTop10 + " Items";
		}
		if( !fSimple1 )
		{
			op1 = doper1.getComparisonOperator();
		}
		else
		{
			if( doper1 instanceof AllBlanksDoper )
			{
				return "Non Blanks";
			}
			if( doper1 instanceof NoBlanksDoper )
			{
				return "Blanks";
			}
		}

		if( !fSimple2 && hasDoper2 )
		{
			op2 = doper2.getComparisonOperator();
		}
		if( !hasDoper2 )    // just == the doper
		{
			return op1 + doper1.toString();
		}
		return op1 + doper1.toString() + ((wJoin) ? " OR " : " AND ") + op2 + doper2.toString();
	}

	/**
	 * Update the record before streaming
	 *
	 * @see com.extentech.formats.XLS.XLSRecord#preStream()
	 */
	@Override
	public void preStream()
	{
		// no need to update unless things have changed ... this.update();
	}

	/**
	 * return the column number (0-based) that this AutoFilter references
	 *
	 * @return int column number
	 */
	public int getCol()
	{
		return iEntry;
	}

	/**
	 * set the column number (0-based) that this AutoFilter references
	 *
	 * @param int col - 0-based column number
	 */
	public void setCol( int col )
	{
		iEntry = (short) col;
		update();
	}

	/**
	 * Sets the custom comparison of this AutoFilter via a String operator and an Object value
	 * <br>Only those records that meet the equation (column value) <operator> value will be shown
	 * <br>e.g show all rows where column value >= 2.99
	 * <p>Object value can be of type
	 * <p>String
	 * <br>Boolean
	 * <br>Error
	 * <br>a Number type object
	 * <p>String operator may be one of: "=", ">", ">=", "<>", "<", "<="
	 *
	 * @param Object val - value to set
	 * @param String op - operator
	 * @see setVal2
	 */
	public void setVal( Object val, String op )
	{
		byte[] doperRec = new byte[10];
		if( val instanceof String )
		{// doper1 should be a string
			String s = (String) val;
			if( !s.startsWith( "!" ) )
			{ // it's not an error val
				doperRec[0] = 0x6;
				doper1 = new StringDoper( doperRec );
				((StringDoper) doper1).setString( s );
			}
			else
			{ // it's an error doper
				doperRec[0] = 0x8;
				doperRec[2] = 1;    // fError
				doper1 = new ErrorDoper( doperRec );
			}
		}
		else if( val instanceof Boolean )
		{
			doperRec[0] = 0x8;
			doperRec[2] = 0;    // fError
			doperRec[3] = (byte) ((Boolean) val ? 1 : 0);
			doper1 = new ErrorDoper( doperRec );
		}
		else
		{   // assume a Number object
			doperRec[0] = 0x2;
			try
			{
				double d = 0;
				if( val instanceof Double )
				{
					d = (Double) val;
				}
				else if( val instanceof Integer )
				{
					d = ((Integer) val).doubleValue();
				}
				else
				{
					throw new NumberFormatException( "Unable to convert to Numeric Object" + val.getClass() );
				}
				doper1 = new RKDoper( doperRec );
				((RKDoper) doper1).setVal( d );
				//doper1= new IEEEDoper(doperRec);
			}
			catch( Exception e )
			{
				Logger.logErr( "AutoFilter.setVal: error setting value to " + val + ":" + e.toString() );
			}
		}
		doper1.setComparisonOperator( op );
		fSimple1 = ("=".equals( op ));
		update();
	}

	/**
	 * Sets the custom comparison of the second condition of this AutoFilter via a String operator and an Object value
	 * <br>This method sets the second condition of a two-condition filter
	 * <p>Only those records that meet the equation:
	 * <br> first condiition AND/OR (column value) <operator> Value will be shown
	 * <br>e.g show all rows where (column value) <= 1.99 AND (column value) >= 2.99
	 * <p>Object value can be of type
	 * <p>String
	 * <br>Boolean
	 * <br>Error
	 * <br>a Number type object
	 * <p>String operator may be one of: "=", ">", ">=", "<>", "<", "<="
	 *
	 * @param Object  val - value to set
	 * @param String  op - operator
	 * @param boolean AND - true if two conditions should be AND'ed, false if OR'd
	 * @see setVal2
	 */
	public void setVal2( Object val, String op, boolean AND )
	{
		byte[] doperRec = new byte[10];
		if( val instanceof String )
		{// doper1 should be a string
			String s = (String) val;
			if( !s.startsWith( "!" ) )
			{ // it's not an error val
				doperRec[0] = 0x6;
				doper2 = new StringDoper( doperRec );
				((StringDoper) doper1).setString( s );
			}
			else
			{ // it's an error doper
				doperRec[0] = 0x8;
				doperRec[2] = 1;    // fError
				doper2 = new ErrorDoper( doperRec );
			}
		}
		else if( val instanceof Boolean )
		{
			doperRec[0] = 0x8;
			doperRec[2] = 0;    // fError
			doperRec[3] = (byte) ((Boolean) val ? 1 : 0);
			doper2 = new ErrorDoper( doperRec );
		}
		else
		{   // assume a Number object
			doperRec[0] = 0x2;
			try
			{
				double d = 0;
				if( val instanceof Double )
				{
					d = (Double) val;
				}
				else if( val instanceof Integer )
				{
					d = ((Integer) val).doubleValue();
				}
				else
				{
					throw new NumberFormatException( "Unable to convert to Numeric Object" + val.getClass() );
				}
				doper2 = new RKDoper( doperRec );
				((RKDoper) doper2).setVal( d );
				//doper1= new IEEEDoper(doperRec);
			}
			catch( Exception e )
			{
				Logger.logErr( "AutoFilter.setVal: error setting value to " + val + ":" + e.toString() );
			}
		}
		doper2.setComparisonOperator( op );
		fSimple2 = ("=".equals( op ));
		wJoin = !AND;
		update();
	}

	/**
	 * return the value of the first comparison of this AutoFilter, if any
	 *
	 * @return Object value
	 * @see getVal2
	 */
	@Override
	public Object getVal()
	{
		return doper1.toString();
	}

	/**
	 * returns the value of the second comparison of this AutoFilter, if any
	 *
	 * @return Object value
	 * @see getVal
	 */
	public Object getVal2()
	{
		return doper2.toString();
	}

	/**
	 * get the operator associated with this AutoFilter
	 * <br>NOTE: this will return the operator in the first condition if this AutoFilter contains two conditions
	 * <br>Use getOp2 to retrieve the second condition operator
	 *
	 * @return String operator: one of "=", ">", ">=", "<>", "<", "<="
	 * @see getOp2
	 */
	public String getOp()
	{
		return doper1.getComparisonOperator();
	}

	/**
	 * get the operator associated with the second condtion of this AutoFilter, if any
	 * <br>NOTE: this will return the operator in the second condition if this AutoFilter contains two conditions
	 *
	 * @return String operator: one of "=", ">", ">=", "<>", "<", "<=", or null, if no second condition
	 * @see getOp
	 */
	public String getOp2()
	{
		if( !(doper2 instanceof UnusedDoper) )
		{
			return doper2.getComparisonOperator();
		}
		return null;
	}

	/**
	 * sets this AutoFilter to be a Top-n type of filter
	 * <br>Top-n filters only show the Top n values or percent in the column
	 * <br>n can be from 1-500, or 0 to turn off Top 10 filtering
	 *
	 * @param int     n - 0-500
	 * @param boolean percent - true if show Top-n percent; false to show Top-n items
	 * @param boolean top10 - true if show Top 10 (items or percent), false to show Bottom N (items or percent)
	 */
	public void setTop10( int n, boolean percent, boolean top10 )
	{
		if( n == 0 )
		{
			fTop = false;
			fTop10 = false;
			fPercent = false;
			wTop10 = 0;
			// TODO: set fSimple1?  remove dopers??
		}
		else if( (n > 0) && (n <= 500) )
		{
			fTop = top10;
			fTop10 = true;
			wTop10 = (short) n;
			fPercent = percent;
			fSimple1 = false;    // true if 1st condition is simple equality
			fSimple2 = false;    // true if 2nd condition is simple equality
			doper1 = new IEEEDoper( new byte[]{ 4, 6, 0, 0, 0, 0, 0, 0, 0, 0 } );
			((IEEEDoper) doper1).setVal( n );
			doper2 = new UnusedDoper( new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 } );
		}
		else
		{
			Logger.logErr( "AutoFilter.setTop10: value " + n + " must be between 0 and 500" );
		}
		update();
	}

	/**
	 * sets this AutoFilter to filter all blank rows
	 */
	public void setFilterBlanks()
	{
		fSimple1 = true;
		fSimple2 = false;
		fTop = false;
		fTop10 = false;
		fPercent = false;
		wTop10 = 0;
		doper1 = new NoBlanksDoper( new byte[]{ 14, 5, 0, 0, 0, 0, 0, 0, 0, 0 } );
		doper2 = new UnusedDoper( new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 } );
		update();
	}

	/**
	 * returns true if this AutoFitler is set to filter blank rows
	 *
	 * @return boolean
	 */
	public boolean isFilterBlanks()
	{
		return (fSimple1 && !fSimple2 && !fTop && !fTop10 && (doper1 instanceof NoBlanksDoper));
	}

	/**
	 * sets this AutoFilter to filter all non-blank rows
	 */
	public void setFilterNonBlanks()
	{
		fSimple1 = true;
		fSimple2 = false;
		fTop = false;
		fTop10 = false;
		fPercent = false;
		wTop10 = 0;
		doper1 = new AllBlanksDoper( new byte[]{ 12, 2, 0, 0, 0, 0, 0, 1, 0, 0 } );
		doper2 = new UnusedDoper( new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 } );
		update();
	}

	/**
	 * returns true if this AutoFitler is set to filter all non-blank rows
	 *
	 * @return boolean
	 */
	public boolean isFilterNonBlanks()
	{
		return (fSimple1 && !fSimple2 && !fTop && !fTop10 && (doper1 instanceof AllBlanksDoper));
	}

	/**
	 * returns true if this is a top-10-type AutoFilter
	 *
	 * @return true if this is a top-10-type AutoFilter, false otherwise
	 */
	public boolean isTop10()
	{
		return (fTop10 && (wTop10 > 0));
	}

	/**
	 * Doper (Database OPER Structure) are parsed definitions of the AutoFilter
	 * <p/>
	 * 10-byte parsed definitions that appear in the Custom AutoFilter dialog box
	 * <p>There are several sub-types of Dopers:
	 * <br>String
	 * <br>IEEE Floating Point
	 * <br>RK
	 * <br>Error or Boolean
	 * <br>All Blanks
	 * <br>All non-blanksfs
	 * <br>Unused -- placeholder doper
	 *
	 * @see Excel bible page 284
	 */
	private class Doper
	{
		byte vt, grbitSign;
		byte[] doperRec;

		protected Doper( byte[] rec )
		{
			vt = rec[0];
			grbitSign = rec[1];    // comparison code
			doperRec = rec;
		}

		public byte getGrbitSign()
		{
			return grbitSign;
		}

		public void setGrbitSign( byte grbitSign )
		{
			this.grbitSign = grbitSign;
		}

		public byte getVt()
		{
			return vt;
		}

		public void setVt( byte vt )
		{
			this.vt = vt;
		}

		public byte[] getRecord()
		{
			return doperRec;
		}

		public String toString()
		{
			return null;
		}

		/**
		 * Returns the comparison operator for this doper, ie '>' '=', etc
		 *
		 * @return
		 */
		public String getComparisonOperator()
		{
			switch( grbitSign )
			{
				case 1:
					return "<";
				case 2:
					return "=";
				case 3:
					return "<=";
				case 4:
					return ">";
				case 5:
					return "<>";
				case 6:
					return ">=";
			}
			return "";
		}

		/**
		 * sets the custom comparison operator to one of:
		 * <br>"<"
		 * <br>"="
		 * <br>"<="
		 * <br>">"
		 * <br>"<>"
		 * <br>">="
		 *
		 * @param String op - custom operator
		 */
		public void setComparisonOperator( String op )
		{
			if( "<".equals( op ) )
			{
				grbitSign = 1;
			}
			if( "=".equals( op ) )
			{
				grbitSign = 2;
			}
			if( "<=".equals( op ) )
			{
				grbitSign = 3;
			}
			if( ">".equals( op ) )
			{
				grbitSign = 4;
			}
			if( "<>".equals( op ) )
			{
				grbitSign = 5;
			}
			if( ">=".equals( op ) )
			{
				grbitSign = 6;
			}
			doperRec[1] = grbitSign;
		}
	}

	/**
	 * A doper representing an unused value/filter condition unused.
	 */
	private class UnusedDoper extends Doper
	{

		protected UnusedDoper( byte[] rec )
		{
			super( rec );
		}

		public String toString()
		{
			return "Unused";
		}

	}

	/**
	 * A doper representing an RK number
	 * <p>Dopers define an AutoFilter value using a 10-byte doperRec
	 * <br>For all dopers, doperRec[0]=vt or code
	 * <br>doperRec[1]=comparison operator
	 * <p>For RK Dopers,
	 * doperRec[2]->[6] = rk number, 6-9= reserved
	 */
	private class RKDoper extends Doper
	{
		double val;

		protected RKDoper( byte[] rec )
		{
			super( rec );
			byte[] b = new byte[4];
			System.arraycopy( doperRec, 2, b, 0, 4 );
			val = Rk.parseRkNumber( b );    // parse bytes in Rk-number format into a double value
		}

		public String toString()
		{
			return new Double( val ).toString();
		}

		/**
		 * set the double value for this Rk-type Doper record
		 *
		 * @param double d - double value to set
		 */
		public void setVal( double d )
		{
			val = d;
			byte[] b = Rk.getRkBytes( d );
			System.arraycopy( b, 0, doperRec, 2, 4 );
		}

		/**
		 * return the double value associated with this Rk-type Doper record
		 *
		 * @return double value
		 */
		public double getVal()
		{
			return val;
		}
	}

	/**
	 * A doper representing a IEEE number
	 * <p>Dopers define an AutoFilter value using a 10-byte doperRec
	 * <br>For all dopers, doperRec[0]=vt or code
	 * <br>doperRec[1]=comparison operator
	 * <p>For IEEE Dopers,
	 * doperRec[2->9] = IEEE floating point number
	 */
	private class IEEEDoper extends Doper
	{
		double val;

		/**
		 * create an IEEEDoper Object from doper record bytes
		 * (10 bytes as part of the AutoFilter record)
		 *
		 * @param byte[] rec 10 byte doper record
		 */
		protected IEEEDoper( byte[] rec )
		{
			super( rec );
			byte[] b = new byte[8];
			System.arraycopy( doperRec, 2, b, 0, 8 );
			val = ByteTools.eightBytetoLEDouble( b );        // TODO: is this correct??
		}

		/**
		 * return the double value associated with this IEEE-type Doper record
		 *
		 * @return double value
		 */
		public double getVal()
		{
			return val;
		}

		/**
		 * set the double value for this IEEE-type Doper record
		 *
		 * @param double d - double value to set
		 */
		public void setVal( double d )
		{
			val = d;
			byte[] b = ByteTools.doubleToLEByteArray( val );
			System.arraycopy( b, 0, doperRec, 2, 8 );
		}

		public String toString()
		{
			return new Double( val ).toString();
		}
	}

	/**
	 * A doper representing a String
	 * <p>Dopers define an AutoFilter value using a 10-byte doperRec
	 * <br>For all dopers, doperRec[0]=vt or code
	 * <br>doperRec[1]=comparison operator
	 * <p>For String Dopers,
	 * doperRec[6]= cch or String length
	 * <br>The actual string data is appended to the end of the entire AutoFilter data record
	 * <br>
	 * NOTE: strings must be in normal or default encoding
	 */
	private class StringDoper extends Doper
	{
		String s;

		protected StringDoper( byte[] rec )
		{
			super( rec );
		}

		/**
		 * sets the string for this StringDoper from a byte array and an index into said array
		 * <br>amount to read from byteArray is stored in cch, bit 6 of the 10-byte doperRec
		 *
		 * @param byte[] rgch - source byte array for string
		 * @param int    start - start index into the byte array
		 */
		public void setString( byte[] rgch, int start )
		{
			int cch = doperRec[6];
			byte[] stringbytes = new byte[cch];
			System.arraycopy( rgch, start, stringbytes, 0, cch );
			s = new String( stringbytes );
		}

		/**
		 * returns cch, the length of this string data
		 *
		 * @return int
		 */
		public int getCch()
		{
			return doperRec[6];
		}

		/**
		 * s the string for this StringDoper from a String Object
		 *
		 * @param String s
		 */
		public void setString( String s )
		{
			this.s = s;
			byte[] b = s.getBytes();
			doperRec[6] = (byte) b.length;
		}

		/**
		 * returns the byte array representing the String referenced by this StringDoper
		 *
		 * @return byte[]
		 */
		public byte[] getStrBytes()
		{
			return s.getBytes();
		}

		/**
		 * returns the String referenced by this StringDoper
		 */
		public String toString()
		{
			return s;
		}
	}

	/**
	 * A doper representing an Error or Boolean
	 * <p>Dopers define an AutoFilter value using a 10-byte doperRec
	 * <br>For all dopers, doperRec[0]=vt or code
	 * <br>doperRec[1]=comparison operator
	 * <p>For Error or Boolean Dopers,
	 * <ul>doperRec[2]= fError; </ul>
	 * <ul>if 0, doperRec[3]= boolean value</ul>
	 * <ul>				 if 1, doperRec[3]= Error value</ul>
	 * Error Values:
	 * <br>	0	= #NULL!
	 * <br>	0x7	= #DIV/0!
	 * <br>	0xF	= #VALUE!
	 * <br>	0x17= #REF!
	 * <br>	0x1D= #NAME?
	 * <br>	0x24= #NUM!
	 * <br>	0x2A= #N/A
	 */
	private class ErrorDoper extends Doper
	{
		boolean bVal;
		int errVal;

		protected ErrorDoper( byte[] rec )
		{
			super( rec );
			if( doperRec[2] == 0 )    // boolean doper
			{
				bVal = (doperRec[3] != 0);
			}
			else
			{
				errVal = doperRec[3];    // see above vals
			}
		}

		/**
		 * returns true if this is a Error Doper
		 *
		 * @return boolean true if this is a Error Doper
		 */
		public boolean isErrVal()
		{
			return (doperRec[2] == 1);
		}

		/**
		 * returns true if this is a Boolean Doper
		 *
		 * @return boolean true if this is a boolean Doper
		 */
		public boolean isBooleanVal()
		{
			return (doperRec[2] == 0);
		}

		/**
		 * returns the boolean value if this is a type Boolean Doper
		 *
		 * @return boolean
		 */
		public boolean getBooleanVal()
		{
			return bVal;
		}

		/**
		 * interprets the error code located in doperRec[2] into a String Error Value
		 * e.g. "#NULL!"
		 *
		 * @return String error Value
		 */
		public String getErrVal()
		{
			if( doperRec[2] == 1 )
			{
				switch( errVal )
				{
					case 0:
						return "#NULL!";
					case 0x7:
						return "#DIV/0!";
					case 0xF:
						return "#VALUE!";
					case 0x17:
						return "#REF!";
					case 0x1D:
						return "#NAME?";
					case 0x24:
						return "#NUM!";
					case 0x2A:
						return "#N/A";
				}
			}
			return "";
		}

		/**
		 * returns a String representation of this Error or Boolean Doper
		 * <br>If this is an Error Doper, returns one of the Error Values
		 * <br>If this is a booelan Doper, returns "true" or "false"
		 *
		 * @return String representation
		 * @see AutoFilter.getErrVal()
		 */
		public String toString()
		{
			if( isErrVal() )
			{
				return getErrVal();
			}
			return bVal ? "true" : "false";
		}
	}

	/**
	 * A doper representing an all blanks selection.
	 * <br>bytes are unused other than identifier
	 */
	private class AllBlanksDoper extends Doper
	{

		protected AllBlanksDoper( byte[] rec )
		{
			super( rec );
		}

		public String toString()
		{
			return "All Blanks";
		}
	}

	/**
	 * A doper representing an all blanks selection.
	 * <br>bytes are unused other than identifier
	 */
	private class NoBlanksDoper extends Doper
	{

		protected NoBlanksDoper( byte[] rec )
		{
			super( rec );
		}

		public String toString()
		{
			return "No Blanks";
		}
	}

}
