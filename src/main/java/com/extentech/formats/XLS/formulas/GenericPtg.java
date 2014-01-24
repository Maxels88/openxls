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

import com.extentech.ExtenXLS.DateConverter;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.FormulaNotFoundException;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.formats.XLS.Name;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Calendar;

public abstract class GenericPtg implements Ptg, Cloneable
{
	private static final Logger log = LoggerFactory.getLogger( GenericPtg.class );
	double doublePrecision = 0.00000001;        // doubles/floats cannot be compared for exactness so use precision comparator
	public static final long serialVersionUID = 666555444333222l;
	byte ptgId;
	byte[] record;

	Ptg[] vars = null;
	int lock_id = -1;
	private int locationLocked = Ptg.PTG_LOCATION_POLICY_UNLOCKED;
	private BiffRec trackercell = null;

	@Override
	public Object clone()
	{
		try
		{
			return super.clone();
		}
		catch( CloneNotSupportedException e )
		{
			// This is, in theory, impossible
			return null;
		}
	}

	/**
	 * a locking mechanism so that Ptgs are not endlessly
	 * re-calculated
	 *
	 * @return
	 */
	@Override
	public int getLock()
	{
		return lock_id;
	}

	/**
	 * a locking mechanism so that Ptgs are not endlessly
	 * re-calculated
	 *
	 * @return
	 */
	@Override
	public void setLock( int x )
	{
		lock_id = x;
	}

	// determine behavior
	@Override
	public boolean getIsOperator()
	{
		return false;
	}

	@Override
	public boolean getIsBinaryOperator()
	{
		return false;
	}

	@Override
	public boolean getIsUnaryOperator()
	{
		return false;
	}

	@Override
	public boolean getIsStandAloneOperator()
	{
		return false;
	}

	@Override
	public boolean getIsPrimitiveOperator()
	{
		return false;
	}

	@Override
	public boolean getIsOperand()
	{
		return false;
	}

	@Override
	public boolean getIsFunction()
	{
		return false;
	}

	@Override
	public boolean getIsControl()
	{
		return false;
	}

	@Override
	public boolean getIsArray()
	{
		return false;
	}

	@Override
	public boolean getIsReference()
	{
		return false;
	}

	/**
	 * returns the Location Policy of the Ptg is locked
	 * used during automated BiffRec movement updates
	 *
	 * @return int
	 */
	@Override
	public int getLocationPolicy()
	{
		return locationLocked;
	}

	/**
	 * lock the Location of the Ptg so that it will not
	 * be updated during automated BiffRec movement updates
	 *
	 * @param b setting of the lock the location policy for this Ptg
	 */
	@Override
	public void setLocationPolicy( int b )
	{
		locationLocked = b;
	}

	/**
	 * update the Ptg
	 */
	@Override
	public void updateRecord()
	{

	}

	/**
	 * Returns the number of Params to pass to the Ptg
	 */
	@Override
	public int getNumParams()
	{
		if( getIsPrimitiveOperator() )
		{
			return 2;
		}
		return 0;
	}

/* ################################################### EXPLANATION ###################################################
   
    1. set string varetvar in all Ptgs
    2. varetvar goes between ptg return vals if any
    3. if this is a funcvar then we loop ptgs and out 
    4. when we call getString or evaluate, we loop into the
        recursive tree and execute on up.
  
   ################################################### EXPLANATION ###################################################*/

	/**
	 * Operator Ptgs take other Ptgs as arguments
	 * so we need to pass them in to get a meaningful
	 * value.
	 */
	@Override
	public void setVars( Ptg[] parr )
	{
		vars = parr;
	}

	/*
		Return all of the cells in this range as an array
		of Ptg's.  This is used for range calculations.
	*/
	@Override
	public Ptg[] getComponents()
	{
		return null;
	}

	/**
	 * pass  in arbitrary number of values (probably other Ptgs)
	 * and return the resultant value.
	 * <p/>
	 * This effectively calculates the Expression.
	 */
	@Override
	public Object evaluate( Object[] obj )
	{
		// do something useful
		return getString();
	}

	/**
	 * return the human-readable String representation of
	 * this ptg -- if applicable
	 */
	@Override
	public String getTextString()
	{

		String strx = "";

		try
		{
			strx = getString();
		}
		catch( Exception e )
		{
			log.error( "Function not supported: " + parent_rec.toString() );
		}

		if( strx == null )
		{
			return "";
		}

		StringBuffer out = new StringBuffer( strx );
		if( vars != null )
		{
			int numvars = vars.length;
			if( getIsPrimitiveOperator() && getIsUnaryOperator() )
			{
				if( numvars > 0 )
				{
					out.append( vars[0].getTextString() );
				}

			}
			else if( getIsPrimitiveOperator() )
			{
				out.setLength( 0 );
				for( int x = 0; x < numvars; x++ )
				{
					out.append( vars[x].getTextString() );
					if( (x + 1) < numvars )
					{
						out.append( getString() );
					}
				}
			}
			else if( getIsControl() )
			{
				for( Ptg var : vars )
				{
					out.append( var.getTextString() );
				}
			}
			else
			{
				for( int x = 0; x < vars.length; x++ )
				{
					if( !((x == 0) && (vars[x] instanceof PtgNameX)) )
					{    // KSC: added to skip External name reference for Add-in Formulas
						String part = vars[x].getTextString();
						// 20060408 KSC: added quoting in PtgStr.getTextString
//	                    if (vars[x] instanceof PtgStr) // 20060214 KSC: Quote string params
//	                    	part= "\"" + part + "\"";
						out.append( part );
		                /*if(!part.equals(""))*/
						out.append( "," );
					}
				}
				if( vars.length > 0 ) // don't strip 1st paren if no params!  20060501 KSC
				{
					out.setLength( out.length() - 1 ); // strip trailing comma
				}
			}
		}
		out.append( getString2() );
		return out.toString();
	}

	/*text1 and 2 for this Ptg
	*/
	@Override
	public String getString()
	{
		return toString();
	}

	/**
	 * return the human-readable String representation of
	 * the "closing" portion of this Ptg
	 * such as a closing parenthesis.
	 */

	@Override
	public String getString2()
	{
		if( getIsPrimitiveOperator() )
		{
			return "";
		}
		if( getIsOperator() )
		{
			return ")";
		}
		return "";
	}

	@Override
	public byte getOpcode()
	{
		return ptgId;
	}

	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
	}

	/**
	 * return a Ptg  consisting of the calculated values
	 * of the ptg's passed in.  Returns null for any non-operand
	 * ptg.
	 *
	 * @throws CalculationException
	 */
	@Override
	public Ptg calculatePtg( Ptg[] parsething ) throws FunctionNotSupportedException, CalculationException
	{
		return null;

	}

	/**
	 * Gets the (return) value of this Ptg as an operand Ptg.
	 */
	@Override
	public Ptg getPtgVal()
	{
		Object value = getValue();
		if( value instanceof Ptg )
		{
			return (Ptg) value;
		}
		if( value instanceof Boolean )
		{
			return new PtgBool( (Boolean) value );
		}
		if( value instanceof Integer )
		{
			return new PtgInt( (Integer) value );
		}
		if( value instanceof Number )
		{
			return new PtgNumber( ((Number) value).doubleValue() );
		}
		if( value instanceof String )
		{
			return new PtgStr( (String) value );
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * returns the value of an operand ptg.
	 *
	 * @return null for non-operand Ptg.
	 */
	@Override
	public Object getValue()
	{
		return null;
	}

	/**
	 * Gets the value of the ptg represented as an int.
	 * <p/>
	 * This can result in loss of precision for floating point values.
	 * <p/>
	 * overridden in PtgInt to natively return value.
	 *
	 * @return integer representing the ptg, or NAN
	 */
	@Override
	public int getIntVal()
	{
		try
		{
			return new Double( getValue().toString() ).intValue();
		}
		catch( NumberFormatException e )
		{
			// we should be throwing something better
			if( !(this instanceof PtgErr) )    // don't report an error if it's already an error
			{
				log.error( "GetIntVal failed for formula: " + getParentRec().toString() + " " + e );
			}
			return 0;
			///  RIIIIGHT!  throw new FormulaCalculationException();
		}
	}

	/**
	 * Gets the value of the ptg represented as an double.
	 * <p/>
	 * This can result in loss of precision for floating point values.
	 * <p/>
	 * NAN will be returned for values that are not translateable to an double
	 * <p/>
	 * overrideen in PtgNumber
	 *
	 * @return integer representing the ptg, or NAN
	 */
	@Override
	public double getDoubleVal()
	{
		Object pob = null;
		Double d = null;
		try
		{
			pob = getValue();
			if( pob == null )
			{
				log.error( "Unable to calculate Formula at " + getLocation() );
				return java.lang.Double.NaN;
			}
			d = (Double) pob;
		}
		catch( ClassCastException e )
		{
			try
			{
				Float f = (Float) pob;
				d = f.doubleValue();
			}
			catch( ClassCastException e2 )
			{
				try
				{
					Integer in = (Integer) pob;
					d = in.doubleValue();
				}
				catch( Exception e3 )
				{
					if( (pob == null) || pob.toString().equals( "" ) )
					{
						d = (double) 0;
					}
					else
					{
						try
						{
							Double dd = new Double( pob.toString() );
							return dd;
						}
						catch( Exception e4 )
						{// Logger.logWarn("Error in Ptg Calculator getting Double Value: " + e3);
							return java.lang.Double.NaN;
						}
					}
				}
			}
		}
		catch( Throwable exp )
		{
			log.error( "Unexpected Exception in PtgCalculator.getDoubleValue()", exp );
		}
		return d;
	}

	/**
	 * So, here you see we can get the static type from the record itself
	 * then format the output record.  Some shorthand techniques are shown.
	 */
	@Override
	public byte[] getRecord()
	{
		return record;
	}

	// these do nothing here...
	@Override
	public void setLocation( String s )
	{
		;
	}

	@Override
	public String getLocation() throws FormulaNotFoundException
	{
		return null;
	}

	@Override
	public int[] getIntLocation() throws FormulaNotFoundException
	{
		return null;
	}

	// Parent Rec is the BiffRec record referenced by Operand Ptgs
	protected XLSRecord parent_rec;

	@Override
	public void setParentRec( XLSRecord f )
	{
		parent_rec = f;
	}

	@Override
	public XLSRecord getParentRec()
	{
		return parent_rec;
	}

	/**
	 * Returns an array of doubles from number-type ptg's sent in.
	 * This should only be referenced by sub-classes.
	 * <p/>
	 * Null values accessed are treated as 0.  Within excel (empty cell values == 0) Tested!
	 * Sometimes as well you can get empty string values, "".  These are NOT EQUAL ("" != 0)
	 *
	 * @param pthings
	 * @return
	 */
	protected static Object[] getValuesFromPtgs( Ptg[] pthings )
	{
		Object[] obar = new Object[pthings.length];
		for( int t = 0; t < obar.length; t++ )
		{
			if( pthings[t] instanceof PtgErr )
			{
				return null;
			}
			if( pthings[t] instanceof PtgArray )
			{
				obar[t] = pthings[t].getComponents();    // get all items in array as Ptgs
				Object v = null;
				try
				{
					v = getValuesFromObjects( (Object[]) obar[t] );    // get value array from the ptgs
				}
				catch( NumberFormatException e )
				{    // string or non-numeric values
					v = getStringValuesFromPtgs( (Ptg[]) obar[t] );
				}
				obar[t] = v;
			}
			else
			{
				Object pval = pthings[t].getValue();
				if( pval instanceof PtgArray )
				{
					obar[t] = ((PtgArray) pval).getComponents();    // get all items in array as Ptgs
					Object v = null;
					try
					{
						v = getValuesFromObjects( (Object[]) obar[t] );    // get value array from the ptgs
					}
					catch( NumberFormatException e )
					{    // string or non-numeric values
						v = getStringValuesFromPtgs( (Ptg[]) obar[t] );
					}
					obar[t] = v;
				}
				else if( pval instanceof Name )
				{    // then get it's components ...
					obar[t] = pthings[t].getComponents();
					Object v = null;
					try
					{
						v = getValuesFromPtgs( (Ptg[]) obar[t] );    // get value array from the ptgs
					}
					catch( NumberFormatException e )
					{    // string or non-numeric values
						v = getStringValuesFromPtgs( (Ptg[]) obar[t] );
					}
					obar[t] = v;
				}
				else
				{    // it's a single value
					try
					{
						obar[t] = getDoubleValueFromObject( pval );
					}
					catch( NumberFormatException e )
					{
						if( pval instanceof CalculationException )
						{
							obar[t] = pval.toString();
						}
						else
						{
							obar[t] = pval;
						}
					}
				}
			}
		}
		return obar;
	}

	/**
	 * Returns an array of doubles from number-type ptg's sent in.
	 * This should only be referenced by sub-classes.
	 * <p/>
	 * Null values accessed are treated as 0.  Within excel (empty cell values == 0) Tested!
	 * Sometimes as well you can get empty string values, "".  These are NOT EQUAL ("" != 0)
	 *
	 * @param pthings
	 * @return
	 */
	protected static double[] getValuesFromObjects( Object[] pthings ) throws NumberFormatException
	{
		double[] returnDbl = new double[pthings.length];
		for( int i = 0; i < pthings.length; i++ )
		{

			// Object o = pthings[i].getValue();
			Object o = pthings[i];

			if( o == null )
			{    // NO!! "" is NOT "0", blank is, but not a zero length string.  Causes calc errors, need to handle diff somehow20081103 KSC: don't error out if "" */
				returnDbl[i] = 0.0;
			}
			else if( o instanceof Double )
			{
				returnDbl[i] = (Double) o;
			}
			else if( o instanceof Integer )
			{
				returnDbl[i] = (double) o;
			}
			else if( o instanceof Boolean )
			{    // Excel converts booleans to numbers in calculations 20090129 KSC
				returnDbl[i] = ((Boolean) o ? 1.0 : 0.0);
			}
			else if( o instanceof PtgBool )
			{
				returnDbl[i] = ((Boolean) (((PtgBool) o).getValue()) ? 1.0 : 0.0);
			}
			else if( o instanceof PtgErr )
			{
				// ?
			}
			else
			{
				String s = o.toString();
				Double d = new Double( s );
				returnDbl[i] = d;
			}
		}
		return returnDbl;
	}

	/**
	 * convert a value to a double, throws exception if cannot
	 *
	 * @param o
	 * @return double value if possible
	 * @throws NumberFormatException
	 */
	public static double getDoubleValue( Object o, XLSRecord parent ) throws NumberFormatException
	{
		if( o instanceof Double )
		{
			return (Double) o;
		}
		if( (o == null) || o.toString().equals( "" ) )
		{
			// empty string is interpreted as 0 if show zero values
			if( (parent != null) && parent.getSheet().getWindow2().getShowZeroValues() )
			{
				return 0.0;
			}
			// otherwise, throw error
			throw new NumberFormatException();
		}
		return new Double( o.toString() );    // will throw NumberFormatException if cannot convert
	}

	/**
	 * converts a single Ptg number-type value to a double
	 */
	public static double getDoubleValueFromObject( Object o )
	{
		double ret = 0.0;
		if( o == null )
		{    // 20081103 KSC: don't error out if "" */
			ret = 0.0;
		}
		else if( o instanceof Double )
		{
			ret = (Double) o;
		}
		else if( o instanceof Integer )
		{
			ret = (double) o;
		}
		else if( o instanceof Boolean )
		{    // Excel converts booleans to numbers in calculations 20090129 KSC
			ret = ((Boolean) o ? 1.0 : 0.0);
		}
		else if( o instanceof PtgErr )
		{
			// ?
		}
		else
		{
			String s = o.toString();
			// handle formatted dates from fields like TEXT() calcs
			if( s.indexOf( "/" ) > -1 )
			{
				try
				{
					Calendar c = DateConverter.convertStringToCalendar( s );
					if( c != null )
					{
						ret = DateConverter.getXLSDateVal( c );
					}
				}
				catch( Exception e )
				{//guess not
				}
				;
			}
			if( ret == 0.0 )
			{
				Double d = new Double( s );
				ret = d;
			}
		}
		return ret;
	}

	/**
	 * returns an array of strings from ptg's sent in.
	 * This should only be referenced by sub-classes.
	 */
	protected static String[] getStringValuesFromPtgs( Ptg[] pthings )
	{
		String[] returnStr = new String[pthings.length];
		for( int i = 0; i < pthings.length; i++ )
		{
			if( pthings[i] instanceof PtgErr )
			{
				return new String[]{ "#VALUE!" };    // 20081202 KSC: return error value ala Excel
			}

			Object o = pthings[i].getValue();
			if( o != null )
			{ // 20070215 KSC: avoid nullpointererror
				try
				{    // 20090205 KSC: try to convert numbers to ints when converting to string as otherwise all numbers come out as x.0
					returnStr[i] = String.valueOf( ((Double) o).intValue() );
				}
				catch( Exception e )
				{
					String s = o.toString();
					returnStr[i] = s;
				}
			}
			else
			{
				returnStr[i] = "null"; // 20070216 KSC: Shouldn't match empty string!
			}
		}
		return returnStr;
	}

	/**
	 * if the Ptg needs to keep a handle to a cell, this is it...
	 * tells the Ptg to get it on its own...
	 */
	@Override
	public void updateAddressFromTrackerCell()
	{
		initTrackerCell();
		BiffRec trk = getTrackercell();
		if( trk != null )
		{
			String nad = trk.getCellAddress();
			setLocation( nad );
		}
	}

	/**
	 * if the Ptg needs to keep a handle to a cell, this is it...
	 * tells the Ptg to get it on its own...
	 */
	@Override
	public void initTrackerCell()
	{
		if( getTrackercell() == null )
		{
			try
			{
				BiffRec trk = getParentRec().getSheet().getCell( getLocation() );
				setTrackercell( trk );
			}
			catch( Exception e )
			{
				log.error( "Formula reference could not initialize:" + e.toString() );
			}
		}
	}

	/**
	 * @return Returns the trackercell.
	 */
	@Override
	public BiffRec getTrackercell()
	{
		return trackercell;
	}

	/**
	 * @param trackercell The trackercell to set.
	 */
	@Override
	public void setTrackercell( BiffRec trackercell )
	{
		this.trackercell = trackercell;
	}

	//TODO: PtgRef.isBlank should override!
	@Override
	public boolean isBlank()
	{
		return false;
	}

	/**
	 * return properly quoted sheetname
	 *
	 * @param s
	 * @return
	 */
	public static final String qualifySheetname( String s )
	{
		if( (s == null) || s.equals( "" ) )
		{
			return s;
		}
		try
		{
			if( (s.charAt( 0 ) != '\'') && ((s.indexOf( ' ' ) > -1) || (s.indexOf( '&' ) > -1) || (s.indexOf( ',' ) > -1) || (s.indexOf( '(' ) > -1)) )
			{
				if( s.indexOf( "'" ) == -1 )    // normal case of no embedded ' s
				{
					return "'" + s + "'";
				}
				return "\"" + s + "\"";
			}
		}
		catch( StringIndexOutOfBoundsException e )
		{
		}
		return s;
	}

	/**
	 * return cell address with $'s e.g.
	 * cell AB12 ==> $AB$12
	 * cell Sheet1!C2=>Sheet1!$C$2
	 * Does NOT handle ranges
	 *
	 * @param s
	 * @return
	 */
	public static String qualifyCellAddress( String s )
	{
		String prefix = "";
		if( s.indexOf( "$" ) == -1 )
		{    // it's not qualified yet
			int i = s.indexOf( "!" );
			if( i > -1 )
			{
				prefix = s.substring( 0, i + 1 );
				s = s.substring( i + 1 );
			}
			s = "$" + s;
			i = 1;
			while( (i < s.length()) && !Character.isDigit( s.charAt( i++ ) ) )
			{
				;
			}
			i--;
			if( (i > 0) && (i < s.length()) )
			{
				s = s.substring( 0, i ) + "$" + s.substring( i );
			}
		}
		return prefix + s;
	}

	public static int getArrayLen( Object o )
	{
		int len = 0;
		if( o instanceof double[] )
		{
			len = ((double[]) o).length;
		}
		return len;
	}

	/**
	 * generic reading of a row byte pair with handling for Excel 2007 if necessary
	 *
	 * @param b0
	 * @param b1
	 * @return int row
	 */
	public int readRow( byte b0, byte b1 )
	{
		if( ((parent_rec != null) && !parent_rec.getWorkBook().getIsExcel2007()) )
		{
			int rw = com.extentech.toolkit.ByteTools.readInt( b0, b1, (byte) 0, (byte) 0 );
			if( (rw >= (MAXROWS_BIFF8 - 1)) || (rw < 0) || (this instanceof PtgRefN) )    // PtgRefN's are ALWAYS relative and therefore never over 32xxx
			{
				rw = ByteTools.readShort( b0, b1 );
			}
			return rw;
		}
		// issue when reading Excel2007 rw from bytes as limits exceed ... try to interpret as best one can
		int rw = com.extentech.toolkit.ByteTools.readInt( b0, b1, (byte) 0, (byte) 0 );
		if( rw == 65535 )
		{    // have to assume that this means a wholeCol reference
			rw = -1;
			((PtgRef) this).wholeCol = true;
		}
		return rw;
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	@Override
	public void close()
	{
		parent_rec = null;
		trackercell = null;
		// vars??

	}

} 