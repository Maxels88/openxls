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
package org.openxls.ExtenXLS;

import org.openxls.formats.XLS.Formula;
import org.openxls.formats.XLS.FormulaNotFoundException;
import org.openxls.formats.XLS.FunctionNotSupportedException;
import org.openxls.formats.XLS.OOXMLAdapter;
import org.openxls.formats.XLS.ReferenceTracker;
import org.openxls.formats.XLS.formulas.CalculationException;
import org.openxls.formats.XLS.formulas.FormulaParser;
import org.openxls.formats.XLS.formulas.FunctionConstants;
import org.openxls.formats.XLS.formulas.Ptg;
import org.openxls.formats.XLS.formulas.PtgName;
import org.openxls.formats.XLS.formulas.PtgRef;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Iterator;
import java.util.List;

/**
 * Formula Handle allows for manipulation of Formulas within a WorkBook.
 *
 * @see WorkBookHandle
 * @see WorkSheetHandle
 * @see CellHandle
 */
public class FormulaHandle
{
	private static final Logger log = LoggerFactory.getLogger( FormulaHandle.class );
	public static String[][] getSupportedFunctions()
	{
		return FunctionConstants.recArr;
	}

	private WorkBook bk;

	/**
	 * Sets the location lock on the Cell Reference at the
	 * specified  location
	 * <p/>
	 * Used to prevent updating of the Cell Reference when
	 * Cells are moved.
	 *
	 * @param location of the Cell Reference to be locked/unlocked
	 * @param lock     status setting
	 * @return boolean whether the Cell Reference was found and modified
	 */
	public boolean setLocationLocked( String loc, boolean l )
	{
		int x = Ptg.PTG_LOCATION_POLICY_UNLOCKED;
		if( l )
		{
			x = Ptg.PTG_LOCATION_POLICY_LOCKED;
		}
		return form.setLocationPolicy( loc, x );
	}

	private Formula form;

	/**
	 * Sets the location lock on the Cell Reference at the
	 * specified  location
	 * <p/>
	 * Used to prevent updating of the Cell Reference when
	 * Cells are moved.
	 *
	 * @param location of the Cell Reference to be locked/unlocked
	 * @param lock     status setting
	 * @return boolean whether the Cell Reference was found and modified
	 */
	public boolean setLocationPolicy( String loc, int l )
	{
		return form.setLocationPolicy( loc, l );
	}

	/**
	 * Create a new FormulaHandle from an Excel Formula
	 *
	 * @param Formula - the formula to create a handle for.
	 */
	protected FormulaHandle( Formula f, WorkBook book )
	{
		bk = book;
		form = f;
	}

	/**
	 * Returns the cell Address of the formula
	 */
	public String getCellAddress()
	{
		return form.getCellAddress();
	}

	/**
	 * Returns the Human-Readable Formula String
	 *
	 * @return String the Formula in Human-readable format
	 */
	public String getFormulaString()
	{
		return form.getFormulaString();
	}

	/**
	 * If the Formula evaluates to a String, return
	 * the value as a String.
	 *
	 * @return String - value of the Formula if stored as a String.
	 */
	public String getStringVal() throws FunctionNotSupportedException
	{
		//this.form.init();
		return form.getStringVal();
	}

	/**
	 * Converts a cell value to a form suitable for the public API.
	 * Currently this converts cached errors ({@link org.openxls.formats.XLS.formulas.CalculationException}s)
	 * to the corresponding error string.
	 */
	static Object sanitizeValue( Object val )
	{
		if( val instanceof CalculationException )
		{
			return ((CalculationException) val).getName();
		}
		return val;
	}

	/**
	 * Return the value of the Formula
	 *
	 * @return Object - value of the Formula
	 */
	public Object getVal() throws FunctionNotSupportedException
	{
		return sanitizeValue( form.calculateFormula() );
	}

	/** Return the cached value of the Formula.
	 *
	 *  This method returns the value as cached by ExtenXLS or Excel of the formula.  Please note
	 *  that cases could exist where a cached value does not exist.  In this case getCachedVal will not try and calculate
	 *  the formula, it will return null.

	 @return Object - cached value of the Formula as a String or a Double dependent on data type.

	 public Object getCachedVal() {
	 return form.getCachedVal();
	 } */

	/**
	 * Calculate the value of the formula and return it as an object
	 * <p/>
	 * Calling calculate will ignore the WorkBook formula calculation flags
	 * and forces calculation of the entire formula stack
	 */
	public Object calculate() throws FunctionNotSupportedException
	{
		form.clearCachedValue();

		return sanitizeValue( form.calculate() );
	}

	/**
	 * Sets the formula to a string passed in excel formula format.
	 *
	 * @param formulaString - String formatted as an excel formula, like Sum(A3+4)
	 */

	public void setFormula( String formulaString ) throws FunctionNotSupportedException
	{
		form = FormulaParser.setFormula( form, formulaString, new int[]{ form.getRowNumber(), form.getColNumber() } );
	}

	/**
	 * If the Formula evaluates to a String, there
	 * will be a Stringrec attached to the Formula
	 * which contains the latest value.
	 *
	 * @return boolean whether this Formula evaluates to a String
	 */
	public boolean evaluatesToString()
	{
		return (form.calculateFormula() instanceof String);
	}

	/**
	 * If the Formula evaluates to a float, return
	 * the value as an float.
	 * <p/>
	 * If the workbook level flag CALCULATE_EXPLICIT is set
	 * then the cached value of the formula (if available) will be returned,
	 * otherwise the latest calculated value will be returned
	 *
	 * @return float - value of the Formula if available as a float.  If the
	 * value cannot be returned as a float NaN will be returned.
	 */
	public float getFloatVal() throws FunctionNotSupportedException
	{
		return form.getFloatVal();
	}

	/**
	 * If the Formula evaluates to a double, return
	 * the value as an double.
	 * <p/>
	 * If the workbook level flag CALCULATE_EXPLICIT is set
	 * then the cached value of the formula (if available) will be returned,
	 * otherwise the latest calculated value will be returned
	 *
	 * @return double - value of the Formula if available as a double.  If the
	 * value cannot be returned as a double NaN will be returned.
	 */
	public double getDoubleVal() throws FunctionNotSupportedException
	{
		return form.getDblVal();
	}

	/**
	 * If the Formula evaluates to an int, return
	 * the value as an int.
	 * <p/>
	 * If the workbook level flag CALCULATE_EXPLICIT is set
	 * then the cached value of the formula (if available) will be returned,
	 * otherwise the latest calculated value will be returned
	 *
	 * @return int - value of the Formula if available as a int.  If the value returned can not be
	 * represented by an int or is a float/double with a non-zero mantissa a runtime NumberFormatException
	 * will be thrown
	 */
	public int getIntVal() throws FunctionNotSupportedException
	{
		return form.getIntVal();
	}

	/**
	 * get CellRange strings referenced by this formula
	 *
	 * @return
	 * @throws org.openxls.formats.XLS.FormulaNotFoundException
	 */
	public String[] getRanges() throws FormulaNotFoundException
	{
		Ptg[] locptgs = form.getCellRangePtgs();
		String[] ret = new String[locptgs.length];
		for( int x = 0; x < locptgs.length; x++ )
		{
			// need sheetname along with address; to ensure, must use explicit method:
//			ret[x]=locptgs[x].getTextString();
			try
			{
				ret[x] = ((PtgRef) locptgs[x]).getLocationWithSheet();
			}
			catch( Exception e )
			{
				if( locptgs[x] instanceof PtgName )
				{//avoid NumberFormatExceptions on parsing missing Named Ranges
					ret[x] = locptgs[x].getLocation();
				}
				else
				{
					ret[x] = locptgs[x].getTextString();
				}
			}
		}
		return ret;
	}

	/**
	 * Initialize CellRanges referenced by this formula
	 *
	 * @return
	 * @throws FormulaNotFoundException
	 */
	public CellRange[] getCellRanges() throws FormulaNotFoundException
	{
		String[] crstrs = getRanges();
		CellRange[] crs = new CellRange[crstrs.length];
		for( int x = 0; x < crs.length; x++ )
		{
			crs[x] = new CellRange( crstrs[x], bk, true );
			try
			{
				crs[x].init();
			}
			catch( Exception e )
			{
				//
			}
		}
		return crs;
	}

	/**
	 * Takes a string as a current formula location, and changes
	 * that pointer in the formula to the new string that is sent.
	 * This can take single cells"A5" and cell ranges,"A3:d4"
	 * Returns true if the cell range specified in formulaLoc exists & can be changed
	 * else false.  This also cannot change a cell pointer to a cell range or vice
	 * versa.
	 *
	 * @param String - range of Cells within Formula to modify
	 * @param String - new range of Cells within Formula
	 */
	public boolean changeFormulaLocation( String formulaLoc, String newaddr ) throws FormulaNotFoundException
	{
		List dx = form.getPtgsByLocation( formulaLoc );
		Iterator lx = dx.iterator();
		while( lx.hasNext() )
		{
			try
			{
				Ptg thisptg = (Ptg) lx.next();
				ReferenceTracker.updateAddressPerPolicy( thisptg, newaddr );
				form.setCachedValue( null ); // flag to recalculate
				return true;
			}
			catch( Exception e )
			{
				log.warn( "updating Formula reference failed {} to {}", formulaLoc, newaddr, e );
				return false;
			}
		}
		return true;
	}

	/**
	 * Changes a range in a formula to expand until it includes the
	 * cell address from CellHandle.
	 * <p/>
	 * Example:
	 * <p/>
	 * CellHandle cell = new Cellhandle("D4")  Formula = SUM(A1:B2)
	 * addCellToRange("A1:B2",cell); would change the formula to look like"SUM(A1:D4)"
	 * <p/>
	 * Returns false if formula does not contain the formulaLoc range.
	 *
	 * @param String     - the Cell Range as a String to add the Cell to
	 * @param CellHandle - the CellHandle to add to the range
	 */
	public boolean addCellToRange( String formulaLoc, CellHandle handle ) throws FormulaNotFoundException
	{
		List dx = form.getPtgsByLocation( formulaLoc );
		Iterator lx = dx.iterator();
		boolean b = false;
		while( lx.hasNext() )
		{
			Ptg ptg = (Ptg) lx.next();
			if( ptg == null )
			{
				return false;
			}
			int[] formulaaddr = ExcelTools.getRangeRowCol( formulaLoc );
			String handleaddr = handle.getCellAddress();
			int[] celladdr = ExcelTools.getRowColFromString( handleaddr );

			// check existing range and set new range vals if the new Cell is outside
			if( celladdr[0] > formulaaddr[2] )
			{
				formulaaddr[2] = celladdr[0];
			}
			if( celladdr[0] < formulaaddr[0] )
			{
				formulaaddr[0] = celladdr[0];
			}
			if( celladdr[1] > formulaaddr[3] )
			{
				formulaaddr[3] = celladdr[1];
			}
			if( celladdr[1] < formulaaddr[1] )
			{
				formulaaddr[1] = celladdr[1];
			}
			String newaddr = ExcelTools.formatRange( formulaaddr );
			b = changeFormulaLocation( formulaLoc, newaddr );

		}
		return b;
	}

	/**
	 * Copy the formula references with offsets
	 *
	 * @param int[row,col] offsets to move the references
	 * @return
	 */
	public static void moveCellRefs( FormulaHandle fmh, int[] offsets ) throws FormulaNotFoundException
	{
		// get the current offsets from the FMH references
		String[] celladdys = fmh.getRanges();

		// iterate
		for( String celladdy : celladdys )
		{
			String[] s = ExcelTools.stripSheetNameFromRange( celladdy );
			String sh = s[0];        // sheet portion of address,if any
			String range = s[1];    // range or single address
			int rangeIdx = range.indexOf( ":" );
			String secondAddress = null;
			if( rangeIdx > -1 )
			{    // separate out addresses within a range
				secondAddress = range.substring( rangeIdx + 1 );
				range = range.substring( 0, rangeIdx );
			}
			int[] orig = ExcelTools.getRowColFromString( range );
			boolean relCol = !(range.startsWith( "$" ));    // 20100603 KSC: handle relative refs
			boolean relRow = !((range.length() > 0) && (range.substring( 1 ).indexOf( '$' ) > -1));
			if( relRow )    // only move if relative ref
			{
				orig[0] += offsets[0]; //row
			}
			if( relCol )// only move if relative ref
			{
				orig[1] += offsets[1]; //col
			}
			String newAddress = ExcelTools.formatLocation( orig, relRow, relCol );
			if( (orig[0] < 0) || (orig[1] < 0) )
			{
				newAddress = "#REF!";
			}
			if( secondAddress != null )
			{
				orig = ExcelTools.getRowColFromString( secondAddress );
				relCol = !(secondAddress.startsWith( "$" ));    // handle relative refs
				relRow = !((secondAddress.length() > 0) && (secondAddress.substring( 1 ).indexOf( '$' ) > -1));
				if( orig[0] >= 0 )
				{
					if( relRow )    // only move if relative ref
					{
						orig[0] += offsets[0]; //row
					}
				}
				if( orig[1] >= 0 )
				{ //if not wholerow/wholecol ref
					if( relCol )// only move if relative ref
					{
						orig[1] += offsets[1]; //col
					}
				}
				String newAddress1 = ExcelTools.formatLocation( orig, relRow, relCol );
				newAddress = newAddress + ":" + newAddress1;
			}
			if( sh != null )        // TODO: handle refs with multiple sheets
			{
				newAddress = sh + "!" + newAddress;
			}
			if( !fmh.changeFormulaLocation( celladdy, newAddress ) )
			{
				log.error( "Could not change Formula Reference: " + celladdy + " to: " + newAddress );
			}
		}
		return;
	}

	public String toString()
	{
		return form.getCellAddress() + ":" + form.getFormulaString();
	}

	/**
	 * return truth of "this formula is shared"
	 *
	 * @return boolean
	 */
	public boolean isSharedFormula()
	{
		return form.isSharedFormula();
	}

	public boolean isArrayFormula()
	{
		return form.isArrayFormula();
	}
	// 20090120 KSC: should be detected auto public void setIsArrayFormula(boolean b) { form.setIsArrayFormula(b); }

	/**
	 * returns the low-level formula rec for this Formulahandle
	 *
	 * @return
	 */
	public Formula getFormulaRec()
	{
		return form;
	}

	/**
	 * Utility method to determine if the calculation works out to an error value.
	 * <p/>
	 * The excel values that will cause this to be true are
	 * #VALUE!, #N/A, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!
	 *
	 * @return
	 */
	public boolean isErrorValue()
	{
		return (form.calculateFormula() instanceof CalculationException);
	}

	/**
	 * return the "Calculate Always" setting for this formula
	 * used for formulas that always need calculating such as TODAY
	 *
	 * @return
	 */
	public boolean getCalcAlways()
	{
		return form.getCalcAlways();
	}

	/**
	 * set the "Calculate Always setting for this formula
	 * used for formulas that always need calculating such as TODAY
	 *
	 * @param fAlwaysCalc
	 */
	public void setCalcAlways( boolean fAlwaysCalc )
	{
		form.setCalcAlways( fAlwaysCalc );
	}

	/**
	 * generate the OOXML necessary to describe this formula
	 * OOXML element <f>
	 *
	 * @return
	 */
	// TODO: Deal with External References ... dataTables
	// common possible attributes:
	// aca= always calculate array, bx=name to assign formula to, ca=calculate cell, r1=data table cell1,
	// t=formula type shared, array, dataTable or normal=default
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		// must have type of formula result
		Object val;
		try
		{
			val = getVal();
			if( val == null )
			{// means cache was cleared (in all cases???) MUST recalc
				if( form.getWorkBook().getCalcMode() != WorkBookHandle.CALCULATE_EXPLICIT )
				{
					calculate();
					val = getVal();
				}
				else
				{
					val = new CalculationException( CalculationException.VALUE );
				}
			}
			else if( (val instanceof String) && ((String) val).startsWith( "#" ) )
			{
				val = new CalculationException( CalculationException.getErrorCode( (String) val ) );
			}
		}
		catch( Exception e )
		{
			val = new CalculationException( CalculationException.VALUE );
		}
		if( val == null )
		{
			log.error( "FormulaHandle.getOOXML:  unexpected null encountered when calculating formula: " + getCellAddress() );
		}
		// Handle attributes for special cached values
		if( val instanceof String )
		{
			ooxml.append( " t=\"str\"" );
			val = OOXMLAdapter.stripNonAscii( (String) val );    // TODO: how can we strip non-ascii? What about Japanese cell text?  ans:  has to be XML-compliant is all ...
		}
		else if( val instanceof Boolean )
		{
			ooxml.append( " t=\"b\"" );
			if( (Boolean) val )
			{
				val = "1";
			}
			else
			{
				val = "0";
			}
		}
		else if( val instanceof Double )
		{
			ooxml.append( " t=\"n\"" );
		}
		else if( val instanceof CalculationException )
		{
			ooxml.append( " t=\"e\"" );
		}

		String fs = "=";
		try
		{
			fs = getFormulaString();
		}
		catch( Exception e )
		{
			log.error( "FormulaHandle.getOOXML: error obtaining formula string: " + e.toString(), e );
		}
		fs = OOXMLAdapter.stripNonAscii( fs ).toString();    // handle non-standard xml chars -- ummm what about Japanese? -- it's all ok
		if( !isArrayFormula() )
		{
			ooxml.append( "><f" );
			fs = fs.substring( 1 );    // ignore =
		}
		else
		{    // array formulas
			if( form.getSheet().isArrayFormulaParent( getCellAddress() ) )
			{    // it's the parent
				ooxml.append( "><f" );
				String refs = form.getSheet().getArrayRef( getCellAddress() );
				if( fs.startsWith( "{=" ) )
				{
					fs = fs.substring( 2, fs.length() - 1 );    // remove "{= }"
				}
				ooxml.append( " t=\"array\"" );
				ooxml.append( " ref=\"" + refs + "\"" );
			}
			else
			{    // it's part of a multi-cell array formula therefore DO NOT add array info here
				fs = null;
				ooxml.append( ">" );    // only output value info
			}
		}
		if( isSharedFormula() )
		{
			// TODO: FINISH 00XML SHARED FORMULAS
			// TODO: need si= shared formula index; when referencing (after shared formula is defined) don't need to include "fs" just the si= ****
			try
			{
				// 20091022 KSC: Shared Formulas do not work 2003->2007
				//ooxml.append(" t=\"shared\" ref=\"" + this.getFormulaRec().getSharedFormula().getCellRange() + "\"");
			}
			catch( Exception e )
			{
			}
		}
		if( getCalcAlways() )
		{
			ooxml.append( " ca=\"1\"" );
		}
		if( fs != null )    // can happen if not a parent array formula
		{
			ooxml.append( ">" + fs + "</f>" );
		}
		ooxml.append( "<v>" + val + "</v>" );
		return ooxml.toString();
	}
}