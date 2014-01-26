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
import org.openxls.formats.XLS.Array;
import org.openxls.formats.XLS.Boundsheet;
import org.openxls.formats.XLS.FunctionNotSupportedException;
import org.openxls.formats.XLS.Name;
import org.openxls.formats.XLS.WorkBook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;

/**
 * LookupReferenceCalculator is a collection of static methods that operate
 * as the Microsoft Excel function calls do.
 * <p/>
 * All methods are called with an array of ptg's, which are then
 * acted upon.  A Ptg of the type that makes sense (ie boolean, number)
 * is returned after each method.
 */

public class LookupReferenceCalculator
{
	private static final Logger log = LoggerFactory.getLogger( LookupReferenceCalculator.class );
	/**
	 * ADDRESS
	 * Creates a cell address as text, given specified row and column numbers.
	 * <p/>
	 * Syntax
	 * ADDRESS(row_num,column_num,abs_num,a1,sheet_text)
	 * Row_num   is the row number to use in the cell reference.
	 * Column_num   is the column number to use in the cell reference.
	 * Abs_num   specifies the type of reference to return.
	 * <p/>
	 * Abs_num	 Returns this type of reference
	 * 1 or omitted	 Absolute
	 * <p/>
	 * 2 	 Absolute row; relative column
	 * <p/>
	 * 3	 Relative row; absolute column
	 * <p/>
	 * 4	 Relative
	 * <p/>
	 * A1   is a logical value that specifies the A1 or R1C1 reference style.
	 * If a1 is TRUE or omitted, ADDRESS returns an A1-style reference;
	 * if FALSE, ADDRESS returns an R1C1-style reference.
	 * <p/>
	 * Sheet_text   is text specifying the name of the worksheet to be
	 * used as the external reference. If sheet_text is omitted, no sheet name is used.
	 * <p/>
	 * Examples
	 * <p/>
	 * ADDRESS(2,3) equals "$C$2"
	 * <p/>
	 * ADDRESS(2,3,2) equals "C$2"
	 * <p/>
	 * ADDRESS(2,3,2,FALSE) equals "R2C[3]"
	 * <p/>
	 * ADDRESS(2,3,1,FALSE,"[Book1]Sheet1") equals "[Book1]Sheet1!R2C3"
	 * <p/>
	 * ADDRESS(2,3,1,FALSE,"EXCEL SHEET") equals "'EXCEL SHEET'!R2C3"
	 */
	protected static Ptg calcAddress( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		// deal with floating point refs
		String rx1 = operands[0].getValue().toString();
		if( rx1.indexOf( "." ) > -1 )
		{
			rx1 = rx1.substring( 0, rx1.indexOf( "." ) );
		}
		int row = Integer.valueOf( rx1 );
		// deal with floating point refs
		String cx1 = operands[1].getValue().toString();
		if( cx1.indexOf( "." ) > -1 )
		{
			cx1 = cx1.substring( 0, cx1.indexOf( "." ) );
		}

		int col = Integer.valueOf( cx1 );
		int abs_num = 1;
		boolean ref_style = true;
		String sheettext = "";
		if( operands.length > 2 )
		{
			if( operands[2].getValue() != null )
			{ //checking for a ptgmissarg
				abs_num = (Integer) operands[2].getValue();
			}
		}
		if( operands.length > 3 )
		{
			if( operands[3].getValue() != null )
			{ //checking for a ptgmissarg
				Boolean b = Boolean.valueOf( String.valueOf( operands[3].getValue() ) );
				ref_style = b;
			}
		}
		if( operands.length > 4 )
		{
			if( operands[4].getValue() != null )
			{ //checking for a ptgmissarg
				sheettext = operands[4].getValue() + "!";
			}
		}
		String loc = "";
		String colstr = ExcelTools.getAlphaVal( col - 1 );
		if( ref_style )
		{
			if( abs_num == 1 )
			{
				loc = "$" + colstr + "$" + row;
			}
			else if( abs_num == 2 )
			{
				loc = colstr + "$" + row;
			}
			else if( abs_num == 3 )
			{
				loc = "$" + colstr + row;
			}
			else if( abs_num == 4 )
			{
				loc = colstr + row;
			}
		}
		else
		{
			if( abs_num == 1 )
			{
				loc = "R" + row + "C" + col; // this is transposed with abs_num 4.  Error in Excel
			}
			else if( abs_num == 2 )
			{
				loc = "R" + row + "C[" + col + "]";
			}
			else if( abs_num == 3 )
			{
				loc = "R[" + row + "]C" + col;
			}
			else if( abs_num == 4 )
			{
				loc = "R[" + row + "]C[" + col + "]";

			}
		}
		loc = sheettext + loc;
		return new PtgStr( loc );

	}

	/**
	 * AREAS
	 * Returns the number of areas in a reference. An area is a range of contiguous cells or a single cell.
	 * <p/>
	 * Reference   is a reference to a cell or range of cells and can refer to multiple areas.
	 * If you want to specify several references as a single argument, then you must include extra sets of parentheses so that Microsoft Excel will not interpret the comma as a field separator.
	 * <p/>
	 * NOTE: this appears to be correct given Excel information but logic is not 100% known
	 * e.g. =AREAS(B2:D4 D3) = 1
	 * =AREAS(B2:D4 E3) gives a #NULL! error - why?
	 * =AREAS(B2:D4,E3) = 2
	 */
	protected static Ptg calcAreas( Ptg[] operands )
	{
		Ptg ref = operands[0];
		String s = ref.toString();
		String[] areas = s.split( ",(?=([^'|\"]*'[^'|\"]*'|\")*[^'|\"]*$)" );
		return new PtgNumber( areas.length );
	}

	/**
	 * CHOOSE
	 * Chooses a value from a list of values
	 * <p/>
	 * Note, this function does not support one specific use-case.  That is choosing a ptgref
	 * and using that ptgref to complete a ptgarea.  Example
	 * =SUM(E6:CHOOSE(3,G4,G5,G6))
	 * =SUM(E6:G6)
	 */
	protected static Ptg calcChoose( Ptg[] operands )
	{
		if( operands.length < 2 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Object o = operands[0].getValue();
		try
		{
			Double dd = new Double( o.toString() ); // this can be non-integer, so truncate it if so...
			double e = dd;
			int i = (int) e;
			if( (i > (operands.length + 1)) || (i < 1) )
			{
				return new PtgErr( PtgErr.ERROR_REF );
			}
			o = operands[i].getValue();
			Double d = (Double) o;
			return new PtgNumber( d );
		}
		catch( Exception ex )
		{
			new PtgErr( PtgErr.ERROR_VALUE );
		}
		;
		return new PtgStr( o.toString() );

	}

	/**
	 * COLUMN
	 * Returns the column number of a reference
	 */
	protected static Ptg calcColumn( Ptg[] operands )
	{
		if( operands[0] instanceof PtgFuncVar )
		{
			// we need to return the col where the formula is.
			PtgFuncVar pfunk = (PtgFuncVar) operands[0];
			try
			{
				int loc = pfunk.getParentRec().getColNumber();
				loc += 1;
				return new PtgInt( loc );
			}
			catch( Exception e )
			{
			}
			;
		}
		else
		{
			// It's ugly, but we are going to handle the four types of references seperately, as there is no good way
			// to generically get this info
			try
			{
				if( operands[0] instanceof PtgArea )
				{
					PtgArea pa = (PtgArea) operands[0];
					int[] loc = pa.getIntLocation();
					return new PtgInt( loc[1] + 1 );
				}
				if( operands[0] instanceof PtgRef )
				{
					PtgRef pref = (PtgRef) operands[0];
					int loc = pref.getIntLocation()[1];
					loc += 1;
					return new PtgInt( loc );
				}
				if( operands[0] instanceof PtgName )
				{    // table???
					String range = ((PtgName) operands[0]).getName().getLocation();
					int[] loc = ExcelTools.getRangeCoords( range );
					return new PtgInt( loc[1] + 1 );
				}
			}
			catch( Exception e )
			{
			}
			;
		}
		return new PtgInt( -1 );
	}

	/**
	 * COLUMNS
	 * Returns the number of columns in an array reference or array formula
	 */
	// TODO: Not finished yet!
	protected static Ptg calcColumns( Ptg[] operands )
	{
		//
		if( operands[0] instanceof PtgFuncVar )
		{
			// we need to return the col where the formula is.
			PtgFuncVar pfunk = (PtgFuncVar) operands[0];
			try
			{
				int loc = pfunk.getParentRec().getColNumber();
				loc += 1;
				return new PtgInt( loc );
			}
			catch( Exception e )
			{
			}
			;
		}
		else
		{
			// It's ugly, but we are going to handle the four types of references seperately, as there is no good way
			// to generically get this info
			try
			{
				if( operands[0] instanceof PtgArea )
				{
					PtgArea pa = (PtgArea) operands[0];
					int[] loc = pa.getIntLocation();
					int ncols = (loc[3] - loc[1]) + 1;
					return new PtgInt( ncols );
				}
				if( operands[0] instanceof PtgArray )
				{
					PtgArray parr = (PtgArray) operands[0];
					return new PtgInt( parr.nc + 1 );
				}
			}
			catch( Exception e )
			{
			}
			;
		}
		return new PtgInt( -1 );
	}

	/**
	 * HLOOKUP: Looks in the top row of an array and returns the value of the
	 * indicated cell.
	 */
	protected static Ptg calcHlookup( Ptg[] operands ) throws FunctionNotSupportedException
	{

		if( operands.length < 3 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		boolean sorted = true;
		boolean isNumber = true;
		Ptg lookup_value = operands[0];
		Ptg table_array = operands[1];
		PtgInt row_index_num = (PtgInt) operands[2];
		int rowNum = row_index_num.getVal() - 1;// reduce 1 for ordinal base off firstcol
		if( operands.length > 3 )
		{
			if( operands[3].getValue() != null )
			{
				Object o = operands[3].getValue();
				if( o instanceof Boolean )
				{
					sorted = (Boolean) o;
				}
				else if( o instanceof Integer )
				{
					sorted = ((Integer) o != 0);
				}
			}
		}
		int[] retarea = { 0, 0 };
		Boundsheet bs = lookup_value.getParentRec().getSheet();
		WorkBook bk = table_array.getParentRec().getWorkBook();
		PtgRef[] lookupComponents = null;
		PtgRef[] valueComponents = null;
		// first, get the lookup Column Vals
		if( table_array instanceof PtgName )
		{
			// Handle getting vals out of name

		}
		else if( (table_array instanceof PtgArea) || (table_array instanceof PtgArea3d) )
		{
			try
			{
				PtgArea pa = (PtgArea) table_array;
				int[] range = table_array.getIntLocation();

				//				 TODO: check rc sanity here
				int firstrow = range[0];
				lookupComponents = (PtgRef[]) pa.getRowComponents( firstrow );
				valueComponents = (PtgRef[]) pa.getRowComponents( firstrow + rowNum );
			}
			catch(Exception e )
			{
				log.warn( "Error in LookupReferenceCalculator: Cannot determine PtgArea location. " + e, e );
			}

		}
		// error check
		if( (lookupComponents == null) || (lookupComponents.length == 0) )
		{
			return new PtgErr( PtgErr.ERROR_REF );
		}
		// lets check if we are dealing with strings or numbers....
		try
		{
			String val = lookupComponents[0].getValue().toString();
			Double d = new Double( val );
		}
		catch( NumberFormatException e )
		{
			isNumber = false;
		}

		if( isNumber )
		{
			double match_num;
			try
			{
				match_num = Double.parseDouble( lookup_value.getValue().toString() );
			}
			catch( NumberFormatException e )
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}

			for( int i = 0; i < lookupComponents.length; i++ )
			{
				double val;
				try
				{
					val = Double.parseDouble( lookupComponents[i].getValue().toString() );
				}
				catch( NumberFormatException e )
				{
					// Ignore entries in the table that aren't numbers.
					continue;
				}

				if( val == match_num )
				{
					return valueComponents[i].getPtgVal();
				}
				if( sorted && (val > match_num) )
				{
					if( i == 0 )
					{
						return new PtgErr( PtgErr.ERROR_NA );
					}
					return valueComponents[i - 1].getPtgVal();
				}
			}

			if( sorted )
			{
				return valueComponents[lookupComponents.length - 1].getPtgVal();
			}
			return new PtgErr( PtgErr.ERROR_NA );

		}
		//TODO: need to handle as string

		return new PtgErr( PtgErr.ERROR_NULL );
	}

	/**
	 * HYPERLINK
	 * Creates a shortcut or jump that opens a document
	 * stored on a network server, an intranet, or the Internet
	 * <p/>
	 * Function just returns the "friendly name" of the link,
	 * Excel doesn't appear to validate the url ...
	 */

	protected static Ptg calcHyperlink( Ptg[] operands )
	{
		try
		{
			if( operands.length == 2 )
			{
				return new PtgStr( operands[1].getValue().toString() );
			}
			return new PtgStr( operands[0].getValue().toString() );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
	}

	/**
	 * INDEX
	 * Returns a value or the reference to a value from within a table or range.
	 * There are two forms of the INDEX function: the array form and the reference form.
	 * <p/>
	 * Array Form:
	 * Returns the value of an element in a table or an array selected by the row and column number indexes.
	 * Use the array form if the first argument to INDEX is an array constant.
	 * INDEX(array,row_num,column_num)
	 * Array   is a range of cells or an array constant.
	 * If array contains only one row or column,
	 * the corresponding row_num or column_num argument is optional.
	 * If array has more than one row and more than one column, and only row_num or column_num is used,
	 * INDEX returns an array of the entire row or column in array.
	 * Row_num   selects the row in array from which to return a value. If row_num is omitted, column_num is required.
	 * Column_num   selects the column in array from which to return a value. If column_num is omitted, row_num is required.
	 * <p/>
	 * Reference Form:
	 * Returns the reference of the cell at the intersection of a particular row and column.
	 * If the reference is made up of nonadjacent selections, you can pick the selection to look in.
	 * INDEX(reference,row_num,column_num,area_num)
	 * Reference   is a reference to one or more cell ranges.
	 * If you are entering a nonadjacent range for the reference, enclose reference in parentheses.
	 * If each area in reference contains only one row or column, the row_num or column_num argument, respectively, is optional. For example, for a single row reference, use INDEX(reference,,column_num).
	 * Row_num   is the number of the row in reference from which to return a reference.
	 * Column_num   is the number of the column in reference from which to return a reference.
	 * Area_num   selects a range in reference from which to return the intersection of row_num and column_num. The first area selected or entered is numbered 1, the second is 2, and so on. If area_num is omitted, INDEX uses area 1.
	 * <p/>
	 * Given a BiffRec Range, choose the cell within referenced by the
	 * row and column operands.
	 * <p/>
	 * example:
	 * =INDEX(G3:L8,6,4)
	 * returns the value of the BiffRec at row 6, col 4 in the following table
	 * which is 0.11
	 * <p/>
	 * G		H		I		J		K		L
	 * 1
	 * 2
	 * 3		0.007	0.005	0.003	0.002	0.002	0.001
	 * 4		0.025	0.017	0.012	0.008	0.006	0.005
	 * 5		0.062	0.044	0.032	0.023	0.018	0.015
	 * 6		0.116	0.086	0.064	0.049	0.04	0.035
	 * 7		0.171	0.13	0.101	0.082	0.07	0.062
	 * 8		0.211	0.165	0.132	0.11	0.096	0.088
	 */
	protected static Ptg calcIndex( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg o = operands[0];
		//  only 1st operand is required; changed below (see doc snippet below) 
		/*If each area in reference contains only one row or column, 
		 * the row_num or column_num argument, respectively, is optional. 
		 * For example, for a single row reference, use INDEX(reference,,column_num). 
		 */
		Object rowref = (double) 1;    // defaults (1-based)
		Object colref = (double) 1;
		double areanum = 1;
		int[] retarea = null;
		String sht = null;
		try
		{
			if( operands.length > 1 )
			{
				Ptg rowrefp = operands[1];
				rowref = rowrefp.getValue();
			}
			if( operands.length > 2 )
			{
				Ptg colrefp = operands[2];
				colref = colrefp.getValue();
			}
			if( operands.length > 3 )
			{
				areanum = operands[3].getDoubleVal();
			}
			if( o instanceof PtgArea )
			{
				retarea = ((PtgArea) o).getIntLocation();
				sht = ((PtgArea) o).getSheetName();
			}
			else if( o instanceof PtgName )
			{
				//Ptg[] p=((PtgName) o).getName().getComponents();	//CellRangePtgs();				
				//o= p[0];
				String r = ((PtgName) o).getName().getLocation();
				retarea = ExcelTools.getRangeRowCol( r );
				sht = (r.indexOf( "!" ) == -1) ? null : r.substring( 0, r.indexOf( "!" ) );
			}
			else if( o instanceof PtgMemFunc )
			{
				Ptg[] ps = o.getComponents();
				areanum--;
				if( (areanum >= 0) && (areanum < ps.length) )
				{
					o = ps[(int) areanum];
				}
				else
				{
					o = ps[0];
				}
				retarea = ((PtgArea) o).getIntLocation();
				sht = ((PtgArea) o).getSheetName();
			}
			else if( o instanceof PtgArray )
			{
				Ptg[] ps = o.getComponents();
				areanum--;
				if( (areanum >= 0) && (areanum < ps.length) )
				{
					o = ps[(int) areanum];
				}
				else
				{
					o = ps[0];
				}
				retarea = ((PtgArea) o).getIntLocation();
				// TODO: Sheet!
			}

			// now should have the correct area
			if( retarea != null )
			{
				// really we can just use the first position to get a ref
				// then the second just checks bounds				
//				 TODO: check rc sanity here
				int rowoff = retarea[0];
				int coloff = retarea[1];
				int rowck = retarea[2] + 1;
				int colck = retarea[3] + 1;

				int[] dims = new int[2];
				try
				{
					int rr;
					if( (rowref instanceof Integer) )
					{
						rr = (Integer) rowref;
					}
					else
					{
						String rw = rowref.toString();
						if( rw.indexOf( "." ) > -1 )
						{
							rw = rw.substring( 0, rw.indexOf( "." ) );
						}
						// string non Integer Chars...
						rr = Integer.parseInt( rw );
					}
					int cr;
					if( (colref instanceof Integer) )
					{
						cr = (Integer) colref;
					}
					else
					{
						String cl = colref.toString();
						if( cl.indexOf( "." ) > -1 )
						{
							cl = cl.substring( 0, cl.indexOf( "." ) );
						}
						// string non Integer Chars...
						cr = Integer.parseInt( cl );
					}

					if( (rr > rowck) || (cr > colck) )
					{
						return new PtgErr( PtgErr.ERROR_REF );
					}
					dims[0] = rr + rowoff - 1;
					dims[1] = cr + coloff - 1;

					// here's a nice new ref...
					PtgRef refp = new PtgRef();
					if( o instanceof PtgArea3d )
					{
						refp = new PtgRef3d();
					}
					refp.setParentRec( o.getParentRec() );
					refp.setUseReferenceTracker( false );
					if( sht != null )
					{
						refp.setSheetName( sht );
					}
					refp.setLocation( dims );
					if( o instanceof PtgArea3d )
					{
						refp.setLocation( ((PtgArea3d) o).getSheetName() + "!" + ExcelTools.formatLocation( dims ) );
					}
					else
					{

					}
					return refp;

				}
				catch( NumberFormatException e )
				{
					//Logger.logWarn("could not calculate INDEX function: " + o.toString() + ":" + e);
					return new PtgErr( PtgErr.ERROR_NULL );    // ERR or #VALUE ??
				}
			}
		}
		catch( Exception e )
		{
			log.warn( "could not calculate INDEX function: " + o.toString() + ":" + e, e );
		}
		return new PtgErr( PtgErr.ERROR_NULL );

	}

	/**
	 * INDIRECT
	 * Returns a reference indicated by a text value
	 * <p/>
	 * INDIRECT(ref_text,a1)
	 * <p/>
	 * Ref_text   is a reference to a cell that contains an A1-style
	 * reference, an R1C1-style reference, a name defined as a reference,
	 * or a reference to a cell as a text string. If ref_text is not a valid
	 * cell reference, INDIRECT returns the #REF! error value.
	 * <p/>
	 * If ref_text refers to another workbook (an external reference),
	 * the other workbook must be open. If the source workbook is not open,
	 * INDIRECT returns the #REF! error value.
	 * <p/>
	 * A1   is a logical value that specifies what type of reference is contained in the cell ref_text.
	 * <p/>
	 * If a1 is TRUE or omitted, ref_text is interpreted as an A1-style reference.
	 * If a1 is FALSE, ref_text is interpreted as an R1C1-style reference.
	 */
	protected static Ptg calcIndirect( Ptg[] operands )
	{
		try
		{
			if( operands[0] instanceof PtgStr )
			{
				PtgStr ps = (PtgStr) operands[0];

				String locx = ps.toString();
				// detect range
				if( !(FormulaParser.isRef( locx ) || FormulaParser.isRange( locx )) )
				{    // see if it's a named range
					Name nmx = ps.getParentRec().getWorkBook().getName( locx );
					if( nmx != null )
					{ // there is a named range
						locx = nmx.getLocation();
					}
					else
					{
						return ps;    //it's just a value
					}
				}
				if( "".equals( locx ) )
				{
					return new PtgInt( 0 );    // that's what Excel does!
				}
				PtgArea3d refp = new PtgArea3d( false );
				refp.setParentRec( ps.getParentRec() );
				refp.setUseReferenceTracker( true );    // very important!!! :)
				refp.setLocation( locx );
				return refp;

			}
			if( operands[0] instanceof PtgRef )
			{
				// check if the ptgRef value is a string representing a Named range
				Object o = operands[0].getValue();
				PtgStr ps = new PtgStr( o.toString() );
				ps.setParentRec( operands[0].getParentRec() );
				operands = new Ptg[1];
				operands[0] = ps;
				return calcIndirect( operands );
			}
			if( operands[0] instanceof PtgName )
			{
				return calcIndirect( operands[0].getComponents() );
			}
		}
		catch( Exception e )
		{
			//log.error("INDIRECT: " + e.toString());
		}
		return new PtgErr( PtgErr.ERROR_REF );    // 's what Excel does ...
	}

	/**
	 * LOOKUP
	 * The LOOKUP function returns a value either from a one-row or one-column range
	 * You can also use the LOOKUP function as an alternative to the IF function for elaborate tests or
	 * tests that exceed the limit for nesting of functions. See the examples in the array form.
	 * For the LOOKUP function to work correctly, the data being looked up must be sorted in
	 * ascending order. If this is not possible, consider using the VLOOKUP, HLOOKUP, or MATCH functions.
	 * <p/>
	 * A vector is a range of only one row or one column.
	 * The vector form of LOOKUP looks in a one-row or one-column range (known as a vector) for a value and returns a value from the same position in a second one-row or one-column range. Use this form of the LOOKUP function when you want to specify the range that contains the values that you want to match. The other form of LOOKUP automatically looks in the first column or row.
	 */
	public static Ptg calcLookup( Ptg[] operands )
	{
		String lookup = operands[0].getValue().toString().toUpperCase();
		if( operands.length > 2 )
		{ //normal version of lookup
			Ptg[] vector = operands[1].getComponents();
			Ptg[] returnvector = operands[2].getComponents();
			if( returnvector == null ) // happens when operands[2] is a PtgRef
			{
				return new PtgNumber( 0 );    // this is what excel does
			}

			//If the LOOKUP function can't find the lookup_value, the function matches the largest value in lookup_vector that is less than or equal to lookup_value.
			//If lookup_value is smaller than the smallest value in lookup_vector, LOOKUP returns the #N/A error value
			Object retval = null;
			for( int i = 0; i < vector.length; i++ )
			{
				if( Calculator.compareCellValue( vector[i].getValue(), lookup, ">" ) )
				{
					break;
				}
				if( i < returnvector.length )
				{
					retval = returnvector[i].getValue();
				}
			}
			if( retval instanceof Number )
			{
				return new PtgNumber( ((Number) retval).doubleValue() );
			}
			if( retval instanceof Boolean )
			{
				return new PtgBool( (Boolean) retval );
			}
			if( retval == null )
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
			// assume string

			return new PtgStr( retval.toString() );
		} //array form of lookup
	/*
		 *  The array form of LOOKUP looks in the first row or column of an array for the specified value
    	 *  and returns a value from the same position in the last row or column of the array.
    	 *  Use this form of LOOKUP when the values that you want to match are in the first row
    	 *  or column of the array. Use the other form of LOOKUP when you want to specify the location
    	 *  of the column or row.
			In general, it's best to use the HLOOKUP or VLOOKUP function instead of the array form of LOOKUP. This form of LOOKUP is provided for compatibility with other spreadsheet programs.

		With the HLOOKUP and VLOOKUP functions, you can index down or across, but LOOKUP always selects the last value in the row or column.

 */
		try
		{
			Ptg[] array = operands[1].getComponents();
			int nrs = ((PtgArray) operands[1]).getNumberOfRows();
			int ncs = ((PtgArray) operands[1]).getNumberOfColumns();
			//If array covers an area that is wider than it is tall (more columns than rows), LOOKUP searches for the value of lookup_value in the first row.
			//If an array is square or is taller than it is wide (more rows than columns), LOOKUP searches in the first column.
			Object retval = null;
			boolean found = false;
			boolean rowbased = (ncs > nrs);
			ncs++; // make 1-based
			int i = 0;
			for( int j = 0; (j < nrs) && !found; j++ )
			{
				int start = i;
				for(; (i < (start + ncs)) && !found; i++ )
				{
					if( Calculator.compareCellValue( array[i].getValue(), lookup, ">" ) )
					{
						found = true;
						break;
					}
					// returns a value from the same position in the last row or column of the array
					if( rowbased )
					{
						retval = array[i + ncs].getValue();
					}
					else
					{
						retval = array[i + 1].getValue();
						i++;
					}
				}
			}
			if( retval instanceof Number )
			{
				return new PtgNumber( ((Number) retval).doubleValue() );
			}
			if( retval instanceof Boolean )
			{
				return new PtgBool( (Boolean) retval );
			}
			if( retval == null )
			{
				return new PtgErr( PtgErr.ERROR_NA );
			}
			// assume string

			return new PtgStr( retval.toString() );
		}
		catch( Exception e )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
	}

	/**
	 * MATCH
	 * Looks up values in a reference or array
	 * Returns the relative position of an item in an array that matches a specified value in a
	 * specified order. Use MATCH instead of one of the LOOKUP functions when you need the position
	 * of an item in a range instead of the item itself.
	 * <p/>
	 * MATCH(lookup_value,lookup_array,match_type)
	 * <p/>
	 * Lookup_value is the value you want to match in lookup_array. For example, when you look up someone's number
	 * in a telephone book, you are using the person's name as the lookup value, but the telephone number is
	 * the value you want.
	 * Lookup_value can be a value (number, text, or logical value) or a cell reference to a number,
	 * text, or logical value.
	 * <p/>
	 * Lookup_array   is a contiguous range of cells containing possible lookup values. Lookup_array must
	 * be an array or an array reference.
	 * <p/>
	 * Match_type   is the number -1, 0, or 1. Match_type specifies how Microsoft Excel matches lookup_value
	 * with values in lookup_array.
	 * <p/>
	 * If match_type is 1, MATCH finds the largest value that is less than or equal to lookup_value.
	 * Lookup_array must be placed in ascending order: ...-2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE.
	 * <p/>
	 * If match_type is 0, MATCH finds the first value that is exactly equal to lookup_value.
	 * Lookup_array can be in any order.
	 * <p/>
	 * If match_type is -1, MATCH finds the smallest value that is greater than or equal to lookup_value.
	 * Lookup_array must be placed in descending order: TRUE, FALSE, Z-A, ...2, 1, 0, -1, -2, ..., and so on.
	 * <p/>
	 * If match_type is omitted, it is assumed to be 1.
	 * <p/>
	 * MATCH returns the position of the matched value within lookup_array, not the value itself. For example, MATCH("b",{"a","b","c"},0) returns 2, the relative position of "b" within the array {"a","b","c"}.
	 * MATCH does not distinguish between uppercase and lowercase letters when matching text values.
	 * If MATCH is unsuccessful in finding a match, it returns the #N/A error value.
	 * If match_type is 0 and lookup_value is text, lookup_value can contain the wildcard characters asterisk (*) and question mark (?). An asterisk matches any sequence of characters; a question mark matches any single character.
	 */
	public static Ptg calcMatch( Ptg[] operands )
	{
		try
		{
			Object lookupValue = operands[0].getValue();    // should be one value or a reference
			Ptg lookupArray = operands[1];    // array or array reference (PtgArea)
			Ptg[] values = null;
			int matchType = 1;
			if( operands.length > 2 )
			{
				Object o = operands[2].getValue();
				if( o instanceof Integer )
				{
					matchType = (Integer) o;
				}
				else
				{
					matchType = ((Double) o).intValue();
				}

			}
			// Step 1- get all the components of the lookupArray (Array or Array Reference
			if( lookupArray instanceof PtgName )
			{
				PtgArea3d pa = new PtgArea3d( false );
				pa.setParentRec( lookupArray.getParentRec() );
				pa.setLocation( ((PtgName) lookupArray).getName().getLocation() );
				lookupArray = pa;
			}
			if( lookupArray instanceof PtgArea )
			{
				PtgArea pa = (PtgArea) lookupArray;
				values = pa.getComponents();
			}
			else if( lookupArray instanceof PtgMemFunc )
			{
				PtgMemFunc pa = (PtgMemFunc) lookupArray;
				values = pa.getComponents();
			}
			else if( lookupArray instanceof PtgArray )
			{
				PtgArray pa = (PtgArray) lookupArray;
				values = pa.getComponents();
			}
			else if( lookupArray instanceof PtgMystery )
			{
				// PtgMystery is return from PtgMemFunc/MemArrays
				ArrayList ptgs = new ArrayList();
				Ptg[] p = ((PtgMystery) lookupArray).vars;
				for( Ptg aP : p )
				{
					if( aP instanceof PtgArea )
					{
						Ptg[] pa = aP.getComponents();
						for( Ptg aPa : pa )
						{
							ptgs.add( aPa );
						}
					}
					else
					{
						ptgs.add( aP );
					}
				}
				values = new Ptg[ptgs.size()];
				ptgs.toArray( values );
			}
			else if( lookupArray instanceof PtgStr )
			{
				PtgArea3d pa = new PtgArea3d( false );
				pa.setParentRec( lookupArray.getParentRec() );
				pa.setLocation( lookupArray.toString() );
				values = pa.getComponents();
			}
			else
			{ // testing!
				log.error( "match: unknown type of lookup array" );
			}

			// Step # 2- traverse thru value array to find lookupValue using matchType rules
			// ALSO must ensure for matchType!=0 that array is in ascending or descending order
			int retIndex = -1;
			// TODO: matchType==0 Strings can match wildcards ...
			for( int i = 1; i <= values.length; i++ )
			{
				Object v0 = values[i - 1].getValue();
				Object v1 = null;
				if( i < values.length )
				{
					v1 = values[i].getValue();
				}
				int mType = -2; // -1 means v0<v1, 0 means v0==v1, 1 means v0>v1
				int match = -2;    // test lookupValue against v0 (i-1)
				if( v0 instanceof Integer )
				{
					if( v1 != null )
					{
						mType = (((Integer) v0).compareTo( (Integer) v1 ));
					}
					match = (((Integer) v0).compareTo( (Integer) lookupValue ));
				}
				else if( v0 instanceof Double )
				{
					if( v1 != null )
					{
						mType = (((Double) v0).compareTo( (Double) v1 ));
					}
					match = (((Double) v0).compareTo( (Double) lookupValue ));
				}
				else if( v0 instanceof Boolean )
				{
					boolean bv0 = (Boolean) v0;
					// 1.6 only if (v1!=null) mType= (((Boolean) v0).compareTo((Boolean)v1)); 
					if( v1 != null )
					{
						boolean bv1 = (Boolean) v1;
						mType = ((bv0 == bv1) ? 0 : ((!bv0 && bv1) ? -1 : +1));
					}
					// 1.6 only match= (((Boolean) v0).compareTo((Boolean)lookupValue));
					boolean bv1 = (Boolean) lookupValue;
					match = ((bv0 == bv1) ? 0 : ((!bv0 && bv1) ? -1 : +1));
				}
				else if( v0 instanceof String )
				{
					if( v1 != null )
					{
						mType = (((String) v0).compareTo( (String) v1 ));
					}
					match = (((String) v0).compareTo( (String) lookupValue ));
				}
				if( i < values.length )
				{ // only check order
					if( ((matchType == 1) && (mType > 0))// not in ascending order
							|| ((matchType == -1) && (mType < 0)) ) // not in descending order
// DOCUMENTATION SEZ MUST BE IN DESCENDING ORDER FOR -1 BUT EXCEL ALLOWS IT IN CERTAIN CIRCUMSTANCES
//						)
					{
						return new PtgErr( PtgErr.ERROR_NA );
					}
				}
				if( (matchType == 0) && (match == 0) )
				{
					retIndex = i;    // 1-based
					break;
				}
				if( (matchType == 1) && (match <= 0) )
				{
					retIndex = i;    // 1-based
				}
				else if( (matchType == -1) && (match >= 0) )
				{
					retIndex = i;    // 1-based
				}
			}
/*			
			for (int i= 0; i < values.length; i++) {
				Object val=values[i].getValue();
				int mType= -2;
				if (val instanceof Integer)
					mType= ((Integer)val).compareTo((Integer)lookupValue);
				else if (val instanceof Double)
					mType= ((Double)val).compareTo((Double)lookupValue);
				else if (val instanceof Boolean)
					mType= ((Boolean)val).compareTo((Boolean)lookupValue);
				else if (val instanceof String)
					mType= ((String)val).toLowerCase().compareTo(((String)lookupValue).toLowerCase());
				if (matchType==0) {// matches if equal
					if (mType==0) {// then got it!
						retIndex= i;
						break;
					}		
				} else if (matchType==-1) {
				} else  {	// default is 1
				}
			}
*/
			if( retIndex > -1 )
			{
				return new PtgInt( retIndex );
			}
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_NA );
	}

	/**
	 * calcOffset
	 * Returns a reference to a range that is a specified number of rows and columns from a cell or range of cells.
	 * The reference that is returned can be a single cell or a range of cells. You can specify the number of rows and the number of columns to be returned.
	 * <p/>
	 * OFFSET(reference,rows,cols,height,width)
	 * <p/>
	 * Reference   is the reference from which you want to base the offset. Reference must refer to a cell or range of adjacent cells; otherwise, OFFSET returns the #VALUE! error value.
	 * Rows   is the number of rows, up or down, that you want the upper-left cell to refer to. Using 5 as the rows argument specifies that the upper-left cell in the reference is five rows below reference. Rows can be positive (which means below the starting reference) or negative (which means above the starting reference).
	 * Cols   is the number of columns, to the left or right, that you want the upper-left cell of the result to refer to. Using 5 as the cols argument specifies that the upper-left cell in the reference is five columns to the right of reference. Cols can be positive (which means to the right of the starting reference) or negative (which means to the left of the starting reference).
	 * Height   is the height, in number of rows, that you want the returned reference to be. Height must be a positive number.
	 * Width   is the width, in number of columns, that you want the returned reference to be. Width must be a positive number.
	 */
	protected static Ptg calcOffset( Ptg[] operands )
	{
		Ptg ref = operands[0];    // Reference must refer to a cell or range of adjacent cells; otherwise, OFFSET returns the #VALUE! error value
		if( !(ref instanceof PtgRef) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		int nrows = operands[1].getIntVal();
		int ncols = operands[2].getIntVal();
		int height = -1;    // Height   if present, how many rows to return - must be positive
		if( operands.length > 3 )
		{
			height = operands[3].getIntVal();
			if( height < 0 )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
		}
		int width = -1;    // Width   	if present, how many columns to return - must be positive
		if( operands.length > 4 )
		{
			width = operands[4].getIntVal();
			if( width < 0 )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
		}
		int[] rc = ((PtgRef) ref).getIntLocation();
		rc[0] += nrows;
		rc[1] += ncols;
		if( rc.length > 3 )
		{ // it's an area
			rc[2] += nrows;
			rc[3] += ncols;
		}
		// A height/width of 1, 1= a single reference
		// When we increase either the row height or column width in the offset function "=OFFSET(A1,2,0,1,1)"
		// to more than 1, the reference is converted to a range.
		// OK, this may be the wrong interpretation of the height and width, but, from research, this is what I've come up with:
		// height and width are only really applicable for an initial single reference
		// in Excel, a height and width value of more than 1 returns #VALUE! unless wrapped in SUM
		// ????
		if( (height == 1) && (width == 1) )
		{ // it's a single reference
			if( rc.length > 3 )
			{    // truncate
				int[] temp = new int[2];
				System.arraycopy( rc, 0, temp, 0, 2 );
				rc = temp;
			}
		}
		else if( !((height == -1) && (width == -1)) )
		{
			if( rc.length < 3 )
			{ // make a range
				int[] temp = new int[4];
				System.arraycopy( rc, 0, temp, 0, 2 );
				rc = temp;
			}
			if( height > 0 )
			{
				rc[2] = rc[0] + height - 1;    // is this correct????
			}
			if( width > 0 )
			{
				rc[3] = rc[1] + width - 1;        // " "
			}
		}
		if( rc.length > 3 )
		{
			// If rows and cols offset reference over the edge of the worksheet, OFFSET returns the #REF! error value.
			if( (rc[0] < 0) || (rc[1] < 0) || (rc[2] < 0) || (rc[3] < 0) )
			{
				return new PtgErr( PtgErr.ERROR_REF );
			}
			PtgArea pa = new PtgArea( false );
			pa.setParentRec( ref.getParentRec() );
			try
			{
				String sh = ref.getLocation();
				int z = sh.indexOf( '!' );
				if( z > 0 )
				{
					pa.setSheetName( sh.substring( 0, z ) );
				}
			}
			catch( Exception e )
			{
				;
			}
			pa.setLocation( rc );
			return pa;
		} // it's a single reference
		// If rows and cols offset reference over the edge of the worksheet, OFFSET returns the #REF! error value.
		if( (rc[0] < 0) || (rc[1] < 0) )
		{
			return new PtgErr( PtgErr.ERROR_REF );
		}
		PtgRef pr = new PtgRef();
		pr.setParentRec( ref.getParentRec() );
		pr.setLocation( rc );
		return pr;
	}

	/**
	 * TRANSPOSE
	 * Returns the transpose of an array
	 * The TRANSPOSE function returns a vertical range of cells as a horizontal range, or vice versa.
	 * The TRANSPOSE function must be entered as an array formula has columns and rows.
	 * Use TRANSPOSE to shift the vertical and horizontal orientation of an array or range on a
	 * worksheet.
	 * <p/>
	 * array  Required. An array or range of cells on a worksheet that you want to transpose. The transpose of an array is created by using the first row of the array as the first column of the new array, the second row of the array as the second column of the new array, and so on.
	 */
	protected static Ptg calcTranspose( Ptg[] operands )
	{
		String retArray = "";
		PtgArray ret = new PtgArray();
		if( !(operands[0] instanceof PtgArray) )
		{
			Ptg[] arr = operands[0].getComponents();
			//it's a list of values, convert to row-based
			for( Ptg anArr : arr )
			{
				retArray = retArray + anArr.getValue().toString() + ";";
			}
			retArray = "{" + retArray.substring( 0, retArray.length() - 1 ) + "}";
			ret.setVal( retArray );
		}
		else
		{    // transpose row/cols of an existing array
			PtgArray pa = (PtgArray) operands[0];
			Ptg[] arr = pa.getComponents();
			int nc = pa.getNumberOfColumns() + 1;
			int nr = pa.getNumberOfRows() + 1;
			for( int i = 0; i < nc; i++ )
			{
				for( int j = 0; j < (nc * nr); j += nc )
				{
					retArray = retArray + arr[i + j].getValue().toString() + ",";
				}
				retArray = retArray.substring( 0, retArray.length() - 1 ) + ";";
			}
			retArray = "{" + retArray.substring( 0, retArray.length() - 1 ) + "}";
			ret.setVal( retArray );
		}
		return ret;
	}

	/**
	 * ROW
	 * Returns the row number of a reference
	 * <p/>
	 * Note this is 1 based, ie Row1 = 1.
	 */
	protected static Ptg calcRow( Ptg[] operands ) throws FunctionNotSupportedException
	{
		if( (operands == null) || (operands.length != 1) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		if( operands[0] instanceof PtgFuncVar )
		{
			// we need to return the col where the formula is.
			PtgFuncVar pfunk = (PtgFuncVar) operands[0];
			try
			{
				int loc = pfunk.getParentRec().getRowNumber() + 1;
				return new PtgInt( loc );
			}
			catch( Exception e )
			{
				log.error( "Error running calcRow " + e );
			}
			;
		}
		try
		{        // process as an array formula ...
			boolean isArray = (operands[0].getParentRec() instanceof Array);
			if( !isArray )
			{
				if( operands[0] instanceof PtgRef )
				{
					return new PtgInt( (((PtgRef) operands[0]).getRowCol()[0]) + 1 );
				}
				if( operands[0] instanceof PtgName )
				{    // table???
					String range = ((PtgName) operands[0]).getName().getLocation();
					return new PtgInt( ExcelTools.getRowColFromString( range )[0] + 1 );
				}
				return new PtgInt( (operands[0].getIntLocation()[0]) + 1 );
			}
			String retArry = "";
			Ptg[] comps = null;
			if( operands[0] instanceof PtgRef )
			{
				comps = operands[0].getComponents();
			}
			else if( operands[0] instanceof PtgName )
			{    // table???
				comps = operands[0].getComponents();
			}
			if( comps == null )
			{
				return new PtgInt( (((PtgRef) operands[0]).getRowCol()[0]) + 1 );
			}
			for( Ptg comp : comps )
			{
				try
				{
					retArry = retArry + (((PtgRef) comp).getIntLocation()[0] + 1) + ",";
				}
				catch( Exception e )
				{
					;
				}
			}
			retArry = "{" + retArry.substring( 0, retArry.length() - 1 ) + "}";
			PtgArray pa = new PtgArray();
			pa.setVal( retArry );
			return pa;
		}
		catch( Exception ex )
		{
			return new PtgRefErr();
		}
	}

	/**
	 * ROWS
	 * Returns the number of rows in a reference
	 * <p/>
	 * ROWS(array)
	 * <p/>
	 * Array   is an array, an array formula, or a reference to a range of cells for which you want the number of rows.
	 * <p/>
	 * =ROWS(C1:E4) Number of rows in the reference (4)
	 * =ROWS({1,2,3;4,5,6}) Number of rows in the array constant (2)
	 */
	protected static Ptg calcRows( Ptg[] operands ) throws FunctionNotSupportedException
	{
		try
		{
			int rsz = 0;
			if( operands[0] instanceof PtgStr )
			{
				String rangestr = operands[0].getValue().toString();
				String startx = rangestr.substring( 0, rangestr.indexOf( ":" ) );
				String endx = rangestr.substring( rangestr.indexOf( ":" ) + 1 );

				int[] startints = ExcelTools.getRowColFromString( startx );
				int[] endints = ExcelTools.getRowColFromString( endx );
				rsz = endints[0] - startints[0];
				rsz++; // inclusive
			}
			else if( operands[0] instanceof PtgName )
			{
				int[] rc = ExcelTools.getRangeCoords( operands[0].getLocation() );
				rsz = rc[2] - rc[0];
				rsz++; // inclusive
			}
			else if( operands[0] instanceof PtgRef )
			{
				int[] rc = ExcelTools.getRangeCoords( ((PtgRef) operands[0]).getLocation() );
				rsz = rc[2] - rc[0];
				rsz++; // inclusive
			}
			else if( operands[0] instanceof PtgMemFunc )
			{
				Ptg[] p = operands[0].getComponents();
				if( (p != null) && (p.length > 0) )
				{
					int[] rc0 = p[0].getIntLocation();
					int[] rc1 = null;
					if( p.length > 1 )
					{
						rc1 = p[p.length - 1].getIntLocation();
					}
					if( rc1 == null )
					{
						rsz = 0;
					}
					else
					{
						rsz = rc1[0] - rc0[0];
					}
					rsz++;
				}
				else
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
			}
			else
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			return new PtgInt( rsz );
		}
		catch( Exception e )
		{
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * VLOOKUP
	 * Looks in the first column of an array and moves across the row to return the value of a cell
	 * Searches for a value in the leftmost column of a table, and then returns a value in the same row
	 * from a column you specify in the table. Use VLOOKUP instead of HLOOKUP when your comparison
	 * values are located in a column to the left of the data you want to find.
	 * <p/>
	 * Syntax
	 * <p/>
	 * VLOOKUP(lookup_value,table_array,col_index_num,range_lookup)
	 * <p/>
	 * Lookup_value   is the value to be found in the first column of the array. Lookup_value can be a value,
	 * a reference, or a text string.
	 * <p/>
	 * Table_array   is the table of information in which data is looked up. Use a reference to a range or
	 * a range name, such as Database or List.
	 * <p/>
	 * If range_lookup is TRUE, the values in the first column of table_array must be placed in ascending order:
	 * ..., -2, -1, 0, 1, 2, ..., A-Z, FALSE, TRUE; otherwise VLOOKUP may not give the correct value.
	 * If range_lookup is FALSE, table_array does not need to be sorted.
	 * You can put the values in ascending order by choosing the Sort command from the Data menu and selecting Ascending.
	 * The values in the first column of table_array can be text, numbers, or logical values.
	 * Uppercase and lowercase text are equivalent.
	 * Col_index_num   is the column number in table_array from which the matching value must be returned.
	 * A col_index_num of 1 returns the value in the first column in table_array; a col_index_num of 2 returns the value
	 * in the second column in table_array, and so on. If col_index_num is less than 1,
	 * VLOOKUP returns the #VALUE! error value; if col_index_num is greater than the number of columns in table_array,
	 * VLOOKUP returns the #REF! error value.
	 * <p/>
	 * Range_lookup   is a logical value that specifies whether you want VLOOKUP to find an exact match or an
	 * approximate match. If TRUE or omitted, an approximate match is returned. In other words,
	 * if an exact match is not found, the next largest value that is less than lookup_value is returned.
	 * If FALSE, VLOOKUP will find an exact match. If one is not found, the error value #N/A is returned.
	 * <p/>
	 * Remarks
	 * <p/>
	 * If VLOOKUP can't find lookup_value, and range_lookup is TRUE, it uses the largest value that is less than or
	 * equal to lookup_value.
	 * If lookup_value is smaller than the smallest value in the first column of table_array, VLOOKUP returns the
	 * #N/A error value.
	 * If VLOOKUP can't find lookup_value, and range_lookup is FALSE, VLOOKUP returns the #N/A value.
	 * <p/>
	 * VLOOKUP(lookup_value,table_array,col_index_num,range_lookup)
	 * <p/>
	 * On the preceding worksheet, where the range A4:C12 is named Range:
	 * <p/>
	 * VLOOKUP(1,Range,1,TRUE) equals 0.946
	 * <p/>
	 * VLOOKUP(1,Range,2) equals 2.17
	 * <p/>
	 * VLOOKUP(1,Range,3,TRUE) equals 100
	 * <p/>
	 * VLOOKUP(.746,Range,3,FALSE) equals 200
	 * <p/>
	 * VLOOKUP(0.1,Range,2,TRUE) equals #N/A, because 0.1 is less than the smallest value in column A
	 * <p/>
	 * VLOOKUP(2,Range,2,TRUE) equals 1.71
	 */
	protected static Ptg calcVlookup( Ptg[] operands ) throws FunctionNotSupportedException
	{

		boolean rangeLookup = true;    // truth of "approximate match"; must be sorted
		boolean isNumber = true;
		try
		{
			Ptg lookup_value = operands[0];
			Ptg table_array = operands[1];

			// can't assume that it's a PtgInt			
			//PtgInt col_index_num 	= (PtgInt)operands[2].getValue();
			//int colNum = col_index_num.getVal() -1;// reduce 1 for ordinal base off firstcol
			Object o = operands[2].getValue();
			int colNum = 0;
			if( o instanceof Double )
			{
				colNum = ((Double) o).intValue() - 1; // reduce 1 for ordinal base off firstcol
			}
			else    // assume int?
			{
				colNum = (Integer) o - 1; // reduce 1 for ordinal base off firstcol
			}
			if( operands.length > 3 )
			{
				Object vx = operands[3].getValue();
				if( vx != null )
				{
					try
					{
						Boolean sort = (Boolean) vx;
						rangeLookup = sort;
					}
					catch( ClassCastException e )
					{
						Integer bool = (Integer) vx;
						if( bool == 0 )
						{
							rangeLookup = false;
						}
					}
				}
			}

			PtgRef[] lookupComponents = null;
			PtgRef[] valueComponents = null;
			// first, get the lookup Column Vals
			if( table_array instanceof PtgName )
			{    // 20090211 KSC:
				PtgArea3d pa = new PtgArea3d( false );
				pa.setParentRec( table_array.getParentRec() );

				pa.setLocation( ((PtgName) table_array).getName().getLocation() );
				table_array = pa;
				if( ((PtgArea3d) table_array).isExternalRef() )
				{
					log.warn( "LookupReferenceCalculator.calcVlookup External References are disallowed" );
					return new PtgErr( PtgErr.ERROR_REF );
				}
			}
			if( table_array instanceof PtgArea )
			{
				try
				{
					PtgArea pa = (PtgArea) table_array;
					int[] range = table_array.getIntLocation();
					//			 TODO: check rc sanity here
					int firstcol = range[1];
					lookupComponents = (PtgRef[]) pa.getColComponents( firstcol );
					valueComponents = (PtgRef[]) pa.getColComponents( firstcol + colNum );
				}
				catch(/*20070209 KSC: FormulaNotFound*/Exception e )
				{
					log.warn( "LookupReferenceCalculator.calcVlookup cannot determine PtgArea location. " + e, e );
				}

			}
			else if( table_array instanceof PtgMemFunc )
			{ //  || table_array instanceof PtgMemArea){
				try
				{
					PtgMemFunc pa = (PtgMemFunc) table_array;
					// int[] range = table_array.getIntLocation();
					int firstcol = -1;

					try
					{
						int[] rc1 = pa.getFirstloc().getIntLocation();

						//   				 TODO: check rc sanity here
						firstcol = rc1[1];
					}
					catch( Exception e )
					{
						log.warn( "LookupReferenceCalculator.calcVlookup could not determine row col from PtgMemFunc.", e );
					}

					lookupComponents = (PtgRef[]) pa.getColComponents( firstcol );
					valueComponents = (PtgRef[]) pa.getColComponents( firstcol + colNum );
				}
				catch(/*20070209 KSC: FormulaNotFound*/Exception e )
				{
					log.warn( "LookupReferenceCalculator.calcVlookup cannot determine PtgArea location. " + e, e );
				}
			}
			// error check
			if( (lookupComponents == null) || (lookupComponents.length == 0) )
			{
				return new PtgErr( PtgErr.ERROR_REF );
			}
			if( (lookup_value == null) || (lookup_value.getValue() == null) ) // 20070221 KSC: Error trap getValue
			{
				return new PtgErr( PtgErr.ERROR_NULL );
			}
			// lets check if we are dealing with strings or numbers....
			try
			{
				String val = lookup_value.getValue().toString();
				if( val.length() == 0 )                    // 20090205 KSC
				{
					return new PtgErr( PtgErr.ERROR_NA );
				}
				Double d = new Double( val );
			}
			catch( NumberFormatException e )
			{
				isNumber = false;
			}

			// TODO:
			//if the value you supply for the lookup_value argument is smaller than the smallest value in the first column of the table_array argument, VLOOKUP returns the #N/A error value.

			if( isNumber )
			{
				double match_num;
				try
				{
					match_num = Double.parseDouble( lookup_value.getValue().toString() );
				}
				catch( NumberFormatException e )
				{
					return new PtgErr( PtgErr.ERROR_NA );
				}

				for( int i = 0; i < lookupComponents.length; i++ )
				{
					double val;

					try
					{
						val = Double.parseDouble( lookupComponents[i].getValue().toString() );
						if( val == 0 )
						{    // VLOOKUP does NOT treat blanks as 0's
							if( lookupComponents[i].refCell[0] == null )
							{
								continue;
							}
						}
					}
					catch( NumberFormatException e )
					{
						// Ignore entries in the table that aren't numbers.
						continue;
					}

					if( val == match_num )
					{
						return valueComponents[i].getPtgVal();
					}
					if( rangeLookup && (val > match_num) )
					{
						if( i == 0 )
						{
							return new PtgErr( PtgErr.ERROR_NA );
						}
						return valueComponents[i - 1].getPtgVal();
					}
				}

				if( rangeLookup )
				{
					return valueComponents[lookupComponents.length - 1].getPtgVal();
				}
				return new PtgErr( PtgErr.ERROR_NA );
			}

			// It's a String

			if( rangeLookup )
			{    // approximate match
				String match_str = lookup_value.getValue().toString();
				int match_len = match_str.length();
				for( int i = 0; i < lookupComponents.length; i++ )
				{
					try
					{
						String val = lookupComponents[i].getValue().toString();
						if( val.equalsIgnoreCase( match_str ) )
						{// we found it
							return valueComponents[i].getPtgVal();
						}
						if( (val.length() >= match_len) && val.substring( 0, match_len ).equalsIgnoreCase( match_str ) )
						{ // matches up to length, but not all, return previous
							return valueComponents[i - 1].getPtgVal();
						}
						if( ExcelTools.getIntVal( val.substring( 0, 1 ) ) > ExcelTools.getIntVal( match_str.substring( 0, 1 ) ) )
						{
							return valueComponents[i - 1].getPtgVal();
						}
						if( i == (lookupComponents.length - 1) )
						{// we reached the last one so use this
							return valueComponents[i].getPtgVal();
						}
					}
					catch( Exception e )
					{
					} // 20070209 KSC: ignore errors in lookup cells
				}
			}
			else
			{ // unsorted
				String match_str = lookup_value.getValue().toString();
				for( int i = 0; i < lookupComponents.length; i++ )
				{
					try
					{
						String val = lookupComponents[i].getValue().toString();
						try
						{
							if( val.equalsIgnoreCase( match_str ) )
							{// we found it
								return valueComponents[i].getPtgVal();
							}
							if( i == (lookupComponents.length - 1) )
							{// we reached the last one so error out
								return new PtgErr( PtgErr.ERROR_NA );
							}
						}
						catch( Exception e )
						{
							log.error( "LookupReferenceCalculator.calcVLookup error: " + e.toString() );
							return new PtgErr( PtgErr.ERROR_NA );
						}
					}
					catch( Exception e )
					{
					} // 20070209 KSC: ignore errors in lookup cells
				}
			}
		}
		catch( Exception e )
		{    // appears that an error with operands results in a #NA error
			return new PtgErr( PtgErr.ERROR_NA );

		}
		return new PtgErr( PtgErr.ERROR_NULL );
	}

}
