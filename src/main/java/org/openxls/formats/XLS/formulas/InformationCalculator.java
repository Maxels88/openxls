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

import org.openxls.formats.XLS.BiffRec;
import org.openxls.formats.XLS.CellNotFoundException;
import org.openxls.formats.XLS.FormatConstants;
import org.openxls.formats.XLS.FunctionNotSupportedException;
import org.openxls.formats.XLS.Labelsst;
import org.openxls.formats.XLS.WorkSheetNotFoundException;
import org.openxls.formats.XLS.XLSRecord;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.text.DecimalFormat;



/*
    InformationCalculator is a collection of static methods that operate
    as the Microsoft Excel function calls do.

    All methods are called with an array of ptg's, which are then
    acted upon.  A Ptg of the type that makes sense (ie boolean, number)
    is returned after each method.
*/

public class InformationCalculator
{
	private static final Logger log = LoggerFactory.getLogger( InformationCalculator.class );
	/**
	 * CELL
	 * Returns information about the formatting, location, or contents of a cell
	 * The CELL function returns information about the formatting, location, or contents of a cell.
	 * <p/>
	 * CELL(info_type, [reference])
	 * info_type  Required. A text value that specifies what type of cell information you want to return. The following list shows the possible values of the info_type argument and the corresponding results.info_type Returns
	 * "address" Reference of the first cell in reference, as text.
	 * "col" Column number of the cell in reference.
	 * "color" The value 1 if the cell is formatted in color for negative values; otherwise returns 0 (zero).
	 * "contents" Value of the upper-left cell in reference; not a formula.
	 * "filename" Filename (including full path) of the file that contains reference, as text. Returns empty text ("") if the worksheet that contains reference has not yet been saved.
	 * "format" Text value corresponding to the number format of the cell. The text values for the various formats are shown in the following table. Returns "-" at the end of the text value if the cell is formatted in color for negative values. Returns "()" at the end of the text value if the cell is formatted with parentheses for positive or all values.
	 * "parentheses" The value 1 if the cell is formatted with parentheses for positive or all values; otherwise returns 0.
	 * "prefix" Text value corresponding to the "label prefix" of the cell. Returns single quotation mark (') if the cell contains left-aligned text, double quotation mark (") if the cell contains right-aligned text, caret (^) if the cell contains centered text, backslash (\) if the cell contains fill-aligned text, and empty text ("") if the cell contains anything else.
	 * "protect" The value 0 if the cell is not locked; otherwise returns 1 if the cell is locked.
	 * "row" Row number of the cell in reference.
	 * "type" Text value corresponding to the type of data in the cell. Returns "b" for blank if the cell is empty, "l" for label if the cell contains a text constant, and "v" for value if the cell contains anything else.
	 * "width" Column width of the cell, rounded off to an integer. Each unit of column width is equal to the width of one character in the default font size.
	 * <p/>
	 * reference  Optional. The cell that you want information about. If omitted, the information specified in the info_type argument is returned for the last cell that was changed. If the reference argument is a range of cells, the CELL function returns the information for only the upper left cell of the range.
	 */
	protected static Ptg calcCell( Ptg[] operands ) throws FunctionNotSupportedException
	{
		String type = operands[0].getValue().toString().toLowerCase();
		PtgRef ref = null;
		BiffRec cell = null;
		if( operands.length > 1 )
		{
			ref = (PtgRef) operands[1];
			try
			{
				cell = ref.getParentRec().getWorkBook().getCell( ref.getLocationWithSheet() );
			}
			catch( CellNotFoundException e )
			{
				try
				{
					String sh = null;
					try
					{
						sh = ref.getSheetName();
					}
					catch( WorkSheetNotFoundException we )
					{
						;
					}
					if( sh == null )
					{
						sh = ref.getParentRec().getSheet().getSheetName();
					}
					cell = ref.getParentRec().getWorkBook().getWorkSheetByName( sh ).addValue( null, ref.getLocation() );
				}
				catch( Exception ex )
				{
					return new PtgErr( PtgErr.ERROR_VALUE );
				}
			}
			catch( Exception e )
			{
				return new PtgErr( PtgErr.ERROR_VALUE );
			}
			//  If ref param is omitted, the information specified in the info_type argument
			// is returned for the last cell that was changed
		}
		else if( !type.equals( "filename" ) )// no ref was passed in and option is not "filename"
		// We cannot determine which is the "last cell" they are referencing;
		{
			throw new FunctionNotSupportedException( "Worsheet function CELL with no reference parameter is not supported" );
		}
		else    // filename option can use any biffrec ...
		{
			cell = operands[0].getParentRec();
		}

		// at this point both ref (PtgRef) and r (BiffRec) should be valid
		try
		{
			if( type.equals( "address" ) )
			{
				PtgRef newref = ref;
				newref.clearLocationCache();
				newref.fColRel = false;        // make absolute
				newref.fRwRel = false;
				return new PtgStr( newref.getLocation() );
			}
			if( type.equals( "col" ) )
			{
				return new PtgNumber( ref.getIntLocation()[1] + 1 );
			}
			if( type.equals( "color" ) )
			{    // The value 1 if the cell is formatted in color for negative values; otherwise returns 0 (zero).
				String s = cell.getFormatPattern();
				if( s.indexOf( ";[Red" ) > -1 )
				{
					return new PtgNumber( 1 );
				}
				return new PtgNumber( 0 );
			}
			if( type.equals( "contents" ) )
			{// Value of the upper-left cell in reference; not a formula.
				return new PtgStr( cell.getStringVal() );
			}
			if( type.equals( "filename" ) )
			{
				String f = cell.getWorkBook().getFileName();
				String sh = cell.getSheet().getSheetName();
				int i = f.lastIndexOf( java.io.File.separatorChar );
				f = f.substring( 0, i + 1 ) + "[" + f.substring( i + 1 );
				f += "]" + sh;
				return new PtgStr( f );
			}
			if( type.equals( "format" ) )
			{    // Text value corresponding to the number format of the cell. The text values for the various formats are shown in the following table. Returns "-" at the end of the text value if the cell is formatted in color for negative values. Returns "()" at the end of the text value if the cell is formatted with parentheses for positive or all values.
				String s = cell.getFormatPattern();
				String ret = "G";    // default?
				if( s.equals( "General" ) ||
						s.equals( "# ?/?" ) ||
						s.equals( "# ??/??" ) )
				{
					ret = "G";
				}
				else if( s.equals( "0" ) )
				{
					ret = "F0";
				}
				else if( s.equals( "#,##0" ) )
				{
					ret = ",0";
				}
				else if( s.equals( "0.00" ) )
				{
					ret = "F2";
				}
				else if( s.equals( "#,##0.00" ) )
				{
					ret = ", 2";
				}
				else if( s.equals( "$#,##0_);($#,##0)" ) )
				{
					ret = "C0";
				}
				else if( s.equals( "$#,##0_);[Red]($#,##0)" ) )
				{
					ret = "C0-";
				}
				else if( s.equals( "$#,##0.00_);($#,##0.00)" ) )
				{
					ret = "C2";
				}
				else if( s.equals( "$#,##0.00_);[Red]($#,##0.00)" ) )
				{
					ret = "C2-";
				}
				else if( s.equals( "0%" ) )
				{
					ret = "P0";
				}
				else if( s.equals( "0.00%" ) )
				{
					ret = "P2";
				}
				else if( s.equals( "0.00E+00" ) )
				{
					ret = "S2";
//					   m/d/yy or m/d/yy h:mm or mm/dd/yy 	"D4"
				}
				else if( s.equals( "m/d/yy" ) ||
						s.equals( "m/d/yy h:mm" ) ||
						s.equals( "mm/dd/yy" ) ||
						s.equals( "mm-dd-yy" ) )
				{        // added last to accomodate Excel's regional short date setting (format #14)
					ret = "D4";
				}
				else if( s.equals( "d-mmm-yy" ) || s.equals( "dd-mmm-yy" ) )
				{
					ret = "D1";
				}
				else if( s.equals( "d-mmm" ) || s.equals( "dd-mmm" ) )
				{
					ret = "D2";
				}
				else if( s.equals( "mmm-yy" ) )
				{
					ret = "D3";
				}
				else if( s.equals( "mm/dd" ) )
				{
					ret = "D5";
				}
				else if( s.equals( "h:mm AM/PM" ) )
				{
					ret = "D7";
				}
				else if( s.equals( "h:mm:ss AM/PM" ) )
				{
					ret = "D6";
				}
				else if( s.equals( "h:mm" ) )
				{
					ret = "D9";
				}
				else if( s.equals( "h:mm:ss" ) )
				{
					ret = "D8";
				}
				return new PtgStr( ret );
			}
			if( type.equals( "parentheses" ) )
			{
				String s = cell.getFormatPattern();
				if( s.startsWith( "(" ) )
				{
					return new PtgNumber( 1 );
				}
				return new PtgNumber( 0 );
			}
			if( type.equals( "prefix" ) )
			{
				// TODO: THIS IS NOT CORRECT - EITHER INFORM USER OR ??
// DOESN'T APPEAR TO MATCH EXCEL
				//Text value corresponding to the "label prefix" of the cell.
				// Returns single quotation mark (') if the cell contains left-aligned text, double quotation mark (") if the cell contains right-aligned text,
				// caret (^) if the cell contains centered text, backslash (\) if the cell contains fill-aligned text, and empty text ("") if the cell contains anything else.
				int al = cell.getXfRec().getHorizontalAlignment();
				if( al == FormatConstants.ALIGN_LEFT )
				{
					return new PtgStr( "'" );
				}
				if( al == FormatConstants.ALIGN_CENTER )
				{
					return new PtgStr( "^" );
				}
				if( al == FormatConstants.ALIGN_RIGHT )
				{
					return new PtgStr( "\"" );
				}
				if( al == FormatConstants.ALIGN_FILL )
				{
					return new PtgStr( "\\" );
				}
				return new PtgStr( "" );
			}
			if( type.equals( "protect" ) )
			{
				if( cell.getXfRec().isLocked() )
				{
					return new PtgNumber( 1 );
				}
				return new PtgNumber( 0 );
			}
			if( type.equals( "row" ) )
			{
				return new PtgNumber( ref.getIntLocation()[0] + 1 );
			}
			if( type.equals( "type" ) )
			{
				//Text value corresponding to the type of data in the cell.
				// Returns "b" for blank if the cell is empty,
				//"l" for label if the cell contains a text constant, and
				// "v" for value if the cell contains anything else.
				if( ((XLSRecord) cell).isBlank )
				{
					return new PtgStr( "b" );
				}
				if( cell instanceof Labelsst )
				{
					return new PtgStr( "l" );
				}
				return new PtgStr( "v" );
			}
			if( type.equals( "width" ) )
			{
				int n = 0;
				n = cell.getSheet().getColInfo( cell.getColNumber() ).getColWidthInChars();
				return new PtgNumber( n );
			}
		}
		catch( Exception e )
		{
			log.warn( "CELL: unable to calculate: " + e.toString(), e );
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * ERROR.TYPE
	 * Returns a number corresponding to an error type
	 * If error_val is
	 * ERROR.TYPE returns
	 * <p/>
	 * #NULL!	 1
	 * #DIV/0! 2
	 * #VALUE! 3
	 * #REF!	 4
	 * #NAME?	 5
	 * #NUM!	 6
	 * #N/A	 7
	 * Anything else	 #N/A
	 */
	protected static Ptg calcErrorType( Ptg[] operands )
	{
		Object o = operands[0].getValue();
		String s = o.toString();
		if( s.equalsIgnoreCase( "#NULL!" ) )
		{
			return new PtgInt( 1 );
		}
		if( s.equalsIgnoreCase( "#DIV/0!" ) )
		{
			return new PtgInt( 2 );
		}
		if( s.equalsIgnoreCase( "#VALUE!" ) )
		{
			return new PtgInt( 3 );
		}
		if( s.equalsIgnoreCase( "#REF!" ) )
		{
			return new PtgInt( 4 );
		}
		if( s.equalsIgnoreCase( "#NAME?" ) )
		{
			return new PtgInt( 5 );
		}
		if( s.equalsIgnoreCase( "#NUM!" ) )
		{
			return new PtgInt( 6 );
		}
		if( s.equalsIgnoreCase( "#N/A" ) )
		{
			return new PtgInt( 7 );
		}
		return new PtgErr( PtgErr.ERROR_NA );
	}

	/**
	 * INFO
	 * Returns information about the current operating environment
	 * INFO(type_text)
	 * <p/>
	 * NOTE: Several options are incomplete:
	 * "osversion"	-- only valid for Windows versions
	 * "system" 	-- only valid for Windows and Mac
	 * "release"	-- incomplete
	 * "origin"	-- does not return R1C1 format
	 */
	protected static Ptg calcInfo( Ptg[] operands )
	{
		// validate
		if( (operands == null) || (operands.length == 0) || (operands[0].getParentRec() == null) )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		String type_text = operands[0].getString();
		String ret = "";
		if( type_text.equals( "directory" ) )        // Path of the current directory or folder
		{
			return new PtgStr( System.getProperty( "user.dir" ).toLowerCase() + "\\" );
		}
		if( type_text.equals( "numfile" ) )    // number of active worksheets in the current workbook
// TODO: what is correct definition of "Active Worksheets"  - hidden state doesn't seem to affect"
		{
			return new PtgNumber( operands[0].getParentRec().getWorkBook().getNumWorkSheets() );
		}
		if( type_text.equals( "origin" ) )
		{		/* Returns the absolute cell reference of the top and leftmost
	   											    cell visible in the window, based on the current scrolling
	   											    position, as text prepended with "$A:".
	   											    This value is intended for for Lotus 1-2-3 release 3.x compatibility.
	   											    The actual value returned depends on the current reference
	   											    style setting. Using D9 as an example, the return value would be:
	   											    A1 reference style   "$A:$D$9".
													R1C1 reference style  "$A:R9C4"
	    										*/
// TODO: FINISH R1C1 reference style
			String cell = operands[0].getParentRec().getSheet().getWindow2().getTopLeftCell();
			for( int i = cell.length() - 1; i >= 0; i-- )
			{
				if( !Character.isDigit( cell.charAt( i ) ) )
				{
					cell = cell.substring( 0, i + 1 ) + "$" + cell.substring( i + 1 );
					break;
				}
			}
			cell = "$A:$" + cell;
			return new PtgStr( cell );
		}
		if( type_text.equals( "osversion" ) )
		{    //Current operating system version, as text.
			// see end of file for os info
			String osversion = System.getProperty( "os.version" );
			String n = System.getProperty( "os.name" );    // Windows Vista
			String os = "";
			// TODO:  need a list of osversions to compare to!  have know idea for mac, linux ...
			if( n.startsWith( "Windows" ) )
			{
				double v = new Double( osversion );
				os = "Windows (32-bit) ";
				if( v >= 5 )
				{
					os += "NT ";
				}
				DecimalFormat df = new DecimalFormat( "##.00" );
				os += df.format( v );
			} // otherwise have NO idea as cannot find any info on net
			else
			{
				os += osversion;
			}
			return new PtgStr( os );
		}
		if( type_text.equals( "recalc" ) )
		{    //Current recalculation mode; returns "Automatic" or "Manual".
			if( operands[0].getParentRec().getWorkBook().getRecalculationMode() == 0 ) // manual
			{
				return new PtgStr( "Manual" );
			}
			return new PtgStr( "Automatic" );
		}
		if( type_text.equals( "release" ) )
		{    //Version of Microsoft Excel, as text.
			// TODO: Finish!  97= 8.0, 2000= 9.0, 2002 (XP)= 10.0, 2003= 11.0, 2007= 12.0
			log.warn( "Worksheet Function INFO(\"release\") is not supported" );
			return new PtgStr( "" );
		}
		if( type_text.equals( "system" ) )        // Name of the operating environment: Macintosh = "mac" Windows = "pcdos"
		// TODO: linux?  ****************
		{
			if( System.getProperty( "os.name" ).indexOf( "Windows" ) >= 0 )
			{
				return new PtgStr( "pcdos" );
			}
			return new PtgStr( "mac" );
		}
		// In previous versions of Microsoft Office Excel, the "memavail", "memused", and "totmem" type_text values, returned memory information.
		// These type_text values are no longer supported and now return a #N/A error value.
		if( type_text.equals( "memavail" ) ||
				type_text.equals( "memused" ) ||
				type_text.equals( "totmem" ) )
		{
			return new PtgErr( PtgErr.ERROR_NA );
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * ISBLANK
	 * ISBLANK determines if the cell referenced is blank, and returns
	 * a boolean ptg based off that
	 */
	protected static Ptg calcIsBlank( Ptg[] operands )
	{
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		for( Ptg allop : allops )
		{
			// 20081120 KSC: blanks are handled differently now as Excel counts blank cells as 0's
			/*Object o = allops[i].getValue();
			if (o != null) return new PtgBool(false);
			*/
			if( !allop.isBlank() )
			{
				return new PtgBool( false );
			}
		}
		return new PtgBool( true );
	}

	/**
	 * ISERROR
	 * Value refers to any error value
	 * (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, or #NULL!).
	 * Usage@ ISERROR(value)
	 * Return@ PtgBool
	 */
	protected static Ptg calcIserror( Ptg[] operands )
	{
		if( operands[0] instanceof PtgErr )
		{
			return new PtgBool( true );
		}
		String[] errorstr = { "#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!" };
		Object o = operands[0].getValue();
		String opval = o.toString();
		for( String anErrorstr : errorstr )
		{
			if( opval.equalsIgnoreCase( anErrorstr ) )
			{
				return new PtgBool( true );
			}
		}
		return new PtgBool( false );
	}

	/**
	 * ISERR
	 * Returns TRUE if the value is any error value except #N/A
	 */
	protected static Ptg calcIserr( Ptg[] operands )
	{
		String[] errorstr = { "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!" };
		if( operands.length != 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Object o = operands[0].getValue();
		String opval = o.toString();
		for( String anErrorstr : errorstr )
		{
			if( opval.equalsIgnoreCase( anErrorstr ) )
			{
				return new PtgBool( true );
			}
		}
		return new PtgBool( false );
	}

	/**
	 * ISEVEN(number)
	 * <p/>
	 * Number   is the value to test. If number is not an integer, it is truncated.
	 * <p/>
	 * Remarks
	 * If number is nonnumeric, ISEVEN returns the #VALUE! error value.
	 * Examples
	 * ISEVEN(-1) equals FALSE
	 * ISEVEN(2.5) equals TRUE
	 * ISEVEN(5) equals FALSE
	 * <p/>
	 * author: John
	 */
	protected static Ptg calcIsEven( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		if( allops.length > 1 )
		{
			return new PtgBool( false );
		}
		Object o = operands[0].getValue();
		if( o != null )
		{
			try
			{    // KSC: mod for different number types + mod typo
				if( o instanceof Integer )
				{
					int s = (Integer) o;
					if( s < 0 )
					{
						return new PtgBool( false );
					}
					return new PtgBool( ((s % 2) == 0) );
				}
				if( o instanceof Float )
				{
					float s = (Float) o;
					if( s < 0 )
					{
						return new PtgBool( false );
					}
					return new PtgBool( ((s % 2) == 0) );
				}
				if( o instanceof Double )
				{
					double s = (Double) o;
					if( s < 0 )
					{
						return new PtgBool( false );
					}
					return new PtgBool( ((s % 2) == 0) );
				}
			}
			catch( Exception e )
			{
			}
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * ISLOGICAL
	 * Returns TRUE if the value is a logical value
	 */
	protected static Ptg calcIsLogical( Ptg[] operands )
	{
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		if( allops.length > 1 )
		{
			return new PtgBool( false );
		}
		Object o = operands[0].getValue();
		// unfortunately we need to know the difference between 
		// "true" and true, if it's a reference this can be difficult
		try
		{
			Boolean b = (Boolean) o;
			return new PtgBool( true );
		}
		catch( ClassCastException e )
		{
		}
		;
		return new PtgBool( false );
	}

	/**
	 * ISNUMBER
	 * Returns TRUE if the value is a number
	 */
	protected static Ptg calcIsNumber( Ptg[] operands )
	{
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		if( allops.length > 1 )
		{
			return new PtgBool( false );
		}
		Object o = operands[0].getValue();
		try
		{
			Float f = (Float) o;
			return new PtgBool( true );
		}
		catch( ClassCastException e )
		{
			try
			{
				Double d = (Double) o;
				return new PtgBool( true );
			}
			catch( ClassCastException ee )
			{
				try
				{
					Integer ii = (Integer) o;
					return new PtgBool( true );
				}
				catch( ClassCastException eee )
				{
				}
				;
			}
			;
		}
		return new PtgBool( false );
	}

	/**
	 * ISNONTEXT
	 * Returns TRUE if the value is not text
	 */
	protected static Ptg calcIsNonText( Ptg[] operands )
	{
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		if( allops.length > 1 )
		{
			return new PtgBool( false );
		}
		// blanks return true for this test
		if( allops[0].isBlank() )
		{
			return new PtgBool( true );
		}
		Object o = operands[0].getValue();
		if( o != null )
		{
			try
			{
				String s = (String) o;
				return new PtgBool( false );
			}
			catch( ClassCastException e )
			{
			}
			;
		}
		return new PtgBool( true );
	}

	/**
	 * ISNA
	 * Value refers to the #N/A
	 * (value not available) error value.
	 * <p/>
	 * usage@ ISNA(value)
	 * return@ PtgBool
	 */
	protected static Ptg calcIsna( Ptg[] operands )
	{
		if( operands.length != 1 )
		{
			return PtgCalculator.getError();
		}
		if( operands[0] instanceof PtgErr )
		{
			PtgErr per = (PtgErr) operands[0];
			if( per.getErrorType() == PtgErr.ERROR_NA )
			{
				return new PtgBool( true );
			}
		}
		else if( operands[0].getIsReference() )
		{
			Object o = operands[0].getValue();
			if( o.toString().equalsIgnoreCase( new PtgErr( PtgErr.ERROR_NA ).toString() ) )
			{
				return new PtgBool( true );
			}
		}
		return new PtgBool( false );
	}

	/**
	 * NA
	 * Returns the error value #N/A
	 */
	protected static Ptg calcNa( Ptg[] operands )
	{
		return new PtgErr( PtgErr.ERROR_NA );
	}

	/**
	 * ISTEXT
	 * Returns TRUE if the value is text
	 */
	protected static Ptg calcIsText( Ptg[] operands )
	{
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		if( allops.length > 1 )
		{
			return new PtgBool( false );
		}
		Object o = operands[0].getValue();
		if( o != null )
		{
			try
			{
				String s = (String) o;
				return new PtgBool( true );
			}
			catch( ClassCastException e )
			{
			}
			;
		}
		return new PtgBool( false );
	}

	/**
	 * ISODD
	 * Returns TRUE if the number is odd
	 * <p/>
	 * author: John
	 */
	protected static Ptg calcIsOdd( Ptg[] operands )
	{
		if( operands.length < 1 )
		{
			return new PtgErr( PtgErr.ERROR_VALUE );
		}
		Ptg[] allops = PtgCalculator.getAllComponents( operands );
		if( allops.length > 1 )
		{
			return new PtgBool( false );
		}
		Object o = operands[0].getValue();
		if( o != null )
		{
			try
			{    // KSC: mod for different number types + mod typo
				if( o instanceof Integer )
				{
					int s = (Integer) o;
					if( s < 0 )
					{
						return new PtgBool( false );
					}
					return new PtgBool( ((s % 2) != 0) );
				}
				if( o instanceof Float )
				{
					float s = (Float) o;
					if( s < 0 )
					{
						return new PtgBool( false );
					}
					return new PtgBool( ((s % 2) != 0) );
				}
				if( o instanceof Double )
				{
					double s = (Double) o;
					if( s < 0 )
					{
						return new PtgBool( false );
					}
					return new PtgBool( ((s % 2) != 0) );
				}
			}
			catch( Exception e )
			{
			}
		}
		return new PtgErr( PtgErr.ERROR_VALUE );
	}

	/**
	 * ISREF
	 * Returns TRUE if the value is a reference
	 */
	protected static Ptg calcIsRef( Ptg[] operands )
	{
		if( operands[0].getIsReference() )
		{
			return new PtgBool( true );
		}
		return new PtgBool( false );
	}

	/**
	 * N
	 * Returns a value converted to a number.
	 * <p/>
	 * Syntax
	 * <p/>
	 * N(value)
	 * <p/>
	 * Value   is the value you want converted. N converts values listed in the following table.
	 * <p/>
	 * If value is or refers to
	 * N returns
	 * <p/>
	 * A number
	 * That number
	 * <p/>
	 * A date, in one of the built-in date formats available in Microsoft Excel
	 * The serial number of that date --- Note that to us, this is just a number, the date
	 * format is just that, a format.
	 * <p/>
	 * TRUE
	 * 1
	 * <p/>
	 * Anything else
	 * 0
	 */
	protected static Ptg calcN( Ptg[] operands )
	{
		Object o = operands[0].getValue();
		if( (o instanceof Double) || (o instanceof Integer) || (o instanceof Float) || (o instanceof Long) )
		{
			Double d = new Double( o.toString() );
			return new PtgNumber( d );
		}
		if( o instanceof Boolean )
		{
			Boolean b = (Boolean) o;
			boolean bo = b;
			if( bo )
			{
				return new PtgInt( 1 );
			}
		}
		return new PtgInt( 0 );
	}

	/**
	 * TYPE
	 * Returns a number indicating the data type of a value
	 * Value   can be any Microsoft Excel value, such as a number, text, logical value, and so on.
	 * If value is TYPE returns
	 * Number 1
	 * Text 2
	 * Logical value 4
	 * Error value 16
	 * Array 64
	 */
	protected static Ptg calcType( Ptg[] operands )
	{
		if( operands[0] instanceof PtgArray )
		{
			return new PtgNumber( 64 );    // avoid value calc for arrays
		}
		if( operands[0] instanceof PtgErr )
		{
			return new PtgNumber( 16 );
		}

		// otherwise, test value of operand
		Object value = operands[0].getValue();
		int type = 0;
		if( value instanceof String )
		{
			type = 2;
		}
		else if( value instanceof Number )
		{
			type = 1;
		}
		else if( value instanceof Boolean )
		{
			type = 4;
		}
		return new PtgNumber( type );
	}

}

/*
 * known INFO function operating systems: 
 * TODO: need complete list
Windows Vista   Windows (32-bit) NT 6.00
Windows XP 		Windows (32-bit) NT 5.01
Windows2000 	Windows (32-bit) NT 5.00
Windows98 		Windows (32-bit) 4.10
Windows95 		Windows (32-bit) 4.00
*/

/*
Linux 	2.0.31 	x86 	IBM Java 1.3
Linux 	(*) 	i386 	Sun Java 1.3.1, 1.4 or Blackdown Java; (*) os.version depends on Linux Kernel version
Linux 	(*) 	x86_64 	Blackdown Java; note x86_64 might change to amd64; (*) os.version depends on Linux Kernel version
Linux 	(*) 	sparc 	Blackdown Java; (*) os.version depends on Linux Kernel version
Linux 	(*) 	ppc 	Blackdown Java; (*) os.version depends on Linux Kernel version
Linux 	(*) 	armv41 	Blackdown Java; (*) os.version depends on Linux Kernel version
Linux 	(*) 	i686 	GNU Java Compiler (GCJ); (*) os.version depends on Linux Kernel version
Linux 	(*) 	ppc64 	IBM Java 1.3; (*) os.version depends on Linux Kernel version
Mac OS 	7.5.1 	PowerPC 	
Mac OS 	8.1 	PowerPC 	
Mac OS 	9.0, 9.2.2 	PowerPC 	MacOS 9.0: java.version=1.1.8, mrj.version=2.2.5; MacOS 9.2.2: java.version=1.1.8 mrj.version=2.2.5
Mac OS X 	10.1.3 	ppc 	
Mac OS X 	10.2.6 	ppc 	Java(TM) 2 Runtime Environment, Standard Edition (build 1.4.1_01-39)
Java HotSpot(TM) Client VM (build 1.4.1_01-14, mixed mode)
Mac OS X 	10.2.8 	ppc 	using 1.3 JVM: java.vm.version=1.3.1_03-74, mrj.version=3.3.2; using 1.4 JVM: java.vm.version=1.4.1_01-24, mrj.version=69.1
Mac OS X 	10.3.1, 10.3.2, 10.3.3, 10.3.4 	ppc 	JDK 1.4.x
Mac OS X 	10.3.8 	ppc 	Mac OS X 10.3.8 Server; using 1.3 JVM: java.vm.version=1.3.1_03-76, mrj.version=3.3.3; using 1.4 JVM: java.vm.version=1.4.2-38; mrj.version=141.3
Windows 95 	4.0 	x86 	
Windows 98 	4.10 	x86 	Note, that if you run Sun JDK 1.2.1 or 1.2.2 Windows 98 identifies itself as Windows 95.
Windows Me 	4.90 	x86 	
Windows NT 	4.0 	x86 	
Windows 2000 	5.0 	x86 	
Windows XP 	5.1 	x86 	Note, that if you run older Java runtimes Windows XP identifies itself as Windows 2000.
Windows 2003 	5.2 	x86 	java.vm.version=1.4.2_06-b03; Note, that Windows Server 2003 identifies itself only as Windows 2003.
Windows CE 	3.0 build 11171 	arm 	Compaq iPAQ 3950 (PocketPC 2002)
OS/2 	20.40 	x86 	
Solaris 	2.x 	sparc 	
SunOS 	5.7 	sparc 	Sun Ultra 5 running Solaris 2.7
SunOS 	5.8 	sparc 	Sun Ultra 2 running Solaris 8
SunOS 	5.9 	sparc 	Java(TM) 2 Runtime Environment, Standard Edition (build 1.4.0_01-b03)
Java HotSpot(TM) Client VM (build 1.4.0_01-b03, mixed mode)
MPE/iX 	C.55.00 	PA-RISC 	
HP-UX 	B.10.20 	PA-RISC 	JDK 1.1.x
HP-UX 	B.11.00 	PA-RISC 	JDK 1.1.x
HP-UX 	B.11.11 	PA-RISC 	JDK 1.1.x
HP-UX 	B.11.11 	PA_RISC 	JDK 1.2.x/1.3.x; note Java 2 returns PA_RISC and Java 1 returns PA-RISC
HP-UX 	B.11.00 	PA_RISC 	JDK 1.2.x/1.3.x
HP-UX 	B.11.23 	IA64N 	JDK 1.4.x
HP-UX 	B.11.11 	PA_RISC2.0 	JDK 1.3.x or JDK 1.4.x, when run on a PA-RISC 2.0 system
HP-UX 	B.11.11 	PA_RISC 	JDK 1.2.x, even when run on a PA-RISC 2.0 system
HP-UX 	B.11.11 	PA-RISC 	JDK 1.1.x, even when run on a PA-RISC 2.0 system
AIX 	5.2 	ppc64 	sun.arch.data.model=64
AIX 	4.3 	Power 	
AIX 	4.1 	POWER_RS 	
OS/390 	390 	02.10.00 	J2RE 1.3.1 IBM OS/390 Persistent Reusable VM
FreeBSD 	2.2.2-RELEASE 	x86 	
Irix 	6.3 	mips 	
Digital Unix 	4.0 	alpha 	
NetWare 4.11 	4.11 	x86 	
OSF1 	V5.1 	alpha 	Java 1.3.1 on Compaq (now HP) Tru64 Unix V5.1
OpenVMS 	V7.2-1 	alpha 	Java 1.3.1_1 on OpenVMS 7.2
*/