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

import com.extentech.ExtenXLS.DateConverter;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.ExtenXLS.ValidationHandle;
import com.extentech.formats.XLS.formulas.FormulaParser;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgArea;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;
import org.xmlpull.v1.XmlPullParser;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Stack;

/**
 * <b>Dv: Data Validity Settings (01BEh)</b><br>
 * <p/>
 * This record is part of the Data Validity Table. It stores data validity settings and a list of cell ranges which
 * contain these settings. The prompt box appears while editing such a cell. The error box appears, if the entered value
 * does not fit the conditions. The data validity settings of a sheet are stored in a sequential list of DV records. This list is
 * preluded by an DVAL record. If a string is empty and the default text should appear in the prompt box or error
 * box, the string must contain a single zero character (string length will be 1).
 * <p><pre>
 * <p/>
 * <p/>
 * Offset          Name              Size                Contents
 * --------------------------------------------------------
 * 0               dwDvFlags      4                   Option flags (see below)
 * 4               dTitlePrompt   var                 Title of the prompt box (Unicode string, 16-bit string length)
 * var.            dTitleError      var.                Title of the error box (Unicode string, 16-bit string length)
 * var.            dTextPrompt  var.                Text of the prompt box (Unicode string, 16-bit string length)
 * var.            dTextError      var.                Text of the error box (Unicode string, 16-bit string length)
 * var.            sz1                 2                    Size of the formula data for first condition (sz1)
 * var.            garbage          2                   Not used
 * var.            firstCond       sz1                Formula data for first condition (RPN token array without size field)
 * var.            sz2                 2                   Size of the formula data for second condition (sz2)
 * var.            garbage           2                   Not used
 * var.            secondCond sz2                  Formula data for second condition (RPN token array without size field)
 * var.            cRangeList    var.                 Cell range address list with all affected ranges
 * <p/>
 * Option flags field:
 * Bit             Mask                    Name                 Contents
 * --
 * 3-0         0000000FH           ValType             Data type: 00H = Any value
 * 01H = Integer values
 * 02H = Decimal values
 * 03H = User defined list
 * 04H = Date
 * 05H = Time
 * 06H = Text length
 * 07H = Formula
 * 6-4     00000070H               ErrStyle               Error style: 00H = Stop
 * 01H = Warning
 * 02H = Info
 * 7         00000080H              fStrLookup         1 = In list type validity the string list is explicitly given in the formula
 * 8         00000100H            fAllowBlank         1 = Empty cells allowed
 * 9         00000200H             fSuppressCombo 1 = Suppress the drop down arrow in list type validity
 * 18      00040000H             fShowInputMsg     1 = Show prompt box if cell selected
 * 19      00080000H             fShowErrorMsg      1 = Show error box if invalid values entered
 * 23-20 00F00000H             typOperator         Condition operator: 00H = Between
 * 01H = Not between
 * 02H = Equal
 * 03H = Not equal
 * 04H = Greater than
 * 05H = Less than
 * 06H = Greater or equal
 * 07H = Less or equal
 * <p/>
 * </pre></p>
 * In list type validity it is possible to enter an explicit string list. This string list is stored as tStr token . The string
 * items are separated by zero characters. There is no zero character at the end of the string list.
 * Example for a string list with the 3 strings A, B, and C: A<00H>B<00H>C (contained in a tStr token, string
 * length is 5).
 */
public class Dv extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7895028832113540094L;
	private int grbit;
	private Unicodestring dTitlePrompt;
	private Unicodestring dTitleError;
	private Unicodestring dTextPrompt;
	private Unicodestring dTextError;
	private Stack firstCond;
	private Stack secondCond;
	private ArrayList cRangeList;
	private byte[] garbageByteOne = new byte[2];
	private byte[] garbageByteTwo = new byte[2];
	byte numLocs;
	//private byte[] garbageByteThree = new byte[1];
	// grbit (dwDvFlags) fields
	private byte valType;
	private byte errStyle;
	private boolean fStrLookup;
	private boolean fAllowBlank;
	private boolean fSuppressCombo;
	private boolean fShowInputMsg;
	private boolean fShowErrMsg;
	private short IMEMode;
	private byte typOperator;
	private boolean dirtyflag = false;

	// 20090606: made prototype completely blank i.e. no prompt, error text or formulas ...
//    private byte[] PROTOTYPE_BYTES = {0,   1,  12,   0,   0,   0,   0,  0,   0,   0,  0,   0,   0,  0,   0,   0,  0,   0,   0,   0,		  
//      0,   0,  89,  84,  0 };//,   0,   0,   0,   0,   0,  0,   0,   0,   0};
	private byte[] PROTOTYPE_BYTES = {
			3, 1, 12, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 1, 0
	};

	// BITMASK
	private static final int BITMASK_VALTYPE = 0x0000000F;
	private static final int BITMASK_ERRSTYLE = 0x00000070;
	private static final int BITMASK_FSTRLOOKUP = 0x00000080;
	private static final int BITMASK_FALLOWBLANK = 0x00000100;
	private static final int BITMASK_FSUPRESSCOMBO = 0x00000200;
	private static final int BITMASK_MDIMEMODE = 0x0003FC00;
	private static final int BITMASK_FSHOWINPUTMSG = 0x00040000;
	private static final int BITMASK_FSHOWERRORMSG = 0x00080000;
	private static final int BITMASK_TYPOPERATOR = 0x00F00000;

	// 20090609 KSC: need to store ranges separately as OOXML ranges addresses may exceed 2003 maximum size
	private String[] ooxmlranges = null;

	/**
	 * Determine if the value passed in is valid for
	 * this validation
	 *
	 * @param value
	 * @return
	 */
	public boolean isValid( Object value ) throws ValidationException
	{

		// TODO: look into whether null is ever a valid value.  for now we assume "no".
		if( value == null )
		{
			throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );
		}

		if( !this.isCorrectDataType( value ) )
		{
			throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );
		}
		if( value instanceof Date )
		{
			double d = DateConverter.getXLSDateVal( (java.util.Date) value );
			value = d + "";
		}
		switch( typOperator )
		{
			case ValidationHandle.CONDITION_BETWEEN:
				if( isBetween( value ) )
				{
					return true;
				}
				throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );
			case ValidationHandle.CONDITION_NOT_BETWEEN:
				if( isNotBetween( value ) )
				{
					return true;
				}
				throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );
			case ValidationHandle.CONDITION_EQUAL:
				if( isEqual( value ) )
				{
					return true;
				}
				throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );
			case ValidationHandle.CONDITION_GREATER_THAN:
				if( isGreaterThan( value ) )
				{
					return true;
				}
				throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );
			case ValidationHandle.CONDITION_GREATER_OR_EQUAL:
				if( isGreaterOrEqual( value ) )
				{
					return true;
				}
				throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );
			case ValidationHandle.CONDITION_LESS_OR_EQUAL:
				if( isLessOrEqual( value ) )
				{
					return true;
				}
				throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );

			case ValidationHandle.CONDITION_LESS_THAN:
				if( isGreaterOrEqual( value ) )
				{
					return true;
				}
				throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );

			case ValidationHandle.CONDITION_NOT_EQUAL:
				if( isNotEqual( value ) )
				{
					return true;
				}
				throw new ValidationException( this.getErrorBoxTitle(), this.getErrorBoxText() );

		}
		return true;
	}

	/**
	 * Validate that the passed in value is between the two values specified in the parameters
	 *
	 * @param value
	 * @return
	 */
	private boolean isBetween( Object value )
	{
		String s1 = StringTool.strip( FormulaParser.getExpressionString( firstCond ), "=" );
		String s2 = StringTool.strip( FormulaParser.getExpressionString( secondCond ), "=" );
		String formulaStr = "=and(" + value.toString() + ">" + s1 + "," + s2 + ">" + value.toString() + ")";
		try
		{
			Formula f = FormulaParser.getFormulaFromString( formulaStr, this.getWorkBook().getWorkSheetByNumber( 0 ), new int[]{ 1, 1 } );
			f.setCachedValue( null );
			Object o = f.calculateFormula();
			if( o instanceof Boolean )
			{
				return (Boolean) o;
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error calculating formula in validation " + e.toString() );
		}
		return false;
	}

	/**
	 * Validate that the passed in value is NOT between the two passed in values
	 *
	 * @param value
	 * @return
	 */
	private boolean isNotBetween( Object value )
	{
		return !isBetween( value );
	}

	/**
	 * Validate that the passed in value is equivalant
	 *
	 * @param value
	 * @return
	 */
	private boolean isEqual( Object value )
	{
		String s1 = StringTool.strip( FormulaParser.getExpressionString( firstCond ), "=" );
		String formulaStr = "=(" + value.toString() + "=" + s1 + ")";
		try
		{
			Formula f = FormulaParser.getFormulaFromString( formulaStr, this.getWorkBook().getWorkSheetByNumber( 0 ), new int[]{ 1, 1 } );
			Object o = f.calculateFormula();
			if( o instanceof Boolean )
			{
				return (Boolean) o;
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error calculating formula in validation " + e.toString() );
		}
		return false;
	}

	/**
	 * Validate that the passed in value is NOT between the two passed in values
	 *
	 * @param value
	 * @return
	 */
	private boolean isNotEqual( Object value )
	{
		return !isEqual( value );
	}

	/**
	 * Validate that the passed in value is greater than
	 *
	 * @param value
	 * @return
	 */
	private boolean isGreaterThan( Object value )
	{
		String s1 = StringTool.strip( FormulaParser.getExpressionString( firstCond ), "=" );
		String formulaStr = "=(" + value.toString() + ">" + s1 + ")";
		try
		{
			Formula f = FormulaParser.getFormulaFromString( formulaStr, this.getWorkBook().getWorkSheetByNumber( 0 ), new int[]{ 1, 1 } );
			Object o = f.calculateFormula();
			if( o instanceof Boolean )
			{
				return (Boolean) o;
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error calculating formula in validation " + e.toString() );
		}
		return false;
	}

	/**
	 * Validate that the passed in value is greater than
	 *
	 * @param value
	 * @return
	 */
	private boolean isGreaterOrEqual( Object value )
	{
		String s1 = StringTool.strip( FormulaParser.getExpressionString( firstCond ), "=" );
		String formulaStr = "=(" + value.toString() + ">=" + s1 + ")";
		try
		{
			Formula f = FormulaParser.getFormulaFromString( formulaStr, this.getWorkBook().getWorkSheetByNumber( 0 ), new int[]{ 1, 1 } );
			Object o = f.calculateFormula();
			if( o instanceof Boolean )
			{
				return (Boolean) o;
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error calculating formula in validation " + e.toString() );
		}
		return false;
	}

	/**
	 * Validate that the passed in value is greater than
	 *
	 * @param value
	 * @return
	 */
	private boolean isLessThan( Object value )
	{
		String s1 = StringTool.strip( FormulaParser.getExpressionString( firstCond ), "=" );
		String formulaStr = "=(" + value.toString() + "<" + s1 + ")";
		try
		{
			Formula f = FormulaParser.getFormulaFromString( formulaStr, this.getWorkBook().getWorkSheetByNumber( 0 ), new int[]{ 1, 1 } );
			Object o = f.calculateFormula();
			if( o instanceof Boolean )
			{
				return (Boolean) o;
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error calculating formula in validation " + e.toString() );
		}
		return false;
	}

	/**
	 * Validate that the passed in value is greater than
	 *
	 * @param value
	 * @return
	 */
	private boolean isLessOrEqual( Object value )
	{
		String s1 = StringTool.strip( FormulaParser.getExpressionString( firstCond ), "=" );
		String formulaStr = "=(" + value.toString() + "<=" + s1 + ")";
		try
		{
			Formula f = FormulaParser.getFormulaFromString( formulaStr, this.getWorkBook().getWorkSheetByNumber( 0 ), new int[]{ 1, 1 } );
			Object o = f.calculateFormula();
			if( o instanceof Boolean )
			{
				return (Boolean) o;
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error calculating formula in validation " + e.toString() );
		}
		return false;
	}

	/**
	 * Determines if the value passed in is of the
	 * correct data type
	 *
	 * @return
	 */
	public boolean isCorrectDataType( Object value )
	{
		switch( valType )
		{
			case ValidationHandle.VALUE_ANY:
				return true;
			case ValidationHandle.VALUE_DATE:
				if( value instanceof Date )
				{
					return true;
				}
				if( value instanceof Calendar )
				{
					return true;
				}
				break;
			case ValidationHandle.VALUE_DECIMAL:
				String possibleDec = value.toString();
				try
				{
					new Double( possibleDec );
					return true;
				}
				catch( NumberFormatException e )
				{
					return false;
				}
			case ValidationHandle.VALUE_FORMULA:
				String possibleFormula = value.toString();
				if( possibleFormula.indexOf( "=" ) == 0 )
				{
					return true;
				}
				break;
			case ValidationHandle.VALUE_INTEGER:
				String possibleInt = value.toString();
				if( possibleInt.indexOf( "." ) > -1 )
				{
					return false;
				}
				try
				{
					Long l = new Long( possibleInt );  // excel ints go past boundary of java ints, so use long
					return true;
				}
				catch( NumberFormatException e )
				{
					return false;
				}
			case ValidationHandle.VALUE_TEXT_LENGTH:
				return true;  //TODO
			case ValidationHandle.VALUE_TIME:
				return true;  // TODO
			case ValidationHandle.VALUE_USER_DEFINED_LIST:
				return true; // TODO
		}
		return false;
	}

	/**
	 * Create a dv record & populate with prototype bytes
	 *
	 * @return
	 */
	protected static XLSRecord getPrototype( WorkBook bk )
	{
		Dv dv = new Dv();
		dv.setOpcode( DV );
		dv.setData( dv.PROTOTYPE_BYTES );
		dv.setWorkBook( bk );
		dv.init();
		return dv;
	}

	/**
	 * Standard init method
	 *
	 * @see com.extentech.formats.XLS.XLSRecord#init()
	 */
	@Override
	public void init()
	{
		super.init();

		int offset = 0;
		grbit = ByteTools.readInt( this.getByteAt( offset++ ), this.getByteAt( offset++ ), this.getByteAt( offset++ ), this.getByteAt(
				offset++ ) );
		short strLen = ByteTools.readShort( this.getByteAt( offset ), this.getByteAt( offset + 1 ) );
		byte strGrbit = this.getByteAt( offset + 2 );
		if( (strGrbit & 0x1) == 0x1 )
		{
			strLen *= 2;
		}
		strLen += 3;
		byte[] namebytes = this.getBytesAt( offset, strLen );
		offset += strLen;
		dTitlePrompt = new Unicodestring();
		dTitlePrompt.init( namebytes, false );

		strLen = ByteTools.readShort( this.getByteAt( offset ), this.getByteAt( offset + 1 ) );
		strGrbit = this.getByteAt( offset + 2 );
		if( (strGrbit & 0x1) == 0x1 )
		{
			strLen *= 2;
		}
		strLen += 3;

		namebytes = this.getBytesAt( offset, strLen );
		offset += strLen;
		dTitleError = new Unicodestring();
		dTitleError.init( namebytes, false );

		strLen = ByteTools.readShort( this.getByteAt( offset ), this.getByteAt( offset + 1 ) );
		strGrbit = this.getByteAt( offset + 2 );
		if( (strGrbit & 0x1) == 0x1 )
		{
			strLen *= 2;
		}
		strLen += 3;
		namebytes = this.getBytesAt( offset, strLen );
		offset += strLen;
		dTextPrompt = new Unicodestring();
		dTextPrompt.init( namebytes, false );

		strLen = ByteTools.readShort( this.getByteAt( offset ), this.getByteAt( offset + 1 ) );
		strGrbit = this.getByteAt( offset + 2 );
		if( (strGrbit & 0x1) == 0x1 )
		{
			strLen *= 2;
		}
		strLen += 3;
		namebytes = this.getBytesAt( offset, strLen );
		offset += strLen;
		dTextError = new Unicodestring();
		dTextError.init( namebytes, false );

		int sz1 = ByteTools.readShort( this.getByteAt( offset++ ), this.getByteAt( offset++ ) );
		// unknown bytes
		garbageByteOne[0] = this.getByteAt( offset++ );
		garbageByteOne[1] = this.getByteAt( offset++ );
		byte[] formulaBytes = this.getBytesAt( offset, sz1 );
		firstCond = ExpressionParser.parseExpression( formulaBytes, this );
		offset += sz1;

		int sz2 = ByteTools.readShort( this.getByteAt( offset++ ), this.getByteAt( offset++ ) );
		// unknown bytes
		garbageByteTwo[0] = this.getByteAt( offset++ );
		garbageByteTwo[1] = this.getByteAt( offset++ );
		formulaBytes = this.getBytesAt( offset, sz2 );
		secondCond = ExpressionParser.parseExpression( formulaBytes, this );
		offset += sz2;

		numLocs = this.getByteAt( offset++ );
		cRangeList = new ArrayList();
		for( int i = 0; i < numLocs; i++ )
		{
			byte[] b = new byte[1];
			b[0] = 0x0;
			b = ByteTools.append( b, this.getBytesAt( offset, 8 ) );
			PtgArea p = new PtgArea( false );
			p.setParentRec( this );
			p.init( b );
			cRangeList.add( p );
			offset += 8;
		}

		// set all the grbit fields
		valType = (byte) ((grbit & BITMASK_VALTYPE));
		errStyle = (byte) ((grbit & BITMASK_ERRSTYLE) >> 4);
		IMEMode = (short) ((grbit & BITMASK_MDIMEMODE) >> 10);
		fStrLookup = ((grbit & BITMASK_FSTRLOOKUP) == BITMASK_FSTRLOOKUP);
		fAllowBlank = ((grbit & BITMASK_FALLOWBLANK) == BITMASK_FALLOWBLANK);
		fSuppressCombo = ((grbit & BITMASK_FSUPRESSCOMBO) == BITMASK_FSUPRESSCOMBO);
		fShowInputMsg = ((BITMASK_FSHOWINPUTMSG) == BITMASK_FSHOWINPUTMSG);
		fShowErrMsg = ((grbit & BITMASK_FSHOWERRORMSG) == BITMASK_FSHOWERRORMSG);
		typOperator = (byte) ((grbit & BITMASK_TYPOPERATOR) >> 20);
	}

	/**
	 * As most of these records have variable lengths we cannot just update part of
	 * the data for the record at a time,  we just don't want to keep up updates.  Rather than this,
	 * we'll update the entire record.  To keep processing down, update the record before streaming
	 * rather than on each internal record change.
	 */
	private void updateRecord()
	{
		this.updateGrbit();
		byte[] recbytes = new byte[0];

		byte[] tmp = ByteTools.cLongToLEBytes( grbit );
		recbytes = ByteTools.append( tmp, recbytes );
		recbytes = ByteTools.append( dTitlePrompt.read(), recbytes );
		recbytes = ByteTools.append( dTitleError.read(), recbytes );
		recbytes = ByteTools.append( dTextPrompt.read(), recbytes );
		recbytes = ByteTools.append( dTextError.read(), recbytes );

		// get the firstCond bytes
		tmp = new byte[0];
		for( int i = 0; i < firstCond.size(); i++ )
		{
			Object o = firstCond.elementAt( i );
			Ptg ptg = (Ptg) o;
			tmp = ByteTools.append( ptg.getRecord(), tmp );
		}
		// get the length and add in.
		short sz = (short) tmp.length;
		recbytes = ByteTools.append( ByteTools.shortToLEBytes( sz ), recbytes );
		// add garbage
		recbytes = ByteTools.append( garbageByteOne, recbytes );
		recbytes = ByteTools.append( tmp, recbytes );

		// get the secondCond bytes
		tmp = new byte[0];
		for( int i = 0; i < secondCond.size(); i++ )
		{
			Object o = secondCond.elementAt( i );
			Ptg ptg = (Ptg) o;
			tmp = ByteTools.append( ptg.getRecord(), tmp );
		}
		// get the length and add in.
		sz = (short) tmp.length;
		recbytes = ByteTools.append( ByteTools.shortToLEBytes( sz ), recbytes );
		// add garbage
		recbytes = ByteTools.append( garbageByteTwo, recbytes );
		recbytes = ByteTools.append( tmp, recbytes );

		tmp = new byte[1];
		if( cRangeList != null )
		{
			tmp[0] = (byte) cRangeList.size();
			;
			recbytes = ByteTools.append( tmp, recbytes );
			for( int i = 0; i < cRangeList.size(); i++ )
			{
				tmp = ((PtgArea) cRangeList.get( i )).getRecord();
				byte[] tmp2 = new byte[8];
				tmp[0] = 0;
				System.arraycopy( tmp, 0, tmp2, 0, tmp2.length );
				recbytes = ByteTools.append( tmp2, recbytes );
			}
			// there is a trailing zero, not sure why...
			if( cRangeList.size() > 0 )
			{
				tmp = new byte[1];
				tmp[0] = 0;
				recbytes = ByteTools.append( tmp, recbytes );
			}
		}
		else if( (ooxmlranges != null) && (ooxmlranges.length > 0) )
		{
			tmp[0] = (byte) ooxmlranges.length;
			recbytes = ByteTools.append( tmp, recbytes );
			for( int i = 0; i < ooxmlranges.length; i++ )
			{
				Ptg/*Ref*/ p = PtgRef.createPtgRefFromString( this.getSheet().getSheetName() + "!" + ooxmlranges[i], this );
				tmp = p.getRecord();
		            /* replace with above PtgArea pa= new PtgArea();
                    try {
                        pa.setParentRec(this);
                        pa.setLocation(this.getSheet().getSheetName() + "!" + ooxmlranges[i]);
                    } catch (Exception e) {
                        // TODO: handle MAXROWS/MAXCOLS
                    }                    
                    tmp = pa.getRecord();
                    /**/
				tmp[0] = 0;
				byte[] tmp2 = new byte[8];
				System.arraycopy( tmp, 0, tmp2, 0, tmp2.length );
				recbytes = ByteTools.append( tmp, recbytes );
			}
			// there is a trailing zero, not sure why...
			if( ooxmlranges.length > 0 )
			{
				tmp = new byte[1];
				tmp[0] = 0;
				recbytes = ByteTools.append( tmp, recbytes );
			}
		}

		this.setData( recbytes );
	}

	/**
	 * update record.
	 *
	 * @see com.extentech.formats.XLS.XLSRecord#preStream()
	 */
	@Override
	public void preStream()
	{
		if( dirtyflag )
		{
			this.updateRecord();
		}
	}

	/**
	 * Apply all the grbit fields into the current grbit int
	 */
	public void updateGrbit()
	{
		grbit = 0;
		grbit |= valType;
		grbit |= (errStyle << 4);
		grbit |= (IMEMode << 10);
		if( fStrLookup )
		{
			grbit = (grbit | BITMASK_FSTRLOOKUP);
		}

		if( fAllowBlank )
		{
			grbit = (grbit | BITMASK_FALLOWBLANK);
		}

		if( fSuppressCombo )
		{
			grbit = (grbit | BITMASK_FSUPRESSCOMBO);
		}

		if( fShowInputMsg )
		{
			grbit = (grbit | BITMASK_FSHOWINPUTMSG);
		}

		if( fShowErrMsg )
		{
			grbit = (grbit | BITMASK_FSHOWERRORMSG);
		}

		grbit |= (typOperator << 20);
	}

	/**
	 * Return the range of data this Dv refers to as a string array
	 * <p/>
	 * Values are stored as absolute ($) references, but should be displayed
	 * as relative
	 *
	 * @return ptgRef.toString()
	 */
	public String[] getRanges()
	{
		if( (cRangeList == null) && (ooxmlranges != null) )
		{
			if( ooxmlranges.length > 0 )
			{
				return ooxmlranges;
			}
		}
		String[] s = new String[cRangeList.size()];
		for( int i = 0; i < s.length; i++ )
		{
			s[i] = ((PtgArea) cRangeList.get( i )).getLocation();
			s[i] = StringTool.strip( s[i], "$" );
		}
		return s;

	}

	/**
	 * Set the range this Dv refers to.   Pass in a range string, sans worksheet
	 * Note that absolute ranges/ptrgrefs are always used, however returning
	 * values should not include the dollar sign
	 *
	 * @param range
	 */
	public void setRange( String range )
	{
		if( range == null )
		{    // for creating a dv and adding range info later
			cRangeList = null;
			return;
		}
		if( range.indexOf( ":" ) == -1 )
		{
			range = range + ":" + range;
		}
		PtgArea p = new PtgArea( range, this, false );
		cRangeList = new ArrayList();
		cRangeList.add( p );
		dirtyflag = true;
	}

	/**
	 * Add a range this Dv refers to.   Pass in a range string, sans worksheet
	 *
	 * @param range
	 */
	public void addRange( String range )
	{
		if( cRangeList == null )
		{
			cRangeList = new ArrayList();    // 20090605 KSC: Added
		}
     	/*int[] i = ExcelTools.getRowColFromString(range);
     	if(i.length==2) {
     	    
     	}else {*/
		PtgArea p = new PtgArea( range, this, false );    // 20090609 KSC: absolute refs if '$' -really should test if row or col
		cRangeList.add( p );
		dirtyflag = true;
	}

	/**
	 * Add a range this Dv refers to.   Pass in a range string, sans worksheet
	 * May need additional handling for records outside bounds?
	 *
	 * @param range
	 */
	public void addOoxmlRange( String range )
	{
		if( cRangeList == null )
		{
			cRangeList = new ArrayList();  // 20090605 KSC: Added
		}
		PtgArea p = new PtgArea( range,
		                         this,
		                         (range != null) ? (range.indexOf( '$' ) == -1) : false ); // 20090609 KSC: absolute refs if '$' -really should test if row or col
//        p.setUseReferenceTracker(false);
		cRangeList.add( p );
		dirtyflag = true;
	}

	/**
	 * Return the text in the error box
	 *
	 * @return
	 */
	public String getErrorBoxText()
	{
		return dTextError.toString().trim();
	}

	/**
	 * Set the text for the error box
	 *
	 * @param textError
	 */
	public void setErrorBoxText( String textError )
	{
		dTextError.updateUnicodeString( textError );
		dirtyflag = true;
	}

	/**
	 * Return the text in the prompt box
	 *
	 * @return
	 */
	public String getPromptBoxText()
	{
		return dTextPrompt.toString().trim();
	}

	/**
	 * Set the text for the prompt box
	 *
	 * @param text
	 */
	public void setPromptBoxText( String text )
	{
		dTextPrompt.updateUnicodeString( text );
		dirtyflag = true;
	}

	/**
	 * Set the title for the error box
	 *
	 * @param textError
	 */
	public void setErrorBoxTitle( String textError )
	{
		dTitleError.updateUnicodeString( textError );
		dirtyflag = true;
	}

	/**
	 * Return the title from the error box
	 *
	 * @return
	 */
	public String getErrorBoxTitle()
	{
		return dTitleError.toString().trim();
	}

	/**
	 * Return the title in the prompt box
	 *
	 * @return
	 */
	public String getPromptBoxTitle()
	{
		return dTitlePrompt.toString().trim();
	}

	/**
	 * Set the title for the prompt box
	 *
	 * @param text
	 */
	public void setPromptBoxTitle( String text )
	{
		dTitlePrompt.updateUnicodeString( text );
		dirtyflag = true;
	}

	/**
	 * Return a byte representing the error style for this DV
	 * <p/>
	 * These map to the static final ints ERROR_* from ValidationHandle
	 *
	 * @return
	 */
	public byte getErrorStyle()
	{
		return errStyle;
	}

	/**
	 * Set the error style for this Dv record
	 * <p/>
	 * These map to the static final ints ERROR_* from ValidationHandle
	 *
	 * @return
	 */
	public void setErrorStyle( byte errstyle )
	{
		errStyle = errstyle;
		dirtyflag = true;

	}

	/**
	 * Allow blank cells in the validation area?
	 *
	 * @return
	 */
	public boolean isAllowBlank()
	{
		return fAllowBlank;
	}

	/**
	 * Allow blank cells in the validation area?
	 *
	 * @return
	 */
	public void setAllowBlank( boolean allowBlank )
	{
		fAllowBlank = allowBlank;
		updateGrbit();
		dirtyflag = true;
	}

	/**
	 * Get the first condition of the validation as
	 * a string representation
	 *
	 * @return
	 */
	public String getFirstCond()
	{
		String s = FormulaParser.getExpressionString( firstCond );
		if( s.substring( 0, 1 ).equals( "=" ) )
		{
			return s.substring( 1, s.length() );
		}
		return s;
	}

	/**
	 * Set the first condition of the validation utilizing
	 * a string.  This value must conform to the Value Type of this
	 * validation or unexpected results may occur.  For example,
	 * entering a string representation of a date here will not work
	 * if your validation is an integer...
	 * <p/>
	 * String passed in should be a vaild XLS formula.  Does not need to include the "="
	 * <p/>
	 * Types of conditions
	 * Integer values
	 * Decimal values
	 * User defined list
	 * Date
	 * Time
	 * Text length
	 * Formula
	 *
	 * @return
	 */
	public void setFirstCond( String firstCond )
	{
		this.firstCond = FormulaParser.getPtgsFromFormulaString( (XLSRecord) this, firstCond );
		dirtyflag = true;
	}

	/**
	 * Get the second condition of the validation as
	 * a string representation
	 *
	 * @return
	 */
	public String getSecondCond()
	{
		String s = FormulaParser.getExpressionString( secondCond );
		if( s.substring( 0, 1 ).equals( "=" ) )
		{
			return s.substring( 1, s.length() );
		}
		return s;
	}

	/**
	 * Set the first condition of the validation utilizing
	 * a string.  This value must conform to the Value Type of this
	 * validation or unexpected results may occur.  For example,
	 * entering a string representation of a date here will not work
	 * if your validation is an integer...
	 * <p/>
	 * String passed in should be a vaild XLS formula.  Does not need to include the "="
	 * <p/>
	 * Types of conditions
	 * Integer values
	 * Decimal values
	 * User defined list
	 * Date
	 * Time
	 * Text length
	 * Formula
	 *
	 * @return
	 */
	public void setSecondCond( String secondCond )
	{
		this.secondCond = FormulaParser.getPtgsFromFormulaString( (XLSRecord) this, secondCond );
		dirtyflag = true;
	}

	/**
	 * Show error box if invalid values entered?
	 *
	 * @return
	 */
	public boolean isShowErrorMsg()
	{
		return fShowErrMsg;
	}

	/**
	 * set show error box if invalid values entered?
	 *
	 * @return
	 */
	public void setShowErrMsg( boolean showErrMsg )
	{
		fShowErrMsg = showErrMsg;
		dirtyflag = true;
	}

	/**
	 * Show prompt box if cell selected?
	 *
	 * @return
	 */
	public boolean getShowInputMsg()
	{
		return fShowInputMsg;
	}

	/**
	 * Set show prompt box if cell selected?
	 * *
	 *
	 * @return
	 */
	public void setShowInputMsg( boolean showInputMsg )
	{
		fShowInputMsg = showInputMsg;
		dirtyflag = true;
	}

	/**
	 * In list type validity the string list is explicitly given in the formula
	 *
	 * @return
	 */
	public boolean isStrLookup()
	{
		return fStrLookup;
	}

	/**
	 * In list type validity the string list is explicitly given in the formula
	 *
	 * @return
	 */
	public void setStrLookup( boolean strLookup )
	{
		fStrLookup = strLookup;
		dirtyflag = true;
	}

	/**
	 * Suppress the drop down arrow in list type validity
	 *
	 * @return
	 */
	public boolean isSuppressCombo()
	{
		return fSuppressCombo;
	}

	/**
	 * Get the IME mode for this validation
	 *
	 * @return
	 */
	public short getIMEMode()
	{
		return IMEMode;
	}

	/**
	 * set the IME mode for this validation
	 *
	 * @return
	 */
	public void setIMEMode( short mode )
	{
		IMEMode = mode;
		dirtyflag = true;
	}

	/**
	 * Suppress the drop down arrow in list type validity
	 *
	 * @return
	 */
	public void setSuppressCombo( boolean suppressCombo )
	{
		fSuppressCombo = suppressCombo;
		dirtyflag = true;
	}

	/**
	 * Get the type operator of this validation as a byte.
	 * <p/>
	 * These bytes map to the CONDITION_* static values in
	 * ValidationHandle
	 *
	 * @return
	 */
	public byte getTypeOperator()
	{
		return typOperator;
	}

	/**
	 * set the type operator of this validation as a byte.
	 * <p/>
	 * These bytes map to the CONDITION_* static values in
	 * ValidationHandle
	 *
	 * @return
	 */
	public void setTypeOperator( byte typOperator )
	{
		this.typOperator = typOperator;
		dirtyflag = true;
	}

	/**
	 * Get the validation type of this Dv as a byte
	 * <p/>
	 * These bytes map to the VALUE_* static values in
	 * ValidationHandle
	 *
	 * @return
	 */
	public byte getValType()
	{
		return valType;
	}

	/**
	 * Set the validation type of this Dv as a byte
	 * <p/>
	 * These bytes map to the VALUE_* static values in
	 * ValidationHandle
	 *
	 * @return
	 */
	public void setValType( byte valtype )
	{
		valType = valtype;
		dirtyflag = true;
	}

	/**
	 * Determines if the Dv contains the cell address passed in
	 *
	 * @param range
	 * @return
	 */
	public boolean isInRange( String celladdy )
	{
		// FIX broken COLROW
		int[] rc = ExcelTools.getRowColFromString( celladdy );
		for( int i = 0; i < cRangeList.size(); i++ )
		{
			if( ((PtgArea) cRangeList.get( i )).contains( rc ) )
			{
				return true;
			}
		}
		return false;
	}
	/**
	 *  OOXML Element:
	 * dataValidation (Data Validation)
	 * A single item of data validation defined on a range of the worksheet
	 *
	 * parent: dataValidations (==Dval)
	 * children: formula1, formula2
	 * attributes:  many
	 */

	/**
	 * create a new Dv record based on OOXML input
	 */
	public static Dv parseOOXML( XmlPullParser xpp, Boundsheet bs )
	{
		Dv dv = bs.createDv( null );
		dv.setSheet( bs );
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "dataValidation" ) )
					{        // get attributes
						for( int i = 0; i < xpp.getAttributeCount(); i++ )
						{
							String n = xpp.getAttributeName( i );
							String v = xpp.getAttributeValue( i );
							if( n.equals( "allowBlank" ) )
							{
								dv.setAllowBlank( true );
							}
							else if( n.equals( "error" ) )
							{
								dv.setErrorBoxText( v );
							}
							else if( n.equals( "errorStyle" ) )
							{    // default= stop
								if( v.equals( "information" ) )
								{
									dv.setErrorStyle( (byte) 2 );
								}
								else if( v.equals( "stop" ) )
								{
									dv.setErrorStyle( (byte) 0 );
								}
								else if( v.equals( "warning" ) )
								{
									dv.setErrorStyle( (byte) 1 );
								}
							}
							else if( n.equals( "errorTitle" ) )
							{
								dv.setErrorBoxTitle( v );
							}
							else if( n.equals( "imeMode" ) )
							{    // TODO: what is the correct mapping??????????????????????
								if( v.equals( "nocontrol" ) )
								{
									dv.setIMEMode( (short) 0 );
								}
								else if( v.equals( "off" ) )
								{
									dv.setIMEMode( (short) 1 );
								}
								else if( v.equals( "on" ) )
								{
									dv.setIMEMode( (short) 2 );
								}
								else if( v.equals( "disabled" ) )
								{
									dv.setIMEMode( (short) 3 );
								}
								else if( v.equals( "hiragana" ) )
								{
									dv.setIMEMode( (short) 4 );
								}
								else if( v.equals( "fullKatakana" ) )
								{
									dv.setIMEMode( (short) 5 );
								}
								else if( v.equals( "halfKatakana" ) )
								{
									dv.setIMEMode( (short) 6 );
								}
								else if( v.equals( "fullAlpha" ) )
								{
									dv.setIMEMode( (short) 7 );
								}
								else if( v.equals( "halfAlpha" ) )
								{
									dv.setIMEMode( (short) 8 );
								}
								else if( v.equals( "fullHangul" ) )
								{
									dv.setIMEMode( (short) 9 );
								}
								else if( v.equals( "halfHangul" ) )
								{
									dv.setIMEMode( (short) 10 );
								}
							}
							else if( n.equals( "operator" ) )
							{    // default= "between"
								if( v.equals( "between" ) )
								{
									dv.setTypeOperator( (byte) 0 );
								}
								else if( v.equals( "equal" ) )
								{
									dv.setTypeOperator( (byte) 2 );
								}
								else if( v.equals( "greaterThan" ) )
								{
									dv.setTypeOperator( (byte) 4 );
								}
								else if( v.equals( "greaterThanOrEqual" ) )
								{
									dv.setTypeOperator( (byte) 6 );
								}
								else if( v.equals( "lessThan" ) )
								{
									dv.setTypeOperator( (byte) 5 );
								}
								else if( v.equals( "lessThanOrEqual" ) )
								{
									dv.setTypeOperator( (byte) 7 );
								}
								else if( v.equals( "notBetween" ) )
								{
									dv.setTypeOperator( (byte) 1 );
								}
								else if( v.equals( "notEqual" ) )
								{
									dv.setTypeOperator( (byte) 3 );
								}
							}
							else if( n.equals( "prompt" ) )
							{
								dv.setPromptBoxText( v );
							}
							else if( n.equals( "promptTitle" ) )
							{
								dv.setPromptBoxTitle( v );
							}
							else if( n.equals( "showDropDown" ) )
							{
//		            			
							}
							else if( n.equals( "showErrorMessage" ) )
							{
								dv.setShowErrMsg( true );
							}
							else if( n.equals( "showInputMessage" ) )
							{
								dv.setShowInputMsg( true );
							}
							else if( n.equals( "sqref" ) )
							{
								dv.ooxmlranges = StringTool.splitString( v, " " );
								// 20090609 KSC: cannot add ranges in 2003 format as 2007 addresses can exceed 2003 limits
								for( int z = 0; z < dv.ooxmlranges.length; z++ )
								{
									dv.addOoxmlRange( dv.ooxmlranges[z] );
								}
							}
							else if( n.equals( "type" ) )
							{        // required
								if( v.equals( "custom" ) )        // custom formula
								{
									dv.setValType( (byte) 7 );
								}
								else if( v.equals( "date" ) )
								{
									dv.setValType( (byte) 4 );
								}
								else if( v.equals( "decimal" ) )
								{
									dv.setValType( (byte) 2 );
								}
								else if( v.equals( "list" ) )
								{
									dv.setValType( (byte) 3 );
								}
								else if( v.equals( "none" ) )
								{
									dv.setValType( (byte) 0 );
								}
								else if( v.equals( "textLength" ) )
								{
									dv.setValType( (byte) 6 );
								}
								else if( v.equals( "time" ) )
								{
									dv.setValType( (byte) 5 );
								}
								else if( v.equals( "whole" ) )
								{
									dv.setValType( (byte) 1 );
								}
							}
						}
					}
					else if( tnm.equals( "formula1" ) )
					{
						dv.setFirstCond( OOXMLAdapter.getNextText( xpp ) );
					}
					else if( tnm.equals( "formula2" ) )
					{
						dv.setSecondCond( OOXMLAdapter.getNextText( xpp ) );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "dataValidation" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "OOXMLELEMENT.parseOOXML: " + e.toString() );
		}
		return dv;
	}

	/**
	 * generate the proper OOXML to define this Dv
	 *
	 * @return
	 */
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<dataValidation" );
		switch( valType )
		{         // required
			case 0:    // any ????
// TODO: this maps to ???    		
				//ooxml.append(" type=");
				break;
			case 1:
				ooxml.append( " type=\"whole\"" );
				break;
			case 2:
				ooxml.append( " type=\"decimal\"" );
				break;
			case 3:
				ooxml.append( " type=\"list\"" );
				break;
			case 4:
				ooxml.append( " type=\"date\"" );
				break;
			case 5:
				ooxml.append( " type=\"time\"" );
				break;
			case 6:
				ooxml.append( " type=\"textLength\"" );
				break;
			case 7:
				ooxml.append( " type=\"custom\"" );
				break;
		}
		switch( typOperator )
		{
			case 0:  // default, leave out
				// ooxml.append(" operator=\"between\"");
				break;
			case 1:
				ooxml.append( " operator=\"notBetween\"" );
				break;
			case 2:
				ooxml.append( " operator=\"equal\"" );
				break;
			case 3:
				ooxml.append( " operator=\"notEqual\"" );
				break;
			case 4:
				ooxml.append( " operator=\"greaterThan\"" );
				break;
			case 5:
				ooxml.append( " operator=\"lessThan\"" );
				break;
			case 6:
				ooxml.append( " operator=\"greaterThanOrEqual\"" );
				break;
			case 7:
				ooxml.append( " operator=\"lessThanOrEqual\"" );
				break;
		}
		switch( errStyle )
		{
			case 0:
				// default no need to outut ooxml.append(" errorStyle=\"stop\"");
				break;
			case 1:
				ooxml.append( " errorStyle=\"warning\"" );
				break;
			case 2:
				ooxml.append( " errorStyle=\"information\"" );
				break;
		}
		if( !this.getErrorBoxText().equals( "" ) )
		{
			ooxml.append( " error=\"" + OOXMLAdapter.stripNonAscii( this.getErrorBoxText() ) + "\"" );
		}
		if( !this.getErrorBoxTitle().equals( "" ) )
		{
			ooxml.append( " errorTitle=\"" + OOXMLAdapter.stripNonAscii( this.getErrorBoxTitle() ) + "\"" );
		}
		//TODO "imeMode"
		if( !this.getPromptBoxText().equals( "" ) )
		{
			ooxml.append( " prompt=\"" + OOXMLAdapter.stripNonAscii( this.getPromptBoxText() ) + "\"" );
		}
		if( !this.getPromptBoxTitle().equals( "" ) )
		{
			ooxml.append( " promptTitle=\"" + OOXMLAdapter.stripNonAscii( this.getPromptBoxTitle() ) + "\"" );
		}
		// This needs to be better thought out, currently it breaks/strips all changes made to the model, as ranges
		// are not automatically added to ooxml ranges.
		/**if (ooxmlranges!=null) {// have stored OOXML ranges
		 ooxml.append(" sqref=\"");
		 for (int i= 0; i < ooxmlranges.length; i++) {
		 if (i>0) ooxml.append(" ");
		 ooxml.append(ooxmlranges[i]);
		 }
		 ooxml.append("\"");
		 } else {	// 2003-style ranges
		 **/
		String[] ranges = this.getRanges();
		if( ranges.length > 0 )
		{
			ooxml.append( " sqref=\"" );
			for( int i = 0; i < ranges.length; i++ )
			{
				if( i > 0 )
				{
					ooxml.append( " " );
				}
				ooxml.append( ranges[i] );
			}
			ooxml.append( "\"" );
		}
		//}

		if( this.isAllowBlank() )
		{
			ooxml.append( " allowBlank=\"1\"" );
		}
		if( this.isShowErrorMsg() )
		{
			ooxml.append( " showErrorMessage=\"1\"" );
		}
		if( this.getShowInputMsg() )
		{
			ooxml.append( " showInputMessage=\"1\"" );
		}
/*    	
	"showDropDown"
*/
		/**
		 * imwMode
		 * TODO: map options correctly!!  where is the info??
		 */
		switch( this.getIMEMode() )
		{
			case 0:    // nocontrol
				break;
			case 1:
				ooxml.append( " imeMode=\"off\"" );
				break;
			case 2:
				ooxml.append( " imeMode=\"on\"" );
				break;
			case 3:
				ooxml.append( " imeMode=\"disabled\"" );
				break;
			case 4:
				ooxml.append( " imeMode=\"hiragana\"" );
				break;
			case 5:
				ooxml.append( " imeMode=\"fullKatakana\"" );
				break;
			case 6:
				ooxml.append( " imeMode=\"halfKatakana\"" );
				break;
			case 7:
				ooxml.append( " imeMode=\"fullAlpha\"" );
				break;
			case 8:
				ooxml.append( " imeMode=\"halfAlpha\"" );
				break;
			case 9:
				ooxml.append( " imeMode=\"fullHangul\"" );
				break;
			case 10:
				ooxml.append( " imeMode=\"halfHangul\"" );
				break;
		}
		ooxml.append( ">" );
		String formula1 = this.getFirstCond();
		if( (formula1 != null) && (formula1.length() > 0) )
		{
			formula1 = formula1.replace( (char) 0, ',' );    // DV Lists are delimited by 0 must replace with commas for OOXML use
			ooxml.append( "<formula1>" + formula1 + "</formula1>" );
		}
		String formula2 = this.getSecondCond();
		if( (formula2 != null) && (formula2.length() > 0) )
		{
			formula2 = formula2.replace( (char) 0, ',' );    // DV Lists are delimited by 0 must replace with commas for OOXML use
			ooxml.append( "<formula2>" + formula2 + "</formula2>" );
		}
		ooxml.append( "</dataValidation>" );
		return ooxml.toString();
	}
}
