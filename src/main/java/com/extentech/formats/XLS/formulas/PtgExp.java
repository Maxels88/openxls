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
import com.extentech.formats.XLS.Array;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.FormulaNotFoundException;
import com.extentech.formats.XLS.Shrfmla;
import com.extentech.formats.XLS.StringRec;
import com.extentech.toolkit.ByteTools;

import java.util.Stack;

/**
 * ptgExp indicates an Array Formula or Shared Formula
 * <p/>
 * When ptgExp occurs in a formula, it's the only token in the formula.
 * this indicates that the cell containing the formula
 * is part of an array or opartof a shared formula.
 * The actual formula is found in an array record.
 * <p/>
 * The value for ptgExp consists of the row and the column of the
 * upper-left corner of the array formula.
 *
 * @see Ptg
 * @see Formula
 * @see Array
 * @see Shrfmla
 */
public class PtgExp extends GenericPtg implements Ptg
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2150560716287810448L;
	int rwFirst;
	int colFirst;

	@Override
	public boolean getIsControl()
	{
		return true;
	}

	@Override
	public boolean getIsStandAloneOperator()
	{
		return true;
	}

	@Override
	public int getLength()
	{
		return PTG_EXP_LENGTH;
	}

	public int getRwFirst()
	{
		return rwFirst;
	}

	public int getColFirst()
	{
		return colFirst;
	}

	/**
	 * init from row, col
	 *
	 * @param row
	 * @param col
	 */
	public void init( int row, int col )
	{
		byte[] r = ByteTools.shortToLEBytes( (short) row );
		byte[] c = ByteTools.shortToLEBytes( (short) col );
		record = new byte[]{ 0x1, r[0], r[1], c[0], c[1] };
		ptgId = record[0];
		this.populateVals();
		//this.addToReferenceTracker();
	}

	@Override
	public void init( byte[] b )
	{
		ptgId = b[0];
		record = b;
		this.populateVals();
		//this.addToReferenceTracker();
	}

	private void populateVals()
	{
		rwFirst = readRow( record[1], record[2] );
		colFirst = ByteTools.readShort( record[3], record[4] );
	}

	/**
	 * Returns the location this PtgExp points to.
	 */
	public String getReferent()
	{
		return ExcelTools.formatLocation( new int[]{ rwFirst, colFirst } );
	}

	/**
	 * Looks up into it's parent shared formula, and returns
	 * the expression as if it were a regular formula.
	 *
	 * @return converted Calculation Expression
	 */
	public Ptg[] getConvertedExpression()
	{
		Formula f = (Formula) this.getParentRec();
		if( f.isSharedFormula() )
		{
			Stack expression = f.shared.instantiate( f );
			Ptg[] retPtg = new Ptg[expression.size()];
			for( int i = 0; i < expression.size(); i++ )
			{
				Ptg p = (Ptg) expression.get( i );
				retPtg[i] = p;
			}
			return retPtg;
//			throw new UnsupportedOperationException (
//					"Shared formulas must be instantiated for calculation");
		}
		else
		{    // if it's an array formula, return ptg's as well
			Array a = (Array) (f.getInternalRecords().get( 0 ));  // this.getParentRec().getSheet().getArrayFormula(getParentLocation());
			Stack calcStack = a.getExpression();
			Ptg[] retPtg = new Ptg[calcStack.size()];
			for( int i = 0; i < calcStack.size(); i++ )
			{
				Ptg p = (Ptg) calcStack.get( i );
				retPtg[i] = p;
			}
			return retPtg;
		}
	}

	@Override
	public Object getValue()
	{
		Object o = null;
		;
		Formula f = (Formula) this.getParentRec();
		if( f.isSharedFormula() )
		{
			o = FormulaCalculator.calculateFormula( f.shared.instantiate( f ) );
//			throw new UnsupportedOperationException (
//					"Shared formulas must be instantiated for calculation");
		}
		else
		{
			Object r = null;
			if( f.getInternalRecords().size() > 0 )
			{
				r = f.getInternalRecords().get( 0 );
			}
			else
			{    // it's part of an array formula but not the parent
				r = this.getParentRec().getSheet().getArrayFormula( getReferent() );
			}
			if( r instanceof Array )
			{
				Array arr = (Array) r;
				o = arr.getValue( this );
			}
			else if( r instanceof StringRec )
			{
				o = ((StringRec) r).getStringVal();
			}
		}
		return o;
	}

	@Override
	public Ptg calculatePtg( Ptg[] parsething )
	{
		Object o = null;
		;
		Formula f = ((Formula) this.getParentRec());
		if( f.isSharedFormula() )
		{
			o = FormulaCalculator.calculateFormula( f.shared.instantiate( f ) );
		}
		else
		{
			Object r = null;
			if( f.getInternalRecords().size() > 0 )
			{
				r = f.getInternalRecords().get( 0 );
			}
			else
			{    // it's part of an array formula but not the parent
				r = this.getParentRec().getSheet().getArrayFormula( getReferent() );
			}
			if( r instanceof Array )
			{
				Array arr = (Array) r;
				o = arr.getValue( this );
			}
			else if( r instanceof StringRec )
			{
				o = ((StringRec) r).getStringVal();
			}
			else    // should never happen
			{
				throw new UnsupportedOperationException( "Expected records parsing Formula were not present" );
			}
		}
		Ptg p = null;
		// conversion isn't necessary 
//		try{
		if( o instanceof Integer )
		{
			return new PtgInt( ((Integer) o).intValue() );
		}
//			Double d = new Double(o.toString());
		else if( o instanceof Double )
		{
			return new PtgNumber( ((Double) o).doubleValue() );
		}
		//p = new PtgNumber(d.doubleValue());
//		}catch(NumberFormatException e){
		if( o.toString().equalsIgnoreCase( "true" ) || o.toString().equalsIgnoreCase( "false" ) )
		{
			p = new PtgBool( o.toString().equalsIgnoreCase( "true" ) );
		}
		else
		{
			p = new PtgStr( o.toString() );
		}
//		}
		return p;
	}

	/**
	 * return the location of this PtgExp
	 * 20060302 KSC
	 */
	@Override
	public String getLocation() throws FormulaNotFoundException
	{
		String s = "";
		try
		{
			s = this.getParentRec().getCellAddress();
			s = this.getParentRec().getSheet().getSheetName() + "!" + s;
		}
		catch( Exception e )
		{

		}
		return s;
	}

	/**
	 * return the human-readable String representation of the linked shared formula
	 */
	@Override
	public String getString()
	{
		try
		{
			try
			{
				// Object o= ((Formula) this.getParentRec()).getInternalRecords().get(0);	PARENT REC of ARRAY or SHRFMLA is determined by referent (record) NOT necessarily same as actual Parent Rec
				Boundsheet sht = this.getParentRec().getSheet();
//	    		Formula pr= (Formula) sht.getCell(this.getReferent());
				Formula f = ((Formula) this.getParentRec());
				Object o;
				if( f.isSharedFormula() )
				{
					o = FormulaParser.getExpressionString( f.shared.instantiate( f ) );
					if( (o != null) && o.toString().startsWith( "=" ) )
					{
						return o.toString().substring( 1 );
					}
					return o.toString();
				}
				Formula pr = (Formula) sht.getCell( this.getReferent() );
				o = pr.getInternalRecords().get( 0 );
				if( o instanceof Array )
				{
					Array a = (Array) o;
					return a.getFormulaString();
				}
				else if( o instanceof StringRec )
				{
					//if this is a shared formula the attached string is the RESULT, not the formula string itself
					if( ((Formula) this.getParentRec()).isSharedFormula() )
					{
						throw new IndexOutOfBoundsException( "parse it" );
					}
					StringRec s = (StringRec) o;
					return s.getStringVal();
				}
			}
			catch( IndexOutOfBoundsException e )
			{    // subsequent formulas use same shared formula rec so find
				throw new UnsupportedOperationException( "Shared formulas must be instantiated for calculation" );
			}
		}
		catch( Exception e )
		{
			return "Array-Entered or Shared Formula";
		}
		return "Array-Entered or Shared Formula";
	}

	/**
	 * updateRecord from local rwFirst and colFirst values
	 *
	 * @see
	 */
	@Override
	public void updateRecord()
	{
		System.arraycopy( ByteTools.shortToLEBytes( (short) rwFirst ), 0, record, 1, 2 );
		System.arraycopy( ByteTools.shortToLEBytes( (short) colFirst ), 0, record, 3, 2 );
	}

	/**
	 * setLocation vars from address string
	 *
	 * @param s String address
	 */
	@Override
	public void setLocation( String s )
	{
		int[] rc = ExcelTools.getRowColFromString( s );
		rwFirst = rc[0];
		colFirst = rc[1];
		updateRecord();
	}

	public void setColFirst( int c )
	{
		this.colFirst = c;
	}

	public void setRowFirst( int r )
	{
		this.rwFirst = r;
	}

	public String toString()
	{
		return "PtgExp: Parent Formula at [" + rwFirst + "," + colFirst + "]";
	}

}