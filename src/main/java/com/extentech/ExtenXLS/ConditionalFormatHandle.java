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
package com.extentech.ExtenXLS;

import com.extentech.formats.XLS.Cf;
import com.extentech.formats.XLS.Condfmt;
import com.extentech.toolkit.Logger;

import java.util.ArrayList;

/**
 * <pre>
 * ConditionalFormatHandle allows for manipulation of the ConditionalFormat cells in Excel
 *
 * Using the ConditionalFormatHandle, the affected range of ConditionalFormats can be modified,
 * along with the formatting applied to the cells when the condition is true.
 *
 * Each ConditionalFormatHandle represents a range of cells and can have a number
 * of formatting rules and formats (ConditionalFormatRule) applied.
 *
 * The ConditionalFormatHandle affected range can either be a contiguous range, or a series of cells and ranges.
 *
 * Each ConditionalFormatRule contains one rule and corresponding format data.
 *
 *
 * Many of these calls are very self-explanatory and can be found in the api.
 * </pre>
 */
public class ConditionalFormatHandle implements Handle
{

	private Condfmt cndfmt = null;
	private WorkSheetHandle worksheet;

	/**
	 * evaluates the criteria for this Conditional Format
	 * <br>if the criteria involves a comparison i.e. equals, less than, etc., it uses the
	 * value from the passed in referenced cell to compare with
	 * <p/>
	 * If there are multiple rules in the ConditionalFormat then the first rule that passes will be returned.
	 * <p/>
	 * If no valid rules pass then a null result is given
	 *
	 * @param CellHandle refcell - the cell to obtain a value from in order for evaluation to occur
	 * @return the ConditionalFormatRule that passes.
	 * @see com.extentech.formats.XLS.Cf#evaluate(com.extentech.formats.XLS.formulas.Ptg)
	 */
	public ConditionalFormatRule evaluate( CellHandle refcell )
	{
		ConditionalFormatRule[] rules = this.getRules();
		for( ConditionalFormatRule rule : rules )
		{
			if( rule.evaluate( refcell ) )
			{
				return rule;
			}
		}
		return null;
	}

	/**
	 * Get all the rules assocated with this conditional format record
	 *
	 * @return
	 */
	public ConditionalFormatRule[] getRules()
	{
		ArrayList cfs = this.cndfmt.getRules();
		ConditionalFormatRule[] rules = new ConditionalFormatRule[cfs.size()];
		for( int i = 0; i < cfs.size(); i++ )
		{
			ConditionalFormatRule cfr = new ConditionalFormatRule( (Cf) cfs.get( i ) );
			rules[i] = cfr;
		}
		return rules;
	}

	/**
	 * For internal use only.  Creates a ConditionalFormat Handle based of the Condfmt passed in.
	 *
	 * @param workBookHandle
	 * @param Condfmt
	 */
	protected ConditionalFormatHandle( Condfmt c, WorkSheetHandle workSheetHandle )
	{
		this.cndfmt = c;
		worksheet = workSheetHandle;
	}

	/**
	 * get the WorkSheetHandle for this ConditionalFormat
	 * <p/>
	 * ConditionalFormats are bound to a specific worksheet and cannot be
	 * applied to multiple worksheets
	 *
	 * @return the WorkSheetHandle for this ConditionalFormat
	 */
	public WorkSheetHandle getWorkSheetHandle()
	{
		return worksheet;
	}

	/**
	 * Return the range of data this ConditionalFormatHandle refers to as a string
	 * This location is the largest bounding rectangle that all cells utilized in this conditional
	 * format can be contained in.
	 *
	 * @return Encompassing range in the format "A2:B12"
	 */
	public String getEncompassingRange()
	{
		int[] rowcols = cndfmt.getEncompassingRange();
		return ExcelTools.formatRangeRowCol( rowcols );
	}

	/**
	 * Return a string representing all ranges that this conditional format handle can affect
	 *
	 * @return range in the format "A2:B3";
	 */
	public String[] getAllAffectedRanges()
	{
		return cndfmt.getAllRanges();
	}

	/**
	 * Determine if the conditional format contains/affects the cell handle passed in
	 *
	 * @param cellHandle
	 * @return
	 */
	public boolean contains( CellHandle cellHandle )
	{
		return this.cndfmt.contains( cellHandle.getIntLocation() );
	}

	/**
	 * Set the range this ConditionalFormatHandle refers to.
	 * Pass in a range string, sans worksheet.
	 * <p/>
	 * This range will overwrite all other ranges this ConditionalFormatHandle refers to.
	 * <p/>
	 * In order to handle multiple ranges, use the addRange(String range) method
	 *
	 * @param range = standard excel range without worksheet information ("A1" or "A1:A10")
	 */
	public void setRange( String range )
	{
		this.cndfmt.resetRange( range );
	}

	/**
	 * Determines if the ConditionalFormatHandle contains the cell address passed in
	 *
	 * @param cellAddress a cell address in the format "A1"
	 * @return if the ce
	 */
	public boolean contains( String celladdy )
	{
		return this.cndfmt.contains( ExcelTools.getRowColFromString( celladdy ) );
	}

	/**
	 * Return an xml representation of the ConditionalFormatHandle
	 *
	 * @return
	 */
	// TODO: more than one cf rule???
	public String getXML()
	{
		ConditionalFormatRule[] rules = this.getRules();
		StringBuffer xml = new StringBuffer();
		xml.append( "<dataConditionalFormat" );
		xml.append( " type=\"" + ConditionalFormatRule.VALUE_TYPE[rules[0].getConditionalFormatType()] + "\"" );
		xml.append( " operator=\"" + ConditionalFormatRule.OPERATORS[rules[0].getTypeOperator()] + "\"" );
		try
		{
			xml.append( " sqref=\"" + this.getEncompassingRange() + "\"" );
		}
		catch( Exception e )
		{
			Logger.logErr( "Problem getting range for ConditionalFormatHandle.getXML().", e );
		}
		xml.append( ">" );
		if( rules[0].getFirstCondition() != null )
		{
			xml.append( "<formula1>" );
			xml.append( rules[0].getFirstCondition() );
			xml.append( "</formula1>" );
		}
		if( rules[0].getSecondCondition() != null )
		{
			xml.append( "<formula2>" );
			xml.append( rules[0].getSecondCondition() );
			xml.append( "</formula2>" );
		}
		xml.append( "</dataConditionalFormat>" );
		return xml.toString();
	}

	/**
	 * return a string representation of this Conditional Format
	 * <p/>
	 * This method is still incomplete as it only returns data for one rule, and only refers to one range
	 */
	// TODO: more than one cf rule???
	public String toString()
	{
		ConditionalFormatRule[] rules = this.getRules();
		String ret = this.getEncompassingRange() + ": " + rules[0].getType() + " " + rules[0].getOperator();    // range, type + operator
		if( rules[0].getFormula1() != null )    // formulas
		{
			ret += " " + rules[0].getFormula1().getFormulaString().substring( 1 );
		}
		if( rules[0].getFormula2() != null )
		{
			ret += " and " + rules[0].getFormula2().getFormulaString().substring( 1 );
		}
		// todo: add formats to this
		return ret;
	}

	/**
	 * @return Returns the cndfmt.
	 */
	protected Condfmt getCndfmt()
	{
		return cndfmt;
	}

	/**
	 * @param cndfmt The cndfmt to set.
	 */
	protected void setCndfmt( Condfmt cndfmt )
	{
		this.cndfmt = cndfmt;
	}

	/**
	 * Add a cell to this conditional format record
	 *
	 * @param cellHandle
	 */
	public void addCell( CellHandle cellHandle )
	{
		if( this.contains( cellHandle ) )
		{
			return;
		}
		cndfmt.addLocation( cellHandle.getCellAddress() );
	}

	/**
	 * returns the formatting for each rule of this Contditional Format Handle
	 *
	 * @return FormatHandle[]
	 */
	public FormatHandle[] getFormats()
	{
		FormatHandle[] fmx = new FormatHandle[this.getCndfmt().getRules().size()];
		if( cndfmt.getFormatHandle() == null )
		{
			//cfm.initCells(this);	// added!
			int cfxe = cndfmt.getCfxe();
			FormatHandle fz = new FormatHandle( cndfmt, worksheet.wbh, cfxe, null );
		}
		for( int t = 0; t < fmx.length; t++ )
		{
			cndfmt.getFormatHandle().updateFromCF( (Cf) this.getCndfmt().getRules().get( t ), worksheet.wbh );
			fmx[t] = cndfmt.getFormatHandle();
		}
		return fmx;
	}

}
