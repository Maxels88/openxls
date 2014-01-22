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

/**
 * Record containing miscellaneous sheet-level boolean values.
 * <p/>
 * A - fShowAutoBreaks (1 bit): A bit that specifies whether page breaks (2) inserted automatically are visible on the sheet.
 * B - reserved1 (3 bits): MUST be zero, and MUST be ignored.
 * C - fDialog (1 bit): A bit that specifies whether the sheet is a dialog sheet.
 * D - fApplyStyles (1 bit): A bit that specifies whether to apply styles in an outline when an outline is applied.
 * E - fRowSumsBelow (1 bit): A bit that specifies whether summary rows appear below an outline's detail rows.
 * F - fColSumsRight (1 bit): A bit that specifies whether summary columns appear to the right or left of an outline's detail columns. Valid values are specified in the following table:
 * Value	    Meaning
 * 0			The summary columns appear to the right, if the sheet is displayed left-to-right, or appear to the left, if the sheet is displayed right-to-left.
 * 1		    The summary columns appear to the left, if the sheet is displayed left-to-right, or appear to the right, if the sheet is displayed right-to-left.
 * G - fFitToPage (1 bit): A bit that specifies whether to fit the printable contents to a single page when printing this sheet.
 * H - reserved2 (1 bit): MUST be zero, and MUST be ignored.
 * I - unused (2 bits): Undefined and MUST be ignored.
 * J - fSyncHoriz (1 bit): A bit that specifies whether horizontal scrolling is synchronized across multiple windows displaying this sheet.
 * K - fSyncVert (1 bit): A bit that specifies whether vertical scrolling is synchronized across multiple windows displaying this sheet.
 * L - fAltExprEval (1 bit): A bit that specifies whether the sheet uses transition formula evaluation.
 * M - fAltFormulaEntry (1 bit): A bit that specifies whether the sheet uses transition formula entry.
 */
public final class WsBool extends XLSRecord
{
	private static final long serialVersionUID = 2794181135988750779L;

	public void init()
	{
		super.init();
		// Make sure the data array is read in
		getData();

	}

	public void setSheet( Sheet sheet )
	{
		super.setSheet( sheet );
		((Boundsheet) sheet).addPrintRec( this );
	}

	/**
	 * Gets whether the sheet will be printed fit to some number of pages.
	 */
	public boolean isFitToPage()
	{
		return (data[1] & 0x01) == 0x01;
	}

	/**
	 * Sets whether the sheet will be printed fit to some number of pages.
	 */
	public void setFitToPage( boolean value )
	{
		if( value )
		{
			data[1] |= 0x01;
		}
		else
		{
			data[1] &= ~0x01;
		}
	}
}