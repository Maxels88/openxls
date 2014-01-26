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
/*
 * Created on Dec 16, 2004
 *
 * To change the template for this generated file go to
 * Window&gt;Preferences&gt;Java&gt;Code Generation&gt;Code and Comments
 */
package org.openxls.formats.XLS;

import org.openxls.formats.XLS.charts.Chart;

import java.io.IOException;
import java.util.Collection;
import java.util.List;

/**
 *
 *
 */
public interface Sheet extends BiffRec
{

	// returns the records belonging to the sheet
	public List getSheetRecs();

	// hidden states from grbit field offset 1
	public static byte VISIBLE = 0x00;

	public static byte HIDDEN = 0x01;

	public static byte VERY_HIDDEN = 0x02;

	// sheet types from grbit field offset 0
	public static byte TYPE_SHEET_DIALOG = 0x00;

	public static byte TYPE_XL4_MACRO = 0x01;

	public static byte TYPE_CHART = 0x02;

	public static byte TYPE_VBMODULE = 0x06;

	public Eof getMyEof();

	public abstract Headerrec getHeader();

	public abstract void setHeader( BiffRec hr );

	public abstract Footerrec getFooter();

	public void setFooter( BiffRec r );

	/**
	 * get the last BiffRec added to this sheet
	 */
	public abstract BiffRec getLastCell();

	/**/
	public abstract List getLocalRecs();

	/**
	 * do all of the expensive updating here
	 * only right before streaming record.
	 */
	@Override
	public abstract void preStream();

	public abstract Bof getMyBof();

	public abstract void setWindow2( Window2 w );

	public abstract Window2 getWindow2();

	/**
	 * Remove a BiffRec from this WorkSheet.
	 */
	public abstract void removeCell( String celladdr );

	/**
	 * Remove a BiffRec from this WorkSheet.
	 */
	public abstract void removeCell( BiffRec c );

	/**
	 * remove rec from the vector, includes firing
	 * a changeevent.
	 */

	public boolean isChartOnlySheet();

	public abstract void removeRecFromVec( BiffRec rec );

	/**
	 * Called from removeCell(), removeMulrk() handles the fact that you
	 * are trying to delete a rk that is really just a part of a Mulrk.  This
	 * is handled by truncating the mulrk at the cell, then creating individual numbers
	 * after the deleted cell.
	 */
	public abstract void removeMulrk( Rk thisrk );

	/**
	 * get a handle to a specific column of cells in this sheet
	 */
	public abstract Colinfo getColInfo( int colin );

	/**
	 * remove all Sheet records from Sheet.
	 */
	public abstract void removeAllRecords();

	/**
	 * set the Bof record for this SheetImpl
	 */
	public abstract void setBOF( Bof b );

	public abstract void setEOF( Eof f );

	/**
	 * return the pos of the Bof for this Sheet
	 */
	public abstract long getLbPlyPos();

	/**
	 * set the pos of the Bof for this Sheet
	 */
	public abstract void setLbPlyPos( long newpos );

	/**
	 * the beginning of the Dimensions record
	 * is the index of the RowBlocks
	 */
	public abstract Dimensions getDimensions();

	public abstract void setDimensions( Dimensions d );

	public abstract int getMinRow();

	public abstract int getMaxRow();

	public abstract int getMinCol();

	public abstract int getMaxCol();

	/**
	 * update the INDEX record with the new max Row #
	 * why we need so many redundant references to the Min/Max Row/Cols
	 * is a question for the Redmond sages.
	 */
	public abstract void updateDimensions( int row, int c );

	/**
	 * set the associated sheet index
	 */
	public abstract void setSheetIDX( Index idx );

	/**
	 * set the numeric sheet number
	 */
	public abstract int getSheetNum();

	/** add a row to the worksheet as well
	 as to the RowBlock which will handle
	 the updating of Dbcell index behavior


	 @param BiffRec the cell being added (can't add a row without one...)

	 public abstract Row addNewRow(BiffRec cell);

	 **/

	/**
	 inserts a row and shifts all of the other rows down by the number
	 of rows specified
	 */
	//public abstract void insertRow(int rownum, int numrows);
	/**
	 inserts a row and shifts all of the other rows down one
	 */
//	public abstract Row insertRow(int rownum, boolean shiftrows);

	/**
	 * Shifts a row down
	 *
	 * @param shiftamount - number of rows shifted
	 */
	//public abstract void shiftRow(Row roe, int shiftamount)
	//	throws RowNotFoundException;

	//	deal with shifting Merged cells. 10-15-04 -jm
	public abstract void updateMergedCells();
	//public abstract void shiftCellRow(BiffRec c, int origRownum, int shiftamount, int flag);
	/* we need to handle shifting any cell references
	 * ie: moving formula C12=SUM(C10:C11) to C13
	 *	should update the formula to C13=SUM(C11:C12)
	 *
	 *	we only do this if the formula references are 'unlocked'
	 *	
	 *	additional restrictions:
	 *		Invoice!F19:F19 - NO, single BiffRec Range
	 *		Sheet2!C3 - NO, not on the same sheet as shifting Cells
	 *		Invoice!F20:F21 - YES, a real range
	 *		Invoice!F20 - YES, a real single ref
	 *
	 */
/* NOTE:  MERGED INTO SHEET.SHIFTCELLROW AS ONE CANNOT CHANGE CELL ROW WITHOUT ALL THE CHECKING MACHINERY CONTAINED WITHIN
	public abstract void setCellRow(BiffRec c, String oldaddr, int newrow);
*/

	/**
	 * get whether this sheet is hidden upon opening
	 */
	public abstract boolean getHidden();

	/**
	 * set whether this sheet is hidden upon opening
	 */
	public abstract void setHidden( int gr );

	/**
	 * returns the selected sheet status
	 */
	public abstract boolean selected();

	/**
	 * set whether this sheet is selected upon opening
	 */
	public abstract void setSelected( boolean b );

	/**
	 * get the number of defined rows on this sheet
	 */
	public abstract int getNumRows();

	/**
	 * get the number of defined cells on this sheet
	 */
	public abstract int getNumCells();

	/**
	 * get the List of columns defined on this sheet
	 */
	public abstract List getColNames();

	/**
	 * get the Number of columns defined on this sheet
	 */
	public abstract int getNumCols();

	/**
	 * get a handle to the Row at the specified
	 * row index
	 */
	public abstract Row getRowByNumber( int r );

	/**
	 * get the List of rows defined on this sheet
	 */
	public abstract List getRowNums();

	/**
	 * return an Array of the Rows in no particular order
	 */
	public abstract Row[] getRows();
	/** Add a Value record to a WorkSheet.


	 @param obj the value of the new Cell
	 @param row & col address of the new Cell

	 */
	//public abstract BiffRec add(Object obj, int[] rc);

	/**
	 * Add a Value record to a WorkSheet.
	 *
	 * @param obj     the value of the new Cell
	 * @param address the address of the new Cell
	 */
	public abstract BiffRec addValue( Object obj, String address );

	/**
	 * Add an BiffRec to a WorkSheet
	 */
	public abstract void addRecord( BiffRec rec, int[] rc );

	public abstract void setCopyPriorCellFormats( boolean f );
//	public abstract boolean copyPriorCellFormatForNewCells(BiffRec c);
	/** add a new cell to the book
	 */
	//public abstract void addNewCell(BiffRec cell);

	/**
	 * add a cell to the worksheet cell collection
	 * <p/>
	 * associates with Row and Col
	 */
	public abstract void addCell( CellRec cell );
	/*
		  Returns the *real* last col num.  Unfortunately the dimensions record
		  cannot be counted on to give a correct value.
	  **/
	//public abstract int getRealMaxCol();

	/**
	 * column formatting records
	 */
	public abstract void addColinfo( Colinfo c );

	/**
	 * get  a colinfo by first col in range
	 */
	public abstract Colinfo getColinfo( String c );

	/**
	 * get the Collection of Colinfos
	 */
	public abstract Collection<Colinfo> getColinfos();
	/** add a new rec to the book where rec already has a cell
	 or does not need one.

	 three things need to happen
	 1. get the record index in the recvec
	 2. get the record offset in the workbook stream
	 3. put the record in the output vector
	 *
	 **/
	//public abstract void addNewRecord(BiffRec vr);
	/**
	 Moves a cell location from one address to another
	 */

	/**
	 * get a cell by address from the worksheet
	 */
	public abstract BiffRec getCell( String s );

	/**
	 * Returns a BiffRec for working with
	 * the value of a BiffRec on a WorkSheet.
	 *
	 * @param int Row the integer row of the Cell
	 * @param int Col the integer col of the Cell
	 * @throws CellNotFoundException is thrown if there is no existing BiffRec at the specified address.
	 */
	public abstract BiffRec getCell( int row, int col ) throws CellNotFoundException;
	/** get an array of all cells for this worksheet

	 public abstract Hashtable getCellMap();*/
	/**
	 * get an array of all cells for this worksheet
	 */
	public abstract BiffRec[] getCells();

	public abstract void addMergedCellsRec( Mergedcells r );

	public abstract List getMergedCellsRecs();

	public abstract Mergedcells getMergedCellsRec();

	/**
	 * get the name of the sheet
	 */
	public abstract String getSheetName();

	/**
	 * get the name of the sheet
	 */
	public abstract String toString();

	/**
	 * initialize the SheetImpl with data from
	 * the byte array.
	 */
	@Override
	public abstract void init();

	public byte getGrbitChr();

	/**
	 * @param b
	 */
	public void setGrbitChr( byte b );

	/**
	 * change the displayed name of the sheet
	 * <p/>
	 * Affects the following byte values:
	 * 10      cch         1       Length of sheet name
	 * 11      grbitChr    1       Compressed/Uncompressed Unicode
	 * 12      rgch        var     Sheet name
	 */
	public abstract void setSheetName( String newname );

	/*
	    Returns a serialized copy of this SheetImpl
	*/
	public abstract byte[] getSheetBytes() throws IOException;

	/* prior to serializing the worksheet,
	   we need to initialize the records which belong to this sheet
	   instance.
	
	*/
	public abstract void setLocalRecs();

	/**
	 * get the type of sheet as a short
	 */
	public abstract short getSheetType();

	/**
	 * get the type of sheet as a string
	 */
	public abstract String getSheetTypeString();

	/* Inserts a serialized boundsheet into the workboook
	*/
	public abstract Chart addChart( byte[] inbytes, String NewChartName, short[] coords );

	public abstract Guts getGuts();

	public abstract void setGuts( Guts g );

	public abstract void setWsBool( WsBool ws );

	public abstract WsBool getWsBool();
	//	----- TODO: implement COMBINATOR SECTION ------ //
	//public abstract void initCombinators(SheetImpl sh);
}