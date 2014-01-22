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

import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.CellPositionConflictException;
import com.extentech.formats.XLS.CellTypeMismatchException;
import com.extentech.formats.XLS.ColumnNotFoundException;
import com.extentech.formats.XLS.FormulaNotFoundException;
import com.extentech.formats.XLS.Mergedcells;
import com.extentech.formats.XLS.RowNotFoundException;
import com.extentech.formats.XLS.Sheet;

import java.util.List;

/**
 * An interface representing an ExtenXLS WorkSheet
 *
 * @see WorkSheetHandle
 */
public interface WorkSheet extends Handle
{

	public WorkBook getWorkBook();

	public Mergedcells getMergedCellsRec();

	public abstract void addChart( byte[] serialchart, String name );

	/**
	 * Get the first row on the Worksheet
	 *
	 * @return the Minimum Row Number on the Worksheet
	 */
	public abstract int getFirstRow();

	/**
	 * Get the first column on the Worksheet
	 *
	 * @return the Minimum Column Number on the Worksheet
	 */
	public abstract int getFirstCol();

	/**
	 * Get the last row on the Worksheet
	 *
	 * @return the Maximum Row Number on the Worksheet
	 */
	public abstract int getLastRow();

	/**
	 * Get the last column on the Worksheet
	 *
	 * @return the Maximum Column Number on the Worksheet
	 */
	public abstract int getLastCol();

	/* set protection on sheet
	
	    @param boolean whether to protect the sheet
	*/
	public abstract void setProtected( boolean b );

	/**
	 * set whether this sheet is VERY hidden opening the file.
	 * <p/>
	 * VERY hidden means users will not be able to unhide
	 * the sheet without using VB code.
	 *
	 * @param boolean b hidden state
	 */
	public abstract void setVeryHidden( boolean b );

	/**
	 * get whether this sheet is selected upon opening the file.
	 *
	 * @return boolean b selected state
	 */
	public abstract boolean getSelected();

	/**
	 * get whether this sheet is hidden from the user opening the file.
	 *
	 * @return boolean b hidden state
	 */
	public abstract boolean getHidden();

	/**
	 * set whether this sheet is hidden from the user opening the file.
	 * <p/>
	 * if the sheet is selected, the API will set the first
	 * visible sheet to selected as you cannot have your selected sheet
	 * be hidden.
	 * <p/>
	 * to override this behavior, set your desired sheet to selected
	 * after calling this method.
	 *
	 * @param boolean b hidden state
	 */
	public abstract void setHidden( boolean b );

	/**
	 * set this WorkSheet as the first visible tab on the left
	 */
	public abstract void setFirstVisibleTab();

	/**
	 * get the tab display order of this Worksheet
	 * <p/>
	 * this is a zero based index with zero representing
	 * the left-most WorkSheet tab.
	 *
	 * @return int idx the index of the sheet tab
	 */
	public abstract int getTabIndex();

	/**
	 * set the tab display order of this Worksheet
	 * <p/>
	 * this is a zero based index with zero representing
	 * the left-most WorkSheet tab.
	 *
	 * @param int idx the new index of the sheet tab
	 */
	public abstract void setTabIndex( int idx );

	/**
	 * set whether this sheet is selected upon
	 * opening the file.
	 *
	 * @param boolean b selected value
	 */
	public abstract void setSelected( boolean b );

	/**
	 * returns the Column at the index position
	 *
	 * @return ColHandle the Column
	 */
	public abstract ColHandle getCol( int clnum ) throws ColumnNotFoundException;

	/**
	 * returns the Column at the named position
	 *
	 * @return ColHandle the Column
	 */
	public abstract ColHandle getCol( String name ) throws ColumnNotFoundException;

	/**
	 * returns all of the Columns in this WorkSheet
	 *
	 * @return ColHandle[] Columns
	 */
	public abstract ColHandle[] getColumns();

	/**
	 * returns a Vector of Column names
	 *
	 * @return CompatibleVector column names
	 */
	public abstract List getColNames();

	/**
	 * returns a Vector of Row numbers
	 *
	 * @return CompatibleVector row numbers
	 */
	public abstract List getRowNums();

	/**
	 * get an a RowHandle for
	 * this WorkSheet by number
	 *
	 * @param int row number to return
	 * @return RowHandle a Row on this WorkSheet
	 */
	public abstract RowHandle getRow( int t ) throws RowNotFoundException;

	/**
	 * get an array of all RowHandles for
	 * this WorkSheet
	 *
	 * @return RowHandle[] all Rows on this WorkSheet
	 */
	public abstract RowHandle[] getRows();

	/**
	 * Returns the number of rows in this WorkSheet
	 *
	 * @return int Number of Rows on this WorkSheet
	 */
	public abstract int getNumRows();

	/**
	 * Returns the number of Columns in this WorkSheet
	 *
	 * @return int Number of Cols on this WorkSheet
	 */
	public abstract int getNumCols();

	/**
	 * Remove a Cell from this WorkSheet.
	 *
	 * @param CellHandle to remove
	 */
	public abstract void removeCell( CellHandle celldel );

	/**
	 * Remove a Cell from this WorkSheet.
	 *
	 * @param String celladdr - the Address of the Cell to remove
	 */
	public abstract void removeCell( String celladdr );

	/**
	 * Remove a Row and all associated Cells from
	 * this WorkSheet.
	 *
	 * @param int rownum - the number of the row to remove
	 */
	public abstract void removeRow( int rownum ) throws RowNotFoundException;

	/**
	 * Remove a Row and all associated Cells from
	 * this WorkSheet. Optionally shift all rows below
	 * target row up one.
	 *
	 * @param int     rownum - the number of the row to remove
	 * @param boolean shiftrows - true will shift all lower rows up one.
	 */
	public abstract void removeRow( int rownum, boolean shiftrows ) throws RowNotFoundException;

	/**
	 * Remove a Column and all associated Cells from
	 * this WorkSheet.
	 *
	 * @param String colstr - the name of the column to remove
	 */
	public abstract void removeCol( String colstr ) throws ColumnNotFoundException;

	/**
	 * Remove a Column and all associated Cells from
	 * this WorkSheet.
	 * <p/>
	 * Optionally shift all cols to the right
	 * of this column to the left by one.
	 *
	 * @param String  colstr - the name of the column to remove
	 * @param boolean shiftcols - true will shift following cols
	 */
	public abstract void removeCol( String colstr, boolean shiftcols ) throws ColumnNotFoundException;

	/**
	 * Returns the low-level ExtenXLS Sheet.
	 *
	 * @return com.extentech.formats.XLS.Sheet
	 */
	public Sheet getSheet();

	/**
	 * Returns the name of the Sheet.
	 *
	 * @return String Sheet Name
	 */
	public abstract String getSheetName();

	/**
	 * Returns the Serialized bytes for this WorkSheet.
	 * <p/>
	 * The output of this method can be used to insert
	 * a copy of this WorkSheet into another WorkBook
	 * using the WorkBook.addWorkSheet(byte[] serialsheet, String NewSheetName)
	 * method.
	 *
	 * @return byte[] the WorkSheet's Serialized bytes
	 * @see WorkBook.addWorkSheet(byte[] serialsheet, String NewSheetName)
	 */
	public abstract byte[] getSerialBytes();

	/**
	 * Set the Object value of the Cell at the given address.
	 *
	 * @param String Cell Address (ie: "D14")
	 * @param Object new Cell Object value
	 * @throws CellNotFoundException is thrown if there is
	 *                               no existing Cell at the specified address.
	 */
	public abstract void setVal( String address, Object val ) throws CellNotFoundException, CellTypeMismatchException;

	/**
	 * Set the double value of the Cell at the given address
	 *
	 * @param String  Cell Address (ie: "D14")
	 * @param double  new Cell double value
	 * @param address
	 * @param d
	 * @throws com.extentech.formats.XLS.CellNotFoundException is
	 *                                                         thrown if there is no existing Cell at the specified address.
	 */
	public abstract void setVal( String address, double d ) throws CellNotFoundException, CellTypeMismatchException;

	/**
	 * Set the String value of the Cell at the given address
	 *
	 * @param String Cell Address (ie: "D14")
	 * @param String new Cell String value
	 * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public abstract void setVal( String address, String s ) throws CellNotFoundException, CellTypeMismatchException;

	/**
	 * Set the name of the Worksheet.  This method will change
	 * the name on the Worksheet's tab as displayed in the
	 * WorkBook, as well as all programmatic and internal
	 * references to the name.
	 * <p/>
	 * This change takes effect immediately, so all attempts
	 * to reference the Worksheet by its previous name will
	 * fail.
	 *
	 * @param String the new name for the Worksheet
	 */
	public abstract void setSheetName( String name );

	/**
	 * Set the int value of the Cell at the given address
	 *
	 * @param String Cell Address (ie: "D14")
	 * @param int    new Cell int value
	 * @throws com.extentech.formats.XLS.CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public abstract void setVal( String address, int i ) throws CellNotFoundException, CellTypeMismatchException;

	/**
	 * Get the Object value of a Cell.
	 * <p/>
	 * Numeric Cell values will return as type Long, Integer, or Double.
	 * String Cell values will return as type String.
	 *
	 * @return the value of the Cell as an Object.
	 * @throws CellNotFoundException is thrown if there is
	 *                               no existing Cell at the specified address.
	 */
	public abstract Object getVal( String address ) throws CellNotFoundException;

	/**
	 * Insert a blank row into the worksheet.
	 * Shift all rows below the cell down one.
	 * <p/>
	 * Adding new cells to non-existent rows will
	 * automatically create new rows in the file,
	 * This method is only necessary to "move" existing cells
	 * by inserting empty rows.
	 *
	 * @param rownum the rownumber to insert
	 */
	public abstract void insertRow( int rownum );

	/**
	 * Insert a blank row into the worksheet.
	 * Shift all rows below the cell down one.
	 * <p/>
	 * Adding new cells to non-existent rows will
	 * automatically create new rows in the file,
	 * <p/>
	 * After calling this method, setVal() can be used on the
	 * newly created cells to update with new values.
	 *
	 * @param rownum  the rownumber to insert
	 * @param whether to shift down existing Cells
	 */
	public abstract void insertRow( int rownum, boolean shiftrows );

	public static final int ROW_KEEP = 0;

	public static final int ROW_INSERT = 1;

	/**
	 * Insert a blank row into the worksheet.
	 * Shift all rows below the cell down one.
	 * <p/>
	 * Adding new cells to non-existent rows will
	 * automatically create new rows in the file,
	 * <p/>
	 * After calling this method, setVal() can be used on the
	 * newly created cells to update with new values.
	 *
	 * @param rownum  the rownumber to insert
	 * @param whether to shift down existing Cells
	 */
	public abstract void insertRow( int rownum, int copyrow, int flag, boolean shiftrows );

	/**
	 * Insert a blank column into the worksheet.
	 * Shift all columns to the right of the cell over one.
	 * <p/>
	 * Adding new cells to non-existent columns will
	 * automatically create new Columns in the file,
	 * This method is only necessary to "move" existing cells
	 * by inserting empty columns.
	 *
	 * @param colstr the Column string to insert
	 */
	public abstract void insertCol( String colnum );

	/**
	 * When adding a new Cell to the sheet, ExtenXLS can
	 * automatically copy the formatting from the Cell directly
	 * above the inserted Cell.
	 * <p/>
	 * ie: if set to true, newly added Cell D19 would take its formatting
	 * from Cell D18.
	 * <p/>
	 * Default is false
	 * <p/>
	 * <p/>
	 * <p/>
	 * boolean copy the formats from the prior Cell
	 */
	public abstract void setCopyFormatsFromPriorWhenAdding( boolean f );

	/**
	 * Add a Cell with the specified value to a WorkSheet.
	 * <p/>
	 * This method determines the Cell type based on
	 * type-compatibility of the value.
	 * <p/>
	 * In other words, if the Object cannot be converted
	 * safely to a Numeric Object type, then it is treated
	 * as a String and a new String value is added to the
	 * WorkSheet at the Cell address specified.
	 *
	 * @param obj the value of the new Cell
	 * @param int row the row of the new Cell
	 * @param int col the column of the new Cell
	 */
	public abstract CellHandle add( Object obj, int row, int col );

	/**
	 * Add a Cell with the specified value to a WorkSheet.
	 * <p/>
	 * This method determines the Cell type based on
	 * type-compatibility of the value.
	 * <p/>
	 * In other words, if the Object cannot be converted
	 * safely to a Numeric Object type, then it is treated
	 * as a String and a new String value is added to the
	 * WorkSheet at the Cell address specified.
	 *
	 * @param obj     the value of the new Cell
	 * @param address the address of the new Cell
	 */
	public abstract CellHandle add( Object obj, String address );

	/**
	 * Add a java.sql.Date Cell to a WorkSheet.
	 * <p/>
	 * You must specify a formatting pattern for the
	 * new date, or null for the default ("m/d/yy h:mm".)
	 * <p/>
	 * valid date format patterns
	 * "m/d/y"
	 * "d-mmm-yy"
	 * "d-mmm"
	 * "mmm-yy"
	 * "h:mm AM/PM"
	 * "h:mm:ss AM/PM"
	 * "h:mm"
	 * "h:mm:ss"
	 * "m/d/yy h:mm"
	 * "mm:ss"
	 * "[h]:mm:ss"
	 * "mm:ss.0"
	 *
	 * @param dt         the value of the new java.sql.Date Cell
	 * @param address    the address of the new java.sql.Date Cell
	 * @param formatting pattern the address of the new java.sql.Date Cell
	 */
	public abstract CellHandle add( java.util.Date dt, String address, String fmt );

	/**
	 * Remove this WorkSheet from the WorkBook
	 */
	public abstract void remove();

	/**
	 * Returns all CellHandles defined on this WorkSheet.
	 *
	 * @return Cell[] - the array of Cells in the Sheet
	 */
	public abstract CellHandle[] getCells();

	/**
	 * Returns a FormulaHandle for working with
	 * the ranges of a formula on a WorkSheet.
	 *
	 * @param addr the address of the Cell
	 * @throws FormulaNotFoundException is thrown if there is no existing formula at the specified address.
	 */
	public abstract FormulaHandle getFormula( String addr ) throws FormulaNotFoundException, CellNotFoundException;

	/**
	 * Returns a CellHandle for working with
	 * the value of a Cell on a WorkSheet.
	 *
	 * @param addr the address of the Cell
	 * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public abstract CellHandle getCell( String addr ) throws CellNotFoundException;

	/**
	 * Returns a CellHandle for working with
	 * the value of a Cell on a WorkSheet.
	 *
	 * @param int Row the integer row of the Cell
	 * @param int Col the integer col of the Cell
	 * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public abstract CellHandle getCell( int row, int col ) throws CellNotFoundException;

	/**
	 * Move a cell on this WorkSheet.
	 *
	 * @param CellHandle c - the cell to be moved
	 * @param String     celladdr - the destination address of the cell
	 */
	public abstract void moveCell( CellHandle c, String addr ) throws CellPositionConflictException;

	/**
	 * Get the text for the Footer printed at the bottom
	 * of the Worksheet
	 *
	 * @return String footer text
	 */
	public abstract String getFooterText();

	/**
	 * Get the text for the Header printed at the top
	 * of the Worksheet
	 *
	 * @return String header text
	 */
	public abstract String getHeaderText();

	/**
	 * Set the text for the Header printed at the top
	 * of the Worksheet
	 *
	 * @param String header text
	 */
	public abstract void setHeaderText( String t );

	/**
	 * Set the text for the Footer printed at the bottom
	 * of the Worksheet
	 *
	 * @param String footer text
	 */
	public abstract void setFooterText( String t );

	/**
	 * Returns the name of this Sheet.
	 *
	 * @see java.lang.Object#toString()
	 */
	public abstract String toString();
}