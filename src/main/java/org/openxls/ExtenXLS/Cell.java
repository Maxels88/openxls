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

/**
 * A lightweight subset of Cell methods allowing for low memory overhead streaming implementations
 * <p/>
 * <p/>
 * <p/>
 * Feature Request:
 * <p/>
 * methods determining the type of the cell on the new Cell class (like CellHandle.isDate() or getCellType()).
 * The current API has a method returning Object, but from I remember at least the date should be explicitly checked with isDate() method.
 * Basically the API should allow us to do what we do now in method ExcelExtenReader.getCellData(CellHandle).
 * Please review the method and let me know what are the plans for the future.
 */
public interface Cell
{

	/**
	 * Cell types
	 */
	public static final int TYPE_BLANK = -1;
	public static final int TYPE_STRING = 0;
	public static final int TYPE_FP = 1;
	public static final int TYPE_INT = 2;
	public static final int TYPE_FORMULA = 3;
	public static final int TYPE_BOOLEAN = 4;
	public static final int TYPE_DOUBLE = 5;

	public boolean isDate();

	public int getCellType();

	/**
	 * Returns the Formatting record ID (FormatId) for this Cell
	 * <br>
	 * This can be used with 'setFormatId(int i)' to copy
	 * the formatting from one Cell to another (e.g.
	 * a template cell to a new cell)
	 *
	 * @return int the FormatId for this Cell
	 */
	public int getFormatId();

	/**
	 * Returns the value of this Cell in the native underlying data type.
	 * <p/>
	 * Formula cells will return the calculated value of the formula in the
	 * calculated data type.
	 * <p/>
	 * Use 'getStringVal()' to return a String regardless of underlying value type.
	 *
	 * @return Object value for this Cell
	 */
	public Object getVal();

	/**
	 * Returns the value of the Cell as a String with formatting
	 * pattern applied..
	 * <br>
	 * see:
	 * <a href="tutorial">http://java.sun.com/docs/books/tutorial/i18n/format/decimalFormat.html</a>
	 * <br>
	 * boolean Cell types will return "true" or "false"
	 * <br>
	 * Negative numbers that are formatted in excel to show as red values rather than using a "-" will return with a minus symbol.
	 *
	 * @return String the formatted value of the Cell
	 */
	public String getFormattedStringVal();

	/**
	 * Returns the column number of this Cell.
	 *
	 * @return int the Column Number of the Cell
	 */
	public int getColNum();

	/**
	 * Returns the row number of this Cell.
	 * <p/>
	 * NOTE: This is the 1-based row number such as you will see in a spreadsheet UI.
	 * <p/>
	 * ie: A1 = row 1
	 *
	 * @return int the ONE-based Row Number of the Cell
	 */
	public int getRowNum();

	/**
	 * Returns the Address of this Cell as a String.
	 *
	 * @return String the address of this Cell in the WorkSheet
	 */
	public String getCellAddress();

	/**
	 * Returns the name of this Cell's WorkSheet as a String.
	 *
	 * @return String the name this Cell's WorkSheet
	 */
	public String getWorkSheetName();

}
