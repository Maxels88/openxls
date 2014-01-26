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
 * Constants used in the generation of JSON output.
 */
interface JSONConstants
{
	public static final String JSON_CELL_VALUE = "v";
	public static final String JSON_CELL_FORMATTED_VALUE = "fv";
	public static final String JSON_CELL_FORMULA = "fm";
	public static final String JSON_CELL = "Cell";
	public static final String JSON_CELLS = "cs";
	public static final String JSON_DATETIME = "DateTime";
	public static final String JSON_DATEVALUE = "DateValue";
	public static final String JSON_DOUBLE = "Double";
	public static final String JSON_RANGE = "Range";
	public static final String JSON_DATA = "d";
	public static final String JSON_FLOAT = "Float";
	public static final String JSON_INTEGER = "Integer";
	public static final String JSON_LOCATION = "loc";
	public static final String JSON_ROW = "Row";
	public static final String JSON_ROW_BORDER_TOP = "BdrT";
	public static final String JSON_ROW_BORDER_BOTTOM = "BdrB";
	public static final String JSON_HEIGHT = "h";
	public static final String JSON_STRING = "String";
	public static final String JSON_STYLEID = "sid";
	public static final String JSON_TYPE = "t";
	public static final String JSON_FORMULA_HIDDEN = "fhd";
	public static final String JSON_LOCKED = "lck";
	public static final String JSON_HIDDEN = "Hidden";
	public static final String JSON_VALIDATION_MESSAGE = "vm";
	public static final String JSON_MERGEACROSS = "MergeAcross";
	public static final String JSON_MERGEDOWN = "MergeDown";
	public static final String JSON_MERGEPARENT = "MergeParent";
	public static final String JSON_MERGECHILD = "MergeChild";
	public static final String JSON_HREF = "HRef";
	public static final String JSON_WORD_WRAP = "wrap";
	public static final String JSON_RED_FORMAT = "negRed";
	public static final String JSON_TEXT_ALIGN = "txtAlign";
}
