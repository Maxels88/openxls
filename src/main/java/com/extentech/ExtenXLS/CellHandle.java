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

import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.Blank;
import com.extentech.formats.XLS.Boolerr;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.CellPositionConflictException;
import com.extentech.formats.XLS.CellTypeMismatchException;
import com.extentech.formats.XLS.Cf;
import com.extentech.formats.XLS.ColumnNotFoundException;
import com.extentech.formats.XLS.Condfmt;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.FormatConstantsImpl;
import com.extentech.formats.XLS.Formula;
import com.extentech.formats.XLS.FormulaNotFoundException;
import com.extentech.formats.XLS.FunctionNotSupportedException;
import com.extentech.formats.XLS.Hlink;
import com.extentech.formats.XLS.Labelsst;
import com.extentech.formats.XLS.Mulblank;
import com.extentech.formats.XLS.Note;
import com.extentech.formats.XLS.NumberRec;
import com.extentech.formats.XLS.ReferenceTracker;
import com.extentech.formats.XLS.Rk;
import com.extentech.formats.XLS.Unicodestring;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.formats.XLS.Xf;
import com.extentech.formats.XLS.charts.Ai;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.formats.cellformat.CellFormatFactory;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;
import org.json.JSONException;
import org.json.JSONObject;

import java.awt.*;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Stack;

import static com.extentech.ExtenXLS.JSONConstants.JSON_CELL_FORMATTED_VALUE;
import static com.extentech.ExtenXLS.JSONConstants.JSON_CELL_FORMULA;
import static com.extentech.ExtenXLS.JSONConstants.JSON_CELL_VALUE;
import static com.extentech.ExtenXLS.JSONConstants.JSON_DATA;
import static com.extentech.ExtenXLS.JSONConstants.JSON_DATETIME;
import static com.extentech.ExtenXLS.JSONConstants.JSON_DOUBLE;
import static com.extentech.ExtenXLS.JSONConstants.JSON_FLOAT;
import static com.extentech.ExtenXLS.JSONConstants.JSON_FORMULA_HIDDEN;
import static com.extentech.ExtenXLS.JSONConstants.JSON_HIDDEN;
import static com.extentech.ExtenXLS.JSONConstants.JSON_HREF;
import static com.extentech.ExtenXLS.JSONConstants.JSON_INTEGER;
import static com.extentech.ExtenXLS.JSONConstants.JSON_LOCATION;
import static com.extentech.ExtenXLS.JSONConstants.JSON_LOCKED;
import static com.extentech.ExtenXLS.JSONConstants.JSON_MERGEACROSS;
import static com.extentech.ExtenXLS.JSONConstants.JSON_MERGECHILD;
import static com.extentech.ExtenXLS.JSONConstants.JSON_MERGEDOWN;
import static com.extentech.ExtenXLS.JSONConstants.JSON_MERGEPARENT;
import static com.extentech.ExtenXLS.JSONConstants.JSON_RED_FORMAT;
import static com.extentech.ExtenXLS.JSONConstants.JSON_STRING;
import static com.extentech.ExtenXLS.JSONConstants.JSON_STYLEID;
import static com.extentech.ExtenXLS.JSONConstants.JSON_TEXT_ALIGN;
import static com.extentech.ExtenXLS.JSONConstants.JSON_TYPE;
import static com.extentech.ExtenXLS.JSONConstants.JSON_VALIDATION_MESSAGE;
import static com.extentech.ExtenXLS.JSONConstants.JSON_WORD_WRAP;

/**
 * The CellHandle provides a handle to an XLS Cell and its values. <br>
 * Use the CellHandle to work with individual Cells in an XLS file. <br>
 * To instantiate a CellHandle, you must first have a valid WorkSheetHandle,
 * which in turn requires a valid WorkBookHandle. <br>
 * <br>
 * for example: <br>
 * <br>
 * <blockquote> WorkBookHandle book = newWorkBookHandle("testxls.xls");<br>
 * WorkSheetHandle sheet1 = book.getWorkSheet("Sheet1");<br>
 * CellHandlecell = sheet1.getCell("B22");<br>
 * <br>
 * <br>
 * </blockquote> With a CellHandle you can:<br>
 * <br>
 * - get the value of a Cell<br>
 * - set the value of a Cell<br>
 * - change the color, font, and background formatting of a Cell<br>
 * - change the formatting pattern of a Cell<br>
 * - change the value of a Cell<br>
 * - get a handle to any Formula for this Cell<br>
 * <br>
 * <p/>
 * "http://www.extentech.com">Extentech Inc.</a>
 *
 * @see WorkBookHandle
 * @see WorkSheetHandle
 * @see FormulaHandle
 * @see CellNotFoundException
 * @see CellTypeMismatchException
 */
public class CellHandle implements Cell, Serializable, Handle, Comparable<CellHandle>
{

	/**
	 *
	 *
	 */
	private static final long serialVersionUID = 4737120893891570607L;
	/**
	 * Cell types
	 */
	public static final int TYPE_BLANK = Cell.TYPE_BLANK;
	public static final int TYPE_STRING = Cell.TYPE_STRING;
	public static final int TYPE_FP = Cell.TYPE_FP;
	public static final int TYPE_INT = Cell.TYPE_INT;
	public static final int TYPE_FORMULA = Cell.TYPE_FORMULA;
	public static final int TYPE_BOOLEAN = Cell.TYPE_BOOLEAN;
	public static final int TYPE_DOUBLE = Cell.TYPE_DOUBLE;
	public static final int NOTATION_STANDARD = 0, NOTATION_SCIENTIFIC = 1, NOTATION_SCIENTIFIC_EXCEL = 2;

	private transient WorkBook wbh = null;
	private transient WorkSheetHandle wsh = null;
	private ColHandle mycol;
	private RowHandle myrow;
	private FormatHandle formatter;
	// reusing or creating new xfs is handled in FormatHandle/cloneXf and
	// updateXf
	// boolean useExistingXF = !false;
	private boolean DEBUG = false;
	private XLSRecord mycell;

	/**
	 * returns the underlying BIFF8 record for the Cell <br>
	 * NOTE: the underlying record is not a part of the public API and may
	 * change at any time.
	 *
	 * @return Returns the underlying biff record.
	 */
	public XLSRecord getRecord()
	{
		return mycell;
	}

	/**
	 * sets the underlying BIFF8 record for the Cell <br>
	 * NOTE: the underlying record is not a part of the public API and may
	 * change at any time.
	 *
	 * @param XLSRecord rec - The BIFF record to set.
	 */
	public void setRecord( XLSRecord rec )
	{
		this.mycell = rec;
	}

	/**
	 * Public Constructor added for use in Bean-context ONLY.
	 * <p/>
	 * Do NOT create CellHandles manually.
	 */
	public CellHandle( BiffRec c )
	{
		mycell = (XLSRecord) c;
	}

	/**
	 * if this cellhandle refers to a mulblank, ensure internal mulblank
	 * properties are set to appropriate cell in the mulblank range
	 */
	private void setMulblank()
	{
		if( mycell.getOpcode() == XLSConstants.MULBLANK )
		{
			if( mulblankcolnum == -1 )
			{ // init
				mulblankcolnum = (short) ((Mulblank) mycell).getColNumber();
				((Mulblank) mycell).setCurrentCell( mulblankcolnum );
				((Mulblank) mycell).getIxfe(); // ensure myxf is set to correct
				// xf for the given cell in the
				// set of mulblanks
			}
			else if( mulblankcolnum != mycell.getColNumber() )
			{
				((Mulblank) mycell).setCurrentCell( mulblankcolnum );
				((Mulblank) mycell).getIxfe(); // ensure myxf is set to correct
				// xf for the given cell in the
				// set of mulblanks
			}
			this.formatter = null;
		}
	}

	/**
	 * Public Constructor added for use in Bean-context ONLY.
	 * <p/>
	 * Do NOT create CellHandles manually.
	 */
	public CellHandle( BiffRec c, WorkBook myb )
	{
		mycell = (XLSRecord) c;
		setMulblank();
		this.wbh = myb;
	}

	/**
	 * Get a FormatHandle (a Format Object describing the formats for this Cell)
	 * referenced by this CellHandle.
	 *
	 * @return FormatHandle
	 * @see FormatHandle
	 */
	void setFormatHandle()
	{
		setMulblank();
		if( (formatter != null) && (formatter.getFormatId() == this.mycell.getIxfe()) )
		{
			return;
		}
		// reusing or creating new xfs is handled in FormatHandle/cloneXf and
		// updateXf
		if( this.mycell.getXfRec() != null )
		{
			formatter = new FormatHandle( this.wbh, this.mycell.myxf );
		}
		else
		{// should ever happen now?
			// useExistingXF = false;
			if( (wbh == null) && (this.mycell.getWorkBook() != null) )
			{
				formatter = new FormatHandle( this.mycell.getWorkBook(), -1 );
			}
			else
			{
				formatter = new FormatHandle( this.wbh, -1 );
			}
		}
		formatter.addCell( mycell );
	}

	/**
	 * Sets a default "empty" value appropriate for the cell type of this
	 * CellHandle <br>
	 * For example, will set the value to 0.0 for TYPE_DOUBLE, an empty String
	 * for TYPE_BLANK
	 */
	public void setToDefault()
	{
		setVal( this.getDefaultVal() );
	}

	/**
	 * Get the default "empty" data value for this CellHandle
	 *
	 * @return Object a default empty value corresponding to the cell type <br>
	 * such as 0.0 for TYPE_DOUBLE or an empty String for TYPE_BLANK
	 */
	public Object getDefaultVal()
	{
		return mycell.getDefaultVal();
	}

	/**
	 * Sets the number format pattern for this cell. All Excel built-in number
	 * formats are supported. Custom formats will not be applied by ExtenXLS
	 * (e.g. the {@link #getFormattedStringVal} method) but they will be written
	 * correctly to the output file. For more information on number format
	 * patterns see <a href="http://support.microsoft.com/kb/264372">Microsoft
	 * KB264372</a>.
	 *
	 * @param pat the Excel number format pattern to apply
	 * @see FormatConstantsImpl#getBuiltinFormats
	 */
	public void setFormatPattern( String pat )
	{
		setFormatHandle();
		formatter.setFormatPattern( pat );
	}

	/**
	 * Gets the number format pattern for this cell, if set. For more
	 * information on number format patterns see <a
	 * href="http://support.microsoft.com/kb/264372">Microsoft KB264372</a>.
	 *
	 * @return the Excel number format pattern for this cell or
	 * <code>null</code> if none is applied
	 */
	public String getFormatPattern()
	{
		if( this.getFont() == null )
		{
			return "";
		}
		return mycell.getFormatPattern();
	}

	/**
	 * Returns whether this Cell has Date formatting applied. <br>
	 * NOTE: This does not guarantee that the value is a valid date.
	 *
	 * @return boolean true if this Cell has a Date Format applied
	 */
	@Override
	public boolean isDate()
	{
		if( mycell.myxf == null )
		{
			return false;
		}
		if( mycell.isString )
		{
			return false;
		}
		if( mycell.isBoolean )
		{
			return false;
		}
		if( mycell.isBlank )
		{
			return false;
		}
		return mycell.myxf.isDatePattern();
	}

	/**
	 * Returns whether this cell is a formula.
	 */
	public boolean isFormula()
	{
		return mycell.isFormula();
	}

	/**
	 * Returns whether the formula for the Cell is hidden
	 *
	 * @return boolean true if formula is hidden
	 */
	public boolean isFormulaHidden()
	{
		return this.getFormatHandle().isFormulaHidden();
	}

	/**
	 * returns whether this Cell is locked for editing
	 *
	 * @return boolean true if the cell is locked
	 */
	public boolean isLocked()
	{
		return this.getFormatHandle().isLocked();
	}

	/**
	 * Returns whether this is a blank cell.
	 *
	 * @return true if this cell is blank
	 */
	public boolean isBlank()
	{
		return this.mycell.isBlank;
	}

	/**
	 * Returns whether this Cell has a Currency format applied.
	 *
	 * @return boolean true if this Cell has a Currency format applied
	 */
	public boolean isCurrency()
	{
		if( mycell.myxf == null )
		{
			return false;
		}
		return mycell.myxf.isCurrencyPattern();
	}

	/**
	 * Returns whether this Cell has a numeric value
	 *
	 * @return boolean true if this Cell contains a numeric value
	 */
	public boolean isNumber()
	{
		return this.mycell.isNumber();
	}

	/**
	 * change the size (in points) of the Font used by this Cell and all other
	 * Cells sharing this FormatId. <br>
	 * NOTE: To add an entirely new Font use the setFont(String fn, int typ, int
	 * sz) method instead.
	 *
	 * @param int sz - Font size in Points.
	 */
	public void setFontSize( int sz )
	{
		setFormatHandle();
		sz *= 20; // excel size is 1/20 pt.
		formatter.setFontHeight( sz );
	}

	/**
	 * change the weight (boldness) of the Font used by this Cell. <br>
	 * Some examples: <br>
	 * a weight of 200 is normal <br>
	 * a weight of 700 is bold <br>
	 *
	 * @param int wt - Font weight range between 100-1000
	 */
	public void setFontWeight( int wt )
	{
		setFormatHandle();
		formatter.setFontWeight( wt );
	}

	/**
	 * Convenience method for toggling the bold state of the Font used by this
	 * Cell.
	 *
	 * @param boolean bold - true if bold
	 */
	public void setBold( boolean bold )
	{
		setFormatHandle();
		formatter.setBold( bold );
	}

	/**
	 * get the weight (boldness) of the Font used by this Cell.
	 *
	 * @return int Font weight range between 100-1000
	 */
	public int getFontWeight()
	{
		if( this.getFont() == null )
		{
			return FormatHandle.DEFAULT_FONT_WEIGHT;
		}
		return mycell.getFont().getFontWeight();
	}

	/**
	 * get the size in points of the Font used by this Cell
	 *
	 * @return int Font size in Points.
	 */
	public int getFontSize()
	{
		if( this.getFont() == null )
		{
			return FormatHandle.DEFAULT_FONT_SIZE;
		}
		return mycell.getFont().getFontHeight() / 20;
	}

	/**
	 * get the Color of the Font used by this Cell. <br>
	 *
	 * @return int Excel color constant for Font color
	 * @see FormatHandle.COLOR constants
	 */
	public Color getFontColor()
	{
		if( this.getFont() == null )
		{
			return FormatHandle.Black;
		}
		int clidx = this.getFont().getColor();

		// handle white on white text issue
		Xf x = mycell.getXfRec();
		int clidb = x.getBackgroundColor();
		if( (clidx == 64) && (clidb == 64) )
		{ // return black
			return FormatHandle.Black;
		}
		// black on black
		if( clidx < this.getWorkBook().getColorTable().length )
		{
			Color mycolr = this.getWorkBook().getColorTable()[clidx];
			return mycolr;
		}
		return Color.black;
	}

	/**
	 * Set the color of the Font used by this Cell.
	 *
	 * @param java .Awt.Color col - color of the font
	 */
	public void setFontColor( Color col )
	{
		setFormatHandle();
		formatter.setFontColor( col );
	}

	/**
	 * set the Foreground Color for this Cell. <br>
	 * <br>
	 * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
	 * Background color is the PATTERN color for all patterns not equal to
	 * PATTERN_SOLID <br>
	 * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
	 * Background Color is 64 (white).
	 *
	 * @param int t - Excel color constant
	 * @see FormatHandle.COLOR constants
	 */
	public void setForegroundColor( int t )
	{
		setFormatHandle();
		formatter.setForegroundColor( t );
	}

	/**
	 * set the Foreground Color for this Cell <br>
	 * NOTE: this is the PATTERN Color
	 *
	 * @param int Excel color constant
	 * @see FormatHandle.COLOR constants
	 */
	// TODO: is this doc correct?
	public void setForeColor( int i )
	{
		if( mycell.myxf == null )
		{
			this.getNewXf();
		}
		mycell.myxf.setForeColor( i, null );
	}

	/**
	 * set the Background Color for this Cell. <br>
	 * <br>
	 * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
	 * Background color is the PATTERN color for all patterns not equal to
	 * PATTERN_SOLID <br>
	 * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
	 * Background Color is 64 (white).
	 *
	 * @param int t - Excel color constant
	 * @see FormatHandle.COLOR constants
	 */
	public void setBackgroundColor( int t )
	{
		setFormatHandle();
		formatter.setBackgroundColor( t );
	}

	/**
	 * get the Color of the Cell Foreground of this Cell and all other cells
	 * sharing this FormatId. <br>
	 * <br>
	 * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
	 * Background color is the PATTERN color for all patterns not equal to
	 * PATTERN_SOLID <br>
	 * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
	 * Background Color is 64 (white).
	 *
	 * @return java.awt.Color of the Font
	 */
	public Color getForegroundColor()
	{
		if( this.mycell.getXfRec() != null )
		{
			int clidx = mycell.getXfRec().getForegroundColor();
			if( clidx < this.getWorkBook().getColorTable().length )
			{
				Color mycolr = this.getWorkBook().getColorTable()[clidx];
				return mycolr;
			}
		}
		return Color.black;
	}

	/**
	 * sets the Color of the Cell Foreground pattern for this Cell.
	 * <p/>
	 * <br>
	 * <br>
	 * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
	 * Background color is the PATTERN color for all patterns not equal to
	 * PATTERN_SOLID <br>
	 * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
	 * Background Color is 64 (white).
	 *
	 * @param java .awt.Color col - color for the foreground
	 */
	public void setForegroundColor( Color col )
	{
		setFormatHandle();
		formatter.setForegroundColor( col );
	}

	/**
	 * get the Color of the Cell Background by this Cell and all other cells
	 * sharing this FormatId. <br>
	 * <br>
	 * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
	 * Background color is the PATTERN color for all patterns not equal to
	 * PATTERN_SOLID <br>
	 * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
	 * Background Color is 64 (white).
	 *
	 * @return java.awt.Color for the Font
	 */
	public Color getBackgroundColor()
	{
		if( this.mycell.getXfRec() != null )
		{
			Xf x = mycell.getXfRec();
			int clidx = x.getBackgroundColor();
			if( clidx < this.getWorkBook().getColorTable().length )
			{
				Color mycolr = this.getWorkBook().getColorTable()[clidx];
				return mycolr;
			}
		}
		return Color.white;
	}

	/**
	 * Return the cell background color i.e. the color if a pattern is set, or
	 * white if none
	 *
	 * @return java.awt.Color cell background color
	 */
	public Color getCellBackgroundColor()
	{
		setFormatHandle();
		int clidx = formatter.getCellBackgroundColor();
		if( clidx < this.wbh.getWorkBook().getColorTable().length )
		{
			Color mycolr = this.getWorkBook().getColorTable()[clidx];
			return mycolr;
		}
		return Color.white;
	}

	/**
	 * set the Color of the Cell Background pattern for this Cell. <br>
	 * <br>
	 * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
	 * Background color is the PATTERN color for all patterns not equal to
	 * PATTERN_SOLID <br>
	 * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
	 * Background Color is 64 (white).
	 *
	 * @param java .awt.Color col - background color
	 */
	public void setBackgroundColor( Color col )
	{
		setFormatHandle();
		formatter.setBackgroundColor( col );
	}

	/**
	 * set the Color of the Cell Background for this Cell. <br>
	 * <br>
	 * see FormatHandle.COLOR constants for valid values <br>
	 * <br>
	 * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
	 * Background color is the PATTERN color for all patterns not equal to
	 * PATTERN_SOLID <br>
	 * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
	 * Background Color is 64 (white).
	 *
	 * @param int t - Excel color constant for Cell Background color
	 */
	public void setCellBackgroundColor( int t )
	{
		setFormatHandle();
		formatter.setCellBackgroundColor( t );
	}

	/**
	 * set the Color of the Cell Background for this Cell. <br>
	 * <br>
	 * NOTE: Foreground color is the CELL BACKGROUND color for all patterns and
	 * Background color is the PATTERN color for all patterns not equal to
	 * PATTERN_SOLID <br>
	 * For PATTERN_SOLID, the Foreground color is the CELL BACKGROUND color and
	 * Background Color is 64 (white).
	 *
	 * @param java .awt.Color col - color of the cell background
	 */
	public void setCellBackgroundColor( Color col )
	{
		setFormatHandle();
		formatter.setCellBackgroundColor( col );
	}

	/**
	 * set the Color of the Cell Background Pattern for this Cell.
	 *
	 * @param java .awt.Color col - color of the pattern background
	 */
	public void setPatternBackgroundColor( Color col )
	{
		setFormatHandle();
		formatter.setBackgroundColor( col );
	}

	/**
	 * set the Background Pattern for this Cell. <br>
	 *
	 * @param int t - Excel pattern constant
	 * @see FormatHandle.PATTERN constants
	 */
	public void setBackgroundPattern( int t )
	{
		setFormatHandle();
		formatter.setPattern( t );
	}

	/**
	 * get the Background Pattern for this cell <br>
	 *
	 * @return int Excel pattern constant
	 * @see FormatHandle.PATTERN constants
	 */
	public int getBackgroundPattern()
	{
		return this.mycell.getXfRec().getFillPattern();
	}

	/**
	 * get the fill pattern for this cell <br>
	 * Same as getBackgroundPattern <br>
	 *
	 * @return int Excel fill pattern constant
	 * @see FormatHandle.PATTERN constants
	 */
	public int getFillPattern()
	{
		setFormatHandle();
		return formatter.getFillPattern();
	}

	/**
	 * set the Color of the Border for this Cell.
	 *
	 * @param java .awt.Color col - border color
	 */
	public void setBorderColor( Color col )
	{
		setFormatHandle();
		formatter.setBorderColor( col );
	}

	/**
	 * set the Color of the right Border line for this Cell.
	 *
	 * @param java .awt.Color col - right border color
	 */
	public void setBorderRightColor( Color col )
	{
		setFormatHandle();
		formatter.setBorderRightColor( col );
	}

	/**
	 * set the Color of the left Border line for this Cell.
	 *
	 * @param java .awt.Color col - left border color
	 */
	public void setBorderLeftColor( Color col )
	{
		setFormatHandle();
		formatter.setBorderLeftColor( col );
	}

	/**
	 * set the Color of the top Border line for this Cell.
	 *
	 * @param java .awt.Color col - top border color
	 */
	public void setBorderTopColor( Color col )
	{
		setFormatHandle();
		formatter.setBorderTopColor( col );
	}

	/**
	 * set the Color of the bottom Border line for this Cell.
	 *
	 * @param java .awt.Color col - bottom border color
	 */
	public void setBorderBottomColor( Color col )
	{
		setFormatHandle();
		formatter.setBorderBottomColor( col );
	}

	/**
	 * get the Font face used by this Cell.
	 *
	 * @return String the system name of the font for this Cell
	 */
	public String getFontFace()
	{
		if( this.getFont() == null )
		{
			return FormatHandle.DEFAULT_FONT_FACE;
		}
		return mycell.getFont().getFontName();
	}

	/**
	 * set the Font face used by this Cell.
	 *
	 * @param String fn - system name of the font
	 */
	public void setFontFace( String fn )
	{
		setFormatHandle();
		formatter.setFontName( fn );
	}

	/**
	 * set the Border line style for this Cell.
	 *
	 * @param short s - border constant
	 * @see FormatHandle.BORDER line style constants
	 */
	public void setBorderLineStyle( short s )
	{
		setFormatHandle();
		formatter.setBorderLineStyle( s );
	}

	/**
	 * set the Right Border line style for this Cell.
	 *
	 * @param short s - border constant
	 * @see FormatHandle.BORDER line style constants
	 */
	public void setRightBorderLineStyle( short s )
	{
		setFormatHandle();
		formatter.setRightBorderLineStyle( s );
	}

	/**
	 * set the Left Border line style for this Cell.
	 *
	 * @param short s - border constant
	 * @see FormatHandle.BORDER line style constants
	 */
	public void setLeftBorderLineStyle( short s )
	{
		setFormatHandle();
		formatter.setLeftBorderLineStyle( s );
	}

	/**
	 * set the Top Border line style for this Cell.
	 *
	 * @param short s - border constant
	 * @see FormatHandle.BORDER line style constants
	 */
	public void setTopBorderLineStyle( short s )
	{
		setFormatHandle();
		formatter.setTopBorderLineStyle( s );
	}

	/**
	 * set the Bottom Border line style for this Cell.
	 *
	 * @param short s - border constant
	 * @see FormatHandle.BORDER line style constants
	 */
	public void setBottomBorderLineStyle( short s )
	{
		setFormatHandle();
		formatter.setBottomBorderLineStyle( s );
	}

	/**
	 * removes the borders for this cell
	 */
	public void removeBorder()
	{
		setFormatHandle();
		formatter.removeBorders();
	}

	/**
	 * Get the ExtenXLS Font for this Cell, which roughly matches the
	 * functionality of the java.awt.Font class. <br>
	 * Due to awt problems on console systems, converting the ExtenXLS font to a
	 * GUI font is up to you. <br>
	 *
	 * @return ExtenXLS font for Cell
	 */
	public Font getFont()
	{
		return mycell.getFont();
	}

	/**
	 * convert the ExtenXLS font used by this Cell to java.awt.Font
	 *
	 * @return java.awt.Font for this cell
	 */
	public java.awt.Font getAwtFont()
	{
		String fface = "Arial";
		try
		{
			fface = this.getFontFace();
		}
		catch( Exception e )
		{
			;
		}
		int sz = 12;
		try
		{
			sz = (int) this.getFontSize();
		}
		catch( Exception e )
		{
			sz = 12; // back to default
		}
		sz += 4; // Excel seems to display 1 point larger than Java
		// Logger.logInfo("Displaying font:" + fface);
		HashMap ftmap = new HashMap();
		// implement underlines
		try
		{
			if( this.getIsUnderlined() )
			{
				ftmap.put( java.awt.font.TextAttribute.UNDERLINE, java.awt.font.TextAttribute.UNDERLINE );
			}
		}
		catch( Exception e )
		{
			;
		}
		// Ahhh, much cooler fonts here!
		ftmap.put( java.awt.font.TextAttribute.FAMILY, fface );
		float fx = this.getFontWeight();
		ftmap.put( java.awt.font.TextAttribute.SIZE, (float) sz );
		// TODO: Interpret other weights- LIGHT, DEMI_LIGHT, DEMI_BOLD, etc.
		if( fx == FormatHandle.BOLD )
		{
			ftmap.put( java.awt.font.TextAttribute.WEIGHT, java.awt.font.TextAttribute.WEIGHT_BOLD );
		}
		else
		{
			ftmap.put( java.awt.font.TextAttribute.WEIGHT, java.awt.font.TextAttribute.WEIGHT_REGULAR );
		}
		return new java.awt.Font( ftmap );
	}

	/**
	 * Set the Font for this Cell via font name, font style and font size <br>
	 * This method adds a new Font to the WorkBook. <br>
	 * Roughly matches the functionality of the java.awt.Font class.
	 *
	 * @param String fn - system name of the font e.g. "Arial"
	 * @param int    stl - font style (either Font.BOLD or Font.PLAIN)
	 * @param int    sz - font size in points
	 */
	public void setFont( String fn, int stl, int sz )
	{
		setFormatHandle();
		formatter.setFont( fn, stl, sz );
	}

	/**
	 * Get a CommentHandle to the note attached to this cell
	 */
	public CommentHandle getComment() throws DocumentObjectNotFoundException
	{
		// this needs significant cleanup. We should not have to iterate notes
		ArrayList notes = mycell.getSheet().getNotes();
		for( int i = 0; i < notes.size(); i++ )
		{
			Note n = (Note) notes.get( i );
			if( n.getCellAddressWithSheet().equals( this.getCellAddressWithSheet() ) )
			{
				return new CommentHandle( n );
			}
		}
		throw new DocumentObjectNotFoundException( "Note record not found at " + this.getCellAddressWithSheet() );
	}

	/**
	 * Removes any note/comment records attached to this cell
	 */
	public void removeComment()
	{
		try
		{
			CommentHandle note = this.getComment();
			note.remove();
		}
		catch( DocumentObjectNotFoundException e )
		{
		}
	}

	/**
	 * Creates a new annotation (Note or Comment) to the cell
	 *
	 * @param comment -- text of note
	 * @param author  -- name of author
	 * @return CommentHandle - handle which allows access to the Note object
	 * @see CommentHandle
	 */
	public CommentHandle createComment( String comment, String author )
	{
		Note n = mycell.getSheet().createNote( this.getCellAddress(), comment, author );
		return new CommentHandle( n );
	}

	/**
	 * returns whether the Font for this Cell is underlined
	 *
	 * @return boolean true if Font is underlined
	 */
	public boolean getIsUnderlined()
	{
		if( this.getFont() == null )
		{
			return false;
		}
		if( this.getFont().getUnderlineStyle() != 0x0 )
		{
			return true;
		}
		return false;
	}

	/**
	 * Set whether the Font for this Cell is underlined
	 *
	 * @param boolean b - true if the Font for this Cell should be underlined
	 *                (single underline only)
	 */
	public void setUnderlined( boolean isUnderlined )
	{
		setFormatHandle();
		if( isUnderlined )
		{
			this.getFont().setUnderlineStyle( Font.STYLE_UNDERLINE_SINGLE );
		}
		else
		{
			this.getFont().setUnderlineStyle( Font.STYLE_UNDERLINE_NONE );
		}
	}

	/**
	 * set the Font Color for this Cell <br>
	 * <br>
	 * see FormatHandle.COLOR constants for valid values
	 *
	 * @param int t - Excel color constant
	 */
	public void setFontColor( int t )
	{
		setFormatHandle();
		formatter.setFontColor( t );
	}

	/**
	 * Returns any other Cells merged with this one, or null if this Cell is not
	 * a part of a merged range. <br>
	 * Adding and/or removing Cells from this CellRange will merge or unmerge
	 * the Cells.
	 *
	 * @return CellRange object containing all Cells in this Cell's merged
	 * range.
	 */
	public CellRange getMergedCellRange()
	{
		return mycell.getMergeRange();
	}

	/**
	 * Returns if the Cell is the parent (cell containing display value) of a
	 * merged cell range.
	 *
	 * @return boolean true if this cell is a merge parent cell
	 */
	public boolean isMergeParent()
	{
		try
		{
			CellRange cr = mycell.getMergeRange();
			if( cr == null )
			{
				return false;
			}
			int[] i = cr.getRangeCoords();
			if( ((this.getRowNum() + 1) == i[0]) && (this.getColNum() == i[1]) )
			{
				return true;
			}
		}
		catch( Exception e )
		{
			return false;
		}
		return false;

	}

	/**
	 * Returns the ColHandle for the Cell.
	 *
	 * @return ColHandle for the Cell
	 */
	public ColHandle getCol()
	{
		try
		{
			if( mycol == null )
			{
				mycol = wsh.getCol( this.getColNum() );
			}
		}
		catch( ColumnNotFoundException ex )
		{
			// can't happen, the column has to exist because we're in it
			throw new RuntimeException( ex );
		}
		return this.mycol;
	}

	/**
	 * Returns the RowHandle for the Cell.
	 *
	 * @return RowHandle representing the Row for the Cell
	 */
	public RowHandle getRow()
	{
		if( myrow == null )
		{
			myrow = new RowHandle( mycell.getRow(), this.wsh );
		}
		return this.myrow;
	}

	/**
	 * Returns the value of this Cell in the native underlying data type.
	 * <p/>
	 * Formula cells will return the calculated value of the formula in the
	 * calculated data type.
	 * <p/>
	 * Use 'getStringVal()' to return a String regardless of underlying value
	 * type.
	 *
	 * @return Object value for this Cell
	 */
	@Override
	public Object getVal()
	{
		return FormulaHandle.sanitizeValue( mycell.getInternalVal() );
	}

	/**/

	/**
	 * returns the java Type string of the Cell <br>
	 * One of: <li>"String" <li>"Float" <li>"Integer" <li>"Formula" <li>"Double"
	 *
	 * @return String java data type
	 */
	public String getCellTypeName()
	{
		String typename = "Object";
		int tp = getCellType();
		switch( tp )
		{
			case XLSConstants.TYPE_BLANK:
				typename = "String";
				break;
			case XLSConstants.TYPE_STRING:
				typename = "String";
				break;
			case XLSConstants.TYPE_FP:
				typename = "Float";
				break;
			case XLSConstants.TYPE_INT:
				typename = "Integer";
				break;
			case XLSConstants.TYPE_FORMULA:
				typename = "Formula";
				break;
			case XLSConstants.TYPE_DOUBLE:
				typename = "Double";
				break;
		}
		return typename;
	}

	/**
	 * Returns the type of the Cell as an int <br>
	 * <li>TYPE_STRING = 0, <li>TYPE_FP = 1, <li>TYPE_INT = 2, <li>TYPE_FORMULA
	 * = 3, <li>TYPE_BOOLEAN = 4, <li>TYPE_DOUBLE = 5;
	 *
	 * @return int type for this Cell
	 */
	@Override
	public int getCellType()
	{
		return mycell.getCellType();
	}

	/**
	 * return the underlying cell record <br>
	 * for internal API use only
	 *
	 * @return XLS cell record
	 */
	public BiffRec getCell()
	{
		return mycell;
	}

	/**
	 * returns Border Colors of Cell in: top, left, bottom, right order <br>
	 * returns null or a java.awt.Color object for each of 4 sides
	 *
	 * @return java.awt.Color array representing the 4 borders of the cell
	 */
	public Color[] getBorderColors()
	{
		getFormatHandle();
		return this.formatter.getBorderColors();
	}

	/**
	 * returns the low-level bytes for the underlying BIFF8 record. <br>
	 * For Internal API use only
	 *
	 * @return bytes for the underlying record
	 */
	public byte[] getBytes()
	{
		return mycell.getData();
	}

	/**
	 * Returns the column number of this Cell.
	 *
	 * @return int the Column Number of the Cell
	 */
	@Override
	public int getColNum()
	{
		setMulblank();
		return mycell.getColNumber();
	}

	/**
	 * Returns the row number of this Cell.
	 * <p/>
	 * NOTE: This is the 1-based row number such as you will see in a
	 * spreadsheet UI.
	 * <p/>
	 * ie: A1 = row 1
	 *
	 * @return 1-based int the Row Number of the Cell
	 */
	@Override
	public int getRowNum()
	{
		return mycell.getRowNumber();
	}

	/**
	 * Returns the Address of this Cell as a String.
	 *
	 * @return String the address of this Cell in the WorkSheet
	 */
	@Override
	public String getCellAddress()
	{
		setMulblank();
		return mycell.getCellAddress();

	}

	public int[] getIntLocation()
	{
		setMulblank();
		return mycell.getIntLocation();
	}

	/**
	 * Returns the Address of this Cell as a String. Includes the sheet name in
	 * the address.
	 *
	 * @return String the address of this Cell in the WorkSheet
	 */
	public String getCellAddressWithSheet()
	{
		setMulblank();
		return mycell.getCellAddressWithSheet();

	}

	/**
	 * sets the column number referenced in the set of multiple blanks <br>
	 * for multiple blank cells only <br>
	 * for Internal Use. Not intended for the End User.
	 */
	short mulblankcolnum = -1;

	public void setBlankRef( int c )
	{
		mulblankcolnum = (short) c;
	}

	/**
	 * Returns the name of this Cell's WorkSheet as a String.
	 *
	 * @return String the name this Cell's WorkSheet
	 */
	@Override
	public String getWorkSheetName()
	{
		if( wsh == null )
		{
			try
			{
				return mycell.getSheet().getSheetName();
			}
			catch( Exception e )
			{
				return "";
			}
		}
		return wsh.getSheetName();
	}

	/**
	 * Returns the value of the Cell as a String. <br>
	 * boolean Cell types will return "true" or "false"
	 *
	 * @return String the value of the Cell
	 */
	public String getStringVal()
	{
		return mycell.getStringVal();
	}

	/**
	 * Gets the value of the cell as a String with the number format applied.
	 * Boolean cell types will return "true" or "false". Custom number formats
	 * are not currently supported, although they will be written correctly to
	 * the output file. Patterns that display negative numbers in red are not
	 * currently supported; the number will be prefixed with a minus sign
	 * instead. For more information on number format patterns see <a
	 * href="http://support.microsoft.com/kb/264372">Microsoft KB264372</a>.
	 *
	 * @return the value of the cell as a string formatted according to the cell
	 * type and, if present, the number format pattern
	 */
	@Override
	public String getFormattedStringVal()
	{
		FormatHandle myfmt = this.getFormatHandle();
		return CellFormatFactory.fromPatternString( myfmt.getFormatPattern() ).format( this );
	}

	/**
	 * Gets the value of the cell as a String with the number format applied.
	 * Boolean cell types will return "true" or "false". Custom number formats
	 * are not currently supported, although they will be written correctly to
	 * the output file. Patterns that display negative numbers in red are not
	 * currently supported; the number will be prefixed with a minus sign
	 * instead. For more information on number format patterns see <a
	 * href="http://support.microsoft.com/kb/264372">Microsoft KB264372</a>.
	 *
	 * @param formatForXML if true non-compliant characters will be properly qualified
	 * @return the value of the cell as a string formatted according to the cell
	 * type and, if present, the number format pattern
	 */
	public String getFormattedStringVal( boolean formatForXML )
	{
		FormatHandle myfmt = this.getFormatHandle();
		String val = this.getVal().toString();
		if( formatForXML )
		{
			val = com.extentech.formats.XLS.OOXMLAdapter.stripNonAscii( val ).toString();
		}
		return CellFormatFactory.fromPatternString( myfmt.getFormatPattern() ).format( val );
	}

	/**
	 * Returns the value of the Cell as a String. <br>
	 * For numeric cell values, including cells containing a formula which
	 * return a numeric value, the notation to be used in representing the value
	 * as a String must be specified. <br>
	 * The Notation choices are: <li>CellHandle.NOTATION_STANDARD e.g.
	 * "8002974342" <li>CellHandle.NOTATION_SCIENTIFIC e.g. "8.002974342E9" <li>
	 * CellHandle.NOTATION_SCIENTIFIC_EXCEL e.g. "8.002974342+E9" <br>
	 * <br>
	 * For non-numeric values, the value of the cell as a string is returned <br>
	 * boolean Cell types will return "true" or "false"
	 *
	 * @param int notation one of the CellHandle.NOTATION constants for numeric
	 *            cell types; ignored for other cell types
	 * @param int notation - notation constant
	 * @return String the value of the Cell
	 */
	public String getStringVal( int notation )
	{
		String numval = mycell.getStringVal();
		int i = this.getCellType();
		if( (i == TYPE_FP) || (i == TYPE_INT) || (i == TYPE_FORMULA) || (i == TYPE_DOUBLE) )
		{
			return ExcelTools.formatNumericNotation( numval, notation );
		}
		return numval;
	}

	/**
	 * Returns the value of the Cell as a String with the specified encoding. <br>
	 * boolean Cell types will return "true" or "false"
	 *
	 * @param String encoding
	 * @return String the value of the Cell
	 */
	public String getStringVal( String encoding )
	{
		return mycell.getStringVal( encoding );
	}

	/**
	 * Returns the Formatting record ID (FormatId) for this Cell <br>
	 * This can be used with 'setFormatId(int i)' to copy the formatting from
	 * one Cell to another (e.g. a template cell to a new cell)
	 *
	 * @return int the FormatId for this Cell
	 */
	@Override
	public int getFormatId()
	{
		setMulblank();
		return mycell.getIxfe();
	}

	/**
	 * get the conditional formatting record ID (FormatId) <br>
	 * returns the normal FormatId if the condition(s) in the conditional format
	 * have not been met
	 *
	 * @return int the conditional format id
	 */
	public int getConditionalFormatId()
	{
		ConditionalFormatHandle[] cfhandles = this.getConditionalFormatHandles();
		if( (cfhandles == null) || (cfhandles.length == 0) )
		{
			return this.getFormatId();
		}
		// TODO: only supporting first cfmat handle again
		Condfmt cfmt = cfhandles[0].getCndfmt();

		Iterator x = cfmt.getRules().iterator();
		while( x.hasNext() )
		{
			Cf cx1 = ((Cf) x.next());
			// create a ptgref for this Cell
			Ptg pref = new PtgRef( this.getCellAddress(), this.mycell, false );

			if( cx1.evaluate( pref ) )
			{
				// TODO: evaluate and combine multiple rules...
				// currently returns on first true format
				int ret = cfmt.getCfxe();
				return ret;
			}
		}
		return this.getFormatId();
	}

	/**
	 * move this cell to another row <br>
	 * throws CellPositionConflictException if there is a cell in that position
	 * already
	 *
	 * @param int newrow - new row number
	 * @throws CellPositionConflictException
	 */
	public void moveToRow( int newrow ) throws CellPositionConflictException
	{
		String newaddr = ExcelTools.getAlphaVal( mycell.getColNumber() );
		newaddr += String.valueOf( newrow );
		this.moveTo( newaddr );
	}

	/**
	 * move this cell to another row <br>
	 * overwrite any cells in the destination
	 *
	 * @param int newrow - new row number
	 */
	public void moveAndOverwriteToRow( int newrow )
	{
		String newaddr = ExcelTools.getAlphaVal( mycell.getColNumber() );
		newaddr += String.valueOf( newrow );
		this.moveAndOverwriteTo( newaddr );
	}

	/**
	 * move this cell to another column <br>
	 * throws CellPositionConflictException if there is a cell in that position
	 * already
	 *
	 * @param String newcol - the new column in alpha format e.g. "A", "B" ...
	 * @throws CellPositionConflictException
	 */
	public void moveToCol( String newcol ) throws CellPositionConflictException
	{
		String newaddr = newcol;
		newaddr += String.valueOf( mycell.getRowNumber() + 1 );
		this.moveTo( newaddr );
	}

	/**
	 * Copy all formats from a source Cell to this Cell
	 *
	 * @param CellHandle source - source cell
	 */
	public void copyFormat( CellHandle source )
	{
		this.getCell().copyFormat( source.getCell() );
	}

	/**
	 * copy this Cell to another location.
	 *
	 * @param String newaddr - address for copy of this Cell in Excel-style e.g.
	 *               "A1"
	 * @return returns the newly copied CellHandle
	 * @throws CellPositionConflictException if there is a cell in the new address already
	 */
	public CellHandle copyTo( String newaddr ) throws CellPositionConflictException
	{

		// check for existing
		Boundsheet bs = this.mycell.getSheet();

		BiffRec rec = this.mycell;
		XLSRecord nucell = (XLSRecord) ((XLSRecord) rec).clone();
		int[] rc = ExcelTools.getRowColFromString( newaddr );
		nucell.setRowNumber( rc[0] );
		nucell.setCol( (short) rc[1] );
		nucell.setXFRecord( this.mycell.getIxfe() );
		bs.addRecord( nucell, rc );

		CellHandle ret = new CellHandle( nucell, wbh );
		if( mycell.hyperlink != null )
		{
			// set the bounds of the mycell.hyperlink
			ret.setURL( this.getURL() );
		}
		ret.setWorkSheetHandle( this.getWorkSheetHandle() );
		// this.getWorkSheetHandle().cellhandles.put(this.getCellAddress(),
		// this);
		return ret;
	}

	/**
	 * Removes this Cell from the WorkSheet
	 *
	 * @param boolean nullme - true if this CellHandle should be nullified after
	 *                removal
	 */
	public void remove( boolean nullme )
	{
		mycell.getSheet().removeCell( mycell );
		if( nullme )
		{
			try
			{
				this.finalize();
			}
			catch( Throwable t )
			{
				;
			}
		}
	}

	/**
	 * move this cell to another location.
	 *
	 * @param String newaddr - the new address for Cell in Excel-style notation
	 *               e.g. "A1"
	 * @throws CellPositionConflictException if there is a cell in the new address already
	 */
	public void moveTo( String newaddr ) throws CellPositionConflictException
	{

		// check for existing
		Boundsheet bs = mycell.getSheet();
		BiffRec oldhand = bs.getCell( newaddr );
		if( oldhand != null )
		{
			throw new CellPositionConflictException( newaddr );
		}
		bs.moveCell( this.getCellAddress(), newaddr );

		if( mycell.hyperlink != null )
		{
			// set the bounds of the mycell.hyperlink
			// int[] bnds =
			// ExcelTools.getRowColFromString(this.getCellAddress());

			int[] bnds = ExcelTools.getRowColFromString( this.getCellAddress() );

			Hlink hl = mycell.hyperlink;
			hl.setRowFirst( bnds[0] );
			hl.setRowLast( bnds[0] );
			hl.setColFirst( bnds[1] );
			hl.setColLast( bnds[1] );
			hl.init();
		}
	}

	/**
	 * move this cell to another location, overwriting any cells that are in the way
	 *
	 * @param String newaddr - the new address for Cell in Excel-style notation
	 *               e.g. "A1"
	 */
	public void moveAndOverwriteTo( String newaddr )
	{

		// check for existing
		Boundsheet bs = mycell.getSheet();
		BiffRec oldhand = bs.getCell( newaddr );
		bs.moveCell( this.getCellAddress(), newaddr );

		if( mycell.hyperlink != null )
		{
			int[] bnds = ExcelTools.getRowColFromString( this.getCellAddress() );

			Hlink hl = mycell.hyperlink;
			hl.setRowFirst( bnds[0] );
			hl.setRowLast( bnds[0] );
			hl.setColFirst( bnds[1] );
			hl.setColLast( bnds[1] );
			hl.init();
		}
	}

	/**
	 * Set a conditional format upon this cell.
	 * <p/>
	 * Note that conditional format handles are bound to a specific worksheet,
	 *
	 * @param format A ConditionalFormatHandle in the same worksheet
	 */
	public void addConditionalFormat( ConditionalFormatHandle format )
	{
		format.addCell( this );
	}

	/**
	 * returns an array of FormatHandles for the Cell that have the current
	 * conditional format applied to them.
	 * <p/>
	 * This behavior is still in testing, and may change
	 *
	 * @return an array of FormatHandles, one for each of the Conditional
	 * Formatting rules
	 */
	public FormatHandle[] getConditionallyFormattedHandles()
	{
		// TODO, should these be read-only?
		// TODO, this is bad, only handles first cf record for the cell
		ConditionalFormatHandle[] cfhandles = this.getConditionalFormatHandles();
		if( cfhandles != null )
		{
			FormatHandle[] fmx = new FormatHandle[cfhandles[0].getCndfmt().getRules().size()];
			for( int t = 0; t < fmx.length; t++ )
			{
				fmx[t++] = new FormatHandle( cfhandles[0].getCndfmt(), wbh, t, this );
			}
			return fmx;
		}
		return null;
	}

	/**
	 * return all the ConditionalFormatHandles for this Cell, if any
	 *
	 * @return
	 */
	public ConditionalFormatHandle[] getConditionalFormatHandles()
	{
		WorkSheetHandle sh = this.getWorkSheetHandle();
		if( sh == null )
		{
			return null;
		}
		ConditionalFormatHandle[] cfs = sh.getConditionalFormatHandles();
		ArrayList cfhandles = new ArrayList();
		for( int i = 0; i < cfs.length; i++ )
		{
			if( cfs[i].contains( this ) )
			{
				cfhandles.add( cfs[i] );
			}
		}
		ConditionalFormatHandle[] c = new ConditionalFormatHandle[cfhandles.size()];
		return (ConditionalFormatHandle[]) cfhandles.toArray( c );
	}

	/**
	 * Gets the FormatHandle (a Format Object describing the formats for this
	 * Cell) for this Cell.
	 *
	 * @return FormatHandle for this Cell
	 */
	public FormatHandle getFormatHandle()
	{
		if( this.formatter == null )
		{
			this.setFormatHandle();
		}

		return this.formatter;
	}

	/**
	 * locks or unlocks this Cell for editing
	 *
	 * @param boolean locked - true if Cell should be locked, false otherwise
	 */
	public void setLocked( boolean locked )
	{
		// create a new xf for this
		// this causes formats to be lost this.useExistingXF = false;
		getFormatHandle().setLocked( locked );
	}

	/**
	 * Hides or shows the formula for this Cell, if present
	 *
	 * @param boolean hidden - setting whether to hide or show formulas for this
	 *                Cell
	 */
	public void setFormulaHidden( boolean hidden )
	{
		// create a new xf for this
		// this causes formats to be lost this.useExistingXF = false;
		getFormatHandle().setFormulaHidden( hidden );
	}

	/**
	 * Sets the FormatHandle (a Format Object describing the formats for this
	 * Cell) for this Cell
	 *
	 * @param FormatHandle to apply to this Cell
	 * @see FormatHandle
	 */
	public void setFormatHandle( FormatHandle f )
	{
		f.addCell( this.mycell );
		this.formatter = f;
	}

	/**
	 * Sets the Formatting record ID (FormatId) for this Cell
	 * <p/>
	 * This can be used with 'getFormatId()' to copy the formatting from one
	 * Cell to another (ie: a template cell to a new cell)
	 *
	 * @param int i - the new index to the Format for this Cell
	 */
	public void setFormatId( int i )
	{
		mycell.setXFRecord( i );
	}

	/**
	 * Resets this cell's format to the default.
	 */
	public void clearFormats()
	{
		this.setFormatId( this.getWorkBook().getWorkBook().getDefaultIxfe() );
	}

	/**
	 * Resets this cells contents to blank.
	 */
	public void clearContents()
	{
		this.setVal( null );
	}

	/**
	 * Resets this cell to the default, as if it had just been added.
	 */
	public void clear()
	{
		this.clearFormats();
		this.clearContents();
	}

	/**
	 * Returns the Formula Handle (an Object describing a Formula) for this
	 * Cell, if it contains a formula
	 *
	 * @return FormulaHandle the Formula of the Cell
	 * @throws FormulaNotFoundException
	 * @see FormulaHandle
	 */
	public FormulaHandle getFormulaHandle() throws FormulaNotFoundException
	{
		Formula f = (Formula) mycell.getFormulaRec();
		if( f == null )
		{
			throw new FormulaNotFoundException( "No Formula for: " + getCellAddress() );
		}
		FormulaHandle fh = new FormulaHandle( f, this.wbh );
		return fh;
	}

	/**
	 * Returns the Hyperlink URL String for this Cell, if any
	 *
	 * @return String URL if this Cell contains a hyperlink
	 */
	public String getURL()
	{
		if( mycell.hyperlink != null )
		{
			return mycell.hyperlink.getURL();
		}
		return null;
	}

	/**
	 * Returns the URL Description String for this Cell, if any
	 *
	 * @return String URL Description, if this Cell contains a hyperlink
	 */
	public String getURLDescription()
	{
		if( mycell.hyperlink != null )
		{
			return mycell.hyperlink.getDescription();
		}
		return "";
	}

	/**
	 * returns true if this Cell contains a hyperlink
	 *
	 * @return boolean true if this Cell contains a hyperlink
	 */
	public boolean hasHyperlink()
	{
		return (mycell.hyperlink != null);
	}

	/**
	 * Creates a new Hyperlink for this Cell from a URL String. Can be any valid
	 * URL. This URL String must include the protocol. <br>
	 * For Example, "http://www.extentech.com/" <br>
	 * To remove a hyperlink, pass in null for the URL String
	 *
	 * @param String urlstr - the URL String for this Cell
	 */
	public void setURL( String urlstr )
	{
		if( urlstr == null )
		{
			mycell.hyperlink = null;
			// TODO: remove existing Hlink from stream
			// mycell.hyperlink.remove(true);
			return;
		}
		setURL( urlstr, "", "" );
	}

	/**
	 * Creates a new Hyperlink for this Cell from a URL String, a descrpiton and
	 * textMark text. <br>
	 * <br>
	 * The URL String Can be any valid URL. This URL String must include the
	 * protocol. For Example, "http://www.extentech.com/" <br>
	 * The textMark text is the porition of the URL that follows # <br>
	 * <br>
	 * NOTE: URL text and textMark text must not be null or ""
	 *
	 * @param String urlstr - the URL String
	 * @param String desc - the description text
	 * @param String textMark - the text that follows #
	 */
	public void setURL( String urlstr, String desc, String textMark )
	{
		if( mycell.hyperlink != null )
		{
			mycell.hyperlink.setURL( urlstr, desc, textMark );
		}
		else
		{
			// create new URL
			mycell.hyperlink = (Hlink) Hlink.getPrototype();
			mycell.hyperlink.setURL( urlstr, desc, textMark );

			// why would we want to set the val during this operation?
			// if (!desc.equals("")) setVal(desc);

			// set the bounds of the mycell.hyperlink
			int[] bnds = ExcelTools.getRowColFromString( this.getCellAddress() );
			mycell.hyperlink.setRowFirst( bnds[0] );
			mycell.hyperlink.setColFirst( bnds[1] );
			mycell.hyperlink.setRowLast( bnds[0] );
			mycell.hyperlink.setColLast( bnds[1] );
		}
	}

	/**
	 * Sets a hyperlink to a location within the current template <br>
	 * The URL String should be prefixed with "file://" <br>
	 *
	 * @param String fileURLStr - the file URL String
	 */
	// TODO: document this: NOTE: Excel File URL in actuality does not match
	// documentation
	public void setFileURL( String fileURLStr )
	{
		setFileURL( fileURLStr, "", "" );
	}

	/**
	 * Sets a hyperlink to a location within the current template, and includes
	 * additional optional information: description + textMark text <br>
	 * <br>
	 * The URL String should be prefixed with "file://" <br>
	 * <br>
	 * The textMark text is the porition of the URL that follows #
	 *
	 * @param String fileURLstr - the file URL String
	 * @param String desc - the description text
	 * @param String textMark - text that follows #
	 */
	// TODO: this documentation is contradictory
	public void setFileURL( String fileURLstr, String desc, String textMark )
	{
		if( mycell.hyperlink != null )
		{
			mycell.hyperlink.setFileURL( fileURLstr, desc, textMark );
		}
		else
		{
			mycell.hyperlink = (Hlink) Hlink.getPrototype();

			mycell.hyperlink.setFileURL( fileURLstr, desc, textMark );
			if( !desc.equals( "" ) )
			{
				setVal( desc );
			}

			// set the bounds of the mycell.hyperlink
			int[] bnds = ExcelTools.getRowColFromString( this.getCellAddress() );
			mycell.hyperlink.setRowFirst( bnds[0] );
			mycell.hyperlink.setColFirst( bnds[1] );
			mycell.hyperlink.setRowLast( bnds[0] );
			mycell.hyperlink.setColLast( bnds[1] );
		}
	}

	/**
	 * Set the val of this Cell to an object <br>
	 * <br>
	 * The object may be one of type: <br>
	 * String, Float, Integer, Double, Long, Boolean, java.sql.Date, or null <br>
	 * <br>
	 * To set a Cell to a formula, obj should be a string begining with "=" <br>
	 * <br>
	 * To set a Cell to an array formula, obj should be a string begining with
	 * "{=" <br>
	 * If you wish to put a line break in a string value, use the newline "\n"
	 * character. Note this will not function unless you also apply a format
	 * handle to the cell with WrapText=true
	 *
	 * @param Object obj - the object to set the value of the Cell to
	 * @throws CellTypeMismatchException
	 */

	public void setVal( Object obj ) throws CellTypeMismatchException
	{
		if( this.wbh.getFormulaCalculationMode() != WorkBook.CALCULATE_EXPLICIT )
		{
			this.clearAffectedCells(); // blow out cache
		}

		if( obj instanceof java.sql.Date )
		{
			this.setVal( (java.sql.Date) obj, null );
			return;
		}
		if( obj instanceof String )
		{
			String formstr = (String) obj;
			if( (formstr.indexOf( "=" ) == 0) || formstr.startsWith( "{=" ) )
			{ // Formula or array string
				try
				{
					this.setFormula( formstr );
					return;
				}
				catch(/* 20070212 KSC: FunctionNotSupported */Exception a )
				{
					Logger.logWarn( "CellHandle.setVal() failed.  Setting Formula to " + obj.toString() + " failed: " + a );
				}
			}
		}
		try
		{
			setBiffRecValue( obj );
		}
		catch( FunctionNotSupportedException fnse )
		{
			// suppress these -- cell has been changed
		}
		catch( Exception e )
		{
			// NOT a CTMME always

			throw new CellTypeMismatchException( e.toString() );
		}
	}

	/**
	 * set the value of this cell to String s <br>
	 * NOTE: this method will not check for formula references or do any data
	 * conversions <br>
	 * This method is useful when a string may start with = but you do not want
	 * to convert to a Formula value
	 * <p/>
	 * If you wish to put a line break in the string use the newline "\n"
	 * character. Note this will not function unless you also apply a format
	 * handle to the cell with WrapText=true
	 *
	 * @param String s - the String value to set the Cell to
	 * @throws CellTypeMismatchException
	 * @see setVal(Object obj)
	 */
	public void setStringVal( String s )
	{
		try
		{
			if( ((s == null) || s.equals( "" )) && !(mycell instanceof Blank) )
			{
				changeCellType( s );
			}
			else if( (s != null) && !s.equals( "" ) )
			{
				if( !(mycell instanceof Labelsst) )
				{
					changeCellType( " " ); // avoid potential issues with string
				}
				// values beginning with "="
				mycell.setStringVal( s );
			}
		}
		catch( Exception e )
		{
			throw new CellTypeMismatchException( e.toString() );
		}
	}

	/**
	 * set the value of this cell to Unicodestring us <br>
	 * NOTE: This method will not check for formula references or do any data
	 * conversions <br>
	 * Useful when strings may start with = but you do not want to convert to a
	 * formula value
	 *
	 * @param Unicodestring us - Unicode String
	 * @throws CellTypeMismatchException
	 */
	public void setStringVal( Unicodestring us )
	{
		try
		{
			if( ((us == null) || us.equals( "" )) && !(mycell instanceof Blank) )
			{
				changeCellType( null ); // set to blank
			}
			else if( (us != null) && !us.equals( "" ) )
			{
				if( !(mycell instanceof Labelsst) )
				{
					changeCellType( " " ); // avoid potential issues with string
				}
				// values beginning with "="
				((Labelsst) mycell).setStringVal( us );
			}
		}
		catch( Exception e )
		{
			throw new CellTypeMismatchException( e.toString() );
		}
	}

	/**
	 * this method will be fired as each record is parsed from an input
	 * Spreadsheet
	 * <p/>
	 * Dec 15, 2010
	 */
	public void fireParserEvent()
	{

	}

	/**
	 * Returns a String representation of this CellHandle
	 *
	 * @see java.lang.Object#toString()
	 */
	public String toString()
	{
		String ret = this.getCellAddress() + ":" + this.getStringVal();
		if( this.getURL() != null )
		{
			ret += this.getURL();
		}
		return ret;
	}

	/**
	 * Set the Value of the Cell to a double
	 *
	 * @param double d- double value to set this Cell to
	 * @throws CellTypeMismatchException
	 */
	public void setVal( double d ) throws CellTypeMismatchException
	{
		this.setVal( new Double( d ) );
	}

	/**
	 * Set value of this Cell to a Float
	 *
	 * @param float f - float value to set this Cell to
	 * @throws CellTypeMismatchException
	 */
	public void setVal( float f ) throws CellTypeMismatchException
	{
		this.setVal( new Float( f ) );
	}

	/**
	 * Sets the value of this Cell to a java.sql.Date. <br>
	 * You must also specify a formatting pattern for the new date, or null for
	 * the default date format ("m/d/yy h:mm".) <br>
	 * <br>
	 * valid date format patterns: <br>
	 * "m/d/y" <br>
	 * "d-mmm-yy" <br>
	 * "d-mmm" <br>
	 * "mmm-yy" <br>
	 * "h:mm AM/PM" <br>
	 * "h:mm:ss AM/PM" <br>
	 * "h:mm" <br>
	 * "h:mm:ss" <br>
	 * "m/d/yy h:mm" <br>
	 * "mm:ss" <br>
	 * "[h]:mm:ss" <br>
	 * "mm:ss.0"
	 *
	 * @param java   .sql.Date dt - the value of the new Cell
	 * @param String fmt - date formatting pattern
	 */
	public void setVal( java.sql.Date dt, String fmt )
	{

		if( this.wbh.getFormulaCalculationMode() != WorkBook.CALCULATE_EXPLICIT )
		{
			this.clearAffectedCells(); // blow out cache
		}
		if( fmt == null )
		{
			fmt = "m/d/yyyy";
		}
		this.setVal( new Double( DateConverter.getXLSDateVal( dt ) ) );
		this.setFormatPattern( fmt );
	}

	/**
	 * Sets the value of this Cell to a boolean value
	 *
	 * @param boolean b - boolean value to set this Cell to
	 * @throws CellTypeMismatchException
	 */
	public void setVal( boolean b ) throws CellTypeMismatchException
	{
		setVal( Boolean.valueOf( b ) );
	}

	/**
	 * Set the value of this Cell to an int value <br>
	 * NOTE: setting a Boolean Cell type to a zero or a negative number will set
	 * the Cell to 'false'; setting it to an int value 1 or greater will set it
	 * to true.
	 *
	 * @param int i - int value to set this Cell to
	 * @throws CellTypeMismatchException
	 */
	public void setVal( int i ) throws CellTypeMismatchException
	{
		if( mycell.getCellType() == XLSConstants.TYPE_BOOLEAN )
		{
			if( i > 0 )
			{
				setVal( Boolean.valueOf( true ) );
			}
			else
			{
				setVal( Boolean.valueOf( false ) );
			}
		}
		else
		{
			setVal( Integer.valueOf( i ) );
		}
	}

	/**
	 * returns the value of this Cell as a double, if possible, or NaN if Cell
	 * value cannot be converted to double
	 *
	 * @return double value or NaN if the Cell value cannot be converted to a
	 * double
	 */
	public double getDoubleVal()
	{
		return mycell.getDblVal();
	}

	/**
	 * returns the value of this Cell as a int, if possible, or NaN if Cell
	 * value cannot be converted to int
	 *
	 * @return int value or NaN if the Cell value cannot be converted to an int
	 */
	public int getIntVal()
	{
		return mycell.getIntVal();
	}

	/**
	 * returns the value of this Cell as a float, if possible, or NaN if Cell
	 * value cannot be converted to float
	 *
	 * @return float value or NaN if the Cell value cannot be converted to an
	 * float
	 */
	public float getFloatVal()
	{
		return mycell.getFloatVal();
	}

	/**
	 * returns the value of this Cell as a boolean <br>
	 * If the Cell is not of type Boolean, returns false
	 *
	 * @return boolean value of cell
	 */
	public boolean getBooleanVal()
	{
		return mycell.getBooleanVal();
	}

	/**
	 * Set a cell to Excel-compatible formula passed in as String. <br>
	 *
	 * @param String formStr - the Formula String
	 * @throws FunctionNotSupportedException if unable to parse string correctly
	 */
	public void setFormula( String formStr ) throws FunctionNotSupportedException
	{
		int ixfe = this.mycell.getIxfe();
		this.remove( true );
		this.mycell = wsh.add( formStr, this.getCellAddress() ).mycell;
		this.mycell.setXFRecord( ixfe );
	}

	/**
	 * Set a cell to formula passed in as String. Sets the cachedValue as well,
	 * so no calculating is necessary.
	 * <p/>
	 * Parses the string to convert into a Excel formula. <br>
	 * IMPORTANT NOTE: if cell is ALREADY a formula String this method will NOT
	 * reset it
	 *
	 * @param String formulaStr - The excel-compatible formula string to pass in
	 * @param Object value - The calculated value of the formula
	 * @throws Exception if unable to parse string correctly
	 */
	public void setFormula( String formStr, Object value ) throws Exception
	{
		if( !(this.mycell instanceof Formula) )
		{
			int ixfe = this.mycell.getIxfe();
			CellRange cr = this.mycell.getMergeRange();
			int r = this.getRowNum();
			int c = this.getColNum();
			this.remove( true );
			this.mycell = wsh.add( formStr, r, c, ixfe ).mycell;
			this.mycell.setMergeRange( cr );
		}
		Formula f = (Formula) this.mycell;
		f.setCachedValue( value );
	}

	/**
	 * sets the formula for this cellhandle using a stack of Ptgs appropriate
	 * for formula records. This method also sets the cachedValue of the formula
	 * as well, so no new calculating is necessary.
	 *
	 * @param Stack  newExp - Stack of Ptgs
	 * @param Object value - calculated value of formula
	 */
	public void setFormula( Stack newExp, Object value )
	{
		if( !(this.mycell instanceof Formula) )
		{
			int ixfe = this.mycell.getIxfe();
			CellRange mccr = this.mycell.getMergeRange();
			int r = this.getRowNum();
			int c = this.getColNum();
			this.remove( true );
			this.mycell = wsh.add( "=0", r, c, ixfe ).mycell; // add the most
			// basic formula so
			// can modify below
			// ((:
			this.mycell.setMergeRange( mccr );
		}
		try
		{
			Formula f = (Formula) this.mycell;
			f.setExpression( newExp );
			f.setCachedValue( value );
		}
		catch( Exception e )
		{
			;// do what??
		}
	}

	/**
	 * Returns the size of the merged cell area, if one exists.
	 *
	 * @param row    this parameter is ignored
	 * @param column this parameter is ignored
	 * @return a 2 position int array with number of rows and number of cols
	 * @deprecated since October 2012. This method duplicates the functionality
	 * of {@link #getMergedCellRange()}, which it calls internally.
	 * That method should be used instead.
	 */
	@Deprecated
	public int[] getSpan( int row, int column )
	{
		CellRange mergerange = getMergedCellRange();
		if( mergerange != null )
		{
			if( DEBUG )
			{
				Logger.logInfo( "CellHandle " + this.toString() + " getSpan() for range: " + mergerange.toString() );
			}
			int[] ret = { 0, 0 };
			// if(check.toString().equals(this.toString())){ //it's the first in
			// the range -- show it!
			try
			{
				ret[0] = mergerange.getRows().length;
				ret[1] = mergerange.getCols().length; // TODO: test!
			}
			catch( Exception e )
			{
				Logger.logWarn( "CellHandle getting CellSpan failed: " + e );
			}
			// }
			return ret;
		}
		return null;
	}

	/**
	 * returns the WorkBookHandle for this Cell
	 *
	 * @return WorkBook the book
	 */
	public WorkBook getWorkBook()
	{
		return wbh;
	}

	/**
	 * get the index of the WorkSheet containing this Cell in the list of sheets
	 *
	 * @return int the WorkSheetHandle index for this Cell
	 */
	public int getSheetNum()
	{
		return this.mycell.getSheet().getSheetNum();
	}

	/**
	 * get the WorkSheetHandle for this Cell
	 *
	 * @return the WorkSheetHandle for this Cell
	 */
	public WorkSheetHandle getWorkSheetHandle()
	{
		return wsh;
	}

	/**
	 * Determines if the cellHandle represents a completely blank/null cell, and
	 * can thus be ignored for many operations.
	 * <p/>
	 * Criteria for returning true is a cell type of BLANK, that has a default
	 * format id (0), is not part of a merge range, does not contain a URL, and
	 * is not a part of a validation
	 *
	 * @return true if cell is truly blank
	 */
	public boolean isDefaultCell()
	{
		if( (this.getCellType() == CellHandle.TYPE_BLANK) && (((this.getFormatId() == 15) && !this.wbh.getWorkBook()
		                                                                                              .getIsExcel2007()) || (this.getFormatId() == 0)) && (this.getMergedCellRange() == null) && (this.getURL() == null) && (this.getValidationHandle() == null) )
		{
			return true;
		}
		return false;
	}

	/**
	 * set the WorkSheetHandle for this Cell
	 *
	 * @param WorkSheetHandle handle - the new worksheet for this Cell
	 * @see WorkSheetHandle
	 */
	public void setWorkSheetHandle( WorkSheetHandle handle )
	{
		wsh = handle;

		// This is redundant, already done in WSH.getCell().

		// if (wsh!=null) //20080616 KSC
		// wsh.cellhandles.put(this.getCellAddress(), this);
	}

	/**
	 * Returns an XML representation of the cell and it's component data
	 *
	 * @return String of XML
	 */
	public String getXML()
	{
		return getXML( null );
	}

	/**
	 * Returns an XML representation of the cell and it's component data
	 *
	 * @param int[] mergedRange - include merged ranges in the XML
	 *              representation if not null
	 * @return String of XML
	 */
	public String getXML( int[] mergedRange )
	{
		String vl = "", fvl = "", sv = "", hd = "", csp = "", hlink = "";
		Object val = null;
		StringBuffer retval = new StringBuffer();
		String typename = this.getCellTypeName();
		// put the formula string in
		if( typename.equals( "Formula" ) )
		{
			try
			{
				FormulaHandle fmh = getFormulaHandle();
				String fms = fmh.getFormulaString();
				// use single quotes around formula value to avoid errors in
				// xslt transform
				if( fms.indexOf( "\"" ) > 0 )
				{
					fvl = " Formula='" + StringTool.convertXMLChars( fms ) + "'";
				}
				else
				{
					fvl = " Formula=\"" + StringTool.convertXMLChars( fms ) + "\"";
				}
				try
				{
					if( this.wbh.getWorkBook().getCalcMode() != WorkBook.CALCULATE_EXPLICIT )
					{
						val = fmh.calculate();
					}
					else
					{
						try
						{
							// changed from getVal() now that getVal returns a
							// null
							val = fmh.getStringVal();
						}
						catch( Exception e )
						{
							Logger.logWarn( "CellHandle.getXML formula calc failed: " + e.toString() );
						}
					}
					if( val instanceof Float )
					{
						typename = "Float";
					}
					else if( val instanceof Double )
					{
						typename = "Double";
					}
					else if( val instanceof Integer )
					{
						typename = "Integer";
					}
					else if( (val instanceof java.util.Date) || (val instanceof java.sql.Date) || (val instanceof java.sql.Timestamp) )
					{
						typename = "DateTime";
					}
					else
					{
						typename = "String";
					}
				}
				catch( Exception e )
				{
					typename = "String"; // default
				}
			}
			catch( Exception e )
			{
				Logger.logErr( "ExtenXLS.getXML() failed getting type of Formula for: " + this.toString(), e );
			}
		}
		if( this.isDate() )
		{
			typename = "DateTime"; // 20060428 KSC: Moved after Formula parsing
		}

		// TODO: when RowHandle.getCells actually contains ALL cells, keep this
		if( this.mycell.getOpcode() != XLSConstants.MULBLANK )
		{
			// Put the style ID in
			sv = " StyleID=\"s" + getFormatId() + "\"";
			if( mergedRange != null )
			{ // TODO: fix!
				csp += " MergeAcross=\"" + (mergedRange[3] - mergedRange[1]) + "\"";
				csp += " MergeDown=\"" + (mergedRange[2] - mergedRange[0]) + "\"";
			}
			if( this.getCol().isHidden() )
			{
				hd = " Hidden=\"true\"";
			}
			// TODO: HRefScreenTip ????
			if( this.getURL() != null )
			{
				hlink = " HRef=\"" + StringTool.convertXMLChars( this.getURL() ) + "\"";
			}

			// put the date formattingin
			// Assemble the string
			retval.append( "<Cell Address=\"" + this.getCellAddress() + "\"" + sv + csp + fvl + hd + hlink + "><Data Type=\"" + typename + "\">" );
			if( typename.equals( "DateTime" ) )
			{
				retval.append( DateConverter.getFormattedDateVal( this ) );
			}
			else if( this.getCellType() == CellHandle.TYPE_STRING )
			{
				val = this.getStringVal(); // (String)getVal();
				if( val.equals( "" ) )
				{ // 20070216 KSC: John, had the same edits,
					// seems to work well in cursory tests ...
					// retval.append(" "); does this screw up formulas expecting
					// empty strings? -jm
				}
				else
				{
					retval.append( StringTool.convertXMLChars( val.toString() ) );
				}
			}
			else
			{
				try
				{
					// if(val == null)
					val = this.getVal();
					retval.append( StringTool.convertXMLChars( val.toString() ) + vl );
				}
				catch( Exception e )
				{
					Logger.logErr( "CellHandle.getXML failed for: " + this.getCellAddress() + " in: " + this.getWorkBook().toString(), e );
					retval.append( "XML ERROR!" );
				}
			}
			retval.append( "</Data>" );
			retval.append( end_cell_xml );
		}
		else
		{
			int c = ((Mulblank) this.mycell).getColFirst();
			int lastcol = ((Mulblank) this.mycell).getColLast();
			for(; c <= lastcol; c++ )
			{
				mulblankcolnum = (short) c;
				// Put the style ID in
				sv = " StyleID=\"s" + getFormatId() + "\"";
				if( this.getCol().isHidden() )
				{
					hd = " Hidden=\"true\"";
				}
				// TODO: HRefScreenTip ????
				if( this.getURL() != null )
				{
					hlink = " HRef=\"" + StringTool.convertXMLChars( this.getURL() ) + "\"";
				}

				// put the date formattingin
				// Assemble the string
				retval.append( "<Cell Address=\"" + this.getCellAddress() + "\"" + sv + csp + fvl + hd + hlink + "><Data Type=\"" + typename + "\"/>" );
				retval.append( end_cell_xml );
			}
		}
		return retval.toString();
	}

	/**
	 * Set the horizontal alignment for this Cell
	 *
	 * @param int align - constant value representing the horizontal alignment.
	 * @see FormatHandle.ALIGN* constants
	 */
	public void setHorizontalAlignment( int align )
	{
		setFormatHandle();
		formatter.setHorizontalAlignment( align );
	}

	/**
	 * Returns an int representing the current horizontal alignment in this
	 * Cell.
	 *
	 * @return int representing horizontal alignment
	 * @see FormatHandle.ALIGN* constants
	 */
	public int getHorizontalAlignment()
	{
		if( this.mycell.getXfRec() != null )
		{
			return this.mycell.getXfRec().getHorizontalAlignment();
		}
		// 0 is default alignment
		return 0;
	}

	/**
	 * Set the Vertical alignment for this Cell
	 *
	 * @param int align - constant value representing the vertical alignment.
	 * @see FormatHandle.ALIGN* constants
	 */
	public void setVerticalAlignment( int align )
	{
		setFormatHandle();
		formatter.setVerticalAlignment( align );
	}

	/**
	 * Returns an int representing the current vertical alignment in this Cell.
	 *
	 * @return int representing vertical alignment
	 * @see FormatHandle.ALIGN* constants
	 */
	public int getVerticalAlignment()
	{
		if( this.mycell.getXfRec() != null )
		{
			return this.mycell.getXfRec().getVerticalAlignment();
		}
		// 1 is default alignment
		return 1;
	}

	/**
	 * Sets the cell wrapping behavior for this cell
	 *
	 * @param boolean wrapit - true if cell text should be wrapped (default is
	 *                false)
	 */
	public void setWrapText( boolean wrapit )
	{
		setFormatHandle();
		formatter.setWrapText( wrapit );
		if( wrapit )
		{ // when wrap text it automatically wraps if row height has
			// NOT been set yet
			try
			{
				if( !this.getRow().isAlteredHeight() ) // has row height been altered??
				{
					this.getRow().setRowHeightAutoFit();
				}
			}
			catch( Exception e )
			{ /* ignore */
			}
		}
	}

	/**
	 * Get the cell wrapping behavior for this cell.
	 *
	 * @return true if cell text is wrapped
	 */
	public boolean getWrapText()
	{
		if( this.mycell.getXfRec() != null )
		{
			return this.mycell.getXfRec().getWrapText();
		}
		// false is default alignment
		return false;
	}

	/**
	 * Set the rotation of the cell in degrees. <br>
	 * Values 0-90 represent rotation up, 0-90degrees. <br>
	 * Values 91-180 represent rotation down, 0-90 degrees. <br>
	 * Value 255 is vertical
	 *
	 * @param int align - an int representing the rotation.
	 */
	public void setCellRotation( int align )
	{
		setFormatHandle();
		formatter.setCellRotation( align );
	}

	/**
	 * Get the rotation of this Cell in degrees. <br>
	 * Values 0-90 represent rotation up, 0-90degrees. <br>
	 * Values 91-180 represent rotation down, 0-90 degrees. <br>
	 * Value 255 is vertical
	 *
	 * @return int representing the degrees of cell rotation
	 */
	public int getCellRotation()
	{
		if( this.mycell.getXfRec() != null )
		{
			return this.mycell.getXfRec().getRotation();
		}
		// false is default alignment
		return 0;
	}

	@Override
	public int compareTo( CellHandle that )
	{
		int comp = this.getRowNum() - that.getRowNum();
		if( comp != 0 )
		{
			return comp;
		}
		return this.getColNum() - that.getColNum();
	}

	public boolean equals( Object that )
	{
		if( !(that instanceof CellHandle) )
		{
			return false;
		}
		return this.mycell.equals( ((CellHandle) that).mycell );
	}

	public int hashCode()
	{
		return this.mycell.hashCode();
	}

	/**
	 * Set the super/sub script for the Font
	 *
	 * @param int ss - super/sub script constant (0 = none, 1 = super, 2 = sub)
	 */
	public void setScript( int ss )
	{
		if( mycell.myxf == null )
		{
			this.getNewXf();
		}
		mycell.myxf.getFont().setScript( ss );
	}

	/**
	 * Set the val of the biffrec with an Object
	 *
	 * @param Object to set the value of the Cell to
	 */
	private void setBiffRecValue( Object obj ) throws CellTypeMismatchException
	{
		if( (mycell.getOpcode() == XLSConstants.BLANK) || (mycell.getOpcode() == XLSConstants.MULBLANK) )
		{
			// no reason for this Blank blank = (Blank)mycell;
			// String addr = mycell.getCellAddress();

			// trim the Mulblank
			/*
			 * KSC: mulblanks are NOT expanded now if (blank.getMyMul() !=
			 * null){ Mulblank mblank = (Mulblank)blank.getMyMul();
			 * mblank.trim(blank); }
			 */
			changeCellType( obj ); // 20080206 KSC: Basically replaces all above
			// code
		}
		else
		{
			if( obj == null )
			{
				// should never be false ??? if (!(mycell instanceof Blank))
				changeCellType( obj ); // will set to blank
			}
			else if( (obj instanceof Float) || (obj instanceof Double) || (obj instanceof Integer) || (obj instanceof Long) )
			{
				if( (mycell instanceof NumberRec) || (mycell instanceof Rk) )
				{
					if( obj instanceof Float )
					{
						Float f = (Float) obj;
						mycell.setFloatVal( f.floatValue() );
					}
					else if( obj instanceof Integer )
					{
						Integer i = (Integer) obj;
						mycell.setIntVal( i.intValue() );
					}
					else if( obj instanceof Double )
					{
						Double d = (Double) obj;
						mycell.setDoubleVal( d.doubleValue() );
					}
					else if( obj instanceof Long )
					{
						Long d = (Long) obj;
						mycell.setDoubleVal( d.longValue() );
					}
				}
				else
				{
					changeCellType( obj );
				}
			}
			else if( obj instanceof Boolean )
			{
				if( mycell instanceof Boolerr )
				{
					((Boolerr) mycell).setBooleanVal( ((Boolean) obj).booleanValue() );
				}
				else
				{
					changeCellType( obj );
				}
			}
			else if( obj instanceof String )
			{
				if( ((String) obj).startsWith( "=" ) )
				{
					changeCellType( obj ); // easier to just redo a formula...
				}
				else if( !obj.toString().equalsIgnoreCase( "" ) )
				{
					if( mycell instanceof Labelsst )
					{
						mycell.setStringVal( String.valueOf( obj ) );
					}
					else
					{
						changeCellType( obj );
					}
				}
				else if( !(mycell instanceof Blank) )
				{
					changeCellType( obj );
				}
			}
		}
	}

	/**
	 * if object type doesn't match current mycell record, remove and add
	 * appropriate record type
	 *
	 * @param obj
	 */
	private void changeCellType( Object obj )
	{
		int[] rc = { mycell.getRowNumber(), mycell.getColNumber() };
		Boundsheet bs = mycell.getSheet();
		int oldXf = mycell.getIxfe();
		bs.removeCell( mycell );
		BiffRec addedrec = bs.addValue( obj, rc, true );
		mycell = (XLSRecord) addedrec;
		mycell.setXFRecord( oldXf );
	}

	/**
	 * retrieves or creates a new xf for this cell
	 *
	 * @return
	 */
	private Xf getNewXf()
	{
		if( mycell.myxf != null )
		{
			return mycell.myxf;
		}
		// reusing or creating new xfs is handled in FormatHandle/cloneXf and
		// updateXf
		// this.useExistingXF = true; // flag to re-use this XF
		try
		{
			mycell.myxf = new Xf( this.getFont().getIdx() );
			// get the recidx of the last Xf
			int insertIdx = mycell.getWorkBook().getXf( mycell.getWorkBook().getNumXfs() - 1 ).getRecordIndex();
			// perform default add rec actions

			mycell.myxf.setSheet( null );
			mycell.getWorkBook().getStreamer().addRecordAt( mycell.myxf, insertIdx + 1 );
			mycell.getWorkBook().addRecord( (BiffRec) mycell.myxf, false );
			// update the pointer
			int xfe = mycell.myxf.getIdx();
			mycell.setIxfe( xfe );
			return mycell.myxf;
		}
		catch( Exception e )
		{
			return null;
		}
	}

	static final String begin_hidden_emptycell_xml = "<Cell Address=\"";
	static final String end_hidden_emptycell_xml = "\" StyleID=\"s15\" Hidden=\"true\"><Data Type=\"String\"></Data></Cell>";

	static final String begin_cell_xml = "<Cell Address=\"";
	static final String end_emptycell_xml = "\" StyleID=\"s15\"><Data Type=\"String\"></Data></Cell>";
	static final String end_cell_xml = "</Cell>";

	/**
	 * Returns an xml representation of an empty cell
	 *
	 * @param loc       - the cell address
	 * @param isVisible - if the cell is visible (not hidden)
	 * @return
	 */
	protected static String getEmptyCellXML( String loc, boolean isVisible )
	{
		if( !isVisible )
		{
			return begin_hidden_emptycell_xml + loc + end_hidden_emptycell_xml;
		}
		else
		{
			return begin_cell_xml + loc + end_emptycell_xml;
		}
	}

	/**
	 * Calculates and returns all formulas that reference this CellHandle. <br>
	 * Please note that these cells may have already been calculated, so in
	 * order to get their values without re-calculating them Extentech suggests
	 * setting the book level non-calculation flag, ie
	 * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT); or
	 * FormulaHandle.getCachedVal()
	 *
	 * @return List of of calculated cells (CellHandles)
	 */
	public List calculateAffectedCells()
	{
		ReferenceTracker rt = this.wbh.getWorkBook().getRefTracker();
		Iterator its = rt.clearAffectedFormulaCells( this ).values().iterator();

		List ret = new ArrayList();
		while( its.hasNext() )
		{
			CellHandle cx = new CellHandle( (BiffRec) its.next(), this.wbh );
			ret.add( cx );
		}
		return ret;
	}

	/**
	 * Internal method for clearing affected cells, does the same thing as
	 * calculateAffectedCells, but does not create a list
	 */
	protected void clearAffectedCells()
	{
		ReferenceTracker rt = this.wbh.getWorkBook().getRefTracker();
		rt.clearAffectedFormulaCells( this );
	}

	/**
	 * Calculates and returns all formulas on the same sheet that reference this
	 * CellHandle. <br>
	 * Please note that these cells may have already been calculated, so in
	 * order to get their values without re-calculating them Extentech suggests
	 * setting the book level non-calculation flag, ie
	 * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT); or
	 * FormulaHandle.getCachedVal()
	 *
	 * @return List of of calculated cells (CellHandles)
	 */
	public List<CellHandle> calculateAffectedCellsOnSheet()
	{
		Iterator its = this.wbh.getWorkBook()
		                       .getRefTracker()
		                       .clearAffectedFormulaCellsOnSheet( this, this.getWorkSheetName() )
		                       .values()
		                       .iterator();
		List<CellHandle> ret = new ArrayList<CellHandle>();
		while( its.hasNext() )
		{
			CellHandle cx = new CellHandle( (BiffRec) its.next(), this.wbh );
			ret.add( cx );
		}
		return ret;
	}

	/**
	 * flags chart references to the particular cell as dirty/
	 * needing caches rebuilt
	 */
	public void clearChartReferences()
	{
		ArrayList<ChartHandle> ret = new ArrayList();
		Iterator ii = this.wbh.getWorkBook().getRefTracker().getChartReferences( this.getCell() ).iterator();
		while( ii.hasNext() )
		{
			Ai ai = (Ai) ii.next();
			if( ai.getParentChart() != null )
			{
				ai.getParentChart().setMetricsDirty();
			}
		}
	}

	/**
	 * Creates a copy of this cell on the given worksheet at the given address.
	 *
	 * @param sourcecell the cell to copy
	 * @param newsheet   the sheet to which the cell should be copied
	 * @param row        the row in which the copied cell should be placed
	 * @param col        the row in which the copied cell should be placed
	 * @return CellHandle representing the new Cell
	 */
	public static final CellHandle copyCellToWorkSheet( CellHandle sourcecell, WorkSheetHandle newsheet, int row, int col ) throws
	                                                                                                                        CellPositionConflictException,
	                                                                                                                        CellNotFoundException
	{
		return copyCellToWorkSheet( sourcecell, newsheet, row, col, false );
	}

	/**
	 * Creates a copy of this cell on the given worksheet at the given address.
	 *
	 * @param sourcecell  the cell to copy
	 * @param newsheet    the sheet to which the cell should be copied
	 * @param row         the row in which the copied cell should be placed
	 * @param col         the row in which the copied cell should be placed
	 * @param copyByValue whether to copy formulas' values instead of the formulas
	 *                    themselves
	 * @return CellHandle representing the new cell
	 */
	public static CellHandle copyCellToWorkSheet( CellHandle sourcecell, WorkSheetHandle newsheet, int row, int col, boolean copyByValue )
	{
		// copy cell values
		CellHandle newcell = null;

		int offsets[] = {
				row - sourcecell.getRowNum(), col - sourcecell.getColNum()
		};

		if( sourcecell.isFormula() && !copyByValue )
		{
			try
			{
				FormulaHandle fmh = sourcecell.getFormulaHandle();
				newcell = newsheet.add( fmh.getFormulaString(), row, col );
				FormulaHandle fm2 = newcell.getFormulaHandle();
				FormulaHandle.moveCellRefs( fm2, offsets );
			}
			catch( FormulaNotFoundException ex )
			{
				newcell = null;
			}
		}

		if( newcell == null )
		{
			newcell = newsheet.add( sourcecell.getVal(), row, col );
		}

		return copyCellHelper( sourcecell, newcell );
	}

	/**
	 * Create a copy of this Cell in another WorkBook
	 *
	 * @param sourcecell the cell to copy
	 * @param target     worksheet to copy this cell into
	 * @return
	 */
	public static final CellHandle copyCellToWorkSheet( CellHandle sourcecell, WorkSheetHandle newsheet )
	{
		// copy cell values
		CellHandle newcell = null;
		try
		{
			FormulaHandle fmh = sourcecell.getFormulaHandle();
			// Logger.logInfo("testFormats Formula encountered:  "+
			// fmh.getFormulaString());

			newcell = newsheet.add( fmh.getFormulaString(), sourcecell.getCellAddress() );
		}
		catch( FormulaNotFoundException ex )
		{
			newcell = newsheet.add( sourcecell.getVal(), sourcecell.getCellAddress() );
		}
		return copyCellHelper( sourcecell, newcell );
	}

	/**
	 * Get the JSON object for this cell.
	 *
	 * @return String representing the JSON for this Cell
	 */
	public String getJSON()
	{
		return getJSONObject().toString();
	}

	/**
	 * Get the JSON object for this cell.
	 */
	public JSONObject getJSONObject()
	{
		CellRange cr = getMergedCellRange();
		int[] mergedCellRange = null;
		if( cr != null )
		{
			try
			{
				mergedCellRange = cr.getRangeCoords();
				if( mycell.getOpcode() == XLSRecord.MULBLANK )
				{
					Mulblank m = (Mulblank) mycell;
					if( !cr.contains( m.getIntLocation() ) )
					{
						mergedCellRange = null;
					}
				}
			}
			catch( CellNotFoundException e )
			{
				;
			}
		}

		return getJSONObject( mergedCellRange );
	}

	/**
	 * Get a JSON Object representation of a cell utilizing a
	 * merged range identifier.
	 *
	 * @deprecated The {@code mergedRange} parameter is unnecessary.
	 * This method will be removed in a future version.
	 * Use {@link #getJSONObject()} instead.
	 */
	@Deprecated
	public JSONObject getJSONObject( int[] mergedRange )
	{
		JSONObject theCell = new JSONObject();
		try
		{
			theCell.put( JSON_LOCATION, getCellAddress() );

			Object val;
			try
			{
				val = getVal();
				if( val == null )
				{
					val = "";
				}
			}
			catch( Exception ex )
			{
				Logger.logWarn( "ExtenXLS.getJSONObject failed: " + ex.toString() );
				val = "#ERR!";
			}

			String typename = getCellTypeName();
			JSONObject dataval = new JSONObject();

			if( typename.equals( "Formula" ) )
			{
				try
				{
					FormulaHandle fmh = getFormulaHandle();
					String fms = fmh.getFormulaString();

					theCell.put( JSON_CELL_FORMULA, fms );

					try
					{
						if( Float.parseFloat( val.toString() ) == (Float.NaN) )
						{
							typename = JSON_FLOAT;
						}
						else if( val instanceof Float )
						{
							typename = JSON_FLOAT;
						}
						else if( val instanceof Double )
						{
							typename = JSON_DOUBLE;
						}
						else if( val instanceof Integer )
						{
							typename = JSON_INTEGER;
						}
						else if( (val instanceof java.util.Date) ||
								(val instanceof java.sql.Date) ||
								(val instanceof java.sql.Timestamp) )
						{
							typename = JSON_DATETIME;
						}
						else
						{
							typename = JSON_STRING;
						}
					}
					catch( Exception e )
					{
						typename = JSON_STRING; // default
					}
				}
				catch( Exception e )
				{
					Logger.logErr( "ExtenXLS.getJSON() failed getting type of Formula for: " + toString(), e );
				}
			}

			if( isDate() )
			{
				typename = JSON_DATETIME;
			}

			dataval.put( JSON_TYPE, typename );

			// TODO: Handle Conditional Format
			// cell should return the style id for its condition
			// this is an ID that begins incrementing after the last Xf
			// and should be contained in the CSS for the output

			// if the conditional format evaluates to TRUE
			// then we use *that* style ID not the default

			// We can have multiple CF styles per cell, one per each rule... we'll need that from CSS standpoint so...

			// style
			theCell.put( JSON_STYLEID, getConditionalFormatId() );

			// merges
			if( mergedRange != null )
			{
				theCell.put( JSON_MERGEACROSS, (mergedRange[3] - mergedRange[1]) );
				theCell.put( JSON_MERGEDOWN, (mergedRange[2] - mergedRange[0]) );
				if( isMergeParent() )
				{
					theCell.put( JSON_MERGEPARENT, true );
				}
				else
				{
					theCell.put( JSON_MERGECHILD, true );
				}
			}

			// handle hidden setting
			try
			{
				if( getCol().isHidden() )
				{
					theCell.put( JSON_HIDDEN, true );
				}
			}
			catch( Exception e )
			{
				;
			}

			// handle the locked/formula hidden setting
			// only active if sheet is protected
			try
			{
				if( isFormulaHidden() )
				{
					theCell.put( JSON_FORMULA_HIDDEN, true );
				}

				theCell.put( JSON_LOCKED, isLocked() );
			}
			catch( Exception e )
			{
				;
			}

			try
			{
				ValidationHandle vh = getValidationHandle();
				if( vh != null )
				{
					theCell.put( JSON_VALIDATION_MESSAGE, vh.getPromptBoxTitle() + ":" + vh.getPromptBoxText() );
				}
			}
			catch( Exception e )
			{
				;
			}

			// hyperlinks
			if( !(getURL() == null) )
			{
				theCell.put( JSON_HREF, getURL() );
			}

			if( getWrapText() )
			{
				theCell.put( JSON_WORD_WRAP, true );
			}

			// store alignment for container issues
			int alignment = getFormatHandle().getHorizontalAlignment();
			if( alignment == FormatHandle.ALIGN_RIGHT )
			{
				theCell.put( JSON_TEXT_ALIGN, "right" );
			}
			else if( alignment == FormatHandle.ALIGN_CENTER )
			{
				theCell.put( JSON_TEXT_ALIGN, "center" );
			}
			else if( alignment == FormatHandle.ALIGN_LEFT )
			{
				theCell.put( JSON_TEXT_ALIGN, "left" );
			}

			// dates
			if( typename.equals( JSON_DATETIME ) && !(val == null) && !val.equals( "" ) )
			{
				dataval.put( JSON_CELL_VALUE, getFormattedStringVal() );
				//dataval.put(JSON_DATEVALUE, ch.getFloatVal());
				dataval.put( "time", DateConverter.getCalendarFromCell( this ).getTimeInMillis() );
			}
			else if( getCellType() == CellHandle.TYPE_STRING )
			{
				// FORCES CALC
				if( ((String) val).indexOf( "\n" ) > -1 )
				{
					val = ((String) val).replaceAll( "\n", "<br/>" );
				}
				if( !val.equals( "" ) )
				{
					dataval.put( JSON_CELL_VALUE, val.toString() );
				}
			}
			else
			{ // other
				dataval.put( JSON_CELL_VALUE, val.toString() );
				try
				{ // formatted pattern
					String s = getFormatPattern();
					if( !(s.equals( "" )) )
					{
						String fmtd = getFormattedStringVal(); // TRIGGERS CALC!
						if( !(val.equals( fmtd )) )
						{
							dataval.put( JSON_CELL_FORMATTED_VALUE, fmtd );
						}
						if( s.indexOf( "Red" ) > -1 )
						{
							Double d = new Double( val.toString() );
							if( d.doubleValue() < 0 )
							{
								theCell.put( JSON_RED_FORMAT, "1" );
								if( fmtd.indexOf( "-" ) == 0 )
								{
									fmtd = fmtd.substring( 1, fmtd.length() );
								}
								dataval.put( JSON_CELL_FORMATTED_VALUE, fmtd );
							}
						}
					}
				}
				catch( Exception x )
				{
				}
				;
			}
			theCell.put( JSON_DATA, dataval );
		}
		catch( JSONException e )
		{
			Logger.logErr( "error getting JSON for the cell: " + e );
		}
		return theCell;
	}

	/**
	 * Returns the validation handle for the cell.
	 *
	 * @return ValidationHandle for this Cell, or null if none
	 */
	public ValidationHandle getValidationHandle()
	{
		ValidationHandle ret = null;
		try
		{
			ret = this.getWorkSheetHandle().getValidationHandle( this.getCellAddress() );
		}
		catch( Exception e )
		{
			; // somewhat normal?
		}
		return ret;
	}

	/**
	 * Copies all formatting - xf and non-xf (such as column width, hidden
	 * state) plus merged cell range from a sourcecell to a new cell (usually in
	 * a new workbook)
	 *
	 * @param sourcecell the cell to copy
	 * @param newcell    the cell to copy to
	 * @return
	 */
	protected static final CellHandle copyCellHelper( CellHandle sourcecell, CellHandle newcell )
	{
		// copy row height & attributes
		int rz = sourcecell.getRow().getHeight();
		newcell.getRow().setHeight( rz );
		if( sourcecell.getRow().isHidden() )
		{
			newcell.getRow().setHidden( true );
		}
		// copy col width & attributes
		int rzx = sourcecell.getCol().getWidth();
		newcell.getCol().setWidth( rzx );
		if( sourcecell.getCol().isHidden() )
		{
			newcell.getCol().setHidden( true );
			// Logger.logInfo("column " + rzx + " is hidden");
		}

		try
		{
			// copy merged ranges
			CellRange rng = sourcecell.getMergedCellRange();
			if( rng != null )
			{
				rng = new CellRange( rng.getRange(), newcell.getWorkBook() );
				rng.addCellToRange( newcell );
				rng.mergeCells( false );
			}
			// Handle formats:
			Xf origxf = sourcecell.getWorkBook().getWorkBook().getXf( sourcecell.getFormatId() );
			newcell.getFormatHandle().addXf( origxf );
			return newcell;
		}
		catch( Exception ex )
		{
			Logger.logErr( "CellHandle.copyCellHelper failed.", ex );
		}
		return newcell;
	}

}
