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
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.CellRec;
import com.extentech.formats.XLS.ColumnNotFoundException;
import com.extentech.formats.XLS.Condfmt;
import com.extentech.formats.XLS.FormulaNotFoundException;
import com.extentech.formats.XLS.Mergedcells;
import com.extentech.formats.XLS.Name;
import com.extentech.formats.XLS.RowNotFoundException;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.formats.XLS.Xf;
import com.extentech.formats.XLS.formulas.FormulaParser;
import com.extentech.formats.XLS.formulas.GenericPtg;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import static com.extentech.ExtenXLS.JSONConstants.JSON_CELL;
import static com.extentech.ExtenXLS.JSONConstants.JSON_CELLS;
import static com.extentech.ExtenXLS.JSONConstants.JSON_CELL_VALUE;
import static com.extentech.ExtenXLS.JSONConstants.JSON_LOCATION;
import static com.extentech.ExtenXLS.JSONConstants.JSON_RANGE;

/**
 * Cell Range is a handle to a range of Workbook Cells
 * <p/>
 * <br>
 * Contains useful methods for working with Collections of Cells. <br>
 * <br>
 * for example: <br>
 * <br>
 * <blockquote> CellRange cr = new CellRange("Sheet1!A1:B10", workbk);<br>
 * cr.addCellToRange("C10");<br>
 * CellHandle mycells = cr.getCells();<br>
 * for(int x=0;x < mycells.length;x++) <br>
 * Logger.logInfo(mycells[x].getCellAddress() + mycells[x].toString());<br>
 * }<br>
 * <br>
 * </blockquote>
 * <p/>
 * "http://www.extentech.com">Extentech Inc.</a>
 *
 * @see DataBoundCellRange
 * @see XLSRecord
 */
public class CellRange implements Serializable
{

	Condfmt cfx = null;

	/**
	 * returns the conditional format object for this range, if any
	 *
	 * @return Condfmt object
	 */
	protected Condfmt getConditionalFormat()
	{
		return cfx;
	}

	/**
	 *
	 *
	 *
	 */
	private static final long serialVersionUID = -3609881364824289079L;
	private boolean ismerged = false;
	private BiffRec parent = null;    // if cell range is child of a named range, must ensure update correctly
	public static final boolean REMOVE_MERGED_CELLS = true;
	public static final boolean RETAIN_MERGED_CELLS = false;
	//private Ptg myptg = null;
	public boolean DEBUG = false;
	private boolean isDirty = false;    // true if addCellsToRange without init
	int firstcellrow = -1;
	int firstcellcol = -1;
	int lastcellrow = -1;
	int lastcellcol = -1;
	private boolean createBlanks = false;
	private boolean initializeCells = true;
	public transient CellHandle[] cells;
	protected transient String range, sheetname;
	protected transient com.extentech.ExtenXLS.WorkBook mybook;
	private transient WorkSheetHandle sheet = null;
	private int[] myrowints;
	private int[] mycolints;
	private transient RowHandle[] myrows;
	private transient ColHandle[] mycols;

	// for OOXML External References
	protected int externalLink1 = 0;
	protected int externalLink2 = 0;

	FormatHandle fmtr = null;
	boolean wholeCol = false, wholeRow = false;

	/**
	 * Protected constructor for creating result ranges.
	 */
	protected CellRange( WorkSheetHandle sheet, int row, int col, int width, int height )
	{
		this.sheet = sheet;
		sheetname = sheet.getSheetName();
		mybook = sheet.getWorkBook();

		firstcellrow = row;
		firstcellcol = col;
		lastcellrow = row + height - 1;
		lastcellcol = col + width - 1;

		range = sheetname + "!" + ExcelTools.formatRange( new int[]{
				firstcellcol, firstcellrow, lastcellcol, lastcellrow
		} );

		cells = new CellHandle[width * height];
	}

	/**
	 * Initializes a <code>CellRange</code> from a <code>CellRangeRef</code>.
	 * The source <code>CellRangeRef</code> instance must be qualified with a
	 * single resolved worksheet.
	 *
	 * @param source the <code>CellRangeRef<code> from which to initialize
	 * @param init   whether to populate the cell array
	 * @param create whether to fill gaps in the range with blank cells
	 * @throws IllegalArgumentException if the source range does not have a
	 *                                  resolved sheet or has more than one sheet
	 */
	public CellRange( CellRangeRef source, boolean init, boolean create )
	{
		initializeCells = init;
		createBlanks = create;

		sheet = source.getFirstSheet();
		if( (sheet == null) || source.isMultiSheet() )
		{
			throw new IllegalArgumentException( "the source range must have a single resolved sheet" );
		}

		mybook = this.sheet.getWorkBook();
		sheetname = this.sheet.getSheetName();

		// This is inefficient, but fixing it would require rewriting init.
		this.range = source.toString();

		try
		{
			this.init();
		}
		catch( CellNotFoundException e )
		{
			// this should be impossible
			throw new RuntimeException( e );
		}
	}

	/**
	 * Initializes a <code>CellRange</code> from a <code>CellRangeRef</code>.
	 * The source <code>CellRangeRef</code> instance must be qualified with a
	 * single resolved worksheet.
	 *
	 * @param source the <code>CellRangeRef<code> from which to initialize
	 * @throws IllegalArgumentException if the source range does not have a
	 *                                  resolved sheet or has more than one sheet
	 */
	public CellRange( CellRangeRef source )
	{
		this( source, false, true );
	}

	public void clearFormats()
	{
		for( int idx = 0; idx < cells.length; idx++ )
		{
			if( cells[idx] != null )
			{
				cells[idx].clearFormats();
			}
		}
	}

	/**
	 * @deprecated use clear()
	 */
	public void clearContents()
	{
		for( int idx = 0; idx < cells.length; idx++ )
		{
			if( cells[idx] != null )
			{
				cells[idx].clearContents();
			}
		}
	}

	/**
	 * clears the contents and formats of the cells referenced by this range
	 * but does not remove the cells from the workbook.
	 */
	public void clear()
	{
		for( int idx = 0; idx < cells.length; idx++ )
		{
			if( cells[idx] != null )
			{
				cells[idx].clear();
			}
		}
	}

	/**
	 * removes the cells referenced by this range from the sheet.
	 * <p/>
	 * NOTE: method does not shift rows or cols.
	 */
	public void removeCells()
	{
		for( int idx = 0; idx < cells.length; idx++ )
		{
			if( cells[idx] != null )
			{
				cells[idx].remove( true );
			}
		}
	}

	/**
	 * Un-Merge the Cells contained in this CellRange
	 *
	 * @throws Exception
	 */
	public void unMergeCells() throws Exception
	{
		BiffRec[] mycells = this.getCellRecs();
		for( int t = 0; t < mycells.length; t++ )
		{
			mycells[t].setMergeRange( null ); // unset the range of merged cells
			mycells[t].getXfRec().setMerged( false );
		}
		Mergedcells mc = this.getSheet().getSheet().getMergedCellsRec();
		if( mc != null )
		{
			mc.removeCellRange( this );
		}
		this.ismerged = false;
	}

	/**
	 * Set the format ID of all cells in this CellRange
	 * <br>FormatID can be obtained through any CellHandle with the getFormatID() call
	 *
	 * @param int fmtID - the format ID to set the cells within the range to
	 */
	public void setFormatID( int fmtID ) throws Exception
	{
		BiffRec[] mycells = this.getCellRecs();
		for( int t = 0; t < mycells.length; t++ )
		{
			mycells[t].setXFRecord( fmtID );
		}
	}

	/**
	 * Set a hyperlink on all cells in this CellRange
	 *
	 * @param String url - the URL String to set
	 */
	public void setURL( String url ) throws Exception
	{
		BiffRec[] mycells = this.getCellRecs();
		for( int t = 0; t < mycells.length; t++ )
		{
			new CellHandle( mycells[t], this.mybook ).setURL( url );
		}
	}

	/**
	 * Merge the Cells contained in this CellRange
	 *
	 * @param boolean remove - true to delete the Cells following the first in the range
	 */
	public void mergeCells( boolean remove )
	{
		this.createBlanks();
		if( remove )
		{
			this.mergeCellsClearFollowingCells();
		}
		else
		{
			this.mergeCellsKeepFollowingCells();
		}
	}

	/**
	 * Merge the Cells contained in this CellRange, clearing or removing
	 * the Cells following the first in the range
	 */
	private void mergeCellsClearFollowingCells()
	{
		BiffRec[] mycells = this.getCellRecs();

		// mycells[0].setMergeRange(this); // set the range of merged cells
		Xf r = mycells[0].getXfRec();
		if( r == null )
		{
			fmtr = new FormatHandle( this.getWorkBook() );
			fmtr.addCell( mycells[0] );
			r = mycells[0].getXfRec();
		}
		r.setMerged( true );
		for( int t = 0; t < mycells.length; t++ )
		{
			if( mycells[t] != null )
			{
				mycells[t].setSheet( this.getSheet().getMysheet() );
			}
			mycells[t].setMergeRange( this ); // set the range of merged cells
			if( t > 0 )
			{
				if( !(mycells[t] instanceof Blank) )
				{
					String cellname = mycells[t].getCellAddress();
					Boundsheet sheet = mycells[t].getSheet();
					mycells[t].remove( true ); // blow it out!
					sheet.addValue( null, cellname );
				}
			}
		}
		Mergedcells mc = this.getSheet().getSheet().getMergedCellsRec();
		if( mc == null )
		{
			mc = this.getSheet().getSheet().addMergedCellRec();
		}
		mc.addCellRange( this );
		this.ismerged = true;
	}

	/**
	 * Merge the Cells contained in this CellRange	 *
	 */
	private void mergeCellsKeepFollowingCells()
	{
		BiffRec[] mycells = this.getCellRecs();
		for( int t = 0; t < mycells.length; t++ )
		{
			mycells[t].setMergeRange( this ); // set the range of merged cells
			Xf r = mycells[t].getXfRec();
			if( r == null )
			{
				fmtr = new FormatHandle( this.getWorkBook() );
				fmtr.addCellRange( this );
				r = mycells[t].getXfRec();
			}
			r.setMerged( true );
		}

		Mergedcells mc = this.getSheet().getSheet().getMergedCellsRec();
		if( mc == null )
		{
			mc = this.getSheet().getSheet().addMergedCellRec();
		}
		mc.addCellRange( this );
		this.ismerged = true;
	}

	/**
	 * Gets the number of columns in the range.
	 */
	public int getWidth()
	{
		return (lastcellcol - firstcellcol) + 1;
	}

	/**
	 * Gets the number of rows in the range.
	 */
	public int getHeight()
	{
		return (lastcellrow - firstcellrow) + 1;
	}

	/**
	 * Returns an array of the row numbers referenced by this CellRange
	 *
	 * @return int[] array of row ints
	 */
	public int[] getRowInts()
	{
		if( myrowints != null )
		{
			return myrowints;
		}
		int numrows = (lastcellrow + 1) - firstcellrow;
		myrowints = new int[numrows];
		for( int t = 0; t < numrows; t++ )
		{
			myrowints[t] = firstcellrow + t;
		}
		return myrowints;
	}

	/**
	 * returns an array of column numbers referenced by this CellRange
	 *
	 * @return int[] array of col ints
	 */
	public int[] getColInts()
	{
		if( mycolints != null )
		{
			return mycolints;
		}
		int numcols = (lastcellcol + 1) - firstcellcol;
		mycolints = new int[numcols];
		for( int t = 0; t < numcols; t++ )
		{
			mycolints[t] = firstcellcol + t;
		}
		return mycolints;
	}

	/**
	 * Returns an array of Rows (RowHandles) referenced by this CellRange
	 *
	 * @return RowHandle[] array of row handles
	 */
	public RowHandle[] getRows() throws RowNotFoundException
	{
		if( myrows != null )
		{
			return myrows;
		}
		int numrows = (lastcellrow + 1) - firstcellrow;
		myrows = new RowHandle[numrows];
		for( int t = 0; t < numrows; t++ )
		{
			RowHandle rx = null;
			try
			{
				rx = sheet.getRow( (firstcellrow - 1) + t );
			}
			catch( Exception x )
			{
				; // typically empty rows
			}
			myrows[t] = rx;
		}
		return myrows;
	}

	/**
	 * Get the number of rows that this CellRange encompasses
	 *
	 * @return
	 */
	public int getNumberOfRows()
	{
		int numRows = (lastcellrow + 1) - firstcellrow;
		return numRows;
	}

	/**
	 * Returns an array of Columns (ColHandles) referenced by this CellRange
	 *
	 * @return ColHandle[] array of columns handles
	 */
	public ColHandle[] getCols() throws ColumnNotFoundException
	{
		if( mycols != null )
		{
			return mycols;
		}
		int numcols = (lastcellcol + 1) - firstcellcol;
		mycols = new ColHandle[numcols];
		for( int t = 0; t < numcols; t++ )
		{
			mycols[t] = sheet.getCol( firstcellcol + t );
		}
		return mycols;
	}

	/**
	 * Get the number of columns that this CellRange encompasses
	 *
	 * @return
	 */
	public int getNumberOfCols()
	{
		int numCols = (lastcellcol + 1) - firstcellcol;
		return numCols;
	}

	/**
	 * returns edge status of the desired CellHandle within this CellRange
	 * ie: top, left, bottom, right
	 * <br>
	 * returns 0 or 1 for 4 sides
	 * <br>
	 * 1,1,1,1 is a single cell in a range 1,1,0,0 is on the top left edge of
	 * the range
	 *
	 * @param CellHandle ch -
	 * @param int        sz -
	 * @return int[] array representing the edge positions
	 */
	// TODO: documentation: Don't quite understand this!
	public int[] getEdgePositions( CellHandle ch, int sz )
	{
		int[] coords = { 0, 0, 0, 0 };
		// get the corners, check for 'edges'
		String adr = ch.getCellAddress();
		int[] rc = ExcelTools.getRowColFromString( adr );
		// increment to one-based
		rc[0]++;
		if( rc[0] == firstcellrow )
		{
			coords[0] = sz;
		}
		if( rc[0] == lastcellrow )
		{
			coords[2] = sz;
		}
		if( rc[1] == firstcellcol )
		{
			coords[1] = sz;
		}
		if( rc[1] == lastcellcol )
		{
			coords[3] = sz;
		}
		return coords;
	}

	/**
	 * returns whether this CellRange intersects with another CellRange
	 *
	 * @param CellRange cr - CellRange to test
	 * @return boolean true if CellRange cr intersects with this CellRange
	 */
	public boolean intersects( CellRange cr )
	{
		// get the corners, check for 'contains'
		try
		{
			int[] rc = cr.getRangeCoords();
			if( (rc[0] >= firstcellrow) && (rc[2] <= lastcellrow) && (rc[1] >= firstcellcol) && (rc[3] <= lastcellcol) )
			{
				return true;
			}
		}
		catch( CellNotFoundException e )
		{
			Logger.logWarn( "CellRange unable to determine intersection of range: " + cr.toString() );
		}
		return false;
	}

	/**
	 * returns whether this CellRange contains a particular Cell
	 *
	 * @param CellHandle ch - the Cell to check
	 * @return true if CellHandle ch is contained within this CellRange
	 */
	public boolean contains( Cell cxx )
	{
		String chsheet = cxx.getWorkSheetName();
		String mysheet = "";
		if( this.getSheet() != null )
		{
			mysheet = this.getSheet().getSheetName();
		}
		if( !chsheet.equalsIgnoreCase( mysheet ) )
		{
			return false;
		}
		String adr = cxx.getCellAddress();
		int[] rc = ExcelTools.getRowColFromString( adr );
		return contains( rc );
	}

	/**
	 * returns whether this CellRange contains the specified Row/Col coordinates
	 *
	 * @param int[] rc - row/col coordinates to test
	 * @return true if the coordinates are contained with this CellRange
	 */
	public boolean contains( int[] rc )
	{
		boolean ret = true;
		if( (rc[0] + 1) < firstcellrow )
		{
			ret = false;
		}
		if( (rc[0] + 1) > lastcellrow )
		{
			ret = false;
		}
		if( rc[1] < firstcellcol )
		{
			ret = false;
		}
		if( rc[1] > lastcellcol )
		{
			ret = false;
		}
		return ret;
	}

	/**
	 * returns the String representation of this CellRange
	 */
	public String toString()
	{
		return range;
	}

	/**
	 * Constructor to create a new CellRange from a WorkSheetHandle and a set of
	 * range coordinates:
	 * <br>coords[0] = first row
	 * <br>coords[1] = first col
	 * <br>coords[2] = last row
	 * <br>coords[3] = last col
	 *
	 * @param WorkSheetHandle sht - handle to the WorkSheet containing the Range's Cells
	 * @param int[]           coords - the cell coordinates
	 * @param boolean         cb - true if should create blank cells
	 * @throws Exception
	 */
	public CellRange( WorkSheetHandle sht, int[] coords, boolean cb ) throws Exception
	{
		this.createBlanks = cb;
		this.sheet = sht;
		this.mybook = sht.wbh;
		sheetname = sht.getSheetName();
		sheetname = GenericPtg.qualifySheetname( sheetname );
		String addr = sheetname + "!";
		String c1 = ExcelTools.getAlphaVal( coords[1] ) + String.valueOf( coords[0] + 1 );
		String c2 = ExcelTools.getAlphaVal( coords[3] ) + String.valueOf( coords[2] + 1 );
		addr += c1 + ":" + c2;
		this.range = addr;
		this.init();
	}

	/**
	 * Set this CellRange to be the current Print Area
	 */
	public void setAsPrintArea()
	{
		if( this.sheet == null )
		{
			Logger.logErr( "CellRange.setAsPrintArea() failed: " + this.toString() + " does not have a valid Sheet reference." );
			return;
		}
		sheet.setPrintArea( this );
	}

	/**
	 * Constructor to create a new CellRange from a WorkSheetHandle and a set of
	 * range coordinates:
	 * <br>coords[0] = first row
	 * <br>coords[1] = first col
	 * <br>coords[2] = last row
	 * <br>coords[3] = last col
	 *
	 * @param WorkSheetHandle sht - handle to the WorkSheet containing the Range's Cells
	 * @param int[]           coords - the cell coordinates
	 */
	public CellRange( WorkSheetHandle sht, int[] coords ) throws Exception
	{
		this( sht, coords, false );
	}

	/**
	 * Constructor to Create a new CellRange from a String range
	 * <br>
	 * The String range must be in the format Sheet!CR:CR
	 * <br>
	 * For Example, "Sheet1!C9:I19"
	 * <br>NOTE:
	 * You MUST Set the WorkBookHandle explicitly on this CellRange or it will generate
	 * NullPointerException when trying to access the Cells.
	 *
	 * @param String r - range String
	 * @see CellRange.setWorkBook
	 */
	public CellRange( String r )
	{
		this.range = r;
	}

	/**
	 * Increase the bounds of the CellRange by including the CellHandle.
	 * <br>
	 * These are the limitations and side-effects of this method:
	 * <br>
	 * - the Cell must be contiguous with the existing Range, ie: you can add a
	 * Cell which either increments the row or the column of the existing range
	 * by one.
	 * <br>
	 * - the Cell must be on the same sheet as the existing range.
	 * <br>
	 * - as a Cell Range is a 2 dimensional rectangle, expanding a multiple
	 * column range by adding a Cell to the end will include the logical Cells
	 * on the row in the range.
	 * <br>
	 * Some Examples:
	 * <br>
	 * <br>
	 * // simple one dimensional range expansion:
	 * <br>existing Range = A1:A20
	 * <br>addCellToRange(A21) new Range = A1:A21
	 * <br>
	 * <br>existing Range = A1:B20
	 * <br>addCellToRange(A21)
	 * <br>new Range = A1:B21 // note B20 is included automatically
	 * <br>
	 * <br>existing Range = A1:A20
	 * <br>addCellToRange(B1)
	 * <br>new Range = A1:B20 //note entire B column of cells are included automatically
	 *
	 * @param CellHandle ch - the Cell to add to the CellRange
	 */
	public boolean addCellToRange( CellHandle ch )
	{
		// check worksheet
		String sheetname = ch.getWorkSheetName();
		if( sheetname == null )
		{
			Logger.logWarn( "Cell " + ch.toString() + " NOT added to Range: " + this.toString() );
			return false;
		}
		if( !sheetname.equalsIgnoreCase( this.getSheet().getSheetName() ) )
		{
			Logger.logWarn( "Cell " + ch.toString() + " NOT added to Range: " + this.toString() );
			return false;
		}
		int[] rc = { ch.getRowNum(), ch.getColNum() };
		// increment to one-based
		rc[0]++;

		// check that it's at the beginning
		if( firstcellrow == -1 )
		{
			firstcellrow = rc[0];
		}
		if( firstcellcol == -1 )
		{
			firstcellcol = rc[1];
		}
		if( lastcellrow == -1 )
		{
			lastcellrow = rc[0];
		}
		if( lastcellcol == -1 )
		{
			lastcellcol = rc[1];
		}
		if( rc[0] < firstcellrow )
		{
			firstcellrow--;
		}
		if( rc[1] < firstcellcol )
		{
			firstcellcol--;
		}
		// check that it's at the end
		if( rc[0] > lastcellrow )
		{
			lastcellrow++;
		}
		if( rc[1] > lastcellcol )
		{
			lastcellcol++;
		}

		// myptg is never set so myptg access never happens... taking out
		//boolean addPtgInfo = false;
		//if (myptg != null && myptg instanceof PtgRef)
		//	addPtgInfo = true;
		// format the new range String
		String newrange = this.getSheet().getSheetName() + "!";
		String newcellrange = "";
		//if (addPtgInfo && !((PtgRef) myptg).isColRel())
		//	newcellrange += "$";
		newcellrange += ExcelTools.getAlphaVal( firstcellcol );
		//if (addPtgInfo && !((PtgRef) myptg).isRowRel())
		//	newcellrange += "$";
		newcellrange += String.valueOf( firstcellrow );
		newcellrange += ":";
		//if (addPtgInfo && !((PtgRef) myptg).isColRel())
		//	newcellrange += "$";
		newcellrange += ExcelTools.getAlphaVal( lastcellcol );
		//if (addPtgInfo && !((PtgRef) myptg).isColRel())
		//	newcellrange += "$";
		newcellrange += String.valueOf( lastcellrow );
		this.range = newrange + newcellrange;
		isDirty = true;

		/*if (addPtgInfo) {
			ReferenceTracker.updateAddressPerPolicy(myptg, newcellrange);
			return true;
		}*/

		if( (this.parent != null) && (this.parent.getOpcode() == XLSConstants.NAME) )
		{
			((Name) parent).setLocation( this.range );    // ensure named range expression is updated, as well as any formula references are cleared of cache
		}

		return false;
	}

	/**
	 * get the Cells in this cell range
	 *
	 * @return CellHandle[] all the Cells in this range
	 */
	public CellHandle[] getCells()
	{
		if( isDirty )
		{
			try
			{
				init();
			}
			catch( CellNotFoundException e )
			{
				;
			}
		}
		return cells;
	}

	/**
	 * Return a list of the cells in this cell range
	 *
	 * @return List of CellHandles
	 */
	public List<CellHandle> getCellList()
	{
		return Arrays.asList( cells );
	}

	/**
	 * get the underlying Cell Records in this range
	 * <br>NOTE: Cell Records are not
	 * a part of the public API and are not intended for use in client
	 * applications.
	 *
	 * @return BiffRec[] array of Cell Records
	 */
	public BiffRec[] getCellRecs()
	{
		CellHandle[] ch = this.getCells();
		BiffRec[] ret = new BiffRec[ch.length];
		for( int t = 0; t < ret.length; t++ )
		{
			if( ch[t] != null )
			{
				ret[t] = ch[t].getCell();
			}
		}
		return ret;
	}
	
	/*
	 * reset the underlying cell records in this range
	 * <br>NOTE: Cell Records are not
	 * a part of the public API and are not intended for use in client
	 * applications.
	 * @return BiffRec[] array of Cell Records
	 * 
	 *NOT USED AT THIS TIME
	public BiffRec[] resetCellRecs() {
		this.isDirty= true;
		return getCellRecs();
	}*/

	/**
	 * If the cells contain an incrementing value that can be transferred into an int,
	 * then return that value, else throw a NPE.  I'm sure there is a better exception to be thrown,
	 * but not sure what that is, and it doesn't make sense to return a value like -1 in these cases.
	 */
	public int getIncrementAmount() throws Exception
	{
		CellHandle[] ch = this.getCells();
		if( ch.length == 1 )
		{
			throw new Exception( "Cannot have increment with non-range cell" );
		}
		boolean initialized = false;
		int incAmount = 0;
		for( int i = 1; i < ch.length; i++ )
		{
			int value1 = ch[i - 1].getIntVal();
			int value2 = ch[i].getIntVal();
			if( !initialized )
			{
				incAmount = (value2 - value1);
				initialized = true;
			}
			else
			{
				if( (value2 - value1) != incAmount )
				{
					throw new Exception( "Inconsistent values across increment" );
				}
			}
		}
		if( !initialized )
		{
			throw new Exception( "Error determining increment" );
		}
		return incAmount;
	}

	/**
	 * Constructor which creates a new CellRange from a String range
	 * <br>
	 * The String range must be in the format Sheet!CR:CR
	 * <br>
	 * For Example, "Sheet1!C9:I19"
	 *
	 * @param String   range - the range string
	 * @param WorkBook bk
	 * @param boolean  createblanks - true if blank cells should be created if necessary
	 * @param boolean  initcells - true if cells should (be initialized)
	 */
	public CellRange( String range, com.extentech.ExtenXLS.WorkBook bk, boolean createblanks, boolean initcells )
	{
		createBlanks = createblanks;
		initializeCells = initcells;
		this.range = range;
		if( bk == null )
		{
			return;
		}
		this.setWorkBook( bk );
		try
		{
			this.init();
		}
		catch( CellNotFoundException e )
		{
			;
		}
	}

	/**
	 * Constructor which creates a new CellRange from a String range
	 * <br>
	 * The String range must be in the format Sheet!CR:CR
	 * <br>
	 * For Example, "Sheet1!C9:I19"
	 *
	 * @param String   range - the range string
	 * @param WorkBook bk
	 * @param boolean  createblanks - true if blank cells should be created (if necessary)
	 */
	public CellRange( String range, com.extentech.ExtenXLS.WorkBook bk, boolean c )
	{
		createBlanks = c;
		this.range = range;
		if( bk == null )
		{
			return;
		}
		this.setWorkBook( bk );
		try
		{
			if( !"".equals( this.range ) )
			{
				this.init();
			}
		}
		catch( CellNotFoundException e )
		{
			;
		}
		catch( NumberFormatException ne )
		{
			;    // happens upon !REF range
		}
	}

	/**
	 * sets the parent of this Cell range
	 * <br>Used Internally.  Not intended for the End User.
	 *
	 * @param b
	 */
	public void setParent( BiffRec b )
	{
		parent = b;
	}

	/**
	 * Re-sort all cells in this cell range according to the column.
	 * <p/>
	 * A custom comparator can be passed in, or the default one can be used with
	 * sort(String, boolean).
	 * <p/>
	 * Comparators will be passed 2 CellHandle objects for comparison.
	 * <p/>
	 * Collections.reverse will be called on the results if ascending is set to false;
	 *
	 * @param rowNumber  the 0 based (row 5 = 4) number of the row to be sorted upon
	 * @param comparator
	 * @throws RowNotFoundException
	 * @throws ColumnNotFoundException
	 */
	public void sort( int rownumber, Comparator<CellHandle> comparator, boolean ascending ) throws RowNotFoundException
	{
		this.createBlanks();
		ArrayList<CellHandle> sortRow = this.getCellsByRow( rownumber );
		Collections.sort( sortRow, comparator );
		if( !ascending )
		{
			Collections.reverse( sortRow );
		}
		// now we have sorted the array list, come up with a map to resort the rows.
		int[] coords = null;
		try
		{
			coords = this.getRangeCoords();
			// fix stupid wrong offsets;
			coords[0] = coords[0]--;
			coords[2] = coords[2]--;
		}
		catch( CellNotFoundException e1 )
		{
		}
		ArrayList<ArrayList> outputCols = new ArrayList<ArrayList>();
		for( int i = 0; i < sortRow.size(); i++ )
		{
			CellHandle cell = sortRow.get( i );
			ArrayList cells = null;
			try
			{
				cells = this.getCellsByCol( ExcelTools.getAlphaVal( cell.getColNum() ) );
			}
			catch( ColumnNotFoundException e )
			{
				// if there are no cells in this column ignore it
			}
			outputCols.add( cells );
		}
		this.removeCells();
		for( int i = coords[1]; i <= coords[3]; i++ )
		{
			ArrayList cells = outputCols.get( i - coords[1] );
			for( int x = 0; x < cells.size(); x++ )
			{
				CellHandle cell = (CellHandle) cells.get( x );
				Boundsheet bs = this.getSheet().getBoundsheet();
				cell.getCell().setCol( (short) i );
				bs.addCell( (CellRec) cell.getCell() );
			}
		}

	}

	/**
	 * Changes the cellRange to a createBlanks cellrange and re-initializes the range,
	 * creating the missing blanks.
	 */
	private void createBlanks()
	{
		if( !this.createBlanks )
		{
			this.createBlanks = true;
			this.initializeCells = true;
			try
			{
				this.init();
			}
			catch( CellNotFoundException e )
			{
			}
		}
	}

	/**
	 * Resort all cells in the range according to the rownumber passed in.
	 *
	 * @param rownumber the 0 based row number
	 * @param ascending
	 * @throws RowNotFoundException
	 * @throws ColumnNotFoundException
	 */
	public void sort( int rownumber, boolean ascending ) throws RowNotFoundException
	{
		Comparator<CellHandle> cp = new CellComparator();
		this.sort( rownumber, cp, ascending );
	}

	/**
	 * Re-sort all cells in this cell range according to the column.
	 * <p/>
	 * A custom comparator can be passed in, or the default one can be used with
	 * sort(String, boolean).
	 * <p/>
	 * Comparators will be passed 2 CellHandle objects for comparison.
	 * <p/>
	 * * Collections.reverse will be called on the results if ascending is set to false;
	 *
	 * @param columnName
	 * @param comparator
	 * @throws RowNotFoundException
	 * @throws ColumnNotFoundException
	 */
	public void sort( String columnName, Comparator comparator, boolean ascending ) throws ColumnNotFoundException
	{
		if( !this.createBlanks )
		{
			// we cannot have empty cells in this operation
			this.createBlanks = true;
			try
			{
				this.init();
			}
			catch( CellNotFoundException e )
			{
			}
		}
		ArrayList sortCol = this.getCellsByCol( columnName );
		Collections.sort( sortCol, comparator );
		if( !ascending )
		{
			Collections.reverse( sortCol );
		}
		// now we have sorted the array list, come up with a map to resort the rows.
		int[] coords = null;
		try
		{
			coords = this.getRangeCoords();
			// fix stupid wrong offsets;
			coords[0] = coords[0]--;
			coords[2] = coords[2]--;
		}
		catch( CellNotFoundException e1 )
		{
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		ArrayList<ArrayList<CellHandle>> outputRows = new ArrayList<ArrayList<CellHandle>>();
		for( int i = 0; i < sortCol.size(); i++ )
		{
			CellHandle cell = (CellHandle) sortCol.get( i );
			ArrayList<CellHandle> cells = null;
			try
			{
				cells = this.getCellsByRow( cell.getRowNum() );
			}
			catch( RowNotFoundException e )
			{
				// ignore if no cells available
			}
			outputRows.add( cells );
		}
		this.removeCells();
		for( int i = coords[0]; i <= coords[2]; i++ )
		{
			ArrayList cells = outputRows.get( i - coords[0] );
			for( int x = 0; x < cells.size(); x++ )
			{
				CellHandle cell = (CellHandle) cells.get( x );
				Boundsheet bs = this.getSheet().getBoundsheet();
				cell.getCell().setRowNumber( i - 1 );
				bs.addCell( (CellRec) cell.getCell() );
			}
		}
	}

	/**
	 * Resort all cells in the range according to the column passed in.
	 *
	 * @param columnName
	 * @param ascending
	 * @throws ColumnNotFoundException
	 * @throws RowNotFoundException
	 */
	@SuppressWarnings( "unchecked" )
	public void sort( String columnName, boolean ascending ) throws ColumnNotFoundException
	{
		Comparator cp = new CellComparator();
		this.sort( columnName, cp, ascending );
	}

	public static String xmlResponsePre = "<CellRange>";
	public static String xmlResponsePost = "</CellRange>";

	/**
	 * Return the XML representation of this CellRange object
	 *
	 * @return String of XML
	 */
	public String getXML()
	{
		StringBuffer sb = new StringBuffer( "<CellRange Range=\"" + this.getRange() + "\">" );
		// StringBuffer sb = new StringBuffer(xmlResponsePre);
		CellHandle[] cx = this.getCells();
		sb.append( "\r\n" );
		// append cellxml
		for( int t = 0; t < cx.length; t++ )
		{
			sb.append( cx[t].getXML() );
			sb.append( "\r\n" );
		}
		sb.append( xmlResponsePost );
		return sb.toString();
	}

	/**
	 * gets the array of Cells in this Name
	 * <p/>
	 * NOTE: this method variation also returns the Sheetname for the name record if not null.
	 * <p/>
	 * Thus this method is limited to use with 2D ranges.
	 *
	 * @param fragment whether to enclose result in NameHandle tag
	 * @return Cell[] all Cells defined in this Name
	 */
	public String getCellRangeXML( boolean fragment )
	{
		StringBuffer sbx = new StringBuffer();
		if( !fragment )
		{
			sbx.append( "<?xml version=\"1\" encoding=\"utf-8\"?>" );
		}
		sbx.append( "<CellRange Range=\"" + this.getRange() + "\">" );
		sbx.append( getXML() );
		sbx.append( "</NameHandle>" );
		return sbx.toString();
	}

	/**
	 * gets the array of Cells in this Name
	 * <p/>
	 * Thus this method is limited to use with 2D ranges.
	 *
	 * @param fragment whether to enclose result in NameHandle tag
	 * @return Cell[] all Cells defined in this Name
	 */
	public String getCellRangeXML()
	{
		StringBuffer sbx = new StringBuffer();
		sbx.append( "<?xml version=\"1\" encoding=\"utf-8\"?>" );
		sbx.append( getXML() );
		return sbx.toString();
	}

	/**
	 * Constructor which creates a new CellRange using an array of cells as it's constructor.
	 * <br>NOTE
	 * that if the array of cells you are adding is not a rectangle of data (ie
	 * [A1][B1][C1]) that you will have null cells in your cell range and
	 * operations on it may cause errors.
	 * <br>
	 * If you wish to populate a cell range that is not contiguous, consider the
	 * constructor CellRange(CellHandle[] newcells, boolean createblanks), which
	 * will populate null cells with blank records, allowing normal operations
	 * such as formatting, merging, etc.
	 *
	 * @param CellHandle[] newcells - the array of cells from which to create the new CellRange
	 * @throws CellNotFoundException
	 */
	public CellRange( CellHandle[] newcells ) throws CellNotFoundException
	{
		this.setWorkBook( newcells[0].getWorkBook() );
		this.sheet = newcells[0].getWorkSheetHandle();
		for( int x = 0; x < newcells.length; x++ )
		{
			this.addCellToRange( newcells[x] );
		}
		this.init();
	}

	/**
	 * create a new CellRange using an array of cells as it's constructor.
	 * <br>
	 * If you wish to populate a cell range that is not contiguous, set
	 * createblanks to true, which will populate null cells with blank records,
	 * allowing normal operations such as formatting, merging, etc.
	 *
	 * @param CellHandle[] newcells - the array of cells from which to create the new CellRange
	 * @param boolean      createblanks - true if should create blank cells if necesary
	 */
	public CellRange( CellHandle[] newcells, boolean createblanks ) throws CellNotFoundException
	{
		this.createBlanks = createblanks;
		this.setWorkBook( newcells[0].getWorkBook() );
		this.sheet = newcells[0].getWorkSheetHandle();
		for( int x = 0; x < newcells.length; x++ )
		{
			this.addCellToRange( newcells[x] );
		}
		this.init();
	}

	/**
	 * Constructor which creates a new CellRange from a String range
	 * <br>
	 * The String range must be in the format Sheet!CR:CR
	 * <br>
	 * For Example, "Sheet1!C9:I19"
	 *
	 * @param String   range - the range string
	 * @param WorkBook bk
	 * @throws CellNotFoundException
	 */
	public CellRange( String range, com.extentech.ExtenXLS.WorkBook bk ) throws CellNotFoundException
	{
		this( range, bk, true );
	}

	/**
	 * attach the workbook for this CellRange
	 *
	 * @param WorkBook bk
	 */
	public void setWorkBook( com.extentech.ExtenXLS.WorkBook bk )
	{
		this.mybook = bk;
	}

	/**
	 * Gets the coordinates of this cell range,
	 *
	 * @return int[5]: [0] first row (zero based, ie row 1=0), [1] first column, [2] last row (zero based, ie row 1=0),
	 * [3] last column, [4] number of cells in range
	 */
	public int[] getCoords() throws CellNotFoundException
	{
		int numrows = 0;
		int numcols = 0;
		int numcells = 0;
		int[] coords = new int[5];
		String temprange = range;
		String[] s = ExcelTools.stripSheetNameFromRange( temprange );
		temprange = s[1];
		// qualify sheet and reset range - necessary if sheetname with spaces is used in formula parsing 
		sheetname = GenericPtg.qualifySheetname( s[0] );
		if( (sheetname != null) && !sheetname.equals( "" ) )
		{
			if( s[2] == null )
			{
				this.range = sheetname + "!" + temprange;
			}
			else
			{
				s[2] = GenericPtg.qualifySheetname( s[2] );
				this.range = sheetname + ":" + s[2] + "!" + temprange;
			}
		}
		
		/*
		 * check for R1C1
		 */
		if( (temprange.indexOf( "R" ) == 0) &&
				(temprange.indexOf( "C" ) > 1) &&
				(Character.isDigit( temprange.charAt( temprange.indexOf( "C" ) - 1 ) )) )
		{

			int[] b = ExcelTools.getRangeRowCol( temprange );

			numrows = (b[2] - b[0]);
			if( numrows <= 0 )
			{
				numrows = 1;
			}

			numcols = (b[3] - b[2]);
			if( numcols <= 0 )
			{
				numcols = 1;
			}

			numcells = numrows * numcols;

			int[] retr = new int[5];
			retr[0] = b[0];
			retr[1] = b[1];
			retr[2] = b[2];
			retr[3] = b[3];
			retr[4] = numcells;
			return retr;
		}
		String startcell = "", endcell = "";
		int lastcolon = temprange.lastIndexOf( ":" );
		endcell = temprange.substring( lastcolon + 1 );
		if( lastcolon == -1 ) // no range
		{
			startcell = endcell;
		}
		else
		{
			startcell = temprange.substring( 0, lastcolon );
		}

		startcell = StringTool.strip( startcell, "$" );
		endcell = StringTool.strip( endcell, "$" );

		// get the first cell's coordinates
		int charct = startcell.length();
		while( charct > 0 )
		{
			if( !Character.isDigit( startcell.charAt( --charct ) ) )
			{
				charct++;
				break;
			}
		}
		String firstcellrowstr = startcell.substring( charct );
		firstcellrow = Integer.parseInt( firstcellrowstr );
		String firstcellcolstr = startcell.substring( 0, charct );
		firstcellcol = ExcelTools.getIntVal( firstcellcolstr );
		// get the last cell's coordinates
		charct = endcell.length();
		while( charct > 0 )
		{
			if( !Character.isDigit( endcell.charAt( --charct ) ) )
			{
				charct++;
				break;
			}
		}
		String lastcellrowstr = endcell.substring( charct );
		lastcellrow = Integer.parseInt( lastcellrowstr );
		String lastcellcolstr = endcell.substring( 0, charct );
		lastcellcol = ExcelTools.getIntVal( lastcellcolstr );
		numrows = (lastcellrow - firstcellrow) + 1;
		numcols = (lastcellcol - firstcellcol) + 1;

		numcells = numrows * numcols;
		if( numcells < 0 )
		{
			numcells *= -1; // handle swapped cells ie: "B1:A1"
		}

		coords[0] = firstcellrow - 1;
		coords[1] = firstcellcol;
		coords[2] = lastcellrow - 1;
		coords[3] = lastcellcol;
		coords[4] = numcells;
		if( ((firstcellrow < 0) && (lastcellrow < 0)) || (firstcellcol < 0) || (lastcellcol < 0) )
		{
			// not an error if it is a whole column or whole row range
			if( (firstcellcol == -1) && (lastcellcol == -1) )
			{
				// what should numcells be for wholerow?
				wholeRow = true;
			}
			else if( (firstcellrow == -1) && (lastcellrow == -1) )
			{
				// what should numcells be for wholecol?
				wholeCol = true;
			}
			else
			{
				Logger.logErr( "CellRange.getRangeCoords: Error in Range " + range );
			}
		}

		// trap OOXML external reference link, if any
		if( s[3] != null )
		{
			externalLink1 = Integer.valueOf( s[3].substring( 1, s[3].length() - 1 ) ).intValue();
		}
		if( s[4] != null )
		{
			externalLink2 = Integer.valueOf( s[4].substring( 1, s[4].length() - 1 ) ).intValue();
		}

		return coords;

	}

	/**
	 * Gets the coordinates of this cell range.
	 *
	 * @return int[5]: [0] first row, [1] first column, [2] last row,
	 * [3] last column, [4] number of cells in range
	 * @deprecated {@link #getCoords()} instead, which returns zero based values for rows.
	 */
	public int[] getRangeCoords() throws CellNotFoundException
	{
		int[] ordinalValues = this.getCoords();
		ordinalValues[0] += 1;
		ordinalValues[2] += 1;
		return ordinalValues;
	}

	/**
	 * Returns the WorkSheet referenced in this CellRange.
	 *
	 * @return WorkSheetHandle sheet referenced in this CellRange.
	 */
	public WorkSheetHandle getSheet()
	{
		return sheet;
	}

	/**
	 * initializes this CellRange
	 *
	 * @throws CellNotFoundException
	 */
	public void init() throws CellNotFoundException
	{
		if( !FormulaParser.isComplexRange( range ) )
		{
			int[] coords = this.getRangeCoords();
			int rowctr = coords[0];
			int firstcellcol = coords[1];
			int lastcellcol = coords[3];
			int numcells = coords[4];

			int cellctr = firstcellcol - 1;
			try
			{
				if( sheetname != null )
				{
					if( sheetname.equals( "" ) ) // 20080214 KSC - is this a good idea?
					// default to work sheet 0
					{
						sheetname = this.mybook.getWorkSheet( 0 ).getSheetName();
					}
				}
				if( sheetname == null )
				{
					if( this.sheet != null )
					{
						sheetname = this.sheet.getSheetName();
					}
					else
					{
						throw new IllegalArgumentException( "sheet name not specified: " + range );
					}
				}

				String s = sheetname;
				if( s != null )
				{
					// handle enclosing apostrophes which are added to PtgRefs
					if( s.charAt( 0 ) == '\'' )
					{
						s = s.substring( 1, s.length() );
						if( s.charAt( s.length() - 1 ) == '\'' )
						{
							s = s.substring( 0, s.length() - 1 );
						}
					}
				}
				sheet = mybook.getWorkSheet( s );
				// if wholerow or wholecol, don't gather cells
				if( this.wholeCol || this.wholeRow )
				{
					return;
				}
				cells = new CellHandle[numcells];
				boolean resetFastAdds = false;
				if( sheet.getFastCellAdds() && this.createBlanks )
				{
					resetFastAdds = true;
					sheet.setFastCellAdds( false );
				}
				for( int i = 0; i < numcells; i++ )
				{
					if( cellctr == lastcellcol )
					{// if its the end of the row,
						// increment row.
						cellctr = firstcellcol - 1;
						rowctr++;
					}
					++cellctr;
					try
					{
						// use caching 20080917 KSC: PROBLEM HERE [Claritas
						// BugTracker 1862]
						if( this.initializeCells )
						{
							cells[i] = sheet.getCell( rowctr - 1, cellctr, sheet.getUseCache() ); // 20080917 KSC: use cache
						}
						// setting instead of
						// defaulting to true);
					}
					catch( CellNotFoundException e )
					{
						if( this.createBlanks )
						{
							cells[i] = sheet.add( null, rowctr - 1, cellctr );
						}

					}
				}
				if( resetFastAdds )
				{
					sheet.setFastCellAdds( true );
				}
			}
			catch( WorkSheetNotFoundException e )
			{
				throw new IllegalArgumentException( e.toString() );
			}
		}
		else
		{    // gather cells for a complex range ...
			com.extentech.formats.XLS.formulas.PtgMemFunc pm = new com.extentech.formats.XLS.formulas.PtgMemFunc();
			XLSRecord b = new XLSRecord();
			b.setWorkBook( this.mybook.getWorkBook() );
			pm.setParentRec( b );
			try
			{
				pm.setLocation( range );
				Ptg[] p = pm.getComponents();
				java.util.ArrayList<CellHandle> cellsfromcomplexrange = new java.util.ArrayList<CellHandle>();
				for( int i = 0; i < p.length; i++ )
				{
					try
					{
						cellsfromcomplexrange.add( mybook.getCell( ((PtgRef) p[i]).getLocation() ) );
					}
					catch( CellNotFoundException e )
					{
						if( this.createBlanks )
						{
							cells[i] = sheet.add( null, p[i].getLocation() );
						}
					}
				}
				cells = new CellHandle[cellsfromcomplexrange.size()];
				cells = cellsfromcomplexrange.toArray( cells );
			}
			catch( Exception e )
			{
				throw new IllegalArgumentException( e.toString() );
			}

		}
		isDirty = false;
	}

	/**
	 * Initializes this <code>CellRange</code>'s cell list if necessary.
	 * This method is useful if this <code>CellRange</code> was created with
	 * <code>initCells</code> set to <code>false</code> and it is later
	 * necessary to retrieve the cell list.
	 *
	 * @param createBlanks whether missing cells should be created as blanks.
	 *                     If this is <code>false</code> they will appear in the cell list
	 *                     as <code>null</code>s.
	 */
	public void initCells( boolean createBlanks )
	{
		// If we don't need to do anything, return
		if( (initializeCells == true) && (createBlanks ? this.createBlanks : true) )
		{
			return;
		}

		this.initializeCells = true;
		this.createBlanks = createBlanks;

		try
		{
			this.init();
		}
		catch( CellNotFoundException e )
		{
			// This will never actually happen but we have to catch it anyway
			throw new Error();
		}
	}

	/**
	 * @return the workbook object attached to this CellRange
	 */
	public WorkBook getWorkBook()
	{
		return mybook;
	}

	/**
	 * gets whether this CellRange will add blank records to the WorkBook for any
	 * missing Cells contained within the range.
	 *
	 * @return true if should create blank records for missing Cells
	 */
	public boolean getCreateBlanks()
	{
		return createBlanks;
	}

	/**
	 * set whether this CellRange will add blank records to the WorkBook for any
	 * missing Cells contained within the range.
	 *
	 * @param boolean b - true if should create blank records for missing Cells
	 */
	public void setCreateBlanks( boolean b )
	{
		createBlanks = b;
	}

	/**
	 * Return the String representation of this range
	 *
	 * @return the String range
	 */
	public String getRange()
	{
		return range;
	}

	/**
	 * Return the String cell address of this range in R1C1 format
	 *
	 * @return String range in R1C1 format
	 */
	public String getR1C1Range()
	{
		String rc1x = "R";
		try
		{
			int[] rc1 = this.getRangeCoords();
			rc1x += rc1[0] + 1;        // rangecoords are already 1-based
			rc1x += "C" + rc1[1];
			rc1x += ":R" + (rc1[2] + 1);
			rc1x += "C" + rc1[3];

		}
		catch( CellNotFoundException e )
		{
			Logger.logErr( "CellRange.getR1C1Range failed", e );
		}
		return rc1x;
	}

	/**
	 * Sets the range of cells for this CellRange to a string range
	 *
	 * @param String rng - Range string
	 */
	public void setRange( String rng )
	{
		range = rng;
		try
		{
			this.init();
		}
		catch( CellNotFoundException e )
		{
			; // don't have to report anything
		}
	}

	/**
	 * sets a border around the range of cells
	 *
	 * @param int            width - line width
	 * @param int            linestyle - line style
	 * @param java.awt.Color colr - color of border line
	 */
	public void setBorder( int width, int linestyle, java.awt.Color colr )
	{
		CellHandle[] ch = getCells();
		for( int t = 0; t < ch.length; t++ )
		{
			int[] coords = getEdgePositions( ch[t], width );
			// create Excel border -- top, left, bottom, right
			if( coords[0] > 0 )
			{
				ch[t].setTopBorderLineStyle( (short) linestyle );
				ch[t].setBorderTopColor( colr );
			}
			if( coords[1] > 0 )
			{
				ch[t].setLeftBorderLineStyle( (short) linestyle );
				ch[t].setBorderLeftColor( colr );
			}
			if( coords[2] > 0 )
			{
				ch[t].setBottomBorderLineStyle( (short) linestyle );
				ch[t].setBorderBottomColor( colr );
			}
			if( coords[3] > 0 )
			{
				ch[t].setRightBorderLineStyle( (short) linestyle );
				ch[t].setBorderRightColor( colr );
			}
		}
	}

	/**
	 * update the CellRange when the underlying Cells change their location
	 *
	 * @return boolean true if the CellRange could be updated, false if there are no cells represented by this range
	 */
	public boolean update()
	{
		if( cells == null )
		{
			return false; // this is an invalid range -- see Mergedcells prob.
		}
		// jm

		// arbitrarily set the initial vals...
		if( cells[0] != null )
		{    // 20100106 KSC: if didn't create blanks it's possible that cells are null
			firstcellrow = cells[0].getRowNum() + 1;
			firstcellcol = cells[0].getColNum();
			lastcellrow = cells[0].getRowNum() + 1;
			lastcellcol = cells[0].getColNum();
			for( int t = 0; t < cells.length; t++ )
			{
				CellHandle cx = cells[t];
				// 20090901 KSC: apparently can be null
				if( cx != null )
				{
					this.addCellToRange( cx );
				}
			}
			this.myrowints = null;
			this.mycolints = null;
			return true;
		}
		else if( this.range != null )
		{// 20100106 KSC: handle ranges containing null cell[0] (i.e. ranges referencing cells not present and createBlanks==false)
			if( this.DEBUG )
			{
				Logger.logWarn( "CellRange.update:  trying to access blank cells in range " + this.range );
			}
			try
			{
				this.getRangeCoords();
				return true;
			}
			catch( CellNotFoundException e )
			{    // shouldn't
				return false;
			}
		}
		return false;    // return false if it doesn't have it's cells defined
	}

	/**
	 * returns the merged state of the CellRange
	 *
	 * @return true if this CellRange is merged
	 */
	public boolean isMerged()
	{
		return ismerged;
	}

	/**
	 * Sets the sheet reference for this CellRange.
	 *
	 * @param WorkSheetHandle aSheet
	 */
	public void setSheet( WorkSheetHandle aSheet )
	{
		this.sheet = aSheet;
		this.sheetname = aSheet.getSheetName();
	}

	/**
	 * Whether to copy the cell contents.
	 */
	public static final int COPY_CONTENTS = 0x01;

	/**
	 * Whether formulas should be copied.
	 * If this bit is not set the formula result will be copied instead.
	 */
	public static final int COPY_FORMULAS = 0x02;

	public static final int COPY_FORMATS = 0x0100;

	/**
	 * Copies this range to another location.
	 * At present only contents and complete formats may be copied.
	 *
	 * @param row  the topmost row of the target area
	 * @param col  the leftmost column of the target area
	 * @param what a set of flags determining what will be copied
	 * @return the destination range
	 */
	public CellRange copy( WorkSheetHandle sheet, int row, int col, int what )
	{
		CellRange result = new CellRange( sheet, row, col, this.getWidth(), this.getHeight() );

		int first_col = col;

		// note these are not currently used, see setting below
		boolean copy_contents = (what & COPY_CONTENTS) != 0;
		boolean copy_formulas = (what & COPY_FORMULAS) != 0;
		boolean copy_formats = (what & COPY_FORMATS) != 0;

		// set to true until this thing is fully implemented
		copy_formats = true;
		copy_formulas = true;

		int cur_row = cells[0].getRowNum();
		for( int idx = 0; idx < cells.length; idx++ )
		{
			CellHandle source = cells[idx];

			if( source.getRowNum() != cur_row )
			{
				cur_row = source.getRowNum();
				row++;
				col = first_col;
			}

			CellHandle target = null;
			try
			{
				target = sheet.getCell( row, col );
			}
			catch( CellNotFoundException e )
			{
			}

			int formatID;
			if( copy_formats )
			{
				formatID = source.getFormatId();
			}
			else if( target != null )
			{
				formatID = target.getFormatId();
			}
			else
			{
				formatID = sheet.getWorkBook().getWorkBook().getDefaultIxfe();
			}

			if( copy_contents )
			{
				Object value;

				if( copy_formulas && source.isFormula() )
				{
					try
					{
						value = source.getFormulaHandle().getFormulaString();
					}
					catch( FormulaNotFoundException e )
					{
						// This shouldn't happen; we known the formula exists.
						// If it does happen it indicates a bug in ExtenXLS,
						// thus we throw an Error.
						throw new Error( "formula cell has no Formula record", e );
					}
				}
				else
				{
					value = source.getVal();
				}

				target = sheet.add( value, row, col, formatID );

				if( target.isFormula() )
				{
					try
					{
						FormulaHandle.moveCellRefs( target.getFormulaHandle(), new int[]{
								row - source.getRowNum(), col - source.getColNum()
						} );
					}
					catch( FormulaNotFoundException e )
					{
					}
				}
			}

			else if( target == null )
			{
				target = sheet.add( null, row, col, formatID );
			}

			if( copy_formats )
			{
				target.setFormatId( formatID );
			}

			result.cells[idx] = target;
			col++;
		}

		return result;
	}

	/**
	 * Fills this range from the given cell.
	 *
	 * @param source    the cell whose attributes should be copied
	 *                  or <code>null</code> to copy from the first cell in the range
	 * @param what      a set of flags determining what will be copied
	 * @param increment the amount by which to increment numeric values
	 *                  or <code>NaN</code> for no increment
	 */
	public void fill( CellHandle source, int what, double increment )
	{
		if( null == source )
		{
			source = cells[0];
		}

		boolean copy_contents = (what & COPY_CONTENTS) != 0;
		boolean copy_formulas = (what & COPY_FORMULAS) != 0;
		boolean copy_formats = (what & COPY_FORMATS) != 0;

		int sourceRow = source.getRowNum();
		int sourceCol = source.getColNum();

		Object value = null;
		if( copy_contents )
		{
			if( copy_formulas && source.isFormula() )
			{
				try
				{
					value = source.getFormulaHandle().getFormulaString();
				}
				catch( FormulaNotFoundException e )
				{
					throw new Error( "formula cell has no Formula record", e );
				}
			}
			else
			{
				value = source.getVal();
			}
		}

		// if increment is set, ensure the value can be incremented
		if( !Double.isNaN( increment ) && !(copy_contents && (value instanceof Number)) )
		{
			throw new IllegalArgumentException( "cannot increment unless filling with a numeric value" );
		}

		for( int idx = 0; idx < cells.length; idx++ )
		{
			CellHandle target = cells[idx];

			// don't overwrite the source cell
			if( source.equals( target ) )
			{
				continue;
			}

			if( !Double.isNaN( increment ) )
			{
				value = ((Number) value).doubleValue() + increment;
			}

			int formatID = (copy_formats ? source : target).getFormatId();

			if( copy_contents )
			{
				cells[idx] = target = sheet.add( value, target.getRowNum(), target.getColNum(), formatID );

				if( target.isFormula() )
				{
					try
					{
						FormulaHandle.moveCellRefs( target.getFormulaHandle(), new int[]{
								target.getRowNum() - sourceRow, target.getColNum() - sourceCol
						} );
					}
					catch( FormulaNotFoundException e )
					{
					}
				}
			}

			// when adding Date values passing the format ID to sheet.add
			// doesn't set the format so we always set it here
			if( copy_formats )
			{
				target.setFormatId( formatID );
			}

		}
	}

	public Collection<CellHandle> calculateAffectedCellsOnSheet()
	{
		Set<CellHandle> affected = new HashSet<CellHandle>();
		for( CellHandle cell : cells )
		{
			if( cell != null )
			{
				affected.add( cell );
				affected.addAll( cell.calculateAffectedCellsOnSheet() );
			}
		}
		return affected;
	}

	/**
	 * return a JSON array of cell values for the given range
	 * <br>
	 * static version
	 *
	 * @param String   range - a string representation of the desired range of cells
	 * @param WorkBook wbh - the source WorkBook for the cell range
	 * @return JSONArray - a JSON representation of the desired cell range
	 */
	public static JSONArray getValuesAsJSON( String range, WorkBook wbh )
	{
		JSONArray rangeArray = new JSONArray();
		try
		{
			CellRange cr = new CellRange( range, wbh, true );
			for( int j = 0; j < cr.getCells().length; j++ )
			{
				rangeArray.put( cr.getCells()[j].getVal() );
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "Error obtaining CellRange " + range + " JSON: " + e );
		}
		return rangeArray;
	}

	/**
	 * Return a json object representing this cell range, entries
	 * contain only address and values for more compact space
	 *
	 * @param range
	 * @param wbh
	 * @return
	 */
	public JSONObject getBasicJSON()
	{
		try
		{
			JSONObject crObj = new JSONObject();
			crObj.put( JSON_RANGE, this.getRange() );
			JSONArray rangeArray = new JSONArray();
			// should this possibly be full
			CellHandle[] cells = this.getCells();
			for( int j = 0; j < cells.length; j++ )
			{
				JSONObject result = new JSONObject();
				String addy = cells[j].getCellAddress();
				String val = cells[j].getVal().toString();
				result.put( JSON_LOCATION, addy );
				result.put( JSON_CELL_VALUE, val );
				rangeArray.put( result );
			}
			crObj.put( JSON_CELLS, rangeArray );
			return crObj;
		}
		catch( Exception e )
		{
			Logger.logErr( "Error obtaining CellRange " + range + " JSON: " + e );
		}
		return new JSONObject();
	}

	/**
	 * Return a json object representing this cell range with full cell
	 * information embedded.
	 */
	public JSONObject getJSON()
	{
		JSONObject theRange = new JSONObject();
		JSONArray cells = new JSONArray();
		try
		{
			theRange.put( JSON_RANGE, getRange() );
			CellHandle[] chandles = getCells();
			for( int i = 0; i < chandles.length; i++ )
			{
				CellHandle thisCell = chandles[i];
				JSONObject result = new JSONObject();

				result.put( JSON_CELL, thisCell.getJSONObject() );
				cells.put( result );
			}
			theRange.put( JSON_CELLS, cells );
		}
		catch( JSONException e )
		{
			Logger.logErr( "Error getting cellRange JSON: " + e );
		}
		return theRange;
	}

	/**
	 * Get the cells from a particular rownumber, constrained by the boundaries of the cellRange
	 *
	 * @param rownumber
	 */
	public ArrayList<CellHandle> getCellsByRow( int rownumber ) throws RowNotFoundException
	{
		ArrayList<CellHandle> al = new ArrayList<CellHandle>();
		RowHandle r = this.getSheet().getRow( rownumber );
		CellHandle[] cells = r.getCells();
		int[] coords = null;
		try
		{
			coords = this.getRangeCoords();
		}
		catch( CellNotFoundException e )
		{
			throw new RowNotFoundException( "Error getting internal coordinates for CellRange" + e );
		}
		for( int i = 0; i < cells.length; i++ )
		{
			if( (cells[i].getColNum() >= coords[1]) && (cells[i].getColNum() <= coords[3]) )
			{
				al.add( cells[i] );
			}
		}
		return al;
	}

	/**
	 * Get the cells from a particular column, constrained by the boundaries of the cellRange
	 *
	 * @param rownumber
	 * @throws ColumnNotFoundException
	 */
	public ArrayList<CellHandle> getCellsByCol( String col ) throws ColumnNotFoundException
	{
		ArrayList<CellHandle> al = new ArrayList<CellHandle>();
		ColHandle r = this.getSheet().getCol( col );
		CellHandle[] cells = r.getCells();
		int[] coords = null;
		try
		{
			coords = this.getRangeCoords();
			coords[0] = coords[0] - 1;
			coords[2] = coords[2] - 1;
		}
		catch( CellNotFoundException e )
		{
			throw new ColumnNotFoundException( "Error getting internal coordinates for CellRange" + e );
		}
		for( int i = 0; i < cells.length; i++ )
		{
			if( (cells[i].getRowNum() >= coords[0]) && (cells[i].getRowNum() <= coords[2]) )
			{
				al.add( cells[i] );
			}
		}
		return al;
	}

	/**
	 * returns the cells for a given range
	 * <br>
	 * static version
	 *
	 * @param String   range - a string representation of the desired range of cells
	 * @param WorkBook wbh - the source WorkBook for the cell range
	 * @return CellHandle[] array of cells represented by the desired cell range
	 */
	public static CellHandle[] getCells( String range, WorkBookHandle wbh )
	{
		CellRange cr = new CellRange( range, wbh, true );
		return cr.getCells();
	}

	/**
	 * removes the border from all of the cells in this range
	 */
	public void removeBorder()
	{
		CellHandle[] ch = getCells();
		for( int t = 0; t < ch.length; t++ )
		{
			ch[t].removeBorder();
		}
	}

	/**
	 * Sets a bottom border on all cells in the cellrange
	 * <p/>
	 * Linestyle should be set through the FormatHandle constants
	 */
	public void setInnerBorderBottom( int linestyle, java.awt.Color colr )
	{
		CellHandle[] ch = getCells();
		for( int t = 0; t < ch.length; t++ )
		{
			ch[t].setBottomBorderLineStyle( (short) linestyle );
			ch[t].setBorderBottomColor( colr );
		}
	}

	/**
	 * Sets a right border on all cells in the cellrange
	 * <p/>
	 * Linestyle should be set through the FormatHandle constants
	 */
	public void setInnerBorderRight( int linestyle, java.awt.Color colr )
	{
		CellHandle[] ch = getCells();
		for( int t = 0; t < ch.length; t++ )
		{
			ch[t].setRightBorderLineStyle( (short) linestyle );
			ch[t].setBorderRightColor( colr );
		}
	}

	/**
	 * Sets a left border on all cells in the cellrange
	 * <p/>
	 * Linestyle should be set through the FormatHandle constants
	 */
	public void setInnerBorderLeft( int linestyle, java.awt.Color colr )
	{
		CellHandle[] ch = getCells();
		for( int t = 0; t < ch.length; t++ )
		{
			ch[t].setLeftBorderLineStyle( (short) linestyle );
			ch[t].setBorderLeftColor( colr );
		}
	}

	/**
	 * Sets a top border on all cells in the cellrange
	 * <p/>
	 * Linestyle should be set through the FormatHandle constants
	 */
	public void setInnerBorderTop( int linestyle, java.awt.Color colr )
	{
		CellHandle[] ch = getCells();
		for( int t = 0; t < ch.length; t++ )
		{
			ch[t].setTopBorderLineStyle( (short) linestyle );
			ch[t].setBorderTopColor( colr );
		}
	}

	/**
	 * Sets a surround border on all cells in the cellrange
	 * <p/>
	 * Linestyle should be set through the FormatHandle constants
	 */
	public void setInnerBorderSurround( int linestyle, java.awt.Color colr )
	{
		CellHandle[] ch = getCells();
		for( int t = 0; t < ch.length; t++ )
		{
			ch[t].setBorderColor( colr );
			ch[t].setBorderLineStyle( (short) linestyle );
		}
	}

}
