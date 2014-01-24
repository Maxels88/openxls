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

import com.extentech.formats.XLS.AutoFilter;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.CellPositionConflictException;
import com.extentech.formats.XLS.CellTypeMismatchException;
import com.extentech.formats.XLS.Cf;
import com.extentech.formats.XLS.Colinfo;
import com.extentech.formats.XLS.ColumnNotFoundException;
import com.extentech.formats.XLS.Condfmt;
import com.extentech.formats.XLS.Dv;
import com.extentech.formats.XLS.FeatHeadr;
import com.extentech.formats.XLS.FormulaNotFoundException;
import com.extentech.formats.XLS.ImageNotFoundException;
import com.extentech.formats.XLS.Mulblank;
import com.extentech.formats.XLS.Name;
import com.extentech.formats.XLS.Note;
import com.extentech.formats.XLS.Password;
import com.extentech.formats.XLS.ReferenceTracker;
import com.extentech.formats.XLS.Row;
import com.extentech.formats.XLS.RowNotFoundException;
import com.extentech.formats.XLS.SheetProtectionManager;
import com.extentech.formats.XLS.Sxview;
import com.extentech.formats.XLS.Unicodestring;
import com.extentech.formats.XLS.ValidationException;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.formats.XLS.charts.Chart;
import com.extentech.formats.XLS.charts.Series;
import com.extentech.formats.XLS.formulas.GenericPtg;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.formats.XLS.formulas.PtgRef;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.StringTool;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectOutputStream;
import java.io.OutputStream;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * The WorkSheetHandle provides a handle to a Worksheet within an XLS file<br>
 * and includes convenience methods for working with the Cell values within the sheet.<br>
 * <br>
 * for example:
 * <br>
 * <br>
 * <blockquote>
 * WorkBookHandle book = new WorkBookHandle("testxls.xls");<br>
 * WorkSheetHandle sheet1 = book.getWorkSheet("Sheet1");<br>
 * CellHandle cell = sheet1.getCell("B22");<br>
 * <p/>
 * <br>to add a cell:<br><br>
 * CellHandle cell = sheet1.add("Hello World","C22");<br>
 * <p/>
 * <br>to add a numeric cell:<br><br>
 * CellHandle cell = sheet1.add(Integer.valueOf(120),"C23");<br>
 * <p/>
 * <br>to add a formula cell:<br><br>
 * CellHandle cell = sheet1.add("=PI()","C24");<br>
 * <p/>
 * </blockquote>
 * <br>
 * <br>
 *
 * @see WorkSheet
 * @see WorkBookHandle
 * @see CellHandle
 */
public class WorkSheetHandle implements Handle
{
	private static final Logger log = LoggerFactory.getLogger( WorkSheetHandle.class );
	private Boundsheet mysheet;
	private WorkBook mybook;
	WorkBookHandle wbh;
	private Hashtable dateFormats = new Hashtable();
	private boolean cache = true;    //20080917 KSC: set var for caching, default to true [BugTracker 1862]
	// public Map cellhandles = new HashMap();

	public void addChart( byte[] serialchart, String name, short[] coords )
	{
		mysheet.addChart( serialchart, name, coords );
	}

	/**
	 * Get the first row on the Worksheet
	 *
	 * @return the Minimum Row Number on the Worksheet
	 */
	public int getFirstRow()
	{
		return mysheet.getMinRow();
	}

	/**
	 * Get the first column on the Worksheet
	 *
	 * @return the Minimum Column Number on the Worksheet
	 */
	public int getFirstCol()
	{
		return mysheet.getMinCol();
	}

	/**
	 * Get the last row on the Worksheet
	 *
	 * @return the Maximum Row Number on the Worksheet
	 */
	public int getLastRow()
	{
		return mysheet.getMaxRow();
	}

	/**
	 * Get the last column on the Worksheet
	 *
	 * @return the Maximum Column Number on the Worksheet
	 */
	public int getLastCol()
	{
		return mysheet.getMaxCol();
	}

	/**
	 * Sets whether the worksheet is protected.
	 * If <code>protect</code> is <code>true</code>, the worksheet will be
	 * protected and the password will be set to <code>password</code>. If it's
	 * <code>false</code>, the worksheet will be unprotected and the password
	 * will be removed.
	 *
	 * @param protect  whether the worksheet should be protected
	 * @param password the password to set if protect is <code>true</code>.
	 *                 ignored when protect is <code>false</code>.
	 * @throws WorkBookException never. This used to be thrown when
	 *                           unprotecting if the password was incorrect.
	 */
	public void setProtected( boolean protect, String password ) throws WorkBookException
	{
		SheetProtectionManager protector = mysheet.getProtectionManager();

		// we need to check if this password can be used to unprotect... 
		// otherwise it is totally insecure...
		String oldpass = protector.getPassword();

		Password pss = new Password();
		pss.setPassword( password );
		String passcheck = pss.getPasswordHashString();

		if( oldpass != null )
		{
			if( !oldpass.equals( passcheck ) && (oldpass != "0000") )
			{
				throw new WorkBookException( "Incorrect Password Attempt to Unprotect Worksheet.",
				                             WorkBookException.SHEETPROTECT_INCORRECT_PASSWORD );
			}
		}
		protector.setProtected( protect );
		protector.setPassword( protect ? password : null );
	}

	/**
	 * Sets whether the worksheet is protected.
	 *
	 * @param protect whether worksheet protection should be enabled
	 */
	public void setProtected( boolean protect )
	{
		mysheet.getProtectionManager().setProtected( protect );
	}

	/**
	 * Sets the password used to unlock the sheet when it is protected.
	 *
	 * @param password the clear text of the password to be applied
	 *                 or null to remove the existing password
	 */
	public void setProtectionPassword( String password )
	{
		mysheet.getProtectionManager().setPassword( password );
	}

	/**
	 * Sets the password used to unlock the sheet when it is protected.
	 * This method is useful in combination with
	 * {@link #getHashedProtectionPassword} to copy the password from one
	 * worksheet to another.
	 *
	 * @param hash the hash of the protection password to be applied
	 *             or null to remove the existing password
	 */
	public void setProtectionPasswordHashed( String hash )
	{
		mysheet.getProtectionManager().setPasswordHashed( hash );
	}

	/**
	 * Gets the hash of the sheet protection password.
	 * This method returns the hashed password as stored in the file. It has
	 * been passed through a one-way hash function. It is therefore not
	 * possible to recover the actual password. You can, however, use
	 * {@link #setProtectionPasswordHashed} to apply the same password to
	 * another worksheet.
	 *
	 * @return the password hash or "0000" if the sheet doesn't have a password
	 */
	public String getHashedProtectionPassword()
	{
		return mysheet.getProtectionManager().getPassword();
	}

	/**
	 * Returns whether the sheet is protected.
	 * Note that this is separate from whether the sheet has a protection
	 * password. It can be protected without a password or have a password but
	 * not be protected.
	 *
	 * @return whether protection is enabled for the sheet
	 */
	public boolean getProtected()
	{
		return mysheet.getProtectionManager().getProtected();
	}

	/**
	 * Checks whether the given password matches the protection password.
	 *
	 * @param guess the password to be checked against the stored hash
	 * @return whether the given password matches the stored hash
	 */
	public boolean checkProtectionPassword( String guess )
	{
		return mysheet.getProtectionManager().checkPassword( guess );
	}

	/**
	 * Sets the worksheet enhanced protection option
	 *
	 * @param int protectionOption
	 * @see WorkBookHandle.iprot options
	 */
	public void setEnhancedProtection( int protectionOption, boolean set )
	{
		mysheet.getProtectionManager().setProtected( protectionOption, set );
	}

	/**
	 * returns true if the indicated Enhanced Protection Setting is turned on
	 *
	 * @param protectionOption
	 * @return boolean true if the indicated Enhanced Protection Setting is turned on
	 * @see WorkBookHandle.iprot options
	 */
	public boolean getEnhancedProtection( int protectionOption )
	{
		return mysheet.getProtectionManager().getProtected( protectionOption );
	}

	/**
	 * set whether this sheet is VERY hidden opening the file.
	 * <p/>
	 * VERY hidden means users will not be able to unhide
	 * the sheet without using VB code.
	 *
	 * @param boolean b hidden state
	 */
	public void setVeryHidden( boolean b )
	{
		int h = 0;
		if( b )
		{
			h = Boundsheet.VERY_HIDDEN;
		}
		mysheet.setHidden( h );
		int t = mysheet.getSheetNum();
		try
		{ // set the next sheet selected...
			Boundsheet s2 = mybook.getWorkSheetByNumber( t + 1 );
			s2.setSelected( true );
		}
		catch( WorkSheetNotFoundException e )
		{
			;
		}
	}

	/**
	 * get whether this sheet is selected upon opening the file.
	 *
	 * @return boolean b selected state
	 */
	public boolean getSelected()
	{
		return mysheet.selected();
	}

	/**
	 * get whether this sheet is hidden from the user opening the file.
	 *
	 * @return boolean b hidden state
	 */
	public boolean getHidden()
	{
		return mysheet.getHidden();
	}

	/**
	 * return the 'veryhidden' state of the sheet
	 *
	 * @return
	 */
	public boolean getVeryHidden()
	{
		return mysheet.getVeryHidden();
	}

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
	public void setHidden( boolean b )
	{
		int h = 0;
		if( b )
		{
			h = Boundsheet.HIDDEN;
		}
		mysheet.setHidden( h );
		if( mysheet.getSheetNum() == 0 )
		{
			try
			{
				Boundsheet s2 = mybook.getWorkSheetByNumber( mysheet.getSheetNum() + 1 );
				mybook.setFirstVisibleSheet( s2 );
			}
			catch( WorkSheetNotFoundException e )
			{
				;
			}
		}
		if( mysheet.selected() )
		{
			try
			{ // set the next sheet selected...
				int x = 1;
				Boundsheet s2 = mybook.getWorkSheetByNumber( mysheet.getSheetNum() + x );
				while( s2.getHidden() )
				{
					s2 = mybook.getWorkSheetByNumber( mysheet.getSheetNum() + x++ );
				}
				s2.setSelected( true );
			}
			catch( WorkSheetNotFoundException e )
			{
				;
			}
		}
	}

	/**
	 * set this WorkSheet as the first visible tab on the left
	 */
	public void setFirstVisibleTab()
	{
		mysheet.getWorkBook().setFirstVisibleSheet( mysheet );
	}

	/**
	 * get the tab display order of this Worksheet
	 * <p/>
	 * this is a zero based index with zero representing
	 * the left-most WorkSheet tab.
	 *
	 * @return int idx the index of the sheet tab
	 */
	public int getTabIndex()
	{
		return mysheet.getSheetNum();
	}

	/**
	 * set the tab display order of this Worksheet
	 * <p/>
	 * this is a zero based index with zero representing
	 * the left-most WorkSheet tab.
	 *
	 * @param int idx the new index of the sheet tab
	 */
	public void setTabIndex( int idx )
	{
		mysheet.getWorkBook().changeWorkSheetOrder( mysheet, idx );
	}

	/**
	 * set whether this sheet is selected upon
	 * opening the file.
	 *
	 * @param boolean b selected value
	 */
	public void setSelected( boolean b )
	{
		mysheet.setSelected( b );
	}

	/**
	 * returns the ColHandle for the column at index position
	 * the column index is zero based ie: column A = 0
	 *
	 * @return ColHandle the Column
	 */
	public ColHandle getCol( int clnum ) throws ColumnNotFoundException
	{
		Colinfo ci = mysheet.getColInfo( clnum );
		ColHandle mycol;
		if( (ci == null) || !ci.isSingleCol() )
		{
			try
			{
				if( ci == null )
				{
					ci = mysheet.createColinfo( clnum, clnum );
				}
				else
				{
					ci = mysheet.createColinfo( clnum, clnum, ci );
				}
				mycol = new ColHandle( ci, this );

			}
			catch( Exception e )
			{
				throw new ColumnNotFoundException( "Unable to getCol for col number " + clnum + " " + e.toString() );
			}
		}
		else
		{
			mycol = new ColHandle( ci, this ); // usual case
		}
		return mycol;
	}

	/**
	 * adds the column (col1st, colLast) and returns the new ColHandle
	 *
	 * @param c1st
	 * @param clast
	 * @return
	 * @deprecated use addCol(int)
	 */
	public ColHandle addCol( int c1st, int clast )
	{
		Colinfo ci = mysheet.createColinfo( c1st, clast );
		ColHandle mycol = new ColHandle( ci, this );
		return mycol;
	}

	/**
	 * adds the column (col1st, colLast) and returns the new ColHandle
	 *
	 * @param colNum, zero based number of the column
	 * @return ColHandle
	 */
	public ColHandle addCol( int colNum )
	{
		Colinfo ci = mysheet.createColinfo( colNum, colNum );
		ColHandle mycol = new ColHandle( ci, this );
		return mycol;
	}

	/**
	 * returns the Column at the named position
	 *
	 * @return ColHandle the Column
	 */
	public ColHandle getCol( String name ) throws ColumnNotFoundException
	{
		return getCol( ExcelTools.getIntVal( name ) );
	}

	/**
	 * returns all of the Columns in this WorkSheet
	 *
	 * @return ColHandle[] Columns
	 */
	public ColHandle[] getColumns()
	{
		List<ColHandle> columns = new CompatibleVector();

		for( Colinfo c : mysheet.getColinfos() )
		{
			try
			{
				int start = c.getColFirst();
				int end = c.getColLast();
				for( int i = start; i <= end; i++ )
				{
					try
					{
						columns.add( getCol( i ) );
					}
					catch( ColumnNotFoundException e )
					{
					}
					;
				}
			}
			catch( Exception ex )
			{
				;
			}
		}

		return columns.toArray( new ColHandle[columns.size()] );
	}

	/**
	 * returns a List of Column names
	 *
	 * @return List column names
	 */
	public List getColNames()
	{
		return mysheet.getColNames();
	}

	/**
	 * returns a List of Row numbers
	 *
	 * @return List of row numbers
	 */
	public List getRowNums()
	{
		return mysheet.getRowNums();
	}

	/**
	 * returns the RowHandle for the row at index position
	 * <p/>
	 * the row index is zero based ie: Excel row 1 = 0
	 *
	 * @param int row number to return
	 * @return RowHandle a Row on this WorkSheet
	 */
	public RowHandle getRow( int t ) throws RowNotFoundException
	{
		Row x = mysheet.getRowByNumber( t );
		if( x == null )
		{
			throw new RowNotFoundException( "Row " + t +
					                                " not found in :" + getSheetName() );
		}
		return new RowHandle( x, this );
	}

	/**
	 * get an array of all RowHandles for
	 * this WorkSheet
	 *
	 * @return RowHandle[] all Rows on this WorkSheet
	 */
	public RowHandle[] getRows()
	{
		Row[] rs = mysheet.getRows();
		RowHandle[] ret = new RowHandle[rs.length];
		for( int t = 0; t < rs.length; t++ )
		{
			ret[t] = new RowHandle( rs[t], this );
		}
		return ret;
	}

	/**
	 * get an array of BIFFREC Rows
	 *
	 * @return RowHandle[] all Rows on this WorkSheet
	 */
	public Map getRowMap()
	{
		return mysheet.getRowMap();
	}

	/**
	 * Returns whether a Cell exists in the WorkSheet.
	 *
	 * @param String the address of the Cell to check for
	 * @return boolean whether the Cell exists
	 */
	boolean hasCell( String addr )
	{
		try
		{
			getCell( addr );
			return true;
		}
		catch( CellNotFoundException e )
		{
			return false;
		}
	}

	/**
	 * Returns the number of rows in this WorkSheet
	 *
	 * @return int Number of Rows on this WorkSheet
	 */
	public int getNumRows()
	{
		return mysheet.getNumRows();
	}

	/**
	 * Returns the number of Columns in this WorkSheet
	 *
	 * @return int Number of Cols on this WorkSheet
	 */
	public int getNumCols()
	{
		return mysheet.getNumCols();
	}

	/**
	 * Remove a Cell from this WorkSheet.
	 *
	 * @param CellHandle to remove
	 */
	public void removeCell( CellHandle celldel )
	{
		mysheet.removeCell( celldel.getCell() );
	}

	/**
	 * removes an Image from the Spreadsheet
	 * <p/>
	 * <p/>
	 * Jan 22, 2010
	 *
	 * @param img
	 */
	public void removeImage( ImageHandle img )
	{
		mysheet.removeImage( img );
		img.remove();
	}

	/**
	 * Remove a Cell from this WorkSheet.
	 *
	 * @param String celladdr - the Address of the Cell to remove
	 */
	public void removeCell( String celladdr )
	{
		mysheet.removeCell( celladdr.toUpperCase() );
	}

	/** Remove a Row and all associated Cells from
	 this WorkSheet.

	 @param int rownum - the number of the row to remove
	  * not used 
	 public void removeRow(int rownum)
	 throws RowNotFoundException{
	 mysheet.removeRow(rownum);
	 }
	 */

	/**
	 * Remove a Row and all associated Cells from
	 * this WorkSheet. Optionally shift all rows below
	 * target row up one.
	 *
	 * @param int     rownum - the number of the row to remove
	 * @param boolean shiftrows - true will shift all lower rows up one.
	 */
	public void removeRow( int rownum, boolean shiftrows ) throws RowNotFoundException
	{
		if( shiftrows )
		{
			removeRow( rownum );
		}
		else
		{
			removeRow( rownum, WorkSheetHandle.ROW_DELETE_NO_REFERENCE_UPDATE );
		}
	}

	/**
	 * Remove a Row and all associated Cells from
	 * this WorkSheet.
	 *
	 * @param int rownum - the number of the row to remove
	 *            uses default row deletion rules regarding updating references
	 */
	public void removeRow( int rownum ) throws RowNotFoundException
	{
		removeRow( rownum, WorkSheetHandle.ROW_DELETE );
	}

	/**
	 * Remove all cells and formatting from a row within
	 * this WorkSheet. Has no other affect upon the workbook
	 *
	 * @param int rownum - the number of the row contents to remove
	 */
	public void removeRowContents( int rownum ) throws RowNotFoundException
	{
		mysheet.removeRowContents( rownum );
	}

	/**
	 * Remove a Row and all associated Cells from
	 * this WorkSheet.
	 *
	 * @param int rownum - the number of the row to remove
	 * @param int flag - controls whether row deletions updates references as well ...
	 */
	public void removeRow( int rownum, int flag ) throws RowNotFoundException
	{

    	/* TODO: deal with merges! */
		mysheet.removeRows( rownum, 1, true );

		// Delete chart series IF SERIES ARE ROW-BASED -- do before updateReferences
		List charts = mysheet.getCharts();
		for( Object chart : charts )
		{
			String sht = GenericPtg.qualifySheetname( getSheetName() );
			Chart c = (Chart) chart;
			HashMap seriesmap = c.getSeriesPtgs();
			Iterator ii = seriesmap.keySet().iterator();
			while( ii.hasNext() )
			{
				Series s = (Series) ii.next();
				Ptg[] ptgs = (Ptg[]) seriesmap.get( s );
				PtgRef pr;
				String cursheet;
				int[] rc;
				for( int j = 0; j < ptgs.length; j++ )
				{
					try
					{
						pr = (PtgRef) ptgs[j];
						cursheet = pr.getSheetName();
						rc = pr.getIntLocation();
						if( (rc[1] != rc[3]) && sht.equalsIgnoreCase( cursheet ) )
						{ // series are in rows, if existing series fall within deleted row 
							if( (rc[0]) == (rownum - 1) )
							{
								c.removeSeries( j );
								break;    // got it
							}
						}
						else
						{
							break;    // isn't row-based so split
						}
					}
					catch( Exception e )
					{
						continue;    // shouldn't happen!
					}
				}
			}
			// also shift chart up if necessary [BugTracker 2858] 
			int row = c.getRow0();
			// only move images whose top is >= rnum
			int rnum = rownum + 1;
			if( row > rnum )
			{
				int h = c.getHeight();
				// move down 1 row
				c.setRow( row - 1 );
				c.setHeight( h );
			}
		}

		if( flag != WorkSheetHandle.ROW_DELETE_NO_REFERENCE_UPDATE )
		{
			ReferenceTracker.updateReferences( rownum, -1, mysheet, true );
		}

		// Adjust image row so that height remains constant
		int rnum = rownum + 1;
		ImageHandle[] images = getImages();
		for( ImageHandle ih : images )
		{
			int row = ih.getRow();
			// only move images whose top is >= rnum
			if( row > rnum )
			{
				short h = ih.getHeight();
				// move down 1 row
				ih.setRow( row - 1 );
				ih.setHeight( h );
			}
		}
	}

	/**
	 * Removes columns and all their associated cells from the sheet.
	 * This method does not shift the subsequent columns left, for that use
	 * {@link #removeCols}.
	 *
	 * @param first the zero-based index of the first column to be removed
	 * @param count the number of columns to remove.
	 */
	public void clearCols( int first, int count )
	{
		removeCols( first, count, false );
	}

	/**
	 * Removes columns from the sheet and shifts the following columns left.
	 *
	 * @param first the zero-based index of the first column to be removed
	 * @param count the number of columns to remove.
	 */
	public void removeCols( int first, int count )
	{
		removeCols( first, count, true );
	}

	private void removeCols( int first, int count, boolean shift )
	{
		if( first < 0 )
		{
			throw new IllegalArgumentException( "column index must be zero or greater" );
		}
		if( count < 1 )
		{
			throw new IllegalArgumentException( "count must be at least one" );
		}
		mysheet.removeCols( first, count, shift );
	}

	/**
	 * Removes a column and all associated cells from this sheet.
	 * This does not shift subsequent columns.
	 *
	 * @param colstr the name of the column to remove
	 * @deprecated Use {@link #clearCols} instead.
	 */
	@Deprecated
	public void removeCol( String colstr ) throws ColumnNotFoundException
	{
		removeCols( ExcelTools.getIntVal( colstr ), 1, false );
	}

	/**
	 * Remove a column and all associated cells from this sheet.
	 * Optionally shift all subsequent columns left to fill the gap.
	 *
	 * @param colstr    the name of the column to remove
	 * @param shiftcols whether to shift subsequent columns
	 * @deprecated Use {@link #removeCols} or {@link #clearCols} instead.
	 */
	@Deprecated
	public void removeCol( String colstr, boolean shiftcols ) throws ColumnNotFoundException
	{
		removeCols( ExcelTools.getIntVal( colstr ), 1, shiftcols );
	}

	/**
	 * Returns the index of the Sheet.
	 *
	 * @return String Sheet Name
	 */
	public int getSheetNum()
	{
		return mysheet.getSheetNum();
	}

	/**
	 * Returns all Named Range Handles scoped to this Worksheet.
	 * <p/>
	 * Note this will not include workbook scoped named ranges
	 *
	 * @return NameHandle[] all of the Named ranges that are scoped to the present worksheet
	 */
	public NameHandle[] getNamedRangesInScope()
	{
		Name[] nand = mysheet.getSheetScopedNames();
		NameHandle[] nands = new NameHandle[nand.length];
		for( int x = 0; x < nand.length; x++ )
		{
			nands[x] = new NameHandle( nand[x], wbh );
		}
		return nands;
	}

	/**
	 * Returns a Named Range Handle if it exists in the specified scope.
	 * <p/>
	 * This can be used to distinguish between multiple named ranges with the same name
	 * but differing scopes
	 *
	 * @return NameHandle a Named range in the Worksheet that exists in the scope
	 */
	public NameHandle getNamedRangeInScope( String rangename ) throws CellNotFoundException
	{
		Name nand = mysheet.getScopedName( rangename );
		if( nand == null )
		{
			throw new CellNotFoundException( rangename );
		}
		return new NameHandle( nand, wbh );
	}

	/**
	 * Returns the name of the Sheet.
	 *
	 * @return String Sheet Name
	 */
	public String getSheetName()
	{
		return mysheet.getSheetName();
	}

	/**
	 * return the sheetname properly qualified or quoted
	 * used when the sheetname contains spaces, commas or parentheses
	 *
	 * @return
	 */
	public String getQualifiedSheetName()
	{
		return GenericPtg.qualifySheetname( mysheet.getSheetName() );
	}

	/**
	 * Returns the underlying low-level Boundsheet
	 * object.
	 *
	 * @return Boundsheet sheet
	 */
	protected Boundsheet getSheet()
	{
		return mysheet;
	}

	/**
	 * Returns the Serialized bytes for this WorkSheet.
	 * <p/>
	 * The output of this method can be used to insert
	 * a copy of this WorkSheet into another WorkBook
	 * using the WorkBookHandle.addWorkSheet(byte[] serialsheet, String NewSheetName)
	 * method.
	 *
	 * @return byte[] the WorkSheet's Serialized bytes
	 * @see WorkBookHandle.addWorkSheet(byte[] serialsheet, String NewSheetName)
	 */
	public byte[] getSerialBytes()
	{
		mysheet.setLocalRecs();
		ObjectOutputStream obs = null;
		byte[] b = null;
		try
		{
			ByteArrayOutputStream baos = new ByteArrayOutputStream();
			BufferedOutputStream bufo = new BufferedOutputStream( baos );
			obs = new ObjectOutputStream( bufo );
			obs.writeObject( mysheet );
			bufo.flush();
			b = baos.toByteArray();
		}
		catch( Throwable e )
		{
			log.error( "Serializing Sheet: " + toString() + " failed: " + e, e );
		}
		return b;
		//return mysheet.getSheetBytes();
	}

	/**
	 * write this sheet as tabbed text output:
	 * <br>
	 * All rows and all characters in each cell are saved. Columns of data are separated by tab characters,
	 * and each row of data ends in a carriage return. If a cell contains a comma, the cell contents are enclosed
	 * in double quotation marks.
	 * All formatting, graphics, objects, and other worksheet contents are lost.
	 * The euro symbol will be converted to a question mark.
	 * If cells display formulas instead of formula values, the formulas are saved as text.
	 */
	public void writeAsTabbedText( OutputStream dest ) throws IOException
	{
		mysheet.writeAsTabbedText( dest );
	}

	/**
	 * Constructor which takes a WorkBook and sheetname as parameters.
	 *
	 * @param sht  The name of the WorkSheet
	 * @param mybk The WorkBook
	 */
	protected WorkSheetHandle( Boundsheet sht, WorkBookHandle b )
	{
		wbh = b;
		mysheet = sht;
		mybook = sht.getWorkBook();
		// 20080624 KSC: add flag for shift formula rules upon row insertion/deletion
		String shiftRule = (String) System.getProperties().get( "com.extentech.extenXLS.WorkSheetHandle.shiftInclusive" );
		if( (shiftRule != null) && shiftRule.equalsIgnoreCase( "true" ) )
		{
			mysheet.setShiftRule( shiftRule.equalsIgnoreCase( "true" ) );
		}
		// 20080917 KSC: set cache setting via system property [BugTracker 1862]
		if( System.getProperty( "com.extentech.ExtenXLS.cacheCellHandles" ) != null )
		{
			cache = Boolean.valueOf( (System.getProperty( "com.extentech.ExtenXLS.cacheCellHandles" )) );
		}
	}

	/**
	 * @return setting on whether to use cache or not
	 */
	public boolean getUseCache()
	{
		return cache;
	}

	/**
	 * Set the Object value of the Cell at the given address.
	 *
	 * @param String Cell Address (ie: "D14")
	 * @param Object new Cell Object value
	 * @throws CellNotFoundException is thrown if there is
	 *                               no existing Cell at the specified address.
	 */
	public void setVal( String address, Object val ) throws CellNotFoundException, CellTypeMismatchException
	{
		CellHandle c = getCell( address );
		c.setVal( val );
	}

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
	public void setVal( String address, double d ) throws CellNotFoundException, CellTypeMismatchException
	{
		CellHandle c = getCell( address );
		c.setVal( d );
	}

	/**
	 * Set the String value of the Cell at the given address
	 *
	 * @param String Cell Address (ie: "D14")
	 * @param String new Cell String value
	 * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public void setVal( String address, String s ) throws CellNotFoundException, CellTypeMismatchException
	{
		CellHandle c = getCell( address );
		c.setVal( s );
	}

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
	public void setSheetName( String name )
	{
		wbh.sheethandles.remove( getSheetName() );    // keep sheethandles (name->wsh) updated
		mysheet.setSheetName( name );
		wbh.sheethandles.put( name, this );
	}

	/**
	 * Set the int value of the Cell at the given address
	 *
	 * @param String Cell Address (ie: "D14")
	 * @param int    new Cell int value
	 * @throws com.extentech.formats.XLS.CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public void setVal( String address, int i ) throws CellNotFoundException, CellTypeMismatchException
	{
		CellHandle c = getCell( address );
		c.setVal( i );
	}

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
	public Object getVal( String address ) throws CellNotFoundException
	{
		CellHandle c = getCell( address );
		return c.getVal();
	}

	/**
	 * Insert a row of Objects into the worksheet.
	 * Automatically shifts all rows below the cell down one.
	 * <p/>
	 * Method takes an array of Objects to insert into the rows.
	 * <p/>
	 * Object array must match columns in number starting with column A.
	 * <p/>
	 * For emptly cells, put a null Object reference in your array.
	 * <p/>
	 * example:
	 * Object[] newCellHandles = {
	 * null, 						// col A
	 * "Hello", 					// col B
	 * Integer.valueOf(120),			// col C
	 * "=sum(A1+B1)",				// col D
	 * null,						// col E
	 * null,						// col F
	 * "World"						// col G
	 * };
	 * <p/>
	 * CellHandle ret = sheet.insertRow(newCellHandles, 1, true);
	 * if(ret !=null)
	 * Logger.log("It worked");
	 *
	 * @param an      array of Objects to insert into the new row
	 * @param rownum  the rownumber to insert
	 * @param whether to shift down existing Cells
	 */
	public CellHandle[] insertRow( int row1, Object[] data )
	{
		return insertRow( row1, data, true );
	}

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
	public boolean insertRow( int rownum )
	{
		return insertRow( rownum, (Row) null, ROW_INSERT, true );
	}

	/**
	 * Insert a blank row into the worksheet.
	 * Shift all rows below the cell down one.
	 * <p/>
	 * Adding new cells to non-existent rows will
	 * automatically create new rows in the file,
	 * This method is only necessary to "move" existing cells
	 * by inserting empty rows.
	 * <p/>
	 * Same as insertRow(rownum) except with addition of flag
	 *
	 * @param rownum the rownumber to insert
	 * @param flag   row insertion rule
	 */
	public void insertRow( int rownum, int flag )
	{
		insertRow( rownum, (Row) null, flag, true );
	}

	private CompatibleVector addedrows = new CompatibleVector();
	private boolean range_init = true;

	/**
	 * Insert a blank row into the worksheet.
	 * Shift all rows below the cell down one.
	 * <p/>
	 * This method differs from insertRow in that it
	 * can be used to repeatedly insert rows at the same
	 * row index.
	 * <p/>
	 * Adding new cells to non-existent rows will
	 * automatically create new rows in the file,
	 * <p/>
	 * After calling this method, setVal() can be used on the
	 * newly created cells to update with new values.
	 *
	 * @param rownum  the rownumber to insert
	 * @param whether to shift down existing Cells
	 * @return whether the insert was successful
	 */
	public boolean insertRowAt( int rownum, boolean shiftrows )
	{
		addedrows.remove( Integer.valueOf( rownum ) );
		return insertRow( rownum, mysheet.getRowByNumber( rownum ), rownum, shiftrows );
	}

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
	 * @param rownum  the rownumber to insert (NOTE: rownum is 0-based)
	 * @param whether to shift down existing Cells
	 * @return whether the insert was successful
	 */
	public boolean insertRow( int rownum, boolean shiftrows )
	{
		Row myr = null;
		try
		{
			myr = mysheet.getRowByNumber( rownum );
		}
		catch( Exception e )
		{
			;
		}
		if( (myr != null) || shiftrows )
		{
			return insertRow( rownum, myr, ROW_INSERT_MULTI, shiftrows );
		}
		// essentially a high performance row insert for the bottom of the workbook, used frequently in streaming workbook insertion
		Row newRow = mysheet.insertRow( rownum, ROW_INSERT_MULTI, shiftrows );
		return true;

	}

	// insert handling flags
	/**
	 * Insert row multiple times allowed, also copies formulas to inserted
	 * row
	 */
	public static final int ROW_INSERT_MULTI = 0;
	/**
	 * Excel standard row insertion behavior
	 */
	public static final int ROW_INSERT = 3;
	/**
	 * Insert row one time, multiple calls ignored
	 */
	public static final int ROW_INSERT_ONCE = 1;
	/**
	 * Insert row but do not update any cell references affected by insert
	 */
	public static final int ROW_INSERT_NO_REFERENCE_UPDATE = 2;

	// 20080619 KSC: Add flag constants for Delete Row 
	public static final int ROW_DELETE = 1;
	public static final int ROW_DELETE_NO_REFERENCE_UPDATE = 2;

	/**
	 * enhanced protection settings:
	 * Edit Object
	 */
	public final static short ALLOWOBJECTS = FeatHeadr.ALLOWOBJECTS;
	/**
	 * enhanced protection settings:
	 * Edit scenario
	 */
	public static final short ALLOWSCENARIOS = FeatHeadr.ALLOWSCENARIOS;
	/**
	 * enhanced protection settings:
	 * Format cells
	 */
	public static final short ALLOWFORMATCELLS = FeatHeadr.ALLOWFORMATCELLS;
	/**
	 * enhanced protection settings:
	 * Format columns
	 */
	public static final short ALLOWFORMATCOLUMNS = FeatHeadr.ALLOWFORMATCOLUMNS;
	/**
	 * enhanced protection settings:
	 * Format rows
	 */
	public static final short ALLOWFORMATROWS = FeatHeadr.ALLOWFORMATROWS;
	/**
	 * enhanced protection settings:
	 * Insert columns
	 */
	public static final short ALLOWINSERTCOLUMNS = FeatHeadr.ALLOWINSERTCOLUMNS;
	/**
	 * enhanced protection settings:
	 * Insert rows
	 */
	public static final short ALLOWINSERTROWS = FeatHeadr.ALLOWINSERTROWS;
	/**
	 * enhanced protection settings:
	 * Insert hyperlinks
	 */
	public static final short ALLOWINSERTHYPERLINKS = FeatHeadr.ALLOWINSERTHYPERLINKS;
	/**
	 * enhanced protection settings:
	 * Delete columns
	 */
	public static final short ALLOWDELETECOLUMNS = FeatHeadr.ALLOWDELETECOLUMNS;
	/**
	 * enhanced protection settings:
	 * Delete rows
	 */
	public static final short ALLOWDELETEROWS = FeatHeadr.ALLOWDELETEROWS;
	/**
	 * enhanced protection settings:
	 * Select locked cells
	 */
	public static final short ALLOWSELLOCKEDCELLS = FeatHeadr.ALLOWSELLOCKEDCELLS;
	/**
	 * enhanced protection settings:
	 * Sort
	 */
	public static final short ALLOWSORT = FeatHeadr.ALLOWSORT;
	/**
	 * enhanced protection settings:
	 * Use Autofilter
	 */
	public static final short ALLOWAUTOFILTER = FeatHeadr.ALLOWAUTOFILTER;
	/**
	 * enhanced protection settings:
	 * Use PivotTable reports
	 */
	public static final short ALLOWPIVOTTABLES = FeatHeadr.ALLOWPIVOTTABLES;
	/**
	 * enhanced protection settings:
	 * Select unlocked cells
	 */
	public static final short ALLOWSELUNLOCKEDCELLS = FeatHeadr.ALLOWSELUNLOCKEDCELLS;

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
	public boolean insertRow( int rownum, RowHandle copyRow, int flag, boolean shiftrows )
	{
		return insertRow( rownum, copyRow.myRow, flag, shiftrows );
	}

	/**
	 * Insert a row of Objects into the worksheet.
	 * Shift all rows below the cell down one.
	 * <p/>
	 * Method takes an array of Objects to insert into the rows.
	 * <p/>
	 * Object array must match columns in number starting with column A.
	 * <p/>
	 * For emptly cells, put a null Object reference in your array.
	 * <p/>
	 * example:
	 * Object[] newCellHandles = {
	 * null, 						// col A
	 * "Hello", 					// col B
	 * Integer.valueOf(120),			// col C
	 * "=sum(A1+B1)",				// col D
	 * null,						// col E
	 * null,						// col F
	 * "World"						// col G
	 * };
	 * <p/>
	 * boolean okay = sheet.insertRow(newCellHandles, 1, true);
	 * if(okay)Logger.log("It worked");
	 *
	 * @param an      array of Objects to insert into the new row
	 * @param rownum  the rownumber to insert
	 * @param whether to shift down existing Cells
	 */
	public CellHandle[] insertRow( int rownum, Object[] data, boolean shiftrows )
	{
		CellHandle[] retc = new CellHandle[data.length];
		try
		{
			insertRow( rownum, shiftrows );
			for( int t = 0; t < data.length; t++ )
			{
				if( data[t] != null )
				{
					retc[t] = add( data[t], rownum, t );
				}
			}
		}
		catch( Exception ex )
		{
			throw new WorkBookException( ex.toString(), WorkBookException.RUNTIME_ERROR );
		}
		return retc;
	}

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
	 * @param rownum  the rownumber to insert (0-based)
	 * @param copyrow the row to copy formats and formulas from
	 * @param flag    determines handling tracking of inserted rows and only allow insertion once
	 * @param whether to shift down existing Cells
	 */
	private boolean insertRow( int rownum, Row copyRow, int flag, boolean shiftrows )
	{
		int offset = 1;
		return shiftRow( rownum, copyRow, flag, shiftrows, offset );
	}

	/**
	 * replacement method for delete row that handles references better
	 *
	 * @param rownum
	 * @param flag
	 * @param shiftrows
	 * @return
	 */
	boolean deleteRow( int rownum, int flag, boolean shiftrows )
	{
		int offset = -1;
		return shiftRow( rownum, null, flag, shiftrows, offset );
	}

	/**
	 * insert/delete agnostic row copy/insert/delete
	 * and formula shifter
	 * <p/>
	 * TODO: Better comments
	 *
	 * @param rownum
	 * @param copyRow
	 * @param flag
	 * @param shiftrows
	 * @param offset
	 * @return
	 */
	private boolean shiftRow( int rownum, Row copyRow, int flag, boolean shiftrows, int offset )
	{

		// If the copyrow is null, such as an insert row on an empty row, create that row, otherwise
		// we end up using different logic for row insertion, which makes no sense.
		if( copyRow == null )
		{
			// insert a blank
			add( null, "A" + rownum + 1 );
			copyRow = mysheet.getRowByNumber( rownum );

		}
		// handle tracking of inserted rows -- if flag is false rows can be inserted multiple times at the same index
		if( flag == WorkSheetHandle.ROW_INSERT_ONCE )
		{ // not inserted
			if( addedrows.contains( Integer.valueOf( rownum ) ) )
			{
				return false; // can't add an existing row!
			}
		}

		// sheetster ui means insert
		// 'on top of' row, shift down
		if( offset == 0 )
		{
			offset = 1;
		}

		// shiftrefs BEFORE inserting new row			
		if( shiftrows && (flag != WorkSheetHandle.ROW_INSERT_NO_REFERENCE_UPDATE) )
		{
			int refUpdateStart = rownum;
			//  ExtenXLS default behavior is to update one row too high.  If we are using ROW_INSERT, update per excel standard,
			// see TestInsertRows.testUpdateFormulaSettings() for testing
			if( flag == WorkSheetHandle.ROW_INSERT )
			{
				refUpdateStart++;
			}
			ReferenceTracker.updateReferences( refUpdateStart,
			                                   offset, mysheet,
			                                   true );    //shift or expand/contract ALL affected references including named ranges
		}

		Row newRow = mysheet.insertRow( rownum,
		                                flag,
		                                shiftrows );    // shifts rows down and inserts a new row, also shifts shared formula refs (see note below) 			

		// *************************************************************************************************************************************************************/
		// Named Range, Formula and AI references:
		// ALL references to the inserted row# (rownum) and rows beyond are shifted in ReferenceTracker.updateReferences.  
		// This method uses the ReferenceTracker collection for the specific sheet in question to iterate through the stored references,
		// shifting them as the shifting rules allow.  
		// In addition, all formulas in the copyrow will be duplicated in the newly inserted row;
		// These formula references are shifted in ReferenceTracker.adjustFormulaRefs (see below) 
		// The only references that are NOT shifted in the schema described above are SharedFormula references (specifically, PtgRefN & PtgAreaA), 
		// which are NOT contained within the ReferenceTracker collection.  
		// These references are shifted "by hand" in ReferenceTracker.moveSharedFormulas, called upon insertRow->->Boundsheet.shiftCellRow 
		// *************************************************************************************************************************************************************/

		if( shiftrows )
		{
			addedrows.add( rownum );
		}

		// TODO:  Why so much logic in here, move this to Boundsheet?

		if( (shiftrows) && (copyRow != null) )
		{
			// Handle shifting reference rules for the newrow only and it's copyrow (NOTE: copyrow may have been shifted via insertRow above although it's references have not yet been shifted)
			int refMovementDiff = (copyRow.getRowNumber() - rownum);    // number of rows to shift
			int refMovementRow = rownum;                                // start row for shifting operation 
			if( refMovementDiff < 0 )
			{
				refMovementRow += refMovementDiff;    // since copyrow < rownum, shifting should be done BEFORE insert row
			}
			newRow.setRowHeight( copyRow.getRowHeight() );

			// Now iterate through all cells in the original row and copy formats and formulas
			Object[] copyRowCells = copyRow.getCellArray();
			Map newmerges = new Hashtable();
			CellRange newmerge = null;
			CellHandle newCellHandle = null;
			Mulblank aMul = null;        // KSC: Mulblank handling
			short c = -1;                // ""
			String sheetname = GenericPtg.qualifySheetname( toString() );
			for( Object copyRowCell1 : copyRowCells )
			{
				BiffRec copyRowCell = (BiffRec) copyRowCell1;
				if( copyRowCell.getOpcode() == XLSConstants.MULBLANK )
				{
					if( copyRowCell.equals( aMul ) )
					{
						c++;    // ref next blank in range - nec. for ixfe (FormatId) see below
					}
					else
					{
						aMul = (Mulblank) copyRowCell;
						c = (short) aMul.getColFirst();
					}
					aMul.setCurrentCell( c );
				}
				CellHandle copyCellHandle = new CellHandle( copyRowCell, wbh );
				copyCellHandle.setWorkSheetHandle( this );
				// this.cellhandles.put(copyCellHandle.getCellAddress(), copyCellHandle);
				int colnum = copyCellHandle.getColNum();

				// insert an empty copy of the cell OR, if it's a formula cell, copy formula and adjust it's cell references appropriate for new cell position  
				if( (copyRowCell.getOpcode() == XLSRecord.FORMULA) &&
						(flag != WorkSheetHandle.ROW_INSERT_NO_REFERENCE_UPDATE) &&
						(flag != WorkSheetHandle.ROW_INSERT) )
				{
					try
					{
						// copy copyrow's formula and then shift it's references relative to copycell's original row and references  
						newCellHandle = add( copyCellHandle.getFormulaHandle().getFormulaString(), rownum, colnum );
						// streaming parser uses fast cell adds which returns null, populate here
						if( newCellHandle == null )
						{
							try
							{
								newCellHandle = getCell( rownum, colnum );
							}
							catch( CellNotFoundException e )
							{
								// should be impossible							}
							}
						}
						// because the original formula has been shifted by one, unshift this sucker...
						ReferenceTracker.adjustFormulaRefs( newCellHandle, refMovementRow, refMovementDiff * -1, true );
						newCellHandle.getFormulaHandle().getFormulaRec().clearCachedValue();
					}
					catch( Exception e )
					{
							log.warn( "WorkSheetHandle.shiftRow() could not adjust formula references in formula: " + copyCellHandle + " while inserting new row." + e
									.toString() );
					}
				}
				else
				{
					newCellHandle = add( null, rownum, colnum );
					// streaming parser uses fast cell adds which returns null, populate here
					if( newCellHandle == null )
					{
						try
						{
							newCellHandle = getCell( rownum, colnum );
						}
						catch( CellNotFoundException e )
						{
							// should be impossible							}
						}
					}
				}
				newCellHandle.setFormatId( copyCellHandle.getFormatId() );

				//				handle merged cells -- assemble the newmerges collection for below
				CellRange oby = copyCellHandle.getMergedCellRange();
				if( oby != null )
				{ // we have a merge
					int[] fr = { rownum, oby.firstcellcol };
					int[] lr = { rownum, oby.lastcellcol };
					String newrng = sheetname + "!" + ExcelTools.formatLocation( fr ) + ":" + ExcelTools.formatLocation( lr );
					newrng = GenericPtg.qualifySheetname( newrng );
						log.debug( "WorksheetHandle.insertRow() created new Merge Range: " + newrng );
					// check if we've already created...
					if( newmerges.get( newrng ) == null )
					{
						newmerge = new CellRange( newrng, wbh, true );
							log.debug( "WorksheetHandle.insertRow() created new Merge Range: " + newrng );
						newmerges.put( newmerge.toString(), newmerge );
					}
				}

			}
			// now update the new merge ranges...
			Collection xl = newmerges.values();
			if( xl != null )
			{
				Iterator itx = xl.iterator();
				while( itx.hasNext() )
				{
					((CellRange) itx.next()).mergeCells( true );
				}
			}
		}
		// Handle Image Movement
		ImageHandle[] images = mysheet.getImages();
		if( images != null )
		{
			for( ImageHandle ih : images )
			{
				int row = ih.getRow();
				// only move images whose top is >= copyRow
				if( row >= rownum )
				{
					// move down 1 row
					short h = ih.getHeight();
					ih.setRow( row + 1 );
					ih.setHeight( h );
				}
			}
		}
		// Insert chart series IF SERIES ARE ROW-BASED			
		List charts = mysheet.getCharts();
		for( Object chart : charts )
		{
			Chart c = (Chart) chart;
			ReferenceTracker.insertChartSeries( c, GenericPtg.qualifySheetname( getSheetName() ), rownum );
			// also shift charts down [BugTracker 2858]
			int row = c.getRow0();
			// only move charts whose top is >= copyRow
			if( row >= rownum )
			{
				// move down 1 row
				int h = c.getHeight();
				c.setRow( row + offset );
				c.setHeight( h );
			}
		}
		return true;
	}

	/**
	 * returns an array of FormatHandles for the ConditionalFormats
	 * applied to this cell
	 *
	 * @return an array of FormatHandles, one for each of the Conditional Formatting rules
	 */
	public ConditionalFormatHandle[] getConditionalFormatHandles()
	{
		ConditionalFormatHandle[] cfx = new ConditionalFormatHandle[mysheet.getConditionalFormats().size()];
		for( int i = 0; i < cfx.length; i++ )
		{
			Condfmt cfmt = (Condfmt) mysheet.getConditionalFormats().get( i );
			cfx[i] = new ConditionalFormatHandle( cfmt, this );
		}
		return cfx;
	}

	/**
	 * Returns the WorkBookHandle for this Sheet
	 *
	 * @return
	 */
	public WorkBookHandle getWorkBook()
	{
		return wbh;
	}

	/**
	 * Get a handle to all of the images in this worksheet
	 *
	 * @return
	 */
	public ImageHandle getImage( String name ) throws ImageNotFoundException
	{
		int idz = mysheet.getImageVect().indexOf( name );
		if( idz > 0 )
		{
			return mysheet.getImageVect().get( idz );
		}
		throw new ImageNotFoundException( "Could not find " + name + " in " + getSheetName() );
	}

	/**
	 * Get a handle to all of the images in this worksheet
	 *
	 * @return
	 */
	public ImageHandle[] getImages()
	{
		return mysheet.getImages();
	}

	/**
	 * returns the actual amount of images contained in the sheet
	 * and is determined by imageMap
	 *
	 * @return
	 */
	public int getNumImages()
	{
		return mysheet.imageMap.size();
	}

	/**
	 * write out all of the images in the Sheet to a directory
	 *
	 * @param imageoutput directory
	 */
	public void extractImagesToDirectory( String outdir )
	{
		ImageHandle[] extracted = getImages();

		// extract and output images
		for( ImageHandle anExtracted : extracted )
		{
			String n = anExtracted.getName();
			if( n.equals( "" ) )
			{
				n = "image" + anExtracted.getMsodrawing().getImageIndex();
			}
			String imgname = n + "." + anExtracted.getType();
				log.debug( "Successfully extracted: " + outdir + imgname );
			try
			{
				FileOutputStream outimg = new FileOutputStream( outdir + imgname );
				anExtracted.write( outimg );
				outimg.flush();
				outimg.close();
			}
			catch( Exception ex )
			{
				log.error( "Could not extract images from: " + this );
			}
		}
	}

	/**
	 * retrieves all charts for this sheet and writes them (in SVG form) to outpdir
	 * <br>Filename is in form of: <sheetname>_Chart<#>.svg
	 *
	 * @param outdir String output folder
	 */
	public void extractChartToDirectory( String outdir )
	{
		ArrayList charts = (ArrayList) mysheet.getCharts();
		String sheetname = getSheetName();
		for( Object chart : charts )
		{
			ChartHandle ch = new ChartHandle( (Chart) chart, getWorkBook() );
			String fname = sheetname + "_Chart" + ch.getId() + ".svg";
			try
			{
				FileOutputStream chartout = new FileOutputStream( outdir + fname );
				chartout.write( ch.getSVG( 1.0 ).getBytes() );    // scaled as necessary in XSL 
			}
			catch( Exception ex )
			{
				log.error( "extractChartToDirectory: Could not extract charts from: " + this + ":" + ex.toString() );
			}

		}
	}

	/**
	 * insert an image into this worksheet
	 *
	 * @param im -- the ImageHandle to insert
	 * @see ImageHandle
	 */
	public void insertImage( ImageHandle im )
	{
		mysheet.insertImage( im );
	}

	/**
	 * Inserts empty columns and shifts the following columns to the right.
	 * This method is used to shift existing columns right to make room for a
	 * new column. To create a column in the file so its size or style can be
	 * set just add a blank cell to its first row with {@link #add}.
	 *
	 * @param first the zero-based index of the first column to insert
	 * @param count the number of columns to insert
	 */
	public void insertCols( int first, int count )
	{
		if( first < 0 )
		{
			throw new IllegalArgumentException( "column index must be zero or greater" );
		}
		if( count < 1 )
		{
			throw new IllegalArgumentException( "count must be at least one" );
		}
		mysheet.insertCols( first, count );
	}

	/**
	 * Inserts an empty column and shifts the following columns to the right.
	 *
	 * @param colnum the zero-based index of the column to be inserted
	 * @deprecated Use {@link #insertCols} instead.
	 */
	@Deprecated
	public void insertCol( int colnum )
	{
		insertCols( colnum, 1 );
	}

	/**
	 * Inserts an empty column and shifts the following columns to the right.
	 *
	 * @param colnum the address of the column to be inserted
	 * @deprecated Use {@link #insertCols} instead.
	 */
	@Deprecated
	public void insertCol( String colnum )
	{
		insertCols( ExcelTools.getIntVal( colnum ), 1 );
	}

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
	public void setCopyFormatsFromPriorWhenAdding( boolean f )
	{
		mysheet.setCopyPriorCellFormats( f );
	}

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
	public CellHandle add( Object obj, int row, int col )
	{
		return add( obj, row, col, getWorkBook().getWorkBook().getDefaultIxfe() );
	}

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
	 * <p/>
	 * If a validation record for the cell exists the validation is checked
	 * for a correct value, if the value does not pass the validation a ValidationException
	 * will be thrown
	 *
	 * @param obj the value of the new Cell
	 * @param int row the row of the new Cell
	 * @param int col the column of the new Cell
	 */
	public CellHandle[] addValidated( Object obj, int row, int col ) throws ValidationException
	{
		return addValidated( obj, row, col, getWorkBook().getWorkBook().getDefaultIxfe() );
	}

	/**
	 * Add a Cell with the specified value to a WorkSheet.
	 * <p/>
	 * This method determines the Cell type based on
	 * type-compatibility of the value.
	 * <p/>
	 * Further, this method allows passing in a format id
	 * <p/>
	 * In other words, if the Object cannot be converted
	 * safely to a Numeric Object type, then it is treated
	 * as a String and a new String value is added to the
	 * WorkSheet at the Cell address specified.
	 * <p/>
	 * If a validation record for the cell exists the validation is checked
	 * for a correct value, if the value does not pass the validation a ValidationException
	 * will be thrown
	 *
	 * @param obj the value of the new Cell
	 * @param int row the row of the new Cell
	 * @param int col the column of the new Cell
	 * @param int the format id to apply to this cell
	 */
	public CellHandle[] addValidated( Object obj, int row, int col, int formatId ) throws ValidationException
	{
		int[] rc = { row, col };
		ValidationHandle vh = getValidationHandle( ExcelTools.formatLocation( rc ) );
		if( vh != null )
		{
			vh.isValid( obj );
		}
		return addValidated( obj, ExcelTools.formatLocation( rc ) );
	}

	/**
	 * Add a Cell with the specified value to a WorkSheet.
	 * <p/>
	 * This method determines the Cell type based on
	 * type-compatibility of the value.
	 * <p/>
	 * Further, this method allows passing in a format id
	 * <p/>
	 * In other words, if the Object cannot be converted
	 * safely to a Numeric Object type, then it is treated
	 * as a String and a new String value is added to the
	 * WorkSheet at the Cell address specified.
	 *
	 * @param obj the value of the new Cell
	 * @param int row the row of the new Cell
	 * @param int col the column of the new Cell
	 * @param int the format id to apply to this cell
	 */
	public CellHandle add( Object obj, int row, int col, int formatId )
	{
		int[] rc = { row, col };

		if( obj instanceof java.util.Date )
		{
			String address = ExcelTools.formatLocation( rc );
			add( (java.util.Date) obj, address, null );
		}
		else
		{
			BiffRec reca = mysheet.addValue( obj, rc, formatId );
		}

		if( !mysheet.fastCellAdds )
		{
			try
			{
				CellHandle c = getCell( row, col );
				if( wbh.getFormulaCalculationMode() != wbh.CALCULATE_EXPLICIT )
				{
					c.clearAffectedCells(); // blow out cache
				}
				return c;
			}
			catch( CellNotFoundException e )
			{
				log.warn( "Adding Cell to row failed row:" + row + " col: " + col + " failed." );
				return null;
			}
		}
		return null;
	}

	/**
	 * Fast-adds a Cell with the specified value to a WorkSheet.
	 * <p/>
	 * This method determines the Cell type based on
	 * type-compatibility of the value.
	 * <p/>
	 * Further, this method allows passing in a format id
	 * <p/>
	 * In other words, if the Object cannot be converted
	 * safely to a Numeric Object type, then it is treated
	 * as a String and a new String value is added to the
	 * WorkSheet at the Cell address specified.
	 *
	 * @param obj the value of the new Cell
	 * @param int row the row of the new Cell
	 * @param int col the column of the new Cell
	 * @param int the format id to apply to this cell
	 */
	public void fastAdd( Object obj, int row, int col, int formatId )
	{
		int[] rc = { row, col };

		// use default format
		if( formatId == -1 )
		{
			formatId = 0;
		}

		if( obj instanceof java.util.Date )
		{
			String address = ExcelTools.formatLocation( rc );
			add( (java.util.Date) obj, address, null );
		}
		else
		{
			BiffRec reca = mysheet.addValue( obj, rc, formatId );
			if( wbh.getFormulaCalculationMode() != wbh.CALCULATE_EXPLICIT )
			{
				ReferenceTracker rt = wbh.getWorkBook().getRefTracker();
				rt.clearAffectedFormulaCells( reca );
			}
				if( reca != null )
				{
					log.debug( "WorkSheetHandle.add() " + reca.toString() + " Successfully Added." );
				}
		}
	}

	/**
	 * Toggle fast cell add mode.
	 * <p/>
	 * Set to true to turn off checking for existing cells, conditional formats and merged ranges in order to
	 * accelerate adding new cells
	 *
	 * @param fastadds whether to disable checking for existing cells and
	 */
	public void setFastCellAdds( boolean fastadds )
	{
		mysheet.setFastCellAdds( fastadds );
	}

	/**
	 * Get the current fast add cell mode for this worksheet
	 */
	public boolean getFastCellAdds()
	{
		return mysheet.fastCellAdds;
	}

	/**
	 * Add a Cell with the specified value to a WorkSheet, optionally attempting to convert
	 * numeric values to appropriate number cells with appropriate number formatting applied.
	 * <p/>
	 * This would allow a value entered such as "$1.00" to be converted to a numeric 1.00d value with
	 * a $0.00 format pattern applied.
	 * <p/>
	 * Note that there is overhead to this method, as the value added needs to be parsed, for higher performance
	 * do not use the autodetectNumberAndPattern = true setting, and instead pass numeric values and format patterns.
	 * <p/>
	 * If the value is non numeric, then it is simply added.
	 *
	 * @param obj             the value of the new Cell
	 * @param address         the address of the new Cell
	 * @param autoDetectValue whether to attempt to store as a number with a format pattern.
	 */
	public CellHandle add( Object obj, String address, boolean autoDetectValue )
	{
		if( obj == null )
		{
			mysheet.addValue( obj, address );  // to add a blank cell
		}
		else if( obj instanceof Date )
		{
			add( (Date) obj, address, null );
		}
		else
		{
			mysheet.addValue( obj, address, autoDetectValue );//
		}

		// fast-adds optimzation -- do not return a CellHandle
		if( mysheet.fastCellAdds )
		{
			return null;
		}

		try
		{
			return getCell( address );
		}
		catch( CellNotFoundException e )
		{
			log.debug( "Adding Cell: " + address + " failed" );
			return null;
		}
	}

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
	public CellHandle add( Object obj, String address )
	{
		if( obj == null )
		{
			mysheet.addValue( obj, address );  // to add a blank cell
		}
		else if( obj instanceof Date )
		{
			add( (Date) obj, address, null );
		}
		else
		{
			mysheet.addValue( obj, address );//
		}

		// fast-adds optimzation -- do not return a CellHandle
		if( !mysheet.fastCellAdds )
		{
			try
			{
				CellHandle c = getCell( address );
				if( wbh.getFormulaCalculationMode() != wbh.CALCULATE_EXPLICIT )
				{
					c.clearAffectedCells(); // blow out cache
				}
				return c;
			}
			catch( CellNotFoundException e )
			{
				log.warn( "Adding Cell: " + address + " failed" );
				return null;
			}
		}
		return null;
	}

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
	 * <p/>
	 * <p/>
	 * If a validation record for the cell exists the validation is checked
	 * for a correct value, if the value does not pass the validation a ValidationException
	 * will be thrown
	 *
	 * @param obj     the value of the new Cell
	 * @param address the address of the new Cell
	 */
	public CellHandle[] addValidated( Object obj, String address ) throws ValidationException
	{
		ValidationHandle vh = getValidationHandle( address );
		if( vh != null )
		{
			vh.isValid( obj );
		}
		CellHandle ch = add( obj, address );
		List cxrs = ch.calculateAffectedCellsOnSheet();

		//	 return the cellhandles
		CellHandle[] cxrx = new CellHandle[cxrs.size() + 1];
		cxrx[0] = ch;
		for( int t = 1; t < cxrx.length; t++ )
		{
			cxrx[t] = (CellHandle) cxrs.get( t - 1 );
			cxrx[t].setWorkSheetHandle( this );
			// this.cellhandles.put(cxrx[t].getCellAddress(), cxrx[t]);
		}

		return cxrx;
	}

	/**
	 * Add a java.sql.Timestamp Cell to a WorkSheet.
	 * <p/>
	 * Will create a default format of:
	 * <p/>
	 * "m/d/yyyy h:mm:ss"
	 * <p/>
	 * if none is specified.
	 *
	 * @param dt         the value of the new java.sql.Timestamp Cell
	 * @param address    the address of the new java.sql.Date Cell
	 * @param formatting pattern the address of the new java.sql.Date Cell
	 */
	public CellHandle add( Timestamp dt, String address, String fmt )
	{
		if( fmt == null )
		{
			fmt = "m/d/yyyy h:mm:ss";
		}
		Date dx = new Date( dt.getTime() );
		return add( dx, address, fmt );
	}

	/**
	 * Add a java.sql.Timestamp Cell to a WorkSheet.
	 * <p/>
	 * Will create a default format of:
	 * <p/>
	 * "m/d/yyyy h:mm:ss"
	 * <p/>
	 * if none is specified.
	 * <p/>
	 * If a validation record for the cell exists the validation is checked
	 * for a correct value, if the value does not pass the validation a ValidationException
	 * will be thrown
	 *
	 * @param dt         the value of the new java.sql.Timestamp Cell
	 * @param address    the address of the new java.sql.Date Cell
	 * @param formatting pattern the address of the new java.sql.Date Cell
	 */
	public CellHandle[] addValidated( Timestamp dt, String address, String fmt ) throws ValidationException
	{
		if( fmt == null )
		{
			fmt = "m/d/yyyy h:mm:ss";
		}
		Date dx = new Date( dt.getTime() );
		return addValidated( dx, address, fmt );
	}

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
	 * @param row        to add the date
	 * @param col        to add the date
	 * @param formatting pattern the address of the new java.sql.Date Cell
	 */
	public CellHandle add( java.util.Date dt, String address, String fmt )
	{
		int[] rc = ExcelTools.getRowColFromString( address );
		return add( dt, rc[0], rc[1], fmt );
	}

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
	 * @param row        to add the date
	 * @param col        to add the date
	 * @param formatting pattern the address of the new java.sql.Date Cell
	 */
	public CellHandle[] addValidated( java.util.Date dt, String address, String fmt ) throws ValidationException
	{
		int[] rc = ExcelTools.getRowColFromString( address );
		return addValidated( dt, rc[0], rc[1], fmt );
	}

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
	public CellHandle add( java.util.Date dt, int row, int col, String fmt )
	{
		double x = DateConverter.getXLSDateVal( dt, mybook.getDateFormat() );
		CellHandle thisCell = add( x, row, col );

		// first handle fast adds
		if( (thisCell == null) && mysheet.fastCellAdds )
		{
			try
			{
				thisCell = getCell( row, col );
			}
			catch( CellNotFoundException exp )
			{
				log.warn( "adding date to WorkSheet failed: " + getSheetName() + ":" + row + ":" + col );
				return null;
			}
		}

		// 20060419 KSC: Use format from cell, if any
		if( fmt == null )
		{
			fmt = thisCell.getFormatPattern();
			if( (fmt == null) || fmt.equals( "General" ) )
			{
				fmt = "m/d/yy h:mm";
			}
		}
		if( dateFormats.get( fmt ) == null )
		{
			FormatHandle fh = thisCell.getFormatHandle();
			fh.setFormatPattern( fmt );
			dateFormats.put( fmt, thisCell.getFormatId() );
		}
		else
		{
			Integer in = (Integer) dateFormats.get( fmt );
			thisCell.setFormatId( in );
		}
		//Logger.logInfo("Date added: " + thisCell.getFormattedStringVal());
		return thisCell;
	}

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
	 * <p/>
	 * If a validation record for the cell exists the validation is checked
	 * for a correct value, if the value does not pass the validation a ValidationException
	 * will be thrown
	 *
	 * @param dt         the value of the new java.sql.Date Cell
	 * @param address    the address of the new java.sql.Date Cell
	 * @param formatting pattern the address of the new java.sql.Date Cell
	 */
	public CellHandle[] addValidated( java.util.Date dt, int row, int col, String fmt ) throws ValidationException
	{
		int[] rc = { row, col };
		ValidationHandle vh = getValidationHandle( ExcelTools.formatLocation( rc ) );
		if( vh != null )
		{
			vh.isValid( dt );
		}
		CellHandle ch = add( dt, row, col, fmt );
		List cxrs = ch.calculateAffectedCellsOnSheet();

		//FIXME: Use List.toArray instead of for loop
		//   return the cellhandles
		CellHandle[] cxrx = new CellHandle[cxrs.size() + 1];
		cxrx[0] = ch;
		for( int t = 1; t < cxrx.length; t++ )
		{
			cxrx[t] = (CellHandle) cxrs.get( t - 1 );
		}

		return cxrx;

	}

	/**
	 * Remove this WorkSheet from the WorkBook
	 * <p/>
	 * NOTE: will throw a WorkBookException if the last sheet is removed.  This results in an invalid output file.
	 */
	public void remove()
	{
		mybook.removeWorkSheet( mysheet );
		wbh.sheethandles.remove( getSheetName() );
	}

	/**
	 * Create a CellRange object from an ExtenXLS string range passed in such as "A1:F6"
	 *
	 * @param rangeName "A1:F6"
	 * @return
	 * @throws CellNotFoundException if the range cannot be created
	 */
	public CellRange getCellRange( String rangeName ) throws CellNotFoundException
	{
		CellRange cr = new CellRange( getSheetName() + "!" + rangeName, getWorkBook() );
		return cr;
	}

	/**
	 * Returns all CellHandles defined on this WorkSheet.
	 *
	 * @return CellHandle[] - the array of Cells in the Sheet
	 */
	public CellHandle[] getCells()
	{
		BiffRec[] cells = mysheet.getCells();
		CellHandle[] retval = new CellHandle[cells.length];
		Mulblank aMul = null;
		short c = -1;
		for( int i = 0; i < retval.length; i++ )
		{
			try
			{
				if( cells[i].getOpcode() != XLSConstants.MULBLANK )
				{
					retval[i] = getCell( cells[i].getRowNumber(), cells[i].getColNumber() );
				}
				else
				{ // Handle MULBLANKS proper column number
					if( cells[i].equals( aMul ) )
					{
						c++;
					}
					else
					{
						aMul = (Mulblank) cells[i];
						c = (short) aMul.getColFirst();
					}
					retval[i] = getCell( cells[i].getRowNumber(), c );
				}
			}
			catch( CellNotFoundException cnfe )
			{
				// try harder
				retval[i] = new CellHandle( cells[i], wbh );
				retval[i].setWorkSheetHandle( this );
				if( cells[i].getOpcode() == XLSConstants.MULBLANK )
				{
					// handle Mulblanks: ref a range of cells; to get correct cell address,
					// traverse thru range and set cellhandle ref to correct column
					if( cells[i].equals( aMul ) )
					{
						c++;
					}
					else
					{
						aMul = (Mulblank) cells[i];
						c = (short) aMul.getColFirst();
					}
					retval[i].setBlankRef( c );    // for Mulblank use only -sets correct column reference for multiple blank cells ...
				}
			}
		}
		return retval;
	}

	/**
	 * Returns a FormulaHandle for working with
	 * the ranges of a formula on a WorkSheet.
	 *
	 * @param addr the address of the Cell
	 * @throws FormulaNotFoundException is thrown if there is no existing formula at the specified address.
	 */
	public FormulaHandle getFormula( String addr ) throws FormulaNotFoundException, CellNotFoundException
	{
		CellHandle c = getCell( addr );
		return c.getFormulaHandle();
	}

	/**
	 * Returns a CellHandle for working with
	 * the value of a Cell on a WorkSheet.
	 *
	 * @param addr the address of the Cell
	 * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public CellHandle getCell( String addr ) throws CellNotFoundException
	{
		if( addr == null )
		{
			throw new IllegalArgumentException("Cell address cannot be null!");
		}

		CellHandle ret = null; // 

		BiffRec c = mysheet.getCell( addr.toUpperCase() );
		if( c == null )
		{
			String sn = "";
			try
			{
				sn = getSheetName();
				sn += "!";
			}
			catch( Exception e )
			{
				;
			}

			throw new CellNotFoundException( sn + addr );
		}
		ret = new CellHandle( c, wbh );
		ret.setWorkSheetHandle( this );
		return ret;
	}

	/**
	 * Returns a CellHandle for working with
	 * the value of a Cell on a WorkSheet.
	 * <p/>
	 * returns a new CellHandle with each call
	 * <p/>
	 * use caching method getCell(int row, int col, boolean cache)
	 * to control caching of CellHandles.
	 *
	 * @param int Row the integer row of the Cell
	 * @param int Col the integer col of the Cell
	 * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public CellHandle getCell( int row, int col ) throws CellNotFoundException
	{
		return getCell( row, col, false );
	}

	/**
	 * Returns a CellHandle for working with
	 * the value of a Cell on a WorkSheet.
	 *
	 * @param int     Row the integer row of the Cell
	 * @param int     Col the integer col of the Cell
	 * @param boolean whether to cache or return a new CellHandle each call
	 * @throws CellNotFoundException is thrown if there is no existing Cell at the specified address.
	 */
	public CellHandle getCell( int row, int col, boolean cache ) throws CellNotFoundException
	{
		CellHandle ret = null;
		if( cache )
		{
			int[] rc = { row, col };
			String address = ExcelTools.formatLocation( rc );
			// ret = (CellHandle)cellhandles.get(address);
			if( ret != null ) // caching!
			{
				return ret;
			}
			ret = new CellHandle( mysheet.getCell( row, col ), wbh );
			ret.setWorkSheetHandle( this );
			//	cellhandles.put(address,ret);
			return ret;
		}
		ret = new CellHandle( mysheet.getCell( row, col ), wbh );
		ret.setWorkSheetHandle( this );
		return ret;
	}

	/**
	 * Move a cell on this WorkSheet.
	 *
	 * @param CellHandle c - the cell to be moved
	 * @param String     celladdr - the destination address of the cell
	 */
	public void moveCell( CellHandle c, String addr ) throws CellPositionConflictException
	{
		mysheet.moveCell( c.getCellAddress(), addr );
	}

	/**
	 * Get the text for the Footer printed at the bottom
	 * of the Worksheet
	 *
	 * @return String footer text
	 */
	public String getFooterText()
	{
		return mysheet.getFooter().getFooterText();
	}

	/**
	 * Get the text for the Header printed at the top
	 * of the Worksheet
	 *
	 * @return String header text
	 */
	public String getHeaderText()
	{
		return mysheet.getHeader().getHeaderText();
	}

	/**
	 * Get the print area set for this WorkSheetHandle.
	 * <p/>
	 * If no print area is set return null;
	 */
	public String getPrintArea()
	{
		return mysheet.getPrintArea();
	}

	/**
	 * Get the Print Titles set for this WorkSheetHandle.
	 * <p/>
	 * If no Print Titles are set, this returns null;
	 */
	public String getPrintTitles()
	{
		return mysheet.getPrintTitles();
	}

	/**
	 * Get the printer settings handle for this WorkSheetHandle.
	 */
	public PrinterSettingsHandle getPrinterSettings()
	{
		return mysheet.getPrinterSetupHandle();
	}

	/**
	 * Sets the print area for the worksheet
	 * <p/>
	 * <p/>
	 * sets the printarea as a CellRange
	 *
	 * @param printarea
	 */
	public void setPrintArea( CellRange printarea )
	{
		mysheet.setPrintArea( printarea.getRange() );
	}

	/**
	 * Set the text for the Header printed at the top
	 * of the Worksheet
	 *
	 * @param String header text
	 */
	public void setHeaderText( String t )
	{
		mysheet.getHeader().setHeaderText( t );
	}

	/**
	 * Set the text for the Footer printed at the bottom
	 * of the Worksheet
	 *
	 * @param String footer text
	 */
	public void setFooterText( String t )
	{
		mysheet.getFooter().setFooterText( t );
	}

	/**
	 * Set the default column width of the worksheet
	 * <p/>
	 * <br/>
	 * <p/>
	 * This setting is roughly the width of the character '0'
	 * The default width of a column is 8.
	 */
	public void setDefaultColWidth( int t )
	{
		mysheet.setDefaultColumnWidth( t );
	}

	public int getDefaultColWidth()
	{
		return (int) mysheet.getDefaultColumnWidth();
	}

	/**
	 * Returns the name of this Sheet.
	 *
	 * @see java.lang.Object#toString()
	 */
	public String toString()
	{
		return mysheet.toString();
	}

	/**
	 * FOR internal Use Only!
	 *
	 * @return Returns the low-level sheet record.
	 */
	public Boundsheet getMysheet()
	{
		return mysheet;
	}

	/**
	 * Calculates all formulas that reference the cell address passed in.
	 * <p/>
	 * Please note that these cells have already been calculated, so in order to get their values without re-calculating them
	 * Extentech suggests setting the book level non-calculation flag, ie
	 * book.setFormulaCalculationMode(WorkBookHandle.CALCULATE_EXPLICIT)
	 * or
	 * FormulaHandle.getCachedVal()
	 *
	 * @return List of of calculated cells
	 */
	public List calculateAffectedCells( String CellAddress )
	{
		CellHandle c = null;
		try
		{
			c = getCell( CellAddress );
		}
		catch( CellNotFoundException e )
		{
			return null;
		}
		return c.calculateAffectedCells();
	}

	/**
	 * Set whether to show calculated formula results in the output sheet.
	 *
	 * @return boolean whether to show calculated formula results
	 */
	public boolean getShowFormulaResults()
	{
		return mysheet.getWindow2().getShowFormulaResults();
	}

	public void setShowFormulaResults( boolean b )
	{
		mysheet.getWindow2().setShowFormulaResults( b );
	}

	/**
	 * Get whether to show gridlines in the output sheet.
	 *
	 * @param boolean whether to show gridlines
	 */
	public boolean getShowGridlines()
	{
		return mysheet.getWindow2().getShowGridlines();
	}

	/**
	 * Set whether to show gridlines in the output sheet.
	 *
	 * @return boolean whether to show gridlines
	 */
	public void setShowGridlines( boolean b )
	{
		mysheet.getWindow2().setShowGridlines( b );
	}

	/**
	 * Get whether to show sheet headers  in the output sheet.
	 *
	 * @return boolean whether to show sheet headers
	 */
	public boolean getShowSheetHeaders()
	{
		return mysheet.getWindow2().getShowSheetHeaders();
	}

	/**
	 * Set whether to show sheet headers in the output sheet.
	 *
	 * @param boolean whether to show sheet headers
	 */
	public void setShowSheetHeaders( boolean b )
	{
		mysheet.getWindow2().setShowSheetHeaders( b );
	}

	/**
	 * Get whether to show zero values in the output sheet.
	 *
	 * @return boolean whether to show zero values
	 */
	public boolean getShowZeroValues()
	{
		return mysheet.getWindow2().getShowZeroValues();
	}

	/**
	 * Set whether to show zero values  in the output sheet.
	 *
	 * @return boolean whether to show zero values
	 */
	public void setShowZeroValues( boolean b )
	{
		mysheet.getWindow2().setShowZeroValues( b );
	}

	/**
	 * Get whether to show outline symbols in the output sheet.
	 *
	 * @return boolean whether to outline symbols
	 */
	public boolean getShowOutlineSymbols()
	{
		return mysheet.getWindow2().getShowOutlineSymbols();
	}

	/**
	 * Set whether to show outline symbols in the output sheet.
	 *
	 * @param boolean whether to show outline symbols
	 */
	public void setShowOutlineSymbols( boolean b )
	{
		mysheet.getWindow2().setShowOutlineSymbols( b );
	}

	/**
	 * Get whether to show normal view or page break preview  view in the output sheet.
	 *
	 * @return boolean whether to show normal view or page break preview view
	 */
	public boolean getShowInNormalView()
	{
		return mysheet.getWindow2().getShowInNormalView();
	}

	/**
	 * Set whether to show normal view or page break preview view in the output sheet.
	 *
	 * @param boolean whether to show normal view or page break preview view
	 */
	public void setShowInNormalView( boolean b )
	{
		mysheet.getWindow2().setShowInNormalView( !b ); // opposite of expected behavior
	}

	/**
	 * Get whether there are  freeze panes in the output sheet.
	 *
	 * @return boolean whether there are freeze panes
	 */
	public boolean hasFrozenPanes()
	{
		return mysheet.getWindow2().getFreezePanes();
	}

	/**
	 * Set whether there are  freeze panes in the output sheet.
	 *
	 * @param boolean whether there are freeze panes
	 */
	public void setHasFrozenPanes( boolean b )
	{
		mysheet.getWindow2().setFreezePanes( b );
		if( !b && (mysheet.getPane() != null) )
		{
			mysheet.removePane();    // remove pane rec if unfreezing ... can also convert to plain splits, but a bit more complicated ...
		}
	}

	/**
	 * sets the zoom for the sheet
	 *
	 * @param the zoom as a float percentage (.25 = 25%)
	 */
	public void setZoom( float zm )
	{
		mysheet.getScl().setZoom( zm );
	}

	/**
	 * if this sheet has freeze panes, return the address of the top left cell
	 * otherwise, return null
	 *
	 * @return
	 */
	public String getTopLeftCell()
	{
		if( mysheet.getPane() != null )
		{
			return mysheet.getPane().getTopLeftCell();
		}
		return null;
	}

	/**
	 * gets the zoom for the sheet
	 *
	 * @return the zoom as a float percentage (.25 = 25%)
	 */
	public float getZoom()
	{
		return mysheet.getScl().getZoom();
	}

	/**
	 * freezes the rows starting at the specified row and creating a scrollable sheet below this row
	 *
	 * @param row the row to start the freeze
	 */
	public void freezeRow( int row )
	{
		if( mysheet.getPane() == null )
		{
			mysheet.setPane( null );    // will add new
		}
		mysheet.getPane().setFrozenRow( row );
	}

	/**
	 * freezes the cols starting at the specified column and creating a scrollable sheet to the right
	 *
	 * @param col the col to start the freeze
	 */
	public void freezeCol( int col )
	{
		if( mysheet.getPane() == null )
		{
			mysheet.setPane( null );    // will add new
		}
		mysheet.getPane().setFrozenColumn( col );
	}

	/**
	 * splits the worksheet at column col for nCols
	 * <p/>
	 * Note: unfreezes panes if frozen
	 *
	 * @param col      col start col to split
	 * @param splitpos position of the horizontal split
	 */
	public void splitCol( int col, int splitpos )
	{
		if( mysheet.getPane() == null )
		{
			mysheet.setPane( null );    // will add new
		}
		mysheet.getPane().setSplitColumn( col, splitpos );
	}

	/**
	 * splits the worksheet at row  for nRows
	 * <p/>
	 * Note: unfreezes panes if frozen
	 *
	 * @param row      start row to split
	 * @param splitpos position of the vertical split
	 */
	public void splitRow( int row, int splitpos )
	{
		if( mysheet.getPane() == null )
		{
			mysheet.setPane( null );    // will add new
		}
		mysheet.getPane().setSplitRow( row, splitpos );
	}

	/**
	 * Gets the row number (0 based) that the
	 * sheet split is located on.  If the sheet is not split
	 * returns -1
	 *
	 * @return
	 */
	public int getSplitRowLocation()
	{
		if( mysheet.getPane() == null )
		{
			return -1;
		}
		return mysheet.getPane().getVisibleRow();
	}

	/**
	 * gets the column number (0-based)that the sheet
	 * split is locaated on; if the sheet is not split,
	 * returns -1
	 *
	 * @return 0-based index of split column, if any
	 */
	public int getSplitColLocation()
	{
		if( mysheet.getPane() == null )
		{
			return -1;
		}
		return mysheet.getPane().getVisibleCol();
	}

	/**
	 * Gets the twips split location
	 * returns -1
	 *
	 * @return
	 */
	public int getSplitLocation()
	{
		if( mysheet.getPane() == null )
		{
			return -1;
		}
		return mysheet.getPane().getRowSplitLoc();
	}

	/**
	 * Get whether to use manual grid color  in the output sheet.
	 *
	 * @return boolean whether to use manual grid color
	 */
	public boolean getManualGridLineColor()
	{
		return mysheet.getWindow2().getManualGridLineColor();
	}

	/**
	 * Set whether to use manual grid color  in the output sheet.
	 *
	 * @param boolean whether to use manual grid color
	 */
	public void setManualGridLineColor( boolean b )
	{
		mysheet.getWindow2().setManualGridLineColor( b );
	}

	public CompatibleVector getAddedrows()
	{
		return addedrows;
	}

	/**
	 * Create a Conditional Format handle for a cell/range
	 *
	 * @param cellAddress     without sheetname.  Can also be a range, such as A1:B5
	 * @param qualifier       = maps to CONDITION_* bytes in ConditionalFormatHandle
	 * @param value1          = the error message
	 * @param value2          = the error title
	 * @param format          = the initial format string to use with the condition
	 * @param firstCondition  = formula string
	 * @param secondCondition = 2nd formula string (optional)
	 * @return
	 */
	public ConditionalFormatHandle createConditionalFormatHandle( String cellAddress,
	                                                              String operator,
	                                                              String value1,
	                                                              String value2,
	                                                              String format,
	                                                              String firstCondition,
	                                                              String secondCondition )
	{

		if( (cellAddress != null) && (cellAddress.indexOf( "!" ) == -1) )
		{
			cellAddress = getSheetName() + "!" + cellAddress;
		}
		Condfmt cfm = mysheet.createCondfmt( cellAddress, wbh );

		Cf cfr = mysheet.createCf( cfm );
		cfr.setOperator( operator );

		// only place this is done..
		cfr.setCondition1( value1 );
		cfr.setCondition2( value2 );

		Cf.setStylePropsFromString( format, cfr );

		ConditionalFormatHandle cfh = new ConditionalFormatHandle( cfm, this );
		// done above in createCondfmt this.mysheet.addConditionalFormat(cfm);
		return cfh;
	}

	/**
	 * Creates a new annotation (Note or Comment) to the worksheeet, attached to a specific cell
	 *
	 * @param address -- address to attach
	 * @param txt     -- text of note
	 * @param author  -- name of author
	 * @return NoteHandle - handle which allows access to the Note object
	 * @see CommentHandle
	 */
	public CommentHandle createNote( String address, String txt, String author )
	{
		Note n = getMysheet().createNote( address, txt, author );
		return new CommentHandle( n );
	}

	/**
	 * Creates a new annotation (Note or Comment) to the worksheeet, attached to a specific cell
	 * <br>The note or comment is a Unicode string, thus it can contain formatting information
	 *
	 * @param address -- address to attach
	 * @param txt     -- Unicode string of note with Formatting
	 * @param author  -- name of author
	 * @return NoteHandle - handle which allows access to the Note object
	 * @see CommentHandle
	 */
	public CommentHandle createNote( String address, Unicodestring txt, String author )
	{
		Note n = getMysheet().createNote( address, txt, author );
		return new CommentHandle( n );
	}

	/**
	 * returns an array of all CommentHandles that exist in the sheet
	 *
	 * @return
	 */
	public CommentHandle[] getCommentHandles()
	{
		ArrayList notes = getMysheet().getNotes();
		CommentHandle[] nHandles = new CommentHandle[notes.size()];
		for( int i = 0; i < nHandles.length; i++ )
		{
			nHandles[i] = new CommentHandle( (Note) notes.get( i ) );
		}
		return nHandles;
	}

	/**
	 * creates a new, blank PivotTable and adds it to the worksheet.
	 *
	 * @param name  pivot table name
	 * @param range source range for the pivot table.  If no sheet is specified, the current sheet will be used.
	 * @param sId   Stream or cachid Id -- links back to SxStream set of records
	 * @return PivotTableHandle
	 */
	public PivotTableHandle createPivotTable( String name, String range, int sId )
	{
		Sxview sx = getMysheet().addPivotTable( range, wbh, sId, name );
		PivotTableHandle pth = new PivotTableHandle( sx, getWorkBook() );
		pth.setSourceDataRange( range );
		return pth;
	}

	/**
	 * Get a validation handle for the cell address passed in.
	 * If the validation is for a range, the handle returned will
	 * modify the entire range, not just the cell address passed in.
	 * <p/>
	 * Returns null if a validation does not exist at the specified location
	 *
	 * @param cell address String
	 */
	public ValidationHandle getValidationHandle( String cellAddress )
	{
		if( mysheet.getDvalRec() != null )
		{
			Dv d = mysheet.getDvalRec().getDv( cellAddress );
			if( d == null )
			{
				return null;
			}
			return new ValidationHandle( d );
		}
		return null;
	}

	/**
	 * Create a validation handle for a cell/range
	 *
	 * @param cellAddress     without sheetname.  Can also be a range, such as A1:B5
	 * @param valueType       = maps to VALUE_* bytes in ValidationHandle
	 * @param condition       = maps to CONDITION_* bytes in ValidationHandle
	 * @param errorBoxText    = the error message
	 * @param errorBoxTitle   = the error title
	 * @param promptBoxText   = the prompt (hover) message
	 * @param promptBoxTitle  = the prompt (hover) title
	 * @param firstCondition  = formula string, seeValidationHandle.setFirstCondition
	 * @param secondCondition = seeValidationHandle.setSecondCondition, this can be left null for validations that do not require a second argument.
	 * @return
	 */
	public ValidationHandle createValidationHandle( String cellAddress,
	                                                byte valueType,
	                                                byte condition,
	                                                String errorBoxText,
	                                                String errorBoxTitle,
	                                                String promptBoxText,
	                                                String promptBoxTitle,
	                                                String firstCondition,
	                                                String secondCondition )
	{
		Dv d = mysheet.createDv( cellAddress );
		d.setValType( valueType );
		/*// KSC: APPARENTLY NOT NEEDED
        if (valueType==ValidationHandle.VALUE_USER_DEFINED_LIST) { // ensure Mso Drop-downs are defined        	
        	int objId = this.mysheet.insertDropDownBox(d.getColNumber()); //TODO: verify that drop down lists are SHARED ****
        	this.mysheet.getDvalRec().setObjectIdentifier(objId);        	
        }*/
		d.setTypeOperator( condition );
		d.setErrorBoxText( errorBoxText );
		d.setErrorBoxTitle( errorBoxTitle );
		d.setPromptBoxText( promptBoxText );
		d.setPromptBoxTitle( promptBoxTitle );
		if( firstCondition != null )
		{
			d.setFirstCond( firstCondition );
		}
		if( secondCondition != null )
		{
			d.setSecondCond( secondCondition );
		}
		ValidationHandle vh = new ValidationHandle( d );
		return vh;
	}

	/**
	 * Return all validation handles that refer to this worksheet
	 *
	 * @return array of all validationhandles valid for this worksheet
	 */
	public ValidationHandle[] getAllValidationHandles()
	{
		if( mysheet.getDvRecs() == null )
		{
			return new ValidationHandle[0];
		}
		ValidationHandle[] vh = new ValidationHandle[mysheet.getDvRecs().size()];
		List dvrecs = mysheet.getDvRecs();
		for( int i = 0; i < vh.length; i++ )
		{
			vh[i] = new ValidationHandle( (Dv) dvrecs.get( i ) );
		}
		return vh;
	}

	/**
	 * return true if sheet contains data validations
	 *
	 * @return boolean
	 */
	public boolean hasDataValidations()
	{
		return (mysheet.getDvRecs() != null);
	}

	/**
	 * Returns a list of all AutoFilterHandles on this sheet
	 * <br>An AutoFilterHandle allows access and manipulation of AutoFilters on the sheet
	 *
	 * @return array of AutoFilterHandles if any exist on sheet, null otherwise
	 */
	public AutoFilterHandle[] getAutoFilterHandles()
	{
		if( mysheet.getAutoFilters() == null )
		{
			return null;
		}
		AutoFilterHandle[] af = new AutoFilterHandle[mysheet.getAutoFilters().size()];
		List afs = mysheet.getAutoFilters();
		for( int i = 0; i < afs.size(); i++ )
		{
			af[i] = new AutoFilterHandle( (AutoFilter) afs.get( i ) );
		}
		return af;
	}

	/**
	 * Adds a new AutoFilter for the specified column (0-based) in this sheet
	 * <br>returns a handle to the new AutoFilter
	 *
	 * @param int column - column number to add an AutoFilter to
	 * @return AutoFilterHandle
	 */
	public AutoFilterHandle addAutoFilter( int column )
	{
		AutoFilter af = mysheet.addAutoFilter( column );
		return new AutoFilterHandle( af );
	}

	/**
	 * Removes all AutoFilters from this sheet
	 * <br>As a consequence, all previously hidden rows are shown or unhidden
	 */
	public void removeAutoFilters()
	{
		mysheet.removeAutoFilter();
	}

	/**
	 * Updates the Row filter (hidden status) for each row on the sheet
	 * by evaluating all AutoFilter conditions
	 * <p>NOTE: This method <b>must</b> be called after Autofilter updates or additions
	 * in order to see the results of the AutoFilter(s)
	 * <p>NOTE: this evaluation is NOT done automatically due to performance considerations, and is designed
	 * to be called after all additions and updating is completed
	 * (as evaluation may be time-consuming)
	 */
	public void evaluateAutoFilters()
	{
		mysheet.evaluateAutoFilters();
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		if( mysheet != null )
		{
			mysheet.close();
		}
		addedrows.clear();
		addedrows = new CompatibleVector();
		mysheet = null;
		mybook = null;
		wbh = null;
		dateFormats.clear();
		dateFormats = null;
	}

	protected Boundsheet getBoundsheet()
	{
		return mysheet;
	}

	/**
	 * Imports the given CSV data into this worksheet.
	 * All rows in the input will be inserted sequentially before any rows
	 * which already exist in this worksheet.
	 * <p/>
	 * To change the value delimiter set the system property
	 * "{@code com.extentech.extenxls.csvdelimiter}".
	 */
	public void readCSV( BufferedReader input ) throws IOException
	{
		int rws = 0;
		String field_delimiter = System.getProperty( "com.extentech.extenxls.csvdelimiter", "," );

		String thisLine = "";
		while( (thisLine = input.readLine()) != null )
		{ // while loop begins here
			String[] vals = StringTool.getTokensUsingDelim( thisLine, field_delimiter );
			Object[] data = new Object[vals.length];
			for( int t = 0; t < vals.length; t++ )
			{
				vals[t] = StringTool.strip( vals[t], '"' );
				try
				{
					int i = Integer.parseInt( vals[t] );
					data[t] = i;

					double d = Double.parseDouble( vals[t] );
					data[t] = d;

				}
				catch( NumberFormatException ax )
				{
					// it's a string!
					data[t] = vals[t];
				}
			}

			insertRow( rws++, data, true );
		}
	}

}