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
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.Mulblank;
import com.extentech.formats.XLS.Row;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.awt.*;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;

import static com.extentech.ExtenXLS.JSONConstants.JSON_CELL;
import static com.extentech.ExtenXLS.JSONConstants.JSON_CELLS;
import static com.extentech.ExtenXLS.JSONConstants.JSON_HEIGHT;
import static com.extentech.ExtenXLS.JSONConstants.JSON_ROW;
import static com.extentech.ExtenXLS.JSONConstants.JSON_ROW_BORDER_BOTTOM;
import static com.extentech.ExtenXLS.JSONConstants.JSON_ROW_BORDER_TOP;

/**
 * The RowHandle provides access to a Worksheet Row and its Cells.
 * <br>
 * Use the RowHandle to work with individual Rows in an XLS file.
 * <br>
 * With a RowHandle you can:
 * <br><blockquote>
 * get a handle to the Cells in a row<br>
 * set the default formatting for a Row<br>
 * <br></blockquote>
 *
 * @see WorkBookHandle
 * @see WorkSheetHandle
 * @see FormulaHandle
 */
public class RowHandle
{
	// FYI: do not change lightly -- these match Excel 2007 almost exactly
	public static int ROW_HEIGHT_DIVISOR = 17;

	public Row myRow;
	private FormatHandle formatter;
	private WorkBook wbh;
	private WorkSheetHandle wsh;

	protected RowHandle( Row c, WorkSheetHandle ws )
	{
		myRow = c;
		wbh = ws.getWorkBook();
		wsh = ws;
	}

	/**
	 * Return the row height of an existing row.
	 * <p/>
	 * These values are returned in twips, 1/20th of a character.
	 *
	 * @return int Height of Row in twips
	 */
	public int getHeight()
	{
		return myRow.getRowHeight();
	}

	/**
	 * returns the row height in Excel units, which depends upon the default font
	 * <br>in Arial 10 pt, standard row height is 12.75 points
	 *
	 * @return int row height in Excel units
	 */
	public int getHeightInChars()
	{
		return myRow.getRowHeight() / 20;
	}

	/**
	 * Return the row height of an existing row.
	 * <p/>
	 * These values are returned in twips, 1/20th of a character.
	 *
	 * @param sheet
	 * @param row
	 * @return
	 */
	public static int getHeight( Boundsheet sheet, int row )
	{
		int h = 255;
		try
		{
			Row r = sheet.getRowByNumber( row );
			if( r != null )
			{
				h = r.getRowHeight();
			}
		}
		catch( Exception e )
		{    // exception if no row defined
			h = 255; // default
		}
		return h;
	}

	/**
	 * sets the row height in Excel units.
	 *
	 * @param double i - row height value in Excel units
	 */
	public void setHeightInChars( int newHeight )
	{
		this.setHeight( newHeight * 20 );    // 20090506 KSC: apparently it's in twips ?? 1/20 of a point
	}

	/**
	 * sets the row height to auto fit
	 * <br>When the row height is set manually, autofit is automatically turned off
	 */
	public void setRowHeightAutoFit()
	{
		// this.myRow.setUnsynched(false);	// firstly, set so excel
		Collection ct = myRow.getCells();
		Iterator it = ct.iterator();
		double h = 0;
		int dpi = Toolkit.getDefaultToolkit().getScreenResolution();        // 96 is "small", 120 dpi is "lg"
		// 1 point= 1/72 of an inch
		// 1 twip=  1 twip= 1/20 of a point
		// this should be a pretty good pixels/twips conversion factor.
		double factorTwip = (double) dpi / 72 / 20;        // .06 is "normal"
		// factorZero is width of 0 char in default font.  If assume Arial 10 pt, it is 6 + 1= 7
		double factorZero = 7;    //java.awt.Toolkit.getDefaultToolkit().getFontMetrics(f).charWidth('0') + 1;

		while( it.hasNext() )
		{
			XLSRecord cellrec = (XLSRecord) it.next();
			try
			{
				double newH = 255;    // default row height
				try
				{
					Font ef = cellrec.getXfRec().getFont();
					int style = java.awt.Font.PLAIN;
					if( ef.getBold() )
					{
						style |= java.awt.Font.BOLD;
					}
					if( ef.getItalic() )
					{
						style |= java.awt.Font.ITALIC;
					}
					java.awt.Font f = new java.awt.Font( ef.getFontName(), style, (int) ef.getFontHeightInPoints() );
					String s = cellrec.getStringVal();
					if( !cellrec.getXfRec().getWrapText() ) // normal case, no wrap
					{
						newH = StringTool.getApproximateHeight( f, s, Double.MAX_VALUE );
					}
					else
					{                // wrap to column width
						// convert column width to pixels
						// factorZero is usually 7		// double factorZero= java.awt.Toolkit.getDefaultToolkit().getFontMetrics(f).charWidth('0') + 1;
						double cw = ColHandle.getWidth( this.wsh.getBoundsheet(), cellrec.getColNumber() ) / 256.0;
						newH = StringTool.getApproximateHeight( f, s, cw * factorZero );
					}
// this doesn't work correctly		    newH/=factorTwip;	// pixels * twips/pixels == twips)		    
					newH *= 20;    // this is better ...
				}
				catch( Exception e )
				{
					Logger.logErr( "RowHandle.setRowHeightAutoFit: " + e.toString() );
				}

				h = Math.max( h, newH );
			}
			catch( Exception e )
			{
			}
		}
		if( h > 0 )
		{
			this.myRow.setRowHeight( (int) Math.ceil( h ) );
		}
	}

	/**
	 * Sets the row height in twips (1/20th of a point)
	 *
	 * @param newHeight
	 */
	public void setHeight( int newHeight )
	{
	    /* 20080604 KSC: if an image falls upon this column,
	     * adjust image width so that it does not change
         */
		ArrayList iAdjust = new ArrayList();
		ImageHandle[] images = myRow.getSheet().getImages();
		if( images != null )
		{
			// for each image that falls over this row, trap index + original width -- to be reset after setting row height
			for( int z = 0; z < images.length; z++ )
			{
				ImageHandle ih = images[z];
				int r0 = ih.getRow();
				int r1 = ih.getRow1();
				int row = myRow.getRowNumber();
				if( (row >= r0) && (row <= r1) )
				{
					int h = ih.getHeight();
					iAdjust.add( new int[]{ z, h } );
				}
			}
		}
		myRow.setRowHeight( newHeight );
		for( Object anIAdjust : iAdjust )
		{
			ImageHandle ih = images[((int[]) anIAdjust)[0]];
			ih.setHeight( ((int[]) anIAdjust)[1] );
		}

	}

	/**
	 * Determines if the row passes through
	 * a vertical merge range
	 *
	 * @return
	 */
	public boolean containsVerticalMergeRange()
	{
		CellHandle[] c = this.getCells();
		for( CellHandle aC : c )
		{
			if( aC.getMergedCellRange() != null )
			{
				CellRange cr = aC.getMergedCellRange();
				try
				{
					if( cr.getRows().length > 1 )
					{
						return true;
					}
				}
				catch( Exception e )
				{
				}
				;
			}
		}
		return false;
	}

	/**
	 * sets the default format id for the Row's Cells
	 *
	 * @param int Format Id for all Cells in Row
	 */
	public void setFormatId( int i )
	{
		myRow.setIxfe( i );
	}

	/**
	 * Gets the FormatHandle for this Row.
	 *
	 * @return FormatHandle - a Format object to apply to this Row
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
	 * Set up the format handle for this row
	 */
	private void setFormatHandle()
	{
		if( formatter != null )
		{
			return;
		}
		formatter = new FormatHandle( wbh, this.getFormatId() );
		formatter.setRowHandle( this );
	}

	/**
	 * gets the current default row format id.  May be overwritten by contained cells
	 *
	 * @return format id of row
	 */
	public int getFormatId()
	{
		if( myRow.getExplicitFormatSet() )
		{
			return myRow.getIxfe();
		}
		return this.getWorkBook().getWorkBook().getDefaultIxfe();
	}

	/**
	 * Returns the array of Cells in this Row
	 *
	 * @param cache cellhandles flag
	 * @return Cell[] all Cells in this Row
	 */
	public CellHandle[] getCells( boolean cached )
	{
		Collection ct = myRow.getCells();
		Iterator it = ct.iterator();
		CellHandle[] ch = new CellHandle[ct.size()];
		int t = 0;
		Mulblank aMul = null;
		short c = -1;
		while( it.hasNext() )
		{
			BiffRec rc = (BiffRec) it.next();
			try
			{  // use cache of Cellhandles!
				if( rc.getOpcode() != XLSConstants.MULBLANK )
				{
					ch[t] = this.wsh.getCell( rc.getRowNumber(), rc.getColNumber(), cached );
				}
				else
				{
					// handle Mulblanks: ref a range of cells; to get correct cell address,
					// traverse thru range and set cellhandle ref to correct column
					if( rc == aMul )
					{
						c++;
					}
					else
					{
						aMul = (Mulblank) rc;
						c = (short) aMul.getColFirst();
					}
					ch[t] = this.wsh.getCell( rc.getRowNumber(), c, cached );
				}
			}
			catch( CellNotFoundException cnfe )
			{
				rc.setXFRecord();
				ch[t] = new CellHandle( rc, null );
				ch[t].setWorkSheetHandle( null ); //TODO: implement if causing grief -jm
				if( rc.getOpcode() == XLSConstants.MULBLANK )
				{
					// handle Mulblanks: ref a range of cells; to get correct cell address,
					// traverse thru range and set cellhandle ref to correct column
					if( rc == aMul )
					{
						c++;
					}
					else
					{
						aMul = (Mulblank) rc;
						c = (short) aMul.getColFirst();
					}
					ch[t].setBlankRef( c );    // for Mulblank use only -sets correct column reference for multiple blank cells ...
				}

			}
			t++;
		}
		return ch;
	}

	/**
	 * Returns the array of Cells in this Row
	 *
	 * @return Cell[] all Cells in this Row
	 */
	public CellHandle[] getCells()
	{
		return getCells( false );    // don't use cache
	}

	/**
	 * Get the JSON object for this row.
	 *
	 * @return
	 */
	public String getJSON()
	{
		return getJSON( 255 ).toString();
	}

	public JSONObject getJSON( int maxcols )
	{
		JSONObject theRange = new JSONObject();
		JSONArray cells = new JSONArray();
		try
		{
			theRange.put( JSON_ROW, getRowNumber() );

			theRange.put( JSON_ROW_BORDER_TOP, getHasAnyThickTopBorder() );
			theRange.put( JSON_ROW_BORDER_BOTTOM, getHasAnyBottomBorder() );
			if( getFormatId() != getWorkBook().getWorkBook().getDefaultIxfe() )
			{
				theRange.put( "xf", getFormatId() );
			}
			theRange.put( JSON_HEIGHT, (getHeight() / ROW_HEIGHT_DIVISOR) + 5 ); // the default is TOO SMALL!
			CellHandle[] chandles = getCells( false );
			for( int i = 0; i < chandles.length; i++ )
			{
				CellHandle thisCell = chandles[i];
				if( !thisCell.isDefaultCell() )
				{
					// do NOT use cached formula vals
					if( thisCell.getCell().getOpcode() == XLSRecord.FORMULA )
					{
						try
						{
							FormulaHandle fh = thisCell.getFormulaHandle();
							fh.getFormulaRec().setCachedValue( null );
						}
						catch( Exception ex )
						{
							;
						}
					}

					if( thisCell.getColNum() >= maxcols )
					{
						i = chandles.length;
					}
					else if( thisCell.getCell().getOpcode() == XLSRecord.MULBLANK )
					{
						Mulblank mb = (Mulblank) thisCell.getCell();
						ArrayList<Integer> columns = mb.getColReferences();
						for( Integer column : columns )
						{
							thisCell.setBlankRef( column );
							thisCell.getCell().setCol( column.shortValue() );
							JSONObject result = new JSONObject();
							thisCell.getCellAddress();
							Object v = "";
							try
							{
								v = thisCell.getJSONObject();
							}
							catch( Exception exz )
							{
								Logger.logErr( "Error getting Row cell value " + thisCell.getCellAddress() + " JSON: " + exz );
								v = "ERROR FETCHING VALUE for:" + thisCell.getCellAddress();
							}
							if( v != null )
							{
								result.put( JSON_CELL, v );
								cells.put( result );
							}
						}
					}
					else
					{
						JSONObject result = new JSONObject();
						Object v = "ERROR FETCHING VALUE for:" + thisCell.getCellAddress();
						try
						{
							v = thisCell.getJSONObject();
						}
						catch( Exception exz )
						{
							Logger.logErr( "Error getting Row cell value " + thisCell.getCellAddress() + " JSON: " + exz );
						}
						if( v != null )
						{
							result.put( JSON_CELL, v );
							cells.put( result );
						}
					}
				}
			}
			theRange.put( JSON_CELLS, cells );
		}
		catch( JSONException e )
		{
			Logger.logErr( "Error getting Row JSON: " + e );
		}
		return theRange;
	}

	/**
	 * Returns the String representation of this Row
	 */
	public String toString()
	{
		return myRow.toString();
	}

	/**
	 * Returns the row number of this RowHandle
	 */
	public int getRowNumber()
	{
		return myRow.getRowNumber();
	}

	/**
	 * Set whether the row is collapsed.
	 * Will hide the current row, and all contiguous rows
	 * with the same outline level.
	 *
	 * @param b
	 */
	public void setCollapsed( boolean b )
	{
		myRow.setCollapsed( b );
	}

	/**
	 * Set whether the row is hidden
	 *
	 * @param b
	 */
	public void setHidden( boolean b )
	{
		myRow.setHidden( b );
	}

	/**
	 * Set the Outline level (depth) of the row
	 *
	 * @param x
	 */
	public void setOutlineLevel( int x )
	{
		myRow.setOutlineLevel( x );
	}

	/**
	 * Returns the Outline level (depth) of the row
	 *
	 * @return
	 */
	public int getOutlineLevel()
	{
		return myRow.getOutlineLevel();
	}

	/**
	 * Returns whether the row is collapsed
	 *
	 * @return
	 */
	public boolean isCollapsed()
	{
		return myRow.isCollapsed();
	}

	/**
	 * Returns whether the row is hidden
	 *
	 * @return
	 */
	public boolean isHidden()
	{
		return myRow.isHidden();
	}

	/**
	 * true if row height has been altered from default
	 * i.e. set manually
	 *
	 * @return
	 */
	public boolean isAlteredHeight()
	{
		return myRow.isAlteredHeight();
	}

	public void setBackgroundColor( java.awt.Color colr )
	{
		setFormatHandle();
		formatter.setCellBackgroundColor( colr );
	}

	/**
	 * returns true if there is a Thick Top border set on the row
	 */
	public boolean getHasThickTopBorder()
	{
		return myRow.getHasThickTopBorder();
	}

	/**
	 * returns true if there is a Thick Bottom border set on the row
	 */
	public boolean getHasThickBottomBorder()
	{
		return myRow.getHasThickBottomBorder();
	}

	/**
	 * returns true if there is a thick top or thick or medium bottom border on previoous row
	 * <p/>
	 * Not useful for public API
	 */
	public boolean getHasAnyThickTopBorder()
	{
		return myRow.getHasAnyThickTopBorder();
	}

	/**
	 * Additional space below the row. This flag is set, if the
	 * lower border of at least one cell in this row or if the upper
	 * border of at least one cell in the row below is formatted with
	 * a medium or thick line style. Thin line styles are not taken
	 * into account.
	 * <p/>
	 * Usage of this method is primarily for UI applications, and is not
	 * needed for standard ExtenXLS functionality
	 */
	public boolean getHasAnyBottomBorder()
	{
		return myRow.getHasAnyBottomBorder();
	}

	/**
	 * sets this row to have a thick top border
	 */
	public void setHasThickTopBorder( boolean hasBorder )
	{
		myRow.setHasThickTopBorder( hasBorder );
	}

	/**
	 * sets this row to have a thick bottom border
	 */
	public void setHasThickBottomBorder( boolean hasBorder )
	{
		myRow.setHasThickBottomBorder( hasBorder );
	}

	/**
	 * return the min/max columns defined for this row
	 *
	 * @return
	 */
	public int[] getColDimensions()
	{
		return myRow.getColDimensions();
	}

	public WorkBook getWorkBook()
	{
		return this.wbh;
	}

	public WorkSheetHandle getWorkSheetHandle()
	{
		return this.wsh;
	}
}
