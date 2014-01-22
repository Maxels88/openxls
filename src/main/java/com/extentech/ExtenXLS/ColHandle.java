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
import com.extentech.formats.XLS.Colinfo;
import com.extentech.formats.XLS.Font;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;

import java.util.ArrayList;
import java.util.List;

/**
 * The ColHandle provides access to an Worksheet Column and its Cells.
 * <br>
 * Use the ColHandle to work with individual Columns in an XLS file.
 * <br>
 * With a ColHandle you can:
 * <br><blockquote>
 * get a handle to the Cells in a column<br>
 * set the default formatting for a column<br>
 * <br></blockquote>
 * <p/>
 * Note: for a discussion of Column widths see:
 * http://support.microsoft.com/?kbid=214123
 *
 * @see WorkBookHandle
 * @see WorkSheetHandle
 * @see FormulaHandle
 */
public class ColHandle
{

	// TODO: read 1st font in file to set DEFAULT_ZERO_CHAR_WIDTH ... eventually ...
	public static final double DEFAULT_ZERO_CHAR_WIDTH = 7.0; // width of '0' char in default font + conversion 1.3
	public static final int COL_UNITS_TO_PIXELS = (int) (256 / DEFAULT_ZERO_CHAR_WIDTH);  // = 36.57
	public static final int DEFAULT_COLWIDTH = Colinfo.DEFAULT_COLWIDTH;

	private Colinfo myCol;
	private FormatHandle formatter;
	private WorkBook wbh;
	private WorkSheetHandle mySheet;

	/**
	 * creates a new  ColHandle from a Colinfo Object and reference to a worksheet (WorkSheetHandle Object)
	 *
	 * @param c
	 * @param sheet
	 */
	protected ColHandle( Colinfo c, WorkSheetHandle sheet )
	{
		myCol = c;
		wbh = sheet.getWorkBook();
		mySheet = sheet;
	}

	private int lastsz = 0; // the last checked col width

	/**
	 * resizes this column to fit the width of all displayed, non-wrapped text.
	 * <br>NOTE: as the Excel autofit implementation is undocumented, this is an approximation
	 */
	public void autoFit()
	{
		// KSC: make more betta :)
		double w = 0;
		CellHandle[] cxt = this.getCells();
		for( int t = 0; t < cxt.length; t++ )
		{
			String s = cxt[t].getFormattedStringVal();    //StringVal();
			FormatHandle fh = cxt[t].getFormatHandle();
			Font ef = fh.getFont();
			int style = java.awt.Font.PLAIN;
			if( ef.getBold() )
			{
				style |= java.awt.Font.BOLD;
			}
			if( ef.getItalic() )
			{
				style |= java.awt.Font.ITALIC;
			}
			int h = (int) ef.getFontHeightInPoints();
			java.awt.Font f = new java.awt.Font( ef.getFontName(), style, h );
			double newW = 0;
			if( !cxt[t].getFormatHandle().getWrapText() ) // normal case, no wrap
			{
				newW = StringTool.getApproximateStringWidth( f, s );
			}
			else    // wrap - use current column width?????
			{
				newW = this.getWidth() / COL_UNITS_TO_PIXELS;
			}
			w = Math.max( w, newW );
		/*
		int strlen = cstr.length();
	    int csz = strlen *= cxt[t].getFontSize();
	    int factor= 28;	// KSC: was 50 + added factor to guard below
	    if((csz*factor)>lastsz)
		this.setWidth(csz*factor);
		*/
		}
		if( w == 0 )
		{
			return;    // keep original width ... that's what Excel does for blank columns ...
		}
		// convert pixels to excel column units basically ExtenXLS.COLUNITSTOPIXELS in double form
		this.setWidth( (int) Math.floor( (w / DEFAULT_ZERO_CHAR_WIDTH) * 256.0 ) );
	}

	/**
	 * sets the width of this Column in Characters or Excel units.
	 * <br>
	 * The default Excel column width is set to 8.43 Characters,
	 * based on the default font and font size,
	 * <br>
	 * NOTE: The last Cell in the column having its width
	 * set will be the resulting width of the column
	 *
	 * @param int i - desired Column width in Characters (Excel units)
	 */
	public void setWidthInChars( int newWidth )
	{
        /* if an image falls upon this column, 
         * adjust image width so that it does not change
         */
		ArrayList iAdjust = new ArrayList();
		ImageHandle[] images = myCol.getSheet().getImages();
		if( images != null )
		{
			// for each image that falls over this column, trap index + original width -- to be reset after setting col width
			for( int z = 0; z < images.length; z++ )
			{
				ImageHandle ih = images[z];
				int c0 = ih.getCol();
				int c1 = ih.getCol1();
				int col = myCol.getColFirst();    // should only be one, right?
				if( col >= c0 && col <= c1 )
				{
					int w = ih.getWidth();
					iAdjust.add( new int[]{ z, w } );
				}
			}
		}

		myCol.setColWidthInChars( newWidth );
		for( int z = 0; z < iAdjust.size(); z++ )
		{
			ImageHandle ih = images[((int[]) iAdjust.get( z ))[0]];
			ih.setWidth( ((int[]) iAdjust.get( z ))[1] );
		}
	}

	/**
	 * sets the width of this Column in internal units, described as follows:
	 * <br>
	 * default width of the columns in 1/256 of the width of the zero character,
	 * using default font.
	 * <br>The Default Excel Column, whose width in Characters or Excel Units, is 8.43, has a width in these units of 2300.
	 * <p>NOTE:
	 * The last Cell in the column having its width
	 * set will be the resulting width of the column
	 *
	 * @param int i - desired Column width in internal units
	 */
	public void setWidth( int newWidth )
	{
        /* if an image falls upon this column, 
         * adjust image width so that it does not change
         */
		ArrayList iAdjust = new ArrayList();
		ImageHandle[] images = myCol.getSheet().getImages();
		if( images != null )
		{
			// for each image that falls over this column, trap index + original width -- to be reset after setting col width  
			for( int z = 0; z < images.length; z++ )
			{
				ImageHandle ih = images[z];
				int c0 = ih.getCol();
				int c1 = ih.getCol1();
				int col = myCol.getColFirst();    // should only be one, right?
				if( col >= c0 && col <= c1 )
				{
					int w = ih.getWidth();
					iAdjust.add( new int[]{ z, w } );
				}
			}
		}
		lastsz = newWidth;
		myCol.setColWidth( newWidth );
		// now adjust any of the images that we noted above
		for( int z = 0; z < iAdjust.size(); z++ )
		{
			ImageHandle ih = images[((int[]) iAdjust.get( z ))[0]];
			ih.setWidth( ((int[]) iAdjust.get( z ))[1] );
		}
	}

	/**
	 * returns the width of this Column in internal units
	 * defined as follows:
	 * <br>
	 * default width of the columns in 1/256 of the width of the zero character,
	 * using default font.
	 * <br>The Default Excel Column, whose width in Excel Units or Characters is 8.43, has a width in these units of 2300.
	 *
	 * @return int Column width in internal units
	 */
	public int getWidth()
	{
		return myCol.getColWidth();
	}

	/**
	 * returns the width of this Column in Characters or regular Excel units
	 * <br>NOTE: this value is a calculated value that should be close but still is an approximation of Excel units
	 *
	 * @return int Column width in Excel units
	 */
	public int getWidthInChars()
	{
		return myCol.getColWidthInChars();
	}

	/**
	 * static utility method to return the Column width of an existing column
	 * in the units as follows:
	 * <br>
	 * default width of the columns in 1/256 of the width of the zero character,
	 * using default font.
	 * <br>For Arial 10 point, the default width of the zero character = 7
	 * <br>The Default Excel Column, whose width in Characters or Excel Units is 8.43, has a width in these units of 2300.
	 *
	 * @param Boundsheet sheet - source Worksheet
	 * @param int        col - 0-based Column number
	 * @return int - Column width in internal units
	 */
	public static int getWidth( Boundsheet sheet, int col )
	{
		int w = Colinfo.DEFAULT_COLWIDTH;
		try
		{
			Colinfo c = sheet.getColInfo( col );
			if( c != null )
			{
				w = (int) c.getColWidth();
			}
		}
		catch( Exception e )
		{    // exception if no col defined
		}
		return w;
	}

	/**
	 * sets the format id (an index to a Format record) for this Column
	 * <br>This sets the default formatting for the Column
	 * such that any cell that does not specifically set it's own formatting
	 * will display this Column formatting
	 *
	 * @param int i - ID representing the Format to set this Column
	 * @see FormatHandle
	 */
	public void setFormatId( int i )
	{
		myCol.setIxfe( i );
	}

	/**
	 * returns the format ID (the index to the format record) for this Column
	 * <br>The Column format is the default formatting for each cell contained
	 * within the column
	 *
	 * @return int formatId - the index of the format record for this Column
	 * @see FormatHandle
	 */
	public int getFormatId()
	{
		return myCol.getIxfe();
	}

	/**
	 * returns the FormatHandle (a Format Object describing visual properties) for this Column
	 * <br>NOTE: The Column format record describes the default formatting for each cell contained
	 * within the column
	 *
	 * @return FormatHandle - a Format object to apply to this Col
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
	 * sets the FormatHandle (a Format Object describing visual properties) for this Column
	 * <br>NOTE: The Column format record describes the default formatting for each cell contained
	 * within the column
	 */
	private void setFormatHandle()
	{
		if( formatter != null )
		{
			return;
		}
		formatter = new FormatHandle( wbh, this.getFormatId() );
		formatter.setColHandle( this );
	}

	/**
	 * returns the first Column referenced by this column handle
	 * <br>NOTE: A Column handle may in some circumstances refer to a range of columns
	 *
	 * @return int first column number referenced by this Column handle
	 */
	public int getColFirst()
	{
		return myCol.getColFirst();
	}

	/**
	 * returns the last Column referenced by this column handle
	 * <br>NOTE: A Column handle may in some circumstances refer to a range of columns
	 *
	 * @return int last column number referenced by this Column handle
	 */
	public int getColLast()
	{
		return myCol.getColLast();
	}

	/**
	 * returns the array of Cells in this Column
	 *
	 * @return CellHandle array
	 */
	public CellHandle[] getCells()
	{
		List mycells;
		try
		{
			mycells = this.mySheet.getBoundsheet().getCellsByCol( this.getColFirst() );
		}
		catch( CellNotFoundException e )
		{
			return new CellHandle[0];
		}
		CellHandle[] ch = new CellHandle[mycells.size()];
		for( int t = 0; t < ch.length; t++ )
		{
			Object o = mycells.get( t );
			Logger.logInfo( "getCells() - processing index " + t + ", " + ((o == null) ? "<NULL>" : o.getClass().getName()) );
			ch[t] = new CellHandle( (BiffRec) o, null );
			ch[t].setWorkSheetHandle( null );
		}
		return ch;
	}

	/**
	 * determines if this Column passes through i.e. contains a
	 * horizontal merge range
	 *
	 * @return true if this Column is part of any merge (horizontally merged cells)
	 */
	public boolean containsMergeRange()
	{
		RowHandle[] r = mySheet.getRows();
		for( int i = 0; i < r.length; i++ )
		{
			BiffRec b;
			try
			{
				b = r[i].myRow.getCell( (short) this.getColFirst() );
				if( b != null && b.getMergeRange() != null )
				{
					return true;
				}
			}
			catch( CellNotFoundException e )
			{
			}

		}
		return false;
	}

	/**
	 * sets whether to collapse this Column
	 *
	 * @param boolean b - true to collapse this Column
	 */
	public void setCollapsed( boolean b )
	{
		this.myCol.setCollapsed( b );
	}

	/**
	 * sets whether to hide or show this Column
	 *
	 * @param boolean b - true to hide this Column, false to show
	 */
	public void setHidden( boolean b )
	{
		this.myCol.setHidden( b );
	}

	/**
	 * Set the Outline level (depth) of this Column
	 *
	 * @param int x - outline level
	 */
	public void setOutlineLevel( int x )
	{
		this.myCol.setOutlineLevel( x );
	}

	/**
	 * Returns the Outline level (depth) of this Column
	 *
	 * @return int outline level
	 */
	public int getOutlineLevel()
	{
		return myCol.getOutlineLevel();
	}

	/**
	 * returns true if this Column is collapsed
	 *
	 * @return true if ths Column is collapsed, false otherwise
	 */
	public boolean isCollapsed()
	{
		return myCol.isCollapsed();
	}

	/**
	 * returns true if this Column is hidden
	 *
	 * @return true if this Column is hidden, false if not
	 */
	public boolean isHidden()
	{
		return myCol.isHidden();
	}
}