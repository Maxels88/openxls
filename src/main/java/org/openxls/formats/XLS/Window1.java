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
package org.openxls.formats.XLS;

import org.openxls.toolkit.ByteTools;

/**
 * <b>WINDOW1 0x3D: Contains window attributes for a Workbook.</b><br>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       xWn         2       Horizontal Position of the window
 * 6       yWn         2       Vertical Position of the window
 * 8       dxWn        2       Width of the window
 * 10      dyWn        2       Height of the window
 * 12      grbit       2       Option Flags
 * 14      itabCur     2       Index of the selected workbook tab (0 based)
 * 16      itabFirst   2       Index of the first displayed workbook tab (0 based)
 * 18      ctabSel     2       Number of workbook tabs that are selected
 * 20      wTabRatio   2       Ratio of the width of the workbook tabs to the width of
 * the horizontal scroll bar; to obtain the ratio, convert to
 * decimal and then divide by 1000
 * </p></pre>
 *
 * @see WorkBook
 * @see BOUNDSHEET
 * @see INDEX
 * @see DBCELL
 * @see ROW
 * @see Cell
 * @see XLSRecord
 */

public class Window1 extends XLSRecord
{

	/**
	 *
	 */
	private static final long serialVersionUID = 2770922305028029883L;
	short xWn = 0;
	short yWn = 0;
	short dxWn = 0;
	short dyWn = 0;
	short grbit = 0;
	short itabCur = 0;
	short itabFirst = 0;
	short ctabSel = 0;
	short wTabRatio = 0;
	Boundsheet mybs = null;

	public int getCurrentTab()
	{
		return itabCur;
	}

	/**
	 * Sets the current tab that is displayed on opening.
	 * Note, this is not really the same thing as "selected".
	 * The selected parameter is from the Window2 record.  As we
	 * don't really have much need to select more than one sheet
	 * on output, this method just delselects every other sheet than
	 * the one that is passed in, and selects that one in it's Window2
	 *
	 * @param bs
	 */
	public void setCurrentTab( Boundsheet bs )
	{
		mybs = bs;
		int t = mybs.getSheetNum();
		Boundsheet[] bounds = getWorkBook().getWorkSheets();
		for( Boundsheet bound : bounds )
		{
			bound.getWindow2().setSelected( false );
		}
		mybs.getWindow2().setSelected( true );
		byte[] mydata = getData();
		itabCur = (short) t;
		byte[] tabbytes = ByteTools.shortToLEBytes( (short) t );
		mydata[10] = tabbytes[0];
		mydata[11] = tabbytes[1];
	}

	/**
	 * Sets which tab will display furthest to the left in the workbook.  Sheets that have
	 * their tabid before this one will be 'pushed off' to the left.  They can be retrieved in the
	 * GUI by clicking the left arrow next to the displayed worksheets.
	 *
	 * @param t
	 */
	public void setFirstTab( int t )
	{
		byte[] mydata = getData();
		itabFirst = (short) t;
		byte[] tabbytes = ByteTools.shortToLEBytes( (short) t );
		mydata[12] = tabbytes[0];
		mydata[13] = tabbytes[1];
	}

	/**
	 * Default init method
	 */
	@Override
	public void init()
	{
		super.init();
		xWn = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		yWn = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		dxWn = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		dyWn = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		grbit = ByteTools.readShort( getByteAt( 8 ), getByteAt( 9 ) );
		itabCur = ByteTools.readShort( getByteAt( 10 ), getByteAt( 11 ) );
		itabFirst = ByteTools.readShort( getByteAt( 12 ), getByteAt( 13 ) );
		ctabSel = ByteTools.readShort( getByteAt( 14 ), getByteAt( 15 ) );
		wTabRatio = ByteTools.readShort( getByteAt( 16 ), getByteAt( 17 ) );

	}

	/**
	 * Returns whether the sheet selection tabs should be shown.
	 */
	public boolean showSheetTabs()
	{
		return (grbit & 0x20) == 0x20;
	}

	/**
	 * Sets whether the sheet selection tabs should be shown.
	 */
	public void setShowSheetTabs( boolean show )
	{
		if( show )
		{
			grbit |= 0x20;
		}
		else
		{
			grbit &= ~0x20;
		}
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[8] = b[0];
		getData()[9] = b[1];
	}
}