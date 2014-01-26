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

import org.openxls.ExtenXLS.ExcelTools;
import org.openxls.toolkit.ByteTools;

/**
 * <b>Pane: Stores the position of window panes. (41h)</b>
 * <br>If the sheet doesn't contain any splits, this record will not occur
 * A sheet can be split in two different ways, with unfrozen or frozen panes.
 * A flag in the WINDOW2 record specifies if the panes are frozen.
 * <p/>
 * <p/>
 * <p><pre>
 * offset  size name		contents
 * ---
 * 0 	2	px 			Position of the vertical split (px, 0 = No vertical split):
 * Unfrozen pane: Width of the left pane(s) (in twips = 1/20 of a point)
 * Frozen pane: Number of visible columns in left pane(s)
 * 2 	2 	py			Position of the horizontal split (py, 0 = No horizontal split):
 * Unfrozen pane: Height of the top pane(s) (in twips = 1/20 of a point)
 * Frozen pane: Number of visible rows in top pane(s)
 * 4 	2 	visRow		Index to first visible row in bottom pane(s)
 * 6 	2	visCol 		Index to first visible column in right pane(s)
 * 8 	1 	pActive		Identifier of pane with active cell cursor (see below)
 * [9] 1 				Not used (BIFF5-BIFF8 only, not written in BIFF2-BIFF4)
 * <p/>
 * If the panes are frozen, pane 0 is always active, regardless of the cursor position. The correct identifiers for all possible
 * combinations of visible panes are shown in the following pictures.
 * px = 0, py = 0		px = 0, py > 0		px > 0, py = 0		px > 0, py > 0
 * 3				3					3 1					3 1
 * 2										2 0
 * <p/>
 * </p></pre>
 */

public final class Pane extends XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5314818835334217157L;
	short px;
	short py;
	short visRow;
	short visCol;
	byte pActive;
	boolean bFrozen;
	Window2 win2;

	@Override
	public void init()
	{
		super.init();
		px = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		;
		py = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		;
		visRow = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		;
		visCol = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		;
		pActive = getByteAt( 8 );

	}

	private byte[] PROTOTYPE_BYTES = { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

	public static XLSRecord getPrototype()
	{
		Pane p = new Pane();
		p.setOpcode( PANE );
		p.setData( p.PROTOTYPE_BYTES );
		p.init();
		return p;
	}

	/**
	 * sets the 1st visible column height in twips + sets frozen panes off
	 *
	 * @param col
	 * @param nCols
	 */
	public void setSplitColumn( int col, int nCols )
	{
		setFrozenColumn( col );
		win2.setFreezePanes( false );
		visCol = (short) col;
		px = (short) nCols;
		py = 0;
		pActive = 3;        // TODO: figure this out!
		updateData();
	}

	/**
	 * sets split columm + freezes panes
	 *
	 * @param col
	 */
	public void setFrozenColumn( int col )
	{
		win2.setFreezePanes( true );
		visCol = (short) col;
		px = (short) col;
		pActive = 0;
		updateData();
	}

	/**
	 * sets the first visible row + width in twips + sets frozen panes off
	 *
	 * @param row to start split
	 * @param rsz row height in twips
	 */
	public void setSplitRow( int row, int rsz )
	{
		setFrozenRow( row );
		win2.setFreezePanes( false );
		visRow = (short) row;
		py = (short) rsz;
		px = 0;
		pActive = 3;        // TODO: figure this out!
		updateData();
	}

	/**
	 * Gets the first visible row of the split pane
	 * this is 0 based
	 *
	 * @return
	 */
	public int getVisibleRow()
	{
		return visRow;
	}

	/**
	 * gets the first visible col of the split pane
	 * this is 0 based
	 *
	 * @return int 0-based first visible column of the split
	 */
	public int getVisibleCol()
	{
		return visCol;
	}

	/**
	 * Returns the position of the row split in twips
	 */
	public int getRowSplitLoc()
	{
		return py;
	}

	/**
	 * sets the split row + freezes pane
	 *
	 * @param row
	 */
	public void setFrozenRow( int row )
	{
		win2.setFreezePanes( true );
		visRow = (short) row;
		py = (short) row;
		pActive = 0;
		updateData();
	}

	/**
	 * return the address of the TopLeft visible Cell
	 *
	 * @return
	 */
	public String getTopLeftCell()
	{
		return ExcelTools.formatLocation( new int[]{ visRow, visCol } );
	}

	protected void updateData()
	{
		getData();
		byte[] b = ByteTools.shortToLEBytes( px );
		data[0] = b[0];
		data[1] = b[1];
		b = ByteTools.shortToLEBytes( py );
		data[2] = b[0];
		data[3] = b[1];
		b = ByteTools.shortToLEBytes( visRow );
		data[4] = b[0];
		data[5] = b[1];
		b = ByteTools.shortToLEBytes( visCol );
		data[6] = b[0];
		data[7] = b[1];
		data[8] = pActive;
	}

	public void setWindow2( Window2 win2 )
	{
		this.win2 = win2;
	}
}