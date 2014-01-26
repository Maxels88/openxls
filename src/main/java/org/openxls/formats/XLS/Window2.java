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
 * <b>WINDOW2 0x23E: Contains window attributes for a Sheet.</b><br>
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       grbit       2       Option Flags
 * 6       rwTop       2       Top row visible in the window
 * 8       colLeft     2       Leftmost visible col in window
 * 10      icvHdr      4       Index to color val for row/col headings & grids
 * 14      wScaleSLV   2       Zoom mag in page break preview
 * 16      wScaleNorm  2       Zoom mag in Normal preview
 * 18      reserved    4
 * <p/>
 * grbit Option flags
 * 0 0001H 0 = Show formula results 1 = Show formulas
 * 1 0002H 0 = Do not show grid lines 1 = Show grid lines
 * 2 0004H 0 = Do not show sheet headers 1 = Show sheet headers
 * 3 0008H 0 = Panes are not frozen 1 = Panes are frozen (freeze)
 * 4 0010H 0 = Show zero values as empty cells 1 = Show zero values
 * 5 0020H 0 = Manual grid line colour 1 = Automatic grid line colour
 * 6 0040H 0 = Columns from left to right 1 = Columns from right to left
 * 7 0080H 0 = Do not show outline symbols 1 = Show outline symbols
 * 8 0100H 0 = Keep splits if pane freeze is removed 1 = Remove splits if pane freeze is removed
 * 9 0200H 0 = Sheet not selected 1 = Sheet selected (BIFF5-BIFF8)
 * 10 0400H 0 = Sheet not visible 1 = Sheet visible (BIFF5-BIFF8)
 * 11 0800H 0 = Show in normal view 1 = Show in page break preview (BIFF8)    </p></pre>
 *
 * @see WorkBook
 * @see BOUNDSHEET
 * @see INDEX
 * @see DBCELL
 * @see ROW
 * @see Cell
 * @see XLSRecord
 */

public class Window2 extends XLSRecord
{
	/**
	 *
	 */
	private static final long serialVersionUID = -8316509425117672619L;
	int grbit = -1;
	int rwTop = -1;
	int colLeft = -1;
	int icvHdr = -1;
	int wScaleSLV = -1;
	int wScaleNorm = -1;

	//20060308 KSC: Added for get/set access to Window2 options
	static final int BITMASK_SHOWFORMULARESULTS = 0x0001;
	static final int BITMASK_SHOWGRIDLINES = 0x0002;
	static final int BITMASK_SHOWSHEETHEADERS = 0x0004;
	static final int BITMASK_FREEZEPANES = 0x0008;
	static final int BITMASK_SHOWZEROVALUES = 0x0010;
	static final int BITMASK_GRIDLINECOLOR = 0x0020;
	static final int BITMASK_COLUMNDIRECTION = 0x0040;
	static final int BITMASK_SHOWOUTLINESYMBOLS = 0x0080;
	static final int BITMASK_KEEPSPLITS = 0x0100;
	static final int BITMASK_SHEETSELECTED = 0x0200;
	static final int BITMASK_SHEETVISIBLE = 0x0400;
	static final int BITMASK_SHOWINPRINTPREVIEW = 0x0800;

	@Override
	public void init()
	{
		super.init();
		short s1;
		short s2;

		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		rwTop = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );
		colLeft = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );

		s1 = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		s2 = ByteTools.readShort( getByteAt( 8 ), getByteAt( 9 ) );
		icvHdr = ByteTools.readInt( s1, s2 );

		// the following do not necessarily exist in VB-manipulated windows
		if( getLength() > 10 )
		{
			wScaleSLV = ByteTools.readShort( getByteAt( 10 ), getByteAt( 11 ) );
			wScaleNorm = ByteTools.readShort( getByteAt( 12 ), getByteAt( 13 ) );
		}
	}

	void setSelected( boolean b )
	{
		if( b )
		{
			getData()[1] |= 0x2;
		}
		else
		{
			getData()[1] &= 0xFD;
		}
		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
	}

	/**
	 * returns the String Address of first visible cell on the sheet
	 *
	 * @return String
	 */
	public String getTopLeftCell()
	{
		return ExcelTools.formatLocation( new int[]{ rwTop, colLeft } );
	}

	// add get/set for Window2 options
	public void setGrbit()
	{
		byte[] data = getData();
		byte[] b = ByteTools.shortToLEBytes( (short) grbit );
		System.arraycopy( b, 0, data, 0, 2 );
		setData( data );

	}

	public boolean getShowFormulaResults()
	{
		return ((grbit & BITMASK_SHOWFORMULARESULTS) == BITMASK_SHOWFORMULARESULTS);
	}

	public void setShowFormulaResults( boolean b )
	{
		if( b )
		{
			grbit = grbit | BITMASK_SHOWFORMULARESULTS;
		}
		else
		{
			grbit = grbit & ~BITMASK_SHOWFORMULARESULTS;
		}
		setGrbit();
	}

	public boolean getShowGridlines()
	{
		return ((grbit & BITMASK_SHOWGRIDLINES) == BITMASK_SHOWGRIDLINES);
	}

	public void setShowGridlines( boolean b )
	{
		if( b )
		{
			grbit = grbit | BITMASK_SHOWGRIDLINES;
		}
		else
		{
			grbit = grbit & ~BITMASK_SHOWGRIDLINES;
		}
		setGrbit();
	}

	public boolean getShowSheetHeaders()
	{
		return ((grbit & BITMASK_SHOWSHEETHEADERS) == BITMASK_SHOWSHEETHEADERS);
	}

	public void setShowSheetHeaders( boolean b )
	{
		if( b )
		{
			grbit = grbit | BITMASK_SHOWSHEETHEADERS;
		}
		else
		{
			grbit = grbit & ~BITMASK_SHOWSHEETHEADERS;
		}
		setGrbit();
	}

	public int getScaleNorm()
	{
		//return wScaleSLV;
		return wScaleNorm;
	}

	public void setScaleNorm( int zm )
	{
		wScaleNorm = zm;
		byte[] data = getData();
		byte[] b = ByteTools.shortToLEBytes( (short) zm );
		// wScaleSLV 10,11
		// wScaleNorm 12,13;
		System.arraycopy( b, 0, data, 12, 2 );
		setData( data );
	}

	public boolean getShowZeroValues()
	{
		return ((grbit & BITMASK_SHOWZEROVALUES) == BITMASK_SHOWZEROVALUES);
	}

	public void setShowZeroValues( boolean b )
	{
		if( b )
		{
			grbit = grbit | BITMASK_SHOWZEROVALUES;
		}
		else
		{
			grbit = grbit & ~BITMASK_SHOWZEROVALUES;
		}
		setGrbit();
	}

	public boolean getShowOutlineSymbols()
	{
		return ((grbit & BITMASK_SHOWOUTLINESYMBOLS) == BITMASK_SHOWOUTLINESYMBOLS);
	}

	public void setShowOutlineSymbols( boolean b )
	{
		if( b )
		{
			grbit = grbit | BITMASK_SHOWOUTLINESYMBOLS;
		}
		else
		{
			grbit = grbit & ~BITMASK_SHOWOUTLINESYMBOLS;
		}
		setGrbit();
	}

	/**
	 * true if sheet is in normal view mode (false if in page break preview mode)
	 *
	 * @return
	 */
	public boolean getShowInNormalView()
	{
		return ((grbit & BITMASK_SHOWINPRINTPREVIEW) == 0);
	}

	public void setShowInNormalView( boolean b )
	{
		if( b )
		{
			grbit = grbit & ~BITMASK_SHOWINPRINTPREVIEW;
		}
		else
		{
			grbit = grbit | BITMASK_SHOWINPRINTPREVIEW;
		}
		setGrbit();
	}

	public boolean getFreezePanes()
	{
		return ((grbit & BITMASK_FREEZEPANES) == BITMASK_FREEZEPANES);
	}

	public void setFreezePanes( boolean b )
	{
		if( b )
		{
			grbit = grbit | BITMASK_FREEZEPANES;
		}
		else
		{
			grbit = grbit & ~BITMASK_FREEZEPANES;
		}
		setGrbit();
	}

	public boolean getManualGridLineColor()
	{
		return ((grbit & BITMASK_GRIDLINECOLOR) == BITMASK_GRIDLINECOLOR);
	}

	public void setManualGridLineColor( boolean b )
	{
		if( b )
		{
			grbit = grbit | BITMASK_GRIDLINECOLOR;
		}
		else
		{
			grbit = grbit & ~BITMASK_GRIDLINECOLOR;
		}
		setGrbit();
	}
	/*	// TODO: finish these
	static final int BITMASK_COLUMNDIRECTION= 0x0040;
    static final int BITMASK_KEEPSPLITS= 0x0100;
    static final int BITMASK_SHEETSELECTED= 0x0200;
    static final int BITMASK_SHEETVISIBLE= 0x0400;
*/

}