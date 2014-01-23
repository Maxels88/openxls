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
package com.extentech.formats.XLS;

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <b>Colinfo: Column Formatting Information (7Dh)</b><br>
 * <p/>
 * Colinfo describes the formatting for a column range
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---------------------------------------------------------------
 * 4       colFirst        2       First formatted column (0)
 * 6       colLast         2       Last formatted column (0)
 * 8       colWidth        2       Column width in 1/256 character units
 * 10      ixfe            2       Index to XF for columns
 * 12      grbit           2       Options
 * 14      reserved        1       Must be zero
 * <p/>
 * <p/>
 * grib options
 * offset	Bits	mask	name		contents
 * ---------------------------------------------------------------
 * 0		0		01h		fHidden		=1 if the column range is hidden
 * 7-1		FEh		UNUSED
 * 1		2-0		07h		iOutLevel	Outline Level of the column
 * 3		08h		Reserved	must be zero
 * 4		10h		iCollapsed	=1 if the col is collapsed in outlining
 * 7-5		E0h		Reserved	must be zero
 * <p/>
 * etc.
 * <p/>
 * Note: for a discussion of Column widths see:
 * http://support.microsoft.com/?kbid=214123
 * </p></pre>
 */

public final class Colinfo extends XLSRecord implements ColumnRange
{
	private static final Logger log = LoggerFactory.getLogger( Colinfo.class );
	private static final long serialVersionUID = 3048724897018541459L;
	public static final int DEFAULT_COLWIDTH = 2340;    // why 2000???? excel reports 2340 ...?
	private int colFirst, colLast, colWidth;
	private short grbit;
	private boolean collapsed, hidden;
	private int outlineLevel = 0;

	/**
	 * set last/first cols/rows
	 */
	public void setColFirst( int c )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = this.getData();
		System.arraycopy( b, 0, dt, 0, 2 );
		this.colFirst = c;
	}

	@Override
	public int getColFirst()
	{
		return colFirst;
	}

	/**
	 * set last/first cols/rows
	 */
	public void setColLast( int c )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) c );
		byte[] dt = this.getData();
		System.arraycopy( b, 0, dt, 2, 2 );
		this.colLast = c;
	}

	@Override
	public int getColLast()
	{
		return colLast;
	}

	/**
	 * Shifts the whole colinfo over the amount of the offset
	 */
	public void moveColInfo( int offset )
	{
		this.setColFirst( this.getColFirst() + offset );
		this.setColLast( this.getColLast() + offset );
	}

	@Override
	public void init()
	{
		super.init();
		colFirst = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		colLast = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		colWidth = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		ixfe = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
		grbit = ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
		decodeGrbit();
			log.trace( "Col: " + ExcelTools.getAlphaVal( colFirst ) + "-" + ExcelTools.getAlphaVal( colLast ) + "  ixfe: " + String.valueOf(
					ixfe ) + " width: " + colWidth );
	}

	public static Colinfo getPrototype( int colF, int colL, int wide, int formatIdx )
	{
		Colinfo ret = new Colinfo();

		ret.originalsize = 12;
		// ret.setLabel("COLINFO");
		ret.setOpcode( COLINFO );
		ret.setLength( (short) ret.originalsize );
		byte[] newbytes = new byte[ret.originalsize];
		// colF
		byte[] b = ByteTools.shortToLEBytes( (short) colF );
		newbytes[0] = b[0];
		newbytes[1] = b[1];

		// colL
		b = ByteTools.shortToLEBytes( (short) colL );
		newbytes[2] = b[0];
		newbytes[3] = b[1];

		// wide
		b = ByteTools.shortToLEBytes( (short) wide );
		newbytes[4] = b[0];
		newbytes[5] = b[1];

		// XF
		b = ByteTools.shortToLEBytes( (short) formatIdx );
		newbytes[6] = b[0];
		newbytes[7] = b[1];

		ret.setData( newbytes );
		ret.init();
		return ret;
	}

	/**
	 * Is this colinfo based on a single column format?
	 */
	public boolean isSingleColColinfo()
	{
		if( getColFirst() == getColLast() )
		{
			return true;
		}
		return false;
	}

	/**
	 * returns whether a given col
	 * is referenced by this Colinfo
	 */
	public boolean inrange( int x )
	{
		if( (x <= colLast) && (x >= colFirst) )
		{
			return true;
		}
		return false;
	}

	/**
	 * Set the width of a column or columns in Excel units
	 *
	 * @param double x	new column width in Excel Units
	 */
	private static double fudgefactor = 0.711;    // calc needs fudge factor to equal excel values !!!

	public void setColWidthInChars( double x )
	{
		// it's a value that needs to be converted to the appropriate units
		x = (int) (x + fudgefactor) * 256.0;
		byte[] cl = ByteTools.shortToLEBytes( (short) x );
		System.arraycopy( cl, 0, this.getData(), 4, 2 );
		colWidth = ByteTools.readShort( this.getData()[4], this.getData()[5] );
		// 20060609 KSC: APPEARS THAT grbit=0 means default column width so must set to either 2 or 6 ---
		// there is NO documentation on this!
		if( grbit == 0 )
		{
			setGrbit( 2 );
		}
	}

	/**
	 * Set the width of a column or columns in internal units
	 * <br>Internal units are units of (defaultfontwidth/256)
	 * <br>Excel column width= default font width/256 * 8.43
	 * <br>More specifically, column width is in 1/256 of the width of the zero character,
	 * using default font (first FONT record in the file)
	 *
	 * @param int x		new column width
	 */
	public void setColWidth( int x )
	{
		byte[] cl = ByteTools.shortToLEBytes( (short) x );
		System.arraycopy( cl, 0, this.getData(), 4, 2 );
		colWidth = ByteTools.readUnsignedShort( this.getData()[4], this.getData()[5] );
		// 20060609 KSC: APPEARS THAT grbit=0 means default column width so must set to either 2 or 6 ---
		// there is NO documentation on this!
		if( grbit == 0 )
		{
			setGrbit( 2 );
		}
	}

	/**
	 * Get the width of a column in internal units
	 * <br>Internal units are units of (defaultfontwidth/256)
	 * <br>Excel column width= default font width/256 * 8.43
	 * <br>More specifically, column width is in 1/256 of the width of the zero character,
	 * using default font (first FONT record in the file)
	 */
	public int getColWidth()
	{
		return colWidth;

	}

	/**
	 * returns Column Width in Chars or Excel-units
	 *
	 * @return
	 */
	public int getColWidthInChars()
	{
		int colwidth = this.getColWidth();
		colwidth = (int) Math.round( ((colwidth - fudgefactor) / 256) );
		return colwidth;
	}

	/**
	 * Flag indicating if the outlining of the affected column(s) is in the collapsed state.
	 *
	 * @param b boolean true if collapsed
	 */
	public void setCollapsed( boolean b )
	{
		this.collapseIt( b );
		// all previous columns are hidden
		for( int i = 0; i < this.colFirst; i++ )
		{
			Colinfo r = this.getSheet().getColInfo( i );
			if( (r != null) && (r.getOutlineLevel() == this.getOutlineLevel()) )
			{
				r.setHidden( b );
			}
		}
	}

	/**
	 * collapse it is called internally, as you need to call the next column from
	 * the colinfo.
	 *
	 * @param b
	 */
	private void collapseIt( boolean b )
	{
		this.collapsed = b;
		updateGrbit();
	}

	/**
	 * Set whether the column is hidden
	 *
	 * @param b
	 */
	public void setHidden( boolean b )
	{
		this.hidden = b;
		updateGrbit();
		this.getWorkBook().getUsersviewbegin().setDisplayOutlines( true );
	}

	/**
	 * Set the Outline level (depth) of the column
	 *
	 * @param x
	 */
	public void setOutlineLevel( int x )
	{
		this.outlineLevel = x;
		updateGrbit();
		this.getSheet().getGuts().setColGutterSize( 10 + (10 * x) );
		this.getSheet().getGuts().setMaxColLevel( x + 1 );
	}

	/**
	 * This should be run at init()
	 * in order to populate grbit values
	 */
	private void decodeGrbit()
	{
		byte[] grbytes = ByteTools.shortToLEBytes( grbit );
		if( (grbytes[0] & 0x1) == 0x1 )
		{
			hidden = true;
		}
		else
		{
			hidden = false;
		}
		if( (grbytes[1] & 0x10) == 0x10 )
		{
			collapsed = true;
		}
		else
		{
			collapsed = false;
		}
		outlineLevel = (grbytes[1] & 0x7);
	}

	public int getGrbit()
	{
		return grbit;
	}

	public void setGrbit( int grbit )
	{
		this.grbit = (short) grbit;
		updateGrbit();
	}

	/**
	 * set the grbit to match
	 * whatever values have been passed in to
	 * modify the grbit functions.  It also updates
	 * underlying byte record of the XLSRecord.
	 */
	public void updateGrbit()
	{
		byte[] grbytes = ByteTools.shortToLEBytes( grbit );
		// set whether collapsed or not
		if( collapsed )
		{
			grbytes[1] = (byte) (0x10 | grbytes[1]);
		}
		else
		{
			grbytes[1] = (byte) (0xEF & grbytes[1]);
		}
		// set if hidden
		if( hidden )
		{
			grbytes[0] = (byte) (0x1 | grbytes[0]);
		}
		else
		{
			grbytes[0] = (byte) (0xFE & grbytes[0]);
		}
		// set the outline level
		grbytes[1] = (byte) (outlineLevel | grbytes[1]);
		// reset the grbit and the body rec
		grbit = ByteTools.readShort( grbytes[0], grbytes[1] );
		byte[] recdata = this.getData();
		recdata[8] = grbytes[0];
		recdata[9] = grbytes[1];
		this.setData( recdata );
	}

	/**
	 * Returns the Outline level (depth) of the column
	 *
	 * @return
	 */
	public int getOutlineLevel()
	{
		return outlineLevel;
	}

	@Override
	public int getIxfe()
	{
		return ixfe;
	}

	/**
	 * Returns whether the column is collapsed
	 *
	 * @return
	 */
	public boolean isCollapsed()
	{
		return collapsed;
	}

	/**
	 * Returns whether the column is hidden
	 *
	 * @return
	 */
	public boolean isHidden()
	{
		return hidden;
	}

	/**
	 * Sets the Ixfe for this record.  For some stupid reason this is the *only* xls record
	 * that has it's ixfe in a different place.  Intern time at microsoft I guess.
	 */
	@Override
	public void setIxfe( int i )
	{
		this.ixfe = i;
		byte[] newxfe = ByteTools.cLongToLEBytes( i );
		byte[] b = this.getData();

		System.arraycopy( newxfe, 0, b, 6, 2 );
		this.setData( b );
	}

	public String toString()
	{
		return "ColInfo: " + ExcelTools.getAlphaVal( colFirst ) + "-" + ExcelTools.getAlphaVal( colLast ) + "  ixfe: " + String.valueOf(
				ixfe ) + " width: " + colWidth;
	}

	@Override
	public boolean isSingleCol()
	{
		return (this.getColFirst() == this.getColLast());

	}
}