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

package com.extentech.formats.XLS.charts;

import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;

/**
 * <b>Tick: Tick Marks and Labels Format (0x101e)</b>
 * <p/>
 * Offset 	Name		Size	Contents
 * 4		tktMajor	1		Type of major tick mark 0= invisible (none) I = inside of axis line 2 = outside of axis line 3 = cross axis line
 * 5		tktMinor	1		Type of minor tick mark 0= invisible (none)	I = inside of axis line 2 = outside of axis line 3 = cross axis line
 * 6		tit			1		Tick label position relative to axis line 0= invisible (none) 1 = low end of plot area 2 = high end of plot area 3 = next to axis
 * 7		wBkgMode	2		Background mode: I = transparent 2 = opaque
 * 8		rgb			4		Tick-label text color; ROB value, high byte = 0
 * 12		(reserved)	16		Reserved; must be zero
 * 28		grbit		2		Display flags
 * 30		icv			2		Index to color of tick label
 * 32		(reserved)	2		Reserved; must be zero
 * <p/>
 * The grbit field contains the following option flags.
 * <p/>
 * Bits	Mask		Name
 * 0		0xOl		fAutoColor		Automatic text color
 * 1		0x02		fAutoMode 		Automatic text back~
 * 4-2		0xlC		rot				0= no rotation (text appears left-to-right), 1=  text appears top-~~ are upright,
 * 2= text is rotated 90 degrees counterclockwise,  3= text is rotated
 * 5		0x20		fAutoRot		Automatic rotation
 * 7-6		0xCO		(reserved)		Reserved; must be zero
 * 7-0		0xFF		(reserved)		Reserved; must be zero
 */
public class Tick extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 3363212452589555220L;
	byte tktMajor, tktMinor, tit;
	short grbit;
	short rot;

	public void init()
	{
		super.init();
		tktMajor = this.getByteAt( 0 );
		tktMinor = this.getByteAt( 1 );
		tit = this.getByteAt( 2 );
		grbit = ByteTools.readShort( this.getByteAt( 24 ), this.getByteAt( 25 ) );
		// TODO: Finish ops
		rot = (short) ((grbit & 0x1C) >> 2);
	}

	// 20070723 KSC: Need to create new records
	public static XLSRecord getPrototype()
	{
		Tick t = new Tick();
		t.setOpcode( TICK );
		t.setData( t.PROTOTYPE_BYTES );
		t.init();    // important when we parse options ...
		return t;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			2, 0, 3, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 35, 0, 77, 0, 0, 0
	};

	private void updateRecord()
	{
		this.getData()[0] = tktMajor;
		this.getData()[1] = tktMinor;
		this.getData()[2] = tit;
		byte[] b = ByteTools.shortToLEBytes( grbit );
		this.getData()[24] = b[0];
		this.getData()[25] = b[1];
	}

	/**
	 * set generic Tick option
	 * <br>op/val can be one of:
	 * <br>tickLblPos			none, low, high or nextTo
	 * <br>majorTickMark		none, in, out, cross
	 * <br>minorTickMark		none, in, out, cross
	 *
	 * @param op
	 * @param val
	 */
	public void setOption( String op, String val )
	{
		if( op.equals( "tickLblPos" ) )
		{        // val= high (=at high end of perp. axis), low (=at low end of perp. axis), nextTo, none (=no axis labels) Tick
			if( val.equals( "high" ) )
			{
				tit = 2;
			}
			else if( val.equals( "low" ) )
			{
				tit = 1;
			}
			else if( val.equals( "none" ) )
			{
				tit = 0;
			}
			else if( val.equals( "nextTo" ) )
			{
				tit = 3;
			}
		}
		else if( op.equals( "majorTickMark" ) )
		{    // major tick marks (cross, in, none, out)
			if( val.equals( "cross" ) )
			{
				tktMajor = 3;
			}
			else if( val.equals( "in" ) )
			{
				tktMajor = 1;
			}
			else if( val.equals( "out" ) )
			{
				tktMajor = 2;
			}
			else if( val.equals( "none" ) )
			{
				tktMajor = 0;
			}
		}
		else if( op.equals( "minorTickMark" ) )
		{    // minor tick marks (cross, in, none, out)
			if( val.equals( "cross" ) )
			{
				tktMinor = 3;
			}
			else if( val.equals( "in" ) )
			{
				tktMinor = 1;
			}
			else if( val.equals( "out" ) )
			{
				tktMinor = 2;
			}
			else if( val.equals( "none" ) )
			{
				tktMinor = 0;
			}
		}
		updateRecord();
	}
	 /*  4		tktMajor	1		Type of major tick mark 0= invisible (none) I = inside of axis line 2 = outside of axis line 3 = cross axis line
	 *  5		tktMinor	1		Type of minor tick mark 0= invisible (none)	I = inside of axis line 2 = outside of axis line 3 = cross axis line
	 *  6		tit			1		Tick label position relative to axis line 0= invisible (none) 1 = low end of plot area 2 = high end of plot area 3 = next to axis
	 */

	/**
	 * retrieve generic Value axis option as OOXML string
	 * <br>can be one of:
	 * <br>tickLblPos			none, low, high or nextTo
	 * <br>majorTickMark		none, in, out, cross
	 * <br>minorTickMark		none, in, out, cross
	 *
	 * @param op
	 * @return
	 */
	public String getOption( String op )
	{
		if( op.equals( "tickLblPos" ) )
		{        // val= high (=at high end of perp. axis), low (=at low end of perp. axis), nextTo, none (=no axis labels) Tick
			switch( tit )
			{
				case 0:
					return "none";
				case 1:
					return "low";
				case 2:
					return "high";
				case 3:
					return "nextTo";
			}
		}
		else if( op.equals( "majorTickMark" ) )
		{// major tick marks (cross, in, none, out)
			switch( tktMajor )
			{
				case 0:
					return "none";
				case 1:
					return "in";
				case 2:
					return "out";
				case 3:
					return "cross";
			}
		}
		else if( op.equals( "minorTickMark" ) )
		{    // minor tick marks (cross, in, none, out)
			switch( tktMinor )
			{
				case 0:
					return "none";
				case 1:
					return "in";
				case 2:
					return "out";
				case 3:
					return "cross";
			}
		}
		return null;
	}

	/**
	 * returns true if should show minor tick marks
	 *
	 * @return
	 */
	public boolean showMinorTicks()
	{
		return tktMinor != 0;
	}

	/**
	 * returns true if should show major tick marks
	 *
	 * @return
	 */
	public boolean showMajorTicks()
	{
		return tktMajor != 0;
	}

	/**
	 * 0= no rotation (text appears left-to-right),
	 * 1= Text is drawn stacked, top-to-bottom, with the letters upright.
	 * 2= text is rotated 90 degrees counterclockwise,
	 * 3= text is rotated at 90 degrees clockwise.
	 *
	 * @return
	 */
	public short getRotation()
	{
		return rot;
	}

}
