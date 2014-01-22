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
 * <b>CatserRange: Defines a Category or Series Axis (0x1020)</b>
 * <p/>
 * 4		catCross	2	Value axis/category crossing point (2-D charts only)
 * If fMaxCross is set to 1, the value this field MUST be ignored.
 * for series axes, must be 0
 * for cat: This field specifies the category at which the value axis crosses.
 * For example, if this field is 2, the value axis crosses this axis at the second category on this axis.
 * MUST be greater than or equal to 1 and less than or equal to 31999.
 * 6		catLabel	2	Frequency of labels. A signed integer that specifies the interval between axis labels on this axis.
 * MUST be greater than or equal to 1 and less than or equal to 31999.
 * MUST be ignored for a date axis.
 * 8		catMark		2	Frequency of tick marks.  A signed integer that specifies the interval at which major tick marks
 * and minor tick marks are displayed on the axis. Major tick marks and minor tick marks that would
 * have been visible are hidden unless they are located at a multiple of this field.
 * MUST be greater than or equal to 1, and less than or equal to 31999.
 * MUST be ignored for a date axis.
 * 10		grbit		2	Format flags
 * <p/>
 * The catCross field defines the point on the category axis where the value axis crosses.
 * A value of 01 indicates that the value axis crosses to the left, or in the center, of the first category
 * (depending on the value of bit 0 of the grbit field); a value of 02 indicates that the value axis crosses
 * to the left or center of the second category, and so on. Bit 2 of the grbit field overrides the value of catCross when set to 1.
 * <p/>
 * The catLabel field defines how often labels appear along the category or series axis. A value of 01 indicates
 * that a category label will appear with each category, a value of 02 means a label appears every other category, and so on.
 * <p/>
 * The catMark field defines how often tick marks appear along the category or series axis. A value of 01 indicates
 * that a tick mark will appear between each category or series; a value of 02 means a label appears between every
 * other category or series, etc.
 * <p/>
 * format flags:
 * 0	0xOl		fBetween		Value axis crossing a = axis crosses midcategory I = axis crosses between categories
 * 1	0x02		fMaxCross		Value axis crosses at the far right category (in a line, bar, column, scatter, or area chart; 2-D charts only)
 * 0 The value axis crosses this axis at the value specified by catCross.
 * 1 The value axis crosses this axis at the last category, the last series, or the maximum date.
 * 2	0x04		fReverse		Display categories in reverse order
 * 7-3	0xF8		(reserved)		Reserved; must be zero
 * 7-0	0xFF		(reserved)		Reserved; must be zero
 */
public class CatserRange extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 905038625844435651L;
	private short grbit, catCross, catLabel, catMark;
	private boolean fBetween, fMaxCross, fReverse;

	// 20070723 KSC: Need to create new records
	public static XLSRecord getPrototype()
	{
		CatserRange c = new CatserRange();
		c.setOpcode( CATSERRANGE );
		c.setData( c.PROTOTYPE_BYTES );
		c.init();
		return c;
	}    // 0, 0, 1, 0, 1, 0, 0, 0	- for axis 2 of surface chart

	private byte[] PROTOTYPE_BYTES = new byte[]{ 1, 0, 1, 0, 1, 0, 1, 0 };

	// 20070727 KSC: TODO: Get data def and parse correctly!
	public void setOpt( int op )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) op );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
		// 20070802 KSC: don't know what this means
		this.getData()[6] = 0;
	}

	// 20070802 KSC: parse data
	@Override
	public void init()
	{
		super.init();
		catCross = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		catLabel = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		catMark = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		grbit = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
		fBetween = (grbit & 0x1) == 0x1;
		fMaxCross = (grbit & 0x2) == 0x2;
		fReverse = (grbit & 0x4) == 0x4;
	}

	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( catCross );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
		b = ByteTools.shortToLEBytes( catLabel );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];
		b = ByteTools.shortToLEBytes( catMark );
		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
		b = ByteTools.shortToLEBytes( grbit );
		this.getData()[6] = b[0];
		this.getData()[7] = b[1];
	}

	// get/set methods
	public int getCatCross()
	{
		return catCross;
	}

	public int getCatLabel()
	{
		return catLabel;
	}

	public int getCatMark()
	{
		return catMark;
	}

	public boolean getCrossBetween()
	{
		return fBetween;
	}

	public boolean getCrossMax()
	{
		return fMaxCross;
	}

	public void setCatCross( int c )
	{
		catCross = (short) c;
		updateRecord();
	}

	public void setCatLabel( int c )
	{
		catLabel = (short) c;
		updateRecord();
	}

	public void setCatMark( int c )
	{
		catMark = (short) c;
		updateRecord();
	}

	public void setCrossBetween( boolean b )
	{
		fBetween = b;
		grbit = ByteTools.updateGrBit( grbit, fBetween, 0 );
		updateRecord();
	}

	public void setCrossMax( boolean b )
	{
		fMaxCross = b;
		grbit = ByteTools.updateGrBit( grbit, fMaxCross, 1 );
		updateRecord();
	}

	/**
	 * sets a specific OOXML axis option
	 * <br>can be one of:
	 * <br>auto
	 * <br>crosses			possible crossing points (autoZero, max, min)
	 * <br>crossesAt		where on axis the perpendicular axis crosses (double val)
	 * <br>lblAlign			text alignment for tick labels (ctr, l, r) (cat only)
	 * <br>lblOffset		distance of labels from the axis (0-1000)  (cat only)
	 * <br>tickLblSkip		how many tick labels to skip between label (int >= 1)
	 * <br>tickMarkSkip		how many tick marks to skip betwen ticks (int >= 1)
	 *
	 * @param op
	 * @param val
	 */
	// TODO: auto
	public boolean setOption( String op, String val )
	{
		if( op.equals( "crossesAt" ) )                // specifies where axis crosses (double value
		{
			catCross = Short.valueOf( val ).shortValue();
		}
		else if( op.equals( "orientation" ) )
		{    // axis orientation minMax or maxMin  -- fReverse
			fReverse = (val.equals( "maxMin" ));    // means in reverse order
			ByteTools.updateGrBit( grbit, fReverse, 2 );
		}
		else if( op.equals( "crosses" ) )
		{            // specifies how axis crosses it's perpendicular axis (val= max, min, autoZero)  -- fbetween + fMaxCross?/fAutoCross + fMaxCross
			if( val.equals( "max" ) )
			{        // TODO: this is probly wrong
				fMaxCross = true;
				ByteTools.updateGrBit( grbit, fMaxCross, 7 );
			}
			else if( val.equals( "mid" ) )
			{
				fBetween = false;
				ByteTools.updateGrBit( grbit, fBetween, 0 );
			}
			else if( val.equals( "autoZero" ) )
			{
				fBetween = true;    // is this correct??
				ByteTools.updateGrBit( grbit, fBetween, 0 );
			}
			else if( val.equals( "min" ) )
			{
				;
			}
			// TODO:  ???
		}
		else if( op.equals( "tickMarkSkip" ) )    //val= how many tick marks to skip before next one is drawn -- catMark -- catsterrange only?
		{
			catMark = Integer.valueOf( val ).shortValue();
		}
		else if( op.equals( "tickLblSkip" ) )
		{
			catLabel = Integer.valueOf( val ).shortValue();
		}
		else
		{
			return false;    // not handled
		}
		this.updateRecord();
		return true;
	}

	/**
	 * retrieve generic Category axis option
	 *
	 * @param op
	 * @return
	 */
	public String getOption( String op )
	{
		// TODO: auto, lblAlign, lblOffset
		if( op.equals( "crossesAt" ) )
		{
			return String.valueOf( catCross );
		}
		if( op.equals( "orientation" ) )
		{
			return (fReverse) ? "maxMin" : "minMax";
		}
		if( op.equals( "crosses" ) )
		{
			if( fMaxCross )
			{
				return "max";
			}
			if( fBetween )
			{
				return "autoZero";    // correct??
			}
			return "min";    // correct??
		}
		if( op.equals( "tickMarkSkip" ) )    //val= how many tick marks to skip before next one is drawn -- catMark -- catsterrange only?
		{
			return String.valueOf( catMark );
		}
		if( op.equals( "tickLblSkip" ) )
		{
			return String.valueOf( catLabel );
		}
		return null;
	}

	/**
	 * returns true if the axis should be displayed at top of chart area
	 * false if axis is displayed in the default bottom location
	 *
	 * @return
	 */
	public boolean isReversed()
	{
		return fReverse;
	}
}
