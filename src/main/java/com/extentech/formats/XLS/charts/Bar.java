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
 * Bar: Chart Group is a Bar or Column Chart Group (0x1017)
 * NOTE: Bar is also base type for Pyramid, Cone and Cylinder charts;
 * actual chart type is determined also by bar shape
 * see ChartFormat.getChartType for more information
 * <p/>
 * <p/>
 * 4	pcOverlap	2	Space between bars (default= 0)
 * values: 	-100 to -1		Size of the separation between data points
 * 0				No overlap.
 * 1 to 100		Size of the overlap between data points
 * 6	pcGap		2	Space between categories (%) (default=50%)
 * An unsigned integer that specifies the width of the gap between the categories and the left and right edges of the plot area
 * as a percentage of the data point width divided by 2. It also specifies the width of the gap between adjacent categories
 * as a percentage of the data point width. MUST be less than or equal to 500.
 * 8	grbit		2
 * <p/>
 * grbit:
 * <p/>
 * 0	0	0x1		fTranspose		1= horizontal bars (=bar chart)	0= vertical bars (= column chart)
 * 1	0x2		fStacked		Stack the displayed values
 * 2	0x4		f100			Each category is displayed as a percentage
 * 3	0x8		fHasShadow		1= this bar has a shadow
 */
public class Bar extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8917510368688674273L;
	private short grbit = 0;
	protected boolean fStacked = false, f100 = false, fHasShadow = false;
	protected short pcOverlap = 0;
	protected short pcGap = 50;

	@Override
	public void init()
	{
		super.init();
		pcOverlap = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		// should be 50 default, but seems to be 150 ??????
		pcGap = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		// pcGap: 0x0096 specifies that the width of the gap between adjacent categories is 150% of the data point width. It also specifies that the width of the gap between the categories (3) and the left and right edges of the plot area is 75% of the data point width. 
		grbit = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		if( (grbit & 0x1) == 0x1 )
		{
			chartType = BARCHART;
		}
		else
		{
			chartType = COLCHART;
		}
		fStacked = ((grbit & 0x2) == 0x2);
		f100 = ((grbit & 0x4) == 0x4);
		fHasShadow = ((grbit & 0x8) == 0x8);
	}

	// 20070716 KSC: get/set methods for format options
	@Override
	public boolean isStacked()
	{
		return fStacked;
	}

	@Override
	public boolean is100Percent()
	{
		return f100;
	}

	@Override
	public boolean hasShadow()
	{
		return fHasShadow;
	}

	public int getGap()
	{
		return pcGap;
	}

	public int getOverlap()
	{
		return pcOverlap;
	}

	/**
	 * sets this bar/col chart to have stacked series:  series are drawn next to each other along the value axis
	 *
	 * @param bIsClustered
	 */
	@Override
	public void setIsStacked( boolean bIsStacked )
	{
		fStacked = bIsStacked;
		grbit = ByteTools.updateGrBit( grbit, fStacked, 1 );
		if( bIsStacked )
		{
			pcOverlap = -100;
			pcGap = 150;
		}
		updateRecord();
	}

	@Override
	public void setIs100Percent( boolean bOn )
	{
		f100 = bOn;
		grbit = ByteTools.updateGrBit( grbit, f100, 2 );
		if( bOn )
		{
			pcOverlap = -100;
			pcGap = 150;
		}
		updateRecord();
	}

	public void setHasShadow( boolean bHasShadow )
	{
		fHasShadow = bHasShadow;
		grbit = ByteTools.updateGrBit( grbit, fHasShadow, 3 );
		updateRecord();
	}

	/**
	 * sets this bar/col chart to have clustered series:  series are drawn next to each other along the category axis
	 *
	 * @param bIsClustered
	 */
	public void setIsClustered()
	{
		setIsStacked( false );
		setIs100Percent( false );
	}

	/**
	 * sets the Space between bars (default= 0)
	 *
	 * @param overlap
	 */
	public void setpcOverlap( int overlap )
	{
		pcOverlap = (short) overlap;
		updateRecord();
	}

	/**
	 * sets the Space between categories (%) (default=50%)
	 *
	 * @param gap
	 */
	public void setpcGap( int gap )
	{
		pcGap = (short) gap;
		updateRecord();
	}

	public void setAsBarChart()
	{
		grbit = ByteTools.updateGrBit( grbit, true, 0 );    // set 0'th bit
		chartType = ChartConstants.BARCHART;
		this.updateRecord();
	}

	public void setAsColumnChart()
	{
		grbit = ByteTools.updateGrBit( grbit, false, 0 );    // clear 0'th bit
		chartType = ChartConstants.COLCHART;
		this.updateRecord();
	}

	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( pcOverlap );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
		b = ByteTools.shortToLEBytes( pcGap );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];
		b = ByteTools.shortToLEBytes( grbit );
		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
	}

	public static XLSRecord getPrototype()
	{
		Bar b = new Bar();
		b.setOpcode( BAR );
		b.setData( b.PROTOTYPE_BYTES );
		b.init();
		b.setAsColumnChart();
		return b;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0, -106, 0, 0, 0 };

	/**
	 * Set specific options
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "Stacked" ) )
		{
			setIsStacked( true );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "PercentageDisplay" ) )
		{
			setIs100Percent( true );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Shadow" ) )
		{
			setHasShadow( true );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Overlap" ) )
		{
			setpcOverlap( Integer.parseInt( val ) );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Gap" ) )
		{
			setpcGap( Integer.parseInt( val ) );
			bHandled = true;
		}
		return bHandled;
	}

	/**
	 * look up bar-specific chart options such as "Gap" or "Overlap" setting
	 */
	@Override
	public String getChartOption( String op )
	{
		if( op.equals( "Gap" ) )
		{ // Bar
			return String.valueOf( this.getGap() );
		}
		else if( op.equals( "Overlap" ) )
		{ // Bar
//    		return String.valueOf(Math.abs(this.getOverlap()));		// KSC: TESTING:  OOXML apparently needs +100 pcOverlap NOT -100 ... WHY and TRUE FOR ALL CASES?????
			return String.valueOf( this.getOverlap() );        // KSC: TESTING:  OOXML apparently needs +100 pcOverlap NOT -100 ... WHY and TRUE FOR ALL CASES?????
		}
		else
		{
			return super.getChartOption( op );
		}
	}
}
