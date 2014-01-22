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
 * <b>Scatter: Chart Group is a Scatter Chart Group(0x101b)</b>
 * <p/>
 * 4	pcBubbleSizeRatio	2		Percent of largest bubble compared to chart in general default= 100
 * 6	wBubbleSize			2		Bubble size: 1= bubble size is area, 2= bubble size is width	default= 1
 * 8	grbit				2		flags
 * <p/>
 * grbit
 * 0	0x1		fBubbles		1= this is a bubble series
 * 1	0x2		fShowNegBubbles	1= show negative bubbles
 * 2	0x4		fHasShadow		1= bubble series has a shadow
 */
public class Scatter extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -289334135036242100L;
	private short grbit = 0;
	private short pcBubbleSizeRatio = 100;
	private short wBubbleSize = 1;
	private boolean fBubbles = false;
	private boolean fShowNegBubbles = false;
	private boolean fHasShadow = false;

	public void init()
	{
		super.init();
		// 20070703 KSC:
		pcBubbleSizeRatio = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		wBubbleSize = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		grbit = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		fBubbles = (grbit & 0x1) == 0x1;
		fShowNegBubbles = (grbit & 0x2) == 0x2;
		fHasShadow = (grbit & 0x4) == 0x4;
		if( fBubbles )
		{
			chartType = ChartConstants.BUBBLECHART;
		}
		else
		{
			chartType = ChartConstants.SCATTERCHART;
		}
	}

	// 20070703 KSC
	private void updateRecord()
	{
		grbit = ByteTools.updateGrBit( grbit, fBubbles, 0 );
		grbit = ByteTools.updateGrBit( grbit, fShowNegBubbles, 1 );
		grbit = ByteTools.updateGrBit( grbit, fHasShadow, 2 );
		byte[] b = ByteTools.shortToLEBytes( pcBubbleSizeRatio );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
		b = ByteTools.shortToLEBytes( wBubbleSize );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];
		b = ByteTools.shortToLEBytes( grbit );
		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
	}

	public void setAsScatterChart()
	{
		fBubbles = false;
		chartType = ChartConstants.SCATTERCHART;
		this.updateRecord();
	}

	public void setAsBubbleChart()
	{
		fBubbles = true;
		chartType = ChartConstants.BUBBLECHART;
		this.updateRecord();
	}

	public static XLSRecord getPrototype()
	{
		Scatter b = new Scatter();
		b.setOpcode( SCATTER );
		b.setData( b.PROTOTYPE_BYTES );
		b.init();
		b.setAsScatterChart();
		return b;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 100, 0, 1, 0, 0, 0 };

	/**
	 * @return String XML representation of this chart-type's options
	 */
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		if( pcBubbleSizeRatio != 100 )
		{
			sb.append( " BubbleSizeRatio=\"" + pcBubbleSizeRatio + "\"" );
		}
		if( wBubbleSize != 1 )
		{
			sb.append( " BubbleSize=\"" + wBubbleSize + "\"" );
		}
		if( fShowNegBubbles )
		{
			sb.append( " ShowNeg=\"true\"" );
		}
		if( fHasShadow )
		{
			sb.append( " Shadow=\"true\"" );
		}
		return sb.toString();
	}

	/**
	 * Handle setting options from XML in a generic manner
	 */
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "BubbleSizeRatio" ) )
		{
			pcBubbleSizeRatio = Short.parseShort( val );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "BubbleSize" ) )
		{
			wBubbleSize = Short.parseShort( val );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ShowNeg" ) )
		{
			fShowNegBubbles = val.equals( "true" );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Shadow" ) )
		{
			fHasShadow = val.equals( "true" );
			bHandled = true;
		}
		if( bHandled )
		{
			updateRecord();
		}
		return bHandled;
	}

}	
