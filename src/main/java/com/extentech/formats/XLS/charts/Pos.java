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
 * <b>Pos: Position Information(0x104f)</b>
 * <p/>
 * <p/>
 * for TextDisp, sets the label position as an offset from the default position
 * for PlotArea, used only for main axis + describes the plot-area bounding box; the tMainPlotArea in the SHTPROPS rec must be 1 or the POS rec is ignored
 * for Legend, describes legend pos + size
 * <p/>
 * 4	rndTopLt		2		for PlotArea, TextDisp= 2; legend= 5 (what=3??? Data Table!)
 * 6	rndTopRt		2		for PlotArea, TextDisp= 2; legend: 1= use x2 and y2 for legend size; 2= autosize legend (ignore x2 + y2; if so, the fAutoSize bit of FRAME rec should be 1)
 * 8	x1				4		for PlotArea, x coord of bounding box; for TextDisp, horiz. offset from default pos; for legend, x coord in 1/4000
 * 12  y1				4		for PlotArea, y coord of bounding box; for TextDisp, vert. offset from default pos; for legend, y coord " "
 * 16	x2				4		for PlotArea, w of bounding box; for TextDisp, ignored; for legend= width
 * 20	y2				4		for PlotArea, h of bounding box; ""; for legend= height
 * <p/>
 * Above is not correct;
 * Correct Information:
 * mdTopLt (2 bytes): A PositionMode structure that specifies the positioning mode for the upper-left corner of a legend,
 * an attached label, or the plot area. The valid combinations of mdTopLt and mdBotRt and the meaning of x1, y1, x2, y2
 * are specified in the Valid Combinations of mdTopLt and mdBotRt by Type table.
 * mdBotRt (2 bytes): A PositionMode structure that specifies the positioning mode for the lower-right corner of a legend,
 * an attached label, or the plot area. The valid combinations of mdTopLt and mdBotRt and the meaning of x1, y1, x2, y2
 * are specified in the following table.
 * <p/>
 * x1 (2 bytes): A signed integer that specifies a position. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
 * unused1 (2 bytes): Undefined and MUST be ignored.
 * y1 (2 bytes): A signed integer that specifies a position. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
 * unused2 (2 bytes): Undefined and MUST be ignored.
 * x2 (2 bytes): A signed integer that specifies a width. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
 * unused3 (2 bytes): Undefined and MUST be ignored.
 * y2 (2 bytes): A signed integer that specifies a height. The meaning is specified in the earlier table showing the valid combinations mdTopLt and mdBotRt by type.
 * unused4 (2 bytes): Undefined and MUST be ignored.
 * <p/>
 * Table:
 * Type			mdTopLtPosition Mode		mdBotRt Position Mode		Meaning
 * plot area (axis group)	MDPARENT			MDPARENT					The values of x1 and y1 specify the horizontal and vertical offsets of the primary axis group's
 * upper-left corner, relative to the upper-left corner of the chart area, in SPRC. The values of x2
 * and y2 specify the width and height of the primary axis group, in SPRC.
 * legend			MDCHART						MDABS						The values x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
 * relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 specify the
 * width and height of the legend, in points.
 * legend			MDCHART						MDPARENT					The values of x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
 * relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 MUST be ignored.
 * The size of the legend is determined by the application.
 * legend			MDKTH						MDPARENT					The values of x1, y1, x2 and y2 MUST be ignored. The legend is located inside a data table.
 * <p/>
 * attached label	MDPARENT					MDPARENT					The meaning of x1 and y1 is specified in the Type of Attached Label table. x2 and y2 MUST be ignored.
 * The size of the attached label is determined by the application.
 * <p/>
 * <p/>
 * The PositionMode structure specifies positioning mode for position information saved in a Pos record.
 * Name			Value					   Meaning
 * MDFX			0x0000						Relative position to the chart, in points.
 * MDABS			0x0001						Absolute width and height in points. It can only be applied to the mdBotRt field of Pos.
 * MDPARENT		0x0002						Owner of Pos determines how to interpret the position data.
 * MDKTH			0x0003						Offset to default position, in 1/1000th of the plot area size.
 * MDCHART			0x0005						Relative position to the chart, in SPRC (A SPRC is a unit of measurement that is 1/4000th of the height or width of the chart).
 * <p/>
 * Type of Attached Label	Meaning
 * Chart title				The value of x1 and y1 specify the horizontal and vertical offset of the title, relative to its default position, in SPRC.
 * Axis title				The value of x1 and y1 specify the offset of the title along the direction of a specific axis. The value of x1 specifies an offset along the category (3) axis, date axis, or horizontal value axis. The value of y1 specifies an offset along the value axis. Both offsets are relative to the title's default position, in 1/1000th of the axis length.
 * Data label				If the chart is not a pie chart group or a radar chart group, x1 and y1 specify the offset of the label along the direction of the specific axis.
 * The x1 value is an offset along the category (3) axis, date axis, or horizontal value axis.
 * The y1 value is an offset along the value axis, opposite to the direction of the value axis.
 * Both offsets are relative to the label's default position, in 1/1000th of the axis length.
 * For a pie chart group, the value of x1 specifies the clockwise angle, in degrees, and the value of y1 specifies the radius offset of the label relative to its default position, in 1/1000th of the pie radius length. A label moved toward the pie center has a negative radius offset.
 * For a radar chart group, the values of x1 and y1 specify the horizontal and vertical offset of the label relative to its default position, in 1/1000th of the axis length.
 */
public class Pos extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 7920967716354683818L;
	short rndTopLt, rndTopRt;
	int x1, y1, x2, y2;
	public static final int TYPE_TEXTDISP = 0;
	public static final int TYPE_LEGEND = 1;
	public static final int TYPE_PLOTAREA = 2;
	public static final int TYPE_DATATABLE = 3;

	@Override
	public void init()
	{
		super.init();
		// 0= in points, relative to the position of the chart
		// 1= absolute, in points
		// 2= parent of this rec determines
		// 3= offset to default
		// 5= relative, in 1/4000 of the w or h of the chart
		rndTopLt = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		rndTopRt = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		x1 = ByteTools.readInt( this.getBytesAt( 4, 4 ) );
		y1 = ByteTools.readInt( this.getBytesAt( 8, 4 ) );
		x2 = ByteTools.readInt( this.getBytesAt( 12, 4 ) );
		y2 = ByteTools.readInt( this.getBytesAt( 16, 4 ) );
	}

	// TODO: Get def and parse options accordingly!
	public static XLSRecord getPrototype( int type )
	{
		Pos p = new Pos();
		p.setOpcode( POS );
		p.setData( p.PROTOTYPE_BYTES );
		p.init();
		p.setType( type );
		return p;
	}

	/**
	 * set the correct bytes for the desired type
	 *
	 * @param type
	 */
	public void setType( int type )
	{
		switch( type )
		{
			case TYPE_PLOTAREA:
			case TYPE_TEXTDISP:
				this.getData()[0] = 2;
				this.getData()[1] = 0;
				break;
			case TYPE_LEGEND:
				this.getData()[0] = 5;
				this.getData()[1] = 2;
				break;
			case TYPE_DATATABLE:
				this.getData()[0] = 3;
				this.getData()[1] = 0;
		}
		this.getData()[2] = 2;
		this.getData()[3] = 0;
	}

	public void setX( int x )
	{
		x1 = x;
		byte[] b = ByteTools.cLongToLEBytes( x );
		System.arraycopy( b, 0, this.getData(), 4, 4 );
	}

	public void setY( int y )
	{
		y1 = y;
		byte[] b = ByteTools.cLongToLEBytes( y );
		System.arraycopy( b, 0, this.getData(), 8, 4 );
	}

	/**
	 * for legends only, set width of bounds
	 *
	 * @param w
	 */
	public void setLegendW( int w )
	{
		if( (rndTopLt == 5) && (rndTopRt == 1) )
		{
			x2 = w;
			byte[] b = ByteTools.cLongToLEBytes( x2 );
			System.arraycopy( b, 0, this.getData(), 12, 4 );
		} // else throw exception?
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 2, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

	/**
	 * set autosize bit for legend Pos's
	 * valid only for legend-type Pos's (rndTopLt==5)
	 * also must set any associated Frame Autosize bits
	 * [BugTracker 2844]
	 */
	public void setAutosizeLegend()
	{
		if( this.rndTopLt == 5 )
		{// it's a legend Pos
			this.rndTopRt = 2;// set to autoposition
			this.getData()[2] = 2;
		}
	}

	/**
	 * return the legend coordinates (x, y in 1/4000 of the chart height or width, w, h in points) or null, depending upon legend options
	 * <br>NOTE: if the w or h are 0, use default.
	 *
	 * @return int[] x, y, w, h
	 */
	public int[] getLegendCoords()
	{
		if( (rndTopLt == 5) && (rndTopRt == 1) )
		{
			/*The values x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
			relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 specify the
		width and height of the legend, in points.*/
			return new int[]{ x1, y1, x2, y2 };
		}
		if( (rndTopLt == 5) && (rndTopRt == 2) )
		{
    		/*The values of x1 and y1 specify the horizontal and vertical offsets of the legend's upper-left corner,
		relative to the upper-left corner of the chart area, in SPRC. The values of x2 and y2 MUST be ignored.
		The size of the legend is determined by the application.*/
			return new int[]{ x1, y1, 0, 0 };
		}
		if( (rndTopLt == 3) && (rndTopRt == 2) )
		{
    		/*The values of x1, y1, x2 and y2 MUST be ignored. The legend is located inside a data table.*/
			return null;
		}
		return null;
	}

	/**
	 * if plot area,
	 * return the Plot Area (x, y, w, h in 1/4000 of the chart height or width) or null, depending upon options
	 * <br>NOTES: Pos record of the Axis Parent:
	 * <br>The Pos record specifies the position and size of the outer plot area. The outer plot area is the bounding rectangle that includes the axis labels,
	 * <br>the axis titles, and data table of the chart.
	 * if attached label,
	 * Chart title			The value of x1 and y1 specify the horizontal and vertical offset of the title, relative to its default position, in SPRC.
	 * Axis title			The value of x1 and y1 specify the offset of the title along the direction of a specific axis.
	 * The value of x1 specifies an offset along the category axis, date axis, or horizontal value axis.
	 * The value of y1 specifies an offset along the value axis. Both offsets are relative to the title's default position, in 1/1000th of the axis length.
	 * Data label			If the chart is not a pie chart group or a radar chart group, x1 and y1 specify the offset of the label along the direction of the specific axis.
	 * The x1 value is an offset along the category axis, date axis, or horizontal value axis. The y1 value is an offset along the value axis,
	 * opposite to the direction of the value axis. Both offsets are relative to the label's default position, in 1/1000th of the axis length.
	 * For a pie chart group, the value of x1 specifies the clockwise angle, in degrees, and the value of y1 specifies the radius offset of the label
	 * relative to its default position, in 1/1000th of the pie radius length. A label moved toward the pie center has a negative radius offset.
	 * For a radar chart group, the values of x1 and y1 specify the horizontal and vertical offset of the label relative to its default position,
	 * in 1/1000th of the axis length.
	 *
	 * @return int[] x, y, w, h
	 */
	public float[] getCoords()
	{
		if( (rndTopLt == 2) && (rndTopRt == 2) )
		{
    		/* 	The values of x1 and y1 specify the horizontal and vertical offsets of the primary axis group's 
				upper-left corner, relative to the upper-left corner of the chart area, in SPRC. The values of x2 
				and y2 specify the width and height of the primary axis group, in SPRC.*/
			return new float[]{ x1, y1, x2, y2 };
		}
		return null;
	}

	/**
	 * convert a coordinate value in SPRC units to points
	 *
	 * @param val
	 * @param w
	 * @param h
	 * @return
	 */
	public static float convertFromSPRC( float val, float w, float h )
	{
		// try this:
		if( w != 0 )
		{
			return (float) (val / 4000.0) * w;
		}
		return (float) (val / 4000.0) * h;
	}

	/**
	 * convert a coordinate value in points to SPRC units
	 * <br>Experimental at this point
	 *
	 * @param val
	 * @param w
	 * @param h
	 * @return
	 */
	public static float convertToSPRC( float val, float w, float h )
	{
		// try this:
		if( w != 0 )
		{
			return (float) (val * 4000.0) / w;
		}
		return (float) (val * 4000.0) / h;
	}

	/**
	 * convert a coordinate value in SPRC units to points
	 *
	 * @param val
	 * @param w
	 * @param h
	 * @return
	 */
	public static float convertFromLabelUnits( float val, float w, float h )
	{
		// try this:
		if( w != 0 )
		{
			return (float) (val / 1000.0) * w;
		}
		return (float) (val / 1000.0) * h;
	}
}
