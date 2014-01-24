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

import com.extentech.toolkit.ByteTools;

/**
 * Chart3DBarShape
 * This record specifies the shape of the data points in a bar or column chart group.
 * This record is used only for a bar or column chart group and MUST be ignored for all other chart group
 * <p/>
 * 1	1	 riser	specifies the shape of the base of the data points in a bar or column chart group
 * 0 =base is a rectangle.  1 =base is an ellipse
 * 2	1	 taper 	specifies how the data points in a bar or column chart group taper from base to tip.
 * 0= no taper
 * 1= The data points of the bar or column chart group taper to a point at the maximum value of each data point
 * 2= he data points of the bar or column chart group taper towards a projected point at the position of the maximum value of all data points in the chart group, but are clipped at the value of each data point.
 */
public class Chart3DBarShape extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 3029030180040933082L;
	byte riser;
	byte taper;

	@Override
	public void init()
	{
		super.init();
		riser = getByteAt( 0 );
		taper = getByteAt( 1 );
	}

	public Chart3DBarShape()
	{
		setData( new byte[2] );
	}

	/**
	 * Bar Shapes
	 * public static final int SHAPECOLUMN= 0;		// default
	 * public static final int SHAPECYLINDER= 1;
	 * public static final int SHAPEPYRAMID= 256;
	 * public static final int SHAPECONE= 257;
	 * public static final int SHAPEPYRAMIDTOMAX= 516;
	 * public static final int SHAPECONETOMAX= 517;
	 */
	public short getShape()
	{
		return ByteTools.readShort( riser, taper );
	}

	/**
	 * set the shape of the bars
	 * <br>the shape is as follows:
	 * public static final int SHAPECOLUMN= 0;		// default
	 * public static final int SHAPECYLINDER= 1;
	 * public static final int SHAPEPYRAMID= 256;
	 * public static final int SHAPECONE= 257;
	 * public static final int SHAPEPYRAMIDTOMAX= 516;
	 * public static final int SHAPECONETOMAX= 517;
	 */
	public void setShape( short shape )
	{
		byte[] b = ByteTools.shortToLEBytes( shape );
		getData()[0] = b[0];
		getData()[1] = b[1];
		riser = getByteAt( 0 );
		taper = getByteAt( 1 );
	}

}
