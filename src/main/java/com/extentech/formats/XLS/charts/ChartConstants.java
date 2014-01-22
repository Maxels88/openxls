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

/**
 * Constants required for Chart functionality
 */

public interface ChartConstants
{
	// Chart Types (used in Chart Creation)
	public static final int COLCHART = 0;
	public static final int BARCHART = 1;
	public static final int LINECHART = 2;
	public static final int PIECHART = 3;
	public static final int AREACHART = 4; // 20070703 KSC:
	public static final int SCATTERCHART = 5; // ""
	public static final int RADARCHART = 6; // ""
	public static final int SURFACECHART = 7; // ""
	public static final int DOUGHNUTCHART = 8; // ""
	public static final int BUBBLECHART = 9; // ""
	public static final int OFPIECHART = 10;
	public static final int PYRAMIDCHART = 11;    // column-type pyramid
	public static final int CYLINDERCHART = 12;    // column-type cylinder
	public static final int CONECHART = 13;        // column-type cone
	public static final int PYRAMIDBARCHART = 14;// bar-type pyramid
	public static final int CYLINDERBARCHART = 15;    // bar-type cylinder
	public static final int CONEBARCHART = 16;    // bar-type cone
	public static final int RADARAREACHART = 17; // ""
	public static final int STOCKCHART = 18;

	// Bar Shapes
	public static final int SHAPEDEFAULT = 0;
	public static final int SHAPECOLUMN = SHAPEDEFAULT;        // default
	public static final int SHAPECYLINDER = 1;
	public static final int SHAPEPYRAMID = 256;
	public static final int SHAPECONE = 257;
	public static final int SHAPEPYRAMIDTOMAX = 516;
	public static final int SHAPECONETOMAX = 517;

	// Axis Types
	public static final int XAXIS = 0;
	public static final int YAXIS = 1;
	public static final int ZAXIS = 2;
	public static final int XVALAXIS = 3;    // an X axis type but VAL records
}