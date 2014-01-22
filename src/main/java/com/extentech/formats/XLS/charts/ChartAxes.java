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

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.cellformat.CellFormatFactory;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;
import org.json.JSONObject;
import org.xmlpull.v1.XmlPullParser;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Stack;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * An axis group is a set of axes that specify a coordinate system, a set of chart groups that are plotted using these and the plot area that defines
 * where the axes are rendered on the chart.
 * <p/>
 * In BIFF8, the AxisParent record governs the Axis Group.  A typical arrangement of records is:
 * AxisParent
 * Pos
 * Axis (X)
 * Axis (Y)
 * PlotArea
 * Frame
 * (Chart Group is one or more of:)
 * ChartFormat (defines the chart type) + Legend
 */
public class ChartAxes implements ChartConstants, Serializable
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7862828186455339066L;
	private AxisParent ap;
	ArrayList axes = new ArrayList();
	HashMap<String, Object> axisMetrics = new HashMap();    // holds important axis metrics such as xAxisReversed, xPattern, for outside use
/**
 * The AxisParent record specifies if the axis group is the primary axis group or the secondary axis group on a chart. 
 * Often the axes of the primary axis group are displayed to the left and bottom sides of the plot area, while axes of the secondary axis group are displayed on the right and top sides of the plot area.
 The Pos record specifies the position and size of the outer plot area. The outer plot area is the bounding rectangle that includes the axis labels, the axis titles,
 and data table of the chart. This record MUST be ignored on a secondary axis group. */
	/** The PlotArea record and the sequence of records that conforms to the FRAME rule in the sequence of records that conform to the AXES rule specify the properties of the inner plot area. The inner plot area is the rectangle bounded by the chart axes. 
	 * The PlotArea record MUST not exist on a secondary axis group.
	 */
	/**
	 * @param ap
	 */
	public ChartAxes( AxisParent ap )
	{
		this.ap = ap;
	}

	/**
	 * store each axis
	 *
	 * @param a
	 */
	public void add( Axis a )
	{
		axes.add( a );
		a.setAP( this.ap );    // ensure axis is linked to it's parent AxisParent
	}

	public void setTd( int axisType, TextDisp td )
	{
		Axis a = getAxis( axisType, false );
		if( a != null ) // shouldn't
		{
			a.setTd( td );
		}
	}

	/**
	 * returns true if axisType is found on chart
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return
	 */
	public boolean hasAxis( int axisType )
	{
		for( int i = 0; i < axes.size(); i++ )
		{
			if( ((Axis) axes.get( i )).getAxis() == axisType )
			{
				return true;
			}
		}
		return false;
	}

	public void createAxis( int axisType )
	{
		Axis a = getAxis( axisType, true );
		return;
	}

	/**
	 * Return the desired axis 
	 * @return Axis
	 *
	public Axis getAxis(int axisType) {	
	return ap.getAxis(axisType);
	}*/

	/**
	 * returns the desired axis if it exists
	 * if bCreateIfNecessary, will create if it doesn't exist
	 * otherwise returns null
	 *
	 * @param axisType           one of: XAXIS, YAXIS, ZAXIS
	 * @param bCreateIfNecessary
	 * @return
	 */
	private Axis getAxis( int axisType, boolean bCreateIfNecessary )
	{
		for( int i = 0; i < axes.size(); i++ )
		{
			if( ((Axis) axes.get( i )).getAxis() == axisType )
			{
				return ((Axis) axes.get( i ));
			}
		}
		if( bCreateIfNecessary )
		{
			axes.add( ap.getAxis( axisType, bCreateIfNecessary ) );
			return (Axis) axes.get( axes.size() - 1 );
		}
		return null;
	}

	/**
	 * returns true if Axis is reversed
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return
	 */
	protected boolean isReversed( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.isReversed();
		}
		return false;
	}

	/**
	 * returns the number format for the desired axis
	 *
	 * @param axisType
	 * @return number format string
	 */
	protected String getNumberFormat( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getNumberFormat();
		}
		return null;
	}

	/**
	 * remove ALL axes from this chart
	 */
	public void removeAxes()
	{
		ap.removeAxes();
		while( axes.size() > 0 )
		{
			axes.remove( 0 );
		}
	}

	/**
	 * remove the desired axis + associated records
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 */
	public void removeAxis( int axisType )
	{
		ap.removeAxis( axisType );
		for( int i = 0; i < axes.size(); i++ )
		{
			if( ((Axis) axes.get( i )).getAxis() == axisType )
			{
				axes.remove( i );
				break;
			}
		}
	}

	/**
	 * return the coordinates of the outer plot area (plot + axes + labels) in pixels
	 *
	 * @return
	 */
	public float[] getPlotAreaCoords( float w, float h )
	{
		Pos p = (Pos) Chart.findRec( ap.chartArr, Pos.class );
		if( p != null )
		{
			float[] plotcoords = p.getCoords();
			if( plotcoords != null )
			{
				plotcoords[0] = Pos.convertFromSPRC( plotcoords[0], w, 0 );
				plotcoords[1] = Pos.convertFromSPRC( plotcoords[1], 0, h );
				plotcoords[2] = Pos.convertFromSPRC( plotcoords[2], w, 0 );    // SPRC units see Pos
				plotcoords[3] = Pos.convertFromSPRC( plotcoords[3], 0, h );    // SPRC units see Pos
				return plotcoords;
			}
		}
		return null;
	}

	/**
	 * returns the axis area background color hex string, or null if not set
	 *
	 * @return color hex string
	 */
	public String getPlotAreaBgColor()
	{
		return ap.getPlotAreaBgColor();
	}

	/**
	 * sets the plot area background color
	 *
	 * @param bg color int
	 */
	public void setPlotAreaBgColor( int bg )
	{
		ap.setPlotAreaBgColor( bg );
	}

	/**
	 * adds a border around the plot area with the desired line width and line color
	 *
	 * @param lw
	 * @param lclr
	 */
	public void setPlotAreaBorder( int lw, int lclr )
	{
		ap.setPlotAreaBorder( lw, lclr );
	}

	/**
	 * obtain the desired axis' label and other options, if present, in XML form
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return String XML representation of desired axis' label and options
	 * @see ObjectLink
	 */
	public String getAxisOptionsXML( int axisType )
	{
		return ap.getAxisOptionsXML( axisType );
	}

	/**
	 * Return the Axis Title, or "" if none
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 */
	public String getTitle( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getTitle();
		}
		return "";
	}

	/**
	 * set the axis title string
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param Title
	 */
	public void setTitle( int axisType, String Title )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			boolean defaultIsVert = ((axisType == YAXIS) && ("".equals( a.getTitle() )));
			a.setTitle( Title );
			if( defaultIsVert )
			{
				a.getTd().setRotation( 90 );
			}
		}
	}

	/**
	 * return the rotation of the axis labels, if any
	 * 0 to 90		Text rotated 0 to 90 degrees counter-clockwise
	 * 91 to 180	Text rotated 1 to 90 degrees clockwise (angle is trot – 90)
	 * 255			Text top-to-bottom with letters upright
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 */
	public int getLabelRotation( int axisType )
	{
		int rot = 0;
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			Tick t = (Tick) Chart.findRec( a.chartArr, Tick.class );
			if( t != null )
			{    // shoudn't
				rot = t.getRotation();
				switch( rot )
				{
					case 0:
						break;
					case 1:        // not correct ...
						rot = 180;
						break;
					case 2:
						rot = -90;
						break;
					case 3:
						rot = 90;
						break;
				}
			}
		}
		return rot;
	}

	/**
	 * return the rotation of the axis Title, if any
	 * 0 to 90		Text rotated 0 to 90 degrees counter-clockwise
	 * 91 to 180	Text rotated 1 to 90 degrees clockwise (angle is trot – 90)
	 * 255			Text top-to-bottom with letters upright
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 */
	public int getTitleRotation( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			TextDisp td = a.getTd();
			return td.getRotation();
		}
		return 0;
	}

	/**
	 * return the coordinates, in pixels, of the title text area, if possible
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return float[]
	 */
	public float[] getCoords( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			TextDisp td = a.getTd();
			return td.getCoords();
		}
		return new float[]{ 0, 0, 0, 0 };
	}

	/**
	 * Returns the YAXIS scale elements (min value, max value, minor (tick) scale, major (tick) scale)
	 * the scale elements are calculated from the minimum and maximum values on the chart
	 *
	 * @param ymin the minimum value of all the series values
	 * @param ymax the maximum value of all the series values
	 * @return double[] min, max, minor, major
	 */
	public double[] getMinMax( double ymin, double ymax )
	{
		return getMinMax( ymin, ymax, YAXIS );
	}

	/**
	 * Returns the scale values of the the desired Value axis
	 * <br>Scale elements: (min value, max value, minor (tick) scale, major (tick) scale)
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 *
	 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @return int Miminum Scale value for the desired axis
	 * @see getAxisMinScale()
	 */
	public double[] getMinMax( double ymin, double ymax, int axisType )
	{
		Axis a = getAxis( axisType, false );
		double[] ret = new double[4];
		if( a != null )
		{
			ValueRange v = (ValueRange) Chart.findRec( a.getChartRecords(), ValueRange.class );
			if( v.isAutomaticScale() ) // major crazy Excel "automatic max/min/tickmarks calc" ... a monster
			{
				v.setMaxMin( ymax, ymin );
			}
			ret[0] = v.getMin();
			ret[1] = v.getMax();
			ret[2] = v.getMinorTick();
			ret[3] = v.getMajorTick();
		}
		return ret;
	}

	/**
	 * Returns the major tick unit of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return int major tick unit
	 * @see getAxisMajorUnit()
	 */
	public int getAxisMajorUnit( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getMajorUnit();
		}
		return 10;    // TODO throw exception instead ???
	}

	/**
	 * Returns the minor tick unit of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return int - minor tick unit of the desired axis
	 * @see getAxisMinorUnit()
	 */
	public int getAxisMinorUnit( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getMinorUnit();
		}
		return 1;    // TODO throw exception instead
	}

	/**
	 * Sets the maximum scale value of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 * <p>Note: The default scale setting for charts is known as Automatic Scaling
	 * <br>When data changes, the chart automatically adjusts the scale as necessary
	 * <br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param int      MaxValue - desired maximum value of the desired axis
	 * @see setAxisMax(int MaxValue)
	 */
	public void setAxisMax( int axisType, int MaxValue )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			a.setMaxScale( MaxValue );
		}
		// TODO: throw exception if no axis?
	}

	/**
	 * Sets the minimum scale value of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 * <p>Note: The default setting for charts is known as Automatic Scaling
	 * <br>When data values change, the chart automatically adjusts the scale as necessary
	 * <br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param int      MinValue - the desired Minimum scale value
	 * @see setAxisMin(int MinValue)
	 */
	public void setAxisMin( int axisType, int MinValue )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			a.setMinScale( MinValue );
		}
		// TODO: throw exception if no axis?
	}

	/**
	 * Sets the automatic scale option on or off for the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 * <p>Automatic Scaling automatically sets the scale maximum, minimum and tick units
	 * upon data changes, and is the default setting for charts
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param boolean  b - true if set Automatic scaling on, false otherwise
	 * @see setAxisAutomaticScale(boolean b)
	 */
	public void setAxisAutomaticScale( int axisType, boolean b )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			a.setAutomaticScale( b );
		}
	}

	/**
	 * Returns true if the desired Value axis is set to automatic scale
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 * <p>Note: The default setting for charts is known as Automatic Scaling
	 * <br>When data changes, the chart automatically adjusts the scale as necessary
	 * <br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return boolean true if Automatic Scaling is turned on
	 * @see getAxisAutomaticScale()
	 */
	public boolean getAxisAutomaticScale( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.isAutomaticScale();    // group min and max together ...
		}
		return false;    // TODO: throw exception?
	}

	/**
	 * Returns true if the Y Axis (Value axis) is set to automatic scale
	 * <p>The default setting for charts is known as Automatic Scaling
	 * <br>When data changes, the chart automatically adjusts the scale (minimum, maximum values
	 * plus major and minor tick units) as necessary
	 *
	 * @return boolean true if Automatic Scaling is turned on
	 * @see getAxisAutomaticScale(int axisType)
	 */
	public boolean getAxisAutomaticScale()
	{
		Axis a = getAxis( YAXIS, false );
		if( a != null )
		{
			return a.isAutomaticScale();    // group min and max together ...
		}
		return false;    // TODO: throw exception?
	}

	/**
	 * Sets the automatic scale option on or off for the Y Axis (Value axis)
	 * <p>Automatic Scaling will automatically set the scale maximum, minimum and tick units
	 * upon data changes, and is the default chart setting
	 *
	 * @param b
	 * @see setAxisAutomaticScale(int axisType boolean b)
	 */
	public void setAxisAutomaticScale( boolean b )
	{
		Axis a = getAxis( YAXIS, false );
		if( a != null )
		{
			a.setAutomaticScale( b );
		}
	}

	/**
	 * Sets the maximum value of the Y Axis (Value Axis) Scale
	 * <p>Note: The default scale setting for charts is known as Automatic Scaling
	 * <br>When data changes, the chart automatically adjusts the scale as necessary
	 * <br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
	 *
	 * @param int MaxValue - the desired maximum scale value
	 * @see ChartHandle.setAxisMax(int axisType, int MaxValue)
	 */
	public void setAxisMax( int MaxValue )
	{
		Axis a = getAxis( YAXIS, false );
		if( a != null )
		{
			a.setMaxScale( MaxValue );
		}
		// TODO: throw exception if no axis?
	}

	/**
	 * Sets the minimum value of the Y Axis (Value Axis) Scale
	 * <p>Note: The default setting for charts is known as Automatic Scaling
	 * <br>When data changes, the chart automatically adjusts the scale as necessary
	 * <br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
	 *
	 * @param int MinValue - the desired minimum scale value
	 * @see ChartHandle.setAxisMin(int axisType, int MinValue)
	 */
	public void setAxisMin( int MinValue )
	{
		Axis a = getAxis( YAXIS, false );
		if( a != null )
		{
			a.setMinScale( MinValue );
		}
		// TODO: throw exception if no axis?
	}

	/**
	 * returns the SVG necesssary to define the desired axis
	 *
	 * @param axisType     one of: XAXIS, YAXIS, ZAXIS
	 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @return
	 */
	public String getSVG( int axisType, java.util.Map<String, Double> chartMetrics, Object[] categories )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getSVG( this, chartMetrics, categories );
		}
		return null;
	}

	/**
	 * returns the OOXML necessary to define the desired axis
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param type
	 * @param id
	 * @param crossId
	 * @return
	 */
	public String getOOXML( int axisType, int type, String id, String crossId )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getOOXML( type, id, crossId );
		}
		return "";
	}

	/**
	 * set the font index for this Axis (for title)
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param fondId
	 */
	public void setTitleFont( int axisType, int fondId )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			a.setFont( fondId );
		}
	}

	/**
	 * returns the font for the title for the desired axis
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return
	 */
	public Font getTitleFont( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getFont();
		}
		return null;
	}

	/**
	 * returns the label font for the desired axis.
	 * If not explicitly set, returns the default font
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return
	 */
	public Font getLabelFont( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getLabelFont();
		}
		return null;
	}

	public JSONObject getJSON( int axisType,
	                           com.extentech.ExtenXLS.WorkBookHandle wbh,
	                           int chartType,
	                           double yMax,
	                           double yMin,
	                           int nSeries )
	{
		return ap.getAxis( axisType, false ).getJSON( wbh, chartType, yMax, yMin, nSeries );
	}

	/**
	 * sets an option for this axis
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param op       option name; one of: CatCross, LabelCross, Marks, CrossBetween, CrossMax,
	 *                 MajorGridLines, AddArea, AreaFg, AreaBg
	 *                 or Linked Text Display options:
	 *                 Label, ShowKey, ShowValue, ShowLabelPct, ShowPct,
	 *                 ShowCatLabel, ShowBubbleSizes, TextRotation, Font
	 * @param val      option value
	 */
	public void setChartOption( int axisType, String op, String val )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			a.setChartOption( op, val );
		}
	}

	/**
	 * returns the Axis Label Placement or position as an int
	 * <p>One of:
	 * <br>Axis.INVISIBLE - axis is hidden
	 * <br>Axis.LOW - low end of plot area
	 * <br>Axis.HIGH - high end of plot area
	 * <br>Axis.NEXTTO- next to axis (default)
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @return int - one of the Axis Label placement constants above
	 */
	public int getAxisPlacement( int axisType )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			return a.getAxisPlacement();
		}
		return -1;
	}

	/**
	 * sets the axis labels position or placement to the desired value (these match Excel placement options)
	 * <p>Possible options:
	 * <br>Axis.INVISIBLE - hides the axis
	 * <br>Axis.LOW - low end of plot area
	 * <br>Axis.HIGH - high end of plot area
	 * <br>Axis.NEXTTO- next to axis (default)
	 *
	 * @param axisType  one of: XAXIS, YAXIS, ZAXIS
	 * @param Placement - int one of the Axis placement constants listed above
	 */
	public void setAxisPlacement( int axisType, int placement )
	{
		Axis a = getAxis( axisType, false );
		if( a != null )
		{
			a.setAxisPlacement( placement );
		}
	}

	/**
	 * parse OOXML axis element
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param xpp      XmlPullParser positioned at correct elemnt
	 * @param axisTag  catAx, valAx, serAx, dateAx
	 * @param lastTag  Stack of element names
	 */
	public void parseOOXML( int axisType, XmlPullParser xpp, String tnm, Stack<String> lastTag, WorkBookHandle bk )
	{
		Axis a = getAxis( axisType, true );
		if( a != null )
		{
			a.removeTitle();
			a.setChartOption( "MajorGridLines", "false" );    // initally, remove any - will be set in parseOOXML if necessary
			a.parseOOXML( xpp, tnm, lastTag, bk );
		}
	}

	public void close()
	{
		axes.clear();
		axes = null;
		ap = null;
		axisMetrics.clear();
	}

	/**
	 * returns already-set axisMetrics
	 * <br>Used only after calling chart.getMetrics
	 *
	 * @return axisMetrics -- map of useful chart display metrics
	 * <br>Contains:
	 * <br>
	 * XAXISLABELOFFSET 	double
	 * XAXISTITLEOFFSET	double
	 * YAXISLABELOFFSET	double
	 * YAXISTITLEOFFSET	double
	 * xAxisRotate		integer x axis lable rotation angle
	 * yPattern		string numeric pattern for y axis
	 * xPattern		string numeric pattern for x axis
	 * double major; // tick units
	 * double minor;
	 */
	public HashMap getMetrics()
	{
		return axisMetrics;
	}

	/**
	 * returns a specific axis Metric
	 * <br>Use only after calling chart.getMetrics
	 *
	 * @param metric String metric option one of
	 *               <br>
	 *               XAXISLABELOFFSET 	double
	 *               XAXISTITLEOFFSET	double
	 *               YAXISLABELOFFSET	double
	 *               YAXISTITLEOFFSET	double
	 *               xAxisRotate		integer x axis lable rotation angle
	 *               yPattern		string numeric pattern for y axis
	 *               xPattern		string numeric pattern for x axis
	 *               double major; // tick units
	 *               double minor;
	 * @return
	 */
	public Object getMetric( String metric )
	{
		return axisMetrics.get( metric );
	}

	/**
	 * generate Axis Metrics -- minor, major scale, title and label offsets, etc.
	 * for use in SVG generation and other operations
	 *
	 * @param charttype    chart type constant
	 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @param plotcoords   == if set, plot-area coordinates (may not be set if automatic)
	 * @return axisMetrics Map:  contains the important axis metrics for use in chart display
	 */
	public HashMap<String, Object> getMetrics( int charttype,
	                                           HashMap<String, Double> chartMetrics,
	                                           float[] plotcoords,
	                                           Object[] categories )
	{
		double[] minmax = this.getMinMax( chartMetrics.get( "min" ),
		                                  chartMetrics.get( "max" ) );    // sets min/max on Value axis, based upon axis settings and actual minimun and maximum values
		chartMetrics.put( "min", minmax[0] );    // set new values, if any
		chartMetrics.put( "max", minmax[1] );    // ""
		axisMetrics.put( "minor", minmax[2] );
		axisMetrics.put( "major", minmax[3] );
		axisMetrics.put( "xAxisReversed", this.isReversed( XAXIS ) ); // default is Value axis crosses at bottom. reverse= crosses at top
		axisMetrics.put( "xPattern", this.getNumberFormat( XAXIS ) );
		axisMetrics.put( "yAxisReversed", this.isReversed( YAXIS ) ); // if value (Y), default is on LHS; reverse= RHS
		axisMetrics.put( "yPattern", this.getNumberFormat( YAXIS ) );
		axisMetrics.put( "XAXISLABELOFFSET", 0.0 );
		axisMetrics.put( "XAXISTITLEOFFSET", 0.0 );
		axisMetrics.put( "YAXISLABELOFFSET", 0.0 );
		axisMetrics.put( "YAXISTITLEOFFSET", 0.0 );

		// X axis title Offset
		if( !this.getTitle( XAXIS ).equals( "" ) )
		{
			com.extentech.formats.XLS.Font ef = this.getTitleFont( XAXIS );
			java.awt.Font f = new java.awt.Font( ef.getFontName(), ef.getFontWeight(), (int) ef.getFontHeightInPoints() );
			AtomicInteger h = new AtomicInteger( 0 );
			double w = getRotatedWidth( f, h, this.getTitle( XAXIS ), this.getTitleRotation( XAXIS ) );
			axisMetrics.put( "XAXISTITLEOFFSET", (double) (h.intValue() + 10) );	/* add a little padding */
		}

		// Y Axis Title Offsets 
		if( this.hasAxis( YAXIS ) )
		{
			String title = this.getTitle( YAXIS );
			if( !title.equals( "" ) )
			{
				com.extentech.formats.XLS.Font ef = this.getTitleFont( YAXIS );
				java.awt.Font f = new java.awt.Font( ef.getFontName(), ef.getFontWeight(), (int) ef.getFontHeightInPoints() );
				AtomicInteger h = new AtomicInteger( 0 );
				double w = getRotatedWidthVert( f, h, this.getTitle( YAXIS ), this.getTitleRotation( YAXIS ) );
				axisMetrics.put( "YAXISTITLEOFFSET", w ); // add padding
/*    			    			
    			int rot= this.getTitleRotation(YAXIS); //0-180 or 255 (=vertical with letters upright)
				if (rot==90 || rot==255 || charttype==BARCHART)	{// Y axis title is almost always rotated, so use font height as offset
					if (charttype!=BARCHART) 		 
						axisMetrics.put("YAXISTITLEOFFSET", Math.ceil(ef.getFontHeightInPoints()*xpaddingfactor));
					else
						axisMetrics.put("XAXISTITLEOFFSET", Math.ceil(ef.getFontHeightInPoints()*ypaddingfactortitle));
				} else {// get width in desired rotation ****
					if (rot > 180) rot-=180;	// ensure angle is 0-90				
	    			java.awt.Font f= new java.awt.Font(ef.getFontName(), ef.getFontWeight(), (int)ef.getFontHeightInPoints());
		    		double len= StringTool.getApproximateStringWidth(f, ExcelTools.getFormattedStringVal(title, null)); 
		    		axisMetrics.put("YAXISTITLEOFFSET", Math.ceil(len*(Math.cos(Math.toRadians(rot)))));
				}*/
			}
			else
			{
				axisMetrics.put( "YAXISTITLEOFFSET", 20.0 );    // no axis title- add a little padding from edge
			}
		}
		// Label offsets
		StringBuffer[] series = new StringBuffer[0];
		try
		{
			double major = (Double) getMetric( "major" );
			double min = chartMetrics.get( "min" );
			double max = chartMetrics.get( "max" );
			// get java awt font so can compute approximate width of largest y axis label in order to compute YAXISLABELOFFSET
			if( major > 0 )
			{
				int nSeries = 0;
				if( min == 0 )    // usual case
				{
					nSeries = ((major != 0) ? ((int) (max / major) + 1) : 0);
				}
				else
				{
					nSeries = ((major != 0) ? ((int) ((max - min) / major) + 1) : 0);
				}
				nSeries = Math.abs( nSeries );
				series = new StringBuffer[nSeries];
				if( Math.floor( major ) != major )
				{    // contains a fractional part ... avoid java floating point issues
					// ensure y value matches scale/avoid double precision issues ... sigh ...
					String s = String.valueOf( major );
					int z = s.indexOf( "." );
					int scale = 0;
					if( z != -1 )
					{
						scale = s.length() - (z + 1);
					}
					z = 0;
					for( double i = min; i <= max; i += major )
					{
						java.math.BigDecimal bd = new java.math.BigDecimal( i ).setScale( scale, java.math.BigDecimal.ROUND_HALF_UP );
						series[z++] = new StringBuffer( CellFormatFactory.fromPatternString( (String) getMetric( "yPattern" ) )
						                                                 .format( bd.toString() ) );
					}
				}
				else
				{ // usual case of an int scale
					int z = 0;
					for( double i = min; i <= max; i += major )
					{
						series[z++] = new StringBuffer( CellFormatFactory.fromPatternString( (String) getMetric( "yPattern" ) )
						                                                 .format( i ) );
					}
				}
			}
		}
		catch( Exception e )
		{
			Logger.logWarn( "ChartAxes.getMetrics.  Error obtaining Series: " + e.toString() );
		}

		// Label Offsets ...
		if( this.hasAxis( XAXIS ) && (charttype != RADARCHART) && (charttype != RADARAREACHART) )
		{ //(Pie, donut, etc. don't have axes labels so disregard
			Object[] s;            // Determine X Axis Label offsets
			double width;
			AtomicInteger rot;
			com.extentech.formats.XLS.Font lf;
			String pattern;
			if( charttype != BARCHART )
			{
				s = categories;
				pattern = (String) getMetric( "xPattern" );
			}
			else
			{        // bar chart - series on x axis
				s = series;
				pattern = (String) getMetric( "yPattern" );
			}
			width = (chartMetrics.get( "w" ) / s.length) - 6; // ensure a bit of padding on either side
			lf = this.getLabelFont( XAXIS );
			rot = new AtomicInteger( this.getLabelRotation( XAXIS ) );        // if rot==0 and xaxis labels do not fit in width, a forced rotation will happen. ...
			double off = getLabelOffsets( lf, width, s, rot, pattern, true );
			axisMetrics.put( "xAxisRotate", rot.intValue() ); // possibly changed when calculating label offsets
			axisMetrics.put( "XAXISLABELOFFSET", off );
		}
		if( this.hasAxis( YAXIS ) && (charttype != RADARCHART) && (charttype != RADARAREACHART) )
		{    //(Pie, Donut, etc. don't have axes labels so disregard
			// for Y axis, determine width of labels and use as offset (except for bar charts, use height as offset) 			
			Object[] s;
			double width;
			AtomicInteger rot;
			com.extentech.formats.XLS.Font lf;
			String pattern;
			if( charttype != BARCHART )
			{
				s = series;
				pattern = (String) getMetric( "yPattern" );
			}
			else
			{    // bar chart- categories on y axis
				s = categories;
				pattern = (String) getMetric( "xPattern" );
			}
			rot = new AtomicInteger( this.getLabelRotation( YAXIS ) );        // if rot==0 and xaxis labels do not fit in width, a forced rotation will happen. ...
			lf = this.getLabelFont( YAXIS );
			width = (chartMetrics.get( "w" ) / 2) - 10; // a good guess?
			double off = getLabelOffsets( lf, width, s, rot, pattern, false );
			axisMetrics.put( "yAxisRotate", rot.intValue() );    // possibly changed when calculating label offsets
			axisMetrics.put( "YAXISLABELOFFSET", off ); // with padding
		}
		return axisMetrics;
	}

	/**
	 * returns the offset of the labels for the axis given the length of the label strings, the width of the
	 * <br>this will attempt to  break apart longer strings; if so will increment offset to accomodate multiple lines
	 *
	 * @param lf      label font
	 * @param width   max width of labels
	 * @param strings label strings
	 * @param rot     desired rotation, if any (0 if none)
	 * @param horiz   true if this is a horizontal axis
	 * @return
	 */
	private double getLabelOffsets( com.extentech.formats.XLS.Font lf,
	                                double width,
	                                Object[] strings,
	                                AtomicInteger rot,
	                                String pattern,
	                                boolean horiz )
	{
		double retwidth = 0;
		java.awt.Font f = null;
		double h = 0;
		try
		{
			// get awt Font so can compute and fit category in width
			f = new java.awt.Font( lf.getFontName(), lf.getFontWeight(), (int) lf.getFontHeightInPoints() );
			for( int i = 0; i < strings.length; i++ )
			{
				double w;
				AtomicInteger height = new AtomicInteger( 0 );
				StringBuffer s = null;
				try
				{
					s = new StringBuffer( CellFormatFactory.fromPatternString( pattern ).format( strings[i].toString() ) );
				}
				catch( IllegalArgumentException e )
				{ // trap error formatting
					s = new StringBuffer( strings[i].toString() );
				}
				if( horiz )    // on horizontal axis
				{
					w = getRotatedWidth( f, height, s.toString(), rot.intValue() );
				}
				else        // on vertical axis
				{
					w = getRotatedWidthVert( f, height, s.toString(), rot.intValue() );
				}
				h = Math.max( height.intValue(), h );
				w = addLinesToFit( f, s, rot, w, width, height );
				if( w > width )
				{
					width = w;
				}

				strings[i] = s;
				retwidth = Math.max( w, retwidth );
			}
		}
		catch( Exception e )
		{
		}
		if( horiz )
		{
			return h;
		}
		return retwidth;
		//		return offset + h;
	}

	/**
	 * get the width of string s in font f rotated by rot (0, 90, -90, 180)
	 *
	 * @param f   Font to display s in
	 * @param s   string to display
	 * @param rot rotation (0= none)
	 */
	private double getRotatedWidth( java.awt.Font f, AtomicInteger height, String s, int rot )
	{
		double retWidth = 0;
		String[] slines = s.split( "\n" );
		for( int i = 0; i < slines.length; i++ )
		{
			double width;
			if( rot == 0 )
			{
				width = StringTool.getApproximateStringWidth( f, CellFormatFactory.fromPatternString( null ).format( slines[i] ) );
				height.set( (int) Math.ceil( f.getSize() * 3 ) ); // width of the font + padding
			}
			else if( Math.abs( rot ) == 90 )
			{
				width = Math.ceil( f.getSize() * 2 ); // width of the font + padding
				height.set( (int) Math.max( height.intValue(), StringTool.getApproximateStringWidth( f, CellFormatFactory.fromPatternString(
						null ).format( slines[i] ) ) ) );
			}
			else
			{ // 45
				width = StringTool.getApproximateStringWidth( f, CellFormatFactory.fromPatternString( null ).format( slines[i] ) );
				width = Math.ceil( width * (Math.cos( Math.toRadians( rot ) )) );
				height.set( (int) Math.ceil( width * (Math.sin( Math.toRadians( rot ) )) ) );
			}
			retWidth = Math.max( width, retWidth );
		}
		return retWidth;
	}

	/**
	 * get the width of string s in font f rotated by rot (0, 90, -90, 180)
	 * on a Vertical axis
	 *
	 * @param f   Font to display s in
	 * @param s   string to display
	 * @param rot rotation (0= none)
	 */
	private double getRotatedWidthVert( java.awt.Font f, AtomicInteger height, String s, int rot )
	{
		double retWidth = 0;
		String[] slines = s.split( "\n" );
		for( int i = 0; i < slines.length; i++ )
		{
			double width;
			if( Math.abs( rot ) == 90 )
			{    // means VERTICAL orientation - default
				height.set( (int) StringTool.getApproximateStringWidth( f,
				                                                        CellFormatFactory.fromPatternString( null ).format( slines[i] ) ) );
				width = Math.ceil( f.getSize() * 2 ); // width of the font + padding
			}
			else if( rot == 0 )
			{    // means HORIZONTAL on vertical axis
				height.set( (int) Math.ceil( f.getSize() * 3 ) ); // width of font
				width = Math.max( height.intValue(), StringTool.getApproximateStringWidth( f,
				                                                                           CellFormatFactory.fromPatternString( null )
				                                                                                            .format( slines[i] ) ) );
			}
			else
			{ // 45
				width = StringTool.getApproximateStringWidth( f, CellFormatFactory.fromPatternString( null ).format( slines[i] ) );
				width = Math.ceil( width * (Math.cos( Math.toRadians( rot ) )) );
				height.set( (int) Math.ceil( width * (Math.sin( Math.toRadians( rot ) )) ) );
			}
			retWidth = Math.max( width, retWidth );
		}
		return retWidth;
	}

	/**
	 * if formatted width of string doesn't fit into width, break apart at spaces with new lines
	 * return
	 *
	 * @param f     font to display string in
	 * @param s     target String
	 * @param rot   rotation (0, 45, 90, -90 or 180)
	 * @param len   formatted string length taking into account display font
	 * @param width maximum width to display string
	 * @return
	 */
	private double addLinesToFit( java.awt.Font f, StringBuffer s, AtomicInteger rot, double len, double width, AtomicInteger height )
	{
		double retLen = Math.min( width, len );
		String str = s.toString().trim();

		while( len > width )
		{
			int lastSpace = -1;
			int j = s.lastIndexOf( "\n" ) + 1;
			len = -1;
			while( (len < width) && (j < str.length()) )
			{
				len += StringTool.getApproximateCharWidth( f, str.charAt( j ) );
				if( str.charAt( j ) == ' ' )
				{
					lastSpace = j;
				}
				j++;
			}
			if( len < width )
			{
				break;    // got it
			}

			if( lastSpace == -1 )
			{    // no spaces to break apart via \n's - rotate!
				if( str.indexOf( ' ' ) == -1 )
				{
					// see if string will fit in 45 degree rotation
					if( rot.intValue() != -90 )
					{
						rot.set( 45 );
						len = getRotatedWidth( f, height, str, rot.intValue() );
						if( len > width )
						{    // then all's fine
						}
						else    // doesn't fit in 45, must be 90 degrees
						{
							rot.set( -90 );
						}
					}
					if( rot.intValue() == -90 )
					{
						len = getRotatedWidth( f, height, str, rot.intValue() );
					}
					retLen = Math.max( len, retLen );
					break;
				}
				lastSpace = s.toString().indexOf( ' ' );
			}
			s.replace( lastSpace, lastSpace + 1, "\n" );    // + str.substring(lastSpace+1));
			str = s.toString();
		}
		return retLen;
	}
}
