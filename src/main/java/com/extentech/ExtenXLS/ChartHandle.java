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
package com.extentech.ExtenXLS;

import com.extentech.formats.OOXML.Layout;
import com.extentech.formats.OOXML.OOXMLConstants;
import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.OOXML.Title;
import com.extentech.formats.OOXML.TwoCellAnchor;
import com.extentech.formats.OOXML.TxPr;
import com.extentech.formats.XLS.Boundsheet;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.FormatConstants;
import com.extentech.formats.XLS.OOXMLAdapter;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import com.extentech.formats.XLS.charts.Axis;
import com.extentech.formats.XLS.charts.Chart;
import com.extentech.formats.XLS.charts.ChartConstants;
import com.extentech.formats.XLS.charts.ChartType;
import com.extentech.formats.XLS.charts.OOXMLChart;
import com.extentech.formats.XLS.charts.Series;
import com.extentech.formats.XLS.charts.ThreeD;
import com.extentech.toolkit.Logger;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserFactory;

import java.util.ArrayList;
import java.util.EnumSet;
import java.util.HashMap;
import java.util.Vector;

/**
 * Chart Handle allows for manipulation of Charts within a WorkBook.
 * <br>
 * Allows for run-time modification of Chart titles, labels, categories,
 * and data Cells.
 * <br>
 * Modification of Chart data cells allows you to completely modify
 * the data shown on the Chart.
 * <br><br>
 * To create a new chart and add to a worksheet, use:
 * ChartHandle.createNewChart(sheet, charttype, chartoptions).
 * <br><br>
 * To Obtain an array of existing chart handles, use
 * <br>WorkBookHandle.getCharts()
 * <br>or
 * WorkBookHandle.getCharts(chart title)
 *
 * @see WorkBookHandle
 */
public class ChartHandle implements ChartConstants
{
	// 20080114 KSC: delegate so visible 
	public static final int BARCHART = ChartConstants.BARCHART;
	public static final int COLCHART = ChartConstants.COLCHART;
	public static final int LINECHART = ChartConstants.LINECHART;
	public static final int PIECHART = ChartConstants.PIECHART;
	public static final int AREACHART = ChartConstants.AREACHART;
	public static final int SCATTERCHART = ChartConstants.SCATTERCHART;
	public static final int RADARCHART = ChartConstants.RADARCHART;
	public static final int SURFACECHART = ChartConstants.SURFACECHART;
	public static final int DOUGHNUTCHART = ChartConstants.DOUGHNUTCHART;
	public static final int BUBBLECHART = ChartConstants.BUBBLECHART;
	public static final int OFPIECHART = ChartConstants.OFPIECHART;
	public static final int PYRAMIDCHART = ChartConstants.PYRAMIDCHART;
	public static final int CYLINDERCHART = ChartConstants.CYLINDERCHART;
	public static final int CONECHART = ChartConstants.CONECHART;
	public static final int PYRAMIDBARCHART = ChartConstants.PYRAMIDBARCHART;
	public static final int CYLINDERBARCHART = ChartConstants.CYLINDERBARCHART;
	public static final int CONEBARCHART = ChartConstants.CONEBARCHART;
	public static final int RADARAREACHART = ChartConstants.RADARAREACHART;
	public static final int STOCKCHART = ChartConstants.STOCKCHART;

	// legacy
	public static final int BAR = BARCHART;
	public static final int COL = COLCHART;
	public static final int LINE = LINECHART;
	public static final int PIE = PIECHART;
	public static final int AREA = AREACHART;
	public static final int SCATTER = SCATTERCHART;
	public static final int RADAR = RADARCHART;
	public static final int SURFACE = SURFACECHART;
	public static final int DOUGHNUT = DOUGHNUTCHART;
	public static final int BUBBLE = BUBBLECHART;
	public static final int RADARAREA = RADARAREACHART;
	public static final int PYRAMID = PYRAMIDCHART;
	public static final int CYLINDER = CYLINDERCHART;
	public static final int CONE = CONECHART;
	public static final int PYRAMIDBAR = PYRAMIDBARCHART;
	public static final int CYLINDERBAR = CYLINDERBARCHART;
	public static final int CONEBAR = CONEBARCHART;
	// axis types
	public static final int XAXIS = ChartConstants.XAXIS;
	public static final int YAXIS = ChartConstants.YAXIS;
	public static final int ZAXIS = ChartConstants.ZAXIS;
	public static final int XVALAXIS = ChartConstants.XVALAXIS;    // an X axis type but VAL records  
	// coordinates
	public static final int X = 0;
	public static final int Y = 1;
	public static final int WIDTH = 2;
	public static final int HEIGHT = 3;

	protected WorkBookHandle wbh;
	private Chart mychart;

	/**
	 * Constructor which creates a new ChartHandle from an existing Chart Object
	 *
	 * @param Chart          c - the source Chart object
	 * @param WorkBookHandle wb - the parent WorkBookHandle
	 */
	public ChartHandle( Chart c, WorkBookHandle wb )
	{
		//super();
		mychart = c;
		wbh = wb;
		if( mychart.getWorkBook() == null )    // TODO: WHY IS THIS NULL????
		{
			mychart.setWorkBook( wb.getWorkBook() );
		}
	}

	/**
	 * Returns the title of the Chart
	 *
	 * @return String title of the Chart
	 */
	public String getTitle()
	{
		return mychart.getTitle();
	}

	/**
	 * Sets the title of the Chart
	 *
	 * @param String title - Chart title
	 */
	public void setTitle( String title )
	{
		this.mychart.setTitle( title );
	}

	/**
	 * returns the data range used by the chart
	 *
	 * @return
	 */
	public String getDataRangeJSON()
	{
		return this.mychart.getChartSeries().getDataRangeJSON().toString();
	}

	public int[] getEncompassingDataRange()
	{
		return getEncompassingDataRange( this.mychart.getChartSeries().getDataRangeJSON() );
	}

	/**
	 * returns the encompassing range for this chart, or null if the chart data
	 * is too complex to represent
	 *
	 * @param jsonDataRange
	 * @return
	 */
	public static int[] getEncompassingDataRange( JSONObject jsonDataRange )
	{
		try
		{
			String catrange = jsonDataRange.get( "c" ).toString();
			String sheet = catrange.substring( 0, catrange.indexOf( '!' ) );
			int[] retVals = ExcelTools.getRangeRowCol( catrange );
			int nSeries = jsonDataRange.getJSONArray( "Series" ).length();
			for( int i = 0; i < nSeries; i++ )
			{
				JSONObject series = (JSONObject) ((JSONArray) jsonDataRange.getJSONArray( "Series" )).get( i );
				String serrange = series.get( "v" ).toString();
				if( !serrange.startsWith( sheet ) )
				{
					continue;
				}
				int[] locs = ExcelTools.getRangeRowCol( serrange );
				try
				{
					if( locs[0] < retVals[0] )
					{
						retVals[0] = locs[0];
					}
					if( locs[1] < retVals[1] )
					{
						retVals[1] = locs[1];
					}
					if( locs[2] > retVals[2] )
					{
						retVals[2] = locs[2];
					}
					if( locs[3] > retVals[3] )
					{
						retVals[3] = locs[3];
					}

					String legendrange = series.get( "l" ).toString();
					locs = ExcelTools.getRowColFromString( legendrange );
					if( locs[0] < retVals[0] )
					{
						retVals[0] = locs[0];
					}
					if( locs[1] < retVals[1] )
					{
						retVals[1] = locs[1];
					}
					if( locs[0] > retVals[2] )
					{
						retVals[2] = locs[0];
					}
					if( locs[1] > retVals[3] )
					{
						retVals[3] = locs[1];
					}

					if( series.has( "b" ) )
					{
						String bubblerange = series.get( "b" ).toString();
						locs = ExcelTools.getRangeRowCol( serrange );
						if( locs[0] < retVals[0] )
						{
							retVals[0] = locs[0];
						}
						if( locs[1] < retVals[1] )
						{
							retVals[1] = locs[1];
						}
						if( locs[2] > retVals[2] )
						{
							retVals[2] = locs[2];
						}
						if( locs[3] > retVals[3] )
						{
							retVals[3] = locs[3];
						}
					}
				}
				catch( Exception e )
				{
					// just continue 	    			
				}
			}
			return retVals;
		}
		catch( Exception e )
		{
		}
		return null;
/*
        while (ptgs.hasNext()) {
        	PtgRef pr= (PtgRef) ptgs.next();
//        	PtgRef pr= (PtgRef)refs[i];
            int[] locs = pr.getIntLocation();
            for (int x=0;x<2;x++) {
                if((locs[x]<retValues[x]))retValues[x]=locs[x];
            }
            for (int x=2;x<4;x++) {
                if((locs[x]>retValues[x]))retValues[x]=locs[x];
            }
            i++;
        }
  */
	}

	/**
	 * returns the ordinal id associated with the underlying Chart Object
	 *
	 * @return int chart id
	 * @see WorkBookHandle.getChartById
	 */
	public int getId()
	{
		return this.mychart.getId();
	}

	/**
	 * returns the string representation of this ChartHandle
	 */
	public String toString()
	{
		return mychart.getTitle();
	}

	/***************************************************************************************************************************************/
	/**
	 * Returns an ordered array of strings representing all the series ranges in the Chart.
	 * <br>Each series can only represent one bar, line or wedge of data.
	 *
	 * @return String[] each item being a Cell Range representing one bar, line or wedge of data in the Chart
	 * @see ChartHandle.getCategories
	 */
	public String[] getSeries()
	{
		return mychart.getSeries( -1 ); // -1 is flag for all rather than for a specific chart
	}

	/**
	 * Returns an ordered array of strings representing all the category ranges in the chart.
	 * <br>This vector corresponds to the getSeries() method so will often contain duplicates,
	 * as while the series data changes frequently, category data
	 * is the same throughout the chart.
	 *
	 * @return String[] each item being a Cell Range representing the Category Data
	 * @see ChartHandle.getSeries
	 */
	public String[] getCategories()
	{
		return getCategories( -1 );    // -1 is flag for all rather than for a specific chart
	}

	/**
	 * Returns an ordered array of strings representing all the category ranges in the chart.
	 * <br>This vector corresponds to the getSeries() method so will often contain duplicates,
	 * as while the series data changes frequently, category data
	 * is the same throughout the chart.
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return String[] each item being a Cell Range representing the Category Data
	 * @see ChartHandle.getSeries
	 */
	private String[] getCategories( int nChart )
	{
		return mychart.getCategories( nChart );
	}

	/**
	 * Returns an array of ChartSeriesHandle Objects, one for each bar, line or wedge of data.
	 *
	 * @return ChartSeriesHandle[] Array of ChartSeriesHandle Objects representing Chart Series Data (Series and Categories)
	 * @see ChartSeriesHandle
	 */
	public ChartSeriesHandle[] getAllChartSeriesHandles()
	{
		return getAllChartSeriesHandles( -1 ); // get ALL
	}

	/**
	 * Returns an array of ChartSeriesHandle Objects for the desired chart, one for each bar, line or wedge of data.
	 * <br>A chart number of 0 means the default chart, 1-9 indicate series for overlay charts
	 * <br>NOTE: using this method returns the series for the desired chart ONLY
	 * <br>You MUST use the corresponding removeSeries(index, nChart) when removing series to properly match
	 * the series index.  Otherwise a mismatch will occur.
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return ChartSeriesHandle[] Array of ChartSeriesHandle Objects representing Chart Series Data (Series and Categories)
	 * @see ChartSeriesHandle
	 */
	public ChartSeriesHandle[] getAllChartSeriesHandles( int nChart )
	{
		Vector v = mychart.getAllSeries( nChart );
		ChartSeriesHandle[] csh = new ChartSeriesHandle[v.size()];
		for( int i = 0; i < v.size(); i++ )
		{
			Series s = (Series) v.get( i );
			csh[i] = new ChartSeriesHandle( s, this.wbh );
		}
		return csh;
	}

	/**
	 * Get the ChartSeriesHandle representing Chart Series Data (Series and Categories) for the specified Series range
	 *
	 * @param String seriesRange - For example, "Sheet1!A12:A21"
	 * @return ChartSeriesHandle
	 * @see ChartSeriesHandle
	 */
	public ChartSeriesHandle getChartSeriesHandle( String seriesRange )
	{
		ChartSeriesHandle[] series = this.getAllChartSeriesHandles();
		for( int i = 0; i < series.length; i++ )
		{
			String sr = series[i].getSeriesRange();
			if( seriesRange.equalsIgnoreCase( sr ) )
			{
				return series[i];
			}
		}
		return null;
	}

	/**
	 * Get the ChartSeriesHandle representing Chart Series Data (Series and Categories) for the specified Series index
	 *
	 * @param int idx - the index (0 based) of the series
	 * @return ChartSeriesHandle
	 * @see ChartSeriesHandle
	 */
	public ChartSeriesHandle getChartSeriesHandle( int idx )
	{
		ChartSeriesHandle[] series = this.getAllChartSeriesHandles();
		if( series.length >= idx )
		{
			return series[idx];
		}
		return null;
	}

	/**
	 * Get the ChartSeriesHandle representing Chart Series Data (Series and Categories) for the Series specified by label (legend)
	 *
	 * @param String legend - label for the desired series
	 * @return ChartSeriesHandle
	 * @see ChartSeriesHandle
	 */
	public ChartSeriesHandle getChartSeriesHandleByName( String legend )
	{
		Series s = mychart.getSeries( legend, -1 ); // -1 is flag for all rather than for a specific chart
		return new ChartSeriesHandle( s, this.wbh );
	}

	/**
	 * sets or removes the axis title
	 *
	 * @param axisType one of: XAXIS, YAXIS, ZAXIS
	 * @param ttl      String new title or null to remove
	 */
	public void setAxisTitle( int axisType, String ttl )
	{
		mychart.getAxes().setTitle( axisType, ttl );
	}

	/**
	 * returns the Y axis Title
	 *
	 * @return String title
	 */
	public String getYAxisLabel()
	{
		return mychart.getAxes().getTitle( YAXIS );
	}

	/**
	 * Sets the Y axis Title
	 *
	 * @param String yTitle	- new Y Axis title
	 */
	public void setYAxisLabel( String yTitle )
	{
		mychart.getAxes().setTitle( YAXIS, yTitle );
	}

	/**
	 * returns the X axis Title
	 *
	 * @return String title
	 */
	public String getXAxisLabel()
	{
		return mychart.getAxes().getTitle( XAXIS );
	}

	/**
	 * Sets the XAxisTitle
	 *
	 * @param String xTitle	- new X Axis title
	 */
	public void setXAxisLabel( String xTitle )
	{
		mychart.getAxes().setTitle( XAXIS, xTitle );
	}

	/**
	 * returns the Z axis Title, if any
	 *
	 * @return String Title
	 */
	public String getZAxisLabel()
	{
		return mychart.getAxes().getTitle( ZAXIS );
	}

	/**
	 * set the Z AxisTitle
	 *
	 * @param String zTitle	- new Z Axis Title
	 */
	public void setZAxisLabel( String zTitle )
	{
		mychart.getAxes().setTitle( ZAXIS, zTitle );
	}

	/**
	 * Sets the automatic scale option on or off for the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 * <p>Automatic Scaling automatically sets the scale maximum, minimum and tick units
	 * upon data changes, and is the default setting for charts
	 *
	 * @param int     axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @param boolean b - true if set Automatic scaling on, false otherwise
	 * @see setAxisAutomaticScale(boolean b)
	 */
	public void setAxisAutomaticScale( int axisType, boolean b )
	{
		mychart.getAxes().setAxisAutomaticScale( axisType, b );
		mychart.setDirtyFlag( true );
	}

	/**
	 * Returns the minimum scale value of the the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 *
	 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @return int Miminum Scale value for the desired axis
	 * @see getAxisMinScale()
	 */
	public double getAxisMinScale( int axisType )
	{
		double[] minmax = mychart.getMinMax( this.wbh );
		return mychart.getAxes().getMinMax( minmax[0], minmax[1], axisType )[0];
	}

	/**
	 * Returns the maximum scale value of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 *
	 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @return int - Maximum Scale value for the desired axis
	 * @see getAxisMaxScale()
	 */
	public double getAxisMaxScale( int axisType )
	{
		double[] minmax = mychart.getMinMax( this.wbh ); // -1 is flag for all rather than for a specific chart
		return mychart.getAxes().getMinMax( minmax[0], minmax[1], axisType )[1];
	}

	/**
	 * Returns the major tick unit of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 *
	 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @return int major tick unit
	 * @see getAxisMajorUnit()
	 */
	public int getAxisMajorUnit( int axisType )
	{
		return mychart.getAxes().getAxisMajorUnit( axisType );
	}

	/**
	 * Returns the minor tick unit of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 *
	 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @return int - minor tick unit of the desired axis
	 * @see getAxisMinorUnit()
	 */
	public int getAxisMinorUnit( int axisType )
	{
		return mychart.getAxes().getAxisMinorUnit( axisType );
	}

	/**
	 * Sets the maximum scale value of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 * <p>Note: The default scale setting for charts is known as Automatic Scaling
	 * <br>When data changes, the chart automatically adjusts the scale as necessary
	 * <br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
	 *
	 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @param int MaxValue - desired maximum value of the desired axis
	 * @see setAxisMax(int MaxValue)
	 */
	public void setAxisMax( int axisType, int MaxValue )
	{
		mychart.getAxes().setAxisMax( axisType, MaxValue );
		mychart.setDirtyFlag( true );
	}

	/**
	 * Sets the minimum scale value of the desired Value axis
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 * <p>Note: The default setting for charts is known as Automatic Scaling
	 * <br>When data values change, the chart automatically adjusts the scale as necessary
	 * <br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
	 *
	 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @param int MinValue - the desired Minimum scale value
	 * @see setAxisMin(int MinValue)
	 */
	public void setAxisMin( int axisType, int MinValue )
	{
		mychart.getAxes().setAxisMin( axisType, MinValue );
		mychart.setDirtyFlag( true );
	}

	/**
	 * Returns true if the desired Value axis is set to automatic scale
	 * <p>The Value axis contains numbers rather than labels, and is normally the Y axis,
	 * but Scatter and Bubble charts may have a value axis on the X Axis as well
	 * <p>Note: The default setting for charts is known as Automatic Scaling
	 * <br>When data changes, the chart automatically adjusts the scale as necessary
	 * <br>Setting the scale manually (either Minimum or Maximum Value) removes Automatic Scaling
	 *
	 * @param int axisType - one of ChartHandle.YAXIS, ChartHandle.XAXIS, ChartHandle.ZAXIS
	 * @return boolean true if Automatic Scaling is turned on
	 * @see getAxisAutomaticScale()
	 */
	public boolean getAxisAutomaticScale( int axisType )
	{
		return mychart.getAxes().getAxisAutomaticScale( axisType );
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
		return mychart.getAxisAutomaticScale();
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
		mychart.getAxes().setAxisAutomaticScale( b );
		mychart.setDirtyFlag( true );
	}

	/**
	 * Returns the minimum value of the Y Axis (Value Axis) scale
	 *
	 * @return int Miminum Scale value for Y axis
	 * @see getAxisMinScale(int axisType)
	 */
	public double getAxisMinScale()
	{
		double[] minmax = mychart.getMinMax( this.wbh ); // -1 is flag for all rather than for a specific chart
		return mychart.getAxes().getMinMax( minmax[0], minmax[1] )[0];
	}

	/**
	 * Returns the maximum value of the Y Axis (Value Axis) scale
	 *
	 * @return int Maximum Scale value for Y axis
	 * @see getAxisMaxScale(int axisType)
	 */
	public double getAxisMaxScale()
	{
		double[] minmax = mychart.getMinMax( this.wbh );
		return mychart.getAxes().getMinMax( minmax[0], minmax[1] )[1];
	}

	/**
	 * Returns the major tick unit of the Y Axis (Value Axis)
	 *
	 * @return int major tick unit
	 * @see getAxisMajorUnit(int axisType)
	 */
	public int getAxisMajorUnit()
	{
		double[] minmax = mychart.getMinMax( this.wbh );
		return (int) mychart.getAxes().getMinMax( minmax[0], minmax[1] )[2];
	}

	/**
	 * Returns the minor tick unit of the Y Axis (Value Axis)
	 *
	 * @return int minor tick unit
	 * @see getAxisMinorUnit(int axisType)
	 */
	public int getAxisMinorUnit()
	{
		double[] minmax = mychart.getMinMax( this.wbh );
		return (int) mychart.getAxes().getMinMax( minmax[0], minmax[1] )[1];
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
		mychart.getAxes().setAxisMax( MaxValue );
		mychart.setDirtyFlag( true );
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
		mychart.getAxes().setAxisMin( MinValue );
		mychart.setDirtyFlag( true );
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
	public void setAxisOption( int axisType, String op, String val )
	{
		mychart.getAxes().setChartOption( axisType, op, val );

	}

	/**
	 * set the font for the chart title
	 *
	 * @param String  name - font name
	 * @param int     height - font height in 1/20 point units
	 * @param boolean bold - true if bold
	 * @param boolean italic - true if italic
	 * @param boolean underline	- true if underlined
	 */
	public void setTitleFont( String name, int height, boolean bold, boolean italic, boolean underline )
	{
		Font f = new Font( name, 200, height );
		f.setBold( bold );
		f.setItalic( italic );
		f.setUnderlined( underline );
		int idx = wbh.getWorkBook().getFontIdx( f );
		if( idx == -1 )
		{ // font doesn't exist yet, add to streamer
			f.setIdx( -1 );        //flag to insert
			idx = wbh.getWorkBook().insertFont( f ) + 1;
		}
		mychart.setTitleFont( idx );
	}

	/**
	 * set the font for the Chart Title
	 *
	 * @param com.extentech.Formats.XLS.Font f - desired font for the Chart Title
	 * @see Font
	 */
	public void setTitleFont( Font f )
	{
		int idx = wbh.getWorkBook().getFontIdx( f );
		if( idx == -1 )
		{ // font doesn't exist yet, add to streamer
			f.setIdx( -1 );        //flag to insert
			idx = wbh.getWorkBook().insertFont( f ) + 1;
		}
		mychart.setTitleFont( idx );
	}

	/**
	 * returns the Font associated with the Chart Title
	 *
	 * @return com.extentech.Formats.XLS.Font
	 * @see Font
	 */
	public Font getTitleFont()
	{
		return mychart.getTitleFont();
	}

	/**
	 * set the font for all axes on the chart
	 *
	 * @param String  name		- font name
	 * @param int     height	- font height in 1/20 point units
	 * @param boolean bold		- true if bold
	 * @param boolean italic	- true if italic
	 * @param boolean underline	- true if underlined
	 */
	public void setAxisFont( String name, int height, boolean bold, boolean italic, boolean underline )
	{
		Font f = new Font( name, 200, height );
		f.setBold( bold );
		f.setItalic( italic );
		f.setUnderlined( underline );
		int idx = wbh.getWorkBook().getFontIdx( f );
		if( idx == -1 )
		{ // font doesn't exist yet, add to streamer
			f.setIdx( -1 );        //flag to insert
			idx = wbh.getWorkBook().insertFont( f ) + 1;
		}
		mychart.getAxes().setTitleFont( XAXIS, idx );
		mychart.getAxes().setTitleFont( YAXIS, idx );
		mychart.getAxes().setTitleFont( ZAXIS, idx );
		mychart.setDirtyFlag( true );
	}

	/**
	 * set the font for all axes on the Chart
	 *
	 * @param com.extentech.Formats.XLS.Font f - desired font for the Chart axes
	 * @see Font
	 */
	public void setAxisFont( Font f )
	{
		int idx = wbh.getWorkBook().getFontIdx( f );
		if( idx == -1 )
		{ // font doesn't exist yet, add to streamer
			f.setIdx( -1 );        //flag to insert
			idx = wbh.getWorkBook().insertFont( f ) + 1;
		}
		mychart.getAxes().setTitleFont( XAXIS, idx );
		mychart.getAxes().setTitleFont( YAXIS, idx );
		mychart.getAxes().setTitleFont( ZAXIS, idx );
		mychart.setDirtyFlag( true );
	}

	/**
	 * return the Font associated with the Chart Axes
	 *
	 * @return com.extentech.Formats.XLS.Font for Chart Axes or null if no axes
	 * @see Font
	 */
	public Font getAxisFont()
	{
		Font f = null;
		f = mychart.getAxes().getTitleFont( XAXIS );
		if( f != null )
		{
			return f;
		}
		f = mychart.getAxes().getTitleFont( YAXIS );
		if( f != null )
		{
			return f;
		}
		f = mychart.getAxes().getTitleFont( ZAXIS );
		if( f != null )
		{
			return f;
		}
		return null;
	}

	/**
	 * resets all fonts in the chart to the default font of the workbook
	 */
	public void resetFonts()
	{
		mychart.resetFonts();
	}

	/**
	 * returns the underlhying Sheet Object this Chart is attached to
	 * <br>For Internal Use
	 *
	 * @return Boundsheet
	 */
	public Boundsheet getSheet()
	{
		return mychart.getSheet();
	}

	/**
	 * return the background color of this chart's Plot Area as an int
	 *
	 * @return int background color constant
	 * @see FormatHandle.COLOR_* constants
	 */
	public int getPlotAreaBgColor()
	{
		String bg = mychart.getPlotAreaBgColor();
		return FormatHandle.HexStringToColorInt( bg, (short) 0 );
	}

	public String getPlotAreaBgColorStr()
	{
		return mychart.getPlotAreaBgColor();
	}

	/**
	 * sets the Plot Area background color
	 *
	 * @param int bg - color constant
	 * @see FormatHandle.COLOR_* constants
	 */
	public void setPlotAreaBgColor( int bg )
	{
		mychart.setPlotAreaBgColor( bg );
	}

	/**
	 * Change the value of a Chart object.
	 * <br><b>NOTE: THIS HAS NOT BEEN 100% IMPLEMENTED YET</b>
	 * <br>You can use this method to change:
	 * <p/>
	 * <br>- the Title of the Chart
	 * <br>- the Text Labels of Categories and Values (X and Y)
	 * <p/>
	 * <br>eg:
	 * <p/>
	 * <br>	To change the value of the Chart title
	 * <br>	chart.changeObjectValue("Template Chart Title", "Widget Sales By Quarter");
	 * <p/>
	 * <br>	To change the text label of the categories
	 * <br>	chart.changeObjectValue("Category X", "Fiscal Year");
	 * <p/>
	 * <br>	To change the text label of the values
	 * <br>	chart.changeObjectValue("Value Y", "Sales in US$");
	 *
	 * @param String originalval - One of: "Template Chart Title", "Category X" or "Value Y"
	 * @param Sring  newval - the new setting
	 * @return whether the change was successful
	 */
	public boolean changeTextValue( String originalval, String newval )
	{
/* KSC: TODO: Refactor ***      
		for(int x=0;x<mychart.aivals.size();x++){
			Ai ser = (Ai)mychart.aivals.get(x);
			if(ser.toString().equalsIgnoreCase(originalval)){
				return	ser.setText(newval);
			}
		}*/
		return mychart.changeTextValue( originalval, newval );
//		return false;
	}

	/**
	 * Sets the location lock on the Cell Reference at the
	 * specified  location
	 * <p/>
	 * Used to prevent updating of the Cell Reference when
	 * Cells are moved.
	 *
	 * @param location of the Cell Reference to be locked/unlocked
	 * @param lock     status setting
	 * @return boolean whether the Cell Reference was found and modified
	 */
	private boolean setLocationPolicy( String loc, int l )
	{
		Logger.logErr( "ChartHandle.setLocationPolicy is broken" );
		
/* TODO: Refactor	    for(int x=0;x<mychart.aivals.size();x++){
			Ai ser = (Ai)mychart.aivals.get(x);
			if(ser.setLocationPolicy(loc,l)){
				return	true;
			}
		}*/
		return false;
	}

	/**
	 * Sets the Chart type to the specified basic type (no 3d, no stacked ...)
	 * <br>To see possible Chart Types, view the public static int's in ChartHandle.
	 * <br>Possible Chart Types:
	 * <br>BARCHART
	 * <br>COLCHART
	 * <br>LINECHART
	 * <br>PIECHART
	 * <br>AREACHART
	 * <br>SCATTERCHART
	 * <br>RADARCHART
	 * <br>SURFACECHART
	 * <br>DOUGHNUTCHART
	 * <br>BUBBLECHART
	 * <br>RADARAREACHART
	 * <br>PYRAMIDCHART
	 * <br>CYLINDERCHART
	 * <br>CONECHART
	 * <br>PYRAMIDBARCHART
	 * <br>CYLINDERBARCHART
	 * <br>CONEBAR
	 *
	 * @param int chartType - representing the chart type
	 */
	public void setChartType( int chartType )
	{
		mychart.setChartType( chartType, 0, EnumSet.noneOf( ChartOptions.class ) );    // no specific options
	}

	/**
	 * Chart Options.
	 * CLUSTERED, 		 bar, col charts only
	 * STACKED,
	 * PERCENTSTACKED,  100% stacked
	 * THREED,			 3d Effect
	 * EXPLODED,		 Pie, Donut
	 * HASLINES,		 Scatter, Line charts ...
	 * WIREFRAME,		 Surface
	 * DROPLINES,
	 * DOWNBARS,		 line, stock
	 * UPDOWNBARS,		 line, stock
	 * SERLINES,		 bar, ofpie
	 * <br>Use these chart options when creating new charts
	 * <br>A chart may have multiple chart options e.g. 3D Exploded pie chart
	 *
	 * @see ChartHandle.createNewChart
	 */
	// need: hasMarkers ****
	public enum ChartOptions
	{
		CLUSTERED, /**
	 * bar, col charts only
	 */
	STACKED,
		PERCENTSTACKED, /**
	 * 100% stacked
	 */
	THREED, /**
	 * 3d Effect
	 */
	EXPLODED, /**
	 * Pie, Donut
	 */
	HASLINES, /**
	 * Scatter, Line
	 */
	SMOOTHLINES, /**
	 * Scatter, Line, Radar
	 */
	WIREFRAME, /**
	 * Surface
	 */
	DROPLINES, /**
	 * line, area, stock charts
	 */
	UPDOWNBARS, /**
	 * line, stock
	 */
	SERLINES, /**
	 * bar, ofpie
	 */
	HILOWLINES, /**
	 * line, stock charts
	 */
	FILLED            /** radar */
	}

	/**
	 * Static method to create a new chart on WorkSheet sheet of type chartType with chart Options options
	 * <br>After creating, you can set the chart title via ChartHandle.setTitle
	 * <br>and Position via ChartHandle.setRelativeBounds (row/col-based) or ChartHandle.setCoords (pixel-based)
	 * <br>as well as several other customizations possible
	 *
	 * @param book      WorkBookHandle
	 * @param sheet     WorkSheetHandle
	 * @param chartType one of:
	 *                  <br>BARCHART
	 *                  <br>COLCHART
	 *                  <br>LINECHART
	 *                  <br>PIECHART
	 *                  <br>AREACHART
	 *                  <br>SCATTERCHART
	 *                  <br>RADARCHART
	 *                  <br>SURFACECHART
	 *                  <br>DOUGHNUTCHART
	 *                  <br>BUBBLECHART
	 *                  <br>RADARAREACHART
	 *                  <br>PYRAMIDCHART
	 *                  <br>CYLINDERCHART
	 *                  <br>CONECHART
	 *                  <br>PYRAMIDBARCHART
	 *                  <br>CYLINDERBARCHART
	 *                  <br>CONEBARCHART
	 * @param options   EnumSet<ChartOptions>
	 * @return
	 * @see ChartHandle.ChartOptions
	 * @see setChartType
	 */
	public static ChartHandle createNewChart( WorkSheetHandle sheet, int chartType, EnumSet<ChartOptions> options )
	{
		// Create Initial Basic Chart
		ChartHandle cht = sheet.getWorkBook().createChart( "", sheet );
		// Change Chart Type with Desired Options:
		cht.setChartType( chartType, 0, options );
		return cht;
	}

	/**
	 * Sets the Chart type to the specified type (no 3d, no stacked ...)
	 * <br>To see possible Chart Types, view the public static int's in ChartHandle.
	 * <br>Possible Chart Types:
	 * <br>BARCHART
	 * <br>COLCHART
	 * <br>LINECHART
	 * <br>PIECHART
	 * <br>AREACHART
	 * <br>SCATTERCHART
	 * <br>RADARCHART
	 * <br>SURFACECHART
	 * <br>DOUGHNUTCHART
	 * <br>BUBBLECHART
	 * <br>RADARAREACHART
	 * <br>PYRAMIDCHART
	 * <br>CYLINDERCHART
	 * <br>CONECHART
	 * <br>PYRAMIDBARCHART
	 * <br>CYLINDERBARCHART
	 * <br>CONEBARCHART
	 *
	 * @param int                   chartType - representing the chart type
	 * @param nChart                - 0 (default) or 1-9 for complex overlay charts
	 * @param EnumSet<ChartOptions> 0 or more chart options (Such as Stacked, Exploded ...)
	 * @see ChartHandle.ChartOptions
	 */
	public void setChartType( int chartType, int nChart, EnumSet<ChartOptions> options )
	{
		mychart.setChartType( chartType, nChart, options );
	}

	/**
	 * Sets the basic chart type (no 3d, stacked...) for multiple or overlay Charts.
	 * <br>You can specify the drawing order of the Chart, where 0 is the default chart,
	 * and 1-9 are overlay charts.
	 * <br>The default chart (chart 0) is always present; however, using this method, you can create a
	 * new overlay chart (up to 9 maximum).
	 * <br>NOTE: The chart number must be <b>unique</b> and <b>in order</b>
	 * <br>If the desired chart number is not present in the chart, a new overlay chart will be created.
	 * <br><b>To set explicit chart options, @see setChartType(chartType, nChart, is3d, isStacked, is100PercentStacked)</b>
	 * <br>
	 * <br>To see possible Chart Types, view the public static int's in ChartHandle.
	 * <br>Possible Chart Types:
	 * <br>BARCHART
	 * <br>COLCHART
	 * <br>LINECHART
	 * <br>PIECHART
	 * <br>AREACHART
	 * <br>SCATTERCHART
	 * <br>RADARCHART
	 * <br>SURFACECHART
	 * <br>DOUGHNUTCHART
	 * <br>BUBBLECHART
	 * <br>RADARAREACHART
	 * <br>PYRAMIDCHART
	 * <br>CYLINDERCHART
	 * <br>CONECHART
	 * <br>PYRAMIDBARCHART
	 * <br>CYLINDERBARCHART
	 * <br>CONEBARCHART
	 *
	 * @param int       chartType - representing the chart type
	 * @param chartType
	 * @param nChart    number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 */
	public void setChartType( int chartType, int nChart )
	{
		mychart.setChartType( chartType, nChart, EnumSet.noneOf( ChartOptions.class ) );    // no specific options
	}

	/**
	 * Return an int corresponding to this ChartHandle's Chart Type
	 * for the default chart
	 * <br>To see possible Chart Types, view the public static int's in ChartHandle.
	 *
	 * @return int chart type
	 * @see ChartHandle static Chart Type Constants
	 * @see ChartHandle.setChartType
	 */
	public int getChartType()
	{
		return mychart.getChartType();
	}

	/**
	 * Return an int corresponding to this ChartHandle's Chart Type
	 * for the specified chart
	 * <br>To see possible Chart Types, view the public static int's in ChartHandle.
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return int chart type
	 * @see ChartHandle static Chart Type Constants
	 * @see ChartHandle.setChartType
	 */
	public int getChartType( int nChart )
	{
		return mychart.getChartType( nChart );
	}

	/** Sets the location lock on the Cell Reference at the
	 *  specified  location
	 *
	 * 	Used to prevent updating of the Cell Reference when
	 *  Cells are moved.
	 *
	 * @param location of the Cell Reference to be locked/unlocked
	 * @param lock status setting
	 * @return boolean whether the Cell Reference was found and modified
	 */
/* TODO: NEEDED??	public boolean setLocationLocked(String loc, boolean l){
		int x = Ptg.PTG_LOCATION_POLICY_UNLOCKED;
		if(l)x = Ptg.PTG_LOCATION_POLICY_LOCKED;
		return setLocationPolicy(loc, x);
	} */

	/**
	 * Change the Cell Range referenced by one of the Series (a bar, line or wedge of data) in the Chart.
	 * <br>
	 * <br>For Example, if the data values for one of the Series in the Chart are obtained from the range Sheet1!A1:A10
	 * and we want to add 5 more values to that Series, use:
	 * <p/>
	 * <br>boolean changedOK = charthandle.changeSeriesRange("Sheet1!A1:A10","Sheet1!A1:A15");
	 * <br>
	 * <br>Please keep in mind this is only the data range; it does not include labels that may have
	 * been automatically created when you chose the chart range.
	 * <br>
	 * <br>To illustrate this, if A1 = "label" A2 = "data" A3 = "data", and we want to add 2 more data points, we would use:
	 * changeSeriesRange("Sheet1!A2:A3", "Sheet1!A2:A5");
	 * <br>
	 * <br>Series are always expressed as one single line of data.  If your chart encompasses a range of rows and columns
	 * you will need to modify each of the series in the chart handle.  To determine the series that already exist in your chart,
	 * utilize the String[] getSeries() method.
	 *
	 * @param String originalrange - the original Series (bar, line or wedge of data) to alter
	 * @param String newrange -the new data range
	 * @return whether the change was successful
	 */
	public boolean changeSeriesRange( String originalrange, String newrange )
	{
		return mychart.changeSeriesRange( originalrange, newrange );
	}

	/**
	 * Change the Cell Range representing the Categories in the Chart.
	 * <br>Categories usually appear on the X Axis and are textual, not numeric
	 * <br>For example: the Category values in the Chart are obtained from the range Sheet1!A1:A10 and
	 * we want to add 5 more categories to the chart:
	 * <br>boolean changedOK = chart.changeCategoryRange("Sheet1!A1:A10","Sheet1!A1:A15");
	 * <br>Note that Category Range is the same for each Series (bar, line or wedge of data)
	 * <br>i.e. there is only one Category Range for the Chart, but there may be many Series Ranges
	 *
	 * @param String originalrange - Original Category Range
	 * @param String newrange - New Category Range
	 * @return true if the change was successful
	 */
	public boolean changeCategoryRange( String originalrange, String newrange )
	{
		return mychart.changeCategoryRange( originalrange, newrange );
	}

	/**
	 * Changes or adds a Series to the chart via Series Index.  Each bar, line or wedge in a chart represents a Series.
	 * <br>If the Series index is greater than the number of series already present in the chart, the series will be added to the end.
	 * <br>Otherwise the Series at the index position will be altered.
	 * <br>This method allows altering of every aspect of the Series:  Data (Series) Range, Legend Cell Address, Category Range and/or Bubble Range.
	 *
	 * @param int    index		- the series index.  If greater than the number of series already present in the chart, the series will be added to the end
	 * @param String legendCell	- String representation of Legend Cell Address
	 * @param String categoryRange	- String representation of Category Range (should be same for all series)
	 * @param String seriesRange - String representation of the Series Data Range for this series
	 * @param String bubbleRange -	String representation of Bubble Range (representing bubble sizes), if bubble chart. null if not
	 * @return a ChartSeriesHandle representing the new or altered Series
	 * @throws CellNotFoundException
	 */
	public ChartSeriesHandle setSeries( int index, String legendCell, String categoryRange, String seriesRange, String bubbleRange ) throws
	                                                                                                                                 CellNotFoundException
	{
		String legendText = "";
		try
		{
			CellHandle ICell = null;
			if( legendCell != null && !legendCell.equals( "" ) )
			{
				// 20070707 KSC: allow addition of new cell ranges for legendCell (see ExtenXLS.handleChartElement)
				try
				{
					ICell = wbh.getCell( legendCell );
				}
				catch( CellNotFoundException c )
				{
					int shtpos = legendCell.indexOf( "!" );
					if( shtpos > 0 )
					{
						String sheetstr = legendCell.substring( 0, shtpos );
						WorkSheetHandle sht = wbh.getWorkSheet( sheetstr );
						String celstr = legendCell.substring( shtpos + 1 );
						ICell = sht.add( "", celstr );
					}
				}
				legendText = ICell.getStringVal();
			}
			return setSeries( index, legendCell, legendText, categoryRange, seriesRange, bubbleRange );
		}
		catch( WorkSheetNotFoundException e )
		{
			throw new CellNotFoundException( "Error locating cell for adding series range: " + legendCell );
		}
	}

	/**
	 * Changes or adds a Series to the chart via Series Index.  Each bar, line or wedge in a chart represents a Series.
	 * <br>If the Series index is greater than the number of series already present in the chart, the series will be added to the end.
	 * <br>Otherwise the Series at the index position will be altered.
	 * <br>This method allows altering of every aspect of the Series:  Data (Series) Range, Legend Text, Legend Cell Address, Category Range and/or Bubble Range.
	 *
	 * @param int    index		- the series index.  If greater than the number of series already present in the chart, the series will be added to the end
	 * @param String legendCell	- String representation of Legend Cell Address
	 * @param String legendText	- String Legend text
	 * @param String categoryRange	- String representation of Category Range (should be same for all series)
	 * @param String seriesRange - String representation of the Series Data Range for this series
	 * @param String bubbleRange -	String representation of Bubble Range (representing bubble sizes), if bubble chart. null if not
	 * @return a ChartSeriesHandle representing the new or altered Series
	 * @throws CellNotFoundException
	 */
	public ChartSeriesHandle setSeries( int index,
	                                    String legendCell,
	                                    String legendText,
	                                    String categoryRange,
	                                    String seriesRange,
	                                    String bubbleRange ) throws CellNotFoundException
	{
		return setSeries( index, legendCell, legendText, categoryRange, seriesRange, bubbleRange, 0 ); // for default chart		
	}

	/**
	 * Changes or adds a Series to the desired Chart (either default or overlay) via Series Index.  Each bar, line or wedge in a chart represents a Series.
	 * <br>If the Series index is greater than the number of series already present in the chart, the series will be added to the end.
	 * <br>Otherwise the Series at the index position will be altered.
	 * <br>This method allows altering of every aspect of the Series:  Data (Series) Range, Legend Text, Legend Cell Address, Category Range and/or Bubble Range.
	 *
	 * @param int    index		- the series index.  If greater than the number of series already present in the chart, the series will be added to the end
	 * @param String legendCell	- String representation of Legend Cell Address
	 * @param String legendText	- String Legend text
	 * @param String categoryRange	- String representation of Category Range (should be same for all series)
	 * @param String seriesRange - String representation of the Series Data Range for this series
	 * @param String bubbleRange -	String representation of Bubble Range (representing bubble sizes), if bubble chart. null if not
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return a ChartSeriesHandle representing the new or altered Series
	 * @throws CellNotFoundException
	 */
	public ChartSeriesHandle setSeries( int index,
	                                    String legendCell,
	                                    String legendText,
	                                    String categoryRange,
	                                    String seriesRange,
	                                    String bubbleRange,
	                                    int nChart ) throws CellNotFoundException
	{
//        if (index < mychart.getAllSeries(nChart).size() && index >= 0) {
		try
		{
			Series s = (Series) mychart.getAllSeries( nChart ).get( index );
			ChartSeriesHandle csh = new ChartSeriesHandle( s, this.wbh );
			csh.setSeries( legendCell, categoryRange, seriesRange, bubbleRange );
			setDimensionsRecord();
			return csh;
		}
		catch( ArrayIndexOutOfBoundsException ae )
		{    // not found - add
			return addSeriesRange( legendCell, legendText, categoryRange, seriesRange, bubbleRange, nChart );
		}
	}

	/**
	 * Adds a new Series to the chart.  Each bar, line or wedge in a chart represents a Series.
	 *
	 * @param String legendAddress -	The cell address defining the legend for the series
	 * @param Srring legendText	- Text of the legend
	 * @param String categoryRange	- Cell Range defining the category (normally will be the same range for all series)
	 * @param String seriesRange - Cell range defining the data points of the series
	 * @param String bubbleRange	- Cell range defining the bubble sizes for this series (bubble charts only)
	 * @return ChartSeriesHandle representing the new series
	 * @throws CellNotFoundException
	 */
	private ChartSeriesHandle addSeriesRange( String legendAddress,
	                                          String legendText,
	                                          String categoryRange,
	                                          String seriesRange,
	                                          String bubbleRange ) throws CellNotFoundException
	{
		return addSeriesRange( legendAddress, legendText, categoryRange, seriesRange, bubbleRange, 0 ); // for default chart
	}

	/**
	 * Adds a new Series to the chart.  Each bar, line or wedge in a chart represents a Series.
	 *
	 * @param String legendAddress -	The cell address defining the legend for the series
	 * @param Srring legendText	- Text of the legend
	 * @param String categoryRange	- Cell Range defining the category (normally will be the same range for all series)
	 * @param String seriesRange - Cell range defining the data points of the series
	 * @param String bubbleRange	- Cell range defining the bubble sizes for this series (bubble charts only)
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return ChartSeriesHandle representing the new series
	 * @throws CellNotFoundException
	 */
	private ChartSeriesHandle addSeriesRange( String legendAddress,
	                                          String legendText,
	                                          String categoryRange,
	                                          String seriesRange,
	                                          String bubbleRange,
	                                          int nChart ) throws CellNotFoundException
	{
		Series s = null;
		if( bubbleRange == null || bubbleRange.equals( "" ) )
		{
			s = mychart.addSeries( seriesRange, categoryRange, "", legendAddress, legendText, nChart );
		}
		else
		{
			s = mychart.addSeries( seriesRange, categoryRange, bubbleRange, legendAddress, legendText, nChart );
		}
		if( nChart > 0 )
		{ // must update SeriesList record for overlay charts
			// TODO: FINISH
		}
		setDimensionsRecord();
		return new ChartSeriesHandle( s, wbh );
	}

	/**
	 * Adds a new Series to the chart.  Each bar, line or wedge in a chart represents a Series.
	 * <br>An Example of adding multiple series to a chart:
	 * <p>ChartHandle.addSeriesRange("Sheet1!A3", "Sheet1!B1:E1", "Sheet1:B3:E3", null);
	 * <br>ChartHandle.addSeriesRange("Sheet1!A4", "Sheet1!B1:E1", "Sheet1:B4:E4", null);
	 * <br>ChartHandle.addSeriesRange("Sheet1!A5", "Sheet1!B1:E1", "Sheet1:B5:E5", null);
	 * <br>etc...
	 * <p>Note that the category does not change, it is usually constant
	 * through series.
	 * <br>Also note that the example above is for a non-bubble-type chart.
	 *
	 * @param String legendCell - Cell reference for the legend cell (e.g. Sheet1!A1)
	 * @param String categoryRange - Category Cell range (e.g. Sheet1!B1:B1);
	 * @param String seriesRange - Series Data range (e.g. Sheet1!B3:E3);
	 * @param String bubbleRange - Cell Range representing Bubble sizes (e.g. Sheet1!A2:A5); or null if chart is not of type Bubble.
	 * @return ChartSeriesHandle referencing the newly added series
	 */
	public ChartSeriesHandle addSeriesRange( String legendCell, String categoryRange, String seriesRange, String bubbleRange ) throws
	                                                                                                                           CellNotFoundException
	{
		return this.addSeriesRange( legendCell, categoryRange, seriesRange, bubbleRange, 0 );    // target default chart
	}

	/**
	 * Adds a new Series to the chart.  Each bar, line or wedge in a chart represents a Series.
	 * <br>An Example of adding multiple series to a chart:
	 * <p>ChartHandle.addSeriesRange("Sheet1!A3", "Sheet1!B1:E1", "Sheet1:B3:E3", null);
	 * <br>ChartHandle.addSeriesRange("Sheet1!A4", "Sheet1!B1:E1", "Sheet1:B4:E4", null);
	 * <br>ChartHandle.addSeriesRange("Sheet1!A5", "Sheet1!B1:E1", "Sheet1:B5:E5", null);
	 * <br>etc...
	 * <p>Note that the category does not change, it is usually constant
	 * through series.
	 * <br>Also note that the example above is for a non-bubble-type chart.
	 *
	 * @param String legendCell - Cell reference for the legend cell (e.g. Sheet1!A1)
	 * @param String categoryRange - Category Cell range (e.g. Sheet1!B1:B1);
	 * @param String seriesRange - Series Data range (e.g. Sheet1!B3:E3);
	 * @param String bubbleRange - Cell Range representing Bubble sizes (e.g. Sheet1!A2:A5); or null if chart is not of type Bubble.
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return ChartSeriesHandle referencing the newly added series
	 */
	public ChartSeriesHandle addSeriesRange( String legendCell,
	                                         String categoryRange,
	                                         String seriesRange,
	                                         String bubbleRange,
	                                         int nChart ) throws CellNotFoundException
	{
		String legendText = "";
		String legendAddr = "";
		try
		{
			CellHandle ICell = null;
			if( legendCell != null && !legendCell.equals( "" ) )
			{
				try
				{
					ICell = wbh.getCell( legendCell );
					if( legendCell.indexOf( "!" ) == -1 )
					{
						legendAddr = ICell.getWorkSheetName() + "!" + ICell.getCellAddress();
					}
					else
					{
						legendAddr = legendCell;
					}
				}
				catch( CellNotFoundException c )
				{
					int shtpos = legendCell.indexOf( "!" );
					if( shtpos > 0 )
					{
						String sheetstr = legendCell.substring( 0, shtpos );
						WorkSheetHandle sht = wbh.getWorkSheet( sheetstr );
						String celstr = legendCell.substring( shtpos + 1 );
						ICell = sht.add( "", celstr );    // TODO: Why is this being added?
						legendAddr = celstr;
					}
				}
				if( ICell != null )
				{
					legendText = ICell.getStringVal();
				}
				else
				{
					legendText = legendCell;
				}
			}
			Series s = null;
			if( bubbleRange == null )
			{
				s = mychart.addSeries( seriesRange, categoryRange, "", legendAddr, legendText, nChart );
			}
			else
			{
				s = mychart.addSeries( seriesRange, categoryRange, bubbleRange, legendAddr, legendText, nChart );
			}
			setDimensionsRecord();    // update chart DIMENSIONS record upon update of series        	
			return new ChartSeriesHandle( s, wbh );
		}
		catch( WorkSheetNotFoundException e )
		{
			throw new CellNotFoundException( "Error locating cell for adding series range: " + legendCell );
		}
	}

	/**
	 * Adds a new Series to the chart via CellHandles and CellRange Objects.  Each bar, line or wedge in a chart represents a Series.
	 *
	 * @param CellHandle legendCell - references the legend cell for this series
	 * @param CellRange  categoryRange - The CellRange referencing the category (should be the same for all Series)
	 * @param CelLRange  seriesRange - The CellRange referencing the data points for one bar, line or wedge in the chart
	 * @param CellRange  bubbleRange -The CellRange referencing bubble sizes for this series, or null if chart is not of type BUBBLE
	 * @return ChartSeriesHandle referencing the newly added series
	 * @see ChartHandle.addSeriesRange(String legendCell, String categoryRange, String seriesRange, String bubbleRange)
	 */
	public ChartSeriesHandle addSeriesRange( CellHandle legendCell, CellRange categoryRange, CellRange seriesRange, CellRange bubbleRange )
	{
		return this.addSeriesRange( legendCell, categoryRange, seriesRange, bubbleRange, 0 );    // 0=default chart
	}

	/**
	 * Adds a new Series to the chart via CellHandles and CellRange Objects.
	 * Each bar, line or wedge in a chart represents a Series.
	 * <br>This method can update the default chart (nChart==0) or overlay charts (nChart 1-9)
	 *
	 * @param CellHandle legendCell - references the legend cell for this series
	 * @param CellRange  categoryRange - The CellRange referencing the category (should be the same for all Series)
	 * @param CelLRange  seriesRange - The CellRange referencing the data points for one bar, line or wedge in the chart
	 * @param CellRange  bubbleRange -The CellRange referencing bubble sizes for this series, or null if chart is not of type BUBBLE
	 * @param nChart     number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return ChartSeriesHandle referencing the newly added series
	 * @see ChartHandle.addSeriesRange(String legendCell, String categoryRange, String seriesRange, String bubbleRange)
	 */
	public ChartSeriesHandle addSeriesRange( CellHandle legendCell,
	                                         CellRange categoryRange,
	                                         CellRange seriesRange,
	                                         CellRange bubbleRange,
	                                         int nChart )
	{
		Series s = null;
		if( bubbleRange == null )
		{
			s = mychart.addSeries( seriesRange.toString(),
			                       categoryRange.toString(),
			                       "",
			                       legendCell.getWorkSheetName() + "!" + legendCell.getCellAddress(),
			                       legendCell.getStringVal(),
			                       nChart );
		}
		else
		{
			s = mychart.addSeries( seriesRange.toString(),
			                       categoryRange.toString(),
			                       bubbleRange.toString(),
			                       legendCell.getWorkSheetName() + "!" + legendCell.getCellAddress(),
			                       legendCell.getStringVal(),
			                       nChart );
		}
		setDimensionsRecord();    // 20080417 KSC: update chart DIMENSIONS record upon update of series
		return new ChartSeriesHandle( s, wbh );
	}

	/**
	 * remove the Series (bar, line or wedge) at the desired index
	 *
	 * @param int index	- series index (valid values:  0 to getAllChartSeriesHandles().length-1)
	 * @see getAllChartSeriesHandles
	 */
	public void removeSeries( int index )
	{
		removeSeries( index, -1 );    // -1 flag for all series
	}

	/**
	 * remove the Series (bar, line or wedge) at the desired index
	 *
	 * @param int    index	- series index (valid values:  0 to getAllChartSeriesHandles().length-1)
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @see getAllChartSeriesHandles
	 */
	public void removeSeries( int index, int nChart )
	{
		Vector seriesperchart = mychart.getAllSeries( nChart );
		Series seriestodelete = (Series) seriesperchart.get( index );    // series to delete
		mychart.removeSeries( seriestodelete );
		setDimensionsRecord();
	}

	/**
	 * updates (replaces) every Chart Series (bar, line or wedge on the Chart) with the data from the array of values,
	 * legends, bubble sizes (optional) and category range.
	 * <br>
	 * NOTE: all arrays must be the same size (the exception is the bubleSizeRanges array, which may be null)
	 * <p/>
	 * NOTE: String arrays come in reverse order from plugins, so this method
	 * adds series LIFO i.e. reversed
	 *
	 * @param String[] valueRanges - Array of Cell Ranges representing the Values or Data points for each series (bar, line or wedge) on the Chart
	 * @param String[] legendCells - Array of Cell Addresses representing the legends for each Series
	 * @param String[] bubbleSizeRanges - Array of Cell ranges representing the bubble sizes for the Chart, or null if chart is not of type BUBBLE
	 * @param String   categoryRange - The Cell Range representing the categories (X Axis) for the entire Chart
	 */
	public void addAllSeries( String[] valueRanges, String[] legendCells, String[] bubbleSizeRanges, String categoryRange )
	{
		addAllSeries( valueRanges, legendCells, bubbleSizeRanges, categoryRange, 0 ); // do for default chart
	}

	/**
	 * updates (replaces) every Chart Series (bar, line or wedge on the Chart) with the data from the array of values,
	 * legends, bubble sizes (optional) and category range For the desired chart (0=default 1-9=overlay charts)
	 * <br>
	 * NOTE: all arrays must be the same size (the exception is the bubleSizeRanges array, which may be null)
	 * <p/>
	 * NOTE: String arrays come in reverse order from plugins, so this method
	 * adds series LIFO i.e. reversed
	 *
	 * @param String[] valueRanges - Array of Cell Ranges representing the Values or Data points for each series (bar, line or wedge) on the Chart
	 * @param String[] legendCells - Array of Cell Addresses representing the legends for each Series
	 * @param String[] bubbleSizeRanges - Array of Cell ranges representing the bubble sizes for the Chart, or null if chart is not of type BUBBLE
	 * @param String   categoryRange - The Cell Range representing the categories (X Axis) for the entire Chart
	 * @param nChart   number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 */
	private void addAllSeries( String[] valueRanges, String[] legendCells, String[] bubbleSizeRanges, String categoryRange, int nChart )
	{
		// first, remove all existing series
		Vector v = mychart.getAllSeries();
		for( int i = 0; i < v.size(); i++ )
		{
			mychart.removeSeries( (Series) v.get( i ) );
		}
		try
		{
			HashMap<String, Double> chartMetrics = mychart.getMetrics( wbh );    // build or retrieve  Chart Metrics --> dimensions + series data ... 
			this.mychart.getLegend().resetPos( chartMetrics.get( "y" ),
			                                   chartMetrics.get( "h" ),
			                                   chartMetrics.get( "canvash" ),
			                                   legendCells.length );
		}
		catch( Exception e )
		{
		}
		setDimensionsRecord();
		// now add series
		boolean hasBubbles = ((bubbleSizeRanges != null && bubbleSizeRanges.length == valueRanges.length));
		for( int i = valueRanges.length - 1; i >= 0; i-- )
		{
			try
			{
				if( !hasBubbles )    // usual case
				{
					this.addSeriesRange( legendCells[i], categoryRange, valueRanges[i], null, nChart );
				}
				else
				{
					this.addSeriesRange( legendCells[i], categoryRange, valueRanges[i], bubbleSizeRanges[i], nChart );
				}
			}
			catch( Exception e )
			{
				Logger.logErr( "Error adding series: " + e.toString() );
			}
		}
	}

	/**
	 * Appends a series one row below the last series in the chart.
	 * <p>This can be utilized when programmatically
	 * adding rows of data that should be reflected in the chart.
	 * <br>Legend cell will be incremented by one row if a reference. Category range
	 * will stay the same.
	 * <p/>
	 * <br>In order for this method to work properly the chart must have row-based series.  If your chart utilizes column-based
	 * series, then you need to append a category.
	 *
	 * @return ChartSeriesHandle representing newly added series
	 * @see ChartHandle.appendRowCategoryToChart
	 */
	public ChartSeriesHandle appendRowSeriesToChart()
	{
		return appendRowSeriesToChart( 0 );    // do for default chart (0)
	}

	/**
	 * Appends a series one row below the last series in the chart for the desired chart
	 * <p>This can be utilized when programmatically
	 * adding rows of data that should be reflected in the chart.
	 * <br>Legend cell will be incremented by one row if a reference. Category range
	 * will stay the same.
	 * <p/>
	 * <br>In order for this method to work properly the chart must have row-based series.  If your chart utilizes column-based
	 * series, then you need to append a category.
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return ChartSeriesHandle representing newly added series
	 * @see ChartHandle.appendRowCategoryToChart
	 */
	public ChartSeriesHandle appendRowSeriesToChart( int nChart )
	{
		ChartSeriesHandle[] handles = this.getAllChartSeriesHandles( nChart );
		ChartSeriesHandle theHandle = handles[handles.length - 1];
		String legendRef = theHandle.getSeriesLegendReference();
		if( legendRef != null && !legendRef.equals( "" ) )
		{
			String sheetnm = legendRef.substring( 0, legendRef.indexOf( "!" ) );
			legendRef = legendRef.substring( legendRef.indexOf( "!" ) + 1, legendRef.length() );
			int[] rc = ExcelTools.getRowColFromString( legendRef );
			rc[0] = rc[0] + 1;
			legendRef = sheetnm + "!" + ExcelTools.formatLocation( rc );
		}
		else if( legendRef == null )
		{
			legendRef = theHandle.getSeriesLegend();
		}
		else
		{
			legendRef = "";
		}
		String categoryRange = theHandle.getCategoryRange();
		String seriesRange = theHandle.getSeriesRange();
		String sheetnm = seriesRange.substring( 0, seriesRange.indexOf( "!" ) );
		seriesRange = seriesRange.substring( seriesRange.indexOf( "!" ) + 1, seriesRange.length() );
		int[] rc = ExcelTools.getRangeRowCol( seriesRange );
		// fiddle it, since exceltools doesn't translate back/forth
		int[] newRc = new int[4];
		newRc[0] = rc[1];
		newRc[1] = rc[0] + 1;
		newRc[2] = rc[3];
		newRc[3] = rc[2] + 1;
		seriesRange = sheetnm + "!" + ExcelTools.formatRange( newRc );
		try
		{
			return this.addSeriesRange( legendRef, "", categoryRange, seriesRange, "", nChart );
		}
		catch( CellNotFoundException e )
		{
			Logger.logErr( "ChartHandle.appendRowSeriesToChart: Unable to append series to chart: " + e );
		}
		return null;
	}

	/**
	 * Append a row of categories to the bottom of the chart.
	 * <br>Expands all Series to include the new bottom row.
	 * <br>To be utilized when expanding a chart to encompass more data that has a col-based series.
	 *
	 * @see ChartHandle.appendRowSeriesToChart
	 */
	public void appendRowCategoryToChart()
	{
		appendRowCategoryToChart( 0 );    // default chart
	}

	/**
	 * Append a row of categories to the bottom of the chart.
	 * <br>Expands all Series to include the new bottom row.
	 * <br>To be utilized when expanding a chart to encompass more data that has a col-based series.
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @see ChartHandle.appendRowSeriesToChart
	 */
	public void appendRowCategoryToChart( int nChart )
	{
		ChartSeriesHandle[] handles = this.getAllChartSeriesHandles( nChart );

		for( int i = 0; i < handles.length; i++ )
		{
			ChartSeriesHandle theHandle = handles[i];
			// update the series
			String seriesRange = theHandle.getSeriesRange();
			String[] s = ExcelTools.stripSheetNameFromRange( seriesRange );
			String sheetnm = s[0];
			seriesRange = s[1];
			//String sheetnm = seriesRange.substring(0, seriesRange.indexOf("!"));
			//seriesRange = seriesRange.substring( seriesRange.indexOf("!")+1,  seriesRange.length());
			// Strip 2nd sheet ref, if any 20080213 KSC
			//int n= seriesRange.indexOf('!');
			//int m= seriesRange.indexOf(':');
			//seriesRange= seriesRange.substring(0, m+1) + seriesRange.substring(n+1);

			int[] rc = ExcelTools.getRangeRowCol( seriesRange );
			int[] newRc = new int[4];
			newRc[0] = rc[1];
			newRc[1] = rc[0];
			newRc[2] = rc[3];
			newRc[3] = rc[2] + 1;
			seriesRange = sheetnm + "!" + ExcelTools.formatRange( newRc );
			theHandle.setSeriesRange( seriesRange );
			// update the category
			seriesRange = theHandle.getCategoryRange();
			s = ExcelTools.stripSheetNameFromRange( seriesRange );
			sheetnm = s[0];
			seriesRange = s[1];
		    /*
            sheetnm = seriesRange.substring(0, seriesRange.indexOf("!"));
            seriesRange = seriesRange.substring( seriesRange.indexOf("!")+1,  seriesRange.length());
            // Strip 2nd sheet ref, if any 20080213 KSC
			n= seriesRange.indexOf('!');
			m= seriesRange.indexOf(':');
			seriesRange= seriesRange.substring(0, m+1) + seriesRange.substring(n+1);
            */
			rc = ExcelTools.getRangeRowCol( seriesRange );
			newRc = new int[4];
			newRc[0] = rc[1];
			newRc[1] = rc[0];
			newRc[2] = rc[3];
			newRc[3] = rc[2] + 1;
			seriesRange = sheetnm + "!" + ExcelTools.formatRange( newRc );
			theHandle.setCategoryRange( seriesRange );
		}
	}
    
    /* NOT IMPLEMENTED YET
     * adjust chart cell references upon row insertion or deletion
     * NOTE: Assumes we're on the correct sheet 
	 * NOT COMPLETELY IMPLEMENTED YEt
     * @param rownum	
     * @param shiftamt	+1= insert row, -1= delete row
     *	
    public void adjustCellRefs(int rownum, int shiftamt) {
        Vector v = mychart.getAllSeries();
        boolean bSeriesRows= false;
        boolean bMod= false;
        for (int i=0;i<v.size();i++) {
            Series s = (Series)v.get(i);
            try { // SERIES
                Ai ai = s.getSeriesValueAi();                
                Ptg[] p=ai.getCellRangePtgs();	// should only be 1 ptg
            	try {
            		int[] loc= p[0].getIntLocation();
            		if (loc.length==4 && loc[0]==loc[2])
            			bSeriesRows= true;
            		if (shiftamt > 0) { // insert row i.e. shift ai location down
                		if ((loc.length==2 && loc[0]>=rownum) || (loc[0]>=rownum || loc[2]>=rownum)) {
                			adjustAiLocation(ai, p[0].getIntLocation(), shiftamt);
                			bMod= true;
                		}
            		} else {
            			if ((loc.length==2 && loc[0]==rownum) || (loc[0]>=rownum && loc[2]<=rownum)) {
                			// remove it
                			this.removeSeries(i);
                			bMod= true;
                			continue;
            			}
            		}
            	} catch(Exception e) {
            		
            	}
            } catch (Exception e) {
            	
            }
            try {	// CATEGORY
            	Ai ai= s.getCategoryValueAi();
                Ptg[] p=ai.getCellRangePtgs();
            	try {
            		int[] loc= p[0].getIntLocation();
            		if (shiftamt > 0) { // insert row i.e. shift ai location down
                		if ((loc.length==2 && loc[0]>=rownum) || (loc[0]>=rownum || loc[2]>=rownum)) {
                			adjustAiLocation(ai, p[0].getIntLocation(), shiftamt);	                			
                			bMod= true;
                		}
            		} else {
            			if ((loc.length==2 && loc[0]==rownum) || (loc[0]>=rownum && loc[2]<=rownum)) {
                			// remove it ????
                			this.removeSeries(i);
                			bMod= true;
            			}
            		}
            	} catch(Exception e) {
            		
            	}
            } catch (Exception e) {
            	
            }
            try {	// LEGEND
            	Ai ai= s.getLegendAi();
                Ptg[] p=ai.getCellRangePtgs();
                int[] loc= p[0].getIntLocation();
        		if (shiftamt > 0) { // insert row i.e. shift ai location down
            		if ((loc.length==2 && loc[0]>=rownum) || (loc[0]>=rownum || loc[2]>=rownum)) {
            			adjustAiLocation(ai, p[0].getIntLocation(), shiftamt);	                			
            			bMod= true;
            		}
        		} else {
        			if ((loc.length==2 && loc[0]==rownum) || (loc[0]>=rownum && loc[2]<=rownum)) {
            			// remove it
            			this.removeSeries(i);
            			bMod= true;
        			}
        		}
        	} catch(Exception e) {
        		
        	}
        }
        if (bMod)// one or more series elements were modified
			setDimensionsRecord();	// update dimensions i.e. Data Range
        	
    }
    */
	
    /* doesn't appear to be used right now
     * move the ai location (row) according to shift amount
     * @param ai
     * @param loc
     * @param shift
     *
    private void adjustAiLocation(Ai ai, int[] loc, int shift) {
    	String oldloc= ExcelTools.formatLocation(loc);
	    if (loc.length>2) // get 2nd part of range  
		       oldloc += ":" + ExcelTools.formatLocation(new int[]{loc[2], loc[3]});
	    oldloc= this.getSheet().getSheetName() + "!" + oldloc;	    
    	if (loc.length==2)// single cell
    		loc[0]+=shift;
    	else {	// range
    		loc[0]+=shift;
    		loc[2]+=shift;
    	}
	    String newloc= ExcelTools.formatLocation(loc);
	    if (loc.length>2) // get 2nd part of range  
	       newloc += ":" + ExcelTools.formatLocation(new int[]{loc[2], loc[3]});	          	
		ai.changeAiLocation(oldloc, newloc);
    }
    
*/

	/**
	 * Get the Chart's bytes
	 * <p/>
	 * This is an internal method that is not useful to the end user.
	 */
	public byte[] getChartBytes()
	{
		return mychart.getChartBytes();
	}

	public byte[] getSerialBytes()
	{
		return mychart.getSerialBytes();
	}

	/**
	 * get the chart-type-specific options in XML form
	 *
	 * @return String XML
	 */
	private String getChartOptionsXML()
	{
		return mychart.getChartOptionsXML( 0 ); // 0 for default chart
	}

	private String t( int n )
	{
		String tabs = "\t\t\t\t\t\t\t\t\t\t\t\t\t";
		return (tabs.substring( 0, n ));
	}

	/**
	 * given a XmlPullParser positioned at the chart element, parse all chart elements to create desired chart
	 *
	 * @param sht WorkSheetHandle
	 * @param xpp XmlPullParser
	 */
	public void parseXML( WorkSheetHandle sht, XmlPullParser xpp, HashMap maps )
	{
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "ChartFonts" ) )
					{
						for( int x = 0; x < xpp.getAttributeCount(); x++ )
						{
							this.setChartFont( xpp.getAttributeName( x ), xpp.getAttributeValue( x ) );
						}
					}
					else if( tnm.equals( "ChartFontRec" ) )
					{
						String fName = "";
						int fontId = 0, fSize = 0, fWeight = 0, fColor = 0, fUnderline = 0;
						boolean bIsBold = false;
						for( int x = 0; x < xpp.getAttributeCount(); x++ )
						{
							String attr = xpp.getAttributeName( x );
							String val = xpp.getAttributeValue( x );
							if( attr.equals( "name" ) )
							{
								fName = val;
							}
							else if( attr.equals( "id" ) )
							{
								fontId = Integer.parseInt( val );
							}
							else if( attr.equals( "size" ) )
							{
								fSize = Font.PointsToFontHeight( Double.parseDouble( val ) );
							}
							else if( attr.equals( "color" ) )
							{
								fColor = FormatHandle.HexStringToColorInt( val, FormatHandle.colorFONT );
								if( fColor == 0 )
								{
									fColor = 32767;    // necessary?
								}
							}
							else if( attr.equals( "weight" ) )
							{
								fWeight = Integer.parseInt( val );
							}
							else if( attr.equals( "bold" ) )
							{
								bIsBold = true;
							}
							else if( attr.equals( "underline" ) )
							{
								fUnderline = Integer.parseInt( val );
							}
						}
						while( this.getWorkBook().getNumFonts() < fontId - 1 )
						{
							this.getWorkBook().insertFont( new Font( "Arial", Font.PLAIN, 10 ) );
						}
						if( this.getWorkBook().getNumFonts() < fontId )
						{
							Font f = new Font( fName, fWeight, fSize );
							f.setColor( fColor );
							f.setBold( bIsBold );
							f.setUnderlineStyle( (byte) fUnderline );
							this.getWorkBook().insertFont( f );
						}
						else
						{    // TODO: this will screw up linked fonts, perhaps, so what to do?
							Font f = this.getWorkBook().getFont( fontId );
							f.setFontWeight( fWeight );
							f.setFontName( fName );
							f.setFontHeight( fSize );
							f.setColor( fColor );
							f.setBold( bIsBold );
							f.setUnderlineStyle( (byte) fUnderline );
						}
					}
					else if( tnm.equals( "FormatChartArea" ) )
					{    // TODO: something!
						// ChartBorder
						// ChartProperties
					}
					else if( tnm.equals( "Series" ) )
					{        // series -->
						// Legend Range Category shape typex typey
						String legend = "", series = "", category = "", bubble = "";
						String dataTypeX = "", dataTypeY = "";
						String shape = "";
						for( int x = 0; x < xpp.getAttributeCount(); x++ )
						{
							if( xpp.getAttributeName( x ).equalsIgnoreCase( "Legend" ) )
							{
								legend = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "Range" ) )
							{
								series = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "Category" ) )
							{
								category = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "Bubbles" ) )
							{
								bubble = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "TypeX" ) )
							{
								dataTypeX = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "TypeY" ) )
							{
								dataTypeY = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "Shape" ) )
							{
								shape = xpp.getAttributeValue( x );
							}
						}
						//20070709 KSC: can't add until all cells are added			        
						// ch.addSeriesRange(legend, category, series);
						String[] s = { legend, series, category, bubble, dataTypeX, dataTypeY, shape };
						HashMap map = (HashMap) maps.get( "chartseries" );
						map.put( s, this );
					}
					else if( tnm.equals( "ChartOptions" ) )
					{    // handle chart-type-specific options such as legend options
						// 20070716 KSC: handle chart-type-specific options in a very generic way ...
						for( int x = 0; x < xpp.getAttributeCount(); x++ )
						{
							this.setChartOption( xpp.getAttributeName( x ), xpp.getAttributeValue( x ) );
						}
					}
					else if( tnm.equals( "ThreeD" ) )
					{        // handle three-d options
						// handle threeD record options in a very generic way ...
						this.make3D(); // default chart - TODO; if mutliple charts, handle	
						for( int x = 0; x < xpp.getAttributeCount(); x++ )
						{    // now add threed rec options
							this.setChartOption( xpp.getAttributeName( x ), xpp.getAttributeValue( x ) );
						}
					}
					else if( tnm.endsWith( "Axis" ) )
					{        // handle axis specs (Label + options ...)
						// 20070720 KSC: handle Axis record options ...
						int type = 0;
						String axis = xpp.getName();
						if( axis.equalsIgnoreCase( "XAxis" ) )
						{
							type = XAXIS;
						}
						else if( axis.equalsIgnoreCase( "YAxis" ) )
						{
							type = YAXIS;
						}
						else if( axis.equalsIgnoreCase( "ZAxis" ) )
						{
							type = ZAXIS;
						}
						if( xpp.getAttributeCount() > 0 )
						{ // then has axis options
							for( int x = 0; x < xpp.getAttributeCount(); x++ )
							{
								this.setAxisOption( type, xpp.getAttributeName( x ), xpp.getAttributeValue( x ) );
							}
						}
						else
						{ // no axis options means no axis present; remove 
							this.removeAxis( type );
						}
					}
					else if( tnm.equals( "Series" ) )
					{        // handle series data
						// Legend Range Category
						String legend = "", series = "", category = "", bubble = "";
						String dataTypeX = "", dataTypeY = "";
						String shape = "";
						for( int x = 0; x < xpp.getAttributeCount(); x++ )
						{
							if( xpp.getAttributeName( x ).equalsIgnoreCase( "Legend" ) )
							{
								legend = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "Range" ) )
							{
								series = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "Category" ) )
							{
								category = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "Bubbles" ) )
							{
								bubble = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "TypeX" ) )
							{
								dataTypeX = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "TypeY" ) )
							{
								dataTypeY = xpp.getAttributeValue( x );
							}
							else if( xpp.getAttributeName( x ).equalsIgnoreCase( "Shape" ) )
							{
								shape = xpp.getAttributeValue( x );
							}
						}
						//20070709 KSC: can't add until all cells are added			        
						// ch.addSeriesRange(legend, category, series);
						String[] s = { legend, series, category, bubble, dataTypeX, dataTypeY, shape };
						HashMap map = (HashMap) maps.get( "chartseries" );
						map.put( s, this );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( "Chart" ) )
					{
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logWarn( "ChartHandle.parseXML <" + xpp.getName() + ">: " + e.toString() );
			// TODO: propogate Exception???
		}
	}

	/**
	 * returns an XML representation of this chart
	 *
	 * @return String XML
	 */
	public String getXML()
	{
		StringBuffer sb = new StringBuffer( t( 1 ) + "<Chart" );
		// Chart Name (=Title)
		sb.append( " Name=\"" + this.getTitle() + "\"" );
		// Type
		sb.append( " type=\"" + this.getChartType() + "\"" );
		// Plot Area Background color 20080429 KSC
		sb.append( " Fill=\"" + this.getPlotAreaBgColor() + "\"" );
		// Position
		short[] coords = mychart.getCoords();
		sb.append( " Left=\"" + coords[0] + "\" Top=\"" + coords[1] + "\" Width=\"" + coords[2] + "\" Height=\"" + coords[3] + "\"" );
		sb.append( ">\n" );

		// Chart Fonts
		sb.append( t( 2 ) + "<ChartFontRecs>" + this.getChartFontRecsXML() );
		sb.append( "\n" + t( 2 ) + "</ChartFontRecs>\n" );
		sb.append( t( 2 ) + "<ChartFonts" + this.getChartFontsXML() + "/>\n" );
		// Format Chart Area
		sb.append( t( 2 ) + "<FormatChartArea>\n" );
		sb.append( t( 3 ) + "<ChartBorder></ChartBorder>\n" ); // KSC: TODO: BORDER
		sb.append( t( 3 ) + "<ChartProperties></ChartProperties>\n" ); // TODO: Properties
		sb.append( t( 2 ) + "</FormatChartArea>\n" );
		// Source Data
		sb.append( t( 2 ) + "<SourceData>\n" );
		ChartSeriesHandle[] series = this.getAllChartSeriesHandles();
		for( int i = 0; i < series.length; i++ )
		{
			sb.append( t( 3 ) + "<Series Legend=\"" + series[i].getSeriesLegendReference() + "\"" );
			sb.append( " Range=\"" + series[i].getSeriesRange() + "\"" );
			sb.append( " Category=\"" + series[i].getCategoryRange() + "\"" );
			if( series[i].hasBubbleSizes() )
			{
				sb.append( " Bubbles=\"" + series[i].getBubbleSizes() + "\"" );
			}
			sb.append( " TypeX=\"" + series[i].getCategoryDataType() + "\"" );
			sb.append( " TypeY=\"" + series[i].getSeriesDataType() + "\"" );
			// controls shape of complex datapoints such as pyramid, cylinder, cone + stacked 3d bars  
			sb.append( " Shape=\"" + series[i].getShape() + "\"" );
			sb.append( "/>\n" );
		}
		sb.append( t( 2 ) + "</SourceData>\n" );
		// Chart Options
		sb.append( t( 2 ) + "<ChartOptions" );
		sb.append( this.getChartOptionsXML() );
		sb.append( "/>\n" );
		// Axis Options
		sb.append( t( 2 ) + "<Axes>\n" );
		sb.append( t( 3 ) + "<XAxis" + this.getAxisOptionsXML( XAXIS ) + "/>\n" );
		sb.append( t( 3 ) + "<YAxis" + this.getAxisOptionsXML( YAXIS ) + "/>\n" );
		sb.append( t( 3 ) + "<ZAxis" + this.getAxisOptionsXML( ZAXIS ) + "/>\n" );
		sb.append( t( 2 ) + "</Axes>\n" );
		// ThreeD rec opts
		if( this.isThreeD() )
		{
			sb.append( t( 2 ) + "<ThreeD" + this.getThreeDXML() + "/>\n" );
		}

		sb.append( t( 1 ) + "</Chart>\n" );
		return sb.toString();
	}

	/******************************************************************** OOXML Generation Methods **********************************************************************/
	/**
	 * Generates OOXML (chartML) for this chart object.
	 * <p/>
	 * <br>NOTE: necessary root chartSpace element + namespaces are not set here
	 *
	 * @param int rId -reference ID for this chart
	 * @return String representing the OOXML describing this Chart
	 */
	public String getOOXML( int rId )
	{
		// TODO: finish 3d options- floor, sideWall, backWall
		// TODO: finish axis options
		// TODO: printSettings

		// generate OOXML (chartML)
		StringBuffer cooxml = new StringBuffer();
		try
		{
			mychart.getChartSeries().resetSeriesNumber();    // reset series idx
			// retrieve pertinent chart data
			// axes id's TODO: HANDLE MULTIPLE AXES per chart ...
			String catAxisId = Integer.toString( (int) (Math.random() * 1000000) );
			String valAxisId = Integer.toString( (int) (Math.random() * 1000000) );
			String serAxisId = Integer.toString( (int) (Math.random() * 1000000) );
			OOXMLChart thischart;
			if( (mychart instanceof OOXMLChart) )
			{
				thischart = (OOXMLChart) mychart;
			}
			else
			{    // XLS->XLSX
				thischart = new OOXMLChart( mychart, wbh );
				mychart = thischart;
				thischart.getChartSeries().setParentChart( thischart );
			}
			thischart.wbh = this.wbh;

			cooxml.append( thischart.getOOXML( catAxisId, valAxisId, serAxisId ) );

			// TODO: <printSettings>
			ArrayList chartEmbeds = thischart.getChartEmbeds();
			if( chartEmbeds != null )
			{
				int j = 0;
				for( int i = 0; i < chartEmbeds.size(); i++ )
				{
					if( ((String[]) chartEmbeds.get( i ))[0].equals( "userShape" ) )
					{
						j++;
						cooxml.append( "<c:userShapes r:id=\"rId" + j + "\"/>" );
					}
				}
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "ChartHandle.getOOXML: error generating OOXML.  Chart not created: " + e.toString() );
		}
		return cooxml.toString();
	}

	/**
	 * generates the OOXML specific for DrawingML, specifying offsets and identifying
	 * the chart object.
	 * <br>this Drawing ML (OOXML) is distinct from Chart ML (OOXML) which actually defines the chart object including series, categories and axes
	 * <br>This is an internal method that is not useful to the end user.
	 *
	 * @param int id - the reference id for this chart
	 * @return String OOXML
	 */
	public String getChartDrawingOOXML( int id )
	{
		TwoCellAnchor t = new TwoCellAnchor( ((OOXMLChart) mychart).getEditMovement() );
		t.setAsChart( id,
		              OOXMLAdapter.stripNonAscii( this.getOOXMLName() ).toString(),
		              TwoCellAnchor.convertBoundsFromBIFF8( this.getSheet(),
		                                                    mychart.getBounds() ) );    // adjust BIFF8 bounds to OOXML units
		return t.getOOXML();
	}

	/******************************************************************************** Parsing OOXML Methods **********************************************************************/
	/**
	 * defines this chart object based on a Chart ML (OOXML) input Stream (root element=c:chartSpace)
	 * <br>This is an internal method that is not useful to the end user.
	 *
	 * @param inputStream ii - representing chart OOXML
	 */
	public void parseOOXML( java.io.InputStream ii )
	{
// overlay in title, after layout
// varyColors val= "0" -- after grouping and before ser    	

		// series colors by theme:
    	/*
    	 * accent1 - 6 
    	 */
/** chartSpace:
 chart (Chart) §5.7.2.27
 clrMapOvr (Color Map Override) §5.7.2.30
 date1904 (1904 Date System) §5.7.2.38
 externalData (External Data Relationship) §5.7.2.63
 extLst (Chart Extensibility) §5.7.2.64
 lang (Editing Language) §5.7.2.87
 pivotSource (Pivot Source) §5.7.2.145
 printSettings (Print Settings) §5.7.2.149
 protection (Protection) §5.7.2.150
 roundedCorners (Rounded Corners) §5.7.2.160
 spPr (Shape Properties) §5.7.2.198
 style (Style) §5.7.2.203
 txPr (Text Properties) §5.7.2.217
 userShapes (Reference to Chart Drawing Part) §5.7.2.222    	
 */
		try
		{
			OOXMLChart thischart = (OOXMLChart) mychart;
			int drawingOrder = 0;    // drawing order of the chart (0=default, 1-9 for multiple charts in 1)
			boolean hasPivotTableSource = false;

			// remove any undesired formatting from default chart:
			this.setTitle( "" );    // clear any previously set 
			mychart.getAxes().setPlotAreaBgColor( FormatConstants.COLOR_WHITE );
			mychart.getAxes().setPlotAreaBorder( -1, -1 );    // remove plot area border 

			java.util.Stack<String> lastTag = new java.util.Stack();        // keep track of element hierarchy

			XmlPullParserFactory factory = XmlPullParserFactory.newInstance();
			factory.setNamespaceAware( true );
			XmlPullParser xpp = factory.newPullParser();

			xpp.setInput( ii, null );    // using XML 1.0 specification 
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName(); // main entry= chartSpace, children: lang, chart
					lastTag.push( tnm );                // keep track of element hierarchy
					if( tnm.equals( "chart" ) )
					{   // beginning of DrawingML for a single image or chart; children:  title, plotArea, legend,
/**
 * <sequence>
 5 <element name="title" type="CT_Title" minOccurs="0" maxOccurs="1"/>
 6 <element name="autoTitleDeleted" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
 7 <element name="pivotFmts" type="CT_PivotFmts" minOccurs="0" maxOccurs="1"/>
 8 <element name="view3D" type="CT_View3D" minOccurs="0" maxOccurs="1"/>
 9 <element name="floor" type="CT_Surface" minOccurs="0" maxOccurs="1"/>
 10 <element name="sideWall" type="CT_Surface" minOccurs="0" maxOccurs="1"/>
 11 <element name="backWall" type="CT_Surface" minOccurs="0" maxOccurs="1"/>
 12 <element name="plotArea" type="CT_PlotArea" minOccurs="1" maxOccurs="1"/>
 plotArea contains layout, <ChartType>, serAx, valAx, catAx, dateAx, spPr,  dTable
 13 <element name="legend" type="CT_Legend" minOccurs="0" maxOccurs="1"/>
 14 <element name="plotVisOnly" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
 15 <element name="dispBlanksAs" type="CT_DispBlanksAs" minOccurs="0" maxOccurs="1"/>
 16 <element name="showDLblsOverMax" type="CT_Boolean" minOccurs="0" maxOccurs="1"/>
 17 <element name="extLst" type="CT_ExtensionList" minOccurs="0" maxOccurs="1"/>
 18 </sequence>		            	 
 */
					}
					else if( tnm.equals( "lang" ) )
					{
						thischart.lang = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "roundedCorners" ) )
					{
						thischart.roundedCorners = xpp.getAttributeValue( 0 ).equals( "1" );
					}
					else if( tnm.equals( "pivotSource" ) )
					{    // has a pivot table
						hasPivotTableSource = true;
					}
					else if( tnm.equals( "view3D" ) )
					{
						ThreeD.parseOOXML( xpp, lastTag, thischart );
					}
					else if( tnm.equals( "layout" ) )
					{
						thischart.plotAreaLayout = (Layout) Layout.parseOOXML( xpp, lastTag ).cloneElement();
					}
					else if( tnm.equals( "legend" ) )
					{
						thischart.showLegend( true, false );
						thischart.ooxmlLegend = (com.extentech.formats.OOXML.Legend) com.extentech.formats.OOXML.Legend.parseOOXML( xpp,
						                                                                                                            lastTag,
						                                                                                                            this.wbh )
						                                                                                               .cloneElement();
						thischart.ooxmlLegend.fill2003Legend( thischart.getLegend() );
						// Parse actual CHART TYPE element (barChart, pieChart, etc.)
					}
					else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.BARCHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.LINECHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.PIECHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.AREACHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.SCATTERCHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.RADARCHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.SURFACECHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.DOUGHNUTCHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.BUBBLECHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.OFPIECHART] ) ||
							tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.STOCKCHART] ) ||
							tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.BARCHART] ) ||
							tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.LINECHART] ) ||
							tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.PIECHART] ) ||
							tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.AREACHART] ) ||
							tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.SCATTERCHART] ) ||
							tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.SURFACECHART] ) )
					{  // specific chart type-
						ChartType.parseOOXML( xpp, this.wbh, mychart, drawingOrder++ );
						lastTag.pop();
					}
					else if( tnm.equals( "title" ) )
					{
						thischart.setOOXMLTitle( (Title) Title.parseOOXML( xpp, lastTag, this.wbh ).cloneElement(), this.wbh );
						this.setTitle( thischart.getOOXMLTitle().getTitle() );
					}
					else if( tnm.equals( "spPr" ) )
					{ // shape properties -- can be for plot area or chart space 
						String parent = (String) lastTag.get( lastTag.size() - 2 );
						if( parent.equals( "plotArea" ) )
						{
							thischart.setSpPr( 0, (SpPr) SpPr.parseOOXML( xpp, lastTag, this.wbh ).cloneElement() );
						}
						else if( parent.equals( "chartSpace" ) )
						{
							thischart.setSpPr( 1, (SpPr) SpPr.parseOOXML( xpp, lastTag, this.wbh ).cloneElement() );
						}
					}
					else if( tnm.equals( "txPr" ) )
					{        // text formatting
						thischart.setTxPr( (TxPr) TxPr.parseOOXML( xpp, lastTag, this.wbh ).cloneElement() );
					}
					else if( tnm.equals( "catAx" ) )
					{        // child of plotArea
						mychart.getAxes().parseOOXML( XAXIS, xpp, tnm, lastTag, this.wbh );
					}
					else if( tnm.equals( "valAx" ) )
					{        // child of plotArea
						if( mychart.getAxes().hasAxis( Axis.XAXIS ) )    // usual, have a catAx then a valAx
						{
							mychart.getAxes().parseOOXML( Axis.YAXIS, xpp, tnm, lastTag, this.wbh );
						}
						else if( mychart.getAxes().hasAxis( Axis.YAXIS ) )    // for bubble charts, has two valAxes and no catAx
						{
							mychart.getAxes().parseOOXML( Axis.XVALAXIS, xpp, tnm, lastTag, this.wbh );
						}
						else        // 2nd val axis is Y axis
						{
							mychart.getAxes().parseOOXML( Axis.YAXIS, xpp, tnm, lastTag, this.wbh );
						}
					}
					else if( tnm.equals( "serAx" ) )
					{        // series axis - 3d charts 
						mychart.getAxes().parseOOXML( ZAXIS, xpp, tnm, lastTag, this.wbh );
					}
					else if( tnm.equals( "dateAx" ) )
					{        // TODO: not finished: figure out!
						// ??
					}

				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					lastTag.pop();
					String endTag = xpp.getName();
					if( endTag.equals( "chartSpace" ) )
					{
						setDimensionsRecord();
						break;    // done processing
					}
				}
				if( xpp.getEventType() != XmlPullParser.END_DOCUMENT )
				{
					eventType = xpp.next();
				}
				else
				{
					eventType = XmlPullParser.END_DOCUMENT;
				}
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "ChartHandle.parseChartOOXML: " + e.toString() );
		}
	}

	/**
	 * Specifies how to resize or move this Chart upon edit
	 * <br>This is an internal method that is not useful to the end user.
	 * <br>Excel 7/OOXML specific
	 *
	 * @param editMovement String OOXML-specific edit movement setting
	 */
	public void setEditMovement( String editMovement )
	{
		((OOXMLChart) mychart).setEditMovement( editMovement );
	}

	/**
	 * return the Excel 7/OOXML-specific name for this chart
	 *
	 * @return String OOXML name
	 */
	private String getOOXMLName()
	{
		return ((OOXMLChart) mychart).getOOXMLName();
	}

	/**
	 * set the Excel 7/OOXML-specific name for this chart
	 *
	 * @param String name
	 */
	public void setOOXMLName( String name )
	{
		((OOXMLChart) mychart).setOOXMLName( name );
	}

	/**
	 * returns the drawingml file name which defines the userShape (if any)
	 * <br>a userShape is a drawing or shape ontop of a chart
	 * associated with this chart
	 *
	 * @return
	 */
	public ArrayList getChartEmbeds()
	{
		return ((OOXMLChart) mychart).getChartEmbeds();
	}

	/**
	 * sets external information linked to or "embedded" in this OOXML chart;
	 * can be a chart user shape, an image ...
	 * <br>NOTE: a userShape is a drawingml file name which defines the userShape (if any)
	 * <br>a userShape is a drawing or shape ontop of a chart
	 *
	 * @param String[] embedType, filename e.g. {"userShape", "userShape file name"}
	 */
	public void addChartEmbed( String[] ce )
	{
		((OOXMLChart) mychart).addChartEmbed( ce );
	}

	/**
	 * set the chart DIMENSIONS record based on the series ranges in the chart
	 * APPEARS THAT for charts, the DIMENSIONS record merely notes the range
	 * of values:
	 * 0, #points in series, 0, #series
	 */
	protected void setDimensionsRecord()
	{
		ChartSeriesHandle[] series = this.getAllChartSeriesHandles();
		int nSeries = series.length;
		int nPoints = 0;
		for( int i = 0; i < series.length; i++ )
		{
			try
			{
				int[] coords = ExcelTools.getRangeCoords( series[i].getSeriesRange() );
				if( coords[3] > coords[1] )
				{
					nPoints = Math.max( nPoints, coords[3] - coords[1] + 1 );    // c1-c0
				}
				else
				{
					nPoints = Math.max( nPoints, coords[2] - coords[0] + 1 );    // r1-r0
				}
			}
			catch( Exception e )
			{
			}
		}
		mychart.setDimensionsRecord( 0, nPoints, 0, nSeries );
	}

	/**
	 * Method for setting Chart-Type-specific options in a generic fashion e.g. charthandle.setChartOption("Stacked", "true");
	 * <p>Note: since most Chart Type Options are interdependent, there are several makeXX methods
	 * that set the desired group of options e.g. makeStacked(); use setChartOption with care
	 * <p>Note that not all Chart Types will have every option available
	 * <p>Possible Options:
	 * <p>"Stacked" - true or false - set Chart Series to be Stacked
	 * <br>"Cluster" - true or false -  set Clustered for Column and Bar Chart Types
	 * <br>"PercentageDisplay" - true or false - Each Category is broken down as a percentge
	 * <br>"Percentage" - Distance of pie slice from center of pie as % for Pie Charts (0 for all others)
	 * <br>"donutSize" - Donut size for Donut Charts Only
	 * <br>"Overlap" - Space between bars (default= 0%)
	 * <br>"Gap" - Space between categories (%) (default=50%)
	 * <br>"SmoothedLine" - true or false - the Line series has a smoothed line
	 * <br>"AnRot" - Rotation Angle (0 to 360 degrees), usually 0 for pie, 20 for others (3D option)
	 * <br>"AnElev" - Elevation Angle (-90 to 90 degrees)   (15 is default) (3D option)
	 * <br>"ThreeDScaling" - true or false - 3d effect
	 * <br>"TwoDWalls" - true if 2D walls (3D option)
	 * <br>"PcDist" - Distance from eye to chart (0 to 100) (30 is default) (3D option)
	 * <br>"ThreeDBubbles" - true or false - Draw bubbles with a 3d effect
	 * <br>"ShowLdrLines" - true or false - Show Pie and Donut charts Leader Lines
	 * <br>"MarkerFormat" - "0" thru "9" for various marker options @see ChartHandle.setMarkerFormat
	 * <br>"ShowLabel"		- true or false - show Series/Data Label
	 * <br>"ShowCatLabel"	- true or false - show Category Label
	 * <br>"ShowLabelPct"	- true or false - show percentage labels for Pie charts
	 * <br>"ShowBubbleSizes"	- true or false - show bubble sizes for Bubble charts
	 * <p>NOTE: all values must be in String form
	 *
	 * @see ChartHandle.getXML
	 */
	public void setChartOption( String op, String val )
	{
		mychart.setChartOption( op, val );
	}

	/**
	 * Method for setting Chart-Type-specific options in a generic fashion
	 * e.g. charthandle.setChartOption("Stacked", "true");
	 *
	 * @param op
	 * @param val
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 */
	private void setChartOption( String op, String val, int nChart )
	{
		mychart.setChartOption( op, val, nChart );
	}

	/**
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return true if Chart has 3D effects, false otherwise
	 */
	private boolean isThreeD( int nChart )
	{
		return mychart.isThreeD( nChart );
	}

	/**
	 * @return true if Chart has 3D effects, false otherwise
	 */
	public boolean isThreeD()
	{
		return mychart.isThreeD( 0 );    // default chart
	}

	/**
	 * @return boolean true if Chart contains Stacked Series, false otherwise
	 */
	public boolean isStacked()
	{
		return mychart.isStacked( 0 );    // default chart
	}

	/**
	 * @return boolean true if Chart is of type 100% Stacked, false otherwise
	 */
	public boolean is100PercentStacked()
	{
		return mychart.is100PercentStacked( 0 );    // default chart
	}

	/**
	 * @return boolean true if Chart contains Clustered Bars or Columns, false otherwise
	 */
	public boolean isClustered()
	{
		return mychart.isClustered( 0 );    // default chart
	}

	/**
	 * @return String ThreeD options in XML form
	 */
	public String getThreeDXML()
	{
		return mychart.getThreeDXML( 0 );    // 0 for default chart
	}

	/**
	 * Make chart 3D if not already
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @return ThreeD rec
	 */
	private ThreeD initThreeD( int nChart )
	{
		return mychart.initThreeD( nChart, this.getChartType( nChart ) );
	}

	/**
	 * @param int Axis - one of the Axis constants (XAXIS, YAXIS or ZAXIS)
	 * @return String XML representation of the desired Axis
	 */
	private String getAxisOptionsXML( int Axis )
	{
		return mychart.getAxes().getAxisOptionsXML( Axis );
	}

	/**
	 * Returns the Axis Label Placement or position as an int
	 * <p>One of:
	 * <br>Axis.INVISIBLE - axis is hidden
	 * <br>Axis.LOW - low end of plot area
	 * <br>Axis.HIGH - high end of plot area
	 * <br>Axis.NEXTTO- next to axis (default)
	 *
	 * @param int Axis - one of the Axis constants (XAXIS, YAXIS or ZAXIS)
	 * @return int - one of the Axis Label placement constants above
	 */
	public int getAxisPlacement( int Axis )
	{
		return mychart.getAxes().getAxisPlacement( Axis );
	}

	/**
	 * Sets the Axis labels position or placement to the desired value (these match Excel placement options)
	 * <p>Possible options:
	 * <br>Axis.INVISIBLE - hides the axis
	 * <br>Axis.LOW - low end of plot area
	 * <br>Axis.HIGH - high end of plot area
	 * <br>Axis.NEXTTO- next to axis (default)
	 *
	 * @param int       Axis - one of the Axis constants (XAXIS, YAXIS or ZAXIS)
	 * @param Placement - int one of the Axis placement constants listed above
	 */
	public void setAxisPlacement( int Axis, int Placement )
	{
		mychart.getAxes().setAxisPlacement( Axis, Placement );
		mychart.setDirtyFlag( true );
	}
	
	/*
	 * returns the desired axis 
	 * If bCreateIfNecessary, will creates if doesn't exist 
	 * @return Axis Object
	 * /
	private Axis getAxis(int axisType, boolean bCreateIfNecessary) {
		return mychart.getAxis(axisType, bCreateIfNecessary);
	}*/

	/**
	 * removes the desired Axis from the Chart
	 *
	 * @param int axisType - one of the Axis constants (XAXIS, YAXIS or ZAXIS)
	 */
	public void removeAxis( int axisType )
	{
		mychart.getAxes().removeAxis( axisType );
		mychart.setDirtyFlag( true );
	}

	/**
	 * returns Chart-specific Font Records in XML form
	 *
	 * @return String Chart Font information in XML format
	 */
	public String getChartFontRecsXML()
	{
		return mychart.getChartFontRecsXML();
	}

	/**
	 * Return non-axis Chart font ids in XML form
	 *
	 * @return String Font information in XML format
	 */
	public String getChartFontsXML()
	{
		return mychart.getChartFontsXML();
	}

	/**
	 * Set non-axis chart font id for title, default, etc
	 * <br>For Internal Use Only
	 *
	 * @param String type - font type
	 * @param String val - font id
	 */
	public void setChartFont( String type, String val )
	{
		mychart.setChartFont( type, val );
	}

	/**
	 * @return the WorkBook Object attached to this Chart
	 */
	public com.extentech.formats.XLS.WorkBook getWorkBook()
	{
		return mychart.getWorkBook();
	}

	public WorkBookHandle getWorkBookHandle()
	{
		return this.wbh;
	}

	public WorkSheetHandle getWorkSheetHandle()
	{
		try
		{
			return this.wbh.getWorkSheet( mychart.getSheet().getSheetNum() );
		}
		catch( WorkSheetNotFoundException e )
		{
			// this should be impossible
			throw new RuntimeException( e );
		}
	}

	/**
	 * sets the coordinates or bounds (position, width and height) of this chart in pixels
	 *
	 * @param short[4] bounds - left or x value, top or y value, width, height
	 * @see ChartHandle.X, ChartHandle.Y, ChartHandle.WIDTH, ChartHandle.HEIGHT
	 */
	//NOTE: THIS SHOULD BE RENAMED TO setCoords as Bounds and Coords are very distinct
	public void setBounds( short[] bounds )
	{
		mychart.setCoords( bounds );
	}

	/**
	 * returns the coordinates or bounds (position, width and height) of this chart in pixels
	 *
	 * @return short[4] bounds - left or x value, top or y value,  width,  height
	 * @see ChartHandle.X, ChartHandle.Y, ChartHandle.WIDTH, ChartHandle.HEIGHT
	 */
	//NOTE: THIS SHOULD BE RENAMED TO setCoords as Bounds and Coords are very distinct
	public short[] getBounds()
	{
		return mychart.getCoords();
	}

	/**
	 * sets the coordinates (position, width and height) for this chart in Excel size units
	 *
	 * @return short[4] pixel coords - left or x value, top or y value,  width,  height
	 * @see ChartHandle.X, ChartHandle.Y, ChartHandle.WIDTH, ChartHandle.HEIGHT
	 */
	public void setCoords( short[] coords )
	{
		mychart.setCoords( coords );
	}

	/**
	 * returns the coordinates (position, width and height) of this chart in Excel size units
	 *
	 * @return short[4] pixel coords - left or x value, top or y value, width, height
	 * @see ChartHandle.X, ChartHandle.Y, ChartHandle.WIDTH, ChartHandle.HEIGHT
	 */
	public short[] getCoords()
	{
		mychart.getMetrics( this.wbh );
		return mychart.getCoords();
	}

	/**
	 * get the bounds of the chart using coordinates relative to row/cols and their offsets
	 *
	 * @return short[8] bounds - COL, COLOFFSET, ROW, ROWOFFST, COL1, COLOFFSET1, ROW1, ROWOFFSET1
	 */
	public short[] getRelativeBounds()
	{
		return mychart.getBounds();
	}

	/**
	 * returns the offset within the column in pixels
	 *
	 * @return
	 */
	public short getColOffset()
	{
		return mychart.getColOffset();
	}

	/**
	 * sets the bounds of the chart using coordinates relative to row/cols and their offsets
	 *
	 * @param short[8] bounds - COL, COLOFFSET, ROW, ROWOFFST, COL1, COLOFFSET1, ROW1, ROWOFFSET1
	 */
	public void setRelativeBounds( short[] bounds )
	{
		mychart.setBounds( bounds );
	}

	// 20070802 KSC: Debug uility to write out chartrecs
	/*
	 * For internal debugging use only
	 *
	public void writeChartRecs(String fName) {
		mychart.writeChartRecs(fName);
	}
	
	/**
	 * set DataLabels option for this chart
	 * @param String type - see below  
	 * @param boolean bShowLegendKey - true if show legend key, false otherwise
	 * <p>possible String type values:
	 * <br>Series
	 * <br>Category
	 * <br>Value 
	 * <br>Percentage	(Only for Pie Charts)
	 * <br>Bubble		(Only for Bubble Charts)
	 * <br>X Value		(Only for Bubble Charts)
	 * <br>Y Value		(Only for Bubble Charts)
	 * <br>CandP
	 * 
	 * <br><br>NOTE: not 100% implemented at this time
	 */
	public void setDataLabel( String/*[] */type, boolean bShowLegendKey )
	{
/* for now, only 1 option is valid 
 * - multiple legend settings e.g. Category and Value are not figured out yet 		
 * 		for (int i= 0; i < type.length; i++)
			mychart.setChartOption("DataLabel", type[i]);
*/
		if( !bShowLegendKey )
		{
			mychart.setChartOption( "DataLabel", type );
		}
		else
		{
			mychart.setChartOption( "DataLabelWithLegendKey", type );
		}
	}

	/**
	 * shows or removes the Data Table for this chart
	 *
	 * @param boolean bShow - true if show data table
	 */
	public void showDataTable( boolean bShow )
	{
		mychart.showDataTable( bShow );
	}

	/**
	 * shows or hides the Chart legend key
	 *
	 * @param booean  bShow - true if show legend, false to hide
	 * @param boolean vertical	- true if show vertically, false for horizontal
	 */
	public void showLegend( boolean bShow, boolean vertical )
	{
		mychart.showLegend( bShow, vertical );
	}

	/**
	 * returns true if Chart has a Data Legend Key showing
	 *
	 * @return true if Chart has a Data Legend Key showing
	 */
	public boolean hasDataLegend()
	{
		return mychart.hasDataLegend();
	}

	public void removeLegend()
	{
		mychart.removeLegend();
	}

	// 20070905 KSC: Group chart options for ease of setting
	// almost all charts have these specific ChartTypes:

	/**
	 * makes this Chart Stacked
	 * <br>sets the group of options necessary to create a stacked chart
	 * <br>For Chart Types:
	 * <br>BAR, COL, LINE, AREA, PYRAMID, PYRAMIDBAR, CYLINDER, CYLINDERBAR, CONE, CONEBAR
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @deprecated
	 */
	public void makeStacked( int nChart )
	{    // bar, col, line, area, pyramid col + bar, cone col + bar, cylinder col + bar
		int chartType = this.getChartType( nChart );
		this.setChartOption( "Stacked", "true", nChart );
		switch( chartType )
		{
			case ChartConstants.BARCHART:
			case ChartConstants.COLCHART:
				this.setChartOption( "Overlap", "-100", nChart );
				break;
			case ChartConstants.CYLINDERCHART:
			case ChartConstants.CYLINDERBARCHART:
			case ChartConstants.CONECHART:
			case ChartConstants.CONEBARCHART:
			case ChartConstants.PYRAMIDCHART:
			case ChartConstants.PYRAMIDBARCHART:
				this.setChartOption( "Overlap", "-100", nChart );
				ThreeD td = this.initThreeD( nChart );
				td.setChartOption( "Cluster", "false" );
				break;
			case ChartConstants.LINECHART:
				this.setChartOption( "Percentage", "0", nChart );
				break;
			case ChartConstants.AREACHART:
				this.setChartOption( "Overlap", "-100", nChart );
				this.setChartOption( "Percentage", "25", nChart );
				this.setChartOption( "SmoothedLine", "true", nChart );
				break;
		}
	}

	/**
	 * makes this Chart 100% Stacked
	 * <br>sets the group of options necessary to create a 100% stacked chart
	 * <br>For Chart Types:
	 * <br>BAR, COL, LINE, PYRAMID, PYRAMIDBAR, CYLINDER, CYLINDERBAR, CONE, CONEBAR
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @deprecated
	 */
	public void make100PercentStacked( int nChart )
	{    // bar, col, line, pyramid col + bar, cone col + bar, cylinder col + bar
		int chartType = this.getChartType( nChart );
		this.setChartOption( "Stacked", "true", nChart );
		this.setChartOption( "PercentageDisplay", "true", nChart );
		switch( chartType )
		{
			case ChartConstants.COLCHART:    // + pyramid
			case ChartConstants.BARCHART:
				this.setChartOption( "Overlap", "-100", nChart );
				break;
			case ChartConstants.CYLINDERCHART:
			case ChartConstants.CYLINDERBARCHART:
			case ChartConstants.CONECHART:
			case ChartConstants.CONEBARCHART:
			case ChartConstants.PYRAMIDCHART:
			case ChartConstants.PYRAMIDBARCHART:
				this.setChartOption( "Overlap", "-100", nChart );
				ThreeD td = this.initThreeD( nChart );
				td.setChartOption( "Cluster", "false" );
				break;
			case ChartConstants.LINECHART:
				break;
		}
	}

	/**
	 * makes this Chart Stacked with a 3D Effect
	 * <br>sets the group of options necessary to create a Stacked 3D chart
	 * <br>For Chart Types:
	 * <br>BAR, COL, AREA
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 */
	public void makeStacked3D( int nChart )
	{    // bar, col, area
		int chartType = this.getChartType( nChart );
		this.setChartOption( "Stacked", "true", nChart );
		ThreeD td = this.initThreeD( nChart );
		td.setChartOption( "AnRot", "20" );
		td.setChartOption( "ThreeDScaling", "true" );
		td.setChartOption( "TwoDWalls", "true" );
		switch( chartType )
		{
			case ChartConstants.COLCHART:
			case ChartConstants.BARCHART:
				this.setChartOption( "Overlap", "-100", nChart );
				break;
			case ChartConstants.AREACHART:
				this.setChartOption( "Percentage", "25", nChart );
				this.setChartOption( "SmoothedLine", "true", nChart );
				break;
		}
	}

	/**
	 * makes this Chart 100% Stacked with a 3D Effect
	 * <br>sets the group of options necessary to create a 100% Stacked 3D chart
	 * <br>For Chart Types:
	 * <br>BAR, COL, AREA
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @deprecated
	 */
	public void make100PercentStacked3D( int nChart )
	{    // bar, col, area
		int chartType = this.getChartType( nChart );
		this.setChartOption( "Stacked", "true", nChart );
		this.setChartOption( "PercentageDisplay", "true", nChart );
		switch( chartType )
		{
			case ChartConstants.COLCHART:    // + pyramid
			case ChartConstants.BARCHART:
				this.setChartOption( "Overlap", "-100", nChart );
				ThreeD td = this.initThreeD( nChart );
				td.setChartOption( "AnRot", "20" );
				td.setChartOption( "ThreeDScaling", "true" );
				td.setChartOption( "TwoDWalls", "true" );
				break;
			case ChartConstants.AREACHART:
				this.setChartOption( "Percentage", "25", nChart );
				this.setChartOption( "SmoothedLine", "true", nChart );
				break;
		}
	}

	/**
	 * makes the default Chart hava a 3D effect
	 * <br>sets the group of options necessary to create a 3D chart
	 * <br>For Chart Types:
	 * <br>BARCHART, COLCHART, LINECHART, PIECHART, AREACHART, BUBBLECHART, PYRAMIDCHART, CYLINDERCHART, CONECHART
	 *
	 * @deprecated use setChartType(chartType, 0, is3d, isStacked, is100PercentStacked) instead
	 */
	public void make3D()
	{    // bar, col, line, pie, area, bubble, pyramid, cone, cylinder
		make3D( 0 );
	}

	/**
	 * makes the desired Chart hava a 3D effect
	 * <br>where nChart 0= default, 1-9=multiple charts in one
	 * <br>sets the group of options necessary to create a 3D chart
	 * <br>For Chart Types:
	 * <br>BARCHART, COLCHART, LINECHART, PIECHART, AREACHART, BUBBLECHART, PYRAMIDCHART, CYLINDERCHART, CONECHART
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @deprecated use setChartType(chartType, nChart, is3d, isStacked, is100%Stacked) instead
	 */
	public void make3D( int nChart )
	{    // bar, col, line, pie, area, bubble, pyramid, cone, cylinder
		int chartType = this.getChartType( nChart );
		ThreeD td = null;
		switch( chartType )
		{
			case ChartConstants.COLCHART:
			case ChartConstants.BARCHART:
				td = this.initThreeD( nChart );
				td.setChartOption( "AnRot", "20" );
				td.setChartOption( "ThreeDScaling", "true" );
				td.setChartOption( "TwoDWalls", "true" );
				break;
			case ChartConstants.CYLINDERCHART:
			case ChartConstants.CONECHART:
			case ChartConstants.PYRAMIDCHART:
				td = this.initThreeD( nChart );
				td.setChartOption( "Cluster", "false" );
				break;
			case ChartConstants.AREACHART:
				this.setChartOption( "Percentage", "25", nChart );
				this.setChartOption( "SmoothedLine", "true", nChart );
				td = this.initThreeD( nChart );
				td.setChartOption( "AnRot", "20" );
				td.setChartOption( "ThreeDScaling", "true" );
				td.setChartOption( "TwoDWalls", "true" );
				td.setChartOption( "Perspective", "true" );
				break;
			case ChartConstants.PIECHART:
			case ChartConstants.LINECHART:
				this.initThreeD( nChart );    // just create a threeD rec w/ no extra options
				break;
			case ChartConstants.BUBBLECHART:
				this.setChartOption( "Percentage", "25", nChart );
				this.setChartOption( "SmoothedLine", "true", nChart );
				this.setChartOption( "ThreeDBubbles", "true", nChart );
				td = this.initThreeD( nChart );    // 20081228 KSC
				break;
		}
	}

	// more specialized option sets

	/**
	 * makes this Chart Clusted with a 3D effect
	 * <br>sets the group of options necessary to create a Clusted 3D chart
	 * <br>For Chart Types:
	 * <br>BAR, COL
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @deprecated
	 */
	public void makeClustered3D( int nChart )
	{    // only for Column and Bar (?)
		int chartType = this.getChartType( nChart );
		switch( chartType )
		{
			case ChartConstants.BARCHART:
			case ChartConstants.COLCHART:
				ThreeD td = this.initThreeD( nChart );
				td.setChartOption( "AnRot", "20" );
				td.setChartOption( "Cluster", "true" );
				td.setChartOption( "ThreeDScaling", "true" );
				td.setChartOption( "TwoDWalls", "true" );
				break;
		}
	}

	/**
	 * makes this chart's wedges exploded i.e. separated
	 * <br>For Chart Types:
	 * <br>PIECHART, DOUGHNUTCHART
	 *
	 * @deprecated
	 */
	public void makeExploded()
	{    // pie, donut
		int chartType = this.getChartType();
		switch( chartType )
		{
			case ChartConstants.DOUGHNUTCHART:
				this.setChartOption( "SmoothedLine", "true" );
			case ChartConstants.PIECHART:
				this.setChartOption( "ShowLdrLines", "true" );
				this.setChartOption( "Percentage", "25" );
				break;
			//ShowLdrLines="true" Percentage="25"/>
			// exploded donut:  ShowLdrLines="true" Donut="50" Percentage="25" SmoothedLine="true"/>
		}
	}

	/**
	 * makes this chart's wedges exploded 3D i.e. separated with a 3D effect
	 * <br>For Chart Types:
	 * <br>PIECHART, DOUGHNUTCHART
	 *
	 * @param nChart number and drawing order of the desired chart (default= 0 max=9 where 1-9 indicate an overlay chart)
	 * @deprecated
	 */
	public void makeExploded3D( int nChart )
	{ // pie
		// ShowLdrLines="true" Percentage="25"
		//AnRot="236"
		int chartType = this.getChartType( nChart );
		switch( chartType )
		{
			case ChartConstants.DOUGHNUTCHART:
			case ChartConstants.PIECHART:
				this.setChartOption( "ShowLdrLines", "true", nChart );
				this.setChartOption( "Percentage", "25", nChart );
				ThreeD td = this.initThreeD( nChart );
				td.setChartOption( "AnRot", "236" );
				break;
		}
	}
	
	/*
	 * NOT IMPLEMENTED YET TODO: IMPLEMENT
	 * Make this Chart have smoothed lines (Scatter only)
	 *
	public void makeSmoothedLines() { // scatter
		//Percentage="25" SmoothedLine="true
	}

	public void makeWireFrame() { // surface
		// NO ColorFill, only 
		//  Percentage="25" SmoothedLine="true"
		// AnRot="20" Perspective="true" ThreeDScaling="true" TwoDWalls="true"/>
		// all else should be default for surface charts 
	}
	public void makeContour() { // surface	-- for wireframe surface, no ColorFill
		//ColorFill="true" Percentage="25" SmoothedLine="true"/>
		// AnElev="90" pcDist="0" Perspective="true" ThreeDScaling="true" TwoDWalls="true"
	}
	*/

	/**
	 * set the marker format style for this chart
	 * <br>one of:
	 * <br>0 = no marker
	 * <br>1 = square
	 * <br>2 = diamond
	 * <br>3 = triangle
	 * <br>4 = X
	 * <br>5 = star
	 * <br>6 = Dow-Jones
	 * <br>7 = standard deviation
	 * <br>8 = circle
	 * <br>9 = plus sign
	 * <br>For Chart Types:
	 * <br>LINE, SCATTER
	 *
	 * @param int imf - marker format constant from list above
	 */
	public void setMarkerFormat( int imf )
	{ // line, scatter ...
		this.setChartOption( "MarkerFormat", String.valueOf( imf ) );
	}

	/**
	 * returs the JSON representation of this chart,
	 * based upon Dojo-charting-specifics
	 *
	 * @return String JSON representation of the chart
	 */
	public String getJSON()
	{
		JSONObject theChart = new JSONObject();
		try
		{
			JSONObject titles = new JSONObject();
			int type = this.getChartType();    // necessary for parsing AXIS options: horizontal charts "switch" axes ...

			// titles/labels
			titles.put( "title", this.getTitle() );
			titles.put( "XAxis",
			            (type != ChartConstants.BARCHART) ? (this.getXAxisLabel()) : this.getYAxisLabel() );    // bar axes are reversed ...
			titles.put( "YAxis", (type != ChartConstants.BARCHART) ? (this.getYAxisLabel()) : this.getXAxisLabel() );
			try
			{
				titles.put( "ZAxis", this.getZAxisLabel() );
			}
			catch( Exception e )
			{
				Logger.logWarn( "ChartHandle.getJSON failed getting zaxislable:" + e.toString() );
			}
			theChart.put( "titles", titles );

			// Chart dimensions (width, height)
			short[] coords = mychart.getCoords();
			theChart.put( "width", coords[ChartHandle.WIDTH] );
			theChart.put( "height", coords[ChartHandle.HEIGHT] );
			theChart.put( "row", mychart.getRow0() );        // TODO: may not be necessary, see usage ...
			theChart.put( "col", mychart.getCol0() );
			// Plot Area Background color
			int plotAreabg = this.getPlotAreaBgColor();
			if( plotAreabg == 0x4D || plotAreabg == 0x4E )
			{
				plotAreabg = FormatConstants.COLOR_WHITE;
			}
			theChart.put( "fill", FormatConstants.SVGCOLORSTRINGS[plotAreabg] );

			Double[] jMinMax = new Double[3];
			JSONObject chartObjectJSON = this.mychart.getChartObject().getJSON( this.mychart.getChartSeries(), this.wbh, jMinMax );
			double yMax = 1.0, yMin = 0.0;
			int nSeries = 0;
			try
			{    //it's possible to not have any series defined ...
				theChart.put( "Series", chartObjectJSON.getJSONArray( "Series" ) );
				// 20080416 KSC: Save SeriesJSON for later comparisons
				mychart.setSeriesJSON( chartObjectJSON.getJSONArray( "Series" ) );
				theChart.put( "SeriesFills",
				              chartObjectJSON.getJSONArray( "SeriesFills" ) );    // 20090729 KSC: capture bar colors or fills
			}
			catch( Exception e )
			{
			}
			theChart.put( "type", chartObjectJSON.getJSONObject( "type" ) );
			yMin = jMinMax[0].doubleValue();
			yMax = jMinMax[1].doubleValue();
			nSeries = jMinMax[2].intValue();

			// Axes + Category Labels + Grid Lines  			
			try
			{
				//inputJSONObject(theChart, this.getAxis(YAXIS, false).getJSON(this.wbh, type, yMax, yMin, nSeries));
				theChart.put( "y", mychart.getAxes().getJSON( YAXIS, this.wbh, type, yMax, yMin, nSeries ).getJSONObject( "y" ) );
				theChart.put( "back_grid", mychart.getAxes().getJSON( YAXIS, this.wbh, type, yMax, yMin, nSeries ).getJSONObject(
						"back_grid" ) );
			}
			catch( Exception e )
			{
			}
			try
			{
//				inputJSONObject(theChart, this.getAxis(XAXIS, false).getJSON(this.wbh, type, yMax, yMin, nSeries));
				theChart.put( "x", mychart.getAxes().getJSON( XAXIS, this.wbh, type, yMax, yMin, nSeries ).getJSONObject( "x" ) );
				theChart.put( "back_grid", mychart.getAxes().getJSON( Axis.YAXIS, this.wbh, type, yMax, yMin, nSeries ).getJSONObject(
						"back_grid" ) );
			}
			catch( Exception e )
			{
			}
			// TODO: 3d Charts (z axis)
			
			/*	        
			/* Chart Fonts
			sb.append(t(2) + "<ChartFontRecs>" + this.getChartFontRecsXML());
			sb.append("\n" + t(2) + "</ChartFontRecs>\n");
			sb.append(t(2) + "<ChartFonts" + this.getChartFontsXML() + "/>\n");
			*/
			/* Format Chart Area
			*/
			/* TODO: read in legend settings */

			// Chart Legend
			if( this.hasDataLegend() )
			{
				short s = this.mychart.getLegend().getLegendPosition();
				String[] legends = this.mychart.getLegends( -1 ); // -1 is flag for all rather than for a specific chart
				String l = "";
				for( int i = 0; i < legends.length; i++ )
				{
					l += legends[i] + ",";
				}
				if( l.length() > 0 )
				{
					l = l.substring( 0, l.length() - 1 );
				}
				theChart.put( "legend", new JSONObject( "{position:" + s + ",labels:[" + l + "]}" ) );
			}

		}
		catch( JSONException e )
		{
			Logger.logErr( "Error getting Chart JSON: " + e );
		}
		return theChart.toString();
	}

	/**
	 * utility to add a JSON object
	 * <br>This is an internal method that is not useful to the end user.
	 *
	 * @param source
	 * @param input
	 */
	protected void inputJSONObject( JSONObject source, JSONObject input )
	{
		if( source != null )
		{
			try
			{
				for( int j = 0; j < input.names().length(); j++ )
				{
					source.put( input.names().getString( j ), input.get( input.names().getString( j ) ) );
				}
			}
			catch( JSONException e )
			{
				Logger.logErr( "Error inputting JSON Object: " + e );
			}
		}
	}

	/**
	 * retrieves the saved Series JSON for comparisons
	 * <br>This is an internal method that is not useful to the end user.
	 *
	 * @return JSONArray
	 */
	public JSONArray getSeriesJSON()
	{
		return mychart.getSeriesJSON();
	}

	/**
	 * sets the saved Series JSON
	 * <br>This is an internal method that is not useful to the end user.
	 *
	 * @param JSONArray s -
	 * @throws JSONException
	 */
	public void setSeriesJSON( JSONArray s ) throws JSONException
	{
		mychart.setSeriesJSON( s );
	}

	/**
	 * retrieves current series and axis scale info in JSONObject form
	 * used upon chart updating
	 * <br>This is an internal method that is not useful to the end user.
	 *
	 * @return JSONObject series and axis info
	 */
	public JSONObject getCurrentSeries()
	{
		JSONObject retJSON = new JSONObject();
		// Retrieve series data + yMin yMax, nSeries
		Double[] jMinMax = new Double[3];
		try
		{
			// Series Data
//20080516 KSC: See above			JSONObject chartObjectJSON= ((GenericChartObject)this.mychart.getChartObject()).getJSON(
//					this.getAllChartSeriesHandles(), this.getCategories()[0], this.wbh, minMax);
			JSONObject chartObjectJSON = this.mychart.getChartObject().getJSON( this.mychart.getChartSeries(), this.wbh, jMinMax );

			try
			{
				retJSON.put( "Series", chartObjectJSON.getJSONArray( "Series" ) );
			}
			catch( Exception e )
			{
				Logger.logWarn( "ChartHandle.getCurrentSeries problem:" + e.toString() );
			}

			// Retrieve Axis Scale info
			double yMax = 0.0, yMin = 0.0;
			int nSeries = 0;
			yMin = jMinMax[0].doubleValue();
			yMax = jMinMax[1].doubleValue();
			nSeries = jMinMax[2].intValue();

			int type = this.getChartType();    // necessary for parsing AXIS options: horizontal charts "switch" axes ...
			// Axes + Category Labels + Grid Lines  			
/*KSC: TAKE OUT JSON STUFF FOR NOW; WILL REFACTOR LATER			try {
				inputJSONObject(retJSON, mychart.getAxes().getMinMaxJSON(YAXIS, this.wbh, type, yMax, yMin, nSeries));
			} catch (Exception e) { }			
			try {  
				inputJSONObject(retJSON, mychart.getAxes().getMinMaxJSON(XAXIS, this.wbh, type, yMax, yMin, nSeries));
			} catch (Exception e) { }*/
			// TODO: 3d Charts (z axis)
		}
		catch( JSONException e )
		{
			Logger.logErr( "ChartHandle.getCurrentSeries: Error getting Series JSON: " + e );
		}
		return retJSON;
	}

	/**
	 * returns a JSON representation of all Series Data (Legend, Categogies, Series Values) for the chart
	 * <br>This is an internal method that is not useful to the end user.
	 *
	 * @return String JSON representation
	 */
	public String getAllSeriesDataJSON()
	{
		JSONArray s = new JSONArray();
		ChartSeriesHandle[] series = this.getAllChartSeriesHandles();
		try
		{
			for( int i = 0; i < series.length; i++ )
			{
				JSONObject ser = new JSONObject();
				ser.put( "l", series[i].getSeriesLegendReference() );
				ser.put( "v", series[i].getSeriesRange() );
				ser.put( "b", series[i].getBubbleSizes() );
				if( i == 0 )
				{
					ser.put( "c", series[i].getCategoryRange() );    // 1 per chart
				}
				s.put( ser );
			}
		}
		catch( JSONException e )
		{
			Logger.logErr( "ChartHandle.getAllSeriesDataJSON: " + e );
		}
		return s.toString();
	}
	/**
	 * Take current Chart object and return the SVG code necessary to define it.
	 */
	/**
	 * TODO:
	 * Less Common Charts:
	 * STOCK
	 * RADAR
	 * SURFACE
	 * COLUMN- 3D, CONE, CYLINDER, PYRAMID
	 * BAR- 3D, CONE, CYLINDER, PYRAMID
	 * 3D PIE
	 * 3D LINE
	 * 3D AREA
	 * <p/>
	 * LINE CHART APPEARS THAT STARTS AND ENDS A BIT TOO EARLY *****************
	 * Z Axis
	 * <p/>
	 * CHART OPTIONS:
	 * STACKED
	 * CLUSTERED
	 */
	public String getSVG()
	{
		return getSVG( 1 );
	}

	/**
	 * /**
	 * Take current Chart object and return the SVG code necessary to define it,
	 * scaled to the desired percentage e.g. 0.75= 75%
	 *
	 * @param scale double scale factor
	 * @return String SVG
	 */
	public String getSVG( double scale )
	{
		HashMap<String, Double> chartMetrics = mychart.getMetrics( wbh );    // build or retrieve  Chart Metrics --> dimensions + series data ... 

		StringBuffer svg = new StringBuffer();
		// required header
		// svg.append("<?xml version=\"1.0\" standalone=\"no\"?>\r\n"); // referneces the DTD
		// svg.append("<!DOCTYPE svg PUBLIC \"-//W3C//DTD SVG 1.1//EN\" \"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd\">\r\n");

		// Define SVG Canvas:
		svg.append( "<svg width='" + (chartMetrics.get( "canvasw" ) * scale) + "px' height='" + (chartMetrics.get( "canvash" ) * scale) + "px' version=\"1.1\" xmlns=\"http://www.w3.org/2000/svg\"  xmlns:xlink=\"http://www.w3.org/1999/xlink\">\r\n" );
		svg.append( "<g transform='scale(" + scale + ")'  onclick='null;' onmousedown='handleClick(evt);' style='width:100%; height:" + (chartMetrics
				.get( "canvash" ) * scale) + "px'>" );        // scale chart -default scale=1 == 100%
		// JavaScript hooks
		svg.append( getJavaScript() );

		// Data Legend Box -- do before drawing plot area as legends box may change plot area coordinates
		String legendSVG = getLegendSVG( chartMetrics );    // but have to append it after because should overlap the plotarea

		String bgclr = this.mychart.getPlotAreaBgColor();
		// setup gradients
		svg.append( "<defs>" );
		svg.append( "<linearGradient id='bg_gradient' x1='0' y1='0' x2='0' y2='100%'>" );
		svg.append( "<stop offset='0' style='stop-color:" + bgclr + "; stop-opacity:1'/>" );
		svg.append( "<stop offset='" + chartMetrics.get( "w" ) + "' style='stop-color:white; stop-opacity:.5'/>" );
		svg.append( "</linearGradient>" );
		svg.append( "</defs>" );

		// PLOT AREA BG + RECT		
		// rectangle around entire chart canvas		
		if( !(mychart instanceof OOXMLChart) || !((OOXMLChart) mychart).roundedCorners )
		{
			svg.append( "<rect x='0' y='0' width='" + chartMetrics.get( "canvasw" ) + "' height='" + chartMetrics.get( "canvash" ) +
					            "' style='fill-opacity:1;fill:white' stroke='#CCCCCC' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'/>\r\n" );
		}
		else    // OOXML rounded corners
		{
			svg.append( "<rect x='0' y='0' width='" + chartMetrics.get( "canvasw" ) + "' height='" + chartMetrics.get( "canvash" ) +
					            "' rx='20' ry='20" +  /* rounded corners */
					            "' style='fill-opacity:1;fill:white' stroke='#CCCCCC' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'/>\r\n" );
		}

		// actual plot area
		svg.append( "<rect fill='" + bgclr + "'  style='fill-opacity:1;fill:url(#bg_gradient)' stroke='none' stroke-opacity='.5' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" +
				            " x='" + chartMetrics.get( "x" ) + "' y='" + chartMetrics.get( "y" ) + "' width='" + chartMetrics.get( "w" ) + "' height='" + chartMetrics
				.get( "h" ) + "' fill-rule='evenodd'/>\r\n" );

		// AXES, IF PRESENT - DO BEFORE ACTUAL SERIES DATA SO DRAWS CORRECTLY + ADJUST CHART DIMENSIONS 
		svg.append( mychart.getAxes().getSVG( XAXIS, chartMetrics, mychart.getChartSeries().getCategories() ) );
		svg.append( mychart.getAxes().getSVG( YAXIS, chartMetrics, mychart.getChartSeries().getCategories() ) );
		// TODO: Z Axis

		// After Axes and gridlines (if present), 
		// ACTUAL bar/series/area/etc. svg generated from series and scale data
		svg.append( this.mychart.getChartObject().getSVG( chartMetrics, mychart.getAxes().getMetrics(), mychart.getChartSeries() ) );

		svg.append( legendSVG );    // append legend SVG obtained above

		// CHART TITLE

		if( mychart.getTitleTd() != null )
		{
			svg.append( mychart.getTitleTd().getSVG( chartMetrics ) );
		}

		svg.append( "</g>" );
		svg.append( "</svg>" );
/*		
//KSC: TESTING: REMOVE WHEN DONE		
if (WorkBookFactory.PID==WorkBookFactory.E360) { // save svg for testing purposes
	try {
		java.io.File f= new java.io.File("c:/eclipse/workspace/testfiles/extenxls/output/charts/FromSheetster.svg");
		java.io.FileOutputStream fos= new java.io.FileOutputStream(f);
		fos.write(svg.toString().getBytes());
		fos.flush();
		fos.close();
	} catch (Exception e) {}
}
*/
		return svg.toString();
	}

	private String getLegendSVG( HashMap<String, Double> chartMetrics )
	{
		try
		{
			return mychart.getLegend().getSVG( chartMetrics, mychart.getChartObject(), mychart.getChartSeries() );
		}
		catch( NullPointerException ne )
		{ // no legend??
			return null;
		}
	}

	/**
	 * returns the svg for javascript for highlight and restore
	 *
	 * @return
	 */
	protected String getJavaScript()
	{
		StringBuffer svg = new StringBuffer();
		svg.append( "<script type='text/ecmascript'>" );
		svg.append( "  <![CDATA[" );

		svg.append( "try{var grid = parent.parent.uiWindowing.getActiveDoc().getContent().contentWindow;" );
		svg.append( "var selection = new grid.cellSelectionResizable();}catch(x){;}" );

		svg.append( "function highLight(evt) {" );
		svg.append( "this.bgc = evt.target.getAttributeNS(null, 'fill');" );
		svg.append( "evt.target.setAttributeNS(null,'fill','gold');" ); // rgb('+ red +','+ green+','+blue+')');");
		svg.append( "evt.target.setAttributeNS(null,'stroke-width','2');" );
		//svg.append("evt.target.setAttributeNS(null,'stroke-color','white');");
		svg.append( "}" );

		svg.append( "function restore(evt) {" );
		svg.append( "evt.target.setAttributeNS(null,'fill',this.bgc);" ); // rgb('+ red +','+ green+','+blue+')');");
		svg.append( "evt.target.setAttributeNS(null,'stroke-width','1');" );
		svg.append( "}" );

		svg.append( "function handleClick(evt) {" );
//		svg.append("try{parent.parent.uiWindowing.getActiveSheet().book.handleMouseClick(evt);}catch(x){;}");
		svg.append( "try{parent.chart.handleClick(evt);}catch(e){;}" );
		svg.append( "}" );

		svg.append( "\r\n" );
		svg.append( "function showRange(range) {" );
		svg.append( "try{selection.select(new grid.cellRange(range.toString())).show();}catch(x){;}" );
		svg.append( "}" );

		svg.append( "\r\n" );
		svg.append( "function hideRange() {" );
		svg.append( "try{selection.clear();}catch(x){;}" );
		svg.append( "}" );

		svg.append( "]]>" );
		svg.append( "</script>" );
		return svg.toString();
	}

	/**
	 * debugging utility remove when done
	 */
	public void WriteMainChartRecs( String fName )
	{
		mychart.getChartObject().WriteMainChartRecs( fName );
	}
}
