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

import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.ChartSeriesHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.FormatConstants;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.toolkit.Logger;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.util.ArrayList;
import java.util.HashMap;

public class ScatterChart extends ChartType
{
	private Scatter scatter = null;

	public ScatterChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		scatter = (Scatter) charttype;
	}

	public JSONObject getJSON( ChartSeriesHandle[] series, WorkBookHandle wbh, Double[] minMax ) throws JSONException
	{
		JSONObject chartObjectJSON = new JSONObject();

		// Type JSON
		chartObjectJSON.put( "type", this.getTypeJSON() );

		// Deal with Series
		double yMax = 0.0, yMin = 0.0;
		int nSeries = 0;
		JSONArray seriesJSON = new JSONArray();
		JSONArray seriesCOLORS = new JSONArray();
		boolean bHasBubbles = false;
		try
		{
			// must trap min and max for axis tick and units
			for( int i = 0; i < series.length; i++ )
			{
				JSONArray seriesvals = CellRange.getValuesAsJSON( series[i].getSeriesRange(), wbh );
				nSeries = Math.max( nSeries, seriesvals.length() );
				for( int j = 0; j < seriesvals.length(); j++ )
				{
					try
					{
						yMax = Math.max( yMax, seriesvals.getDouble( j ) );
						yMin = Math.min( yMin, seriesvals.getDouble( j ) );
					}
					catch( NumberFormatException n )
					{
						;
					}
				}
				if( !series[i].hasBubbleSizes() )
				{
					seriesJSON.put( seriesvals );
				}
				else
				{
					bHasBubbles = true;
				}
				seriesCOLORS.put( FormatConstants.SVGCOLORSTRINGS[series[i].getSeriesColor()] );
			}
			if( bHasBubbles )
			{
				// 20080423 KSC: Go thru a second time, after obtaining yMax and yMin, for bubble sizes ...
				for( int i = 0; i < series.length; i++ )
				{
					JSONArray bubbles = new JSONArray();
					JSONArray seriesvals = CellRange.getValuesAsJSON( series[i].getSeriesRange(), wbh );
					JSONArray catvals = CellRange.getValuesAsJSON( series[i].getCategoryRange(), wbh );
					JSONArray bubblesizes = CellRange.getValuesAsJSON( series[i].getBubbleSizes(), wbh );
					for( int j = 0; j < catvals.length(); j++ )
					{
						JSONObject jo = new JSONObject();
						try
						{
							jo.put( "x", catvals.getDouble( j ) );
						}
						catch( Exception e )
						{
							jo.put( "x", j + 1 );
						}
						jo.put( "y", seriesvals.getDouble( j ) );
						jo.put( "size",
						        Math.round( bubblesizes.getDouble( j ) / ((yMax - yMin) / nSeries) ) );        // TODO: bubble sizes ration is a guess!!
						bubbles.put( jo );
					}
					seriesJSON.put( bubbles );
				}
			}
			chartObjectJSON.put( "Series", seriesJSON );
			chartObjectJSON.put( "SeriesFills", seriesCOLORS );
		}
		catch( JSONException je )
		{
			// TODO: Log error
		}
		minMax[0] = new Double( yMin );
		minMax[1] = new Double( yMax );
		minMax[2] = new Double( nSeries );
		return chartObjectJSON;
	}

	/**
	 * return the type JSON for this Chart Object
	 *
	 * @return
	 */
	@Override
	public JSONObject getTypeJSON() throws JSONException
	{
		JSONObject typeJSON = new JSONObject();
/*    	String dojoType;
		if (this.chartType==ChartConstants.SCATTER) {		
			dojoType="Default";	// = MarkersOnly            			
			typeJSON.put("markers", true);
		} else { // Bubble
			dojoType="Bubble";		// "MarkersOnly";
			// shadows: {dx: 2, dy: 2, dw: 2}
		}
    	//controls Data Legends, % Distance of sections, Line Format, Area Format, Bar Shapes ...
    	DataFormat df= this.getParentChart().getChartFormat().getDataFormatRec(false);
    	if (df!=null) {
	    	for (int z= 0; z < df.chartArr.size(); z++) {
	            BiffRec b = (BiffRec)df.chartArr.get(z);
	            if (b instanceof LineFormat) {
	            	if (chartType==ChartConstants.SCATTER)
	                	typeJSON.put("type", "MarkersOnly");	// if Has LineFormat means NO lines!
	            } else if (b instanceof Serfmt) {
	            	Serfmt s= (Serfmt) b;
	            	if (s.getSmoothLine()) {
	            		// can be:  Scatter with markers and smoothed lines                        	
	            		typeJSON.put("type", "Default");	// change to Line with Markers
	    				typeJSON.put("markers", true);		// default, apparently- doesn't require a MarkerFormat record
	            	}
	            }
	        }
    	}
    	typeJSON.put("type", dojoType);
*/
		return typeJSON;
	}

	/**
	 * returns SVG to represent the actual chart object i.e. the representation of the series data in the particular format (BAR, LINE, AREA, etc.)
	 *
	 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @param axisMetrics  maps specific axis options such as xAxisReversed, xPattern ...
	 * @param categories
	 * @param series       arraylist of double[] series values
	 * @param seriescolors int[] of series or bar colors color ints
	 * @return String svg
	 */
	@Override
	public String getSVG( HashMap<String, Double> chartMetrics, HashMap<String, Object> axisMetrics, ChartSeries s )
	{
		double x = chartMetrics.get( "x" );    // + (!yAxisReversed?0:w);	// x is constant at x origin unless reversed
		double y = chartMetrics.get( "y" );
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double max = chartMetrics.get( "max" );
		Object[] categories = s.getCategories();
		ArrayList series = s.getSeriesValues();
		String[] seriescolors = s.getSeriesBarColors();
		String[] legends = s.getLegends();
		// x value for each point= w/(ncategories + 1) 1st one is xv*2 then increases from there
		// y value for each point= h/YMAX 
		StringBuffer svg = new StringBuffer();

		if( series.size() == 0 )
		{
			Logger.logErr( "Scatter.getSVG: error in series" );
			return "";
		}
		// gather data labels, markers, has lines, series colors for chart
		boolean threeD = cf.isThreeD( ChartConstants.SCATTERCHART );
		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)
		int n = series.size();
		boolean hasLines = cf.getHasLines(); // should be per-series?
		int[] markers = getMarkerFormats();    // get an array of marker formats per series
		if( !hasLines && (markers[0] == 0) )
		{
			// if no lines AND no markers set, MUST use default markers (this is what excel does ...)
			int[] defaultmarkers = { 2, 3, 1, 4, 8, 9, 5, 6, 7 };
			for( int i = 0; i < n; i++ )
			{
				markers[i] = defaultmarkers[i];
			}
		}
		/**
		 * A Scatter chart has two value axes, showing one set of numerical data along the x-axis 
		 * 	and another along the y-axis. 
		 * It combines these values into single data points and displays them in uneven intervals, 
		 * or clusters.
		 */
		double[] seriesx = null;
		double xfactor = 0, yfactor = 0;    //
		boolean TEXTUALXAXIS = true;
		// get x axis max/min for an x axis which is a value axis 
		double xmin = Double.MAX_VALUE, xmax = Double.MIN_VALUE;
		seriesx = new double[categories.length];
		for( int j = 0; j < categories.length; j++ )
		{
			try
			{
				seriesx[j] = new Double( categories[j].toString() );
				xmax = Math.max( xmax, seriesx[j] );
				xmin = Math.min( xmin, seriesx[j] );
				TEXTUALXAXIS = false;        // if ANY category val is a double, assume it's a normal xyvalue axis
			}
			catch( Exception e )
			{
				; /* keep going */
			}
		}
		if( !TEXTUALXAXIS )
		{
			double d[] = ValueRange.calcMaxMin( xmax, xmin, w );
			xfactor = w / (d[2]);        // w/maximum scale
		}
		else
		{
			xfactor = w / (categories.length + 1);    // w/#categories
		}

		if( max != 0 )
		{
			yfactor = h / max;    // h/YMAXSCALE
		}
		svg.append( "<g>\r\n" );
		// define marker shapes for later use 
		svg.append( MarkerFormat.getMarkerSVGDefs() );
		// for each series 
		for( int i = 0; i < n; i++ )
		{
			// two lines- 1 black, 1 color
			String points = "";
			String labels = "";
			double[] seriesy = (double[]) series.get( i );
			for( int j = 0; j < seriesy.length; j++ )
			{
				double xval = 0;
				if( TEXTUALXAXIS /*|| i > 0*/ )
				{
					xval = j + 1;
				}
				else
				{
					xval = seriesx[j];
				}
				points += ((x) + xval * xfactor) + "," + ((y + h) - (seriesy[j] * yfactor));
				points += " ";
				String l = getSVGDataLabels( dls, axisMetrics, seriesy[j], 0, i, legends, categories[j].toString() );
				if( l != null )
				{
					labels += "<text x='" + (12 + (x) + xval * xfactor) + "' y='" + (((y + h) - (seriesy[j] * yfactor))) +
							"' " + this.getDataLabelFontSVG() + ">" + l + "</text>\r\n";
				}

			}
			if( hasLines )
			{
				svg.append( getLineSVG( points, seriescolors[i] ) );
			}
			// Markers, if any, along data points in series
			if( markers[i] > 0 )
			{
				String[] markerpoints = points.split( " " );
				for( int j = 0; j < markerpoints.length; j++ )
				{
					String markerpoint = markerpoints[j];
					String[] xy = markerpoint.split( "," );
					double xx = Double.valueOf( xy[0] );
					double yy = Double.valueOf( xy[1] );
					svg.append( MarkerFormat.getMarkerSVG( xx, yy, seriescolors[i], markers[i] ) + "\r\n" );
				}
			}
			// labels after lines and markers  
			svg.append( labels );
		}
		svg.append( "</g>\r\n" );
		return svg.toString();
	}

	/**
	 * returns the SVG necessary to define a line at points in color clr
	 *
	 * @param points String of SVG points
	 * @param clr    SVG color String
	 * @return
	 */
	private String getLineSVG( String points, String clr )
	{
		String s = "";
		// each line is comprised of 1 black line and 1 series color line:			
		// 1st line is black
		s = "<polyline fill='none' fill-opacity='0' " + getStrokeSVG( 1, this.getDarkColor() ) +
				" points='" + points + "'" +
				"/>\r\n";
		// 2nd line is the series color
		s += "<polyline fill='none' fill-opacity='0' " + getStrokeSVG( 2, clr ) +
				" points='" + points + "'" +
				"/>\r\n";
		return s;
	}

	/**
	 * gets the chart-type specific ooxml representation: <scatterChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:scatterChart>" );
		cooxml.append( "\r\n" );
		int[] markers = getMarkerFormats();
		String style = null;
		for( int m : markers )
		{
			if( m != 0 )
			{
				if( this.getHasSmoothLines() )
				{
					style = "smoothMarker";
				}
				else if( this.getHasLines() )
				{
					style = "lineMarker";
				}
				else
				{
					style = "marker";
				}
				break;
			}
		}
		if( style == null && this.getHasLines() )
		{
			style = "line";
		}
		if( style == null && this.getHasSmoothLines() )
		{
			style = "smooth";
		}
		if( style == null )
		{
			style = "none";
		}

		cooxml.append( "<c:scatterStyle val=\"" + style + "\"/>" );
		// vary colors???

		// *** Series Data: ser, cat, val for most chart types
		cooxml.append( this.getParentChart().getChartSeries().getOOXML( this.getChartType(), false, 0 ) );

		// chart data labels, if any
		// TODO: FINISH
		// cooxml.append(getDataLabelsOOXML(cf));

		// axis ids - unsigned int strings
		cooxml.append( "<c:axId val=\"" + catAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:axId val=\"" + valAxisId + "\"/>" );
		cooxml.append( "\r\n" );

		cooxml.append( "</c:scatterChart>" );
		cooxml.append( "\r\n" );
		return cooxml;
	}
}
