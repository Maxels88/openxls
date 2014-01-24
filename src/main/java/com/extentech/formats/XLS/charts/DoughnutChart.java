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
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashMap;

public class DoughnutChart extends ChartType
{
	private static final Logger log = LoggerFactory.getLogger( DoughnutChart.class );
	private Pie doughnut = null;

	public DoughnutChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		doughnut = (Pie) charttype;
	}

	/**
	 * return the JSON that
	 *
	 * @param seriesvals
	 * @param range
	 * @return
	 */
	public JSONObject getJSON( ChartSeriesHandle[] series, WorkBookHandle wbh, Double[] minMax ) throws JSONException
	{
		JSONObject chartObjectJSON = new JSONObject();
		// Type JSON
		chartObjectJSON.put( "type", getTypeJSON() );

		// TODO: check out labels options chosen: default label is Category name + percentage ...
		double yMax = 0.0;
		double yMin = 0.0;
		double len = 0.0;
		JSONArray pieSeries = new JSONArray();
		try
		{
			String range = series[0].getCategoryRange();        // 20080516 KSC: retrieve cat range instead of parameter
			JSONArray cats = CellRange.getValuesAsJSON( range, wbh );    // parse category range into JSON Array
			JSONArray seriesvals = CellRange.getValuesAsJSON( series[0].getSeriesRange(), wbh );
			double piesum = 0;
			for( int k = 0; k < seriesvals.length(); k++ )
			{
				piesum += seriesvals.getDouble( k );
				yMax = Math.max( yMax, seriesvals.getDouble( k ) );
				yMin = Math.min( yMin, seriesvals.getDouble( k ) );
			}
			double percent = 100 / piesum;
			for( int k = 0; k < seriesvals.length(); k++ )
			{
				JSONObject piepoint = new JSONObject();
				piepoint.put( "y", seriesvals.getDouble( k ) );
				piepoint.put( "text", cats.getString( k ) + "\n" + Math.round( percent * seriesvals.getDouble( k ) ) + "%" );
				piepoint.put( "color", FormatConstants.SVGCOLORSTRINGS[series[k].getPieChartSliceColor( k )] );
				piepoint.put( "stroke", getDarkColor() );
				pieSeries.put( piepoint );
			}
			len = seriesvals.length();
		}
		catch( Exception e )
		{
			// TODO: warning ...?
		}
		// 20090717 KSC: input outside of try/catch to always set
		minMax[0] = yMin;
		minMax[1] = yMax;
		minMax[2] = len;
		chartObjectJSON.put( "Series", pieSeries );
		chartObjectJSON.put( "SeriesFills", "" );    // not applicable for pie charts; color is set above
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
		typeJSON.put( "labelOffset", -25 );    // default
		typeJSON.put( "precision", 0 );        // default - rounds percentages up
		typeJSON.put( "type", "Pie" );
		// TODO: Interpret distance ...
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
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double max = chartMetrics.get( "max" );
		double min = chartMetrics.get( "min" );
		Object[] categories = s.getCategories();
		ArrayList series = s.getSeriesValues();
		String[] seriescolors = s.getSeriesBarColors();
		String[] legends = s.getLegends();
		StringBuffer svg = new StringBuffer();
		if( series.size() == 0 )
		{
			log.error( "DoughnutChart.getSVG: error in series" );
			return "";
		}
		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)
		boolean threeD = isThreeD();
		final int LABELOFFSET = 15;

		int n = series.size();
		double centerx = 0.0;
		double centery = 0.0;
		double radius = 0.0;
		double radiusy = 0.0;
		centerx = w / 2 + chartMetrics.get( "x" );
		centery = h / 2 + chartMetrics.get( "y" );
		if( threeD )
		{
			svg.append( "<defs>\r\n" );
			svg.append( "<filter id=\"multiply\">\r\n" );
			svg.append( "<feBlend mode=\"multiply\" in2=\"\"/>\r\n" );
			svg.append( "</filter>\r\n" );
			svg.append( "</defs>\r\n" );
		}
		svg.append( "<g>\r\n" );

		//radius= Math.min(w, h)/2.3;	// should take up almost entire w/h of chart  		
		radius = Math.min( w, h ) / 1.9;    // should take up almost entire w/h of chart
		double r0 = (radius / 2) / n;    // radius/2 to account for hole
		double r = radius;    // start at outside and work inside
		for( int i = n - 1; i >= 0; i-- )
		{    // for each series
			double[] curseries = (double[]) series.get( i );        // series data points
			double total = 0.0;
			for( double cursery : curseries )
			{    // get total in order to calculate percentages
				total += cursery;
			}
			if( dls.length == 1 )
			{    // no seriess-specific data labels; expand to entire series for loop below
				int dl = dls[0];
				dls = new int[curseries.length];
				for( int z = 0; z < dls.length; z++ )
				{
					dls[z] = dl;
				}
			}
			svg.append( "<circle " + getScript( "" ) + " cx='" + centerx + "' cy='" + centery + "' r='" + r + "' " + getStrokeSVG( 2,
			                                                                                                                       getDarkColor() ) + " fill='none'/>\r\n" );
			double x = centerx + r;
			double y = centery;
			String path = "";
			double percentage = 0;
			double lasta = 0;
			int largearcflag = 0;
			int sweepflag = 0;
			// Now create each pie wedge according to it's percentage value 
			for( int j = 0; j < curseries.length; j++ )
			{
				percentage = curseries[j] / total;
				double angle = (percentage * 360) + lasta;
				double x1 = centerx + (r * (Math.cos( Math.toRadians( angle ) )));
				double y1 = centery - (r * (Math.sin( Math.toRadians( angle ) )));
				if( (percentage * 360) > 180 )
				{
					sweepflag = 0;
					largearcflag = 1;
				}
				else
				{
					largearcflag = 0;
				}
				path = "M" + centerx + " " + centery + " L" + x + " " + y + " A" + r + " " + r + " 0 " + largearcflag + " " + sweepflag + " " + x1 + " " + y1 + " L" + centerx + " " + centery + "Z";
				// paint wedge of color to center of chart -- the inner will overwrite so don't have to worry about segments and arcs, etc
				svg.append( "<path  " + getScript( "" ) + "  fill='" + seriescolors[j] + "'   id='series_" + (j + 1) + "' fill-opacity='" + getFillOpacity() + "' " + getStrokeSVG() + " path='' d='" + path + "' fill-rule='evenodd'/>\r\n" );
				// data labels
				String l = getSVGDataLabels( dls, axisMetrics, curseries[j], percentage, j, legends, categories[j].toString() );
				if( l != null )
				{
					double halfa = ((percentage / 2) * 360) + lasta;    // center in area
					double x2 = centerx + ((r - (r0 / 2)) * (Math.cos( Math.toRadians( halfa ) )));
					double y2 = centery - ((r - (r0 / 2)) * (Math.sin( Math.toRadians( halfa ) )));
					svg.append( "<text x='" + (x2) + "' y='" + (y2) + "' vertical-align='bottom' " + getDataLabelFontSVG() + " style='text-anchor: middle;'>" + l + "</text>\r\n" );
				}
				lasta = angle;
				x = x1;
				y = y1;
			}    // each point in current series
			r -= r0;
		}    // each series
		// complete inner circle & create "hole"
		svg.append( "<circle " + getScript( "" ) + " cx='" + centerx + "' cy='" + centery + "' r='" + r + "' " + getStrokeSVG( 2,
		                                                                                                                       getDarkColor() ) + " fill='white'/>\r\n" );
		svg.append( "</g>\r\n" );

		return svg.toString();
	}
	/**
	 * 		Of the four candidate arc sweeps, two will represent an arc sweep of greater than or equal to 180 degrees (the "large-arc"), 
	 *      and two will represent an arc sweep of less than or equal to 180 degrees (the "small-arc"). 
	 *      If large-arc-flag is '1', then one of the two larger arc sweeps will be chosen; otherwise, if large-arc-flag is '0', one of the smaller arc sweeps will be chosen,
	 * 	If sweep-flag is '1', then the arc will be drawn in a "positive-angle" direction (i.e., the ellipse formula x=cx+rx*cos(theta)
	 *   and y=cy+ry*sin(theta) is evaluated such that theta starts at an angle corresponding to the current point and increases positively until the arc reaches (x,y)).
	 *   A value of 0 causes the arc to be drawn in a "negative-angle" direction
	 *  (i.e., theta starts at an angle value corresponding to the current point and decreases until the arc reaches (x,y)).
	 */

	/**
	 * gets the chart-type specific ooxml representation: <doughnutChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:doughnutChart>" );
		cooxml.append( "\r\n" );
		// vary colors???
		cooxml.append( "<c:varyColors val=\"1\"/>" );

		// *** Series Data:	ser, cat, val for most chart types
		cooxml.append( getParentChart().getChartSeries().getOOXML( getChartType(), false, 0 ) );

		// chart data labels, if any
		//TODO: FINISH		    	
		//cooxml.append(getDataLabelsOOXML(cf));
		// TODO: firstSLiceAng
		cooxml.append( "<c:firstSliceAng val=\"0\"/>" );
		cooxml.append( "<c:holeSize val=\"" + getChartOption( "donutSize" ) + "\"/>" );

		cooxml.append( "</c:doughnutChart>" );
		cooxml.append( "\r\n" );

		return cooxml;
	}

}
