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

import com.extentech.formats.XLS.WorkBook;
import com.extentech.toolkit.Logger;
import org.json.JSONException;
import org.json.JSONObject;

import java.util.ArrayList;
import java.util.HashMap;

public class LineChart extends ChartType
{
	Line line = null;

	public LineChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		line = (Line) charttype;
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
		if (!this.isStacked()) {
			dojoType="Default";
			typeJSON.put("markers", true);		// default, apparently- doesn't require a MarkerFormat record
		} else
			dojoType="StackedLines";
    	//controls Data Legends, % Distance of sections, Line Format, Area Format, Bar Shapes ...
    	DataFormat df= this.getParentChart().getChartFormat().getDataFormatRec(false);
    	if (df!=null) {
	    	for (int z= 0; z < df.chartArr.size(); z++) {
	            BiffRec b = (BiffRec)df.chartArr.get(z);
	            // LineFormat,AreaFormat,PieFormat,MarkerFormat
	            if (b instanceof Serfmt) {            	
	            	Serfmt s= (Serfmt) b;
	            	if (s.getSmoothLine()) {
	            		// can be:  Scatter with markers and smoothed lines                        	
	            		typeJSON.put("type", "Default");	// change to Line with Markers
	    				typeJSON.put("markers", true);		// default, apparently- doesn't require a MarkerFormat record
	            	}
	            } else if (b instanceof MarkerFormat) {	// markers for legend (attached label should follow) BUT also can mean NO markers
	            	if (((MarkerFormat)b).getMarkerFormat()==0 && dojoType.equals("Default"))
	            		typeJSON.remove("markers");
	            }
	    	}
    	}		
    	typeJSON.put("type", dojoType);*/
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
		double min = chartMetrics.get( "min" );
		Object[] categories = s.getCategories();
		ArrayList series = s.getSeriesValues();
		String[] seriescolors = s.getSeriesBarColors();
		String[] legends = s.getLegends();
		if( series.size() == 0 )
		{
			Logger.logErr( "Line.getSVG: error in series" );
			return "";
		}

		// x value for each point= w/(ncategories + 1) 1st one is xv*2 then increases from there
		// y value for each point= h/YMAX 
		StringBuffer svg = new StringBuffer();
		// get data labels, marker formats, series colors
		int n = series.size();
		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)
		int[] markers = getMarkerFormats();
		double xfactor = 0, yfactor = 0;    //
		if( categories.length != 0 )
		{
			xfactor = w / (categories.length);    // w/#categories
		}
		else
		{
			xfactor = w;
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
			// each visible line on the chart consists of two lines- 1 black, 1 color
			String points = "";
			String labels = "";
			double[] curseries = (double[]) series.get( i );
			String[] curranges = (String[]) s.getSeriesRanges().get( i );
			for( int j = 0; j < curseries.length; j++ )
			{
				points += ((x) + ((j + .5) * xfactor)) + "," + ((y + h) - (curseries[j] * yfactor));
				points += " ";
				String l = getSVGDataLabels( dls, axisMetrics, curseries[j], 0, i, legends, categories[j].toString() );
				if( l != null )
				{
					double xx = (2 + (x) + ((j + .5) * xfactor));
					if( markers[i] > 0 )
					{
						xx += 10;    // scoot over for markers
					}
					labels += "<text x='" + xx + "' y='" + (((y + h) - (curseries[j] * yfactor))) +
							"' " + this.getDataLabelFontSVG() + ">" + l + "</text>\r\n";
				}
			}
			// 1st line is black
			svg.append( "<polyline " + getScript( "" ) + " fill='none' fill-opacity='0' " + getStrokeSVG( 4,
			                                                                                              this.getDarkColor() ) + " points='" + points + "'" + "/>\r\n" );
			// 2nd line is the series color
			svg.append( "<polyline " + getScript( "" ) + "  id='series_" + (i + 1) + "' fill='none' fill-opacity='0' stroke='" + seriescolors[i] + "' stroke-opacity='1' stroke-width='3' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" +
					            " points='" + points + "'" +
					            "/>\r\n" );
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
			// data labels, if any, after lines and markers  
			svg.append( labels );
		}

		svg.append( "</g>\r\n" );
		return svg.toString();
	}

	/**
	 * gets the chart-type specific ooxml representation: <lineChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:lineChart>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:grouping val=\"" );

		if( this.is100PercentStacked() )
		{
			cooxml.append( "percentStacked" );
		}
		else if( this.isStacked() )
		{
			cooxml.append( "stacked" );
		}
		//			} else if (this.isClustered())
		//				grouping="clustered";
		else
		{
			cooxml.append( "standard" );
		}
		cooxml.append( "\"/>" );
		cooxml.append( "\r\n" );
		// vary colors???

		// *** Series Data:	ser, cat, val for most chart types
		cooxml.append( this.getParentChart().getChartSeries().getOOXML( this.getChartType(), false, 0 ) );

		// chart data labels, if any
		//TODO: FINISH dlbls		    	
		//cooxml.append(getDataLabelsOOXML(cf));

		//dropLines
		ChartLine cl = this.cf.getChartLinesRec( ChartLine.TYPE_DROPLINE );
		if( cl != null )
		{
			cooxml.append( cl.getOOXML() );
		}
		// hiLowLines
		cl = this.cf.getChartLinesRec( ChartLine.TYPE_HILOWLINE );
		if( cl != null )
		{
			cooxml.append( cl.getOOXML() );
		}
		// upDownBars
		cooxml.append( cf.getUpDownBarOOXML() );
		// marker
		// smooth

		// axis ids	 - unsigned int strings
		cooxml.append( "<c:axId val=\"" + catAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:axId val=\"" + valAxisId + "\"/>" );
		cooxml.append( "\r\n" );

		cooxml.append( "</c:lineChart>" );
		cooxml.append( "\r\n" );

		return cooxml;
	}
}
