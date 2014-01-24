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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashMap;

public class RadarChart extends ChartType
{
	private static final Logger log = LoggerFactory.getLogger( RadarChart.class );
	//private Radar radar = null; can be Radar or RadarArea

	public RadarChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
//		radar = (Radar) charttype;
	}

	public void setFilled( boolean isFilled )
	{
		if( isFilled )
		{
			chartobj = (RadarArea) RadarArea.getPrototype();
		}
		else
		{
			chartobj = (Radar) Radar.getPrototype();
		}
		chartobj.setParentChart( cf.parentChart );
		cf.chartArr.remove( 0 );
		cf.chartArr.add( chartobj );
	}

	public boolean getIsFilled()
	{
		return (chartobj.chartType == ChartConstants.RADARAREACHART);
	}

	/**
	 * returns SVG to represent the actual chart object i.e. the representation
	 * of the series data in the particular format (BAR, LINE, AREA, etc.)
	 *
	 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min,
	 *                     max
	 * @param axisMetrics  maps specific axis options such as xAxisReversed, xPattern ...
	 * @param categories
	 * @param series       arraylist of double[] series values
	 * @param seriescolors int[] of series or bar colors color ints
	 * @return String svg
	 */
	@Override
	public String getSVG( HashMap<String, Double> chartMetrics, HashMap<String, Object> axisMetrics, ChartSeries s )
	{
		double x = chartMetrics.get( "x" ); // + (!yAxisReversed?0:w); // x is
		// constant at x origin unless
		// reversed
		double y = chartMetrics.get( "y" );
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double max = chartMetrics.get( "max" );
		double min = chartMetrics.get( "min" );
		Object[] categories = s.getCategories();
		ArrayList series = s.getSeriesValues();
		String[] seriescolors = s.getSeriesBarColors();
		String[] legends = s.getLegends();
		// x value for each point= w/(ncategories + 1) 1st one is xv*2 then
		// increases from there
		// y value for each point= h/YMAX
		if( series.size() == 0 )
		{
			log.error( "Radar.getSVG: error in series" );
			return "";
		}
		StringBuffer svg = new StringBuffer();

		// obtain data label, marker info for chart + series colors
		int[] dls = getDataLabelInts(); // get array of data labels (can be
		// specific per series ...)
		int[] markers = getMarkerFormats(); // get an array of marker formats
		// per series
		int n = series.size();
		int nseries = categories.length;

		double centerx = (w / 2) + x;
		double centery = (h / 2) + y;
		double percentage = 1.0 / nseries; // divide into equal sections
		double radius = Math.min( w, h ) / 2.3; // should take up almost entire
		// w/h of chart

		svg.append( "<g>\r\n" );
		// define marker shapes for later use
		svg.append( MarkerFormat.getMarkerSVGDefs() );
		// for each series
		for( int i = 0; i < n; i++ )
		{
			String points = "";
			String labels = "";
			double angle = 90; // starts straight up
			double[] curseries = (double[]) series.get( i );
			String[] curranges = (String[]) s.getSeriesRanges().get( i );
			double x0 = 0;
			double y0 = 0;
			for( int j = 0; j < curseries.length; j++ )
			{
				// get next point as a percentage of the radius
				double r = radius * (curseries[j] / (max - min));
				double x1 = centerx + (r * (Math.cos( Math.toRadians( angle ) ))); //
				double y1 = centery - (r * (Math.sin( Math.toRadians( angle ) )));
				if( j == 0 )
				{ // save initial points so can close the loop at
					// end
					x0 = x1;
					y0 = y1;
				}
				points += x1 + "," + y1 + " ";
				String l = getSVGDataLabels( dls, axisMetrics, curseries[j], percentage, j, legends, categories[j].toString() );
				if( l != null )
				{
					double labelx1 = centerx + ((r + 5) * (Math.cos( Math.toRadians( angle ) ))); //
					double labely1 = centery - ((r + 5) * (Math.sin( Math.toRadians( angle ) )));
					labels += "<text x='" + (labelx1) + "' y='" + (labely1) + "' style='text-anchor: middle;' " + getDataLabelFontSVG() + ">" + l + "</text>\r\n";
				}
				angle -= (percentage * 360); // next point on next category
				// radial line
			}
			// close loop
			points += x0 + "," + y0;
			// 1st line is black
			svg.append( "<polyline " + getScript( "" ) + " fill-opacity='0' " + getStrokeSVG( 4,
			                                                                                  "black" ) + " points='" + points + "'" + "/>\r\n" );
			// 2nd line is the series color
			svg.append( "<polyline " + getScript( "" ) + "   id='series_" + (i + 1) + "' fill='none' fill-opacity='0' " + getStrokeSVG( 3,
			                                                                                                                            seriescolors[i] ) + " points='" + points + "'" + "/>\r\n" );
			// Markers, if any, along data points in series
			if( markers[i] > 0 )
			{
				String[] markerpoints = points.split( " " );
				for( String markerpoint : markerpoints )
				{
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
	 * gets the chart-type specific ooxml representation: <radarChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:radarChart>" );
		cooxml.append( "\r\n" );

		String style = "standard";
		if( getIsFilled() )
		{
			style = "filled";
		}
		else
		{
			int[] markers = getMarkerFormats();
			for( int m : markers )
			{
				if( m != 0 )
				{
					style = "marker";
					break;
				}
			}
		}
		cooxml.append( "<c:radarStyle val=\"" + style + "\"/>" );
		// vary colors???

		// *** Series Data: ser, cat, val for most chart types
		cooxml.append( getParentChart().getChartSeries().getOOXML( getChartType(), false, 0 ) );

		// chart data labels, if any
		// TODO: FINISH
		// cooxml.append(getDataLabelsOOXML(cf));

		// axis ids - unsigned int strings
		cooxml.append( "<c:axId val=\"" + catAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:axId val=\"" + valAxisId + "\"/>" );
		cooxml.append( "\r\n" );

		cooxml.append( "</c:radarChart>" );
		cooxml.append( "\r\n" );
		return cooxml;
	}
}
