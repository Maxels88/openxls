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

import java.util.ArrayList;
import java.util.HashMap;

public class StackedAreaChart extends AreaChart
{

	public StackedAreaChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		area = (Area) charttype;
	}

	@Override
	public boolean isStacked()
	{
		return true;
	}

	/**
	 * returns SVG to represent the actual chart object (BAR, LINE, AREA, etc.)
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
		double x = chartMetrics.get( "x" );
		double y = chartMetrics.get( "y" );
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double max = chartMetrics.get( "max" );
		double min = chartMetrics.get( "min" );
		Object[] categories = s.getCategories();
		ArrayList series = s.getSeriesValues();
		String[] seriescolors = s.getSeriesBarColors();
		String[] legends = s.getLegends();
		// x value for each point= w/(ncategories + 1) 1st one is xv*2 then increases from there
		// y value for each point= h/YMAX 
		if( series.size() == 0 )
		{
			Logger.logErr( "Area.getSVG: error in series" );
			return "";
		}
		StringBuffer svg = new StringBuffer();
		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)

		double xfactor = 0, yfactor = 0;    //
		if( categories.length > 1 )
		{
			xfactor = w / (categories.length - 1);    // w/#categories
		}
		if( max != 0 )
		{
			yfactor = h / max;    // h/YMAXSCALE
		}

		// first, calculate Area points- which are summed per series
		int n = series.size();
		int nSeries = ((double[]) series.get( 0 )).length;
		double[][] areapoints = new double[nSeries][n];
		for( int i = 0; i < n; i++ )
		{
			double[] seriesy = (double[]) series.get( i );
			for( int j = 0; j < seriesy.length; j++ )
			{
				double yval = seriesy[j];
				areapoints[j][i] = yval + ((i > 0) ? areapoints[j][i - 1] : 0);
			}
		}
		// for each series 
		for( int i = n - 1; i >= 0; i-- )
		{    // "paint" right to left
			svg.append( "<g>\r\n" );
			String points = "";
			double x1 = 0;
			double y1 = 0;
			String labels = null;
			double[] curseries = (double[]) series.get( i );
			for( int j = 0; j < curseries.length; j++ )
			{
				x1 = (x) + j * xfactor;
				double yval = areapoints[j][i];    // current point
				points += ((x) + ((j) * xfactor)) + "," + ((y + h) - (yval * yfactor));

				if( j == 0 )
				{
					y1 = ((y + h) - (yval * yfactor));    // end point (==start point) for path statement below
				}
				points += " ";
				// DATA LABELS
				String l = getSVGDataLabels( dls, axisMetrics, curseries[j], 0, i, legends, categories[j].toString() );
				if( l != null )
				{
					// if only category label, center over all series; anything else, position at data point
					boolean showCategories = (dls[i] & AttachedLabel.CATEGORYLABEL) == AttachedLabel.CATEGORYLABEL;
					boolean showValueLabel = (dls[i] & AttachedLabel.VALUELABEL) == AttachedLabel.VALUELABEL;
					boolean showValue = (dls[i] & AttachedLabel.VALUE) == AttachedLabel.VALUE;
					if( showCategories && !(showValue || showValueLabel) && (j == 0) )
					{    // only 1 label, centered along category axis within area
						//y1+= (seriesx[seriesx.length/2]/2)*yfactor;
						double hh = (areapoints[areapoints.length / 2][i] * yfactor);
						double yy = ((y + h) - hh) + 10;
						if( labels == null )
						{
							labels = "";
						}
						labels = "<text x='" + (x + (w / 2)) + "' y='" + yy + "' vertical-align='middle' " + this.getDataLabelFontSVG() + " style='text-align:middle;'>" + l + "</text>\r\n";
					}
					else if( showValue || showValueLabel )
					{ // labels at each data point
						if( labels == null )
						{
							labels = "";
						}
						double yy = (((y + h) - ((yval - (curseries[j] * .5)) * yfactor)));
						labels += "<text x='" + x1 + "' y='" + yy + "' style='text-anchor: middle;' " + this.getDataLabelFontSVG()/*+" fill='"+getDarkColor()+"'*/ + ">" + l + "</text>\r\n";
					}
				}
			}
			// pointsends connects up area to beginning
			double x0 = x;
			String pointsend = x1 + "," + (y + h) +
					" " + x0 + "," + (y + h) +
					" " + x0 + "," + y1;
			svg.append( "<polyline  id='series_" + (i + 1) + "' " + getScript( "" ) + " fill='" + seriescolors[i] + "' fill-opacity='1' " + getStrokeSVG() + " points='" + points + pointsend + "' fill-rule='evenodd'/>\r\n" );

/* john took out			
			
			// do twice to make slightly thicker
			svg.append("<polyline fill='none' fill-opacity='0' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" + 
					" points='" + points + "'/>\r\n");
			// do twice to make slightly thicker
			svg.append("<polyline fill='none' fill-opacity='0' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" + 
					" points='" + points + "'/>\r\n");*/
			// Now print data labels, if any
			if( labels != null )
			{
				svg.append( labels );
			}
			svg.append( "</g>\r\n" );
		}
		return svg.toString();
	}
}
