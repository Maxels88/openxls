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

public class StackedColumn extends ColChart
{
	int defaultShape = 0; //????

	public StackedColumn( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		col = (Bar) charttype;
	}

	@Override
	public boolean isStacked()
	{
		return true;
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
		if( series.size() == 0 )
		{
			Logger.logErr( "Bar.getSVG: error in series" );
			return "";
		}
		/*
		 * TODO: Stacked vs 100% Stacked		
		 */
		StringBuffer svg = new StringBuffer();
		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)

		double barw = 0, yfactor = 0;
		if( categories.length > 0 )
		{
			barw = w / (categories.length * 2);    // w/#categories (only 1 column per series)
		}
		if( max != 0 )
		{
			yfactor = h / max;    // h/YMAXSCALE
		}

		double[] totalperseries = null;
		boolean f100 = col.is100Percent();
		if( f100 )
		{  //FOR 100% STACKED
			// first, calculate points- which are summed per series
			int n = series.size();
			int nSeries = ((double[]) series.get( 0 )).length;
			totalperseries = new double[nSeries];
			for( int i = 0; i < n; i++ )
			{
				double[] seriesy = (double[]) series.get( i );
				for( int j = 0; j < seriesy.length; j++ )
				{
					double yval = seriesy[j];
					totalperseries[j] = yval + totalperseries[j];
				}
			}
		}
		// for each series - ONE COLUMN per series
		double[] previousY = new double[0];
		for( int i = 0; i < series.size(); i++ )
		{
			svg.append( "<g>\r\n" );
			double[] curseries = (double[]) series.get( i );    // for each data point - stacked on series column
			String[] curranges = (String[]) s.getSeriesRanges().get( i );
			double xx, yy = y + h;    // origin
			if( i == 0 )
			{
				previousY = new double[curseries.length];
				for( int j = 0; j < previousY.length; j++ )
				{
					previousY[j] = yy;    // origin
				}
			}
			double barh;
			for( int j = 0; j < curseries.length; j++ )
			{
				xx = (x) + (j * barw * 2) + barw / 2;
				if( previousY.length > j )    // should
				{
					yy = previousY[j];
				}
				if( f100 )
				{
					barh = ((curseries[j] / totalperseries[j]) * (h - 10));    // height of current point as a percentage of total points per series
				}
				else
				{
					barh = (curseries[j] * yfactor);    // height of current point
				}
				svg.append( "<rect " + getScript( curranges[j] ) + " fill='" + seriescolors[i] + "' fill-opacity='1' " + getStrokeSVG() +
						            " x='" + xx + "' y='" + (yy - barh) + "' width='" + barw + "' height='" + barh + "'/>" );
				//TODO: DATA LABELS
				// Now print data labels, if any
				//if (labels!=null)  svg.append(labels);
				if( previousY.length > j ) // should
				{
					previousY[j] = yy - barh;
				}
			}
			svg.append( "</g>\r\n" );
		}
		return svg.toString();
	}
}
