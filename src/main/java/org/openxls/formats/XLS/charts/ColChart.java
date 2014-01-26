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
package org.openxls.formats.XLS.charts;

import org.openxls.formats.XLS.WorkBook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashMap;

public class ColChart extends ChartType
{
	private static final Logger log = LoggerFactory.getLogger( ColChart.class );
	protected Bar col = null;

	public ColChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		col = (Bar) charttype;
		defaultShape = SHAPEDEFAULT;
	}

	/**
	 * @return truth of "Chart is Clustered"  (Bar/Col only)
	 */
	@Override
	public boolean isClustered()
	{
		return (/*cf.isClustered());	*/ (!isStacked() && !is100PercentStacked()));
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
			log.error( "Bar.getSVG: error in series" );
			return "";
		}
		int n = series.size();
		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)
		boolean isXReversed = (Boolean) axisMetrics.get( "xAxisReversed" );
		boolean isYReversed = (Boolean) axisMetrics.get( "yAxisReversed" );

		StringBuffer svg = new StringBuffer();        // return svg
		svg.append( "<g>\r\n" );

		double barw = 0;    //
		double yfactor = 0;
		if( n > 0 )
		{
			barw = w / ((categories.length * (n + 1.0)) + 1);    // w/(#cats * nseries+1) + 1
			if( max != 0 )
			{
				yfactor = h / max;    // h/YMAXSCALE
			}
		}
		int rfX = (!isXReversed ? 1 : -1); // reverse factor
		int rfY = (!isYReversed ? 1 : -1);
		// for each series 
		for( int i = 0; i < n; i++ )
		{    // each series group
			svg.append( "<g>\r\n" );
			double y0 = y + (!isYReversed ? h : 0);    // start from bottom and work up (unless reversed)
			double[] curseries = (double[]) series.get( i );
			String[] curranges = (String[]) s.getSeriesRanges().get( i );

			//x+=barw;	// a barwidth separates each series group
			for( int j = 0; j < curseries.length; j++ )
			{        // each series
				double xx = x + (barw * (i + 1)) + (j * (n + 1) * barw);        // x goes from 1 series to next, corresponding to bar/column color
				double hh = yfactor * curseries[j];                        // bar height = measure of series value
				double yy = y0 - (!isYReversed ? hh : 0);                // start drawing column
				svg.append( "<rect id='series_" + (i + 1) + "' " + getScript( curranges[j] ) + " fill='" + seriescolors[i] + "' fill-opacity='1' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" +
						            " x='" + xx + "' y='" + yy + "' width='" + barw + "' height='" + hh + "' fill-rule='evenodd'/>" );
				String l = getSVGDataLabels( dls, axisMetrics, curseries[j], 0, i, legends, categories[j].toString() );
				if( l != null )
				{
					svg.append( "<text x='" + (xx + (barw / 2)) + "' y='" + (y0 - ((hh + 10) * rfY)) +
							            "' style='text-anchor: middle;' " + getDataLabelFontSVG() + ">" + l + "</text>\r\n" );
				}
			}
			svg.append( "</g>\r\n" );        // each series group
		}
		svg.append( "</g>\r\n" );
		return svg.toString();
	}

	/**
	 * gets the chart-type specific ooxml representation: <colChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:barChart>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:barDir val=\"col\"/>" );
		cooxml.append( "<c:grouping val=\"" );

		if( is100PercentStacked() )
		{
			cooxml.append( "percentStacked" );
		}
		else if( isStacked() )
		{
			cooxml.append( "stacked" );
		}
		else if( isClustered() )
		{
			cooxml.append( "clustered" );
		}
		else
		{
			cooxml.append( "standard" );
		}
		cooxml.append( "\"/>" );
		cooxml.append( "\r\n" );
		// vary colors???

		// *** Series Data:	ser, cat, val for most chart types
		cooxml.append( getParentChart().getChartSeries().getOOXML( getChartType(), false, 0 ) );

		// chart data labels, if any
		//TODO: FINISH		    	
		//cooxml.append(getDataLabelsOOXML(cf));
		//TODO: FINISH DR0P LINES    			
		///    			if (this.hasDropLines() )	/* Area, Line, Stock*/
		//				cooxml.append()
		if( !getChartOption( "Gap" ).equals( "150" ) )
		{
			cooxml.append( "<c:gapWidth val=\"" + getChartOption( "Gap" ) + "\"/>" );    // default= 0
		}
		if( !getChartOption( "Overlap" ).equals( "0" ) )
		{
			cooxml.append( "<c:overlap val=\"" + getChartOption( "Overlap" ) + "\"/>" );    // default= 0
		}
		// Series Lines	
		ChartLine cl = cf.getChartLinesRec();
		if( cl != null )
		{
			cooxml.append( cl.getOOXML() );
		}

		// axis ids	 - unsigned int strings
		cooxml.append( "<c:axId val=\"" + catAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:axId val=\"" + valAxisId + "\"/>" );
		cooxml.append( "\r\n" );

		cooxml.append( "</c:barChart>" );
		cooxml.append( "\r\n" );

		return cooxml;
	}
}
