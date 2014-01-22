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

public class BubbleChart extends ChartType
{

	private Scatter bubble = null;
	private boolean is3d = false;

	public BubbleChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		bubble = (Scatter) charttype;
	}

	/**
	 * Bubble charts handle 3d differently
	 *
	 * @param is3d
	 */
	public void setIs3d( boolean is3d )
	{
		this.is3d = is3d;
	}

	/**
	 * return true if this bubble chart series should be displayed as 3d
	 *
	 * @return
	 */
	public boolean is3d()
	{
		return is3d;
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
		StringBuffer svg = new StringBuffer();

		if( series.size() == 0 )
		{
			Logger.logErr( "Scatter.getSVG: error in series" );
			return "";
		}
		// gather data labels, markers, has lines, series colors for chart
		boolean threeD = cf.isThreeD( BUBBLECHART );
		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)
		boolean hasLines = cf.getHasLines(); // should be per-series?
		int[] markers = getMarkerFormats();    // get an array of marker formats per series
		int n = series.size();
		double[] seriesx = null;
		double xfactor = 0, yfactor = 0, bfactor = 0;    //
		boolean TEXTUALXAXIS = true;
		// get x axis max/min for an x axis which is a value axis
		seriesx = new double[categories.length];
		double xmin = Double.MAX_VALUE, xmax = Double.MIN_VALUE;
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

		if( (max - min) != 0 )
		{
			yfactor = h / Math.abs( max - min );    // h/YMAXSCALE
		}
		// get Bubble Size Factor
		for( int i = 0; i < n; i++ )
		{
			double[] seriesy = (double[]) series.get( i );
			String[] curranges = (String[]) s.getSeriesRanges().get( i );
			int nseries = seriesy.length / 2;
			double bmin = Double.MAX_VALUE, bmax = Double.MIN_VALUE;
			for( int j = nseries; j < seriesy.length; j++ )
			{
				bmax = Math.max( bmax, seriesy[j] );
				bmin = Math.min( bmin, seriesy[j] );
			}
			if( (bmax - bmin) != 0 )
			{
				bfactor = h / Math.abs( bmax - bmin ) / 5;
			}
		}
		svg.append( "<g>\r\n" );
		if( threeD )
		{
			svg.append( get3DBubbleSVG( seriescolors ) );
		}
		// for each series 
		for( int i = 0; i < n; i++ )
		{
			String labels = "";
			double[] seriesy = (double[]) series.get( i );
			String[] curranges = (String[]) s.getSeriesRanges().get( i );
			int nseries = seriesy.length / 2;
			for( int j = 0; j < nseries; j++ )
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
				double cx = x + xval * xfactor;
				double cy = (y + h) - (seriesy[j] * yfactor);
				double r = seriesy[j + nseries] * bfactor;
//System.out.println("x: " + xval + " val: " + seriesy[j] + " size: " + seriesy[j+nseries]);				
				if( !threeD )
				{
					svg.append( "<circle " + getScript( curranges[j] ) + " cx='" + cx + "' cy='" + cy + "' r='" + r + "' " + getStrokeSVG() + " fill='" + seriescolors[i] + "'/>\r\n" );
				}
				else
				{
					svg.append( "<circle " + getScript( curranges[j] ) + "   id='series_" + (i + 1) + "' cx='" + cx + "' cy='" + cy + "' r='" + r + "' " + getStrokeSVG() + " style='fill:url(#fill" + i + ")'/>\r\n" );
				}
				String l = getSVGDataLabels( dls, axisMetrics, seriesy[j + nseries], 0, i, legends, categories[j].toString() );
				if( l != null )
				{
					labels += "<text x='" + (r + 10 + (x) + xval * xfactor) + "' y='" + (((y + h) - (seriesy[j] * yfactor))) + "' " + this.getDataLabelFontSVG() + ">" + l + "</text>\r\n";
				}
			}
			// labels after lines and markers  
			svg.append( labels );
		}
//System.out.println("Bubble svg: " + svg.toString());				
		svg.append( "</g>\r\n" );
		return svg.toString();
	}

	/**
	 * returns the SVG necessary to define 3D bubbles (circles) for each series color
	 *
	 * @param seriescolors
	 * @return String SVG
	 */
	private String get3DBubbleSVG( String[] seriescolors )
	{
		StringBuffer svg = new StringBuffer();
		svg.append( "<defs>\r\n" );
		for( int i = 0; i < seriescolors.length; i++ )
		{
			svg.append( "<radialGradient id='fill" + i + "' " +
					            "gradientUnits=\"objectBoundingBox\" fx=\"40%\" fy=\"30%\">" );
			svg.append( "<stop offset='0%' style='stop-color:#FFFFFF' />" );
			svg.append( "<stop offset='40%' style='stop-color:" + seriescolors[i] + "' stop-opacity='.65' />" );
			svg.append( "<stop offset='95%' style='stop-color:" + seriescolors[i] + "' stop-opacity='1' />" );
			svg.append( "<stop offset='99%' style='stop-color:" + seriescolors[i] + "' stop-opacity='.3' />" );
			svg.append( "<stop offset='100%' style='stop-color:" + seriescolors[i] + "'/>" );
			svg.append( "</radialGradient>\r\n" );
		}
		svg.append( "</defs>\r\n" );
		return svg.toString();
	}

	/**
	 * gets the chart-type specific ooxml representation: <bubbleChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:bubbleChart>" );
		cooxml.append( "\r\n" );
		// vary colors???

		// *** Series Data:	ser, cat, val for most chart types
		cooxml.append( this.getParentChart().getChartSeries().getOOXML( this.getChartType(), false, 0 ) );

		// chart data labels, if any
		//TODO: FINISH		    	
		//cooxml.append(getDataLabelsOOXML(cf));		
		if( this.is3d )
		{
			cooxml.append( "bubble3d val=\"1\"" );
		}
		// bubblescale
		cooxml.append( "<c:bubbleScale val=\"100\"/>" );    // TODO: read correct value
		// showNegBubbles
		// sizeRepresents

		// axis ids	 - unsigned int strings
		cooxml.append( "<c:axId val=\"" + catAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:axId val=\"" + valAxisId + "\"/>" );
		cooxml.append( "\r\n" );

		cooxml.append( "</c:bubbleChart>" );
		cooxml.append( "\r\n" );

		return cooxml;
	}

}