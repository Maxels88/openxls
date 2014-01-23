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

public class PieChart extends ChartType
{
	private static final Logger log = LoggerFactory.getLogger( PieChart.class );
	Pie pie = null;

	public PieChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		pie = (Pie) charttype;
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
		chartObjectJSON.put( "type", this.getTypeJSON() );

		// TODO: check out labels options chosen: default label is Category name + percentage ...
		double yMax = 0.0, yMin = 0.0;
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
		/*
		DataFormat df= this.getParentChart().getChartFormat().getDataFormatRec();
    	for (int z= 0; z < df.chartArr.size(); z++) {
    		// LineFormat = color, ticks.., AreaFormat= color/patterns ..., PieFormat= distance as percentage, MarkerFormat, AttachedLabel
            BiffRec b = (BiffRec)df.chartArr.get(z);
            if (b instanceof PieFormat) {
            	//if (((PieFormat) b).getPercentage()!=0)
            	//	typeJSON.put("percentage", ((PieFormat) b).getPercentage());
            	// TODO: Convert Excel percentage to Dojo labelOffset
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
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		Object[] categories = s.getCategories();
		ArrayList series = s.getSeriesValues();
		String[] seriescolors = s.getSeriesBarColors();
		String[] legends = s.getLegends();
		StringBuffer svg = new StringBuffer();
		// if have any series - should :)
		if( series.size() == 0 )
		{
			log.error( "Pie.getSVG: error in series" );
			return "";
		}
		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)
		final int LABELOFFSET = 15;

		int n = series.size();
		double centerx = 0.0, centery = 0.0, radius = 0.0;
		double radiusy = 0.0;
		centerx = w / 2 + chartMetrics.get( "x" );
		centery = h / 2 + chartMetrics.get( "y" );
		svg.append( "<g>\r\n" );

		radius = (Math.min( w, h )) / 2;
		//radius= Math.min(w, h)/2;	    // too small
		//radius= Math.min(w, h);		// ????????????? just an estimate?
		radiusy = radius;
		if( n > 0 )
		{
			double[] oneseries = (double[]) series.get( 0 );    //FOR PIE CHARTS ONLY 1 SERIES VALUE POINTS ARE USED **********************************
			String[] curranges = (String[]) s.getSeriesRanges().get( 0 );
			double total = 0.0;
			for( double onesery : oneseries )
			{    // get total in order to calculate percentages
				total += onesery;
			}
			if( dls.length == 1 )
			{    // no series-specific data labels; expand to entire series for loop below
				int dl = dls[0];
				dls = new int[oneseries.length];
				for( int i = 0; i < dls.length; i++ )
				{
					dls[i] = dl;
				}
			}
			String path = "";
			double x = centerx + radius;
			double y = centery;
			double percentage = 0;
			double lasta = 0;
			int largearcflag = 0;
			int sweepflag = 0;

			// Now create each pie wedge according to it's percentage value 
			for( int j = 0; j < oneseries.length; j++ )
			{
				if( total > 0 )
				{
					percentage = oneseries[j] / total;
				}
				double angle = (percentage * 360) + lasta;
				double x1 = centerx + (radius * (Math.cos( Math.toRadians( angle ) )));
				double y1 = centery - (radiusy * (Math.sin( Math.toRadians( angle ) )));
				if( (percentage * 360)/*angle*/ > 180 )
				{
					largearcflag = 1;
				}
				else
				{
					largearcflag = 0;
				}
				path = "M" + centerx + " " + centery + " L" + x + " " + y + " A" + radius + " " + radiusy + " 0 " + largearcflag + " " + sweepflag + " " + x1 + " " + y1 + " L" + centerx + " " + centery + "Z";
				svg.append( "<path " + getScript( curranges[j] ) + "   id='series_" + (j + 1) + "'  fill='" + seriescolors[j] + "' fill-opacity='1' " + getStrokeSVG() + " path='' d='" + path + "' fill-rule='evenodd'/>\r\n" );
//						" path='' d='" + path + "' fill-rule='evenodd' transform='matrix(1, 0, 0, 1, 0, 0)'/>\r\n");

				String l = getSVGDataLabels( dls, axisMetrics, oneseries[j], percentage, j, legends, categories[j].toString() );
				if( l != null )
				{
					// apparently labels are outside of wedge unless angle is >= 30 ...
					// category labels
					double halfa = ((percentage / 2) * 360) + lasta;    // center in area
					double x2, y2;
					if( percentage < .3 )
					{    // display label on outside with leader lines
						x2 = centerx + (radius + LABELOFFSET) * (Math.cos( Math.toRadians( halfa ) ));
						y2 = centery - (radiusy + LABELOFFSET) * (Math.sin( Math.toRadians( halfa ) ));
					}
					else
					{    // display label within wedge for > 30%
						x2 = centerx + (radius / 2) * (Math.cos( Math.toRadians( halfa ) ));
						y2 = centery - (radiusy / 2) * (Math.sin( Math.toRadians( halfa ) ));
					}
					String style = "";
					if( percentage >= .3 )
					{
						style = " style='text-anchor: middle;'";
					}
					else if( (lasta > 90) && (lasta < 270) )
					{    // right-align text for wedges on left side of pie
						style = " style='text-anchor: end;'";
						// TODO: dec x2 
					}
					svg.append( "<text x='" + (x2) + "' y='" + (y2) + "' vertical-align='bottom' " + this.getDataLabelFontSVG() + " " + style + ">" + l + "</text>\r\n" );
					// leaderline - not exactly like Excel's but ... :) do when NOT putting text within wedge
					if( percentage < .3 )
					{
						double x0 = centerx + ((radius) * (Math.cos( Math.toRadians( halfa ) )));
						double y0 = centery - ((radiusy) * (Math.sin( Math.toRadians( halfa ) )));
						svg.append( "<line " + getScript( curranges[j] ) + " x1='" + x0 + "' y1 ='" + y0 + "' x2='" + (x2 - 3) + "' y2='" + (y2 - 3) + "'" + getStrokeSVG() + "/>\r\n" );
					}
				}
				lasta = angle;
				x = x1;
				y = y1;
			}
		}
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
	 * gets the chart-type specific ooxml representation: <pieChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:pieChart>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:varyColors val=\"1\"/>" );

		// *** Series Data:	ser, cat, val for most chart types
		cooxml.append( this.getParentChart().getChartSeries().getOOXML( this.getChartType(), false, 0 ) );

		// chart data labels, if any
		//TODO: FINISH		    	
		//cooxml.append(getDataLabelsOOXML(cf));

		// TODO: firstSLiceAng
//		if (pie.getAnStart()!=0)
		cooxml.append( "<c:firstSliceAng val=\"" + pie.getAnStart() + "\"/>" );
		cooxml.append( "</c:pieChart>" );
		cooxml.append( "\r\n" );

		return cooxml;
	}

}
