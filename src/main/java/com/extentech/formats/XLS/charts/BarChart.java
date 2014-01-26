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

import com.extentech.ExtenXLS.ChartSeriesHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.XLS.WorkBook;
import org.json.JSONException;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashMap;

public class BarChart extends ChartType
{
	private static final Logger log = LoggerFactory.getLogger( BarChart.class );
	protected Bar bar = null;

	public BarChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		bar = (Bar) charttype;
		defaultShape = SHAPEDEFAULT;
	}

	/**
	 * set the default shape for all bars in this chart
	 *
	 * @param shape
	 */
	public void setDefaultShape( int shape )
	{
		defaultShape = shape;
	}

	/**
	 * return the default shape for all bars in this chart
	 *
	 * @return
	 */
	public int getDefaultShape()
	{
		return defaultShape;
	}

	/**
	 * @return truth of "Chart is Clustered"  (Bar/Col only)
	 */
	@Override
	public boolean isClustered()
	{
		return (/*cf.isClustered() || */(!isStacked() && !is100PercentStacked()));
	}

	/**
	 * returns the bar shape for a column or bar type chart
	 * can be one of:
	 * <br>SHAPECOLUMN	default
	 * <br>SHAPECONEd
	 * <br>SHAPECONETOMAX
	 * <br>SHAPECYLINDER
	 * <br>SHAPEPYRAMID
	 * <br>SHAPEPYRAMIDTOMAX
	 *
	 * @return int bar shape
	 */
	@Override
	public int getBarShape()
	{
		return cf.getBarShape();
	}

	/**
	 * @return String XML representation of this chart-type's options
	 */
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		if( bar.fStacked )
		{
			sb.append( " Stacked=\"true\"" );
		}
		if( bar.f100 )
		{
			sb.append( " PercentageDisplay=\"true\"" );
		}
		if( bar.fHasShadow )
		{
			sb.append( " Shadow=\"true\"" );
		}
		if( bar.pcOverlap != 0 )
		{
			sb.append( " Overlap=\"" + bar.pcOverlap + "\"" );
		}
		if( (bar.pcGap != 50) && (bar.pcGap != 150) )
		{
			sb.append( " Gap=\"" + bar.pcGap + "\"" );
		}
		return sb.toString();
	}

	@Override
	public JSONObject getTypeJSON() throws JSONException
	{
		JSONObject typeJSON = new JSONObject();
		// Dojo: Regular, Stacked, Clustered
/*			String dojoType;
	    	int shape=ChartConstants.SHAPECOLUMN;
	    	//controls Data Legends, % Distance of sections, Line Format, Area Format, Bar Shapes ...
	    	ChartFormat cf= this.getParentChart().getChartFormat();
	    	DataFormat df= cf.getDataFormatRec(false);
	    	
	    	if (df!=null)
	    		shape= df.getShape();
			if (shape==ChartConstants.SHAPECONE ||
				shape==ChartConstants.SHAPECYLINDER ||
				shape==ChartConstants.SHAPEPYRAMID)
				;	// TODO: SOMETHING!!! Interpret Shape for Pyramid, etc. charts
			if (this.chartType==ChartConstants.BAR) {		
				if (!cf.isThreeD(ChartConstants.BAR)) {
					dojoType="ClusteredBars";	// clustered dojo is actually Excel normal
					typeJSON.put("gap", 5);		//default val that looks pretty good ...
				}
				else if (cf.isClustered())
					dojoType="Bars";			// = 
				else
					dojoType="StackedBars";
			} else { 	// Column
				if (!cf.isThreeD(ChartConstants.COL)) {
					dojoType="ClusteredColumns";	// clustered dojo is actually Excel normal col	
					typeJSON.put("gap", 5);			//default val that looks pretty good ...
				}
				else if (cf.isClustered())
					dojoType="Columns";
				else
					dojoType="StackedColumns";
			}
	    	typeJSON.put("type", dojoType);
*/
		return typeJSON;
	}

	/**
	 * replaces generic getJSON with Bar/Column specifics (necessary for stacked-type charts)
	 */
	public static JSONObject getJSON( ChartSeriesHandle[] series, WorkBookHandle wbh, Double[] minMax ) throws JSONException
	{
		JSONObject chartObjectJSON = new JSONObject();

/*	    	// 20080428 KSC: get stacked state     	
	    	boolean bIsStacked= (cf.isThreeD() && !cf.isClustered());
	    	// Type JSON
	    	chartObjectJSON.put("type", this.getTypeJSON());
	    	
	    	// Deal with Series 
			double yMax= 0.0, yMin= 0.0;
			int nSeries= 0;
			// 20080428 KSC: stacked charts max is sum of all points
			double[] ySum= null;
			if (bIsStacked) {	
				// preprocess series to get # series values
		        for (int i= 0; i < series.length; i++) {
					CellRange cr = new CellRange(series[i].getSeriesRange(), wbh, true);
		        	nSeries= Math.max(nSeries, cr.getCells().length);
		        }
				ySum= new double[nSeries];		
			}
	        JSONArray seriesJSON= new JSONArray();
	        JSONArray seriesCOLORS= new JSONArray();
	    	try {
		        for (int i= 0; i < series.length; i++) {
		        	try{
			        	JSONArray seriesvals= CellRange.getJSON(series[i].getSeriesRange(), wbh);
			        	// must trap min and max for axis tick and units
			        	nSeries= Math.max(nSeries, seriesvals.length());
			        	for (int j= 0; j < seriesvals.length(); j++) {
			        		try {
			        			if (!bIsStacked)
			        				yMax= Math.max(yMax, seriesvals.getDouble(j));
			        			else 
			        				ySum[j]+= seriesvals.getDouble(j);		
			        			yMin= Math.min(yMin, seriesvals.getDouble(j));
			        		} catch (NumberFormatException n) {;}
			        	}
			        	seriesJSON.put(seriesvals);
			        	seriesCOLORS.put(FormatConstants.SVGCOLORSTRINGS[series[i].getSeriesColor()]);
		        	}catch(Exception x){
		        		;
		        	}
		        }
	        	if (bIsStacked) {
	        		for (int i= 0; i < nSeries; i++)
	        			yMax= Math.max(yMax, ySum[i]);
	        	}
		    	chartObjectJSON.put("Series", seriesJSON);
		    	chartObjectJSON.put("SeriesFills", seriesCOLORS);
	    	} catch (JSONException je) {
	    		// TODO: Log error
	    	}
	    	minMax[0]= new Double(yMin);
	    	minMax[1]= new Double(yMax);
	    	minMax[2]= new Double(nSeries);
*/
		return chartObjectJSON;
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
		if( series.size() == 0 )
		{
			log.error( "Bar.getSVG: error in series" );
			return "";
		}
		StringBuffer svg = new StringBuffer();
		int shape = cf.getBarShape();
		if( (shape == ChartConstants.SHAPECONE) ||
				(shape == ChartConstants.SHAPECYLINDER) ||
				(shape == ChartConstants.SHAPEPYRAMID) )
		{
			;    // TODO: SOMETHING!!! Interpret Shape for Pyramid, etc. charts
		}

		int[] dls = getDataLabelInts(); // get array of data labels (can be specific per series ...)
		int n = series.size();
		svg.append( "<g>\r\n" );

		double barw = 0;    //
		double yfactor = 0;
		barw = h / ((categories.length + 1.0) * (n + 1));    // bar width= height/total number of bars+separators
		if( max != 0 )
		{
			yfactor = w / max;        // w/yMax= scale unit
		}

		int rf = 1;    //(!yAxisReversed?1:-1);
		double y0 = (h + y) - (barw * 0.5);    //start at bottom and work up (unless reversed)
		// for each series
		for( int i = 0; i < n; i++ )
		{    // each series group i.e. each bar with a different color
			svg.append( "<g>\r\n" );
			double[] curseries = (double[]) series.get( i );
			String[] curranges = (String[]) s.getSeriesRanges().get( i );
			for( int j = 0; j < curseries.length; j++ )
			{        // each category (group of series)
//						double y= y0 - barw*(i+1)*rf - (j*n*barw*1.5)*rf;		// y goes from 1 series to next, corresponding to bar/column color
				double yy = y0 - (barw * (n + 1) * 1.1 * (j)) - (barw * (i + 1)) - ((barw * 1.5) * rf) /* start */;
				double ww = yfactor * curseries[j];                            // width of bar -- measure of series
				svg.append( "<rect " + getScript( curranges[j] ) + " fill='" + seriescolors[i] + "' fill-opacity='1' stroke='black' stroke-opacity='1' stroke-width='1' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'" +
						            " x='" + x + "' y='" + yy + "' width='" + ww + "' height='" + barw + "' fill-rule='evenodd'/>" );
				String l = getSVGDataLabels( dls, axisMetrics, curseries[j], 0, i, legends, categories[j].toString() );
				if( l != null )
				{
					svg.append( "<text x='" + (x + ww + 10) + "' y='" + (yy + (barw * rf)) +
							            "' " + getDataLabelFontSVG() + ">" + l + "</text>\r\n" );
				}
			}
			svg.append( "</g>\r\n" );        // each series group
		}
		svg.append( "</g>\r\n" );
		return svg.toString();
	}

	/**
	 * gets the chart-type specific ooxml representation: <barChart>
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
		cooxml.append( "<c:barDir val=\"bar\"/>" );
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
