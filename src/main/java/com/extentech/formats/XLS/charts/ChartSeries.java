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

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.DLbls;
import com.extentech.formats.OOXML.DPt;
import com.extentech.formats.OOXML.Marker;
import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.OOXMLAdapter;
import com.extentech.formats.XLS.WorkBookException;
import com.extentech.formats.XLS.XLSConstants;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.formats.XLS.Xf;
import com.extentech.formats.XLS.formulas.Ptg;
import com.extentech.toolkit.CompatibleVector;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Stack;
import java.util.Vector;

public class ChartSeries implements ChartConstants, Serializable
{
	private static final Logger log = LoggerFactory.getLogger( ChartSeries.class );
	private static final long serialVersionUID = -7862828186455339066L;
	private ArrayList<Object[]> series = new ArrayList();    // ALL series in chart MAPPED to chart that "owns" it
	private JSONArray seriesJSON = null;            // save series JSON for update comparisons later
	protected double[] minmaxcache = null;    // stores minimum/maximum/minor/major for chart scale; cached;
	protected String[] legends;
	protected ArrayList seriesranges;
	protected Object[] categories;
	protected ArrayList seriesvalues;
	protected String[] seriescolors;
	protected Chart parentChart;
//	protected transient WorkBookHandle wbh;

	/**
	 * series stores new Object[] {Series Record, Integer.valueOf(nCharts)}
	 *
	 * @param o
	 */
	public void add( Object[] o )
	{
		series.add( o );
	}

	public void setParentChart( Chart c )
	{
		parentChart = c;
	}
//	public void setWorkBook(WorkBookHandle wbh) { this.wbh= wbh; }

	/**
	 * returns an array containing the and maximum values of Y (Value) axis, along with the
	 * maximum number of series values of each series (bar, line, etc.) ...
	 * <br>sets all series cache values:  seriesvalues, seriesranges, seriescolors, legends and minmaxcache
	 * <p/>
	 * <br>Note: this will only reset cached values if isDirty
	 *
	 * @return double[]
	 */
	public double[] getMetrics( boolean isDirty )
	{
		if( !isDirty && (minmaxcache != null) )
		{
			return minmaxcache;
		}
		// trap minimum, maximum + number of series
		ChartType co = parentChart.getChartObject();    // default chart object TODO: overlay charts
		seriesvalues = new ArrayList();
		seriesranges = new ArrayList();
		com.extentech.formats.XLS.Boundsheet sht = parentChart.getSheet();
		java.util.Vector s = getAllSeries( -1 );
		// Category values *******************************************************************************
		if( s.size() > 0 )
		{
			try
			{
				CellRange cr = new CellRange( ((Series) s.get( 0 )).getCategoryValueAi().toString(), parentChart.wbh, true );
				CellHandle[] ch = cr.getCells();
				if( ch != null )
				{ // found a template with a chart with series but no categories
					categories = new Object[ch.length];
					for( int j = 0; j < ch.length; j++ )
					{
						try
						{
							categories[j] = ch[j].getFormattedStringVal( true );
						}
						catch( IllegalArgumentException e )
						{ // catch format exceptions
							categories[j] = ch[j].getStringVal();
						}
					}
				}
				else if( s.size() > 0 )
				{
					cr = new CellRange( ((Series) s.get( 0 )).getSeriesValueAi().toString(), parentChart.wbh, true );
					int sz = cr.getCells().length;
					categories = new Object[sz];
					for( int j = 0; j < sz; j++ )
					{
						categories[j] = Integer.valueOf( j + 1 ).toString();
					}
				}
			}
			catch( Exception e )
			{
				log.error( "ChartSeries.getMinMax: " + e.toString() );
			}
		}
		// Series colors, labels and values ***************************************************************
		double yMax = 0.0;
		double yMin = Double.MAX_VALUE;
		int nseries = 0;
		seriescolors = null;
		legends = null;
		int charttype = co.getChartType();
		// obtain/store series colors, store series values and trap maximum and
		// minimun values so can be used below for axis scale
		/*
		 * A Scatter chart has two value axes, showing one set of numerical data along the x-axis and another along the y-axis.
    	 * It combines these values into single data points and displays them in uneven intervals, or clusters		
    	 */
		if( (charttype != PIECHART) && (charttype != DOUGHNUTCHART) )
		{
			seriescolors = new String[s.size()];
			legends = new String[s.size()];
			for( int i = 0; i < s.size(); i++ )
			{
				Series myseries = ((Series) s.get( i ));
				seriescolors[i] = myseries.getSeriesColor();
				legends[i] = com.extentech.formats.XLS.OOXMLAdapter.stripNonAscii( myseries.getLegendText() ).toString();
				CellRange cr = new CellRange( myseries.getSeriesValueAi().toString(), parentChart.wbh, true );
				CellHandle[] ch = cr.getCells();
				nseries = Math.max( nseries, ch.length );
				double[] seriesvals;
				String[] sranges;
				//				String[] series_strings;
				if( !myseries.hasBubbleSizes() )
				{
					seriesvals = new double[nseries];
					sranges = new String[nseries];
				}
				else
				{
					seriesvals = new double[nseries * 2];
					sranges = new String[nseries * 2];
				}
				//				series_strings = new String[seriesvals.length];

				for( int j = 0; j < ch.length; j++ )
				{
					try
					{
						sranges[j] = ch[j].getCellAddressWithSheet();
						seriesvals[j] = ch[j].getDoubleVal();
						if( Double.isNaN( seriesvals[j] ) )
						{
							seriesvals[j] = 0.0;
						}
						yMax = Math.max( yMax, seriesvals[j] );
						yMin = Math.min( yMin, seriesvals[j] );
					}
					catch( NumberFormatException n )
					{
						;
					}
				}
				if( myseries.hasBubbleSizes() )
				{ // append bubble sizes to series values ... see BubbleChart.getSVG for parsing
					int z = ch.length;
					CellRange crb = new CellRange( myseries.getBubbleValueAi().toString(), parentChart.wbh, true );
					CellHandle[] chb = crb.getCells();
					for( int j = 0; j < ch.length; j++ )
					{
						seriesvals[j + z] = chb[j].getDoubleVal();
						sranges[j + z] = chb[j].getCellAddressWithSheet();
					}
				}
				seriesvalues.add( seriesvals );   // trap and add series value points
				seriesranges.add( sranges );        // trap series range
			}
		}
		else if( (charttype == DOUGHNUTCHART) && (s.size() > 1) )
		{ // like a PIE chart but can have multiple series
			legends = new String[categories.length];        // for PIE/DONUT charts, legends are actually category labels, not series labels
			for( int i = 0; i < categories.length; i++ )
			{
				legends[i] = com.extentech.formats.XLS.OOXMLAdapter.stripNonAscii( categories[i].toString() ).toString();
			}
			for( int i = 0; i < s.size(); i++ )
			{
				Series myseries = ((Series) s.get( i ));
				// legends[i]=
				// com.extentech.formats.XLS.OOXMLAdapter.stripNonAscii(myseries.getLegend());
				CellRange cr = new CellRange( myseries.getSeriesValueAi().toString(), parentChart.wbh, true );
				CellHandle[] ch = cr.getCells();
				double[] seriesvals = new double[ch.length];
				String[] sranges = new String[ch.length];
				if( seriescolors == null )
				{
					seriescolors = new String[ch.length];
				}
				for( int j = 0; j < ch.length; j++ )
				{
					try
					{
						seriesvals[j] = ch[j].getDoubleVal();
						if( ch[j].getWorkSheetHandle().getMysheet().equals( sht ) )
						{
							sranges[j] = ch[j].getCellAddress();
						}
						yMax = Math.max( yMax, seriesvals[j] );
						yMin = Math.min( yMin, seriesvals[j] );
						if( i == 0 )
						{ // only do for 1st series; will be the
							// same for rest
							seriescolors[j] = myseries.getPieSliceColor( j );
    			    /*if (seriescolors[j] == 0x4D
									|| seriescolors[j] == 0x4E)
								seriescolors[j] = com.extentech.formats.XLS.FormatConstants.COLOR_WHITE;*/
						}

					}
					catch( NumberFormatException n )
					{
						;
					}
				}
				seriesvalues.add( seriesvals ); // trap and add series value points
				seriesranges.add( sranges );        // trap series range
			}
		}
		else
		{ // PIES - only 1 series
			if( s.size() > 0 )
			{
				// PIE: 1 series data
				CellHandle[] cats = new CellRange( ((Series) s.get( 0 )).getCategoryValueAi().toString(),
				                                   parentChart.wbh,
				                                   true ).getCells();
				if( cats != null )
				{
					nseries = cats.length;
					legends = new String[cats.length]; // for PIE charts, legends are actually category labels, not series labels
					for( int i = 0; i < cats.length; i++ )
					{
						legends[i] = cats[i].getFormattedStringVal( true );
					}
				}
				seriescolors = new String[nseries];
				Series myseries = ((Series) s.get( 0 ));
				try
				{
					CellRange cr = new CellRange( myseries.getSeriesValueAi().toString(), parentChart.wbh, true );
					CellHandle[] ch = cr.getCells();
					// error trap - shouldn't happen
					if( ch.length != nseries )
					{
						log.warn( "ChartHandle.getSeriesInfo: unexpected Pie Chart structure" );
						nseries = Math.min( nseries, ch.length );
					}
					double[] seriesvals = new double[nseries];
					String[] sranges = new String[nseries];
					for( int i = 0; i < nseries; i++ )
					{
						seriescolors[i] = myseries.getPieSliceColor( i );
    			/*if (seriescolors[i] == 0x4D || seriescolors[i] == 0x4E)
							seriescolors[i] = com.extentech.formats.XLS.FormatConstants.COLOR_WHITE;*/
						// legends[i]=
						// com.extentech.formats.XLS.OOXMLAdapter.stripNonAscii(myseries.getLegend());
						// // same for every series ...
						seriesvals[i] = ch[i].getDoubleVal();
						if( ch[i].getWorkSheetHandle().getMysheet().equals( sht ) )
						{
							sranges[i] = ch[i].getCellAddress();
						}
						yMax = Math.max( yMax, seriesvals[i] );
						yMin = Math.min( yMin, seriesvals[i] );
					}
					seriesvalues.add( seriesvals ); // trap and add series value points
					seriesranges.add( sranges );        // trap series range
				}
				catch( IllegalArgumentException e )
				{
					; // error in cell range sheet ...
				}
			}
		}
		// For stacked-type charts, must sum values for ymax
		if( co.isStacked() )
		{
			// scale is SUM of values, yMax is maximum total per series point
			double[] sum = new double[nseries];
			for( Object seriesvalue : seriesvalues )
			{
				double[] seriesv = (double[]) seriesvalue;
				for( int j = 0; j < seriesv.length; j++ )
				{
					sum[j] = sum[j] + seriesv[j];
				}
			}
			yMax = 0;
			for( int i = 0; i < nseries; i++ )
			{
				yMax = Math.max( sum[i], yMax );
			}
		}
		minmaxcache = new double[2];
		minmaxcache[0] = yMin;
		minmaxcache[1] = yMax;
		//    	minmaxcache[2]= nSeries;
		//    	minmaxcache= new double[3][];	// 3 possible axes: x, y and z
		//    	minmaxcache[axisType]= minMax;
		return minmaxcache;
	}

	/**
	 * Change series ranges for ALL matching series
	 *
	 * @param originalrange
	 * @param newrange
	 * @return
	 */
	public boolean changeSeriesRange( String originalrange, String newrange )
	{
		boolean changed = false;
		for( Object[] sery : series )
		{
			Series s = (Series) sery[0];
			Ai ai = s.getSeriesValueAi();
			if( ai != null )
			{
				if( ai.toString().equalsIgnoreCase( originalrange ) )
				{
					changed = ai.changeAiLocation( originalrange, newrange );
				}
			}
		}

		return changed;
	}

	/**
	 * Return a string representing all series in this chart
	 *
	 * @param nChart 0=default, 1-9= overlay charts -1 for ALL charts
	 * @return
	 */
	public String[] getSeries( int nChart )
	{
		Vector seriesperchart = getAllSeries( nChart );
		String[] retStr = new String[seriesperchart.size()];
		for( int i = 0; i < seriesperchart.size(); i++ )
		{
			Series s = (Series) seriesperchart.get( i );
			Ai a = s.getSeriesValueAi();
			retStr[i] = a.toString();
		}
		return retStr;
	}

	/**
	 * Return an array of strings, one for each category
	 *
	 * @param nChart 0=default, 1-9= overlay charts -1 for All
	 * @return
	 */
	public String[] getCategories( int nChart )
	{
		Vector seriesperchart = getAllSeries( nChart );
		String[] retStr = new String[seriesperchart.size()];
		for( int i = 0; i < seriesperchart.size(); i++ )
		{
			Series s = (Series) seriesperchart.get( i );
			Ai a = s.getCategoryValueAi();
			retStr[i] = a.toString();
		}
		return retStr;
	}

	/**
	 * get all the series objects for ALL charts
	 * (i.e. even in overlay charts)
	 *
	 * @return
	 */
	public Vector getAllSeries()
	{
		return getAllSeries( -1 );
	}

	/**
	 * get all the series objects in the specified chart (-1 for ALL)
	 *
	 * @param nChart 0=default, 1-9= overlay charts -1 for ALL series
	 * @return
	 */
	public Vector getAllSeries( int nChart )
	{
/*    	if (nChart==-1) { //return all series   		
    		return new Vector(series.keySet()); // unordered!!!
    	}
*/
		Vector retVec = new Vector();
		for( Object[] sery : series )
		{
			Integer chart = (Integer) sery[1];
			if( (nChart == -1) || (nChart == chart) )
			{
				Series s = (Series) sery[0];
				retVec.add( s );
			}
		}
		return retVec;
	}

	/**
	 * Add a series object to the array.
	 *
	 * @param seriesRange   = one row range, expressed as (Sheet1!A1:A12);
	 * @param categoryRange = category range
	 * @param bubbleRange=  bubble range, if any 20070731 KSC
	 * @param seriesText    = label for the series;
	 * @param nChart        0=default, 1-9= overlay charts
	 *                      s     * @return
	 */
	public Series addSeries( String seriesRange,
	                         String categoryRange,
	                         String bubbleRange,
	                         String legendRange,
	                         String legendText,
	                         ChartType chartObject,
	                         int nChart )
	{
		Series s = Series.getPrototype( seriesRange, categoryRange, bubbleRange, legendRange, legendText, chartObject );
		s.setParentChart( chartObject.getParentChart() );
		s.setShape( chartObject.getBarShape() );
		series.add( new Object[]{ s, nChart } );
		// Update parent chartArr
		ArrayList<XLSRecord> chartArr = s.getParentChart().chartArr;
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec br = chartArr.get( i );
			BiffRec br2 = null;
			if( i < (chartArr.size() - 1) )
			{
				br2 = chartArr.get( i + 1 );
			}
			if( (br != null) && (((br.getOpcode() == XLSConstants.SERIES) && (br2.getOpcode() != XLSConstants.SERIES)) || ((br.getOpcode() == XLSConstants.FRAME) && (br2
					.getOpcode() == XLSConstants.SHTPROPS))) )
			{
				chartArr.add( i + 1, s );
				break;
			}
		}
		if( chartObject.getChartType() == STOCKCHART )
		{
			s.setHasLines( 5 );
		}
		return s;
	}

	/**
	 * specialty method to take absolute index of series and remove it
	 * <br>only used in WorkSheetHandle
	 *
	 * @param index
	 */
	public void removeSeries( int index )
	{
		Vector v = getAllSeries();
		removeSeries( (Series) v.get( index ) );
	}

	/**
	 * remove desired series from chart
	 *
	 * @param index
	 */
	public void removeSeries( Series seriestodelete )
	{
		for( int z = 0; z < series.size(); z++ )
		{
			Series ss = (Series) series.get( z )[0];
			if( ss.equals( seriestodelete ) )
			{
				series.remove( z );
				break;
			}
		}
	}

	/**
	 * get a chart series handle based off the series name
	 *
	 * @param seriesName
	 * @param nChart     0=default, 1-9= overlay charts
	 * @return
	 */
	public Series getSeries( String seriesName, int nChart )
	{
		Vector seriesperchart = getAllSeries( nChart );
		for( Object aSeriesperchart : seriesperchart )
		{
			Series s = (Series) aSeriesperchart;
			if( s.getLegendText().equalsIgnoreCase( seriesName ) )
			{
				return s;
			}
		}
		return null;
	}

	/**
	 * changes the category range which matches originalrange to new range
	 *
	 * @param originalrange
	 * @param newrange
	 * @return
	 */
	public boolean changeCategoryRange( String originalrange, String newrange )
	{
		boolean changed = false;
		for( Object[] sery : series )
		{
			Series s = (Series) sery[0];
			Ai ai = s.getCategoryValueAi();
			if( ai.toString().equalsIgnoreCase( originalrange ) )
			{
				changed = ai.changeAiLocation( originalrange, newrange );
			}
		}
		return changed;
	}

	/**
	 * attempts to replace all category elements containing originalval text to newval
	 *
	 * @param originalval
	 * @param newval
	 * @return
	 */
	public boolean changeTextValue( String originalval, String newval )
	{
		boolean changed = false;
		for( Object[] sery : series )
		{
			Series s = (Series) sery[0];
			Ai ai = s.getCategoryValueAi();
			if( ai.getText().equals( originalval ) )
			{
				ai.setText( newval );
			}

		}
		return changed;
	}

	/**
	 * retrieve the current data range JSON for comparisons
	 *
	 * @return JSONObject in form of: c:catgeory range, {v:series range, l:legendrange, b: bubble sizes} for each series
	 */
	public org.json.JSONObject getDataRangeJSON()
	{
		JSONObject seriesJSON = new JSONObject();
		java.util.Vector allseries = getAllSeries( -1 );
		try
		{
			JSONArray series = new JSONArray();
			for( int i = 0; i < allseries.size(); i++ )
			{
				try
				{
					//v=series val range, l=legend cell [,b= bubble sizes]
					Series thisseries = (Series) allseries.get( i );
					Ai serAi = thisseries.getSeriesValueAi();
					if( i == 0 )
					{
						Ai catAi = thisseries.getCategoryValueAi();
						seriesJSON.put( "c", catAi.toString() );
					}
					JSONObject seriesvals = new JSONObject();
					seriesvals.put( "v", serAi.toString() );
					seriesvals.put( "l", thisseries.getLegendRef() );
					if( thisseries.hasBubbleSizes() )
					{
						seriesvals.put( "b", thisseries.getBubbleValueAi().toString() );
					}
					series.put( seriesvals );
				}
				catch( Exception e )
				{ // keep going
				}
			}
			seriesJSON.put( "Series", series );
		}
		catch( JSONException e )
		{
			log.error( "ChartSeries.getDataRangeJSON:  Error retrieving Series Information: " + e.toString() );
		}
		// seriesJSON.getJSONArray("Series").length()
		// seriesJSON.getJSONArray("Series").get(1) 
		// ((JSONObject)seriesJSON.getJSONArray("Series").get(1)).get("v")	value range
		// ((JSONObject)seriesJSON.getJSONArray("Series").get(1)).get("l")	legend ref
		// ((JSONObject)seriesJSON.getJSONArray("Series").get(1)).has("b")	bubble sizes
		// seriesJSON.get("c")		category range
		return seriesJSON;
	}

	/**
	 * retrieve the current Series JSON for comparisons
	 *
	 * @return JSONArray
	 */
	public JSONArray getSeriesJSON()
	{
		return seriesJSON;
	}

	/**
	 * set the current Series JSON
	 *
	 * @param s JSONArray
	 */
	public void setSeriesJSON( JSONArray s ) throws JSONException
	{
		seriesJSON = new JSONArray( s.toString() );
	}

	/**
	 * return an array of ALL cell references of the chart
	 */
	public Ptg[] getCellRangePtgs()
	{
		CompatibleVector locptgs = new CompatibleVector();
		for( Object[] sery : series )
		{
			Series s = (Series) sery[0];
			for( int j = 0; j < s.chartArr.size(); j++ )
			{
				BiffRec br = s.chartArr.get( j );
				if( br.getOpcode() == XLSConstants.AI )
				{
					try
					{
						Ptg[] ps = ((Ai) br).getCellRangePtgs();
						for( Ptg p : ps )
						{
							locptgs.add( p );
						}
					}
					catch( Exception e )
					{
					}
				}
			}
		}
		Ptg[] ret = new Ptg[locptgs.size()];
		locptgs.toArray( ret );
		return ret;
	}

	/**
	 * @return an HashMap of Series Range Ptgs mapped by Series representing all the series in the chart
	 */
	public HashMap getSeriesPtgs()
	{
		HashMap seriesPtgs = new HashMap();
		for( Object[] sery : series )
		{
			Series s = (Series) sery[0];
			for( int j = 0; j < s.chartArr.size(); j++ )
			{
				BiffRec br = s.chartArr.get( j );
				if( br.getOpcode() == XLSConstants.AI )
				{
					if( ((Ai) br).getType() == Ai.TYPE_VALS )
					{
						try
						{
							Ptg[] ps = ((Ai) br).getCellRangePtgs();
							seriesPtgs.put( s, ps );
							break;    // we're done, move onto next
						}
						catch( Exception e )
						{
						}
					}
				}
			}
		}
		return seriesPtgs;
	}

	/**
	 * for overlay charts, store series list
	 *
	 * @param nCharts
	 * @param seriesList
	 */
	public void addSeriesMapping( int nCharts, int[] seriesList )
	{
		for( int aSeriesList : seriesList )
		{
			try
			{
				int idx = aSeriesList - 1;
				// ALL series in chart MAPPED to chart that "owns" it            			
				Series s = (Series) series.get( idx )[0];
				series.add( idx, new Object[]{ s, nCharts } );
			}
			catch( ArrayIndexOutOfBoundsException ae )
			{
				; // happens -- are they deleted series?????
			}
		}
	}

	/**
	 * return an array of legend text
	 *
	 * @param nChart 0=default, 1-9= overlay charts -1 for ALL
	 */
	public String[] getLegends( int nChart )
	{
		Vector seriesperchart = getAllSeries( nChart );
		String[] ret = new String[seriesperchart.size()];
		for( int i = 0; i < seriesperchart.size(); i++ )
		{
			ret[i] = ((Series) seriesperchart.get( i )).getLegendText();
		}
		return ret;
	}

	/**
	 * return the type of markers for each series:
	 * <br>0 = no marker
	 * <br>1 = square
	 * <br>2 = diamond
	 * <br>3 = triangle
	 * <br>4 = X
	 * <br>5 = star
	 * <br>6 = Dow-Jones
	 * <br>7 = standard deviation
	 * <br>8 = circle
	 * <br>9 = plus sign
	 */
	public int[] getMarkerFormats()
	{
		int[] markers = new int[series.size()];
		for( int i = 0; i < series.size(); i++ )
		{
			Series s = (Series) series.get( i )[0];
			markers[i] = s.getMarkerFormat();
		}
		return markers;
	}

	/**
	 * parse a chartSpace->chartType->ser element into our Series record/structure
	 *
	 * @param xpp         XML pullparser positioned at ser element
	 * @param wbh         WorkBookHandle
	 * @param parentChart parent chart object
	 * @param lastTag
	 * @return
	 */
	public static Series parseOOXML( XmlPullParser xpp,
	                                 WorkBookHandle wbh,
	                                 ChartType parentChart,
	                                 boolean hasPivotTableSource,
	                                 Stack<String> lastTag )
	{
		try
		{
			int eventType = xpp.getEventType();
			int idx = 0;
			int seriesidx = parentChart.getParentChart().getAllSeries().size();
			String[] ranges = { "", "", "", "" };     //legend, cat, ser/value, bubble cell references
			String legendText = "";
			SpPr sp = null;
			DLbls d = null;
			Marker m = null;
			boolean smooth = false;
			ArrayList dpts = null;
			String cache = null;
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "ser" ) )
					{ // Represents a single field in the PivotTable. This complex type contains
						idx = 0;
					}
					else if( tnm.equals( "order" ) )
					{// attr val= order
					}
					else if( tnm.equals( "cat" ) || tnm.equals( "xVal" ) )
					{        // children: CHOICE OF: multiLvlStrRef, numLiteral, numRef, strLit, strRef
						idx = 1;
					}
					else if( tnm.equals( "val" ) || tnm.equals( "yVal" ) )
					{        // children: CHOICE OF: numLit, numRef
						idx = 2;
					}
					else if( tnm.equals( "dLbls" ) )
					{        // data labels
						lastTag.push( tnm );                // keep track of element hierarchy
						d = (DLbls) DLbls.parseOOXML( xpp, lastTag, wbh ).cloneElement();
						if( d.showBubbleSize() )
						{
							parentChart.setChartOption( "ShowBubbleSizes", "1" );
						}
						if( d.showCatName() )
						{
							parentChart.setChartOption( "ShowCatLabel", "1" );
						}
						if( d.showLeaderLines() )
						{
							parentChart.setChartOption( "ShowLdrLines", "1" );
						}
						if( d.showLegendKey() )
						{
							; // TODO: handle show legend key
						}
						if( d.showPercent() )
						{
							parentChart.setChartOption( "ShowLabelPct", "1" );
						}
						if( d.showSerName() )
						{
							parentChart.setChartOption( "ShowLabel", "1" );
						}
						if( d.showVal() )
						{
							parentChart.setChartOption( "ShowValueLabel", "1" );
						}
						// data label options
					}
					else if( tnm.equals( "dPt" ) )
					{            // data point(s)
						if( dpts == null )
						{
							dpts = new ArrayList();
						}
						lastTag.push( tnm );                // keep track of element hierarchy
						dpts.add( DPt.parseOOXML( xpp, lastTag, wbh ).cloneElement() );
					}
					else if( tnm.equals( "spPr" ) )
					{    // series spPr
						lastTag.push( tnm );                // keep track of element hierarchy
						sp = (SpPr) SpPr.parseOOXML( xpp, lastTag, wbh ).cloneElement();
					}
					else if( tnm.equals( "marker" ) )
					{
						lastTag.push( tnm );                // keep track of element hierarchy
						m = (Marker) Marker.parseOOXML( xpp, lastTag, wbh ).cloneElement();
					}
					else if( tnm.equals( "bubbleSize" ) )
					{
						idx = 3;
					}
					else if( tnm.equals( "shape" ) )
					{    // bar only
						parentChart.convertShape( xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "smooth" ) )
					{  // line chart
						smooth = ((xpp.getAttributeCount() == 0) || !xpp.getAttributeValue( 0 ).equals( "0" ));
					}
					else if( tnm.equals( "explosion" ) )
					{
						String v = xpp.getAttributeValue( 0 );
						parentChart.setChartOption( "Percentage", v );
						// NOTE: two types of values; 1- reference denoted by f element parents can be numRef, strRef or multiLvlStrRef
						//						  2- text value denoted by v element parent is strRef
					}
					else if( tnm.equals( "formatCode" ) )
					{    // part of numCache element
						Xf.addFormatPattern( wbh.getWorkBook(), OOXMLAdapter.getNextText( xpp ) );
// *******************************************						
// TODO: add to y pattern ***						
// *******************************************						
					}
					else if( tnm.equals( "f" ) )
					{    // range element  -- legend cell, Cat range, Value range, Bubble data reference
						ranges[idx] = OOXMLAdapter.getNextText( xpp );
					}
					else if( tnm.equals( "v" ) )
					{    // value or text of series or category (parent=tx)
						if( idx == 0 ) // legend text
						{
							legendText = OOXMLAdapter.getNextText( xpp );    // legend text; possible to have legend text without a legend cell range (ranges[1])
						}
						else if( (idx == -1) || ranges[idx].equals( "" ) )
						{ // shoudln't!! can't have a textual refernce in place of a series or cat value (can you?)
							log.warn( "ChartSeries.parseOOXML: unexpected text value" );
						}
					}
					else if( tnm.equals( "numCache" ) || tnm.equals( "strCache" ) || tnm.equals( "multiLvlStrRef" ) )
					{    // parent= cat or vals (series values)
						cache = tnm;
					}
					else if( tnm.equals( "ptCount" ) )
					{    // parent= numCache or strCache (governs either f element)
						if( hasPivotTableSource )
						{    // OK, if have a pivot table source then the range referenced in f is only a SUBSET
							// unclear if at any other time the range referenced is a subset ... [NOTE: in testing, only pivot charts hit]
							// another assumption:  assume that range is only TRUNCATED -- in testing, true
							int npoints = Integer.valueOf( xpp.getAttributeValue( 0 ) );
							if( !ranges[idx].equals( "" ) && (ranges[idx].indexOf( "," ) == -1) )
							{
								try
								{
									CellRange cells = new CellRange( ranges[idx], wbh, false, true );
									if( cells.getCells().length != npoints )
									{    //must adjust
										int z = 0;
										CellHandle[] clist = cells.getCells();
										while( eventType != XmlPullParser.END_DOCUMENT )
										{
											if( eventType == XmlPullParser.START_TAG )
											{
												tnm = xpp.getName();
												if( tnm.equals( "pt" ) )
												{
													// format code idx
												}
												else if( tnm.equals( "v" ) )
												{
/* this case should NOT happen
													String s= OOXMLAdapter.getNextText(xpp);
 													if (z < clist.length) 
														if (!clist[z].getVal().toString().equals(s)) 															
															Logger.logWarn("ChartSeries.parseOOXML: unexpected pivot value order- skipping");
													z++;
*/
												}
											}
											else if( eventType == XmlPullParser.END_TAG )
											{
												if( xpp.getName().equals( cache ) )
												{
													cache = null;
													break;
												}
											}
											eventType = xpp.next();
										}
										// pivot charts: apparently always truncate/skip last cell in range (which represents the grand total)
										if( npoints < clist.length )
										{// truncate!
											int[] rc = cells.getRangeCoords();
											rc[0]--;
											rc[2]--;    // make 0-based
											if( rc[0] == rc[2] )
											{
												rc[3] -= (clist.length - npoints);
											}
											else
											{
												rc[2] -= (clist.length - npoints);
											}
// KSC: TESTING: REMOVE WHEN DONE											
//System.out.println("Truncate list: old range: " + ranges[idx] + " new range: " + cells.getSheet().getQualifiedSheetName() + "!" + ExcelTools.formatLocation(rc));
											ranges[idx] = cells.getSheet().getQualifiedSheetName() + "!" + ExcelTools.formatLocation( rc );
										}

										continue;    // don't hit xpp.next() below
									}
								}
								catch( Exception e )
								{
									log.error( "ChartSeries.parseOOXML: Error adjusting pivot range for " + parentChart + ":" + e.toString() );
								} // problems parsing range - skp
							}
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( "ser" ) )
					{
						lastTag.pop();
						// must only have 1 for pie-type charts (pie, of pie - bar of, pie of)
						Series s = parentChart.getParentChart().getChartSeries().addSeries( ranges[2],
						                                                                    ranges[1],
						                                                                    ranges[3],
						                                                                    ranges[0],
						                                                                    legendText,
						                                                                    parentChart,
						                                                                    parentChart.getParentChart().getChartOrder(
								                                                                    parentChart ) );
						if( sp != null )
						{
							s.setSpPr( sp );
						}
						// TODO: " When you create a chart, by default - the first six series are the six accent colors in order - but not the exact color or any variation that appears in the palette. They're typically (unless the primary accent color being modified is extremely dark) a bit darker than the primary accent color. Chart series 7 - 12 use the actual primary accent colors 1 through 6 ... and then chart series 13 starts a set of lighter variations of the six accent colors that are also slightly different from any position in the palette."
						else if( seriesidx < 7 ) // TODO: figure out where to get colors past 6
						{
							if( (seriesidx > 0) && (parentChart instanceof PieChart) )
							{
								log.warn( "ChartSeries.parseOOXML:  more than 1 series encountered for a Pie-style chart" );
							}
							else
							{
								s.setColor( wbh.getWorkBook().getTheme().genericThemeClrs[seriesidx + 4] ); // series colors start at 4
							}
						}
						if( d != null )
						{
							s.setDLbls( d );
						}
						if( dpts != null )
						{
							for( Object dpt : dpts )
							{
								s.addDpt( (DPt) dpt );
							}
						}
						if( m != null )
						{
							s.setMarker( m );
						}
						if( smooth )
						{
							s.setHasSmoothLines( smooth );
						}
						return s;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "ChartSeries.parseOOXML: Error parsing series for " + parentChart + ":" + e.toString() );
		}
		return null;
	}

	/**
	 * Generate the OOXML used to represent all series for this chart type
	 * @param ct        chart type
	 * @return
	 */
	/**
	 * all contain:
	 * idx			(index)
	 * order		(order)
	 * tx			(series text)
	 * spPr		(shape properties)
	 * then after, may contain:
	 * bubble:  invertIfNegative, dPt, dLbls, trendline, errBars, xVal, yVal, bubbleSize, bubble3D		(bubbleChart)
	 * line 	marker, dPt, dLbls, trendline, errBars, cat, val, smooth								(line3DChart, lineChart, stockChart)
	 * pie:		explosion, dPt,dLbls, cat, val															(doughnutChart, ofPieChart, pie3DChart, pieChart)
	 * surface: cat, val																				(surfaceChart, surface3DChart)
	 * scatter: marker, dPt, dLbls, trendline, errBars, xVal, yVal, smooth								(scatterChart)
	 * radar:	marker, dPt, dLbls, cat, val															(radarChart)
	 * area:	pictureOptions, dPt, dLbls, trendline, errBars, cat, val								(area3DChart, areaChart)
	 * bar:		invertIfNegative, pictureOptions, dPt, dLbls, trendline, errBars, cat, val, shape		(bar3DChart, barChart)
	 */
	// TODO: finish options
	// TODO: refactor !!!
	public void resetSeriesNumber()
	{
		seriesNumber = 0;
	}

	int seriesNumber = 0;

	public String getOOXML( int ct, boolean isBubble3d, int nChart )
	{
		String catstr = ((ct == SCATTERCHART) || (ct == BUBBLECHART)) ? ("xVal") : ("cat");
		String valstr = ((ct == SCATTERCHART) || (ct == BUBBLECHART)) ? ("yVal") : ("val");
		StringBuffer ooxml = new StringBuffer();

		Vector v = parentChart.getAllSeries( nChart );
		int defaultDL = parentChart.getDataLabel();
		boolean from2003 = (!parentChart.getWorkBook().getIsExcel2007());
		String[] cats = getCategories( nChart );    // do 1x

		for( int i = 0; i < v.size(); i++ )
		{
			Series s = (Series) v.get( i );
			ooxml.append( "<c:ser>" );
			ooxml.append( "\r\n" );
			ooxml.append( "<c:idx val=\"" + seriesNumber + "\"/>" );
			ooxml.append( "\r\n" );
			ooxml.append( "<c:order val=\"" + seriesNumber++ + "\"/>" );
			ooxml.append( "\r\n" );
			// Series Legend
			ooxml.append( s.getLegendOOXML( from2003 ) );
			// Options for current series
			if( ct == PIECHART /*&& i==0*/ )
			{
				ooxml.append( "<c:explosion val=\"" + parentChart.getChartOption( "Percentage" ) + "\"/>" );
			}
			if( s.getMarker() != null )
			{
				ooxml.append( s.getMarker().getOOXML() );        // only for Radar, Line or Scatter
			}
			if( s.getDPt() != null )
			{
				DPt[] datapoints = s.getDPt();
				for( DPt datapoint : datapoints )
				{
					ooxml.append( datapoint.getOOXML() );
				}
			}
			if( s.getDLbls() != null )
			{
				ooxml.append( s.getDLbls().getOOXML() );
			}
			else if( from2003 )
			{
				int dl = s.getDataLabel() | defaultDL;
				if( dl > 0 )
				{    // todo: showLegendKey catpercent ? sppr + txpr
					// TODO: spPr, txPr
					DLbls dlbl = new DLbls( ((dl | 0x1) == 0x1),
					                        ((dl | 0x08) == 0x08),
					                        false,
					                        ((dl | 0x10) == 0x10),
					                        ((dl | 0x40) == 0x40),
					                        ((dl | 0x2) == 0x2),
					                        ((dl | 0x20) == 0x20),
					                        null /*sppr*/,
					                        null /*txpr*/ );
					ooxml.append( dlbl.getOOXML() );
				}
			}
			if( s.getHasSmoothedLines() )
			{
				ooxml.append( "<c:smooth val=\"1\"/>" );
				ooxml.append( "\r\n" );
			}

			// Categories 							NOTE:  Categories==xVals for Scatter charts, cat for all others
			ooxml.append( s.getCatOOXML( cats[i], catstr ) );
			// Series ("vals")						NOTE:  Series==yVals for Scatter charts, val for all others
			ooxml.append( s.getValOOXML( valstr ) );    // gets the numeric data reference to define the series (values)

			if( ct == BUBBLECHART )
			{ // also include bubble sizes
				ooxml.append( s.getBubbleOOXML( isBubble3d ) );
			}
			ooxml.append( "</c:ser>" );
			ooxml.append( "\r\n" );
		}
		return ooxml.toString();
	}

	/**
	 * if has multiple or overlay charts, update series mappings
	 *
	 * @param sl
	 * @param thischartnumber
	 */
	protected void updateSeriesMappings( SeriesList sl, int thischartnumber )
	{
		if( sl == null )
		{
			return;    // NO seriesList record means no mapping
		}
		ArrayList seriesmappings = new ArrayList();
		for( int z = 0; z < series.size(); z++ )
		{
			int chartnumber = (Integer) series.get( z )[1];
			if( chartnumber == thischartnumber )    // mappped to this chart
			{
				seriesmappings.add( z + 1 );
			}
		}
		int[] mappings = new int[seriesmappings.size()];
		for( int z = 0; z < seriesmappings.size(); z++ )
		{
			mappings[z] = (Integer) seriesmappings.get( z );
		}
		try
		{
			sl.setSeriesMappings( mappings );
		}
		catch( Exception e )
		{
			throw new WorkBookException( "ChartSeries.updateSeriesMappings failed:" + e.toString(), WorkBookException.RUNTIME_ERROR );
		}
	}

	/**
	 * return Data Labels Per Series or default Data Labels, if no overrides specified
	 *
	 * @param defaultDL Default Data labels
	 * @param charttype Chart Type Int
	 * @return
	 */
	public int[] getDataLabelsPerSeries( int defaultDL, int charttype )
	{
		if( (charttype == PIECHART) || (charttype == DOUGHNUTCHART) )
		{    // handled differently
			if( series.size() > 0 )
			{
				Series s = (Series) series.get( 0 )[0];
				int[] dls = s.getDataLabelsPIE( defaultDL );
				if( dls == null )
				{
					dls = new int[]{ defaultDL };
				}
				return dls;
			}
		}
		int[] datalabels = new int[series.size()];
		for( int i = 0; i < series.size(); i++ )
		{
			Series s = (Series) series.get( i )[0];
			datalabels[i] = s.getDataLabel();
			datalabels[i] |= defaultDL; // if no per-series setting use overall chart setting
		}
		return datalabels;
	}

	// TODO: FINISH -- include cell range ... + overlay charts ...?
	public String[] getLegends()
	{
		if( legends == null )
		{
			getMetrics( true );
		}
		return legends;
	}

	public Object[] getCategories()
	{
		if( categories == null )
		{
			getMetrics( true );
		}
		return categories;
	}

	public ArrayList getSeriesRanges()
	{
		if( seriesranges == null )
		{
			getMetrics( true );
		}
		return seriesranges;
	}

	public ArrayList getSeriesValues()
	{
		if( seriesvalues == null )
		{
			getMetrics( true );
		}
		return seriesvalues;
	}

	public String[] getSeriesBarColors()
	{
		if( seriescolors == null )
		{
			getMetrics( true );
		}
		return seriescolors;
	}

}
