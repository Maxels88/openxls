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

//OOXML-specific structures

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.Layout;
import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.OOXML.Title;
import com.extentech.formats.OOXML.TxPr;
import com.extentech.toolkit.Logger;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;

public class OOXMLChart extends Chart
{
	public String lang = "en-US";                                // default
	public boolean roundedCorners = false;
	public Title ttl = null;                                    // title element
	public com.extentech.formats.OOXML.Legend ooxmlLegend = null;
	public Layout plotAreaLayout = null;
	private SpPr plotareashapeProps = null;                    // defines the shape properties for this chart (line and fill)
	private SpPr csshapeProps = null;                        // defines the shape properties for this chart space
	private TxPr txpr = null;
	private String editMovement = "twoCell";                    // default
	private String name = null;                                // name property of cNvPr
	private ArrayList chartEmbeds = null;                    // if present, name(s) of chart embeds (images, userShape definitions (drawingml files that define shapes ontop of chart))

	/**
	 * create a new OOXMLChart from a 2003-v chart object
	 *
	 * @param c
	 * @param wbh
	 */
	public OOXMLChart( Chart c, WorkBookHandle wbh )
	{
		// Walk up the superclass hierarchy
		//System.out.println("BEFORE: chartArr: " + Arrays.toString(chartArr.toArray()));		
		for( Class obj = c.getClass(); !obj.equals( Object.class ); obj = obj.getSuperclass() )
		{
			java.lang.reflect.Field[] fields = obj.getDeclaredFields();
			for( Field field : fields )
			{
				field.setAccessible( true );
				try
				{
					// for each class/suerclass, copy all fields
					// from this object to the clone
					field.set( this, field.get( c ) );
				}
				catch( IllegalArgumentException e )
				{
				}
				catch( IllegalAccessException e )
				{
				}
			}
		}

		// ttl?
		// chartlaout?
		// txpr?
		// name?
		this.name = c.getTitle();
		if( c.hasDataLegend() )
		{
			ooxmlLegend = com.extentech.formats.OOXML.Legend.createLegend( c.getLegend() );
		}
		this.wbh = wbh;
	}

	public String toString()
	{
		String t = getTitle();
		if( !t.equals( "" ) )
		{
			return t;
		}
		return name;    // if no title, return OOXMLName
	}

	/**
	 * return the OOXML shape property for this chart
	 *
	 * @param type 0= chart shape props, 1=plot area shape props 2= chartspace shape props, 3= legend shape props
	 * @return
	 */
	public SpPr getSpPr( int type )
	{
		if( type == 0 )
		{
			return plotareashapeProps;
		}
		if( type == 1 )
		{
			return csshapeProps;
		}
		return null;
	}

	/**
	 * store the OOXML text formatting element for this chart
	 */
	public void setTxPr( TxPr txPr )
	{
		txpr = txPr;
	}

	/**
	 * return the OOXML text formatting element for this chart, if present
	 *
	 * @return
	 */
	public TxPr getTxPr()
	{
		return txpr;
	}

	/**
	 * define the OOXML shape property for this chart from an existing spPr element
	 *
	 * @param type 0=plot area shape props 1= chartspace shape props, 2= legend shape props
	 */
	public void setSpPr( int type, SpPr spPr )
	{
		if( type == 0 )
		{
			plotareashapeProps = spPr;    // plot area
			int lw = -1, lclr = 0, bgcolor = -1;
			lw = spPr.getLineWidth();    // TO DO: Style
			lclr = spPr.getLineColor();
			//bgcolor= spPr.getColor();
			this.getAxes().setPlotAreaBorder( lw, lclr );
		}
		else if( type == 1 )
		{
			csshapeProps = spPr;
		}

	}

	/**
	 * return the OOXML title element for this chart
	 *
	 * @return
	 */
	public Title getOOXMLTitle()
	{
		return ttl;
	}

	/**
	 * set the OOXML title element for this chart
	 *
	 * @param t
	 */
	public void setOOXMLTitle( Title t, WorkBookHandle wb )
	{
		ttl = t;
		int fid = ttl.getFontId( wb );
		if( fid == -1 )
		{
			fid = 5;    // default ...?
		}
		float[] coords = null;
		int lw = -1, lclr = 0, bgcolor = 0;
		if( ttl.getLayout() != null )
		{    // pos
			coords = ttl.getLayout().getCoords();
		}
		if( ttl.getSpPr() != null )
		{    // Area Fill, Line Format
			SpPr sp = ttl.getSpPr();
			lw = sp.getLineWidth();    // TO DO: Style, fill/color ...
			lclr = sp.getLineColor();
			bgcolor = sp.getColor();
		}
		if( coords != null )
		{
			charttitle.setFrame( lw, lclr, bgcolor, coords );
		}

		// must also set the fontx id for the title
		if( charttitle != null )
		{
			charttitle.setFontId( fid );
		}

	}

	/**
	 * specify how to resize or move upon edit OOXML specific
	 *
	 * @param editMovement
	 */
	public void setEditMovement( String editMovement )
	{
		this.editMovement = editMovement;
		dirtyflag = true;
	}

	/**
	 * return state of  how to resize or move upon edit OOXML specific
	 *
	 * @return editMovement string
	 */
	public String getEditMovement()
	{
		return this.editMovement;
	}

	/**
	 * remove the legend from the chart
	 */
	@Override
	public void removeLegend()
	{
		showLegend( false, false );
		ooxmlLegend = null;
	}

	/**
	 */
	public String getOOXMLName()
	{
		return this.name;
	}

	/**
	 * set the OOXML-specific name for this chart
	 *
	 * @param name
	 */
	public void setOOXMLName( String name )
	{
		this.name = name;
		dirtyflag = true;
	}

	/**
	 * returns information regarding external files associated with this chart
	 * <br>e.g. a chart user shape, an image
	 *
	 * @return
	 */
	public ArrayList getChartEmbeds()
	{
		return chartEmbeds;
	}

	/**
	 * sets external information linked to or "embedded" in this OOXML chart;
	 * can be a chart user shape, an image ...
	 * <br>NOTE: a userShape is a drawingml file name which defines the userShape (if any)
	 * <br>a userShape is a drawing or shape ontop of a chart
	 *
	 * @param String[] embedType, filename e.g. {"userShape", "userShape file name"}
	 */
	public void addChartEmbed( String[] ce )
	{
		if( chartEmbeds == null )
		{
			chartEmbeds = new ArrayList();
		}
		chartEmbeds.add( ce );
	}

	/**
	 * return the OOXML representation of this chart object "c:chart" representing OOXML element in chartX.xml
	 * <br>below is complete the ordered sequence of child elements of the chart element
	 * <br>c:chart - parent= chartSpace
	 * <li>title
	 * <li>autoTitleDeleted
	 * <li>pivotFmts
	 * <li>view3d
	 * <li>floor
	 * <li>sideWall
	 * <li>backWall
	 * <li>plotArea (see below)
	 * <li>legend
	 * <li>plotVisOnly
	 * <li>dispBlankAs
	 * <li>showDlblsOverMax
	 * <p/>
	 * <br>plotArea:
	 * <li>layout
	 * <li>chart type (see below)
	 * <li> axes ***
	 * <li>dTable
	 * <li>spPr
	 * <p/>
	 * <br>chart type:
	 * <li>barDir		Bar, Bar3d only
	 * <li>radarStyle || scatterStyle
	 * <li>ofPieType
	 * <li>wireFrame	surface
	 * <li>grouping	Area, Area3d, Line, Line3d, Bar, Bar3D
	 * <li>varyColors		not for Stock
	 * <li>ser  *n series
	 * <li>dLbls			not for surface
	 * Area Chart, AreaChart3D, LineChart, Line3D, Stock
	 * <li>dropLines
	 * Bar Chart, Bar3d, ofPieChart
	 * <li>gapWidth
	 * AreaChart3D, Line3D, Bar3D
	 * <li>gapDepth
	 * Line, Stock
	 * <li>hiLowLines
	 * <li>upDownBars
	 * Line
	 * <li>marker
	 * <li>smooth
	 * BarChart only
	 * <li>overlap
	 * <li>serLines
	 * Bar 3d only
	 * <li>shape
	 * ofPieChart
	 * <li>splitType
	 * <li>splitPos
	 * <li>custSplit
	 * <li>secondPieSize
	 * <li>serLines
	 * Pie, Doughnut
	 * <li>firstSliceAng
	 * Doughnut
	 * <li>holeSize
	 * Surface
	 * <li>bandFmts
	 * Bubble
	 * <li>bubble3D
	 * <li>bubbleScale
	 * <li>showNegBubbles
	 * <li>sizeRepresents
	 * <li>axId
	 *
	 * @return StringBuffer
	 */
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();
		try
		{
			int[] allCharts = this.getAllChartTypes();        // usually only 1 chart but some have overlay charts in addition to the default chart (chart 0)

			// lang
			cooxml.append( "<c:lang val=\"" + this.lang + "\"/>" );
			cooxml.append( "\r\n" );
			// rounded corners
			if( this.roundedCorners )
			{
				cooxml.append( "<c:roundedCorners val=\"1\"/>" );
			}
			// chart
			cooxml.append( "<c:chart>" );
			cooxml.append( "\r\n" );
			// title
			if( this.getOOXMLTitle() == null )
			{// if no OOXML title, see if have a BIFF8 title
				if( !this.getTitle().equals( "" ) )
				{
					this.setOOXMLTitle( new Title( this.getTitleTd(), this.wbh.getWorkBook() ), this.wbh );
				}
			}
			if( this.getOOXMLTitle() != null )    // otherwise there's no title
			{
				cooxml.append( this.getOOXMLTitle().getOOXML() );
			}

			if( allCharts[0] != BUBBLECHART )
			{    // bubble threeD handled in series for some reason
				// Q: what if overlay charts are not 3D?  what if default isn't and overlay is?
				// view 3D
				ThreeD td = this.getThreeDRec( 0 );
				if( td != null )
				{
					cooxml.append( td.getOOXML() );
				}
			}
			// TODO: Handle:
			// floor
			// sideWall
			// backWall
			// plot area (contains all chart types such as barChart, lineChart ...)
			cooxml.append( "<c:plotArea>" );
			cooxml.append( "\r\n" );
			// layout: size and position
			if( this.plotAreaLayout == null )
			{    // if converted from XLS will hit here
				HashMap<String, Double> chartMetrics = this.getMetrics( wbh );
				double x = chartMetrics.get( "x" ) / chartMetrics.get( "canvasw" );
				double y = chartMetrics.get( "y" ) / chartMetrics.get( "canvash" );
				double w = chartMetrics.get( "w" ) / chartMetrics.get( "canvasw" );
				double h = chartMetrics.get( "h" ) / chartMetrics.get( "canvash" );
				this.plotAreaLayout = new Layout( "inner", new double[]{ x, y, w, h } );
			}
			cooxml.append( this.plotAreaLayout.getOOXML() );

			for( ChartType ch : chartgroup )
			{
				cooxml.append( ch.getOOXML( catAxisId, valAxisId, serAxisId ) );
			}

/* TODO:
 * 1- varyColors  ???
 * 2- getChartSeries.getOOXML  -> nchart?
 * 3- data labels
 * 4- drop lines
 * 
 * area charts -- bar colors!!!
 * bar - serLines
 * pie, doughnut - firstSliceArg
 * radar -- filled?
 * bubble	-- bubbleScale, showNegBubbles, sizeRepresents
 * surface -- wireframe, bandfmts
 * 
 * pie of pie, bar of pie
 * surface3d
 * stock
 * 
 * 
 */

			// ******************************************************************************
			// after chart type ooxml, axes (if present)
			cooxml.append( this.getAxes().getOOXML( XAXIS, 0, catAxisId, valAxisId ) );
			cooxml.append( this.getAxes().getOOXML( XVALAXIS, 2, catAxisId, valAxisId ) );    // valAx - for bubble/scatter
			// val axis
			cooxml.append( this.getAxes().getOOXML( YAXIS, 1, valAxisId, catAxisId ) );    // val axis
			// ser axis
			cooxml.append( this.getAxes().getOOXML( ZAXIS, 3, serAxisId, valAxisId ) );    // ser axis (crosses val axis)
			// TODO: dateAx
			if( this.getSpPr( 0 ) != null )
			{    // plot area shape props
				cooxml.append( this.getSpPr( 0 ).getOOXML() );
			}
			else if( !this.wbh.getIsExcel2007() )
			{
				SpPr sp = new SpPr( "c", this.getPlotAreaBgColor().substring( 1 ), 12700, this.getPlotAreaLnColor().substring( 1 ) );
				cooxml.append( sp.getOOXML() );

			}

			cooxml.append( "</c:plotArea>" );
			cooxml.append( "\r\n" );
			// legend
			if( this.ooxmlLegend != null )
			{
				cooxml.append( this.ooxmlLegend.getOOXML() );
			}

			cooxml.append( "<c:plotVisOnly val=\"1\"/>" );        // specifies that only visible cells should be plotted on the chart
//	    	<c:dispBlanksAs val="gap"/>	"gap", "span", "zero"  --> default	    	
			cooxml.append( "\r\n" );
			cooxml.append( "</c:chart>" );
			cooxml.append( "\r\n" );
			if( this.getSpPr( 1 ) != null )
			{ // chart space shape props
				cooxml.append( this.getSpPr( 1 ).getOOXML() );
			}
			if( this.getTxPr() != null )
			{ // text formatting
				cooxml.append( this.getTxPr().getOOXML() );
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "OOXMLChart.getOOXML: error generating OOXML.  Chart not created: " + e.toString() );
		}
		return cooxml;    //.toString();
	}

}
