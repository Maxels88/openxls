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
import com.extentech.ExtenXLS.ChartHandle;
import com.extentech.ExtenXLS.ChartHandle.ChartOptions;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.OOXMLConstants;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.Label;
import com.extentech.formats.XLS.MSODrawing;
import com.extentech.formats.XLS.Obj;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.formats.cellformat.CellFormatFactory;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.xmlpull.v1.XmlPullParser;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.EnumSet;
import java.util.HashMap;

/**
 * This abstract class defines a Chart Group for a BIFF Chart.  An Excel Chart may have up to 4 (2003 and previous versions) or 9 (post-2003) charts within one chart.
 * The 0th chart type in the Chart Group is the default chart.
 * <br>
 * NOTES:
 * <br>
 * CRT = ChartFormat Begin (Bar / Line / (BopPop [BopPopCustom]) / Pie / Area / Scatter / Radar / RadarArea / Surf) CrtLink [SeriesList] [Chart3d] [LD] [2DROPBAR] *4(CrtLine LineFormat) *2DFTTEXT [DataLabExtContents] [SS] *4SHAPEPROPS End
 * <p/>
 * MUST CHANGE ChartObject if change chart type
 * MUST ADD/REMOVE ChartObject if add/remove multiple charts
 * axisparent  --> when add a new chart a new crt is added  axes + titles/labels, number format, etc
 * <p/>
 * REFACTOR:  TO FINISH
 * showDataTable -- Finish, TODO
 * NOTE: The SeriesList record specifies the series of the chart. This record MUST NOT exist in the first chart group in the chart sheet substream.
 * This record MUST exist when not in the first chart group in the chart sheet substream
 * <p/>
 * NOTE:
 * The Chart3d record specifies that the plot area, axis group, and chart group are rendered in a 3-D scene, rather than a 2-D scene, and specifies properties of the 3-D scene. If this record exists in the chart sheet substream, the chart sheet substream  MUST have exactly one chart group. This record MUST NOT exist in a bar of pie, bubble, doughnut, filled radar, pie of pie, radar, or scatter chart group.
 * <p/>
 * NOTE: legends only in 1st chart group
 */
public abstract class ChartType implements ChartConstants, Serializable
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7862828186455339066L;
	protected GenericChartObject chartobj;
	protected Legend legend = null;
	protected ChartFormat cf = null;
	//    protected ChartSeries chartseries= new ChartSeries();
	protected transient WorkBook wb = null;
	protected int defaultShape = 0;                            // controls default bar shape for all bars in the chart; used when adding or setting series

	public ChartType()
	{
	}

	public ChartType( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		chartobj = charttype;    // record that defines the chart type (Bar, Area, Line ...)
		this.wb = wb;
		this.cf = cf;
	}

	public Chart getParentChart()
	{
		return chartobj.getParentChart();
	}

	/**
	 * creates a new chart type object of the desired chart type.  Will set options if already set via ChartFormat
	 *
	 * @param chartType
	 * @param parentChart
	 * @param cf
	 * @return
	 */
	public static ChartType create( int chartType, Chart parentChart, ChartFormat cf )
	{
		GenericChartObject co = ChartType.createUnderlyingChartObject( chartType, parentChart, cf );    //, ((ChartType)chartobj.get(0)));

		ChartType ct = ChartType.createChartTypeObject( co, cf, parentChart.getWorkBook() );
		return ct;
	}

	/**
	 * create and return the appropriate ChartType for the given
	 * chart record (Bar, Radar, Line, etc)
	 *
	 * @param ch GenericChartObject representing the Chart Type
	 *           <br>Must be one of:  Area, Bar, Line, Surface, Pie, Radar, Scatter, BopPop, RadarArea
	 * @return ChartType chart object
	 */
	public static ChartType createChartTypeObject( GenericChartObject ch, ChartFormat cf, WorkBook wb )
	{
		if( cf == null )    // TODO: throw exception
		{
			return null;
		}

		int barshape = cf.getBarShape();
		boolean threeD = cf.isThreeD( ch.chartType );
		boolean isStacked = ch.isStacked();
		boolean is100Percent = ch.is100Percent();

		switch( ch.chartType )
		{
			case ChartConstants.COLCHART:
				if( barshape == SHAPEDEFAULT )
				{ // regular column
					if( isStacked || is100Percent )
					{
						return new StackedColumn( ch, cf, wb );
					}
					if( threeD )
					{
						return new Col3DChart( ch, cf, wb );
					}
					return new ColChart( ch, cf, wb );
				}
				if( barshape == SHAPECONE )
				{ // Cone chart	 always 3d
					return new ConeChart( ch, cf, wb );
				}
				if( barshape == SHAPECYLINDER )
				{ // Cylinder chart always 3d
					return new CylinderChart( ch, cf, wb );
				}
				if( barshape == SHAPEPYRAMID )
				{    // Pyramid chart	alwasy 3d
					return new PyramidChart( ch, cf, wb );
				}
			case ChartConstants.BARCHART:
				if( barshape == SHAPEDEFAULT )
				{ // regular Bar
					if( threeD )
					{
						return new Bar3DChart( ch, cf, wb );
					}
					return new BarChart( ch, cf, wb );
				}
				if( barshape == SHAPECONE )
				{ // ConeBarchart	always 3d
					return new ConeBarChart( ch, cf, wb );
				}
				if( barshape == SHAPECYLINDER )
				{ // CylinderBar chart	always 3d
					return new CylinderBarChart( ch, cf, wb );
				}
				if( barshape == SHAPEPYRAMID )
				{    // PyramidBar chart	always 3d
					return new PyramidBarChart( ch, cf, wb );
				}
			case ChartConstants.LINECHART:
				if( threeD )
				{
					return new Line3DChart( ch, cf, wb );
				}
				return new LineChart( ch, cf, wb );
			case ChartConstants.STOCKCHART:
				return new StockChart( ch, cf, wb );
			case ChartConstants.PIECHART:
				cf.setPercentage( 0 );
				if( threeD )
				{
					return new Pie3dChart( ch, cf, wb );
				}
				return new PieChart( ch, cf, wb );
			case ChartConstants.AREACHART:
				if( ch.isStacked() )
				{
					return new StackedAreaChart( ch, cf, wb );
				}
				if( threeD )
				{
					return new Area3DChart( ch, cf, wb );
				}
				return new AreaChart( ch, cf, wb );
			case ChartConstants.SCATTERCHART:
				return new ScatterChart( ch, cf, wb );
			case ChartConstants.RADARCHART:
				return new RadarChart( ch, cf, wb );
			case ChartConstants.SURFACECHART:
				if( ((Surface) ch).getIs3d() )
				{
					return new Surface3DChart( ch, cf, wb );
				}
				return new SurfaceChart( ch, cf, wb );
			case ChartConstants.DOUGHNUTCHART:
				cf.setPercentage( 0 );
				return new DoughnutChart( ch, cf, wb );
			case ChartConstants.BUBBLECHART:
				return new BubbleChart( ch, cf, wb );
			case ChartConstants.RADARAREACHART:
				return new RadarAreaChart( ch, cf, wb );
			case ChartConstants.OFPIECHART:
				cf.setPercentage( 0 );
				return new OfPieChart( ch, cf, wb );
			default:
				return null;
		}
	}

	/**
	 * create a new low-level chart object which determines the ChartTypeObject, and update the existing chartobj
	 *
	 * @param chartType
	 * @param parentchart
	 * @param chartgroup
	 * @return
	 */
	public static GenericChartObject createUnderlyingChartObject( int chartType, Chart parentchart, ChartFormat cf )
	{//, ChartType chartobj) {
		GenericChartObject c = null;
		switch( chartType )
		{
			case ChartConstants.COLCHART: // column-type
				Bar col = (Bar) Bar.getPrototype();
				col.setAsColumnChart();
				c = col;
				break;

			case ChartConstants.BARCHART: // bar-type
				Bar bar = (Bar) Bar.getPrototype();
				bar.setAsBarChart();
				c = bar;
				break;

			case ChartConstants.PIECHART: // Pie
				Pie p = (Pie) Pie.getPrototype();
				p.setAsPieChart();
				c = p;
				break;

			case ChartConstants.STOCKCHART:
				Line st = (Line) Line.getPrototype();
				st.setAsStockChart();
				c = st;
				break;

			case ChartConstants.LINECHART:
				Line l = (Line) Line.getPrototype();
				c = l;
				break;

			case ChartConstants.AREACHART:
				Area a = (Area) Area.getPrototype();
				c = a;
				break;

			case ChartConstants.SCATTERCHART:
				Scatter s = (Scatter) Scatter.getPrototype();
				s.setAsScatterChart();
				c = s;
				break;

			case ChartConstants.RADARCHART:
				Radar r = (Radar) Radar.getPrototype();
				c = r;
				break;

			case ChartConstants.SURFACECHART:
				Surface su = (Surface) Surface.getPrototype();
				c = su;
				break;

			case ChartConstants.DOUGHNUTCHART:
				Pie d = (Pie) Pie.getPrototype();
				d.setAsDoughnutChart();
				c = d;
				break;

			case ChartConstants.BUBBLECHART:
				Scatter bu = (Scatter) Scatter.getPrototype();
				bu.setAsBubbleChart();
				c = bu;
				break;

			case ChartConstants.RADARAREACHART:
				RadarArea ra = (RadarArea) RadarArea.getPrototype();
				c = ra;
				break;

			// note that, for the below chart types, the underlying type will be either COL or BAR
			// which actual chart is determined also by the bar shape
			case ChartConstants.PYRAMIDCHART:
				Bar pyramid = (Bar) Bar.getPrototype();
				pyramid.setAsColumnChart();
				cf.setBarShape( SHAPEPYRAMID );
				c = pyramid;
				break;

			case ChartConstants.CONECHART:
				Bar cone = (Bar) Bar.getPrototype();
				cone.setAsColumnChart();
				cf.setBarShape( SHAPECONE );
				c = cone;
				break;

			case ChartConstants.CYLINDERCHART:
				Bar cy = (Bar) Bar.getPrototype();
				cy.setAsColumnChart();
				cf.setBarShape( SHAPECYLINDER );
				c = cy;
				break;

			case ChartConstants.PYRAMIDBARCHART:
				Bar pb = (Bar) Bar.getPrototype();
				pb.setAsBarChart();
				cf.setBarShape( SHAPEPYRAMID );
				c = pb;
				break;

			case ChartConstants.CONEBARCHART:
				Bar cb = (Bar) Bar.getPrototype();
				cb.setAsBarChart();
				cf.setBarShape( SHAPECONE );
				c = cb;
				break;

			case ChartConstants.CYLINDERBARCHART:
				Bar cyb = (Bar) Bar.getPrototype();
				cyb.setAsBarChart();
				cf.setBarShape( SHAPECYLINDER );
				c = cyb;
				break;

			case ChartConstants.OFPIECHART:
				Boppop ofpie = (Boppop) Boppop.getPrototype();
				c = ofpie;
				break;
		}
		if( c != null )
		{
			c.setParentChart( parentchart );
		}
		return c;
	}

	public void setOptions( EnumSet<ChartOptions> options )
	{
		// FYI: The CrtLine (section 2.4.68) LineFormat (section 2.4.156) record pairs and the sequences of records that conform to the SHAPEPROPS rule (section 2.1.7.20.1) 
		// specify the drop lines, high-low lines, series lines, and leader lines for the chart (section 2.2.3.3).
		// NO 3d record for: bar of pie, bubble, doughnut, filled radar, pie of pie, radar, or scatter chart group.
		// Has Ser (Z) axis:  Surface, fStacked==0 & Line, Area, fStacked==0 && fClustered==0 && Bar (Col)  (Must also have ThreeD record)
		// 2 Value Axes:  Scatter, Bubble
		// NO Axes:  Pie, Doughnut, PieOfPie, BarOfPie

		ChartAxes ca = this.getParentChart().getAxes();
		int chartType = this.getChartType();
		if( options.contains( ChartOptions.STACKED ) )            // bar/col types, line, area
		{
			chartobj.setIsStacked( true );
		}

		if( options.contains( ChartOptions.PERCENTSTACKED ) )    // bar/col types, line, area
		{
			chartobj.setIs100Percent( true );
		}

		if( options.contains( ChartOptions.CLUSTERED ) )        // bar/col only
		{
			cf.setIsClustered( true );
		}

		if( options.contains( ChartOptions.SERLINES ) )        // bar, line, stock
		{
			cf.addChartLines( ChartLine.TYPE_SERIESLINE );
		}

		if( options.contains( ChartOptions.HILOWLINES ) )        // bar, OfPie
		{
			cf.addChartLines( ChartLine.TYPE_HILOWLINE );
		}

		if( options.contains( ChartOptions.DROPLINES ) )        // Surface chart
		{
			cf.addChartLines( ChartLine.TYPE_DROPLINE );
		}

		if( options.contains( ChartOptions.UPDOWNBARS ) )        // line, area,stock
		{
			cf.addUpDownBars();
		}

		if( options.contains( ChartOptions.HASLINES ) )        // line, scatter ...
		{
			cf.setHasLines();
		}

		if( options.contains( ChartOptions.SMOOTHLINES ) )        // line, scatter, radar
		{
			cf.setHasSmoothLines( true );
		}
// HANDLE FILLED for radar
// HANDLE bubble 3d
// HANDLE bar shapes ***    	

		cf.setChartObject( chartobj );
		boolean use3Ddefaults = true;            // init 3D record with default values for specific chart type
		ThreeD threeD = cf.getThreeDRec( false );
		if( threeD == null )
		{
			if( options.contains( ChartOptions.THREED ) || (chartType == SURFACECHART) )
			{    // surface charts ALWAYS have a 3 record as does pyramid, cone and cylinder charts
				if( (chartType != BUBBLECHART) && (chartType != SCATTERCHART) )    // supposed to be also donught, radar as well ...
				{
					threeD = this.initThreeD( chartType );
				}
				else if( chartType == BUBBLECHART )
				{ // scatter charts have no 3d option
					cf.setHas3DBubbles( true );
				}
			}
		}
		else    // 3D record already set (via OOXML) - do not use defaults
		{
			use3Ddefaults = false;
		}
		switch( chartType )
		{
			case BARCHART:
			case COLCHART:
				if( use3Ddefaults && (threeD != null) )
				{
					threeD.setChartOption( "AnRot", "20" );
					threeD.setChartOption( "AnElev", "15" );
					threeD.setChartOption( "TwoDWalls", "true" );
					threeD.setChartOption( "ThreeDScaling", "false" );
					threeD.setChartOption( "Cluster", "false" );
					threeD.setChartOption( "PcDepth", "100" );
					threeD.setChartOption( "PcDist", "30" );
					threeD.setChartOption( "PcGap", "150" );
					threeD.setChartOption( "PcHeight", "72" );    // ??????
					threeD.setChartOption( "Perspective", "false" );
					if( options.contains( ChartOptions.CLUSTERED ) )
					{
						threeD.setChartOption( "Cluster", "true" );
						threeD.setChartOption( "PcHeight", "62" );    // ??????
						if( chartType == COLCHART )
						{
							threeD.setChartOption( "Perspective", "true" );
						}
						else
						{    // bar chart
							threeD.setChartOption( "Perspective", "false" );
							threeD.setChartOption( "ThreeDScaling", "true" );
							threeD.setChartOption( "PcHeight", "150" );    // ??????
						}
					}
				}
				ca.setChartOption( XAXIS, "AddArea", "true" );
				ca.setChartOption( YAXIS, "AddArea", "true" );
				if( !isStacked() && !isClustered() )
				{
					ca.createAxis( ZAXIS );
					ca.setChartOption( ZAXIS, "AddArea", "true" );
				}
				break;
			case LINECHART:
				if( options.contains( ChartOptions.THREED ) )
				{
					ca.setChartOption( XAXIS, "AddArea", "true" );
					ca.setChartOption( YAXIS, "AddArea", "true" );
					if( !isStacked() )
					{
						ca.createAxis( ZAXIS );
						ca.setChartOption( ZAXIS, "AddArea", "true" );
					}
				}
				break;
			case STOCKCHART:
				cf.addChartLines( ChartLine.TYPE_HILOWLINE );
				cf.setMarkers( 0 );
				break;
			case SCATTERCHART:
				if( !options.contains( ChartOptions.HASLINES ) )
				{
					cf.setHasLines( 5 );    // no line style -- doesn't appear to work
				}
				break;
			case AREACHART:
				if( use3Ddefaults && (threeD != null) )
				{
					threeD.setChartOption( "AnRot", "20" );
					threeD.setChartOption( "TwoDWalls", "true" );
					threeD.setChartOption( "ThreeDScaling", "false" );
					threeD.setChartOption( "Perspective", "true" );
				}
				this.setChartOption( "Percentage", "25" );
				this.setChartOption( "SmoothedLine", "true" );
				ca.setChartOption( XAXIS, "CrossBetween", "false" );
				if( options.contains( ChartOptions.THREED ) )
				{
					ca.setChartOption( XAXIS, "AddArea", "true" );
					ca.setChartOption( YAXIS, "AddArea", "true" );
					if( !isStacked() )
					{
						ca.createAxis( ZAXIS );
						ca.setChartOption( ZAXIS, "AddArea", "true" );
					}
				}
				break;
			case BUBBLECHART:
				if( options.contains( ChartOptions.THREED ) )
				{
					this.setChartOption( "Percentage", "25" );
					this.setChartOption( "SmoothedLine", "true" );
					this.setChartOption( "ThreeDBubbles", "true" );
				}
				break;
			case PIECHART: // Pie
				cf.setVaryColor( true ); // Should be true for all pie charts ...
				if( options.contains( ChartOptions.EXPLODED ) )
				{
					this.setChartOption( "SmoothedLine", "true" );
					this.setChartOption( "Percentage", "25" );
				}
				if( use3Ddefaults && options.contains( ChartOptions.THREED ) )
				{
					this.setChartOption( "AnRot", "236" );
				}
				ca.removeAxes();
				break;
			case DOUGHNUTCHART:
				if( options.contains( ChartOptions.EXPLODED ) )
				{
					this.setChartOption( "SmoothedLine", "true" );
					this.setChartOption( "Percentage", "25" );
				}
				cf.setVaryColor( true ); // Should be true for all pie charts ...
				ca.removeAxes();
				break;
			case SURFACECHART:    // NOTE: For Surface charts, non-threeD==Contour
				if( use3Ddefaults && (threeD != null) )
				{    // shouldn't
					threeD.setChartOption( "Cluster", "false" );
					threeD.setChartOption( "TwoDWalls", "true" );
					threeD.setChartOption( "ThreeDScaling", "true" );
					threeD.setChartOption( "Perspective", "true" );
				}
				ca.setChartOption( XAXIS, "AddArea", "true" );
				ca.setChartOption( XAXIS, "AreaFg", "8" );
				ca.setChartOption( XAXIS, "AreaBg", "78" );
				ca.setChartOption( YAXIS, "AddArea", "true" );
				ca.setChartOption( YAXIS, "AreaFg", "22" );
				ca.setChartOption( YAXIS, "AreaBg", "78" );
				ca.createAxis( ZAXIS );
				if( use3Ddefaults && options.contains( ChartOptions.THREED ) )
				{    // "regular" 3d surface
					threeD.setChartOption( "AnElev", "15" );
					threeD.setChartOption( "AnRot", "20" );
					threeD.setChartOption( "PcDepth", "100" );
					threeD.setChartOption( "PcDist", "30" );
					threeD.setChartOption( "PcGap", "150" );
					threeD.setChartOption( "PcHeight", "50" );    // ??????
				}
				else if( use3Ddefaults )
				{                                        // contour (non-3d)
					threeD.setChartOption( "AnElev", "90" );
					threeD.setChartOption( "AnRot", "0" );
					threeD.setChartOption( "PcDepth", "100" );
					threeD.setChartOption( "PcDist", "0" );
					threeD.setChartOption( "PcGap", "150" );
					threeD.setChartOption( "PcHeight", "50" );    // ??????
				}
				if( options.contains( ChartOptions.WIREFRAME ) )
				{
					((Surface) chartobj).setIsWireframe( true );
				}
				if( !options.contains( ChartOptions.WIREFRAME ) && options.contains( ChartOptions.THREED ) )
				{
					chartobj.setChartOption( "ColorFill", "true" );
					chartobj.setChartOption( "Shading", "true" );
				}
				else if( options.contains( ChartOptions.WIREFRAME ) )
				{    // conotur (flat, non-3d) wireframe
					chartobj.setChartOption( "Shading", "true" );
				}
				else                                            // contour filled (i.e. plain Surface)
				{
					chartobj.setChartOption( "ColorFill", "true" );
				}
				break;
			case RADARCHART:
				if( options.contains( ChartOptions.FILLED ) )
				{
					((RadarChart) this).setFilled( true );
				}
				break;
			case PYRAMIDCHART:
			case CONECHART:
			case CYLINDERCHART:
			case PYRAMIDBARCHART:
			case CONEBARCHART:
			case CYLINDERBARCHART:
				// Shaped Bar/Col charts are all 3d 
				if( threeD == null )
				{
					threeD = this.initThreeD( chartType );
				}
				if( use3Ddefaults )
				{
					threeD.setChartOption( "AnElev", "15" );
					threeD.setChartOption( "AnRot", "20" );
					threeD.setChartOption( "Cluster", "true" ); // only for regular pyramid charts; stacked, etc. will alter
					threeD.setChartOption( "TwoDWalls", "true" );
					if( !options.contains( ChartOptions.THREED ) )
					{
						threeD.setChartOption( "ThreeDScaling", "false" );
						threeD.setChartOption( "Cluster", "true" );
					}
					threeD.setChartOption( "Perspective", "false" );
					threeD.setChartOption( "PcDepth", "100" );
					threeD.setChartOption( "PcDist", "30" );
					threeD.setChartOption( "PcGap", "150" );
					threeD.setChartOption( "PcHeight", "52" );    // ??????
				}
				if( (chartType == PYRAMIDCHART) || (chartType == PYRAMIDBARCHART) )
				{
					cf.setBarShape( ChartConstants.SHAPEPYRAMID );
				}
				else if( (chartType == CONECHART) || (chartType == CONEBARCHART) )
				{
					cf.setBarShape( ChartConstants.SHAPECONE );
				}
				else if( (chartType == CYLINDERCHART) || (chartType == CYLINDERBARCHART) )
				{
					cf.setBarShape( ChartConstants.SHAPECYLINDER );
				}
				break;
		}
	}

	/**
	 * parse the chart object OOXML element (barchart, area3DChart, etc.) and create the corresponding chart type objects
	 *
	 * @param xpp         XmlPullParser
	 * @param wbh         workBookHandle
	 * @param parentChart parent Chart Object
	 * @param nChart      chart grouping number, 0 for default, 1-9 for overlay
	 * @return
	 */
	public static ChartType parseOOXML( XmlPullParser xpp, WorkBookHandle wbh, Chart parentChart, int nChart )
	{
		try
		{
			String endTag = xpp.getName();
			String tnm = xpp.getName();
			int eventType = xpp.getEventType();
			int chartType = BARCHART;
			ChartAxes ca = parentChart.getAxes();

			java.util.Stack<String> lastTag = new java.util.Stack();        // keep track of element hierarchy
			ChartFormat cf = parentChart.getChartOjectParent( nChart );

			EnumSet<ChartOptions> options = EnumSet.noneOf( ChartOptions.class );    // chart-specific options such as threed, stacked ...
			if( (tnm.equals( "bubble3D" )) )
			{    // bubble3D tag appears for each series in 3D bubble chart
				options.add( ChartOptions.THREED );
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.BARCHART] ) )
			{
				chartType = BARCHART;
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.LINECHART] ) )
			{
				chartType = LINECHART;
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.PIECHART] ) )
			{
				chartType = PIECHART;
				parentChart.getAxes().removeAxes();
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.AREACHART] ) )
			{
				chartType = AREACHART;
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.SCATTERCHART] ) )
			{
				chartType = SCATTERCHART;
				ca.removeAxis( XAXIS );
				ca.createAxis( XVALAXIS );
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.RADARCHART] ) )
			{
				chartType = RADARCHART;
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.SURFACECHART] ) )
			{
				chartType = SURFACECHART;
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.DOUGHNUTCHART] ) )
			{
				chartType = DOUGHNUTCHART;
				parentChart.getAxes().removeAxes();
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.BUBBLECHART] ) )
			{
				chartType = BUBBLECHART;
				ca.removeAxis( XAXIS );
				ca.createAxis( XVALAXIS );
			}
			else if( tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.BARCHART] ) )
			{
				chartType = BARCHART;
				options.add( ChartOptions.THREED );
			}
			else if( tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.LINECHART] ) )
			{
				chartType = LINECHART;
				options.add( ChartOptions.THREED );
			}
			else if( tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.PIECHART] ) )
			{
				chartType = PIECHART;
				options.add( ChartOptions.THREED );
				parentChart.getAxes().removeAxes();
			}
			else if( tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.AREACHART] ) )
			{
				chartType = AREACHART;
				options.add( ChartOptions.THREED );
			}
			else if( tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.SCATTERCHART] ) )
			{
				chartType = SCATTERCHART;
				ca.removeAxis( XAXIS );
				ca.createAxis( XVALAXIS );
				options.add( ChartOptions.THREED );
			}
			else if( tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.RADARCHART] ) )
			{
				chartType = RADARCHART;
				options.add( ChartOptions.THREED );
			}
			else if( tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.SURFACECHART] ) )
			{
				chartType = SURFACECHART;
				options.add( ChartOptions.THREED );
			}
			else if( tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.DOUGHNUTCHART] ) )
			{
				chartType = DOUGHNUTCHART;
				options.add( ChartOptions.THREED );
				parentChart.getAxes().removeAxes();
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.OFPIECHART] ) )
			{
				chartType = OFPIECHART;
				ca.removeAxis( XAXIS );
				ca.removeAxis( YAXIS );
			}
			else if( tnm.equals( OOXMLConstants.twoDchartTypes[ChartHandle.STOCKCHART] ) )
			{
				chartType = STOCKCHART;
			}

			GenericChartObject co = ChartType.createUnderlyingChartObject( chartType,
			                                                               parentChart,
			                                                               cf );    //, ((ChartType)chartobj.get(0)));
			cf.setChartObject( co );    // sets the chart format (parent of chart item) to the specific chart item
			// exception for surface charts in 3d
			if( (chartType == SURFACECHART) && (tnm.equals( OOXMLConstants.threeDchartTypes[ChartHandle.SURFACECHART] )) )
			{
				((Surface) co).setIs3d( true );
			}
			ChartType ct = ChartType.createChartTypeObject( co, cf, parentChart.getWorkBook() );
			parentChart.addChartType( ct, nChart );

			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					tnm = xpp.getName();
					lastTag.push( tnm );
					String v = null;
					try
					{
						v = xpp.getAttributeValue( 0 );
					}
					catch( IndexOutOfBoundsException e )
					{
					}
// TODO: 
// hasMarkers ****		             
// dLbls		             
// smooth	-- line
// marker -- 	line
// bandfmts -- surface, surface3d
// bubble3D,
// bubbleScale
// showNegBubbles
// sizeRepresents		         
// custSplit	-- OfPie
					if( tnm.equals( "grouping" ) )
					{    //This element specifies the type of grouping for a column, line, or area chart + 3d versions
						if( v.equals( "stacked" ) )
						{
							co.setIsStacked( true );
						}
						else if( v.equals( "percentStacked" ) )
						{
							co.setIs100Percent( true );
						}
						else if( v.equals( "clustered" ) )
						{
							cf.setIsClustered( true );        // bar/col only
						}
						else if( v.equals( "standard" ) )
						{    // for Line, Line3d, Area, Area3d and Bar3d, Col3d - does not appear valid for Bar 2d and Col 2d
							co.setIsStacked( false );
							co.setIs100Percent( false );
						}
					}
					else if( tnm.equals( "barDir" ) )
					{    //
						if( v.equals( "col" ) )
						{
							((Bar) co).setAsColumnChart();
							ct = ChartType.createChartTypeObject( co, cf, parentChart.getWorkBook() );
							parentChart.addChartType( ct, nChart );
						}
					}
					else if( tnm.equals( "shape" ) )
					{    // bar3d
						cf.setBarShape( ct.convertShape( v ) );
						if( ct.defaultShape != 0 )
						{
							ct = ChartType.createChartTypeObject( co, cf, parentChart.getWorkBook() );
							parentChart.addChartType( ct, nChart );
						}
					}
					else if( tnm.equals( "radarStyle" ) )
					{
						if( v.equals( "filled" ) )
						{
							((RadarChart) ct).setFilled( true );
						}
						else if( v.equals( "marker" ) )
						{

						}
					}
					else if( tnm.equals( "wireframe" ) )
					{    // surface
						((Surface) co).setIsWireframe( (v != null) && v.equals( "1" ) );
					}
					else if( tnm.equals( "scatterStyle" ) )
					{
						if( v.equals( "lineMarker" ) )
						{
							cf.setHasLines();
						}
						else if( v.equals( "smoothMarker" ) )
						{
							cf.setHasSmoothLines( true );
						}
						else if( v.equals( "marker" ) )
						{
							; //cf.seth;
						}
						else if( v.equals( "line" ) )
						{
							cf.setHasLines();
						}
						else if( v.equals( "smooth" ) )
						{
							cf.setHasSmoothLines( true );
						}
					}
					else if( tnm.equals( "varyColors" ) )
					{
						cf.setVaryColor( xpp.getAttributeValue( 0 ).equals( "1" ) );
					}
					else if( tnm.equals( "dropLines" ) )
					{
						ChartLine cl = cf.addChartLines( ChartLine.TYPE_DROPLINE );
						cl.parseOOXML( xpp, lastTag, cf, wbh );
					}
					else if( tnm.equals( "hiLowLines" ) )
					{
						ChartLine cl = cf.addChartLines( ChartLine.TYPE_HILOWLINE );
						cl.parseOOXML( xpp, lastTag, cf, wbh );
					}
					else if( tnm.equals( "upDownBars" ) )
					{
						cf.parseUpDownBarsOOXML( xpp, lastTag, wbh );
					}
					else if( tnm.equals( "serLines" ) )
					{
						ChartLine cl = cf.addChartLines( ChartLine.TYPE_SERIESLINE );
						cl.parseOOXML( xpp, lastTag, cf, wbh );
					}
					else if( tnm.equals( "overlap" ) )
					{    // bar
						co.setChartOption( "Overlap", v );
					}
					else if( tnm.equals( "gapWidth" ) )
					{
						co.setChartOption( "Gap", v );        // bar
					}
					else if( tnm.equals( "ofPieType" ) )
					{
						((Boppop) co).setIsPieOfPie( "pie".equals( v ) );
					}
					else if( tnm.equals( "gapDepth" ) )
					{
						cf.setGapDepth( Integer.valueOf( v ) );    // bar3d, area3d, line3d
					}
					else if( tnm.equals( "firstSliceAn" ) )
					{
						((Pie) co).setAnStart( Integer.valueOf( v ) );        // pie or doughnut
					}
					else if( tnm.equals( "holeSize" ) )
					{
						((Pie) co).setDoughnutSize( Integer.valueOf( v ) );        // pie or doughnut
					}
					else if( tnm.equals( "secondPieSize" ) )
					{
						((Boppop) co).setSecondPieSize( Integer.valueOf( v ) );    // OfPie (Pie of Pie, Bar of Pie)
					}
					else if( tnm.equals( "splitType" ) )
					{
						((Boppop) co).setSplitType( v );    // OfPie (Pie of Pie, Bar of Pie)
					}
					else if( tnm.equals( "splitPos" ) )
					{
						((Boppop) co).setSplitPos( Integer.valueOf( v ) );    // OfPie (Pie of Pie, Bar of Pie)
					}
					else if( tnm.equals( "ser" ) )
					{
						Series s = ChartSeries.parseOOXML( xpp, wbh, ct, false, lastTag );
					}
					else if( tnm.equals( "dLbls" ) )
					{
					}
					else if( tnm.equals( "marker" ) )
					{        // line only?
//				       	 	m= (Marker) Marker.parseOOXML(xpp, lastTag).cloneElement();			       	 	
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( endTag ) )
					{
						break;
					}

				}
				eventType = xpp.next();
			}
			return ct;

//		    lastTag.pop();	// chart type tag will be added in parseOOXML     	
//mychart.getChartObject(nChart).parseOOXML(xpp, this.wbh, lastTag);
		}
		catch( Exception e )
		{
			Logger.logErr( "ChartType.parseChartType: " + e.toString() );
		}
		return null;
	}

	/**
	 * return data label options for each series as an int array
	 * <br>each can be one or more of:
	 * <br>VALUELABEL= 0x1;
	 * <br>VALUEPERCENT= 0x2;
	 * <br>CATEGORYPERCENT= 0x4;
	 * <br>SMOOTHEDLINE= 0x8;
	 * <br>CATEGORYLABEL= 0x10;
	 * <br>BUBBLELABEL= 0x20;
	 * <br>SERIESLABEL= 0x40;
	 *
	 * @return int array
	 * @see AttachedLabel
	 */
	protected int[] getDataLabelInts()
	{
		return chartobj.getParentChart().getDataLabelsPerSeries( cf.getDataLabelsInt() );    // data label options, if any, per series
	}

	/**
	 * return the default data label setting for the chart, if any
	 * <br>NOTE: each series can override the default data label for the chart
	 * <br>can be one or more of:
	 * <br>VALUELABEL= 0x1;
	 * <br>VALUEPERCENT= 0x2;
	 * <br>CATEGORYPERCENT= 0x4;
	 * <br>SMOOTHEDLINE= 0x8;
	 * <br>CATEGORYLABEL= 0x10;
	 * <br>BUBBLELABEL= 0x20;
	 * <br>SERIESLABEL= 0x40;
	 *
	 * @return int default data label for chart
	 */
	protected int getDataLabel()
	{
		return cf.getDataLabelsInt();
	}

	/**
	 * return an array of the type of markers for each series:
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
	protected int[] getMarkerFormats()
	{
		int mf = cf.getMarkerFormat();
		int[] markers = new int[getParentChart().getAllSeries( getParentChart().getChartOrder( this ) ).size()];
		if( mf > 0 )
		{
			for( int i = 0; i < markers.length; i++ )
			{
				markers[i] = mf;
			}
		}
		return markers;
	}

	/**
	 * specifies whether
	 * the color for each data point and the color and type for each data marker
	 * vary
	 *
	 * @param b
	 */
	protected void setVaryColor( boolean b )
	{
		cf.setVaryColor( b );
	}

	/**
	 * returns true if this chart has smoothed lines
	 *
	 * @return
	 */
	public boolean getHasSmoothLines()
	{
		return cf.getHasSmoothLines();
	}

	public void setHasSmoothLines( boolean b )
	{
		cf.setHasSmoothLines( b );
	}

	/**
	 * returns the chart type for the default chart
	 */
	public int getChartType()
	{
		return chartobj.chartType;
	}

	public void addLegend( Legend l )
	{
		legend = l;
	}

	public String getSVG()
	{
		return null;
	}

	public String getJSON()
	{
		return null;
	}

	public JSONObject getOptionsJSON()
	{
		return null;
	}

	/**
	 * gets the chart-type specific OOXML representation (representing child element of plotArea element)
	 *
	 * @return
	 */
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		return null;
	}

	public JSONObject getJSON( ChartSeries s, WorkBookHandle wbh, Double[] minMax ) throws JSONException
	{
		JSONObject chartObjectJSON = new JSONObject();

		// Type JSON
		chartObjectJSON.put( "type", this.getTypeJSON() );

		// Deal with Series
		double yMax = 0.0, yMin = 0.0;
		int nSeries = 0;
		JSONArray seriesJSON = new JSONArray();
		JSONArray seriesCOLORS = new JSONArray();
		try
		{
			ArrayList series = s.getSeriesRanges();
			String[] scolors = s.getSeriesBarColors();
			for( int i = 0; i < series.size(); i++ )
			{
				JSONArray seriesvals = CellRange.getValuesAsJSON( series.get( i ).toString(), wbh );
				// must trap min and max for axis tick and units
				nSeries = Math.max( nSeries, seriesvals.length() );
				for( int j = 0; j < seriesvals.length(); j++ )
				{
					try
					{
						yMax = Math.max( yMax, seriesvals.getDouble( j ) );
						yMin = Math.min( yMin, seriesvals.getDouble( j ) );
					}
					catch( NumberFormatException n )
					{
						;
					}
				}
				seriesJSON.put( seriesvals );
				seriesCOLORS.put( scolors[i] );
			}
			chartObjectJSON.put( "Series", seriesJSON );
			chartObjectJSON.put( "SeriesFills", seriesCOLORS );
		}
		catch( JSONException je )
		{
			// TODO: Log error
		}
		minMax[0] = yMin;
		minMax[1] = yMax;
		minMax[2] = (double) nSeries;
		return chartObjectJSON;
	}

	/**
	 * return Type JSON for generic chart types
	 *
	 * @return
	 * @throws JSONException
	 */
	public JSONObject getTypeJSON() throws JSONException
	{
		JSONObject typeJSON = new JSONObject();
		typeJSON.put( "type", "Default" );
		return typeJSON;
	}

	/**
	 * returns SVG to represent the actual chart object (BAR, LINE, AREA, etc.)
	 *
	 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @param axisMetrics  maps specific axis options such as xAxisReversed, xPattern ...
	 * @param s            ChartSeries - holds legends, categories, seriesdata ...
	 * @return String svg
	 */
	public String getSVG( HashMap<String, Double> chartMetrics, HashMap<String, Object> axisMetrics, ChartSeries s )
	{
		return "";
	}

	/**
	 * returns the SVG-ready data label for the given set of data label options
	 * TODO: get separator character
	 *
	 * @param datalabel  int[] Data label options (indexed by current series # s)
	 *                   <br>can be one or more of:
	 *                   <br>VALUELABEL= 0x1;
	 *                   <br>VALUEPERCENT= 0x2;
	 *                   <br>CATEGORYPERCENT= 0x4;
	 *                   <br>SMOOTHEDLINE= 0x8;
	 *                   <br>CATEGORYLABEL= 0x10;
	 *                   <br>BUBBLELABEL= 0x20;
	 *                   <br>SERIESLABEL= 0x40;
	 * @param series     ArrayList of series and category data: structure:  category data, series 0-n data, series colors, series Labels (Eventually will be an object)
	 * @param val        double current series value
	 * @param percentage double percentage value (pie charts only)
	 * @param s          int current series #
	 * @param cat        string current cat
	 * @return
	 */
	public String getSVGDataLabels( int[] datalabels,
	                                HashMap<String, Object> axisMetrics,
	                                double val,
	                                double percentage,
	                                int s,
	                                String[] legends,
	                                String cat )
	{

		if( s >= datalabels.length )    // can happen with Pie-style charts
		{
			return null;
		}
		boolean showValueLabel = (datalabels[s] & AttachedLabel.VALUELABEL) == AttachedLabel.VALUELABEL;
		boolean showValuePercent = (datalabels[s] & AttachedLabel.VALUEPERCENT) == AttachedLabel.VALUEPERCENT;
		boolean showCatPercent = (datalabels[s] & AttachedLabel.CATEGORYPERCENT) == AttachedLabel.CATEGORYPERCENT;
		boolean showCategories = (datalabels[s] & AttachedLabel.CATEGORYLABEL) == AttachedLabel.CATEGORYLABEL;
		boolean showBubbleLabel = (datalabels[s] & AttachedLabel.BUBBLELABEL) == AttachedLabel.BUBBLELABEL;
		boolean showValue = (datalabels[s] & AttachedLabel.VALUE) == AttachedLabel.VALUE;
		if( showValue || showCategories || showValueLabel || showValuePercent || showBubbleLabel )
		{
			String l = "";
			if( showValueLabel )
			{
				l += legends[s] + " ";    // series names
			}
			if( showCategories )
			{
				l += CellFormatFactory.fromPatternString( (String) axisMetrics.get( "xPattern" ) ).format( cat ) + " ";    // categories
			}
			if( showValue || showBubbleLabel )
			{
				l += CellFormatFactory.fromPatternString( (String) axisMetrics.get( "yPattern" ) ).format( String.valueOf( val ) );
				/*try {
					int v= new Double(val).intValue(); 
					if (v==val)	
						l+=v + " ";
					else
						l+=val + " ";
				} catch (Exception e) {
					l+= val + " ";
				}*/
			}
			if( showValuePercent )
			{
				l += (int) Math.round( percentage * 100 ) + "%";
			}
			return l;
		}
		return null;
	}

	/** Show or remove Data Table for Chart
	 *  NOTE:  METHOD IS STILL EXPERIMENTAL
	 * @param bShow
	 *
	public void showDataTable(boolean bShow) {
	int i= Chart.findRecPosition(chartArr, Dat.class);
	if (bShow) {
	if (i==-1) { // add Dat
	Dat d= (Dat)Dat.getPrototype(true); // create data table
	i= Chart.findRecPosition(chartArr, AxisParent.class);
	this.chartArr.add(++i, d);
	}
	} else if (i > 0) {
	chartArr.remove(i);	// remove Dat - Data Table options + all associated recs
	}
	}*/

	/**
	 * return truth if Chart has a data legend key showing
	 *
	 * @return
	 */
	public boolean hasDataLegend()
	{
		return (legend != null);
	}

	/**
	 * return the data legend for this chart
	 *
	 * @return
	 */
	public Legend getDataLegend()
	{
		return legend;
	}

	/**
	 * show or hide chart legend key
	 *
	 * @param bShow    boolean show or hide
	 * @param vertical boolean show as vertical or horizontal
	 */
	public void showLegend( boolean bShow, boolean vertical )
	{
		if( bShow && (legend == null) )
		{
			legend = Legend.createDefaultLegend( wb );
			legend.setParentChart( chartobj.getParentChart() );
			for( int j = 0; j < legend.chartArr.size(); j++ )
			{
				((GenericChartObject) legend.chartArr.get( j )).setParentChart( chartobj.getParentChart() );
			}
			legend.setVertical( vertical );
			cf.chartArr.add( legend );
		}
		else if( bShow )
		{
			legend.setVertical( vertical );
		}
		else if( legend != null )
		{
			int i = Chart.findRecPosition( cf.chartArr, Legend.class );
			cf.chartArr.remove( i );
			legend = null;
		}
	}

	/**
	 * return the Data Labels chosen for this chart, if any
	 * can be one or more of:
	 * <br>Value
	 * <br>ValuePerecentage
	 * <br>CategoryPercentage
	 * <br>CategoryLabel
	 * <br>BubbleLabel
	 * <br>SeriesLabel
	 * or an empty string if no data labels are chosen for the chart
	 *
	 * @return
	 */
	public String getDataLabels()
	{
		return cf.getDataLabels();
	}

	/**
	 * return data label options as an int
	 * <br>can be one or more of:
	 * <br>VALUELABEL= 0x1;
	 * <br>VALUEPERCENT= 0x2;
	 * <br>CATEGORYPERCENT= 0x4;
	 * <br>SMOOTHEDLINE= 0x8;
	 * <br>CATEGORYLABEL= 0x10;
	 * <br>BUBBLELABEL= 0x20;
	 * <br>SERIESLABEL= 0x40;
	 *
	 * @return a combination of data label options above or 0 if none
	 * @see AttachedLabel
	 */
	public int getDataLabelsInt()
	{
		return cf.getDataLabelsInt();
	}

	/**
	 * returns the bar shape for a column or bar type chart
	 * can be one of:
	 * <br>ChartConstants.SHAPECOLUMN	default
	 * <br>ChartConstants.SHAPECONEd
	 * <br>ChartConstants.SHAPECONETOMAX
	 * <br>ChartConstants.SHAPECYLINDER
	 * <br>ChartConstants.SHAPEPYRAMID
	 * <br>ChartConstants.SHAPEPYRAMIDTOMAX
	 *
	 * @return int bar shape
	 */
	public int getBarShape()
	{
		return SHAPEDEFAULT;
	}

	/**
	 * returns type of marker for this chart, if any
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
	 *
	 * @return int marker type
	 */
	public int getMarkerFormat()
	{
		return cf.getMarkerFormat();
	}

	/**
	 * returns true if this chart has lines (see Scatter, Line Charts amongst others)
	 *
	 * @return
	 */
	public boolean getHasLines()
	{
		return cf.getHasLines();
	}

	/**
	 * returns true if this chart has drop lines
	 *
	 * @return
	 */
	public boolean getHasDropLines()
	{
		return cf.getHasDropLines();
	}

	/**
	 * sets this chart (Area, Line, Scatter) to have drop lines
	 */
	public void setHasDropLines()
	{
		cf.setHasDropLines();
	}

	/**
	 * look up a generic chart option
	 *
	 * @param op
	 * @return
	 */
	public String getChartOption( String op )
	{
		String ret = chartobj.getChartOption( op );    // if not a chart-specific option, see if it's more generic
		if( ret == null )
		{
			return cf.getChartOption( op );
		}
		return ret;
	}

	/**
	 * return chart-type-specific options in XML form
	 *
	 * @param nChart 0=default, 1-9= overlay charts
	 * @return String XML
	 */
	public String getChartOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		sb.append( chartobj.getOptionsXML() );    // chart-type-specific
		sb.append( cf.getChartOptionsXML() );    // governs threed settings and other misc. options
		return sb.toString();
	}

	/**
	 * interface for setting chart-type-specific options
	 * in a generic fashion
	 *
	 * @param op  option
	 * @param val value
	 * @see ExtenXLS.handleChartElement
	 * @see ChartHandle.getXML
	 */
	public boolean setChartOption( String op, String val )
	{
		if( !chartobj.setChartOption( op, val ) )  // if not handled,
		{
			cf.setOption( op, val );
		}
		return true;
	}

	/**
	 * @return truth of "Chart is Three-D"
	 */
	public boolean isThreeD()
	{
		return cf.isThreeD( chartobj.chartType );
	}

	/**
	 * return ThreeD settings for this chart in XML form
	 *
	 * @return String XML
	 */
	public String getThreeDXML()
	{
		return cf.getThreeDXML();
	}

	/**
	 * returns the 3d record of the desired chart
	 *
	 * @param chartType one of the chart type constants
	 * @return
	 */
	public ThreeD initThreeD( int charttype )
	{
/* TODO: test if this is necessary        AxisParent ap = this.getAxisParent();
    	try {
        	// first thing, remove the PlotArea and Frame - don't know why but is necessary!! 
    		int x= Chart.findRecPosition(ap.chartArr, PlotArea.class);
    		if (x!=-1) ap.chartArr.remove(x);
    		x= Chart.findRecPosition(ap.chartArr, Frame.class);
    		if (x!=-1) ap.chartArr.remove(x);
    	} catch (Exception e) {}*/
		ThreeD td = cf.getThreeDRec( true );
		td.setIsPie( (charttype == PIECHART) || (charttype == DOUGHNUTCHART) );
		return td;
	}

	/**
	 * return the 3d record of the chart
	 * <br>Creates if not present, if bCreate is true
	 *
	 * @return
	 */
	public ThreeD getThreeDRec( boolean bCreate )
	{
		return cf.getThreeDRec( bCreate );
	}

	/**
	 * @return truth of "Chart is Stacked"
	 */
	public boolean isStacked()
	{
		return chartobj.isStacked();
	}

	/**
	 * return truth of "Chart is 100% Stacked"
	 *
	 * @return
	 */
	public boolean is100PercentStacked()
	{
		return (chartobj.is100Percent());
	}

	public void setIsStacked( boolean isstacked )
	{
		chartobj.setIsStacked( isstacked );
	}

	public void setIs100Psercent( boolean ispercentage )
	{
		chartobj.setIs100Percent( ispercentage );
	}

	/**
	 * @return truth of "Chart is Clustered"  (Bar/Col only)
	 */
	public boolean isClustered()
	{
		return false;
	}

	public void addLegend()
	{
		if( legend == null )
		{
			// TODO: ADD LEGEND to cf
			//int i = Chart.findRecPosition(cf.chartArr, Legend.class);
		}
	}

	// SVG Convenience Methods
	public static String getScript( String range )
	{
		return "onmouseover='highLight(evt); showRange(\"" + range + "\");' onclick='handleClick(evt);' onmouseout='restore(evt); hideRange();'";
//		return "onmouseover='highLight(evt);' onclick='handleClick(evt);' onmouseout='restore(evt);'";
	}

	public static String getFillOpacity()
	{
		return ".75";
	}

	public static String getTextColor()
	{
		return "#222222";
	}

	public static String getLightColor()
	{
		return "#CCCCCC";
	}

	public static String getMediumColor()
	{
		return "#555555";
	}

	public static String getDarkColor()
	{
		return "#333333";
	}

	/**
	 * returns the SVG for the font style of this object
	 */
	protected String getFontSVG()
	{
		return getFontSVG( -1 );
	}

	/**
	 * Returns SVG used to define font for data labels
	 * TODO: read correct value from chart recs
	 *
	 * @return
	 */
	public static String getDataLabelFontSVG()
	{
		int sz = 9;
		return "font-family='Arial' font-size='" + sz + "' fill='" + ChartType.getDarkColor() + "' ";
	}

	/**
	 * returns the SVG for the font style of this object
	 *
	 * @param stroke size in pt
	 */
	public static String getStrokeSVG()
	{
		return getStrokeSVG( 1f, getMediumColor() );
	}

	/**
	 * returns the SVG for the font style of this object
	 *
	 * @param stroke size in pt
	 * @param String stroke color in HTML format
	 */
	public static String getStrokeSVG( float sz, String strokeclr )
	{
		String stk = " stroke='" + strokeclr + "'  stroke-opacity='1' stroke-width='" + sz + "' stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'";
		return stk;
	}

	/**
	 * returns the SVG for the font style of this object
	 *
	 * @param font size in pt
	 */
	protected String getFontSVG( int sz )
	{
//		com.extentech.formats.XLS.Font f = this.getFont();	// this is not correct!
		//	if(f == null) // return a default
		return "font-family='Arial' font-size='" + sz + "' fill='" + getDarkColor() + "' ";

		//return f.getSVG();
	}

	/**
	 * Convert the user-friendly OOXML shape string to 2003-v int shape flag
	 *
	 * @param shape
	 * @return
	 */
	public int convertShape( String shape )
	{
		defaultShape = 0;
		if( shape.equals( "box" ) )
		{
			defaultShape = 0;
		}
		if( shape.equals( "cone" ) )    // 1 1
		{
			defaultShape = 257;
		}
		if( shape.equals( "coneToMax" ) )    // 1 2
		{
			defaultShape = 513;
		}
		if( shape.equals( "cylinder" ) )    // 1 0
		{
			defaultShape = 1;
		}
		if( shape.equals( "pyramid" ) )    // 0 1
		{
			defaultShape = 256;
		}
		if( shape.equals( "pyramidToMax" ) ) // 0 2
		{
			defaultShape = 512;
		}
		return defaultShape;
	}

	/**
	 * convert the default shape flag to a user-friendly (OOXML-compliant) String
	 *
	 * @param shape int
	 * @return
	 */
	public String getShape()
	{
		switch( defaultShape )
		{
			case 0:
				return "box";
			case 1:
				return "cylinder";
			case 256:
				return "pyramid";
			case 257:
				return "cone";
			case 512:
				return "pyramidToMax";
			case 513:
				return "coneToMax";
		}
		return null;
	}

	/**
	 * returns an int representing the space between points in a 3d area, bar or line chart, or 0 if not 3d
	 *
	 * @return
	 */
	public int getGapDepth()
	{
		return cf.getGapDepth();
	}

	/**
	 * return the SeriesList record for this chart object
	 * The SeriesList record maps the series for the chart.
	 *
	 * @return
	 */
	protected SeriesList getSeriesList()
	{
		return (SeriesList) Chart.findRec( cf.chartArr, SeriesList.class );
	}

	/**
	 * @param fName
	 */
	public void WriteMainChartRecs( String fName )
	{
		class util
		{
			public void writeRecs( BiffRec b, BufferedWriter writer, int level ) throws IOException
			{
				String tabs = "\t\t\t\t\t\t\t\t\t\t";
				if( b == null )
				{
					return;
				}
				writer.write( tabs.substring( 0, level ) + b.getClass().toString().substring( b.getClass()
				                                                                               .toString()
				                                                                               .lastIndexOf( '.' ) + 1 ) );
				if( b instanceof com.extentech.formats.XLS.charts.SeriesText )
				{
					writer.write( "\t[" + b.toString() + "]" );
				}
				else if( b instanceof MSODrawing )
				{
					writer.write( "\t[" + b.toString() + "]" );
					//								writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
					writer.write( ((MSODrawing) b).debugOutput() );
					writer.write( "\t[" + ByteTools.getByteDump( b.getData(), 0 ).substring( 11 ) + "]" );
				}
				else if( b instanceof Obj )
				{
					writer.write( ((Obj) b).debugOutput() );
				}
				else if( b instanceof Label )
				{
					writer.write( "\t[" + b.getStringVal() + "]" );
				}
				else // all else, write bytes
				{
					writer.write( "\t[" + ByteTools.getByteDump( ByteTools.shortToLEBytes( b.getOpcode() ),
					                                             0 ) + "][" + ByteTools.getByteDump( b.getData(), 0 )
					                                                                   .substring( 11 ) + "]" );
				}
				writer.newLine();
				try
				{
					if( ((GenericChartObject) b).chartArr.size() > 0 )
					{
						ArrayList<com.extentech.formats.XLS.XLSRecord> chartArr = ((GenericChartObject) b).chartArr;
						for( XLSRecord aChartArr : chartArr )
						{
							writeRecs( aChartArr, writer, level + 1 );

						}
					}
				}
				catch( ClassCastException ce )
				{
					;
				}
			}
		}
		;

		try
		{
			java.io.File f = new java.io.File( fName );
			BufferedWriter writer = new BufferedWriter( new FileWriter( f ) );
			util u = new util();

			java.util.Vector v = this.getParentChart().getAllSeries();
			for( Object aV : v )
			{
				u.writeRecs( (BiffRec) aV, writer, 0 );
			}
			writer.newLine();
			ArrayList<com.extentech.formats.XLS.XLSRecord> chartArr = this.cf.chartArr;
			for( XLSRecord aChartArr : chartArr )
			{
				u.writeRecs( aChartArr, writer, 0 );
			}

			writer.flush();
			writer.close();
			writer = null;
		}
		catch( Exception e )
		{
		}
	}
}
