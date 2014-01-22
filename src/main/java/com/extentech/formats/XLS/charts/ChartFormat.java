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

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * <b>ChartFormat: Parent Record for Chart Group (0x1014)</b>
 * <p/>
 * 4 reserved 16 0 20 grbit 2 format flags 22 icrt 2 drawing order (0= bottom of
 * z-order)
 * <p/>
 * <p/>
 * 16 bytes- reserved must be 0
 * fVaried (1 bit): A bit that specifies whether the color for each data point and
 * the color and type for each data marker vary. If the chart group has multiple series or the chart group has one
 * series and the chart group type is a surface, stock, or area, this field MUST
 * be ignored, and the data points do not vary. For all other chart group types,
 * if the chart group has one series, a value of 0x1 specifies that the data
 * points vary.
 * 15 bits - reserved - 0
 * icrt (2 bytes): An unsigned integer that specifies the drawing order of the chart group relative to the other chart
 * groups, where 0x0000 is the bottom of the z-order.
 * This value MUST be unique for each instance of this record and MUST be less than or equal to 0x0009.
 * <p/>
 * <p/>
 * ORDER OF SUBRECS:
 * Bar/Pie/Scatter ...
 * ChartFormatLink
 * [SeriesList]
 * [ThreeD]
 * [Legend]
 * [DropBar]
 * [ChartLine, LineFormat]
 * [DataLabExt]
 * [DefaultText, Text]
 * [DataLabExtContents]
 * [DataFormat]
 * [ShapePropsStream]
 * <p/>
 * "http://www.extentech.com">Extentech Inc.</a>
 */
public class ChartFormat extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4000704166442059677L;
	private short grbit = 0;
	private boolean fVaried = false;
	private short drawingOrder = 0;

	@Override
	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( this.getByteAt( 16 ), this.getByteAt( 17 ) );
		drawingOrder = ByteTools.readShort( this.getByteAt( 18 ), this.getByteAt( 19 ) );
		fVaried = ((grbit & 0x1) == 0x1);
	}

	/**
	 *
	 *
	 */
	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( grbit );
		this.getData()[16] = b[0];
		this.getData()[17] = b[1];
	}

	/**
	 * specifies whether
	 * the color for each data point and the color and type for each data marker
	 * vary
	 *
	 * @param b
	 */
	public void setVaryColor( boolean vary )
	{
		fVaried = vary;
		grbit = ByteTools.updateGrBit( grbit, fVaried, 0 );
		updateRecord();
	}

	/**
	 * returns whether
	 * the color for each data point and the color and type for each data marker
	 * vary
	 *
	 * @param b
	 */
	public boolean getVaryColor()
	{
		return fVaried;
	}

	/**
	 * replace the existing chart object with the desired ChartObject,
	 * effectively changing the type of the chart
	 *
	 * @param co
	 */
	protected void setChartObject( ChartObject co )
	{
		chartArr.remove( 0 );
		chartArr.add( 0, (XLSRecord) co );

	}

	/**
	 * @return truth of "Chart is Three D"
	 */
	public boolean isThreeD( int chartType )
	{
		if( chartType != ChartConstants.BUBBLECHART )
		{
			for( int i = 0; i < chartArr.size(); i++ )
			{
				BiffRec b = (BiffRec) chartArr.get( i );
				if( b.getOpcode() == THREED )
				{
					return true;
				}
			}
		}
		else
		{
			DataFormat df = this.getDataFormatRec( false );
			if( df != null )
			{
				return df.getHas3DBubbles();
			}
		}
		return false;
	}

	/**
	 * return truth of "chart is 3d clustered"
	 *
	 * @return
	 */
	public boolean is3DClustered()
	{
		ThreeD td = (ThreeD) Chart.findRec( this.chartArr, ThreeD.class );
		if( td != null )
		{
			return td.isClustered();
		}
		return false;
	}

	/**
	 * sets if this chart has clustered bar/columns
	 *
	 * @param bIsClustered
	 */
	public void setIsClustered( boolean bIsClustered )
	{
		ThreeD td = getThreeDRec( false );
		if( td != null )
		{
			td.setIsClustered( bIsClustered );
		}
		else
		{
			if( chartArr.get( 0 ).getOpcode() == BAR )
			{
				((Bar) chartArr.get( 0 )).setIsClustered();
			}
		}

	}

	/**
	 * sets the Space between points (50 or 150 is default)
	 *
	 * @param gap
	 */
	public void setGapDepth( int gap )
	{
		ThreeD td = getThreeDRec( true );
		td.setPcGap( gap );
	}

	/**
	 * returns an int representing the space between points in a 3d area, bar or line chart, or 0 if not 3d
	 *
	 * @return
	 */
	public int getGapDepth()
	{
		ThreeD td = getThreeDRec( false );
		if( td != null )
		{
			td.getPcGap();
		}
		return 0;
	}

	/**
	 * return ThreeD options in XML form
	 *
	 * @return String XML
	 */
	public String getThreeDXML()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b.getOpcode() == THREED )
			{
				return ((ThreeD) b).getOptionsXML();
			}
		}
		return "";
	}

	/**
	 * percentage=distance of pie slice from center of pie as %
	 *
	 * @param p
	 */
	public void setPercentage( int p )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setPercentage( p );
	}

	/**
	 * bar shape:
	 * <br>the shape is as follows:
	 * public static final int SHAPECOLUMN= 0;		// default
	 * public static final int SHAPECYLINDER= 1;
	 * public static final int SHAPEPYRAMID= 256;
	 * public static final int SHAPECONE= 257;
	 * public static final int SHAPEPYRAMIDTOMAX= 516;
	 * public static final int SHAPECONETOMAX= 517;
	 *
	 * @param shape
	 */
	public void setBarShape( int shape )
	{
		DataFormat df = this.getDataFormatRec( true );
		// THIS DOES NOT MAKE SENSE ACCORDING TO DOC BUT IS WHAT EXCEL DOES
		df.setPointNumber( 0 );
		df.setSeriesIndex( 0 );
		df.setSeriesNumber( -3 );
		df.setShape( shape );
	}

	/**
	 * sets chart options such as threed options, grid lines, etc ... these options
	 * are distinct from chart-type-specific options, which are handled by the
	 * appropriate chart record (pie, bar ...)
	 *
	 * @param op  String option name
	 * @param val Object value
	 */
	public void setOption( String op, String val )
	{
		if( op.equalsIgnoreCase( "Percentage" ) )
		{
			setPercentage( Short.valueOf( val ).shortValue() );
		}
		else if( op.equalsIgnoreCase( "Shape" ) )
		{
			setBarShape( Integer.parseInt( val ) );
		}
		else if( op.equals( "ShowBubbleSizes" ) || // TextDisp options
				op.equals( "ShowLabelPct" ) || op.equals( "ShowCatLabel" ) || op.equals( "ShowPct" ) || op.equals( "Rotation" ) ||
				// op.equals("ShowValue") || unknown
				op.equals( "Label" ) || op.equals( "TextRotation" ) )
		{
			TextDisp td = getDataLegendTextDisp( 0 );
			td.setChartOption( op, val );
		}
		else if( op.equals( "Perspective" ) || // ThreeD options
				op.equals( "Cluster" ) || op.equals( "ThreeDScaling" ) || op.equals( "TwoDWalls" ) || op.equals( "PcGap" ) || op.equals(
				"PcDepth" ) || op.equals( "PcHeight" ) || op.equals( "PcDist" ) || op.equals( "AnElev" ) || op.equals( "AnRot" ) )
		{
			ThreeD td = this.getThreeDRec( true );
			td.setChartOption( op, val );
		}
		else if( op.equals( "ShowValueLabel" ) || // Attached Label Options
				op.equals( "ShowValueAsPercent" ) || op.equals( "ShowLabelAsPercent" ) || op.equals( "ShowLabel" ) || op.equals(
				"ShowSeriesName" ) || op.equals( "ShowBubbleLabel" ) )
		{
			DataFormat df = this.getDataFormatRec( true );
			df.setDataLabels( op );
		}
		else if( op.equalsIgnoreCase( "SmoothedLine" ) || op.equalsIgnoreCase( "ThreeDBubbles" ) || op.equalsIgnoreCase( "ArShadow" ) )
		{
			DataFormat df = this.getDataFormatRec( true );
			if( op.equalsIgnoreCase( "SmoothedLine" ) )
			{
				df.setSmoothLines( true );
			}
			else if( op.equalsIgnoreCase( "ThreeDBubbles" ) )
			{
				df.setHas3DBubbles( true );
			}
			else
			{
				df.setHasShadow( true );
			}
		}
	}

	/**
	 * Return the ThreeD rec associated with this ChartFormat, create if not
	 * present
	 */
	// NOTES: 
	// The Chart3d record specifies that the plot area of the chart group is rendered in a 3-D scene 
	// and also specifies the attributes of the 3-D plot area. 
	// The preceding chart group type MUST be of type bar, pie, line, area, or surface.
	public ThreeD getThreeDRec( boolean bCreate )
	{
		ThreeD td = (ThreeD) Chart.findRec( this.chartArr, ThreeD.class );
		if( td == null && bCreate )
		{ // add ThreeD rec
			for( int i = 0; i < this.chartArr.size(); i++ )
			{
				BiffRec b = (BiffRec) this.chartArr.get( i );
				if( b.getOpcode() == CHARTFORMATLINK )
				{
					if( (i + 1) < chartArr.size() && ((BiffRec) this.chartArr.get( i + 1 )).getOpcode() == SERIESLIST )
					{
						i++;    // rare that SeriesList record appears
					}
					td = (ThreeD) ThreeD.getPrototype();
					td.setParentChart( this.getParentChart() );
					this.chartArr.add( i + 1, td );
					break;
				}
			}
		}
		return td;
	}

	/**
	 * Add or Retrieve TextDisp and assoc records specific for Data Legends
	 *
	 * @return
	 */
	private TextDisp getDataLegendTextDisp( int type )
	{
		int i = Chart.findRecPosition( this.chartArr, Legend.class );
		TextDisp td = null;
		if( this.chartArr.size() <= (i + 1) || this.chartArr.get( i + 1 ).getClass() != DefaultText.class )
		{ // then add one
			DefaultText d = (DefaultText) DefaultText.getPrototype();
			d.setType( (short) type );
			d.setParentChart( this.getParentChart() );
			this.chartArr.add( ++i, d );
			td = (TextDisp) TextDisp.getPrototype( ObjectLink.TYPE_DATAPOINTS, "", this.getWorkBook() );
			td.setParentChart( this.getParentChart() );
			this.chartArr.add( ++i, td );
		}
		else
		{
			DefaultText d = (DefaultText) this.chartArr.get( i + 1 );
			if( d.getType() != type )
			{ // / add a new one
				i += 2; // add after TextDisp
				d = (DefaultText) DefaultText.getPrototype();
				d.setType( (short) type );
				d.setParentChart( this.getParentChart() );
				this.chartArr.add( ++i, d );
				td = (TextDisp) TextDisp.getPrototype( ObjectLink.TYPE_DATAPOINTS, "", this.getWorkBook() );
				td.setParentChart( this.getParentChart() );
				this.chartArr.add( ++i, td );
			}
			else
			{ // it's the correct one
				i += 2;
				td = (TextDisp) this.chartArr.get( i );
			}
		}
		return td;
	}

	/**
	 * Gets the dataformat record associated with this ChartFormat If none
	 * present, creates a basic DataFormat set of records DataFormat controls
	 * Data Labels, % Distance from sections, line formats ...
	 *
	 * @return DataFormat Record
	 */
	private DataFormat getDataFormatRec( boolean bCreate )
	{
		DataFormat df = (DataFormat) Chart.findRec( this.chartArr, DataFormat.class );
		if( df == null && bCreate )
		{ // create dataformat
			df = (DataFormat) DataFormat.getPrototypeWithFormatRecs( this.getParentChart() );
			this.addChartRecord( df );
		}
		return df;
	}

	/**
	 * return the Data Labels chosen for this chart, if any can be one or more
	 * of: <br>
	 * Value <br>
	 * ValuePerecentage <br>
	 * CategoryPercentage <br>
	 * CategoryLabel <br>
	 * BubbleLabel <br>
	 * SeriesLabel or an empty string if no data labels are chosen for the chart
	 *
	 * @return
	 */
	public String getDataLabels()
	{
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			return df.getDataLabelType();
		}
		return "";
	}

	/**
	 * returns true if this chart displays lines (Line, Scatter, Radar)
	 *
	 * @return true if chart has lines (see Scatter, Line chart ...)
	 */
	public boolean getHasLines()
	{
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			return df.getHasLines();
		}
		return false;
	}

	/**
	 * sets this chart to have default lines (Scatter, Line chart ...)
	 */
	public void setHasLines()
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setHasLines();
	}

	/**
	 * sets this chart to have lines (Scatter, Line chart ...) of style lineStyle
	 * <br>Style of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
	 */
	public void setHasLines( int lineStyle )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setHasLines( lineStyle );
	}

	/**
	 * returns true if this chart has smoothed lines (Scatter, Line, Radar)
	 *
	 * @return
	 */
	public boolean getHasSmoothLines()
	{
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			return df.getSmoothedLines();
		}
		return false;
	}

	/**
	 * sets this chart to have smoothed lines (Scatter, Line, Radar)
	 */
	public void setHasSmoothLines( boolean b )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setSmoothLines( b );
	}

	/**
	 * returns true if chart has Drop Lines
	 *
	 * @return
	 */
	public boolean getHasDropLines()
	{
		/* chartline:  
		line, chartformatlink, <serieslist>, <3d>, <legend>, chartline, lineformat, startblock, 
		shapepropsstream, [-92, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 
		                   0, 0, 
		                   -51, -110, 
		                   -66, 74, 1, -88, 0, 0, 0, 0]
		endblock
		
	dropbar= chartline, lineformat	
	*/
		return false;
	}

	/**
	 * sets this chart (Area, Line or Stock) to have Drop Lines
	 */
	public void setHasDropLines()
	{

	}

	/**
	 * sets 3d bubble state
	 *
	 * @param has3dBubbles
	 */
	public void setHas3DBubbles( boolean has3dBubbles )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setHas3DBubbles( has3dBubbles );

	}

	/**
	 * return data label options as an int <br>
	 * can be one or more of: <br>
	 * SHOWVALUE= 0x1; <br>
	 * SHOWVALUEPERCENT= 0x2; <br>
	 * SHOWCATEGORYPERCENT= 0x4; <br>
	 * SHOWCATEGORYLABEL= 0x10; <br>
	 * SHOWBUBBLELABEL= 0x20; <br>
	 * SHOWSERIESLABEL= 0x40;
	 *
	 * @return a combination of data label options above or 0 if none
	 * <p/>
	 * NOTE: this returns the Data Labels settings for the entire chart,
	 * not a particular series
	 * @see AttachedLabel
	 */
	public int getDataLabelsInt()
	{
		int datalabels = 0;
		int z = Chart.findRecPosition( this.chartArr, DataLabExtContents.class );

		// here we are assuming that the TextDisp is of the proper ObjectLink=4 type ... 
		if( z > 0 && this.chartArr.get( z - 1 ) instanceof TextDisp )
		{ // Extended Label -- add to attachedlabel, if any
			DataLabExtContents dl = (DataLabExtContents) chartArr.get( z );
			datalabels = dl.getTypeInt(); // if so, no fontx record ... use default???
		}
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			datalabels |= df.getDataLabelTypeInt();
		}
		return datalabels;
	}

	/**
	 * sets the data labels for the entire chart (as opposed to a specific series/data point).
	 * A combination of:
	 * <li>SHOWVALUE= 0x1;
	 * <li>SHOWVALUEPERCENT= 0x2;
	 * <li>SHOWCATEGORYPERCENT= 0x4;
	 * <li>SHOWCATEGORYLABEL= 0x10;
	 * <li>SHOWBUBBLELABEL= 0x20;
	 * <li>SHOWSERIESLABEL= 0x40;
	 */
	public void setHasDataLabels( int dl )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setHasDataLabels( dl );
	}

	/**
	 * returns the bar shape for a column or bar type chart can be one of: <br>
	 * ChartConstants.SHAPECOLUMN default <br>
	 * ChartConstants.SHAPECONE <br>
	 * ChartConstants.SHAPECONETOMAX <br>
	 * ChartConstants.SHAPECYLINDER <br>
	 * ChartConstants.SHAPEPYRAMID <br>
	 * ChartConstants.SHAPEPYRAMIDTOMAX
	 *
	 * @return
	 */
	public int getBarShape()
	{
		int shape = SHAPEDEFAULT;
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			shape = df.getShape();
		}
		return shape;
	}

	/**
	 * returns type of marker, if any <br>
	 * 0 = no marker <br>
	 * 1 = square <br>
	 * 2 = diamond <br>
	 * 3 = triangle <br>
	 * 4 = X <br>
	 * 5 = star <br>
	 * 6 = Dow-Jones <br>
	 * 7 = standard deviation <br>
	 * 8 = circle <br>
	 * 9 = plus sign
	 *
	 * @return
	 */
	public int getMarkerFormat()
	{
		int markertype = 8; // default= circles
		// default actually looks like: 2, 1, 5, 4 ...
		DataFormat df = this.getDataFormatRec( false );
		if( df != null )
		{
			return df.getMarkerFormat();
		}
		return 0;
	}

	/**
	 * 0 = no marker <br>
	 * 1 = square <br>
	 * 2 = diamond <br>
	 * 3 = triangle <br>
	 * 4 = X <br>
	 * 5 = star <br>
	 * 6 = Dow-Jones <br>
	 * 7 = standard deviation <br>
	 * 8 = circle <br>
	 * 9 = plus sign
	 */
	public void setMarkers( int markerFormat )
	{
		DataFormat df = this.getDataFormatRec( true );
		df.setMarkerFormat( markerFormat );
	}

	/**
	 * Chart Line options (not available on all charts):
	 * <br>Drop Lines available on Line, Area and Stock charts
	 * <br>HiLow Lines are available on Line and Stock charts
	 * <br>Series Lines are available on Bar and OfPie charts
	 * <pre>
	 * 0= drop lines below the data points of Line, Area and Stock charts
	 * 1= High-low lines around the data points of Line and Stock charts
	 * 2- Series Line connecting data points of stacked column and bar charts and OfPie Charts
	 * </pre>
	 * <br>
	 *
	 * @param lineType
	 */
	public ChartLine addChartLines( int lineType )
	{
//		ChartLine cl= (ChartLine) Chart.findRec(chartArr, ChartLine.class);
//		if (cl==null) {
		ChartLine cl = null;
		// goes After DropBar or legend or 3d (or CFL)
		for( int i = chartArr.size() - 1; i >= 1; i-- )
		{
			BiffRec br = (BiffRec) chartArr.get( i );
			short op = br.getOpcode();
			if( op == DROPBAR || op == LEGEND || op == THREED || op == CHARTFORMATLINK )
			{
				cl = (ChartLine) cl.getPrototype();
				cl.setParentChart( this.getParentChart() );
				chartArr.add( ++i, cl );
				cl.setParentChart( this.getParentChart() );    // ensure can find
				LineFormat lf = (LineFormat) LineFormat.getPrototype( 0, 1 );
				lf.setParentChart( this.getParentChart() );
				chartArr.add( ++i, lf );
				break;
			}
		}
//		}
		cl.setLineType( lineType );
		return cl;
	}

	/**
	 * return ChartLines option, if any
	 * <pre>
	 * 0= drop lines below the data points of Line, Area and Stock charts
	 * 1= High-low lines around the data points of Line and Stock charts
	 * 2- Series Line connecting data points of stacked column and bar charts and OfPie Charts
	 * </pre>
	 *
	 * @return 0-2 or -1 if no chart lines
	 */
	public int getChartLines()
	{
		ChartLine cl = (ChartLine) Chart.findRec( chartArr, ChartLine.class );
		if( cl != null )
		{
			return cl.getLineType();
		}
		return -1;
	}

	/**
	 * return the record governing chart lines: dropLines, Hi-low lines or Series Lines
	 *
	 * @return
	 */
	public ChartLine getChartLinesRec()
	{
		return (ChartLine) Chart.findRec( chartArr, ChartLine.class );
	}

	public ChartLine getChartLinesRec( int id )
	{
		int i = Chart.findRecPosition( chartArr, ChartLine.class );
		if( i > -1 )
		{
			while( i < chartArr.size() )
			{
				if( ((BiffRec) chartArr.get( i )).getOpcode() == CHARTLINE )
				{
					ChartLine cl = (ChartLine) chartArr.get( i );
					if( cl.getLineType() == id )
					{
						return cl;
					}
				}
				else
				{
					break;
				}
				i += 2;
			}
		}
		return null;
	}

	/**
	 * add up/down bars (line, area stock)
	 */
	public void addUpDownBars()
	{
		if( Chart.findRec( chartArr, Dropbar.class ) == null )
		{
			// create necessary records to describe up/down bars  
			Dropbar upBar = (Dropbar) Dropbar.getPrototype();
			upBar.setParentChart( this.getParentChart() );
			Dropbar downBar = (Dropbar) Dropbar.getPrototype();
			downBar.setParentChart( this.getParentChart() );
			LineFormat lf = (LineFormat) LineFormat.getPrototype();
			lf.setParentChart( this.getParentChart() );
			AreaFormat af = (AreaFormat) AreaFormat.getPrototype();
			af.setParentChart( this.getParentChart() );
			upBar.chartArr.add( lf );
			upBar.chartArr.add( af );
			lf = (LineFormat) LineFormat.getPrototype();
			lf.setParentChart( this.getParentChart() );
			af = (AreaFormat) AreaFormat.getPrototype();
			af.setParentChart( this.getParentChart() );
			downBar.chartArr.add( lf );
			downBar.chartArr.add( af );

			// add dropbar records to subarray
			for( int i = chartArr.size() - 1; i >= 0; i-- )
			{
				BiffRec br = (BiffRec) chartArr.get( i );
				short op = br.getOpcode();
				if( op == SERIESLIST || op == LEGEND || op == THREED || op == CHARTFORMATLINK )
				{
					chartArr.add( ++i, upBar );
					chartArr.add( ++i, downBar );
					break;
				}
			}
		}
	}

	/**
	 * parse upDownBars OOXML element (controled by 2 DropBar records in this subArray)
	 * <br>Valid for Line and Stock charts only
	 */
	public void parseUpDownBarsOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		// assume dropbar records are not present ...
		addUpDownBars();
		int z = Chart.findRecPosition( chartArr, Dropbar.class );
		Dropbar downBar = (Dropbar) chartArr.get( z++ );
		Dropbar upBar = (Dropbar) chartArr.get( z );

		try
		{
			Dropbar curbar = null;
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "downBars" ) )
					{
						curbar = downBar;
					}
					else if( tnm.equals( "upBars" ) )
					{
						curbar = upBar;
					}
					else if( tnm.equals( "gapWidth" ) )
					{    // default=150
						upBar.setGapWidth( Integer.valueOf( xpp.getAttributeValue( 0 ) ) );    // TODO: should this be in 1st dropbar or both??
						downBar.setGapWidth( Integer.valueOf( xpp.getAttributeValue( 0 ) ) );    // TODO: should this be in 1st dropbar or both??
					}
					else if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						SpPr sppr = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk ).cloneElement();
						LineFormat lf = (LineFormat) curbar.chartArr.get( 0 );
						if( lf != null )
						{
							lf.setFromOOXML( sppr );
						}
						// TODO: fill AreaFormat with sppr
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( "upDownBars" ) )
					{
						lastTag.pop();
						break;
					}
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "parseUpDownBarsOOXML: " + e.toString() );
		}
	}

	/**
	 * return the OOXML to define the upDownBars element
	 * <br>defined by 2 Dropbar records in this subarray
	 * <br>Only valid for Line and Stock charts
	 *
	 * @return
	 */
	public String getUpDownBarOOXML()
	{
		int z = Chart.findRecPosition( chartArr, Dropbar.class );    //
		if( z == -1 )
		{
			return "";
		}
		StringBuffer ooxml = new StringBuffer();
		try
		{
			Dropbar upBars = (Dropbar) chartArr.get( z++ );
			Dropbar downBars = (Dropbar) chartArr.get( z );
			z++;
			ooxml.append( "<c:upDownBars>" );
			// c:gapWidth
			int gw = upBars.getGapWidth();    // TODO: only upBar?  Should they match?
			if( gw != 150 )// default
			{
				ooxml.append( "<c:gapWidth val=\"" + gw + "\"/>" );
			}
			// c:upBars
			ooxml.append( upBars.getOOXML( true ) );
			// c:downBars
			ooxml.append( downBars.getOOXML( false ) );
			ooxml.append( "</c:upDownBars>" );
		}
		catch( Exception e )
		{
		}
		return ooxml.toString();
	}

	/**
	 * return the drawing order of this ChartFormat <br>
	 * For multiple charts-in-one, drawing order determines the order of the
	 * charts
	 *
	 * @return
	 */
	public int getDrawingOrder()
	{
		return drawingOrder;
	}

	/**
	 * set the drawing order of this ChartFormat <br>
	 * For multiple charts-in-one, drawing order determines the order of the
	 * charts
	 *
	 * @param order
	 */
	public void setDrawingOrder( int order )
	{
		drawingOrder = (short) order;
		byte[] b = ByteTools.shortToLEBytes( drawingOrder );
		this.getData()[18] = b[0];
		this.getData()[19] = b[1];
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
	};

	public static XLSRecord getPrototype()
	{
		ChartFormat cf = new ChartFormat();
		cf.setOpcode( CHARTFORMAT );
		cf.setData( cf.PROTOTYPE_BYTES );
		cf.init();
		return cf;
	}

	/**
	 * access chart-type record and return any options specific for this chart
	 * in XML form Gathers chart options such as show legend, grid lines, etc
	 * ... these options are distinct from chart-type-specific options, which
	 * are handled by the appropriate chart record (pie, bar ...) Also, since
	 * both ThreeD options and Axis-specific options are quite extensive, these
	 * are handled separately
	 *
	 * @return String of options XML ("" if no options set)
	 * @see setChartOption
	 * @see setOption
	 * @see getThreeDXML()
	 * @see getAxesXML()
	 */
	public String getChartOptionsXML()
	{
		boolean bFoundDefaultText0 = false;
		StringBuffer sb = new StringBuffer();
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			if( b instanceof ChartObject )
			{
				ChartObject co = (ChartObject) b;
				if( b instanceof DataFormat )
				{
					// controls Data Legends, % Distance of sections, Line
					// Format, Area Format, Bar Shapes ...
					DataFormat df = (DataFormat) b;
					int shape = df.getShape();
					if( shape != 0 )
					{
						sb.append( " Shape=\"" + shape + "\"" );
					}
					for( int z = 0; z < df.chartArr.size(); z++ )
					{
						b = (BiffRec) df.chartArr.get( z );
						if( b instanceof PieFormat )
						{
							sb.append( ((PieFormat) b).getOptionsXML() );
						}
						else if( b instanceof AttachedLabel )
						{
							String type = ((AttachedLabel) b).getType();
							sb.append( " DataLabel=\"" + type + "\"" );
						}
						else if( b instanceof Serfmt )
						{
							sb.append( ((Serfmt) b).getOptionsXML() );
						}
						else if( b instanceof MarkerFormat )
						{
							sb.append( ((MarkerFormat) b).getOptionsXML() );
						}
					}
				}
				else if( b instanceof DefaultText )
				{ // controls Show Legend + // some Data Legends
					if( !bFoundDefaultText0 )
					{
						bFoundDefaultText0 = (((DefaultText) b).getType() == 0);
					}
					if( ((DefaultText) b).getType() == 1 && bFoundDefaultText0 )
					{
						sb.append( " ShowLegendKey=\"true\"" );
					}
				}
			}
		}
		return sb.toString();
	}

	/**
	 * get the value of *almost* any chart option (axis options are in Axis)
	 *
	 * @param op String option e.g. Shadow or Percentage
	 * @return String value of option
	 */
	@Override
	public String getChartOption( String op )
	{
		DataFormat df = this.getDataFormatRec( false );
		try
		{
			if( op.equals( "Percentage" ) )
			{ // pieformat
				return String.valueOf( df.getPercentage() );
			}
			else if( op.equals( "ShowValueLabel" ) || // Attached Label Options
					op.equals( "ShowValueAsPercent" ) || op.equals( "ShowLabelAsPercent" ) || op.equals( "ShowLabel" ) || op.equals(
					"ShowBubbleLabel" ) )
			{
				return df.getDataLabelType( op );
			}
			else if( op.equals( "ShowBubbleSizes" ) || // TextDisp options
					op.equals( "ShowLabelPct" ) || op.equals( "ShowPct" ) || op.equals( "ShowCatLabel" ) ||
					// op.equals("ShowValue") || unknown
					op.equals( "Rotation" ) || op.equals( "Label" ) || op.equals( "TextRotation" ) )
			{
				TextDisp td = getDataLegendTextDisp( 0 );
				return td.getChartOption( op );
			}
			else if( op.equals( "Perspective" ) || // ThreeD options
					op.equals( "Cluster" ) || op.equals( "ThreeDScaling" ) || op.equals( "TwoDWalls" ) || op.equals( "PcGap" ) || op.equals(
					"PcDepth" ) || op.equals( "PcHeight" ) || op.equals( "PcDist" ) || op.equals( "AnElev" ) || op.equals( "AnRot" ) )
			{
				ThreeD td = this.getThreeDRec( false );
				if( td != null )
				{
					return td.getChartOption( op );
				}
				return "";
			}
			else if( op.equals( "ThreeDBubbles" ) )
			{
				return String.valueOf( df.getHas3DBubbles() );
			}
			else if( op.equals( "ArShadow" ) )
			{
				return String.valueOf( df.getHasShadow() );
			}
			else if( op.equals( "SmoothLines" ) )
			{
				return String.valueOf( df.getSmoothedLines() );
				// TODO: FINSIH REST!
			}
			else if( op.equals( "AxisLabels" ) )
			{ // Radar, RadarArea
			}
			else if( op.equals( "BubbleSizeRatio" ) )
			{ // Scatter
			}
			else if( op.equals( "BubbleSize" ) )
			{ // Scatter
			}
			else if( op.equals( "ShowNeg" ) )
			{ // Scatter
			}
			else if( op.equals( "ColorFill" ) )
			{ // Surface
			}
			else if( op.equals( "Shading" ) )
			{ // Surface
			}
			else if( op.equals( "MarkerFormat" ) )
			{ // MarkerFormat
			}
		}
		catch( NullPointerException e )
		{

		}
		return "";
	}

}
