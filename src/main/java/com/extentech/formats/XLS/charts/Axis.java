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
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.NumFmt;
import com.extentech.formats.OOXML.OOXMLElement;
import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.OOXML.Title;
import com.extentech.formats.OOXML.TxPr;
import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.Font;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.formats.cellformat.CellFormatFactory;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;
import com.extentech.toolkit.StringTool;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

//OOXML-specific structures
//OOXML-specific structures

/**
 * <b>Axis: Axis Type (0x101d)</b>
 * <p/>
 * 4	wType		2	axis type (0= category axis or x axis on a scatter chart, 1= value axis, 2= series axis
 * 6	(reserved)	16	0
 * <p/>
 * Order of Axis Subrecords:
 * X (Cat Axis)
 * <p/>
 * CatSerRange
 * AxcExt
 * CatLab
 * [IfmtRecord]
 * [Tick]
 * [FontX]
 * [AxisLineFormat, LineFormat] - if gridlines (major or minor) or border around wall/floor
 * [AreaFormat, GelFrame]
 * [Shape or TextPropsStream]
 * [CrtMltFrt]
 * <p/>
 * Y (Value Axis)
 * ValueRange
 * [YMult -->Text-->Pos...] -- display units
 * [IfmtRecord]
 * [Tick]
 * [FontX]
 * [AxisLineFormat, LineFormat] - if gridlines (major or minor) or border around wall/floor
 * [AreaFormat, GelFrame]
 * [Shape or TextPropsStream]
 * [CrtMltFrt]
 * <p/>
 * XY (Series Axis)
 * CatSerRange
 * [IfmtRecord]
 * [Tick]
 * [FontX]
 * [AxisLineFormat, LineFormat] - if gridlines (major or minor) or border around wall/floor
 * [AreaFormat, GelFrame]
 * [Shape or TextPropsStream]
 * [CrtMltFrt]
 */
/* General info on Excel axis types:
 *  In Microsoft Excel charts, there are different types of X axes. 
 *  While the Y axis is a Value type axis (i.e. containing a ValueRange record), 
 *  the X axis can be a Category type axis or a Value type axis. 
 *  Using a Value axis, the data is treated as continuously varying numerical data, 
 *  and the marker is placed at a point along the axis which varies according to its 
 *  numerical value. 
 *  Using a Category axis, the data is treated as a sequence of non-numerical text labels, 
 *  and the marker is placed at a point along the axis according to its position in 
 *  the sequence.     	    	
 *  
 *  Note that Scatter (x/y) charts and Bubble charts have two Value axes, pie and donut charts have no axes
 *  
 *  How do you arrange your chart so the categories are displayed along the Y axis? 
 *  The method involves adding a dummy series along the Y axis, 
 *  applying data labels to its points for category labels, 
 *  and making the original Y axis disappear.
 */
public class Axis extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -8592219101790307789L;
	short wType = 0;
	private TextDisp linkedtd = null;    // 20070730 KSC: hold axis legend TextDisp for this axis
	private AxisParent ap = null;        // 20090108 KSC: links to this axes parent --
	// Axis placement
	public static final int INVISIBLE = 0;
	public static final int LOW = 1;
	public static final int HIGH = 2;
	public static final int NEXTTO = 3;
	// OOXML-specific
	private SpPr shapeProps = null;    // 20081224 KSC:  OOXML-specific holds the shape properties (line and fill) for this axis
	private Title ttl = null;        // OOXML title element
	private TxPr txpr = null;        // text properties for axis
	private NumFmt nf = null;        // NumFmt prop for axis
	String axPos = null;

	@Override
	public void init()
	{
		super.init();
		wType = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

	public short getAxis()
	{
		return wType;
	}

	public void setAxis( int wType )
	{
		this.wType = (short) wType;
		if( wType == XVALAXIS )
		{
			wType = XAXIS;        // 20090108 KSC: XVALAXIS is type of X axis with VAL records
		}
		byte[] b = ByteTools.shortToLEBytes( (short) wType );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
	}

	public static XLSRecord getPrototype( int wType )
	{
		Axis a = new Axis();
		a.setOpcode( AXIS );
		a.setData( a.PROTOTYPE_BYTES );
//        if (wType!=XVALAXIS)
		a.setAxis( wType );
//        else
//        	a.setAxis(XAXIS);
		// also add the associated records
		switch( wType )
		{
			case XAXIS:
				a.addChartRecord( (CatserRange) CatserRange.getPrototype() );
				a.addChartRecord( (Axcent) Axcent.getPrototype() );
				a.addChartRecord( (Tick) Tick.getPrototype() );
				break;
			case YAXIS:
			case XVALAXIS:
				a.addChartRecord( (ValueRange) ValueRange.getPrototype() );
				a.addChartRecord( (Tick) Tick.getPrototype() );
				AxisLineFormat alf = (AxisLineFormat) AxisLineFormat.getPrototype();
				alf.setId( AxisLineFormat.ID_MAJOR_GRID );    // default has major gridlines
				a.addChartRecord( alf );
				a.addChartRecord( (LineFormat) LineFormat.getPrototype() );
				break;
			case ZAXIS:
				// KSC: TODO: Set CatserRange options correctly when get def!!! **********
				CatserRange c = (CatserRange) CatserRange.getPrototype();
				a.addChartRecord( c );
				a.addChartRecord( (Tick) Tick.getPrototype() );    // TODO: Tick should have 
				break;
		}
		return a;
	}

	/**
	 * @return String XML representation of this chart-type's options
	 */
	@Override
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		boolean hasMajorGridLines = false;
		// Axis 0: CatserRange, Axcent, Tick [AxisLineFormat, LineFormat, AreaFormat] 					last 3 recs are for 3d formatting	
		// Axis 1: ValueRange, Tick, AxisLineFormat, LineFormat [AreaFormat, LineFormat, AreaFormat]		"	"
		// Axis 2: [CatserRange, Tick]		Z axis, for surface charts (only??)
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = (BiffRec) chartArr.get( i );
			// Handle subordinate record options here rather than in the specific rec
			if( b instanceof CatserRange )
			{
				CatserRange c = ((CatserRange) b);
				// record deviations from defaults
				if( c.getCatCross() != 1 )
				{
					sb.append( " CatCross=\"" + c.getCatCross() + "\"" );
				}
				if( c.getCatLabel() != 1 )
				{
					sb.append( " LabelCross=\"" + c.getCatLabel() + "\"" );
				}
				if( c.getCatMark() != 1 )
				{
					sb.append( " Marks=\"" + c.getCatMark() + "\"" );
				}
				if( !c.getCrossBetween() )
				{
					sb.append( " CrossBetween=\"false\"" );
				}
				if( c.getCrossMax() )
				{
					sb.append( " CrossMax=\"true\"" );
				}
			}
			else if( b instanceof AxisLineFormat )
			{    // necessary for 3d charts and bg colors ...
				int id = ((AxisLineFormat) b).getId();
				if( id == AxisLineFormat.ID_MAJOR_GRID ) // Y Axis draw gridline usually true, trap if do not have (see below)          		
				{
					hasMajorGridLines = true;
				}
				else if( id == AxisLineFormat.ID_WALLORFLOOR ) // indicates area
				{
					sb.append( " AddArea=\"true\"" );
				}
			}
			else if( b instanceof AreaFormat )
			{ // should ONLY be present if AxisLineFormat AddArea
				// if not default background colors, record and reset
				int icvFore = ((AreaFormat) b).geticvFore();
				int icvBack = ((AreaFormat) b).geticvBack();
				if( icvBack == -1 )
				{
					icvBack = 0x4D;
				}
				if( icvFore == 1 )
				{
					icvFore = 0x4E;
				}
				// if not defaults, set fore and back of walls, sides or floor
				if( (wType == 0 && icvFore != 22) // xaxis
						|| (wType == 1 && icvFore != -1) ) // yaxis
				{
					sb.append( " AreaFg=\"" + icvFore + "\"" );
				}
				if( (wType == 0 && icvBack != 0) // xaxis
						|| (wType == 1 && icvBack != 1) ) // yaxis
				{
					sb.append( " AreaBg=\"" + icvBack + "\"" );
				}
			}
			// TODO: parse LineFormat and, if not standard, add as option
			// TODO: Parse Tick for options            
		}
		if( wType == 1 && !hasMajorGridLines ) // most Y Axes have major grid lines; flag if do not
		{
			sb.append( " MajorGridLines=\"false\"" );
		}
		return sb.toString();
	}

	/**
	 * get/set linked TextDisp (legend for this axis)
	 * used for setting axis options
	 *
	 * @param td
	 */
	public void setTd( TextDisp td )
	{
		linkedtd = td;
	}

	public TextDisp getTd()
	{
		if( linkedtd == null )
		{
			getTitleTD( false );
		}
		return linkedtd;
	}

	/**
	 * remove the title for this axis
	 */
	public void removeTitle()
	{
		if( linkedtd != null )
		{
			int x = Chart.findRecPosition( ap.chartArr, TextDisp.class );
			while( x > -1 )
			{
				if( ((TextDisp) ap.chartArr.get( x )).getType() == TextDisp.convertType( wType ) )
				{
					ap.chartArr.remove( x );
					break;
				}
				x++;
				if( ((BiffRec) ap.chartArr.get( x )).getOpcode() != TEXTDISP )
				{
					break;
				}
			}
		}
		linkedtd = null;
	}

	/**
	 * returns true if plot area has a bonding box/border
	 *
	 * @return
	 */
	public boolean hasPlotAreaBorder()
	{
		Frame f = (Frame) Chart.findRec( ap.chartArr, Frame.class );
		if( f != null )
		{
			return f.hasBox();
		}
		return false;
	}

	/**
	 * link this axis with it's parent
	 *
	 * @param a
	 */
	public void setAP( AxisParent a )
	{
		ap = a;
	}

	/**
	 * return the Title associated with this Axis (through the linked TextDisp record)
	 *
	 * @return
	 * @deprecated use getTitle
	 */
	public String getLabel()
	{
		return getTitle();

	}

	public String getTitle()
	{
		if( linkedtd == null )
		{
			getTitleTD( false );
		}
		if( linkedtd != null )
		{
			return linkedtd.toString();
		}
		return "";
	}

	/**
	 * return the Font object associated with the Axis Title, or null if  none
	 */
	public Font getTitleFont()
	{
		try
		{
			if( linkedtd != null )
			{
				Fontx fx = (Fontx) Chart.findRec( linkedtd.chartArr, Fontx.class );
				return this.getParentChart().getWorkBook().getFont( fx.getIfnt() );
			}
		}
		catch( Exception e )
		{
		}
		return null;
	}

	/**
	 * finds or adds TextDisp records for this axis (for labels and fonts)
	 */
	private void getTitleTD( boolean add )
	{
		int tdtype = TextDisp.convertType( wType );
		int pos = 0;
		// Must add if can't find existing label
		int x = Chart.findRecPosition( ap.chartArr, TextDisp.class );
		pos = x;
		while( x > 0 && linkedtd == null )
		{
			TextDisp td = (TextDisp) ap.chartArr.get( x );
			if( td.getType() == tdtype )    // found it
			{
				linkedtd = td;
			}
			else
			{
				if( !(ap.chartArr.get( ++x ) instanceof TextDisp) )
				{
					x = -1;
				}
			}
		}
		if( linkedtd == null && add )
		{
			linkedtd = (TextDisp) TextDisp.getPrototype( tdtype, "", this.wkbook );
			if( pos < 0 )
			{
				pos = Chart.findRecPosition( ap.chartArr, PlotArea.class );
			}
			else
			{    // otherwise, TextDisp(s) exist already; position correctly
				// set font to other axis font as default
				TextDisp td = (TextDisp) ap.chartArr.get( pos );
				linkedtd.setFontId( td.getFontId() );
				if( wType != XAXIS )
				{
					pos++;
				}
			}
			if( pos < 0 )
			{
				pos = Chart.findRecPosition( ap.chartArr, ChartFormat.class );
			}
			linkedtd.setParentChart( this.getParentChart() );
			ap.chartArr.add( pos, linkedtd );
		}
	}

	/**
	 * set the Title associated with this Axis (through the linked TextDisp record)
	 *
	 * @param l
	 */
	public void setTitle( String l )
	{
		if( l == null )
		{
			this.removeTitle();
			return;
		}
		if( linkedtd == null )
		{
			getTitleTD( true );    // finds or adds TextDisp records for this axis (for labels and fonts)
		}
		linkedtd.setText( l );
	}

	/**
	 * set the font index for this Axis (for title)
	 *
	 * @param fondId
	 */
	public void setFont( int fondId )
	{
		if( linkedtd == null )
		{
			getTitleTD( true );    // finds or adds TextDisp records for this axis (for labels and fonts)
		}
		linkedtd.setFontId( fondId );
	}

	/**
	 * returns the font of the Axis Title
	 */
	@Override
	public com.extentech.formats.XLS.Font getFont()
	{
		if( linkedtd == null )
		{
			getTitleTD( true );    // finds or adds TextDisp records for this axis (for labels and fonts)
		}
		int idx = linkedtd.getFontId();
		return this.getWorkBook().getFont( idx );
	}

	/**
	 * return the Font object used for Axis labels
	 *
	 * @return
	 */
	public com.extentech.formats.XLS.Font getLabelFont()
	{
		try
		{
			Fontx fx = (Fontx) Chart.findRec( chartArr, Fontx.class );
			return this.getParentChart().getWorkBook().getFont( fx.getIfnt() );
		}
		catch( NullPointerException e )
		{
			return this.getParentChart().getDefaultFont();
		}

	}

	/**
	 * utility method to find the CatserRange rec associated with this Axis
	 *
	 * @return
	 */
	protected CatserRange getCatserRange( boolean bCreate )
	{
		CatserRange csr = (CatserRange) Chart.findRec( chartArr, CatserRange.class );
		if( csr == null )
		{
			csr = (CatserRange) CatserRange.getPrototype();
			csr.setParentChart( this.getParentChart() );
			chartArr.add( 0, csr );
		}
		return csr;
	}

	/**
	 * returns true of this axis displays major gridlines
	 *
	 * @return
	 */
	protected boolean hasGridlines( int type )
	{
		int j = Chart.findRecPosition( chartArr, AxisLineFormat.class );
		if( j != -1 )
		{
			try
			{
				while( j < chartArr.size() )
				{
					AxisLineFormat al = (AxisLineFormat) chartArr.get( j );
					int id = al.getId();
					if( id == type ) // Y Axis draw gridline usually true, trap if do not have (see below)          		
					{
						return true;
					}
					j += 2;    // Skip line format
				}
			}
			catch( ClassCastException e )
			{
			}
		}
		return false;
	}

	/**
	 * returns the SVG necessary to describe the desired line (referenced by id)
	 *
	 * @param id @see AxisLineFormat id types
	 * @return
	 */
	protected String getLineSVG( int id )
	{
		LineFormat lf = getAxisLine( id );
		if( lf != null )
		{
			return lf.getSVG();
		}
		return "";
	}

	/**
	 * return the lineformat rec for the given axis line type
	 *
	 * @param type AxisLineFormat type
	 * @return Line Format
	 * @see AxisLineFormat.ID_MAJOR_GRID, et. atl.
	 */
	protected LineFormat getAxisLine( int type )
	{
		int j = getAxisLineFormat( type );
		if( j > -1 )
		{
			return (LineFormat) chartArr.get( j + 1 );
		}

		return null;
	}

	/**
	 * return the AxisLineFormat of the desired type, or create if none and bCreate
	 *
	 * @param type    AxisLineFormat Type:  AxisLineFormat.ID_AXIS_LINE, AxisLineFormat.ID_MAJOR_GRID, AxisLineFormat.ID_MINOR_GRID, AxisLineFormat.ID_WALLORFLOOR
	 * @param bCreate
	 * @return
	 */
	protected AxisLineFormat getAxisLineFormat( int type, boolean bCreate )
	{
		AxisLineFormat alf = null;
		int j = Chart.findRecPosition( chartArr, AxisLineFormat.class );
		if( j == -1 && !bCreate )
		{
			return null;
		}
		if( j > -1 )
		{
			try
			{
				while( j < chartArr.size() )
				{
					alf = (AxisLineFormat) chartArr.get( j );
					if( alf.getId() == type )
					{
						return alf;
					}
					else if( alf.getId() > type )
					{
						break;
					}
					j += 2;
				}
			}
			catch( ClassCastException e )
			{
				;
			}
		}
		j = 1;
		for(; j < chartArr.size(); j++ )
		{
			BiffRec b = (BiffRec) chartArr.get( j );
			if( b.getOpcode() == AREAFORMAT ||
					b.getOpcode() == GELFRAME ||
					b.getOpcode() == 2213 ||	/* TextPropsStream */
					b.getOpcode() == 2212 ||  /* ShapePropsStream */
					b.getOpcode() == 2206 )	/* CtrlMlFrt */
			{
				break;
			}
		}
		alf = (AxisLineFormat) AxisLineFormat.getPrototype();
		alf.setId( type );    // default has major gridlines
		chartArr.add( j++, alf );
		alf.setParentChart( this.getParentChart() );
		LineFormat lf = (LineFormat) LineFormat.getPrototype();
		lf.setParentChart( this.getParentChart() );
		chartArr.add( j, lf );
		return alf;

	}

	/**
	 * return the index of the AxisLineFormat of the desired type, -1 if none
	 *
	 * @param type AxisLineFormat Type:  AxisLineFormat.ID_AXIS_LINE, AxisLineFormat.ID_MAJOR_GRID, AxisLineFormat.ID_MINOR_GRID, AxisLineFormat.ID_WALLORFLOOR
	 * @return
	 */
	protected int getAxisLineFormat( int type )
	{
		AxisLineFormat alf = null;
		int j = Chart.findRecPosition( chartArr, AxisLineFormat.class );
		if( j == -1 )
		{
			return j;
		}
		try
		{
			while( j < chartArr.size() )
			{
				alf = (AxisLineFormat) chartArr.get( j );
				if( alf.getId() == type )
				{
					return j;
				}
				else if( alf.getId() > type )
				{
					break;
				}
				j += 2;
			}
		}
		catch( ClassCastException e )
		{
			;
		}
		if( j > chartArr.size() )
		{
			return -1;    // not found
		}
		return j;

	}

	/**
	 * gets or creates YMult (value multiplier) record
	 *
	 * @param bCreate
	 * @return
	 */
	protected YMult getYMultRec( boolean bCreate )
	{
		YMult ym = (YMult) Chart.findRec( chartArr, YMult.class );
		if( ym == null && bCreate )
		{
			ym = (YMult) YMult.getPrototype();
			ym.setParentChart( this.getParentChart() );
			chartArr.add( 1, ym );    // 2nd, after ValueRange
		}
		return ym;
	}

	/**
	 * returns true if this axis is reversed
	 * <br>If horizontal axis, default= on bottom, reversed= on top
	 * <br>If vertical axis, default= LHS, reversed= RHS
	 *
	 * @return
	 */
	public boolean isReversed()
	{
		if( wType == XAXIS )
		{
			CatserRange c = getCatserRange( false );
			if( c != null )
			{
				return c.isReversed();
			}
		}
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{ // shouldn't
			return v.isReversed();
		}
		return false;
	}

	public String getNumberFormat()
	{
		Ifmt f = (Ifmt) Chart.findRec( this.chartArr, Ifmt.class );
		int i = 0;
		if( f != null )
		{
			i = f.getFmt();
		}
		else
		{
			// see if have series-specific formats
			java.util.Vector s = this.getParentChart().getAllSeries( -1 );
			if( s.size() > 0 )
			{
				if( wType == YAXIS )
				{
					return ((Series) s.get( 0 )).getSeriesFormatPattern();
				}
				else    // see if it's a value X axis with a custom 
				{
					return ((Series) s.get( 0 )).getCategoryFormatPattern();
				}
			}
		}
		return "General";
	}

	/**
	 * return the JSON/Dojo representation of this axis
	 * chartType int necessary for parsing AXIS options: horizontal charts "switch" axes ...
	 * All options have NOT been gathered at this point
	 *
	 * @return JSONObject
	 */
	public JSONObject getJSON( com.extentech.ExtenXLS.WorkBookHandle wbh, int chartType, double yMax, double yMin, int nSeries )
	{
		JSONObject axisJSON = new JSONObject();
		JSONObject axisOptions = new JSONObject();
		try
		{
			if( wType == YAXIS && chartType != ChartConstants.BARCHART )
			{
				axisOptions.put( "vertical", true );
			}
			else if( wType == XAXIS && chartType == ChartConstants.BARCHART )
			{
				axisOptions.put( "vertical", true );
			}

			// 20090721 KSC: dojo 1.3.1 has label element
			axisOptions.put( "label", getTitle() );

			// TODO: Dojo Axis Options:
			// fixLower, fixUpper	("minor"/"major")
			// includeZero, natural (true)
			// min, max
			// minorTicks, majorTicks, microTicks (true/false)
			// minorTickStep, majorTickStep, microTickStep (#)
			for( int i = 0; i < chartArr.size(); i++ )
			{
				BiffRec b = (BiffRec) chartArr.get( i );
				if( b instanceof CatserRange )
				{
					CatserRange c = ((CatserRange) b);
					// for x/Category axis:  if has labels, gather and input into axis label JSON array
					String[] categories = this.getParentChart().getCategories( -1 );
					if( categories != null )
					{
						// Category Labels
						JSONArray labels = new JSONArray();
						if( c.getCrossBetween() )
						{ // categories appear mid-axis so put "spacers" at 0 and max
							axisOptions.put( "includeZero", true );        // always true????
							JSONObject nullCat = new JSONObject();
							nullCat.put( "value", 0 );
							nullCat.put( "text", "" );
							labels.put( nullCat );
						}
						JSONArray cats = CellRange.getValuesAsJSON( categories[0], wbh );    // parse category range into JSON Array
						for( int z = 0; z < cats.length(); z++ )
						{
							JSONObject aCat = new JSONObject();
							aCat.put( "value",
							          z + c.getCatLabel() );    // should= +1 i.e. a category label appears with each category			            	
							aCat.put( "text", cats.get( z ) );
							labels.put( aCat );
						}
						if( c.getCrossBetween() )
						{ // categories appear mid-axis so put "spacers" at 0 and max
							JSONObject nullCat = new JSONObject();
							nullCat.put( "value", cats.length() + c.getCatLabel() );
							nullCat.put( "text", "" );
							labels.put( nullCat );
							axisOptions.put( "max", cats.length() + c.getCatLabel()/*+1*/ );
						}
						axisOptions.put( "labels", labels );
						// Defaults axis values ...
						axisOptions.put( "fixLower", "major" );        // what do these do??
						axisOptions.put( "fixUpper", "major" );
					}

//	            	if (c.getCatCross()!=1)
					if( c.getCatMark() != 1 )
					{
/*	            		The catMark field defines how often tick marks appear along the category or series axis. A value of 01 indicates 
	            		that a tick mark will appear between each category or series; a value of 02 means a label appears between every 
	            		other category or series, etc.
*/
					}

//	            	if (c.getCrossMax())
				}
				else if( b instanceof ValueRange )
				{
					ValueRange v = (ValueRange) b;
					if( wType == YAXIS )    // normal 
					{
						v.setMaxMin( yMax, yMin ); // must do first
					}
					else
					{
						v.setMaxMin( nSeries,
						             0 );    // scatter and bubble charts have X axis with Value Range, scale is 0 to number of series
					}

					// y major/minor scales
					axisOptions.put( "min", v.getMin() );
					axisOptions.put( "max", v.getMax() );
					axisOptions.put( "majorTickStep", v.getMajorTick() );
				}
				else if( b instanceof AxisLineFormat )
				{
					JSONObject gridJSON = new JSONObject();
					gridJSON.put( "type", "Grid" );
					int id = ((AxisLineFormat) b).getId();
					switch( id )
					{
						case AxisLineFormat.ID_MAJOR_GRID:
							if( wType == XAXIS || chartType == ChartConstants.BARCHART )
							{
								gridJSON.put( "hMajorLines", false );
							}
							else
							{
								gridJSON.put( "vMajorLines", false );
							}
							break;
						case AxisLineFormat.ID_MINOR_GRID:
							break;
					}
					axisJSON.put( "back_grid", gridJSON );
				}
			}
			if( wType == YAXIS )
			{
				axisJSON.put( "y", axisOptions );
			}
			if( wType == XAXIS )
			{
				axisJSON.put( "x", axisOptions );
			}
		}
		catch( JSONException e )
		{
			Logger.logErr( "Error getting Axis JSON: " + e );
		}
		return axisJSON;
	}

	/**
	 * interface for setting Axis rec-specific XML options
	 * in a generic fashion
	 *
	 * @see ExtenXLS.handleChartElement
	 * @see ChartHandle.getXML
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		if( op.equalsIgnoreCase( "Label" ) )
		{
			this.setTitle( val );
			return true;
		}
		else if( op.equalsIgnoreCase( "CatCross" ) )
		{
			getCatserRange( true ).setCatCross( Integer.parseInt( val ) );
			return true;
		}
		else if( op.equalsIgnoreCase( "LabelCross" ) )
		{
			getCatserRange( true ).setCatLabel( Integer.parseInt( val ) );
			return true;
		}
		else if( op.equalsIgnoreCase( "Marks" ) )
		{
			getCatserRange( true ).setCatMark( Integer.parseInt( val ) );
			return true;
		}
		else if( op.equalsIgnoreCase( "CrossBetween" ) )
		{
			getCatserRange( true ).setCrossBetween( val.equals( "true" ) );
			return true;
		}
		else if( op.equalsIgnoreCase( "CrossMax" ) )
		{
			getCatserRange( true ).setCrossMax( val.equals( "true" ) );
			return true;
		}
		else if( op.equalsIgnoreCase( "MajorGridLines" ) )
		{
			if( val.equals( "false" ) )
			{
				int j = getAxisLineFormat( AxisLineFormat.ID_MAJOR_GRID );
				if( j > -1 )
				{
					chartArr.remove( j );    // remove AxisLineFormat
					chartArr.remove( j );    // remove corresponding Line Format
				}
			}
			else
			{    // add major grid lines
				getAxisLineFormat( AxisLineFormat.ID_MAJOR_GRID, true );
			}
		}
		else if( op.equalsIgnoreCase( "MinorGridLines" ) )
		{
			if( val.equals( "false" ) )
			{
				int j = getAxisLineFormat( AxisLineFormat.ID_MINOR_GRID );
				if( j > -1 )
				{
					chartArr.remove( j );    // remove AxisLineFormat
					chartArr.remove( j );    // remove corresponding Line Format
				}
			}
			else
			{    // add major grid lines
				getAxisLineFormat( AxisLineFormat.ID_MINOR_GRID, true );
			}
		}
		else if( op.equalsIgnoreCase( "AddArea" ) )
		{
			if( wType == XAXIS )
			{
				AxisLineFormat alf0 = (AxisLineFormat) AxisLineFormat.getPrototype();
				alf0.setId( AxisLineFormat.ID_AXIS_LINE );
				this.addChartRecord( alf0 );
				LineFormat lf0 = (LineFormat) LineFormat.getPrototype( 0, 0 );
				this.addChartRecord( lf0 );
			}
			AxisLineFormat alf = (AxisLineFormat) AxisLineFormat.getPrototype();
			alf.setId( AxisLineFormat.ID_WALLORFLOOR );
			this.addChartRecord( alf );
			LineFormat lf = (LineFormat) LineFormat.getPrototype( 0, -1 );
			if( wType == 1 )
			{
				lf.setLineStyle( 5 );    //none
			}
			this.addChartRecord( lf );
			AreaFormat af = (AreaFormat) AreaFormat.getPrototype( wType );
			this.addChartRecord( af );
			return true;
		}
		else if( op.equals( "AreaFg" ) )
		{    // custom foreground on Wall, Side or Floor
			AreaFormat af = (AreaFormat) Chart.findRec( chartArr, AreaFormat.class );
			af.seticvFore( Integer.valueOf( val ).intValue() );
			return true;
		}
		else if( op.equals( "AreaBg" ) )
		{    // custom bg on Wall, SIde or Floor
			AreaFormat af = (AreaFormat) Chart.findRec( chartArr, AreaFormat.class );
			af.seticvBack( Integer.valueOf( val ).intValue() );
			return true;
		}
		else if( linkedtd != null )
		{ // see if associated TextDisp can handle
			return linkedtd.setChartOption( op, val );
		}
		else if( op.equalsIgnoreCase( "MajorGridLines" ) )
		{
			if( val.equals( "false" ) )
			{
				int j = getAxisLineFormat( AxisLineFormat.ID_MAJOR_GRID );
				if( j > -1 )
				{
					chartArr.remove( j );    // remove AxisLineFormat
					chartArr.remove( j );    // remove corresponding Line Format
				}
			}
			else
			{    // add major grid lines
				getAxisLineFormat( AxisLineFormat.ID_MAJOR_GRID, true );
			}
		}
		else if( op.equalsIgnoreCase( "MinorGridLines" ) )
		{
			if( val.equals( "false" ) )
			{
				int j = getAxisLineFormat( AxisLineFormat.ID_MINOR_GRID );
				if( j > -1 )
				{
					chartArr.remove( j );    // remove AxisLineFormat
					chartArr.remove( j );    // remove corresponding Line Format
				}
			}
			else
			{    // add major grid lines
				getAxisLineFormat( AxisLineFormat.ID_MINOR_GRID, true );
			}
		}
		return false;
	}

	/**
	 * sets a specific OOXML axis option; most options apply to any type of axis (Cat, Value, Ser, Date)
	 * <br>can be one of:
	 * <br>axPos		-   position of the axis (b, t, r, l)
	 * <br>crosses			possible crossing points (autoZero, max, min)
	 * <br>crossBeteween	whether axis crosses the cat. axis between or on categories (value axis only)  (between, midCat)
	 * <br>crossesAt		where on axis the perpendicular axis crosses (double val)
	 * <br>lblAlign			text alignment for tick labels (ctr, l, r) (cat only)
	 * <br>lblOffset		distance of labels from the axis (0-1000)  (cat only)
	 * <br>majorTickMark	major tick mark position (cross, in, none, out)
	 * <br>minorTickMark	minor tick mark position ("")
	 * <br>tickLblPos		tick label position (high, low, nextTo, none)
	 * <br>tickLblSkip		how many tick labels to skip between label (int >= 1)	(cat only)
	 * <br>tickMarkSkip		how many tick marks to skip betwen ticks (int >= 1)		(cat only)
	 * <br>majorUnit		distance between major tick marks (val, date ax only) (double >= 0)
	 * <br>minorUnit		distance between minor tick marks (val, date ax only) (double >= 0)
	 * <br>MajorGridLines
	 * <br>MinorGridLines
	 *
	 * @param op
	 * @param val
	 */
	private void setOption( String op, String val )
	{
		if( op.equals( "axPos" ) )
		{                    // val= "b" (bottom) "l", "t", "r"   -->?????
			axPos = val;    // for now	    
		}
		else if( op.equals( "lblOffset" ) || op.equals( "lblAlgn" ) )
		{
			CatLab cl = (CatLab) Chart.findRec( chartArr, CatLab.class );
			if( cl == null )
			{
				cl = (CatLab) CatLab.getPrototype();
				cl.setParentChart( this.getParentChart() );
				chartArr.add( 1, cl );    // second in chart array, after CatSerRange
			}
			cl.setOption( op, val );
		}
		else if( op.equals( "tickLblPos" ) ||        // val= high (=at high end of perp. axis), low (=at low end of perp. axis), nextTo, none (=no axis labels) Tick
				op.equals( "majorTickMark" ) ||    // major tick marks (cross, in, none, out)
				op.equals( "minorTickMark" ) )
		{    // minor tick marks (cross, in, none, out)
			Tick t = (Tick) Chart.findRec( chartArr, Tick.class );
			t.setOption( op, val );
		}
		else if( op.equalsIgnoreCase( "MajorGridLines" ) )
		{
			if( val.equals( "false" ) )
			{
				int j = getAxisLineFormat( AxisLineFormat.ID_MAJOR_GRID );
				if( j > -1 )
				{
					chartArr.remove( j );    // remove AxisLineFormat
					chartArr.remove( j );    // remove corresponding Line Format
				}
			}
			else
			{    // add major grid lines
				getAxisLineFormat( AxisLineFormat.ID_MAJOR_GRID, true );
			}
		}
		else if( op.equalsIgnoreCase( "MinorGridLines" ) )
		{
			if( val.equals( "false" ) )
			{
				int j = getAxisLineFormat( AxisLineFormat.ID_MINOR_GRID );
				if( j > -1 )
				{
					chartArr.remove( j );    // remove AxisLineFormat
					chartArr.remove( j );    // remove corresponding Line Format
				}
			}
			else
			{    // add major grid lines
				getAxisLineFormat( AxisLineFormat.ID_MINOR_GRID, true );
			}
		}
		else
		{        // valuerange, caterrange options	-- crosses, crossBetween, crossesAt, tickMarkSkip (cat only), tickLblSkip (cat only), majorUnit (val only), minorUnit (val only)  
// KSC: TESTING
//Logger.logInfo("Setting option: " + op + "=" + val);
			for( int i = 0; i < chartArr.size(); i++ )
			{
				BiffRec b = (BiffRec) chartArr.get( i );
				if( b instanceof CatserRange )
				{
					if( ((CatserRange) b).setOption( op, val ) )
					{
						break;
					}
				}
				else if( b instanceof ValueRange )
				{
					if( ((ValueRange) b).setOption( op, val ) )
					{
						break;
					}
				}
			}
		}
	}

	/**
	 * get the desired axis option for this axis
	 * <br>can be one of:
	 * <br>axPos		-   position of the axis (b, t, r, l)
	 * <br>crosses			possible crossing points (autoZero, max, min)
	 * <br>crossBeteween	whether axis crosses the cat. axis between or on categories (value axis only)  (between, midCat)
	 * <br>crossesAt		where on axis the perpendicular axis crosses (double val)
	 * <br>lblAlign			text alignment for tick labels (ctr, l, r) (cat only)
	 * <br>lblOffset		distance of labels from the axis (0-1000)  (cat only)
	 * <br>majorTickMark	major tick mark position (cross, in, none, out)
	 * <br>minorTickMark	minor tick mark position ("")
	 * <br>tickLblPos		tick label position (high, low, nextTo, none)
	 * <br>tickLblSkip		how many tick labels to skip between label (int >= 1)	(cat only)
	 * <br>tickMarkSkip		how many tick marks to skip betwen ticks (int >= 1)		(cat only)
	 * <br>majorUnit		distance between major tick marks (val, date ax only) (double >= 0)
	 * <br>minorUnit		distance between minor tick marks (val, date ax only) (double >= 0)
	 *
	 * @param op String option name
	 * @return String val of option or ""
	 */
	public String getOption( String op )
	{
		if( op.equals( "axPos" ) )
		{                    // val= "b" (bottom) "l", "t", "r"   -->?????
			return axPos;    // for now -- can't find matching Axis attribute	    	
		}
		else if( op.equals( "lblAlign" ) || op.equals( "lblOffset" ) )
		{
			CatLab c = (CatLab) Chart.findRec( chartArr, CatLab.class );
			if( c != null )
			{
				return ((CatLab) c).getOption( op );
			}
			return null;    // use defaults	    	
		}
		else if( op.equals( "crossesAt" ) ||        // specifies where axis crosses		  -- numCross or catCross
				op.equals( "orientation" ) ||        // axis orientation minMax or maxMin  -- fReverse 
				op.equals( "crosses" ) ||            // specifies how axis crosses it's perpendicular axis (val= max, min, autoZero)  -- fbetween + fMaxCross?/fAutoCross + fMaxCross	    		
				op.equals( "max" ) ||                // axis max - valueRange only?
				op.equals( "max" ) ||                // axis min- valueRange only?
				op.equals( "tickLblSkip" ) ||    //val= how many tick labels to skip btwn label -- catLabel -- Catserrange only??
				op.equals( "tickMarkSkip" ) ||    //val= how many tick marks to skip before next one is drawn -- catMark -- catsterrange only?
				op.equals( "crossBetween" ) )
		{    // value axis only -- val= between, midCat, crossBetween 	    										
			// logScale-- ValueRange
			for( int i = 0; i < chartArr.size(); i++ )
			{
				BiffRec b = (BiffRec) chartArr.get( i );
				if( b instanceof CatserRange )
				{
					return ((CatserRange) b).getOption( op );
				}
				else if( b instanceof ValueRange )
				{
					return ((ValueRange) b).getOption( op );
				}
			}
			//TICK Options
		}
		else if( op.equals( "tickLblPos" ) ||        // val= high (=at high end of perp. axis), low (=at low end of perp. axis), nextTo, none (=no axis labels) Tick
				op.equals( "majorTickMark" ) ||    // major tick marks (cross, in, none, out)
				op.equals( "minorTickMark" ) )
		{    // minor tick marks (cross, in, none, out)
			Tick t = (Tick) Chart.findRec( chartArr, Tick.class );
			return t.getOption( op );
		}
		return null;
	}

	/**
	 * return the JSON/Dojo representation of this axis
	 * chartType int necessary for parsing AXIS options: horizontal charts "switch" axes ...
	 *
	 * @return JSONObject
	 */
	public JSONObject getMinMaxJSON( com.extentech.ExtenXLS.WorkBookHandle wbh, int chartType, double yMax, double yMin, int nSeries )
	{
		JSONObject axisJSON = new JSONObject();
		JSONObject axisOptions = new JSONObject();
		try
		{
			for( int i = 0; i < chartArr.size(); i++ )
			{
				BiffRec b = (BiffRec) chartArr.get( i );
				if( b instanceof CatserRange )
				{
					CatserRange c = ((CatserRange) b);
					// for x/Category axis:  if has labels, gather and input into axis label JSON array
					String[] categories = this.getParentChart().getCategories( -1 );
					if( categories != null )
					{
						JSONArray cats = CellRange.getValuesAsJSON( categories[0], wbh );    // parse category range into JSON Array
						axisOptions.put( "max", cats.length() + c.getCatLabel()/*+1*/ );
					}
					break;    // only need this per axis
				}
				else if( b instanceof ValueRange )
				{
					ValueRange v = (ValueRange) b;
					if( wType == YAXIS )    // normal 
					{
						v.setMaxMin( yMax, yMin );    // must do first
					}
					else
					{
						v.setMaxMin( nSeries,
						             0 );    // scatter and bubble charts have X axis with Value Range, scale is 0 to number of series
					}

					// y major/minor scales
					axisOptions.put( "min", v.getMin() );
					axisOptions.put( "max", v.getMax() );
					axisOptions.put( "majorTickStep", v.getMajorTick() );
					break;    // only need this per axis
				}
			}
			if( wType == YAXIS )
			{
				axisJSON.put( "y", axisOptions );
			}
			if( wType == XAXIS )
			{
				axisJSON.put( "x", axisOptions );
			}
		}
		catch( JSONException e )
		{
			Logger.logErr( "Error getting Axis JSON: " + e );
		}
		return axisJSON;
	}

	/**
	 * return the OOXML shape property for this axis
	 *
	 * @return
	 */
	public SpPr getSpPr()
	{
		return shapeProps;
	}

	/**
	 * define the OOXML shape property for this axis from an existing spPr element
	 */
	public void setSpPr( SpPr sp )
	{
		shapeProps = sp;
		//shapeProps.setNS("c");
	}

	/**
	 * return the OOXML title element for this axis
	 *
	 * @return
	 */
	public com.extentech.formats.OOXML.Title getOOXMLTitle()
	{
		return ttl;
	}

	/**
	 * set the OOXML title element for this axis
	 *
	 * @param t
	 */
	public void setOOXMLTitle( Title t )
	{
		ttl = t;
	}

	/**
	 * return the OOXML txPr element for this axis
	 *
	 * @return
	 */
	public TxPr gettxPr()
	{
		return txpr;
	}

	/**
	 * set the OOXML title element for this axis
	 *
	 * @param t
	 */
	public void settxPr( TxPr t )
	{
		txpr = (TxPr) t.cloneElement();
	}

	public String toString()
	{
		String s = "";
		switch( wType )
		{
			case XAXIS:
				s = "XAxis";
				break;
			case YAXIS:
				s = "YAxis";
				break;
			case ZAXIS:
				s = "ZAxis";
				break;
			case XVALAXIS:
				s = "XValAxis";
				break;
		}
		if( linkedtd != null )
		{
			s = s + " " + linkedtd.toString();
		}
		return s;
	}

	/**
	 * return the maximum value of this Value or Y axis scale
	 *
	 * @return
	 */
	public double getMaxScale( double[] minmax )
	{
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{ // shouldn't
			if( v.isAutomaticMax() )
			{
				v.setMaxMin( minmax[1], minmax[0] );
			}
			return v.getMax();
		}
		return -1;
	}

	protected double[] getMinMax()
	{
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{ // shouldn't
			return new double[]{ v.getMin(), v.getMax() };
		}
		return new double[]{ 0, 0 };

	}

	/**
	 * return the minimum value of this value or Y axis scale
	 *
	 * @return
	 */
	public double getMinScale( double[] minmax )
	{
/*
 * Because a horizontal (category) axis (axis: A line bordering the chart plot area used as 
 * a frame of reference for measurement.  * The y axis is usually the vertical axis and contains data. The x-axis is usually the horizontal axis and contains categories.) 
 * displays text labels instead of numeric intervals, 
 * there are fewer scaling options that you can change than there are for a vertical (value) axis. 
 * However, you can change the number of categories to display between tick marks, the order in which 
 * to display categories, and the point where the two axes cross.        	
 */
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{
			/*
			 if (wType==YAXIS)	// normal 
            	v.setMaxMin(yMax, yMin); 	// must do first
            else
            	v.setMaxMin(nSeries, 0);	// scatter and bubble charts have X axis with Value Range, scale is 0 to number of series  
			 */
// why? should already			v.setParentChart(this.getParentChart());
			if( v.isAutomaticMin() )
			{
				v.setMaxMin( minmax[1], minmax[0] );
			}
			return v.getMin();
		}
		return -1;
	}

	/**
	 * return the major tick unit of this Y or Value axis
	 *
	 * @return
	 */
	public int getMajorUnit()
	{
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{
			return new Double( v.getMajorTick() ).intValue();
		}
		return 10;    // try a default
	}

	/**
	 * return the minor tick unit of this Y or Value axis
	 *
	 * @return
	 */
	public int getMinorUnit()
	{
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{
			return new Double( v.getMinorTick() ).intValue();
		}
		return 0;    // try a default
	}

	/**
	 * returns true if either Automatic min or max scale is set for the Y or Value axis
	 *
	 * @return
	 */
	public boolean isAutomaticScale()
	{
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{
			return (v.isAutomaticMin() || v.isAutomaticMax());
		}
		return false;
	}

	/**
	 * sets the automatic scale option on or off for the Y or Value axis
	 * <br>Automatic Scaling will automatically set the scale maximum, minimum and tick units
	 *
	 * @param b
	 */
	public void setAutomaticScale( boolean b )
	{
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{
			v.setAutomaticMin( b );
			v.setAutomaticMax( b );
		}
	}

	/**
	 * set the minimum value of this axis scale
	 * <br>Note: this disables automatic scaling
	 *
	 * @param Min
	 */
	public void setMinScale( int min )
	{
		// TODO: also update ticks?   
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{ // shouldn't
			v.setMin( min );
		}
	}

	/**
	 * set the maximum value of this axis scale
	 * <br>Note: this disables automatic scaling
	 *
	 * @param Max
	 */
	public void setMaxScale( int max )
	{
		// TODO: also update ticks?   
		ValueRange v = (ValueRange) Chart.findRec( this.chartArr, ValueRange.class );
		if( v != null )
		{ // shouldn't
			v.setMax( max );
		}
	}

	/**
	 * sets the axis labels position or placement to the desired value (these match Excel placement options)
	 * <p>Possible options:
	 * <br>Axis.INVISIBLE - hides the axis
	 * <br>Axis.LOW - low end of plot area
	 * <br>Axis.HIGH - high end of plot area
	 * <br>Axis.NEXTTO- next to axis (default)
	 *
	 * @param Placement - int one of the Axis placement constants listed above
	 */
	public void setAxisPlacement( int Placement )
	{
		Tick t = (Tick) Chart.findRec( this.chartArr, Tick.class );
		if( t != null )
		{    // shoudn't 
			switch( Placement )
			{
				case Axis.INVISIBLE:
					t.setOption( "tickLblPos", "none" );
					break;
				case Axis.LOW:
					t.setOption( "tickLblPos", "low" );
					break;
				case Axis.HIGH:
					t.setOption( "tickLblPos", "high" );
					break;
				case Axis.NEXTTO:
					t.setOption( "tickLblPos", "nextTo" );
			}
		}
	}

	/**
	 * returns the Axis Label Placement or position as an int
	 * <p>One of:
	 * <br>Axis.INVISIBLE - axis is hidden
	 * <br>Axis.LOW - low end of plot area
	 * <br>Axis.HIGH - high end of plot area
	 * <br>Axis.NEXTTO- next to axis (default)
	 *
	 * @return int - one of the Axis Label placement constants above
	 */
	public int getAxisPlacement()
	{
		Tick t = (Tick) Chart.findRec( this.chartArr, Tick.class );
		if( t != null )
		{    // shoudn't 
			String p = t.getOption( "tickLblPos" );
			if( p == null || p.equals( "none" ) )
			{
				return Axis.INVISIBLE;
			}
			else if( p.equals( "low" ) )
			{
				return Axis.LOW;
			}
			else if( p.equals( "high" ) )
			{
				return Axis.HIGH;
			}
			else if( p.equals( "nextTo" ) )
			{
				return Axis.NEXTTO;
			}
		}
		return Axis.INVISIBLE;
	}

	/**
	 * parse OOXML axis element
	 *
	 * @param xpp     XmlPullParser positioned at correct elemnt
	 * @param axisTag catAx, valAx, serAx, dateAx
	 * @param lastTag Stack of element names
	 */
	// noMultiLvlLbl -- val= 1 (true) means draw labels as flat text; not included or 0 (false)= draw labels as a heirarchy
	public void parseOOXML( XmlPullParser xpp, String axisTag, Stack<String> lastTag, WorkBookHandle bk )
	{
// crossAx -- need to parse?
// auto -- need to parse?    	
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "scaling" ) )
					{        // additional axis settings
						lastTag.push( tnm );
						Scaling sc = (Scaling) Scaling.parseOOXML( xpp, lastTag );
						String s = sc.getOption( "orientation" );
						if( s != null )
						{
							this.setOption( "orientation", s );
						}
						s = sc.getOption( "min" );
						if( s != null )
						{
							this.setOption( "min", s );
						}
						s = sc.getOption( "max" );
						if( s != null )
						{
							this.setOption( "max", s );
						}
						// the below children only have 1 attribute: val
					}
					else if( tnm.equals( "axPos" ) )
					{        // // position of the axis (b, t, r, l)
						this.setOption( tnm, xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "majorGridlines" ) || tnm.equals( "minorGridlines" ) )
					{
						lastTag.push( tnm );
						parseGridlinesOOXML( xpp, lastTag, bk );
					}
					else if( tnm.equals( "title" ) )
					{
						lastTag.push( tnm );
						this.setOOXMLTitle( (Title) Title.parseOOXML( xpp, lastTag, bk ).cloneElement() );
						this.setTitle( this.getOOXMLTitle().getTitle() );
					}
					else if( tnm.equals( "numFmt" ) )
					{
						this.nf = (NumFmt) NumFmt.parseOOXML( xpp ).cloneElement();
					}
					else if( tnm.equals( "majorTickMark" ) ||    // major tick mark position (cross, in, none, out)
							tnm.equals( "minorTickMark" ) ||        // minor tick mark position ("")
							tnm.equals( "tickLblPos" ) )
					{        // tick label position (high, low, nextTo, none)
						this.setOption( tnm, xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "spPr" ) )
					{    // axis shape properties - for axis or gridlines
						lastTag.push( tnm );
						this.setSpPr( (SpPr) SpPr.parseOOXML( xpp, lastTag, bk ).cloneElement() );
					}
					else if( tnm.equals( "txPr" ) )
					{        // text Properties for axis
						lastTag.push( tnm );
						this.settxPr( (TxPr) TxPr.parseOOXML( xpp, lastTag, bk ).cloneElement() );
						// crossesAx = crossing axis id - need ?
					}
					else if( tnm.equals( "crosses" ) ||            // possible crossing points (autoZero, max, min) 
							tnm.equals( "crossesAt" ) )
					{            // where on axis the perpendicular axis crosses (double val)
						this.setOption( tnm, xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "crossBetween" ) )
					{        // whether axis crosses the cat. axis between or on categories (value axis only)  (between, midCat)
						this.setOption( tnm, xpp.getAttributeValue( 0 ) );
					}
					else if(        // cat, date, ser ax only
						// auto -- Date only?
							tnm.equals( "lblAlign" ) ||            // text alignment for tick labels (ctr, l, r)  only for cat
									tnm.equals( "lblOffset" ) ||            // distance of labels from the axis (0-1000)	only for cat, date
									tnm.equals( "tickLblSkip" ) ||        // how many tick labels to skip between label (int >= 1)
									tnm.equals( "tickMarkSkip" ) )
					{        // how many tick marks to skip betwen ticks (int >= 1)
						this.setOption( tnm, xpp.getAttributeValue( 0 ) );
// TODO: noMultiLvlLbl		            	 
					}
					else if(    // val, ser ax + some date ax
							tnm.equals( "crossBeteween" ) ||        // whether axis crosses the cat. axis between or on categories (value axis only)  (between, midCat)
									tnm.equals( "majorUnit" ) ||            // distance between major tick marks (val, date ax only) (double >= 0)
									tnm.equals( "minorUnit" ) )
					{            // distance between minor tick marks (val, date ax only) (double >= 0)
						this.setOption( tnm, xpp.getAttributeValue( 0 ) );
					}
					else if( tnm.equals( "dispUnits" ) )
					{    // valAx only
						parseDispUnitsOOXML( xpp, lastTag );
					}
// TODO: date ax specifics		             
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( axisTag ) )
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
			Logger.logErr( "Axis: " + e.toString() );
		}
	}

	/**
	 * parse major or minor Gridlines element
	 *
	 * @param xpp
	 * @param lastTag
	 */
	private void parseGridlinesOOXML( XmlPullParser xpp, Stack<String> lastTag, WorkBookHandle bk )
	{
		String endTag = (String) lastTag.peek();
		this.setOption( endTag, "true" );
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "spPr" ) )
					{
						lastTag.push( tnm );
						SpPr sppr = (SpPr) SpPr.parseOOXML( xpp, lastTag, bk ).cloneElement();
						LineFormat lf = getAxisLine( (endTag.equals( "majorGridlines" ) ? AxisLineFormat.ID_MAJOR_GRID : AxisLineFormat.ID_MINOR_GRID) );
						lf.setFromOOXML( sppr );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( endTag ) )
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
			Logger.logErr( "parseGridLinesOOXML: " + e.toString() );
		}
	}

	/**
	 * parse the dispUnits child element of valAx
	 * <br>TODO: do not know how to interpret most of these options
	 *
	 * @param xpp
	 * @param lastTag
	 */
	private void parseDispUnitsOOXML( XmlPullParser xpp, Stack<String> lastTag )
	{
		try
		{
			int eventType = xpp.getEventType();
			YMult ym = null;
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "custUnit" ) )
					{
						ym = getYMultRec( true );
						ym.setCustomMultiplier( Double.valueOf( xpp.getAttributeValue( 0 ) ) );
					}
					else if( tnm.equals( "builtInUnit" ) )
					{
						ym = getYMultRec( true );
						ym.setAxMultiplierId( xpp.getAttributeValue( 0 ) );
					}
					// TODO: dispUnitLbl
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( "dispUnits" ) )
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
			Logger.logErr( "parseDispUnitsOOXML: " + e.toString() );
		}
	}

	/**
	 * generate the appropriate ooxml for the given axis
	 *
	 * @param type    0= Category Axis, 1= Value Axis, 2= Ser Axis, 3= Date Axis
	 * @param id
	 * @param crossId crossing axis id
	 * @return ORDER:
	 * Axis Common:
	 * axisId		REQ
	 * scaling	-- orientation
	 * delete
	 * axPos		REQ?
	 * majorGridlines
	 * minorGridlines
	 * title
	 * numFmt
	 * majorTickMark
	 * minorTickMark
	 * tickLblPos
	 * spPr
	 * txPr
	 * crossAx
	 * crosses OR
	 * crossesAt
	 * after:
	 * valAx:
	 * crossBetween
	 * majorUnit
	 * minorUnit
	 * dispUnits
	 * catAx:
	 * auto
	 * lblAlign
	 * lblOffet
	 * tickLblSkip
	 * tickMarkSkip
	 * noMultiLvlLbl
	 * serAx:
	 * tickLblSkip
	 * tickMarkSkip
	 */

	public String getOOXML( int type, String id, String crossId )
	{
		if( this.getParentChart() == null )    // happens on ZAxis, XYValAxis ...
		{
			this.setParentChart( this.ap.getParentChart() );
		}
		boolean from2003 = (!parentChart.getWorkBook().getIsExcel2007());

		StringBuffer axisooxml = new StringBuffer();
		String axis = "";
		switch( type )
		{
			case 0:    // cat axis "X"
				axis = "catAx";
				break;
			case 1:    // val axis "Y" 
				axis = "valAx";
				break;
			case 2: // xval axis "X" axis for multiple val axes: bubble & scatter charts
				axis = "valAx";
				break;
			case 3: // ser	 ("Z" axis)
				axis = "serAx";
				break;
			case 4:  // date		TODO: Not correct or tested!
				axis = "dateAx";
		}
		// axis main element
		axisooxml.append( "<c:" + axis + ">" );
		axisooxml.append( "\r\n" );
		// axId - required
		axisooxml.append( "<c:axId val=\"" + id + "\"/>" );
		axisooxml.append( "\r\n" );
		// scaling - required
		String s = this.getOption( "orientation" );
		double[] d = this.getMinMax();
		if( s != null || d[0] != d[1] )
		{    // if have orientation or min/max set ..
			axisooxml.append( "<c:scaling>\r\n" );
			if( s != null )
			{
				axisooxml.append( "<c:orientation val=\"" + s + "\"/>\r\n" );
			}
			axisooxml.append( "</c:scaling>\r\n" );
		}
		// axPos - required
		if( this.getOption( "axPos" ) != null )
		{
			axisooxml.append( "<c:axPos val=\"" + this.getOption( "axPos" ) + "\"/>" );
			axisooxml.append( "\r\n" );
		}
		else
		{// it's required
			if( this.getParentChart().getChartType() != BARCHART )
			{
				if( axis.equals( "catAx" ) || axis.equals( "serAx" ) )
				{
					axisooxml.append( "<c:axPos val=\"b\"/>" );
				}
				else
				{
					axisooxml.append( "<c:axPos val=\"l\"/>" );
				}
			}
			else
			{
				if( axis.equals( "catAx" ) || axis.equals( "serAx" ) )
				{
					axisooxml.append( "<c:axPos val=\"l\"/>" );
				}
				else
				{
					axisooxml.append( "<c:axPos val=\"b\"/>" );
				}
			}
			axisooxml.append( "\r\n" );
		}
		// major Gridlines
		if( this.hasGridlines( AxisLineFormat.ID_MAJOR_GRID ) )
		{
			axisooxml.append( "<c:majorGridlines>" );
			axisooxml.append( getAxisLine( AxisLineFormat.ID_MAJOR_GRID ).getOOXML() );
			axisooxml.append( "</c:majorGridlines>\r\n" );
		}
		// minor Gridlines
		if( this.hasGridlines( AxisLineFormat.ID_MINOR_GRID ) )
		{
			axisooxml.append( "<c:minorGridlines>" );
			axisooxml.append( getAxisLine( AxisLineFormat.ID_MINOR_GRID ).getOOXML() );
			axisooxml.append( "</c:minorGridlines>\r\n" );
		}
		// Title
		if( this.getOOXMLTitle() != null )
		{
			axisooxml.append( this.getOOXMLTitle().getOOXML() );
		}
		else if( from2003 )
		{    // create OOXML title     		
			if( !this.getTitle().equals( "" ) )
			{
				com.extentech.formats.OOXML.Title ttl = new com.extentech.formats.OOXML.Title( this.getTitle() );
				if( type == 0 )
				{
					ttl.setLayout( .026, .378 );
				}
				else if( type == 1 )
				{
					ttl.setLayout( .468, .863 );
				}
				axisooxml.append( ttl.getOOXML() );
			}
		}
		// numFmt
		if( this.nf != null )
		{
			axisooxml.append( nf.getOOXML( "c:" ) );        //need a default???: axisooxml.append("<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/>");	axisooxml.append("\r\n");
		}
		// majorTickMark
		s = this.getOption( "majorTickMark" );    // default= "cross"
		if( s != null )
		{
			axisooxml.append( "<c:majorTickMark val=\"" + s + "\"/>" );
		}
		// minorTickMark
		s = this.getOption( "minorTickMark" );    // default= "cross"
		if( s != null )
		{
			axisooxml.append( "<c:minorTickMark val=\"" + s + "\"/>" );
		}
		// tickLblPos
		s = this.getOption( "tickLblPos" );    // default= "nextTo"
		if( s != null )
		{
			axisooxml.append( "<c:tickLblPos val=\"" + s + "\"/>" );
		}
		// shape properties
		if( this.getSpPr() != null )
		{
			axisooxml.append( this.getSpPr().getOOXML() );
		}
		// text props
		if( this.gettxPr() != null )
		{
			axisooxml.append( this.gettxPr().getOOXML() );
		}
		else if( from2003 )
		{    // XLS->XLSX
/*     		
 * label font:
 * Fontx fx= (Fontx) Chart.findRec(chartArr, Fontx.class);
   return  this.getParentChart().getWorkBook().getFont(fx.getIfnt());
*/
			int rot = 0;
			Tick t = (Tick) Chart.findRec( this.chartArr, Tick.class );
			if( t != null )
			{    // shoudn't
				rot = t.getRotation();
				/**
				 0= no rotation (text appears left-to-right), 
				 1=  text appears top-~~ are upright, 
				 2= text is rotated 90 degrees counterclockwise,  
				 3= text is rotated
				 */
				// convert BIFF8 rotation to TxPr rotation:  
				switch( rot )
				{
					case 1:
						//????
						break;
					case 2:
						rot = -5400000;
						break;
					case 3:
						rot = 5400000;
						break;
				}
				// TODO: is vert rotation from td? 
				TxPr txpr = new TxPr( this.getLabelFont(), rot, null );
				axisooxml.append( txpr.getOOXML() );
			}
			axisooxml.append( "\r\n" );
		}
		// crossesAx
		axisooxml.append( "<c:crossAx val=\"" + crossId + "\"/>" );
		axisooxml.append( "\r\n" ); // crosses axis ...
		// crosses -- autoZero, max, min
		if( this.getOption( "crosses" ) != null )
		{
			axisooxml.append( "<c:crosses val=\"" + this.getOption( "crosses" ) + "\"/>" );
		}
		axisooxml.append( "\r\n" );// where axis crosses it's perpendicular axis
		if( axis.equals( "catAx" ) || axis.equals( "serAx" ) )
		{
			// auto
			axisooxml.append( "<c:auto val=\"1\"/>\r\n" );
			s = this.getOption( "lblAlgn" );
			if( s != null )
			{
				axisooxml.append( "<c:lblAlgn val=\"" + s + "\"/>\r\n" );
			}
			s = this.getOption( "lblOffset" );
			if( s != null )
			{
				axisooxml.append( "<c:lblOffset val=\"" + s + "\"/>\r\n" );
			}
			s = this.getOption( "tickLblSkip" );
			if( s != null )
			{
				axisooxml.append( "<c:tickLblSkip val=\"" + s + "\"/>\r\n" );
			}
			s = this.getOption( "tickMarkSkip" );
			if( s != null )
			{
				axisooxml.append( "<c:tickMarkSkip val=\"" + s + "\"/>\r\n" );
			}
			// TODO: noMutliLvlLbl
		}
		else
		{    // val or date
			s = this.getOption( "crossBetween" );
			if( s != null )
			{
				axisooxml.append( "<c:crossBetween val=\"" + s + "\"/>\r\n" );
			}
			s = this.getOption( "majorUnit" );
			if( s != null )
			{
				axisooxml.append( "<c:majorUnit val=\"" + s + "\"/>\r\n" );
			}
			s = this.getOption( "minorUnit" );
			if( s != null )
			{
				axisooxml.append( "<c:minorUnit val=\"" + s + "\"/>\r\n" );
			}
// TODO: dispUnit ************************************

		}
		axisooxml.append( "</c:" + axis + ">" );
		axisooxml.append( "\r\n" );
		return axisooxml.toString();
	}

	/**
	 * returns the SVG representation of the desired axis
	 *
	 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @return String SVG
	 */
	public String getSVG( ChartAxes ca, java.util.Map<String, Double> chartMetrics, Object[] categories )
	{
		StringBuffer svg = new StringBuffer();

		// for all axies-- label and title fonts, rotation, tick info ...
		// Title + Label SVG
		String labelfontSVG = "";
		String titlefontSVG = "";

		try
		{
			labelfontSVG = this.getLabelFont().getSVG();    // uses specific or default for chart
		}
		catch( Exception e )
		{    // shouldn't
			labelfontSVG = "font-family='Arial' font-size='9pt' fill='" + ChartType.getDarkColor() + "' ";
		}
		try
		{
			titlefontSVG = linkedtd.getFont( this.getParentChart().getWorkBook() ).getSVG();
		}
		catch( NullPointerException e )
		{
		}
		boolean showMinorTickMarks = false, showMajorTickMarks = true;
		try
		{
			Tick t = (Tick) Chart.findRec( chartArr, Tick.class );
			showMinorTickMarks = t.showMinorTicks();
			showMajorTickMarks = t.showMajorTicks();
		}
		catch( Exception e )
		{
		}

		// BAR CHART AXES ARE SWITCHED - handle seperately for clarity; radar axes are also handled separately as are significantly different than regualr charts
		int charttype = this.getParentChart().getChartType();
		if( charttype == ChartConstants.BARCHART )
		{
			return getSVGBARCHART( ca, titlefontSVG, labelfontSVG, showMinorTickMarks, showMajorTickMarks, chartMetrics, categories );
		}
		if( charttype == ChartConstants.RADARCHART )
		{
			return getSVGRADARCHART( ca, titlefontSVG, labelfontSVG, chartMetrics, categories );
		}

		int wtype = wType;
		if( wtype == XAXIS && (charttype == ChartConstants.SCATTERCHART || charttype == ChartConstants.BUBBLECHART) )
		{    // XY Charts - X Axis is a Value Axis
			wtype = XVALAXIS;
		}

		switch( wtype )
		{
			case XAXIS:
				svg.append( drawXAxisSVG( ca,
				                          titlefontSVG,
				                          labelfontSVG,
				                          showMinorTickMarks,
				                          showMajorTickMarks,
				                          chartMetrics,
				                          categories ) );

				break;
			case YAXIS:
				svg.append( drawYAxisSVG( ca, titlefontSVG, labelfontSVG, showMinorTickMarks, showMajorTickMarks, chartMetrics ) );

				break;
			case ZAXIS:        // ??
				break;

			case XVALAXIS:    // Scatter/Bubble Chart X Value Axis
				svg.append( drawXYValAxisSVG( ca,
				                              titlefontSVG,
				                              labelfontSVG,
				                              showMinorTickMarks,
				                              showMajorTickMarks,
				                              chartMetrics,
				                              categories ) );
				break;
		}
		return svg.toString();
	}

	/**
	 * generate SVG for a basic (non-bar-chart, non-textual) X Axis
	 *
	 * @param ca
	 * @param titlefontSVG
	 * @param labelfontSVG
	 * @param rot
	 * @param showMinorTickMarks
	 * @param showMajorTickMarks
	 * @param chartMetrics       maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @return
	 */
	private String drawXAxisSVG( ChartAxes ca,
	                             String titlefontSVG,
	                             String labelfontSVG,
	                             boolean showMinorTickMarks,
	                             boolean showMajorTickMarks,
	                             java.util.Map<String, Double> chartMetrics,
	                             Object[] categories )
	{
		StringBuffer svg = new StringBuffer();
		// X Axis TICKS x= rectX, x2= x1, y1= recty+h, y2=y1+5 (major) x1 increments by (#Major Ticks-1)/Width (usually 4)
		//              x= rectX+17, x2= x1, y1= recty+h y2=y1+2 (minor) x1 increments by (approx 8)
		// LABELS (CATEGORIES): - do before ticks NOTE: patterns and formatting is applied .adjustCoordinates in order to account for fitting in space

		// when x axis is reversed means that categories are right to left and the y axis is on the RHS
		// when y axis is reversed means the categories are on TOP of the chart and y axis labels are reversed
		double x0, x1, y0, y1;
		double inc;
		java.awt.Font f = null;
		double x = chartMetrics.get( "x" );
		double y = chartMetrics.get( "y" );
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double canvash = chartMetrics.get( "canvash" );
		boolean yAxisReversed = (Boolean) ca.getMetric( "yAxisReversed" );
		boolean xAxisReversed = (Boolean) ca.getMetric( "xAxisReversed" );
		int xAxisRotate = (Integer) ca.getMetric( "xAxisRotate" );
		double XAXISLABELOFFSET = (Double) ca.getMetric( "XAXISLABELOFFSET" );
		double XAXISTITLEOFFSET = (Double) ca.getMetric( "XAXISTITLEOFFSET" );

		int labelRot = (wType == XAXIS ? xAxisRotate : 0);    // TODO: handle Y axis rotation 
		if( labelRot != 0 )
		{
			// get font object so can calculate rotation point
			com.extentech.formats.XLS.Font lf = this.getLabelFont();
			try
			{
				// get awt Font so can compute and fit category in width
				f = new java.awt.Font( lf.getFontName(), lf.getFontWeight(), (int) lf.getFontHeightInPoints() );
			}
			catch( Exception e )
			{
			}
		}
		if( categories != null && categories.length > 0 )
		{ // shouldn't
			// Category Labels - centered within area on X Axis
			inc = w / categories.length;
			svg.append( getCategoriesSVG( x,
			                              y,
			                              w,
			                              h,
			                              inc,
			                              labelRot,
			                              categories,
			                              f,
			                              labelfontSVG,
			                              yAxisReversed,
			                              xAxisReversed,
			                              XAXISLABELOFFSET ) );
			// TICK MARKS
			y0 = y + (!yAxisReversed ? h : 0);    // ticks at bottom edge of axis unless Y axiis is reversed
			x0 = x;        // start at chart x
			int rfY = (!yAxisReversed ? 1 : -1);    // reverse factor :)
			int rfX = (!xAxisReversed ? 1 : -1);    // reverse factor :)
			svg.append( "<g>\r\n" );
			inc = w / (categories.length);    // w/scale factor
			double minorinc = 0;
			if( showMinorTickMarks )
			{
				minorinc = inc / 2;    // half-marks for category axis
			}
			for( double i = 0; i <= categories.length; i++ )
			{
				y1 = y0 + 2 * rfY;    // minor tick mark
				if( showMinorTickMarks )
				{
					for( int j = 0; j < 2; j++ )
					{    // for categories, only option is 1/2 major
						svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n" );
						x0 += minorinc;
					}
				}
				y1 = y0 + 5 * rfY;    // Major tick marks
				if( showMajorTickMarks )
				{
					svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "'" + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n" );
				}
				x0 += inc;
			}
			// bounding edge line
			if( hasPlotAreaBorder() )
			{
				x0 = x + (!xAxisReversed ? w : 0);
				svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "'" + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y + "' x2='" + x0 + "' y2='" + (y + h) + "'/>\r\n" );
			}
			svg.append( "</g>\r\n" );
		}
		// X AXIS TITLE
		int titleRot = (linkedtd != null) ? linkedtd.getRotation() : 0;
		x0 = x + w / 2;
		if( !yAxisReversed )    // TODO: why doesn't "normal" calc work???????
		{
			y0 = canvash - XAXISTITLEOFFSET;
		}
		else
		{
			y0 = y - XAXISTITLEOFFSET - XAXISLABELOFFSET;
		}

		svg.append( getAxisTitleSVG( x0, y0, titlefontSVG, titleRot, "xaxistitle" ) );
		return svg.toString();
	}

	static int YLABELSSPACER_X = 10;
	static int YLABELSPACER_Y = 4;

	/**
	 * generate SVG for a basic (non-bar-chart, value or numeric) Y Axis
	 *
	 * @param ca
	 * @param titlefontSVG
	 * @param labelfontSVG
	 * @param rot
	 * @param showMinorTickMarks
	 * @param showMajorTickMarks
	 * @param chartMetrics       maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @return
	 */
	// TODO: label rotation
	private String drawYAxisSVG( ChartAxes ca,
	                             String titlefontSVG,
	                             String labelfontSVG,
	                             boolean showMinorTickMarks,
	                             boolean showMajorTickMarks,
	                             java.util.Map<String, Double> chartMetrics )
	{
		StringBuffer svg = new StringBuffer();
		// Y or Value Axis- must be non-textual; values obtained in calling method
		// When Y Axis is reversed, scale is reversed and x axis labels and title are on top
		// When X Axis is reversed, Y scale/labels and title are on RHS
		double x0, x1, y0, y1;
		double inc;
		// major and minor tick marks, max and min axis scales NOTE: for MOST category (usually x) axes, they are textual; the max=# of categories; min=0
		double x = chartMetrics.get( "x" );
		double y = chartMetrics.get( "y" );
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double max = chartMetrics.get( "max" );
		double min = chartMetrics.get( "min" );
		double minor = (Double) ca.getMetric( "minor" );
		double major = (Double) ca.getMetric( "major" );
		boolean scaleIsInteger = (major == Math.floor( major ));    // true if scale is in integer form rather than double (used for formatting label text)
		int titleRot = (linkedtd != null) ? linkedtd.getRotation() : 0;
		boolean xAxisReversed = (Boolean) ca.getMetric( "xAxisReversed" );
		boolean yAxisReversed = (Boolean) ca.getMetric( "yAxisReversed" );
		String xPattern = (String) ca.getMetric( "xPattern" );
		String yPattern = (String) ca.getMetric( "yPattern" );
		double YAXISLABELOFFSET = (Double) ca.getMetric( "YAXISLABELOFFSET" );
		double YAXISTITLEOFFSET = (Double) ca.getMetric( "YAXISTITLEOFFSET" );

		// Y Axis GRIDLINES: -- x1=rectX x2=rectx+rect w y1=y2  y1 starts with recty+h, decrements by= approx 27 ==rectheight/# lines *****************************			
		// Major/Minor tick marks
		if( major > 0 )
		{    // if is displaying tick marks/grid lines - usual case
			inc = h / ((max - min) / major);
			double minorinc = 0;
			if( minor > 0 )
			{
				minorinc = inc / (major / minor);
			}
			x0 = x;
			y0 = y + h;
			x1 = x + w;    // entire width
			// GRIDLINES
			String lineSVG = getLineSVG( AxisLineFormat.ID_MAJOR_GRID );
			svg.append( "<g>\r\n" );
			for( double i = min; i <= max; i += major )
			{
				svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "'" + lineSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x1 + "' y2='" + y0 + "'/>\r\n" );
				y0 -= inc;
			}
			svg.append( "</g>\r\n" );
			// TICK MARKS
			svg.append( "<g>\r\n" );
			int rfY = (!yAxisReversed ? 1 : -1); // reverse factor :)
			int rfX = (!xAxisReversed ? 1 : -1); // reverse factor :)
			y0 = y + (!yAxisReversed ? h : 0);    // starts at bottom, goes up to top except if reversed
			lineSVG = ChartType.getStrokeSVG();
			int scale = 0;
			{ // figure out if scale has a fractional portion; if so, determine desired scale and keep to that  
				String s = String.valueOf( major );
				int z = s.indexOf( "." );
				if( z != -1 )
				{
					scale = s.length() - (z + 1);
				}
			}
			int k = 0;    // axis label index
			for( double i = min; i <= max; i += major )
			{
				x0 = x + (!xAxisReversed ? 0 : w);
				x1 = x0 - 2 * rfX;// minor ticks
				y1 = y0;
				if( i < max && minor > 0 )
				{
					if( showMinorTickMarks )
					{
						for( int j = 0; j < (major / minor); j++ )
						{
							svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + lineSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x1 + "' y2='" + y0 + "'/>\r\n" );
							y0 -= minorinc;
						}
					}
				}
				y0 = y1;
				x1 = x0 - 5 * rfX;// Major tick Marks
				if( showMajorTickMarks )
				{
					svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + lineSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x1 + "' y2='" + y0 + "'/>\r\n" );
				}
				// Y Axis Labels
				java.math.BigDecimal bd = new java.math.BigDecimal( i ).setScale( scale,
				                                                                  java.math.BigDecimal.ROUND_HALF_UP );    // if fractional, ensure java's floating point crap is handled corectly
				if( !xAxisReversed )
				{
					svg.append( "<text id='yaxislabels" + (k++) + "' x='" + (x0 - YLABELSSPACER_X) + "' y='" + (y1 + YLABELSPACER_Y) +
							            "' style='text-anchor: end;' direction='rtl' alignment-baseline='text-after-edge' " + labelfontSVG + ">" + CellFormatFactory
							.fromPatternString( yPattern )
							.format( bd ) + "</text>\r\n" );
				}
				else
				{
					svg.append( "<text id='yaxislabels" + (k++) + "' x='" + (x0 + YLABELSSPACER_X) + "' y='" + (y1 + YLABELSPACER_Y) +
							            "' style='text-anchor: start;' alignment-baseline='text-after-edge' " + labelfontSVG + ">" + CellFormatFactory
							.fromPatternString( yPattern )
							.format( bd ) + "</text>\r\n" );
				}
				y0 -= inc * rfY;
			}
			svg.append( "</g>\r\n" );
			// AXIS bounding line
			x0 = x + (!xAxisReversed ? 0 : w);
			y0 = y;
			svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + (y0 + h) + "'/>\r\n" );
			// Y AXIS TITLE
			if( !xAxisReversed )
			{
				x0 = /*YAXISTITLEOFFSET + */10;    // TODO: 10 should be actual font height (when rotated 90 as is the norm)
			}
			else
			{
				x0 = x + w + YAXISTITLEOFFSET;
			}
			svg.append( getAxisTitleSVG( x0, (y + h / 2), titlefontSVG, titleRot, "yaxistitle" ) );
		}
		return svg.toString();
	}

	/**
	 * return the svg to display an XYValue axis (bubble, scatter)
	 *
	 * @param ca
	 * @param titlefontSVG
	 * @param labelfontSVG
	 * @param showMinorTickMarks
	 * @param showMajorTickMarks
	 * @param chartMetrics
	 * @return
	 */
	private String drawXYValAxisSVG( ChartAxes ca,
	                                 String titlefontSVG,
	                                 String labelfontSVG,
	                                 boolean showMinorTickMarks,
	                                 boolean showMajorTickMarks,
	                                 java.util.Map<String, Double> chartMetrics,
	                                 Object[] categories )
	{
		StringBuffer svg = new StringBuffer();
		double x0, x1, y0, y1;
		double minorinc = 0;
		double inc;
		// Y or Value Axis- must be non-textual; values obtained in calling method
		// major and minor tick marks, max and min axis scales NOTE: for MOST category (usually x) axes, they are textual; the max=# of categories; min=0
		double x = chartMetrics.get( "x" );
		double y = chartMetrics.get( "y" );
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double max = chartMetrics.get( "max" );
		double min = chartMetrics.get( "min" );
		double minor = (Double) ca.getMetric( "minor" );
		double major = (Double) ca.getMetric( "major" );
		boolean scaleIsInteger = (major == Math.floor( major ));    // true if scale is in integer form rather than double (used for formatting label text)
		String yPattern = (String) ca.getMetric( "yPattern" );
		double XAXISLABELOFFSET = (Double) ca.getMetric( "XAXISLABELOFFSET" );
		double XAXISTITLEOFFSET = (Double) ca.getMetric( "XAXISTITLEOFFSET" );

		if( categories != null && categories.length > 0 )
		{ // shouldn't
			double xmin = Double.MAX_VALUE, xmax = Double.MIN_VALUE;
			boolean TEXTUALXAXIS = true;
			for( int j = 0; j < categories.length; j++ )
			{
				try
				{
					double d = new Double( categories[j].toString() );
					xmax = Math.max( xmax, d );
					xmin = Math.min( xmin, d );
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
				minor = (int) d[0];
				major = (int) d[1];
				xmax = d[2];

			}
			else
			{    // NO category values are numeric --  x axis is Textual - just use count
				major = 1;
				minor = 0;
				xmax = categories.length + 1;
				XAXISLABELOFFSET = 30;    // TODO: calculate	
			}

			// CATEGORY LABELS- centered within area on X Axis
			y0 = y + h + XAXISLABELOFFSET;
			inc = w / (xmax / major);
			x0 = x;
			int scale = 0;
			{ // figure out if scale has a fractional portion; if so, determine desired scale and keep to that  
				String s = String.valueOf( major );
				int z = s.indexOf( "." );
				if( z != -1 )
				{
					scale = s.length() - (z + 1);
				}
			}
			int k = 0;    // axis label index
			for( double i = 0; i <= xmax; i += major )
			{
				if( !TEXTUALXAXIS )
				{
					java.math.BigDecimal bd = new java.math.BigDecimal( i ).setScale( scale, java.math.BigDecimal.ROUND_HALF_UP );
					svg.append( "<text id='xaxislabels" + (k++) + "' x='" + x0 + "' y='" + y0 +											/* TODO: should really trap xyPattern */
							            "' style='text-anchor: middle;' " + labelfontSVG + ">" + CellFormatFactory.fromPatternString(
							yPattern ).format( bd ) + "</text>\r\n" );
				}
				else
				{
					svg.append( "<text id='xaxislabels" + (k++) + "' x='" + x0 + "' y='" + y0 +
							            "' style='text-anchor: middle;' " + labelfontSVG + ">" + CellFormatFactory.fromPatternString(
							yPattern ).format( i ) + "</text>\r\n" );
				}
				x0 += inc;
			}

			// TICK MARKS
			y0 = h + y;    // origin at y+h
			x0 = x;        // start at chart x
			svg.append( "<g>\r\n" );
			if( minor > 0 )
			{
				minorinc = inc / minor;
			}
			for( int i = 0; i <= xmax; i += major )
			{
				y1 = y0 + 2;    // minor tick mark
				for( int j = 0; j < minor; j++ )
				{
					svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n" );
					x0 += minorinc;
				}
				y1 = y0 + 5;    // Major tick marks
				svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n" );
				if( minorinc == 0 )
				{
					x0 += inc;
				}
			}
			svg.append( "</g>\r\n" );
		}
		// X AXIS TITLE
		int titleRot = (linkedtd != null) ? linkedtd.getRotation() : 0;
		svg.append( getAxisTitleSVG( (x + w / 2), (y + h + XAXISTITLEOFFSET), titlefontSVG, titleRot, "zaxistitle" ) );
		return svg.toString();
	}

	/**
	 * Bar chart Axes are switched so handle seperately
	 * basically x axis holds Y labels and title, and visa versa + gridlines go up and down (traverse y) rather than across (traverse x)
	 *
	 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @return String SVG
	 */
	private String getSVGBARCHART( ChartAxes ca,
	                               String titlefontSVG,
	                               String labelFontSVG,
	                               boolean showMinorTicks,
	                               boolean showMajorTicks,
	                               java.util.Map<String, Double> chartMetrics,
	                               Object[] categories )
	{
		StringBuffer svg = new StringBuffer();

		// major and minor tick marks, max and min axis scales NOTE: for MOST category (usually x) axes, they are textual; the max=# of categories; min=0
		// X Axis/Cats in reversed order means X axis on Top, Y axis labels in reversed order (along with bars)
		double x0, x1, y0, y1;
		double inc;
		int rfX = 1;    // reverse factor used to reverse order + position of axes
		int rfY = 1;    // ""
		boolean scaleIsInteger = true;    // usual case
		double x = chartMetrics.get( "x" );
		double y = chartMetrics.get( "y" );
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double canvasw = chartMetrics.get( "canvasw" );
		double canvash = chartMetrics.get( "canvash" );
		double max = chartMetrics.get( "max" );
		double min = chartMetrics.get( "min" );
		boolean xAxisReversed = (Boolean) ca.getMetric( "xAxisReversed" );
		boolean yAxisReversed = (Boolean) ca.getMetric( "yAxisReversed" );
		double YAXISTITLEOFFSET = (Double) ca.getMetric( "YAXISTITLEOFFSET" );
		String yPattern = (String) ca.getMetric( "yPattern" );
		double XAXISLABELOFFSET = (Double) ca.getMetric( "XAXISLABELOFFSET" );
		double XAXISTITLEOFFSET = (Double) ca.getMetric( "XAXISTITLEOFFSET" );
		double major = (Double) ca.getMetric( "major" );

		if( wType == XAXIS )
		{
			scaleIsInteger = (major == Math.floor( major ));    // true if scale is in integer form rather than double (used for formatting label text)
		}
		switch( wType )
		{
			case XAXIS:
				// X Axis TICKS x= rectX, x2= x1, y1= recty+h, y2=y1+5 (major) x1 increments by (#Major Ticks-1)/Width (usually 4)
				//              x= rectX+17, x2= x1, y1= recty+h y2=y1+2 (minor) x1 increments by (approx 8)
				x0 = x;    // start at chart x 
				y0 = y + (!xAxisReversed ? h : 0);    // origin at y+h (unless reversed)
				rfX = (!xAxisReversed ? 1 : -1);
				rfY = (!yAxisReversed ? 1 : -1);
				if( major > 0 )
				{
					svg.append( "<g>\r\n" );
					inc = w / ((max - min) / major);    // w/scale factor
					int scale = 0;
					{ // figure out if scale has a fractional portion; if so, determine desired scale and keep to that  
						String s = String.valueOf( major );
						int z = s.indexOf( "." );
						if( z != -1 )
						{
							scale = s.length() - (z + 1);
						}
					}
					int k = 0;    // axis label index
					for( double i = min; i <= max; i += major )
					{// traverse across bottom (or top, if reversed) of x axis) incrementing x value keeping y value constant		
						y1 = y0 + 5 * rfX;    // Major tick marks
						if( showMajorTicks )
						{
							svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "'" + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + y1 + "'/>\r\n" );
						}
						// X axis labels= (Values)
						java.math.BigDecimal bd = new java.math.BigDecimal( i ).setScale( scale, java.math.BigDecimal.ROUND_HALF_UP );
						if( !xAxisReversed )
						{
							svg.append( "<text id='xaxislabels" + (k++) + "' x='" + (x0 + (!yAxisReversed ? 0 : w)) + "' y='" + (y0 + XAXISLABELOFFSET) + "' style='text-anchor: end;' alignment-baseline='middle' " + labelFontSVG + ">" + CellFormatFactory
									.fromPatternString( yPattern )
									.format( bd ) + "</text>\r\n" );
						}
						else
						{
							svg.append( "<text id='xaxislabels" + (k++) + "' x='" + (x0) + "' y='" + (y0 - 4) + "' style='text-anchor: end;' " + labelFontSVG + ">" + CellFormatFactory
									.fromPatternString( yPattern )
									.format( bd ) + "</text>\r\n" );
						}
						x0 += inc * rfY;
					}
					svg.append( "</g>\r\n" );
				}
				// AXIS bounding line
				y0 = y + (!xAxisReversed ? h : 0);        // origin at y+h (unless reversed)
				x0 = x; // + (!yAxisReversed?0:ci.w);		// start at chart x
				svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y0 + "' x2='" + (x0 + w) + "' y2='" + y0 + "'/>\r\n" );
				// is the Y AXIS TITLE i.e. goes alongside Y AXIS
				if( !yAxisReversed )
				{
					x0 = YAXISTITLEOFFSET;
				}
				else
				{
					x0 = x + w + YAXISTITLEOFFSET;
				}
				svg.append( getAxisTitleSVG( x0, (y + h / 2), titlefontSVG, 90, "xaxistitle" ) );
				break;
			case YAXIS:
				rfX = (!xAxisReversed ? 1 : -1);    // if reversed, y vals on top (x axis), cats are in reversed order (y axis) 
				rfY = (!yAxisReversed ? 1 : -1);    // if reversed, y vals in reverse order (x axis), cats are on RHS (y axis)
				if( categories != null && categories.length > 0 )
				{ // should! 
					inc = h / categories.length;
					x0 = x;    // + (!yAxisReversed?0:ci.w);	// draw y axis tick marks + y axis labels (= categories) 
					y0 = y + (!xAxisReversed ? h : 0);    // starts at bottom, goes up to top unless reversed
					int k = 0;    // axis label index
					for( int i = 0; i < categories.length; i++ )
					{    // traverse Y axis, spacing category labels (x is constant, y is segmented)
						x1 = x0 - 5 * rfY;// Major tick Marks
						if( showMajorTicks )
						{
							svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + ChartType.getStrokeSVG() + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x1 + "' y2='" + y0 + "'/>\r\n" );
						}
						// Y Axis Labels = categories
						y1 = y0 - inc * rfX + inc * .5 * rfX;
						if( !yAxisReversed )
						{
							svg.append( "<text id='yaxislabels" + (k++) + "' x='" + (x0 - 8) + "' y='" + y1 +
									            "' style='text-anchor: end;' direction='rtl' dominant-baseline='text-before-edge' " + labelFontSVG + ">" + categories[i]
									.toString() + "</text>\r\n" );
						}
						else
						{
							svg.append( "<text id='yaxislabels" + (k++) + "' x='" + (x0 + w + 8) + "' y='" + (y1 + 4) +
									            "' style='text-anchor: start;' alignment-baseline='text-after-edge' " + labelFontSVG + ">" + categories[i]
									.toString() + "</text>\r\n" );
						}
						y0 -= inc * rfX;
					}
					// show gridlines (top to bottom)
					String lineSVG = getLineSVG( AxisLineFormat.ID_MAJOR_GRID );
					if( !lineSVG.equals( "" ) )
					{
						y0 = y;
						x0 = x;    // start at chart x 
						inc = w / ((max - min) / major);    // w/scale factor
						// 1st line is axis line
						svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + getLineSVG( AxisLineFormat.ID_AXIS_LINE ) + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + (y0 + h) + "'/>\r\n" );
						// rest are grid lines
						for( double i = min; i < max; i += major )
						{
							x0 += inc;
							svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + lineSVG + " x1='" + x0 + "' y1='" + y0 + "' x2='" + x0 + "' y2='" + (y0 + h) + "'/>\r\n" );
						}
					}
					// is the X AXIS TITLE i.e. GOES ON THE X AXIS
					y0 = (!xAxisReversed ? canvash - XAXISTITLEOFFSET : y - XAXISTITLEOFFSET);
					int titleRot = (linkedtd != null) ? linkedtd.getRotation() : 0;
					svg.append( getAxisTitleSVG( (x + w / 2), y0, titlefontSVG, titleRot, "yaxistitle" ) );
				}
				break;
		}
		return svg.toString();
	}

	/**
	 * returns the SVG representation of this Radar Chart Axes --
	 * <br>Radar Chart Axis look like a spider web
	 * <br>Note that these coordinates, like all axis scales, match the coordinates and calculations in Rader.getSVG
	 *
	 * @param chartMetrics maps chart coords in pixels x, y, w, h, canvasw, canvash, min, max
	 * @return String SVG
	 */
	private String getSVGRADARCHART( ChartAxes ca,
	                                 String titleFontSVG,
	                                 String labelFontSVG,
	                                 java.util.Map<String, Double> chartMetrics,
	                                 Object[] categories )
	{
		double x = chartMetrics.get( "x" );
		double y = chartMetrics.get( "y" );
		double w = chartMetrics.get( "w" );
		double h = chartMetrics.get( "h" );
		double max = chartMetrics.get( "max" );
		double min = chartMetrics.get( "min" );
		double major = (Double) ca.getMetric( "major" );
		StringBuffer svg = new StringBuffer();
		if( wType == YAXIS )
		{
			svg.append( "<g>\r\n" );
			if( categories != null && categories.length > 0 )
			{ // shouldn't
				major = major * 2;    // appears the scale calc doesn't follow other charts ...	
				double n = categories.length;
				double centerx = w / 2 + x;
				double centery = h / 2 + y;
				double percentage = 1 / n;        // divide into equal sections
				double radius = Math.min( w, h ) / 2.3;    // should take up almost entire w/h of chart
				double radiusinc = radius / (max / major);
				double lastx = centerx, lasty = centery - radius;    // again, start straight up
				int k = 0;    // axis label index
				for( double j = min; j <= max; j += major )
				{    // each major unit is a concentric line				
					double angle = 90;            // starts straight up
					for( int i = 0; i <= n; i++ )
					{        // each category is a radial line; <= n so can complete the spider web line path
						// get next point on circumference 
						double x1 = centerx + radius * (Math.cos( Math.toRadians( angle ) ));
						double y1 = centery - radius * (Math.sin( Math.toRadians( angle ) ));
						svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "'" + ChartType.getStrokeSVG() + " x1='" + lastx + "' y1='" + lasty + "' x2='" + x1 + "' y2='" + y1 + "'/>\r\n" );
						if( j == 0 && i < n )
						{    // print radial lines & labels at very end of chart/top of each line  -- only do for 1st round
							svg.append( "<line fill='none' fill-opacity='" + ChartType.getFillOpacity() + "' " + ChartType.getStrokeSVG() + " x1='" + centerx + "' y1='" + centery + "' x2='" + x1 + "' y2='" + y1 + "'/>\r\n" );
							double labelx1 = centerx + (radius + 10) * (Math.cos( Math.toRadians( angle ) ));
							double labely1 = centery - (radius + 10) * (Math.sin( Math.toRadians( angle ) ));
							svg.append( "<text id='xaxislabels" + (k++) + "' x='" + labelx1 + "' y='" + labely1 + "' style='text-anchor: middle;' " + labelFontSVG + ">" + categories[i]
									.toString() + "</text>\r\n" );
						}
						// next angle
						angle -= (percentage * 360);
						lastx = x1;
						lasty = y1;
					}
					radius -= radiusinc;
				}
			}
			svg.append( "</g>\r\n" );
		}
		return svg.toString();
	}

	/**
	 * return the svg necessary to display categories along an x axis
	 *
	 * @param x
	 * @param y
	 * @param w
	 * @param h
	 * @param inc
	 * @param labelRot
	 * @param categories
	 * @param f
	 * @param labelfontSVG
	 * @param yAxisReversed
	 * @param xAxisReversed
	 * @param XAXISLABELOFFSET
	 * @return
	 */
	private String getCategoriesSVG( double x,
	                                 double y,
	                                 double w,
	                                 double h,
	                                 double inc,
	                                 int labelRot,
	                                 Object[] categories,
	                                 java.awt.Font f,
	                                 String labelfontSVG,
	                                 boolean yAxisReversed,
	                                 boolean xAxisReversed,
	                                 double XAXISLABELOFFSET )
	{
		// Category Labels - centered within area on X Axis
		StringBuffer svg = new StringBuffer();
		double x0, x1, y0, y1;
		int k = labelfontSVG.indexOf( "font-size=" ) + 11;
		double fh = Double.parseDouble( labelfontSVG.substring( k,
		                                                        labelfontSVG.indexOf( "pt" ) ) );    // approximate height of a line of labels
		y0 = y + (!yAxisReversed ? h + XAXISLABELOFFSET / 3 : -XAXISLABELOFFSET);    // draw on bottom edge of axis unless Y axis is reversed
		int m = 0;    // axis label index
		for( int i = 0; i < categories.length; i++ )
		{
			if( !xAxisReversed )
			{
				x0 = x + inc * i + (inc / 2);
			}
			else    // reversed:  category labels start from LHS and go to RHS 
			{
				x0 = x + w - inc * i - (inc / 2);
			}
			if( labelRot != 0 )
			{
				double len = StringTool.getApproximateStringWidthLB( f,
				                                                     CellFormatFactory.fromPatternString( null ).format( categories[i] ) );
				if( labelRot == 45 )
				{
					len = (int) Math.ceil( len * (Math.cos( Math.toRadians( labelRot ) )) );
				}
				int offset = (int) (len / 2) + 5;
				y0 = y + (!yAxisReversed ? h + offset : -offset);
				if( labelRot == 45 )
				{
					x0 += inc / 2;
				}

			}
			// handle multiple lines in X axis labels - must do "by hand"
			String[] s = categories[i].toString().split( "\n" );
			svg.append( "<text id='xaxislabels" + (m++) + "' x='" + x0 + "' y='" + y0 +
					            (labelRot == 0 ? "" : "' transform='rotate(" + labelRot + ", " + (x0) + " , " + (y0) + ")") +
					            "' style='text-anchor: middle;' alignment-baseline='text-after-edge' " + labelfontSVG + ">" );
			for( int z = 0; z < s.length; z++ )
			{
				svg.append( "<tspan x='" + (x0) + "' dy='" + (fh * 1.4) + "'>" + s[z] + "</tspan>\r\n" );
			}
			svg.append( "</text>\r\n" );
		}
		return svg.toString();
	}

	/**
	 * return the svg necessary to display an axis title
	 *
	 * @param x
	 * @param y
	 * @param titlefontSVG
	 * @param titleRot     0, 45 or -90 degrees
	 * @param scriptTitle  either "xaxistitle", "yaxistitle" or "zaxistitle"
	 * @return
	 */
	private String getAxisTitleSVG( double x, double y, String titlefontSVG, int titleRot, String scriptTitle )
	{
		StringBuffer svg = new StringBuffer();
		svg.append( "<g>\r\n" );
		svg.append( "<text " + getScript( scriptTitle ) + " x='" + x + "' y='" + y +
				            (titleRot == 0 ? "" : "' transform='rotate(-" + titleRot + ", " + x + " ," + y + ")") +
				            "' style='text-anchor: middle;' " + titlefontSVG + ">" + this.getTitle() + "</text>\r\n" );
		svg.append( "</g>\r\n" );
		return svg.toString();
	}

}

class Scaling implements OOXMLElement
{
	private String logBase, max, min, orientation;

	public Scaling()
	{    // no-param constructor, set up common defaults 
	}

	public Scaling( String logBase, String max, String min, String orientation )
	{
		this.logBase = logBase;
		this.max = max;
		this.min = min;
		this.orientation = orientation;
	}

	public Scaling( Scaling sc )
	{
		this.logBase = sc.logBase;
		this.max = sc.max;
		this.min = sc.min;
		this.orientation = sc.orientation;
	}

	/**
	 * parse Axis OOXML element (catAx, valAx, serAx or dateAx)
	 *
	 * @param xpp     XmlPullParser
	 * @param lastTag element stack
	 * @return spPr object
	 */
	public static OOXMLElement parseOOXML( XmlPullParser xpp, Stack lastTag )
	{
		String logBase = null, max = null, min = null, orientation = null;
		try
		{
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "scaling" ) )
					{
					}
					else if( tnm.equals( "logBase" ) )
					{
						logBase = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "max" ) )
					{
						max = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "min" ) )
					{
						min = xpp.getAttributeValue( 0 );
					}
					else if( tnm.equals( "orientation" ) )
					{
						orientation = xpp.getAttributeValue( 0 );
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					String endTag = xpp.getName();
					if( endTag.equals( "scaling" ) )
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
			Logger.logErr( "scaling.parseOOXML: " + e.toString() );
		}
		Scaling s = new Scaling( logBase, max, min, orientation );
		return s;
	}

	/**
	 * return the specific scaling option
	 *
	 * @param op
	 * @return
	 */
	public String getOption( String op )
	{
		if( op.equals( "logBase" ) )
		{
			return logBase;
		}
		else if( op.equals( "max" ) )
		{
			return max;
		}
		else if( op.equals( "min" ) )
		{
			return min;
		}
		else if( op.equals( "orientation" ) )
		{
			return orientation;
		}
		return null;
	}

	@Override
	public String getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		// sequence:  logBase, orientation, max, min
		ooxml.append( "<c:scaling" );
		if( logBase != null )
		{
			ooxml.append( "<c:logBase val=\"" + logBase + "\"/>" );
		}
		if( orientation != null )
		{
			ooxml.append( "<c:orientation val=\"" + orientation + "\"/>" );
		}
		if( max != null )
		{
			ooxml.append( "<c:max val=\"" + max + "\"/>" );
		}
		if( min != null )
		{
			ooxml.append( "<c:min val=\"" + min + "\"/>" );
		}
		ooxml.append( "</scaling>\r\n" );
		return ooxml.toString();
	}

	@Override
	public OOXMLElement cloneElement()
	{
		return new Scaling( this );
	}
}