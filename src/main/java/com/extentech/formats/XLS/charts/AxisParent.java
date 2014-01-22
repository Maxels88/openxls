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

import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;

/**
 * <b>AxisParent: Axis Size and Location (0x1014)</b>
 * <p/>
 * This record specifies properties of an axis group and the beginning of a collection of records as defined by the chart sheet substream.
 * <p/>
 * <p/>
 * The Axis parent record stores most of the actual chart information,  what type of chart, x and y labels, etc.
 * <p/>
 * 4	iax	2		axis index (0= main, 1= secondary)  This field MUST be set to zero when it is in the first AxisParent record in the chart sheet substream,
 * This field MUST be set to 1 when it is in the second AxisParent record in the chart sheet substream.
 * 16  	unused (16 bytes): Undefined and MUST be ignored.
 * <p/>
 * <p/>
 * this doesn't appear correct given ms doc
 * 6	x	4
 * 10	y	4
 * 14	dx	4		len of x axis
 * 18	dy	4		len of y axis
 */
public class AxisParent extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2247258217367570732L;
	private short iax = 0;
	private int x = 0, y = 0, dx = 0, dy = 0;

	@Override
	public void init()
	{
		super.init();
		iax = ByteTools.readShort( this.getData()[0], this.getData()[1] );
		x = ByteTools.readInt( this.getBytesAt( 2, 4 ) );
		y = ByteTools.readInt( this.getBytesAt( 6, 4 ) );
		dx = ByteTools.readInt( this.getBytesAt( 10, 4 ) );
		dy = ByteTools.readInt( this.getBytesAt( 14, 4 ) );

	}

	/**
	 * default version, returns the 1sst chart format record
	 * <br>if there are more than 1 chart type in the chart
	 * there will be multiple chart format records
	 *
	 * @return
	 */
	protected ChartFormat getChartFormat()
	{
		return getChartFormat( 0, false );
	}

	/**
	 * get the chart format collection
	 *
	 * @return
	 */
	protected ChartFormat getChartFormat( int nChart, boolean addNew )
	{
		for( XLSRecord aChartArr : chartArr )
		{
			BiffRec b = aChartArr;
			if( b.getOpcode() == CHARTFORMAT )
			{
				ChartFormat cf = (ChartFormat) b;
				if( cf.getDrawingOrder() == nChart )
				{
					return cf;
				}
			}
		}
		if( addNew )
		{
			return createNewChart( nChart );
		}
		return null;
	}

	/**
	 * create basic cf necessary to define multiple or overlay charts
	 * and add to this AxisParent
	 *
	 * @param nChart overlay chart number (>0 and <=9)
	 * @return new ChartFormat
	 */
	protected ChartFormat createNewChart( int nChart )
	{
		// muliple charts, create one
		if( (nChart > 0) && (nChart <= 9) )
		{
			ChartFormat cf = (ChartFormat) ChartFormat.getPrototype();
			cf.setParentChart( this.getParentChart() );
			Bar b = (Bar) Bar.getPrototype();
			cf.chartArr.add( b );    // add a dummy chart object - will be replaced later
			ChartFormatLink cfl = (ChartFormatLink) ChartFormatLink.getPrototype();
			cf.chartArr.add( cfl );
			SeriesList sl = new SeriesList();
			sl.setOpcode( SERIESLIST );
			sl.setParentChart( this.getParentChart() );
			cf.chartArr.add( sl );
			cf.setDrawingOrder( nChart );
			this.chartArr.add( cf );    // add chartformat to chart array of axis parent
			return cf;
		}
		return null;
	}

	/**
	 * remove axis records for pie-type charts ...
	 */
	public void removeAxes()
	{
		if( chartArr.get( 1 ) instanceof Axis )
		{
			chartArr.remove( 1 );    // remove Axis
		}
		if( chartArr.get( 1 ) instanceof Axis )
		{
			chartArr.remove( 1 );    // remove Axis
		}
		if( chartArr.get( 1 ) instanceof TextDisp )
		{
			chartArr.remove( 1 );    // remove Text for Axis
		}
		if( chartArr.get( 1 ) instanceof TextDisp )
		{
			chartArr.remove( 1 );    // remove Text for Axis
		}
		if( chartArr.get( 1 ) instanceof PlotArea )
		{
			chartArr.remove( 1 );    // remove
		}
		if( chartArr.get( 1 ) instanceof Frame )
		{
			chartArr.remove( 1 );    // remove Frame
		}
		// all should be left is pos + chartFormat
	}

	/**
	 * remove the desired axis + associated records
	 *
	 * @param axisType
	 */
	public void removeAxis( int axisType )
	{
		// Remove axis
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = chartArr.get( i );
			if( b.getOpcode() == AXIS )
			{
				if( ((Axis) b).getAxis() == axisType )
				{
					chartArr.remove( i );
					break;
				}
			}
		}
		int tdType = TextDisp.convertType( axisType );
		// Remove TextDisp assoc with Axis
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = chartArr.get( i );
			if( b.getOpcode() == TEXTDISP )
			{
				TextDisp td = (TextDisp) b;
				if( tdType == td.getType() )
				{
					chartArr.remove( i );
					break;
				}
			}
		}
	}

	/**
	 * for those charts such as Radar, remove the
	 * axis frame and plot area records as are not necessary
	 */
	protected void removePlotArea()
	{
		boolean remove = false;
		for( int i = 0; i < chartArr.size(); i++ )
		{
			BiffRec b = chartArr.get( i );
			if( b.getOpcode() == AXIS )
			{
				remove = true;
			}
			else if( b.getOpcode() == CHARTFORMAT )
			{
				break;
			}
			else if( remove )
			{
				chartArr.remove( i );
				i--;
			}
		}

	}

	/**
	 * return XML of axis label and specific options for the desired axis
	 * (both the Axis record + it's "associated" TextDisp record are consulted)
	 *
	 * @param Axis int desired Axis
	 * @return String XML of axis label + specific options
	 * @see ObjectLink
	 */ // KSC: TODO: Get all axis ops ...
	protected String getAxisOptionsXML( int axis )
	{
		boolean bHasAxis = false;
		boolean bHasLabel = false;
		StringBuffer sb = new StringBuffer();
		int tdType = TextDisp.convertType( axis );
		for( XLSRecord aChartArr : chartArr )
		{
			BiffRec b = aChartArr;
			if( b.getOpcode() == AXIS )
			{
				if( ((Axis) b).getAxis() == axis )
				{
					sb.append( ((Axis) b).getOptionsXML() );
					bHasAxis = true;
				}
			}
			else if( b.getOpcode() == TEXTDISP )
			{
				if( ((TextDisp) b).getType() == tdType )
				{
					sb.append( ((TextDisp) b).getOptionsXML() );
					bHasLabel = true;
					break;
				}
			}
		}
		if( bHasAxis && !bHasLabel )    // must include axis even if no textdisp/label
		{
			sb.append( " Label=\"\"" );
		}
		return sb.toString();
	}

	/**
	 * return the desired axis according to AxisType
	 * (will create the axis if it doesn't exist)
	 * Also finds and associates the axis rec with it's TextDisp
	 * so as to be able to set axis optins ...
	 * will create Axis + TextDisp if not present ...
	 *
	 * @param axisType
	 * @return Desired axis
	 */
	public Axis getAxis( int axisType )
	{
		return getAxis( axisType, true );
	}

	/**
	 * return the desired axis according to AxisType
	 * Also finds and associates the axis rec with it's TextDisp
	 * so as to be able to set axis optins ...
	 * will create Axis + TextDisp if not present ...
	 *
	 * @param axisType
	 * @param bCreateIfNecessary - if true, will create the axis if it doesn't exist
	 * @return Desired axis
	 */
	public Axis getAxis( int axisType, boolean bCreateIfNecessary )
	{
		int lastTd = 0;
		int lastAxis = 0;
		Axis a = null;
		TextDisp td = null;
		int tdType = TextDisp.convertType( axisType );
		for( int i = 0; (i < chartArr.size()) && (td == null); i++ )
		{
			BiffRec b = chartArr.get( i );
			if( b.getOpcode() == AXIS )
			{
				lastAxis = i;
				if( ((Axis) b).getAxis() == axisType )
				{
					a = ((Axis) b);
					a.setAP( this );    // ensure axis is linked to it's parent AxisParent 20090108 KSC:
					//if (!bCreateIfNecessary) return a;
				}
			}
			else if( b.getOpcode() == TEXTDISP )
			{
				lastTd = i;
				if( tdType == ((TextDisp) b).getType() )
				{
					td = (TextDisp) b;
					if( bCreateIfNecessary )
					{// 20080723 KSC: added guard - but when is clearing td text necessary????
						td.setText( "" );    // clear out axis legend
					}
				}
			}
		}
		if( (a == null) && bCreateIfNecessary )
		{
			// if didn't find axis, then add axis + text disp ...
			// first, add TD
			td = (TextDisp) TextDisp.getPrototype( axisType, "", this.wkbook );
			td.setParentChart( this.getParentChart() );
			this.chartArr.add( lastTd + 1, td );
			// next, add axis
			a = (Axis) Axis.getPrototype( axisType );
			a.setParentChart( this.getParentChart() );
			this.chartArr.add( lastAxis + 1, a );
			a.setAP( this );    // ensure axis is linked to it's parent AxisParent 20090108 KSC:
		}
		if( a != null )
		{    // associate this axis with the textdisp if any
			a.setTd( td );
		}
		return a;
	}

	/**
	 * returns the plot area background color
	 *
	 * @return plot area background color hex string
	 */
	public String getPlotAreaBgColor()
	{
		for( XLSRecord aChartArr : chartArr )
		{
			BiffRec b = aChartArr;
			if( b instanceof Frame )
			{
				return ((Frame) b).getBgColor();
			}
		}
		return null;    //FormatHandle.COLOR_WHITE;
	}

	/**
	 * sets the plot area background color
	 *
	 * @param bg color int
	 */
	public void setPlotAreaBgColor( int bg )
	{
		for( XLSRecord aChartArr : chartArr )
		{
			BiffRec b = aChartArr;
			if( b instanceof Frame )
			{
				((Frame) b).setBgColor( bg );
				break;
			}
		}
	}

	/**
	 * adds a border around the plot area with the desired line width and line color
	 *
	 * @param lw
	 * @param lclr
	 */
	public void setPlotAreaBorder( int lw, int lclr )
	{
		for( XLSRecord aChartArr : chartArr )
		{
			BiffRec b = aChartArr;
			if( b instanceof Frame )
			{
				((Frame) b).addBox( lw, lclr, -1 );
				break;
			}
		}
	}

	/**
	 * returns true if this is a secondary axis
	 *
	 * @return
	 */
	public boolean isSecondaryAxis()
	{
		return (iax == 1);
	}

	/**
	 * sets if this is a secondary axis
	 *
	 * @param b
	 */
	public void setIsSecondaryAxis( boolean b )
	{
		if( b )
		{
			iax = 1;
		}
		else
		{
			iax = 0;
		}
		this.getData()[0] = (byte) iax;
	}
}
