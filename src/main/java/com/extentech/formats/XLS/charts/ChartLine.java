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
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;
import org.xmlpull.v1.XmlPullParser;

import java.util.Stack;

/**
 * <b>ChartLine: Drop/Hi-Lo/Series Lines on a Line Chart (0x101c)</b>
 * <p/>
 * The CrtLine record specifies the presence of drop lines, high-low lines, series lines or leader lines on the chart group.
 * This record is followed by a LineFormat record which specifies the format of the lines.
 * <p/>
 * <br> id (2 bytes): An unsigned integer that specifies the type of line that is present on the chart group.
 * This field value MUST be unique among the other id field values in CrtLine records in the current chart group.
 * This field MUST be greater than the id field values in preceding CrtLine records in the current chart group. MUST be a value from the following table:
 * <p/>
 * Value		Type of Line
 * 0x0000		Drop lines below the data points of line, area, and stock chart groups.
 * 0x0001		High-low lines around the data points of line and stock chart groups.
 * 0x0002		Series lines connecting data points of stacked column and bar chart groups, and the primary pie to the secondary bar/pie of bar of pie and pie of pie chart groups.
 * 0x0003		Leader lines with non-default formatting connecting data labels to the data point of pie and pie of pie chart groups.
 * <p/>
 * But also there is:
 * DROPBAR = DropBar Begin LineFormat AreaFormat [GELFRAME] [SHAPEPROPS] End
 */
public class ChartLine extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -8311605814020380069L;
	private short id;
	private ChartFormat cf;    // necessary to link to corresponding ChartFormat
	public static byte TYPE_DROPLINE = 0;
	public static byte TYPE_HILOWLINE = 1;
	public static byte TYPE_SERIESLINE = 2;
	public static byte TYPE_LEADERLINE = 3;

	public void init()
	{
		super.init();
		id = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
	}

	public static XLSRecord getPrototype()
	{
		ChartLine cl = new ChartLine();
		cl.setOpcode( CHARTLINE );
		cl.setData( new byte[]{ 0, 0 } );
		cl.init();
		return cl;
	}

	/**
	 * <br>0= drop lines below the data points of line, area and stock charts
	 * <br>1= High-low lines around the data points of line and stock charts
	 * <br>2- Series Line connecting data points of stacked column and bar charts + some pie chart configurations
	 * <br>3= Leader lines with non-default formatting for pie and pie of pie
	 *
	 * @return
	 */
	public int getLineType()
	{
		return id;
	}

	/**
	 * sets the chart line type:
	 * <br>0= drop lines below the data points of line, area and stock charts
	 * <br>1= High-low lines around the data points of line and stock charts
	 * <br>2- Series Line connecting data points of stacked column and bar charts + some pie chart configurations
	 * <br>3= Leader lines with non-default formatting for pie and pie of pie
	 */
	public void setLineType( int id )
	{
		this.id = (short) id;
		byte[] b = ByteTools.shortToLEBytes( this.id );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
	}

	/**
	 * return the LineFormat rec associated with this ChartLine
	 * (== the next record in the cf chart array)
	 *
	 * @return
	 */
	private LineFormat findLineFormatRec()
	{
		LineFormat lf = null;
		for( int i = 0; i < cf.chartArr.size(); i++ )
		{
			if( cf.chartArr.get( i ).equals( this ) )
			{
				lf = (LineFormat) cf.chartArr.get( i + 1 );
				break;
			}
		}
		return lf;
	}

	/**
	 * parse a Chart Line OOXML element: either
	 * <li>dropLines
	 * <li>hiLowLines
	 * <li>serLines
	 *
	 * @param xpp
	 * @param lastTag
	 * @param cf
	 */
	public void parseOOXML( XmlPullParser xpp, Stack<String> lastTag, ChartFormat cf, WorkBookHandle bk )
	{
		this.cf = cf;
		String endTag = (String) lastTag.peek();
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
						LineFormat lf = findLineFormatRec();
						if( lf != null )
						{
							lf.setFromOOXML( sppr );
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
					if( xpp.getName().equals( endTag ) )
					{
						lastTag.pop();
					}
					break;
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			Logger.logErr( "ChartLine.parseOOXML: " + e.toString() );
		}
	}

	/**
	 * return the OOXML to define this ChartLine
	 *
	 * @return
	 */
	public StringBuffer getOOXML()
	{
		StringBuffer cooxml = new StringBuffer();
		String tag = null;
		if( id == TYPE_DROPLINE )
		{
			tag = "c:dropLines>";
		}
		else if( id == TYPE_HILOWLINE )
		{
			tag = "c:hiLowLines>";
		}
		else if( id == TYPE_SERIESLINE )
		{
			tag = "c:serLines>";
		}
		cooxml.append( "<" + tag );
		LineFormat lf = findLineFormatRec();
		if( lf != null )
		{
			cooxml.append( lf.getOOXML() );
		}
		cooxml.append( "</" + tag );
		return cooxml;
	}
}
