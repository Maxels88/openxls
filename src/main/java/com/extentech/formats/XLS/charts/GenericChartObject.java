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

import java.util.ArrayList;

public class GenericChartObject extends XLSRecord implements ChartObject, ChartConstants
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -1919254120575019160L;

	protected int chartType = -1;        // this will be >=0 when record defines the type of chart
	protected ArrayList<XLSRecord> chartArr = new ArrayList();
	protected Chart parentChart;

	public boolean setChartOption( String op, String val )
	{
		return false;
	}

	public boolean isStacked()
	{
		return false;
	}

	public boolean is100Percent()
	{
		return false;
	}

	public boolean hasShadow()
	{
		return false;
	}

	public void setIsStacked( boolean bIsStacked )
	{
		;
	}

	public void setIs100Percent( boolean bOn )
	{
		;
	}

	/**
	 * get chart option common to almost all chart types
	 *
	 * @param op
	 * @return
	 */
	public String getChartOption( String op )
	{
		if( op.equals( "Stacked" ) )
		{ // Area, Bar, Pie, Line
			return String.valueOf( isStacked() );
		}
		else if( op.equals( "Shadow" ) )
		{ // Pie, Area, Bar, Line, Radar, Scatter
			return String.valueOf( hasShadow() );
		}
		else if( op.equals( "PercentageDisplay" ) )
		{ // Area, Bar,Line
			return String.valueOf( is100Percent() );
		}
		return null;
	}

	/**
	 * @return String XML representation of this chart-type's options
	 */
	public String getOptionsXML()
	{
		return "";
	}

	/**
	 * adds html id and handlers for generic chart svg element
	 *
	 * @param title id of element
	 * @return String html
	 */
	public static String getScript( String id )
	{
		String ret = "";
		if( id != null )
		{
			ret = "id='" + id + "' ";
		}
		ret += "onmouseover='highLight(evt);' onclick='handleClick(evt);' onmouseout='restore(evt)'";
		return ret;
	}

	public ArrayList getChartRecords()
	{
		return chartArr;
	}

	public void addChartRecord( XLSRecord b )
	{
		chartArr.add( b );
	}

	public Chart getParentChart()
	{
		return parentChart;
	}

	public void setParentChart( Chart c )
	{
		parentChart = c;
	}

	/**
	 * Get the output array of records, including begin/end records and those of it's children.
	 */
	public ArrayList getRecordArray()
	{
		ArrayList outputArr = new ArrayList();
		outputArr.add( this );
		int nChart = -1;
		for( int i = 0; i < chartArr.size(); i++ )
		{
			if( i == 0 )
			{
				Begin b = (Begin) Begin.getPrototype();
				outputArr.add( b );
			}

			Object o = chartArr.get( i );
			if( o instanceof ChartObject )
			{
				ChartObject co = (ChartObject) o;
				outputArr.addAll( co.getRecordArray() );
			}
			else
			{
				BiffRec b = (BiffRec) o;
				outputArr.add( b );    // 20070712 KSC: missed some recs!
			}

			if( i == chartArr.size() - 1 )
			{
				End e = (End) End.getPrototype();
				outputArr.add( e );
			}
		}
		return outputArr;
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		for( int i = 0; i < chartArr.size(); i++ )
		{
			XLSRecord r = (XLSRecord) chartArr.get( i );
			r.close();
		}
		chartArr.clear();
		parentChart = null;
	}
}
