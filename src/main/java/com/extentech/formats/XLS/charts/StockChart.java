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

import com.extentech.formats.XLS.WorkBook;

public class StockChart extends LineChart
{

	public StockChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		defaultShape = STOCKCHART;
	}

	/**
	 * gets the chart-type specific ooxml representation: <stockChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:stockChart>" );
		cooxml.append( "\r\n" );

		// *** Series Data:	ser, cat, val for most chart types
		cooxml.append( this.getParentChart().getChartSeries().getOOXML( this.getChartType(), false, 0 ) );

		// chart data labels, if any
		//TODO: FINISH dlbls
		//cooxml.append(getDataLabelsOOXML(cf));

		//dropLines
		ChartLine cl = this.cf.getChartLinesRec( ChartLine.TYPE_DROPLINE );
		if( cl != null )
		{
			cooxml.append( cl.getOOXML() );
		}
		// hiLowLines
		cl = this.cf.getChartLinesRec( ChartLine.TYPE_HILOWLINE );
		if( cl != null )
		{
			cooxml.append( cl.getOOXML() );
		}
		// upDownBars
		cooxml.append( cf.getUpDownBarOOXML() );

		// axis ids	 - unsigned int strings
		cooxml.append( "<c:axId val=\"" + catAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:axId val=\"" + valAxisId + "\"/>" );
		cooxml.append( "\r\n" );

		cooxml.append( "</c:stockChart>" );
		cooxml.append( "\r\n" );

		return cooxml;
	}
}
