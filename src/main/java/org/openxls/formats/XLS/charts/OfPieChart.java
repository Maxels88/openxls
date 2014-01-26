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
package org.openxls.formats.XLS.charts;

import org.openxls.formats.XLS.WorkBook;

public class OfPieChart extends ChartType
{
	Boppop ofPie = null;

	public OfPieChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		ofPie = (Boppop) charttype;
	}

	/**
	 * gets the chart-type specific ooxml representation: <ofPieChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:ofPieChart>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:ofPieType val=\"" + (ofPie.isPieOfPie() ? "pie" : "bar") + "\"/>" );
		cooxml.append( "<c:varyColors val=\"1\"/>" );

		// *** Series Data:	ser, cat, val for most chart types
		cooxml.append( getParentChart().getChartSeries().getOOXML( getChartType(), false, 0 ) );

		// chart data labels, if any
		//TODO: FINISH		    	
		//cooxml.append(getDataLabelsOOXML(cf));
		// gapWidth
		if( ofPie.getpcGap() != 150 )
		{
			cooxml.append( "<c:gapWidth val=\"" + ofPie.getpcGap() + "\"/>" );
		}
		// splitType
		if( ofPie.getSplitType() != -1 )
		{
			cooxml.append( "<c:splitType val=\"" + ofPie.getSplitTypeOOXML() + "\"/>" );
		}
		// splitPos
		if( ofPie.getSplitPos() != 0 )
		{
			cooxml.append( "<c:splitPos val=\"" + ofPie.getSplitPos() + "\"/>" );
		}
		// custSplit TODO: FINISH
		// secondPieSize
		if( ofPie.getSecondPieSize() != 75 )
		{
			cooxml.append( "<c:secondPieSize val=\"" + ofPie.getSecondPieSize() + "\"/>" );
		}
		// serLines 
		ChartLine cl = cf.getChartLinesRec();
		if( cl != null )
		{
			cooxml.append( cl.getOOXML() );
		}
		cooxml.append( "</c:ofPieChart>" );
		cooxml.append( "\r\n" );

		return cooxml;
	}
}
