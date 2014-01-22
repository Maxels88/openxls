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

public class Bar3DChart extends BarChart
{

	public Bar3DChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		bar = (Bar) charttype;
	}

	/**
	 * gets the chart-type specific ooxml representation: <barChart>
	 *
	 * @return
	 */
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:bar3DChart>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:barDir val=\"bar\"/>" );
		cooxml.append( "<c:grouping val=\"" );

		if( this.is100PercentStacked() )
		{
			cooxml.append( "percentStacked" );
		}
		else if( this.isStacked() )
		{
			cooxml.append( "stacked" );
		}
		else if( cf.is3DClustered() )
		{
			cooxml.append( "clustered" );
		}
		else
		{
			cooxml.append( "standard" );
		}
		cooxml.append( "\"/>" );
		cooxml.append( "\r\n" );
		// vary colors???

		// *** Series Data:	ser, cat, val for most chart types
		cooxml.append( this.getParentChart().getChartSeries().getOOXML( this.getChartType(), false, 0 ) );

		// chart data labels, if any
		//TODO: FINISH		    	
		//cooxml.append(getDataLabelsOOXML(cf));

		if( !this.getChartOption( "Gap" ).equals( "150" ) )
		{
			cooxml.append( "<c:gapWidth val=\"" + this.getChartOption( "Gap" ) + "\"/>" );    // default= 0
		}
		int gapdepth = this.getGapDepth();
		if( gapdepth != 0 )
		{
			cooxml.append( "<c:gapDepth val=\"" + gapdepth + "\"/>" );
		}
		cooxml.append( "<c:shape val=\"" + this.getShape() + "\"/>" );

		// axis ids	 - unsigned int strings
		cooxml.append( "<c:axId val=\"" + catAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:axId val=\"" + valAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		if( this.getParentChart().getAxes().hasAxis( ZAXIS ) )
		{
			cooxml.append( "<c:axId val=\"" + serAxisId + "\"/>" );
			cooxml.append( "\r\n" );
		}
		else
		{// KSC: appears to be necessary but very unclear as to why
			cooxml.append( "<c:axId val=\"0\"/>" );
			cooxml.append( "\r\n" );
		}

		cooxml.append( "</c:bar3DChart>" );
		cooxml.append( "\r\n" );

		return cooxml;
	}
}