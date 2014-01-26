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

public class RadarAreaChart extends RadarChart
{
	protected RadarArea radararea = null;

	public RadarAreaChart( GenericChartObject charttype, ChartFormat cf, WorkBook wb )
	{
		super( charttype, cf, wb );
		radararea = (RadarArea) charttype;
	}

	/**
	 * gets the chart-type specific ooxml representation: <radarChart>
	 *
	 * @return
	 */
	@Override
	public StringBuffer getOOXML( String catAxisId, String valAxisId, String serAxisId )
	{
		StringBuffer cooxml = new StringBuffer();

		// chart type: contains chart options and series data
		cooxml.append( "<c:radarChart>" );
		cooxml.append( "\r\n" );

		String style = "filled";

		cooxml.append( "<c:radarStyle val=\"" + style + "\"/>" );
		// vary colors???

		// *** Series Data: ser, cat, val for most chart types
		cooxml.append( getParentChart().getChartSeries().getOOXML( getChartType(), false, 0 ) );

		// chart data labels, if any
		// TODO: FINISH
		// cooxml.append(getDataLabelsOOXML(cf));

		// axis ids - unsigned int strings
		cooxml.append( "<c:axId val=\"" + catAxisId + "\"/>" );
		cooxml.append( "\r\n" );
		cooxml.append( "<c:axId val=\"" + valAxisId + "\"/>" );
		cooxml.append( "\r\n" );

		cooxml.append( "</c:radarChart>" );
		cooxml.append( "\r\n" );
		return cooxml;
	}
}
