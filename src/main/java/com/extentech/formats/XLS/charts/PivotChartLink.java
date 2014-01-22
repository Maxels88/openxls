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

import com.extentech.formats.XLS.XLSRecord;

/**
 * <b>PIVOTCHARTLINK: Pivot Chart Link (861h)</b>
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for Charts.
 * This record stores the link to a PivotTable for a PivotChart. Similar in function
 * to SXVIEWLINK but used only during copy & paste of a chart via BIFF.
 * <p/>
 * Record Data
 * Offset	Field Name	Size	Contents
 * 4		rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =0861h
 * 6		grbitFrt	2		FRT flags; must be zero
 * 8		ai			var		same as AI record
 */
public class PivotChartLink extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -6202325538826559210L;

	@Override
	public void init()
	{
		super.init();
	}    // TODO: Prototype Bytes

	private byte[] PROTOTYPE_BYTES = new byte[]{ };

	public static XLSRecord getPrototype()
	{
		PivotChartLink pcl = new PivotChartLink();
		pcl.setOpcode( PIVOTCHARTLINK );
		pcl.setData( pcl.PROTOTYPE_BYTES );
		pcl.init();
		return pcl;
	}

	private void updateRecord()
	{
	}

}
