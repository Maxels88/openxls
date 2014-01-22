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
 * <b>PIVOTCHARTBITS: PivotChart Bits (859h)</b>
 * Introduced in Excel 9 (2000), this BIFF record is an FRT
 * record for Charts. This stores flags for a PivotChart.
 * <p/>
 * Record Data
 * Offset	Field Name	Size	Contents
 * 4		rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =0859h
 * 6		grbitFrt	2		FRT flags; must be zero
 * 8		grbit		2		Option flags for PivotCharts (see description below)
 * 10		(unused)	6		Reserved; must be zero
 * <p/>
 * The grbit field contains the following PivotChart option flags:
 * Bits	Mask	Flag Name	Contents
 * 0		0001h	fGXHide		=1 if the field buttons are hidden =0 otherwise
 * 15-1	FFFEh	(unused)	Reserved; must be zero
 */
public class PivotChartBits extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -4587483948421928667L;

	@Override
	public void init()
	{
		super.init();
	}

	// TODO: Prototype Bytes
	private byte[] PROTOTYPE_BYTES = new byte[]{ };

	public static XLSRecord getPrototype()
	{
		PivotChartBits pcb = new PivotChartBits();
		pcb.setOpcode( PIVOTCHARTBITS );
		pcb.setData( pcb.PROTOTYPE_BYTES );
		pcb.init();
		return pcb;
	}

	private void updateRecord()
	{
	}

}
