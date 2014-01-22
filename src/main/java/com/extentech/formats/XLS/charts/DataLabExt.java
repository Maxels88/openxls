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
 * <b>DATALABEXT: Chart Data Label Extension (86Ah)</b>
 * Introduced in Excel 10 (2002) this BIFF record is an FRT record
 * for Charts. This record is the parent of DATALABEXTCONTENTS, but
 * contains no other information.
 * <p/>
 * Record Data
 * Offset	Field Name	Size	Contents
 * 4		rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =086Ah
 * 6		grbitFrt	2		FRT flags; must be zero
 * 8		(unused)	8		Reserved; must be zero
 */
public class DataLabExt extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -1868700214505277636L;

	@Override
	public void init()
	{
		super.init();
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 106, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

	public static XLSRecord getPrototype()
	{
		DataLabExt dl = new DataLabExt();
		dl.setOpcode( DATALABEXTCONTENTS );
		dl.setData( dl.PROTOTYPE_BYTES );
		dl.init();
		return dl;
	}

	private void updateRecord()
	{
	}

}
