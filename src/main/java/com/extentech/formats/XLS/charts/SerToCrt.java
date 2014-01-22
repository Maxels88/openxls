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
 * <b>SerToCrt: Series Chart-Broup Index (0x1045)</b>
 * 0	chartGroup	2	chart-group index:  the number of the chart group (specified by a CHARTFORMAT record, starts at 0)
 */
public class SerToCrt extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8217594656389677975L;

	public void init()
	{
		super.init();
	}

	protected static XLSRecord getPrototype()
	{
		SerToCrt bl = new SerToCrt();
		bl.setOpcode( SERTOCRT );
		bl.setData( new byte[]{ 0, 0 } );
		return bl;
	}
}
