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

import org.openxls.formats.XLS.XLSRecord;

/**
 * <b>ChartFormatLink: Not Used.  Great. (0x1022)</b>
 */
public class ChartFormatLink extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6928761103400718842L;

	@Override
	public void init()
	{
		super.init();
	}

	public static XLSRecord getPrototype()
	{
		ChartFormatLink cfl = new ChartFormatLink();
		cfl.setOpcode( CHARTFORMATLINK );
		cfl.setData( new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 } );    // unused
		return cfl;
	}
}