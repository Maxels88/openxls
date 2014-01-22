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
 * <b>Begin:  Begin identifier (0x1033)</b>
 * <p/>
 * Begin is an identifier record for the chart record type.  There is no data to a begin record,
 * and every begin record must have a corrosponding end record
 */
public class Begin extends GenericChartObject implements ChartObject
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2732787714371128511L;

	@Override
	public void init()
	{
		super.init();
	}

	protected static XLSRecord getPrototype()
	{
		Begin bl = new Begin();
		bl.setOpcode( BEGIN );
		bl.setData( new byte[]{ } );
		return bl;
	}

}
