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
 * <b> End:  End of chart substream (0x1034)</b>
 * <p/>
 * End is an identifier record for the chart record type.  There is no data to a end record,
 * and every end record must have a corrosponding begin record
 */
public class End extends GenericChartObject implements ChartObject
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 9022736093645720842L;

	@Override
	public void init()
	{
		super.init();
	}

	protected static XLSRecord getPrototype()
	{
		End bl = new End();
		bl.setOpcode( END );
		bl.setData( new byte[]{ } );
		return bl;
	}
}
