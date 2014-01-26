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
import org.openxls.toolkit.ByteTools;

/**
 * <b>AxisLineFormat:  Defines a Line that spans an Axis (0x1021)</b>
 * The AxisLineFormat  record specifies which part of the axis is specified by the
 * LineFormat record that follows.
 * <br>
 * 4		id		2		Axis Line identifier:
 * 0= the axis line itself,
 * 1= major grid line,
 * 2= minor grid line,
 * 3= walls or floor
 * if 3, MUST be preceded by an Axis record with the wType set to:
 * 0x0000	The walls of a 3-D chart
 * 0x0001	The floor of a 3-D chart
 */
public class AxisLineFormat extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5243346695500373630L;
	private short id = 0;

	@Override
	public void init()
	{
		super.init();
		id = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
	}

	//  id MUST be greater than the id field values in preceding AxisLine records in the current axis
	public static XLSRecord getPrototype()
	{
		AxisLineFormat a = new AxisLineFormat();
		a.setOpcode( AXISLINEFORMAT );
		a.setData( a.PROTOTYPE_BYTES );
		a.init();
		return a;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0 };
	public static final int ID_AXIS_LINE = 0;
	public static final int ID_MAJOR_GRID = 1;
	public static final int ID_MINOR_GRID = 2;
	public static final int ID_WALLORFLOOR = 3;

	public short getId()
	{
		return id;
	}

	public void setId( int id )
	{
		this.id = (short) id;
		byte[] b = ByteTools.shortToLEBytes( this.id );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	public String toString()
	{
		return "AxisLineFormat: " + ((id == 0) ? "Axis" : ((id == 1) ? "Major" : ((id == 2) ? "Minor" : ((id == 3) ? "Wall or Floor" : "Unknown"))));
	}

}
