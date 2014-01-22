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
import com.extentech.toolkit.ByteTools;

/**
 * <b>PieFormat: Position of the Pie Slice(0x100b)</b>
 * <p/>
 * percentage		2		distance of pie slice from center of pie as %
 */
public class PieFormat extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 28305957039802849L;
	short percentage = 0;

	public void init()
	{
		super.init();
		this.getData();
		percentage = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );

	}

	// 20070716 KSC: Need to create new records
	public static XLSRecord getPrototype()
	{
		PieFormat pf = new PieFormat();
		pf.setOpcode( PIEFORMAT );
		pf.setData( pf.PROTOTYPE_BYTES );
		pf.init();
		return pf;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0 };

	// 20070717 KSC: get/set methods for format options
	public short getPercentage()
	{
		return percentage;
	}

	public void setPercentage( short p )
	{
		percentage = p;
		updateRecord();
	}

	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( percentage );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
	}

	public String getOptionsXML()
	{
		return " Percentage=\"" + percentage + "\"";
	}

	public String toString()
	{
		return "PieFormat:  Percentage=" + percentage;
	}
}
