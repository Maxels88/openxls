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
 * <b>Fontx: Font Index (0x1026)</b>
 * Child of a TEXT record and defines a text font by indexing the appropriate font in
 * the font table.
 * <p/>
 * 4		iFont		2		Index number into the font table
 * If this field is less than or equal to the number of Font records
 * in the workbook, this field is a one-based index to a Font record in
 * the workbook. Otherwise, this field is a one-based index into the
 * collection of Font records in the chart sheet substream, where the index is equal to
 * iFont minus n and n is the number of Font records in the workbook.
 */
public class Fontx extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4255798925225768809L;
	// 20070806 KSC: Add init/update to control FontX opts
	private short ifnt = 0;

	@Override
	public void init()
	{
		super.init();
		ifnt = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
	}

	public static XLSRecord getPrototype()
	{
		Fontx f = new Fontx();
		f.setOpcode( FONTX );
		f.setData( f.PROTOTYPE_BYTES );
		f.init();
		return f;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 5, 0 };

	/**
	 * returns the index into the wb font table referenced by this Fontx record for the chart
	 */
	/*
	If this field is less than or equal to the number of Font records 
	in the workbook, this field is a one-based index to a Font record in 
	the workbook. Otherwise, this field is a one-based index into the 
	collection of Font records in the chart sheet substream, where the index is equal to 
	iFont minus n and n is the number of Font records in the workbook.
	*/
	public int getIfnt()
	{
		int n = 0;
		try
		{
			n = this.getWorkBook().getNumFonts();
		}
		catch( Exception e )
		{
		}
		if( ifnt <= n )
		{
			return ifnt;
		}
		return (ifnt - n) + 1;
	}

	public void setIfnt( int id )
	{
		this.ifnt = (short) id;
		byte[] b = ByteTools.shortToLEBytes( this.ifnt );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
	}
}
