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
 * <b>DefaultText: Default Data Label Text Properties(0x1024)</b>
 * <p/>
 * id (2 bytes): An unsigned integer that specifies the text elements that are formatted using the position and appearance information specified by the Text record immediately following this record. MUST be a value from the following table.
 * If this record is in a sequence of records that conforms to the CRT rule as specified by the Chart Sheet Substream ABNF, then this field MUST be 0x0000 or 0x0001. If this record is not in a sequence of records that conforms to the CRT rule as specified by the Chart Sheet Substream ABNF, then this field MUST be 0x0002 or 0x0003.
 * Value		Meaning
 * 0x0000		Format all Text records in the chart group where fShowPercent is equal to 0 or fShowValue is equal to 0.
 * 0x0001		Format all Text records in the chart group where fShowPercent is equal to 1 or fShowValue is equal to 1.
 * 0x0002		Format all Text records in the chart where the value of fScaled of the associated FontInfo structure is equal to 0.
 * 0x0003		Format all Text records in the chart where the value of fScaled of the associated FontInfo structure is equal to 1. *
 */
public class DefaultText extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6845131434077760152L;

	private short grbit = 0;

	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
	}

	// 20070716 KSC: Need to create new records
	public static XLSRecord getPrototype()
	{
		DefaultText t = new DefaultText();
		t.setOpcode( DEFAULTTEXT );
		t.setData( t.PROTOTYPE_BYTES );
		t.init();
		return t;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0 };

	public static final int TYPE_SHOWLABELS = 0;
	public static final int TYPE_VALUELABELS = 1;
	public static final int TYPE_ALLTEXT = 2;
	public static final int TYPE_UNKNOWN = 3;

	// try to interpret!!!
	public int getType()
	{
		return grbit;
	}

	public void setType( short type )
	{
		grbit = type;
		byte[] b = ByteTools.shortToLEBytes( grbit );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
	}

}
