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
 * <b>CATLAB: Category Labels (856h)</b>
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for Charts.
 * <p/>
 * Record Data
 * Offset		Field Name	Size	Contents
 * 10			wOffset		2		Distance between category label levels
 * 12			at			2		Category axis label alignment
 * 14			grbit		2		Option flags for category axis labels (see description below)
 * 16			(unused)	2		Reserved; must be zero
 * <p/>
 * The grbit field contains the following category axis label option flags:
 * Bits	Mask	Flag Name			Contents
 * 0		0001h	fAutoCatLabelReal	=1 if the category label skip is automatic =0 otherwise
 * 15-1	FFFEh	(unused)			Reserved; must be zero
 * <p/>
 * <p/>
 * wOffset (2 bytes): An unsigned integer that specifies the distance between the axis and axis label.
 * It contains the offset as a percentage of the default distance. The default distance is equal to 1/3 the height of the font calculated in pixels.
 * MUST be a value greater than or equal to 0 (0%) and less than or equal to 1000 (1000%).
 * at (2 bytes): An unsigned integer that specifies the alignment of the axis label. MUST be a value from the following table:
 * Value			Alignment
 * 0x0001			Top-aligned if the trot field of the Text record of the axis is not equal to 0. Left-aligned if the iReadingOrder field of the Text record of the axis specifies left-to-rightreading order; otherwise, right-aligned.
 * 0x0002		    Center-alignment
 * 0x0003			Bottom-aligned if the trot field of the Text record of the axis is not equal to 0. Right-aligned if the iReadingOrder field of the Text record of the axis specifies left-to-right reading order; otherwise, left-aligned.
 * <p/>
 * A - cAutoCatLabelReal (1 bit): A bit that specifies whether the number of categories (3) between axis labels is set to the default value. MUST be a value from the following table:
 * Value	Description
 * 0	    The value is set to catLabel field as specified by CatSerRange record.
 * 1	    The value is set to the default value. The number of category (3) labels is automatically calculated by the application based on the data in the chart.
 * unused (15 bits): Undefined, and MUST be ignored.
 * reserved (2 bytes): MUST be zero, and MUST be ignored.
 */
public class CatLab extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	short wOffset, at;
	private static final long serialVersionUID = 3042712098138741496L;

	@Override
	public void init()
	{
		super.init();
		wOffset = ByteTools.readShort( this.getData()[4], this.getData()[5] );
		at = ByteTools.readShort( this.getData()[6], this.getData()[7] );
		// cAutoCatLabelReal= (record[8] & 0x1==0x1
	}

	// TODO:  need prototype bytes
	private byte[] PROTOTYPE_BYTES = new byte[]{ (byte) 86, 8, 0, 0, (byte) 100, 0, 2, 0, 86, 66, 0, 0 };

	public static XLSRecord getPrototype()
	{
		CatLab cl = new CatLab();
		cl.setOpcode( CATLAB );
		cl.setData( cl.PROTOTYPE_BYTES );
		cl.init();
		return cl;
	}

	public void setOption( String op, String val )
	{
		if( op.equals( "lblAlign" ) )
		{            // ctr, l, r
			if( val.equals( "ctr" ) )
			{
				at = 2;
			}
			else if( val.equals( "l" ) )
			{
				at = 1;
			}
			else
			{
				at = 3;
			}
		}
		else if( op.equals( "lblOffset" ) )
		{    // 0-100
			wOffset = (short) Integer.parseInt( val );
		}
		updateRecord();
	}

	public String getOption( String op )
	{
		if( op.equals( "lblAlign" ) )
		{            // ctr, l, r
			if( at == 2 )
			{
				return "ctr";
			}
			if( at == 1 )
			{
				return "l";
			}
			return "r";
		}
		if( op.equals( "lblOffset" ) )    // 0-100
		{
			return Integer.toString( wOffset );
		}
		return null;
	}

	private void updateRecord()
	{
		byte[] b = new byte[2];
		b = ByteTools.shortToLEBytes( wOffset );
		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
		b = ByteTools.shortToLEBytes( at );
		this.getData()[6] = b[0];
		this.getData()[7] = b[1];
	}

}
