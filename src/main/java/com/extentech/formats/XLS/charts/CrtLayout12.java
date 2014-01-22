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

import com.extentech.toolkit.ByteTools;

/**
 * The CrtLayout12 record specifies the layout information for attached label
 * (data label or legend)
 * 12  frtHeader	  0x89D
 * 4	dwCheckSum	: An unsigned integer that specifies the checksum of the values in the order as follows, if the checksum is incorrect,
 * the layout information specified in this record MUST be ignored.
 * Checksum for type	Values
 * AttachedLabel		x1 field of the Pos record in the sequence of records that contains this CrtLayout12 record and conforms to the ATTACHEDLABEL rule.
 * y1 field of the Pos record in the in the sequence of records that contains this CrtLayout12 record and conforms to the ATTACHEDLABEL rule.
 * An unsigned integer that specifies whether the attached label is at its default position. MUST be 1 if the dlp field of the Text record in the in the sequence of records that contains this CrtLayout12 record and conforms to the ATTACHEDLABEL rule is equal to 0xA. Otherwise, MUST be zero.
 * Legend				x1 field of the Pos record in the in the sequence of records that contains this CrtLayout12 record and conforms to the LD rule.
 * y1 field of the Pos record in the in the sequence of records that contains this CrtLayout12 record and conforms to the LD rule.
 * Width of the legend in pixels.
 * Height of the legend in pixels.
 * The fAutoPosX field of Legend record.
 * The fAutoPosY field of Legend record.
 * The fAutoSize of the Frame record in the in the sequence of records that contains this CrtLayout12 record and conforms to the LD rule.
 * The width and height of legend in pixels are calculated with the following steps:
 * Get chart area width in pixels
 * chart area width in pixels = (dx field of Chart record - 8) * DPI of the display device / 72
 * If the frt field of the Frame record following the Chart record is 0x0004 and the chart is not embedded, add the shadow size:
 * chart area width in pixels -= 2 * line width of the display device in pixels
 * Get chart area height in pixels
 * chart area height in pixels = (dy field of Chart record - 8) * DPI of the display device / 72
 * If the frt field of the Frame record following the Chart record is 0x0004 and the chart is not embedded, add the shadow size:
 * chart area height in pixels -= 2 * line height of the display device in pixels
 * Compute legend size in pixels
 * legend width in pixels = dx field of Legend / 4000 * chart area width in pixels
 * legend height in pixels = dy field of Legend / 4000 * chart area height in pixels
 * 2					1 bit- unused
 * autolayouttype (4 bits): An unsigned integer that specifies the automatic layout type of the legend.
 * MUST be ignored when this record is in the sequence of records that conforms to the ATTACHEDLABEL rule.
 * MUST be a value from the following table:
 * Value			Meaning
 * 0x0				Align to the bottom
 * 0x1				Align to top right corner
 * 0x2				Align to the top
 * 0x3				Align to the right
 * 0x4				Align to the left
 * reserved1 (11 bits): MUST be zero, and MUST be ignored.
 * 2			    	wXMode 			A CrtLayout12Mode structure that specifies the meaning of x.
 * 2					wYMode 			A CrtLayout12Mode structure that specifies the meaning of y.
 * 2					wWidthMode		A CrtLayout12Mode structure that specifies the meaning of dx.
 * 2					wHeightMode		A CrtLayout12Mode structure that specifies the meaning of dy.
 * 8					x (8 bytes): An Xnum value that specifies a horizontal offset. The meaning is determined by wXMode.
 * 8					y (8 bytes): An Xnum value that specifies a vertical offset. The meaning is determined by wYMode.
 * 8					dx (8 bytes): An Xnum value that specifies a width or an horizontal offset. The meaning is determined by wWidthMode.
 * 8					dy (8 bytes): An Xnum value that specifies a height or an vertical offset. The meaning is determined by wHeightMode.
 * 2					reserved2 (2 bytes): MUST be zero, and MUST be ignored.
 */
public class CrtLayout12 extends GenericChartObject implements ChartObject
{
	byte autolayouttype;
	short wXMode, wYMode;
	short wWidthMode, wHeightMode;
	float x, y, dx, dy;

	@Override
	public void init()
	{
		super.init();
		byte[] data = this.getData();
		autolayouttype = (byte) (data[16] >> 1);
		wXMode = ByteTools.readShort( data[18], data[19] );
		wYMode = ByteTools.readShort( data[20], data[21] );
		wWidthMode = ByteTools.readShort( data[22], data[23] );
		wHeightMode = ByteTools.readShort( data[24], data[25] );
		x = (float) ByteTools.eightBytetoLEDouble( this.getBytesAt( 26, 8 ) );
		y = (float) ByteTools.eightBytetoLEDouble( this.getBytesAt( 34, 8 ) );
		dx = (float) ByteTools.eightBytetoLEDouble( this.getBytesAt( 42, 8 ) );
		dy = (float) ByteTools.eightBytetoLEDouble( this.getBytesAt( 50, 8 ) );
	}

	/**
	 * returns the legend autolayout type:
	 * <br>
	 * 0x0				Align to the bottom
	 * 0x1				Align to top right corner
	 * 0x2				Align to the top
	 * 0x3				Align to the right
	 * 0x4				Align to the left
	 *
	 * @return
	 */
	public int getLayout()
	{
		return autolayouttype;
	}

	/**
	 * sets the layout position of the attached legend or attached label
	 * * <br>
	 * 0x0				Align to the bottom
	 * 0x1				Align to top right corner
	 * 0x2				Align to the top
	 * 0x3				Align to the right
	 * 0x4				Align to the left
	 *
	 * @param pos
	 */
	public void setLayout( int pos )
	{
		autolayouttype = (byte) pos;
		this.getData()[16] = (byte) (autolayouttype << 1);
	}

	public float[] getCoords()
	{
/*		return new float[] { Pos.convertFromSPRC(xTL, w, 0),	// x offset
				 Pos.convertFromSPRC(yTL, 0, h),	// y offset
				 Pos.convertFromSPRC(xBR, w, 0),	// w								
				 Pos.convertFromSPRC(yBR, 0, h) };	// h
	*/
		if( wWidthMode == 0 )
		{
			return null;
		}
		if( wWidthMode == 1 )
		{
			//x and y specify the offset of the top left corner, relative to its default position, as a fraction of the chart area. MUST be greater than or equal to -1.0 and MUST be less than or equal to 1.0. dx and dy specify the width and height, as a fraction of the chart area, MUST be greater than or equal to 0.0, and MUST be less than or equal to 1.0.

		}
		if( wWidthMode == 2 )
		{
			// x and y specify the offset of the upper-left corner; dx and dy specify the offset of the bottom-right corner. x, y, dx and dy are specified relative to the upper-left corner of the chart area as a fraction of the chart area. x, y, dx and dy MUST be greater than or equal to 0.0, and MUST be less than or equal to 1.0.
		}
		return null;
	}
}

