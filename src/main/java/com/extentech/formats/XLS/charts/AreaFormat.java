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

import com.extentech.ExtenXLS.FormatHandle;
import com.extentech.formats.OOXML.FillGroup;
import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.XLS.FormatConstants;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;

;
/**
 * <b>AreaFormat: Colors and Patterns for an area(0x100a)</b>
 *
 * 4	rgbFore		4		FgColor RGB Value (High byte is 0)
 * 8	rgbBack		4		Bg Color ""
 * 12	fls			2		Pattern
 * 14	grbit		2		flags
 * 16	icvFore		2 		index to fg color
 * 18	icvBack		2		index to bg color
 *
 * grbit:
 * 		0		0x1		fAuto		Automatic Format
 * 		1		0x2		fInvertNeg	Fg + Bg is swapped when data is neg.
 *
 */

/**
 * more info on colors:
 * <p/>
 * The chart color table is a subset of the full color table.
 * icv (2 bytes): An Icv that specifies a color from the chart color table.
 * MUST be greater than or equal to 0x0008 and less than or equal to 0x003F, or greater than or equal to 0x004D and less than or equal to 0x004F.
 * <p/>
 * This info is not yet verified:
 * <p/>
 * For icvFore, icvBack, must be either 0-7 or 0x40 or 0x41 or icv????
 * "The default value of this field is selected automatically from the next available color in the Chart color table."
 * <p/>
 * For icvBack, must be either 0-7 or 0x40 or 0x41
 * The default value of this field is 0x0009.
 * <p/>
 * icv (chart color table index):  > 0x0008 < 003F OR 0x4D, 0x4E, 0x4F OR 0x7FFF
 * <p/>
 * 0-7 are bascially the normal COLORTABLE entries
 * The icv values greater than or equal to 0x0008 and less than or equal to 0x003F, specify the palette colors in the table.
 * If a Palette record exists in this file, these icv values specify colors from the rgColor array in the Palette record.
 * If no Palette record exists, these values specify colors in the default palette.
 * <p/>
 * The next 56 values in this part of the color table are specified as follows:
 * 0x0040
 * Default foreground color. This is the window text color in the data sheet display.
 * 0x0041
 * Default background color. This is the window background color in the data sheet display and is the default background color for a cell.
 * 0x004D
 * Default chart foreground color. This is the window text color in the chart display.
 * 0x004E
 * Default chart background color. This is the window background color in the chart display.
 * 0x004F
 * Chart neutral color which is black, an RGB value of (0,0,0).
 * 0x7FFF
 * Font automatic color. This is the window text color.
 */
public class AreaFormat extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -437132913972684937L;
	private java.awt.Color rgbFore;
	private java.awt.Color rgbBack;
	private short fls = 0;
	private short grbit = 0;
	private short icvFore = 0;
	private short icvBack = 0;
	boolean fAuto;
	boolean fInvertNeg;

	@Override
	public void init()
	{
		super.init();
		byte[] data = getData();
		grbit = ByteTools.readShort( data[10], data[11] );
		fAuto = (grbit & 0x1) == 0x1;
		fInvertNeg = (grbit & 0x2) == 0x2;
		rgbFore = new java.awt.Color( ((data[0] < 0) ? (255 + data[0]) : data[0]),
		                              ((data[1] < 0) ? (255 + data[1]) : data[1]),
		                              ((data[2] < 0) ? (255 + data[2]) : data[2]) );
		rgbBack = new java.awt.Color( ((data[4] < 0) ? (255 + data[4]) : data[4]),
		                              ((data[5] < 0) ? (255 + data[5]) : data[5]),
		                              ((data[6] < 0) ? (255 + data[6]) : data[6]) );
		fls = ByteTools.readShort( data[8], data[9] );
		icvFore = ByteTools.readShort( data[12], data[13] );
		icvBack = ByteTools.readShort( data[14], data[15] );
	}

	// 20070716 KSC: Need to create new records
	public static XLSRecord getPrototype()
	{
		AreaFormat af = new AreaFormat();
		af.setOpcode( AREAFORMAT );
		af.setData( af.PROTOTYPE_BYTES );
		af.init();
		return af;
	}                        // changed prototype bytes to ensure has default settings i.e. no pattern and automatic color

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 77, 0, 77, 0 };
	private byte[] PROTOTYPE_BYTES_1 = new byte[]{ -64, -64, -64, 0, 0, 0, 0, 0, 1, 0, 0, 0, 22, 0, 79, 0 };
	private byte[] PROTOTYPE_BYTES_2 = new byte[]{ 0, 0, 0, 0, -1, -1, -1, 0, 1, 0, 1, 0, 77, 0, 78, 0 };

	public static XLSRecord getPrototype( int type )
	{
		AreaFormat af = new AreaFormat();
		af.setOpcode( AREAFORMAT );
		if( type == 0 )
		{
			af.setData( af.PROTOTYPE_BYTES_1 );    // default is certain color combo
		}
		else if( type == 1 )
		{
			af.setData( af.PROTOTYPE_BYTES_2 );    // ""
		}
		else
		{
			af.setData( af.PROTOTYPE_BYTES );
		}
		af.init();
		return af;
	}

	private void updateRecord()
	{
		byte[] b = new byte[4];
		b[0] = (byte) rgbFore.getRed();
		b[1] = (byte) rgbFore.getGreen();
		b[2] = (byte) rgbFore.getBlue();
		b[3] = 0;    // reserved/0
		System.arraycopy( b, 0, getData(), 0, 4 );
		b[0] = (byte) rgbBack.getRed();
		b[1] = (byte) rgbBack.getGreen();
		b[2] = (byte) rgbBack.getBlue();
		b[3] = 0;    // reserved/0
		System.arraycopy( b, 0, getData(), 4, 4 );
		b = ByteTools.shortToLEBytes( fls );
		getData()[8] = b[0];
		getData()[9] = b[1];
		grbit = 0;
		if( fAuto )
		{
			grbit = 0x1;
		}
		if( fInvertNeg )
		{
			grbit |= 0x2;
		}
		b = ByteTools.shortToLEBytes( grbit );
		getData()[10] = b[0];
		getData()[11] = b[1];
		b = ByteTools.shortToLEBytes( icvFore );
		getData()[12] = b[0];
		getData()[13] = b[1];
		b = ByteTools.shortToLEBytes( icvBack );
		getData()[14] = b[0];
		getData()[15] = b[1];
	}

	public String toString()
	{
		return "AreaFormat: Pattern=" + fls + " ForeColor=" + icvFore + " BackColor=" + icvBack + " Automatic Format=" + ((grbit & 0x1) == 0x1);
	}

	/**
	 * return the bg color index for this Area Format
	 *
	 * @return
	 */
	public int geticvBack()
	{
		if( icvBack > getColorTable().length )
		{    // then it's one of the special codes:
			return FormatHandle.interpretSpecialColorIndex( icvBack );
		}
		return icvBack;
	}

	/**
	 * return the area fill color
	 *
	 * @return int index into color table
	 */
	public int getFillColor()
	{
		if( fls == 0 ) // no fill
		{
			return FormatConstants.COLOR_WHITE;
		}
		if( fls == 1 ) // solid; forecolor is bg
		{
			return geticvFore();
		}
		if( fls == 2 ) // medium grey
		{
			return FormatConstants.COLOR_GRAY50;
		}
		if( fls == 3 ) // dark grey (possible for more options not yet handled
		{
			return FormatConstants.COLOR_GRAY80;
		}
		if( fls == 4 ) // light grey
		{
			return FormatConstants.COLOR_GRAY25;
		}
		// rest are actual fill patterns TODO handle
		return geticvFore();
	}

	/**
	 * return the area fill color
	 *
	 * @return String color hex string
	 */
	public String getFillColorStr()
	{
		if( fAuto )
		{
			return null;
		}
		if( fls == 0 ) // no fill
		{
			return "#FFFFFF";
		}
		if( fls == 1 ) // solid; forecolor is bg
		{
			return FormatHandle.colorToHexString( rgbFore );
		}
		if( fls == 2 ) // medium grey
		{
			return FormatHandle.colorToHexString( FormatConstants.Gray50 );
		}
		if( fls == 3 ) // dark grey (possible for more options not yet handled
		{
			return FormatHandle.colorToHexString( FormatConstants.Gray80 );
		}
		if( fls == 4 ) // light grey
		{
			return FormatHandle.colorToHexString( FormatConstants.Gray25 );
		}
		// rest are actual fill patterns TODO handle
		return FormatHandle.colorToHexString( rgbFore );
	}

	/**
	 * return the fg color index for this Area Format
	 *
	 * @return
	 */
	public int geticvFore()
	{
		if( icvFore > getColorTable().length )
		{    // then it's one of the special codes:
			return FormatHandle.interpretSpecialColorIndex( icvFore );
		}
		return icvFore;
	}

	public void seticvBack( int clr )
	{
		if( (clr > -1) && (clr < getColorTable().length) )
		{
			icvBack = (short) clr;
			rgbBack = getColorTable()[clr];
			updateRecord();
		}
		else if( clr == 0x4D )
		{ // special flag, default bg
			icvBack = (short) clr;
			rgbBack = getColorTable()[0];
			updateRecord();
		}
	}

	/**
	 * sets the fill color for the area
	 *
	 * @param clr color index
	 */
	public void seticvFore( int clr )
	{

		// fls= 1 The fill pattern is solid. When solid is specified,
		// rgbFore is the only color rendered, even if rgbBack is specified
		if( (clr > -1) && (clr < getColorTable().length) )
		{
			fAuto = false;
			fls = 1;
			icvFore = (short) clr;
			rgbFore = getColorTable()[clr];
			// must also set bg to 9
			seticvBack( 9 );
			//updateRecord();
		}
		else if( clr == 0x4E )
		{ // default fg
			icvFore = (short) clr;
			rgbFore = getColorTable()[1];
			updateRecord();
		}
	}

	/**
	 * sets the fill color for the area
	 *
	 * @param clr Hex Color String
	 */
	public void seticvFore( String clr )
	{
		fAuto = false;
		fls = 1;
		rgbFore = FormatHandle.HexStringToColor( clr );
		icvFore = (short) FormatHandle.HexStringToColorInt( clr, FormatHandle.colorBACKGROUND );        // finds best match
		updateRecord();
	}

	/**
	 * sets the OOXML settings for this Area Format
	 *
	 * @param sp
	 */
	public void setFromOOXML( SpPr sp )
	{
		FillGroup f = sp.getFill();
		if( f != null )
		{
			seticvFore( f.getColor() );
			//this.seticvBack()
			// fls= fill pattern

		}
	}

	public StringBuffer getOOXML()
	{
		StringBuffer ooxml = new StringBuffer();
		ooxml.append( "<a:solidFill>" );
//    	ooxml.append("<a:srgbClr val=\"" + FormatHandle.colorToHexString(rgb) + "\"/>");
		ooxml.append( "</a:solidFill>" );
		return ooxml;
	}

}
