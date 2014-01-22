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
import com.extentech.formats.OOXML.SpPr;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.toolkit.ByteTools;

/**
 * <b>LineFormat: Style of a Line or border(0x1007)</b>
 * <p/>
 * 4		rgb		4		Color of line: high byte must be 0
 * 8		lnx		2		Pattern of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
 * When the value of this field is 0x0005 (None), the values of we and icv MUST be set to: Line thickness (we)= 0xFFFF (Hairline)   Line color (icv)  0x004D
 * 10		we		2		Weight of line	(-1= hairline, 0= narrow, 1= med (double), 2= wide (triple)
 * 12		grbit	2		flags
 * 14		icv				Index to color of line.
 * An Icv that specifies a color from the chart color table. This value MUST be greater than or equal to 0x0000 and less than or equal to 0x0041,
 * or greater than or equal to 0x004D and less than or equal to 0x00004F. This value SHOULD NOT be less than 0x0008.
 * 0x0040 == Default foreground color. This is the window text color in the sheet display.
 * 0x0041 == Default background color. This is the window background color in the sheet display and is the default background color for a cell.
 * 0x004D == Default chart foreground color. This is the window text color in the chart display.
 * <p/>
 * grbit:
 * 0		0x1		fAuto		Automatic format
 * 1				reserved, 0
 * 2		0x4		fAxisOn		specifies whether axis line is displayed
 * If the previous record is AxisLine and the value of the id field of the AxisLine record is equal to 0x0000, this field MUST be a value from the following table:
 * fAxisOn			Lns								Meaning
 * 0				0x0005							The axis line is not displayed.
 * 0				Any legal value except 0x0005	The axis line is displayed.
 * 1				Any legal value					The axis line is displayed.
 * If the previous record is not AxisLine and the value of the id field of the AxisLine record is equal to 0x0000, this field MUST be zero, and MUST be ignored.
 * 3		0x8		fAutoColor		specifies whether icv= 0x4D.  if 1, icv must= 0x4D.
 */
public class LineFormat extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 3051781109844837056L;
	private java.awt.Color rgb;
	private short lnx = 0;
	private short we = 0;
	private short grbit = 0;
	private short icv = 0;
	private SpPr sppr = null;
	public static final int SOLID = 0;
	public static final int DASH = 1;
	public static final int DOT = 2;
	public static final int DASHDOT = 3;
	public static final int DASHDASHDOT = 4;
	public static final int NONE = 5;
	public static final int DKGRAY = 6;
	public static final int MEDGRAY = 7;
	public static final int LTGRAY = 8;

	@Override
	public void init()
	{
		super.init();
		byte[] data = this.getData();
		rgb = new java.awt.Color( ((data[0] < 0) ? (255 + data[0]) : data[0]),
		                          ((data[1] < 0) ? (255 + data[1]) : data[1]),
		                          ((data[2] < 0) ? (255 + data[2]) : data[2]) );
		lnx = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
		we = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
		grbit = ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
		icv = ByteTools.readShort( this.getByteAt( 10 ), this.getByteAt( 11 ) );
	}

	public static XLSRecord getPrototype()
	{
		LineFormat lf = new LineFormat();
		lf.setOpcode( LINEFORMAT );
		lf.setData( lf.PROTOTYPE_BYTES );
		lf.init();
		return lf;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0, 0, 0, 0, 0, -1, -1, 9, 0, 77, 0 };    // no line default
	// TODO: Figure this out!!
	private byte[] PROTOTYPE_BYTES_1 = new byte[]{ -128, -128, -128, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

	/**
	 * get new Line Format in desired pattern and weight
	 * <br>pattern= Pattern of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
	 * <br>Weight of line	(-1= hairline, 0= narrow, 1= med (double), 2= wide (triple)
	 * <br>note color is set to default
	 *
	 * @return new LineFormat record
	 */
	public static XLSRecord getPrototype( int style, int weight )
	{
		LineFormat lf = new LineFormat();
		lf.setOpcode( LINEFORMAT );
		lf.setData( lf.PROTOTYPE_BYTES_1 );
		lf.init();
		lf.setLineStyle( style );
		lf.setLineWeight( weight );
		return lf;
	}

	/**
	 * 10		we		2		Weight of line	(-1= hairline, 0= narrow, 1= med (double), 2= wide (triple)
	 */
	public void setLineWeight( int weight )
	{
		we = (short) weight;
		updateRecord();
	}

	/**
	 * Style of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
	 *
	 * @param style
	 */
	public void setLineStyle( int style )
	{
		lnx = (short) style;
		//  When the value of this field is 0x0005 (None), the values of we and icv MUST be set to: Line thickness (we)= 0xFFFF (Hairline)   Line color (icv)  0x004D
		//* 10		we		2		Weight of line	(-1= hairline, 0= narrow, 1= med (double), 2= wide (triple)
		if( lnx == 5 )
		{
			we = -1;
			grbit = 0x8;    // auto color
			setLineColor( 0x4D );
		}
		updateRecord();
	}

	/**
	 * return Pattern of line (0= solid, 1= dash, 2= dot, 3= dash-dot,4= dash dash-dot, 5= none, 6= dk gray pattern, 7= med. gray, 8= light gray
	 */
	public int getLineStyle()
	{
		return lnx;
	}

	/**
	 * Index to color of line
	 *
	 * @param clr
	 */
	public void setLineColor( int clr )
	{
		if( (clr > -1) && (clr < this.getColorTable().length) )
		{
			icv = (short) clr;
			rgb = this.getColorTable()[clr];
			updateRecord();
		}
		else if( clr == 0x4D )
		{ // special flag, default fg
			icv = (short) clr;
			rgb = this.getColorTable()[0];
			updateRecord();
		}
		// TOOD: finish
		if( sppr != null )
		{
//    		sppr.setLine(we, clr);
		}
	}

	/**
	 * return the line color as a hex String
	 *
	 * @return
	 */
	public String getLineColor()
	{
		return FormatHandle.colorToHexString( rgb );
	}

	/**
	 * update the underlying data
	 */
	private void updateRecord()
	{
		byte[] b = new byte[4];
		b[0] = (byte) rgb.getRed();
		b[1] = (byte) rgb.getGreen();
		b[2] = (byte) rgb.getBlue();
		b[3] = 0;    // reserved/0
		System.arraycopy( b, 0, this.getData(), 0, 4 );
		b = ByteTools.shortToLEBytes( lnx );
		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
		b = ByteTools.shortToLEBytes( we );
		this.getData()[6] = b[0];
		this.getData()[7] = b[1];
		b = ByteTools.shortToLEBytes( grbit );
		this.getData()[8] = b[0];
		this.getData()[9] = b[1];
		b = ByteTools.shortToLEBytes( icv );
		this.getData()[10] = b[0];
		this.getData()[11] = b[1];
	}

	public String toString()
	{
		return "LineFormat: LinePattern=" + lnx + " Weight=" + we + " Draw Ticks=" + ((grbit & 0x4) == 0x4);
	}

	public String getSVG()
	{
		if( lnx == 5 )
		{
			return "";    // no line
		}
		float sz = 1f;
		if( we == -1 )    // hairline
		{
			sz = 1f;
		}
		else if( we == 0 )    // narrow	- rest are just guesses really
		{
			sz = 2f;
		}
		else if( we == 1 )    // medium
		{
			sz = 4f;
		}
		else if( we == 2 ) // wide
		{
			sz = 6f;
		}
		String clr = ChartType.getMediumColor();
		if( lnx == DKGRAY )    // dark grey pattern
		{
			clr = ChartType.getDarkColor();
		}
		else if( lnx == MEDGRAY )    // medium grey pattern
		{
			clr = ChartType.getMediumColor();
		}
		else if( lnx == LTGRAY )    // light grey pattern
		{
			clr = ChartType.getLightColor();
		}
		String style = "";
		if( lnx == DASH )
		{
			style = " style='stroke-dasharray: 9, 5;' ";
		}
		else if( lnx == DOT )
		{
			style = " style='stroke-dasharray:2, 2;' ";
		}
		else if( lnx == DASHDOT )
		{
			style = " style='stroke-dasharray: 3, 2, 9, 2;' ";
		}
		else if( lnx == DASHDASHDOT )
		{
			style = " style='stroke-dasharray: 9, 5, 9, 5, 3, 2;' ";
		}
		return " stroke='" + clr + "'  stroke-opacity='1' stroke-width='" + sz + "' " + style + "stroke-linecap='butt' stroke-linejoin='miter' stroke-miterlimit='4'";
	}

	/**
	 * if have 2007+v settings, use them.  Otherwise, interpret line settings to OOXML
	 *
	 * @return
	 */
	public String getOOXML()
	{
		// TODO:  if changed lineformat info, change in sp
		if( sppr != null )
		{
			return sppr.getOOXML();
		}

		if( !parentChart.getWorkBook().getIsExcel2007() )
		{
			StringBuffer ooxml = new StringBuffer();
			ooxml.append( "<c:spPr>" );
			// TODO: line styles + weight
			ooxml.append( "<a:ln w=\"" + we + "\">" );
			ooxml.append( "<a:solidFill>" );
			ooxml.append( "<a:srgbClr val=\"" + FormatHandle.colorToHexString( rgb ) + "\"/>" );
			ooxml.append( "</a:solidFill>" );
			ooxml.append( "</a:ln>" );
			ooxml.append( "</c:spPr>" );
			return ooxml.toString();
		}
		return "";
	}

	/**
	 * sets the OOXML settings for this line
	 *
	 * @param sp
	 */
	public void setFromOOXML( SpPr sp )
	{
		this.sppr = sp;
		int lw = sp.getLineWidth();
		this.setLineWeight( lw );        // sp lw in emus.  1 pt= 12700 emus.
		this.setLineColor( sp.getLineColor() );
		this.setLineStyle( sp.getLineStyle() );
	}
}
