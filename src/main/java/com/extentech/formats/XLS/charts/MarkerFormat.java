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
 * <b>MarkerFormat: Style of a Line Marker(0x1009)</b>
 * <p/>
 * The MarkerFormat record specifies the color, size, and shape of the associated data markers that appear on line,
 * radar, and scatter chart groups. The associated data markers are specified by the preceding DataFormat record.
 * If this record is not present  then the color, size, and shape of the associated data markers are specified by the
 * default values of the fields of this record.
 * <p/>
 * Offset		Name		Size		Contents
 * 4			rgbFore		4			Foreground color: RGB value (high byte = 0)
 * 8			rgbBack		4			Background color: RGB value (high byte = 0)
 * 12			imk			2			Type of marker
 * 0 = no marker
 * 1 = square
 * 2 = diamond
 * 3 = triangle
 * 4 = X
 * 5 = star
 * 6 = Dow-Jones
 * 7 = standard deviation
 * 8 = circle
 * 9 = plus sign
 * 14			grbit		2			Format flags
 * 16			icvFore		2			Index to color of marker border
 * 18			icvBack		2			Index to color of marker fill
 * 20			miSize		4			Size of line markers.  An unsigned integer that specifies the size in twips of the data marker. MUST be greater than or equal to 40 and less than or equal to 1440. The default value for this field is 100.
 * <p/>
 * The icvBack field describes the color of the marker's background, such as the center of the square,
 * while the icvFore field describes the color of the border or the marker itself. The imk field defines
 * the type of marker.
 * <p/>
 * The grbit field contains the following option flags.
 * <p/>
 * Offset		Bits		Mask		Name		Contents
 * 0			0			01h			fAuto		Automatic color
 * 0			3-1			0Eh			(reserved)	Reserved; must be zero
 * 0			4			10h			fNotShowInt	1 = "background = none"
 * 0			5			20h			fNotShowBrd	1 = "foreground = none"
 * 0			7-6			C0h			(reserved)	Reserved; must be zero
 * 1			7-0			FFh			(reserved)	Reserved; must be zero
 */
public class MarkerFormat extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 7526015026522467305L;
	private java.awt.Color rgbFore, rgbBack;
	private int miSize = 0;
	private short imk = 0;
	private short icvFore = 0;
	private short icvBack = 0;
	private short grbit = 0;

	public void init()
	{
		super.init();
		byte[] data = this.getData();
		rgbFore = new java.awt.Color( (data[0] < 0 ? 255 + data[0] : data[0]),
		                              (data[1] < 0 ? 255 + data[1] : data[1]),
		                              (data[2] < 0 ? 255 + data[2] : data[2]) );
		rgbBack = new java.awt.Color( (data[4] < 0 ? 255 + data[4] : data[4]),
		                              (data[5] < 0 ? 255 + data[5] : data[5]),
		                              (data[6] < 0 ? 255 + data[6] : data[6]) );
		imk = ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
		grbit = ByteTools.readShort( this.getByteAt( 10 ), this.getByteAt( 11 ) );
		icvFore = ByteTools.readShort( this.getByteAt( 12 ), this.getByteAt( 13 ) );
		icvBack = ByteTools.readShort( this.getByteAt( 12 ), this.getByteAt( 13 ) );
		miSize = ByteTools.readInt( this.getBytesAt( 14, 4 ) );
	}

	// 20070716 KSC: Need to create new records
	public static XLSRecord getPrototype()
	{
		MarkerFormat mf = new MarkerFormat();
		mf.setOpcode( MARKERFORMAT );
		mf.setData( mf.PROTOTYPE_BYTES );
		return mf;
	}        // imk was 2 in default

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 77, 0, 77, 0, 60, 0, 0, 0 };

	/**
	 * @return String XML representation of this chart-type's options
	 */
	// TODO: Finish MarkerFormat Options
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		//if (imk!=0)
		sb.append( " MarkerFormat=\"" + imk + "\"" );
		return sb.toString();
	}

	/**
	 * Handle setting options from XML in a generic manner
	 */
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "MarkerFormat" ) )
		{
			imk = Short.parseShort( val );
			bHandled = true;
		}
		if( bHandled )
		{
			updateRecord();
		}
		return bHandled;
	}

	private void updateRecord()
	{
		imk = ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
		byte[] b = ByteTools.shortToLEBytes( imk );
		this.getData()[8] = b[0];
		this.getData()[9] = b[1];
	}

	/**
	 * return the Marker Format field imk
	 *
	 * @return
	 */
	public int getMarkerFormat()
	{
		return imk;
	}

	/**
	 * set the Marker Format field imk + update the record
	 *
	 * @param imk
	 */
	public void setMarkerFormat( short imk )
	{
		this.imk = imk;
		updateRecord();
	}

	public String toString()
	{
		return "MarkerFormat: imk= " + imk + " grbit=" + grbit;
	}

	/**
	 * returns the SVG defs element used, in conjuction with getMarkerSVG, to define markers on charts in SVG
	 *
	 * @return String SVG marker defs
	 */
	public static String getMarkerSVGDefs()
	{
		return "<defs>\r\n" +        //-5   -5
				"<rect onmouseover='highLight(evt);'  onmouseout='restore(evt)' id='square1' x='0' y='-5' width='8' height='8'  stroke-width='1'/>\r\n" +
				"<rect onmouseover='highLight(evt);'  onmouseout='restore(evt)' id='diamond1' x='0' y='-5' width='8' height='8'  stroke-width='1' transform='rotate(45 1 1)'/>\r\n" +
				"<path id='triangle1' d='M 0,0 L 5,-10 L 10,0 z'/>\r\n" +
				"<circle id='circle1' cx='1' cy='1' r='5' stroke-width='1'/>\r\n" +
				"<path id='cross1' d='M0,0 H10 M5,5 V-5' stroke-width='2' transform='rotate(45 5 0)'/>\r\n" +
				"<path id='plus1' d='M 0,0 H10 M5,5 V-5' stroke-width='2'/>\r\n" +
				// "star
				// "dow-jones
				// "stddev
				"</defs>\r\n";
	}

	/**
	 * returns the svg line to define the marker at the point x,y, in color clr, of marker style marker (1-9)
	 * <br>NOTE: you must use getMarkerSVGdefs before using this method
	 *
	 * @param x      double x point
	 * @param y      double y point
	 * @param clr    String SVG color string
	 * @param marker int marker format (1-9)
	 * @return SVG
	 */
	public static String getMarkerSVG( double x, double y, String clr, int marker )
	{
		String markersvg = "<use x='" + x + "' y='" + y + "' stroke='" + clr + "' fill='" + clr + "' xlink:href='#";
		if( marker == 1 ) // square
		{
			markersvg += "square1'/>";
		}
		else if( marker == 2 ) // diamond
		{
			markersvg += "diamond1'/>";
		}
		else if( marker == 3 ) // triangle
		{
			markersvg += "triangle1'/>";
		}
		else if( marker == 4 ) // X
		{
			markersvg += "cross1'/>";
		}
		else if( marker == 5 ) // Star	?????????????
		{
			markersvg += "circle1'/>";
		}
		else if( marker == 6 ) // Dow-Jones ??????????
		{
			markersvg += "circle1'/>";
		}
		else if( marker == 7 ) // Std Dev ?????????????
		{
			markersvg += "circle1'/>";
		}
		else if( marker == 8 )    // circle
		{
			markersvg += "circle1'/>";
		}
		else if( marker == 9 ) // + sign
		{
			markersvg += "plus1'/>";
		}
		return markersvg;
	}
}
    
 
