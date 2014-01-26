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

import org.openxls.formats.XLS.BiffRec;
import org.openxls.formats.XLS.XLSRecord;
import org.openxls.toolkit.ByteTools;
/**
 * <b>Frame: Defines Border Shape Around Displayed Text (0x1032)</b>
 *
 */

/**
 * frt (2 bytes): An unsigned integer that specifies the type of frame to be drawn. MUST be a value from the following table:
 * Type of frame
 * 0x0000	    A frame surrounding the chart element.
 * <p/>
 * 0x0004		A frame with shadow surrounding the chart element.
 * <p/>
 * A - fAutoSize (1 bit): A bit that specifies if the size of the frame is automatically calculated. If the value is 1, the size of the frame is automatically calculated. In this case, the width and height specified by the chart element are ignored and the size of the frame is calculated automatically. If the value is 0, the width and height specified by the chart element are used as the size of the frame.
 * <p/>
 * B - fAutoPosition (1 bit): A bit that specifies if the position of the frame is automatically calculated. If the value is 1, the position of the frame is automatically calculated. In this case, the (x, y) specified by the chart element are ignored, and the position of the frame is automatically calculated. If the value is 0, the (x, y) location specified by the chart element are used as the position of the frame.
 */
public class Frame extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6302152932127918650L;
	boolean fAutoSize;
	boolean fAutoPosition;
	int frt;

	@Override
	public void init()
	{
		frt = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		byte flag = getByteAt( 2 );// 3= autosize, autoposition
		fAutoSize = (flag & 0x01) == 0x01;
		fAutoPosition = (flag & 0x02) == 0x02;
		super.init();
	}

	public static XLSRecord getPrototype()
	{
		Frame f = new Frame();
		f.setOpcode( FRAME );
		f.setData( new byte[]{ 0, 0, 2, 0 } );
		f.init();
		return f;
	}

	/**
	 * return the bg color assoc with this frame rec
	 * in main chart record
	 * NOTE that bg color is defined by the frame rec's associated
	 * AreaFormat Record's FOREGROUND color (icvFore)
	 *
	 * @return bg color Hex String
	 */
	public String getBgColor()
	{
		return getBgColor( chartArr );
	}

	/**
	 * static utility to return the bg color assoc with the desired object
	 * <br>first checks if a gelFrame record exists; if so, it uses that color.
	 * <br>if no gelFrame record exists, it looks for an AreaFormat record.
	 *
	 * @return bg color Hex String
	 */
	public static String getBgColor( java.util.ArrayList chartArr )
	{
		GelFrame gf = (GelFrame) Chart.findRec( chartArr, GelFrame.class );
		if( gf != null )
		{
			return gf.getFillColor();
		}
		AreaFormat af = (AreaFormat) Chart.findRec( chartArr, AreaFormat.class );
		if( af != null )
		{
			return af.getFillColorStr();
		}
		return null;    // use default
	}

	/**
	 * sets the background color assoc with this frame rec
	 * NOTE that bg color is defined by the frame rec's associated
	 * AreaFormat Record's FOREGROUND color (icvFore)
	 *
	 * @param bg color int
	 */
	public void setBgColor( int bg )
	{
		for( XLSRecord aChartArr : chartArr )
		{
			BiffRec b = aChartArr;
			if( b instanceof AreaFormat )
			{
				((AreaFormat) b).seticvFore( bg );
			}
		}
	}

	/**
	 * set the frame autosize/autoposition, necessary for legend expansion
	 *
	 * @see Legend.setAutoPosition, Pos.setAutosizeLegend
	 * [BugTracker 2844]
	 */
	public void setAutosize()
	{
		getData()[2] = 3; // sets both fAutoSize and fAutoPostion
	}

	/**
	 * adds a frame box with the desired lineweight and fill color
	 *
	 * @param lw      if -1, none
	 * @param lclr
	 * @param bgcolor if -1, none
	 */
	public void addBox( int lw, int lclr, int bgcolor )
	{
		LineFormat lf = (LineFormat) Chart.findRec( chartArr, LineFormat.class );
		if( lf == null )
		{
			lf = (LineFormat) LineFormat.getPrototype( 0, 0 );
			addChartRecord( lf );
		}
		if( lw != -1 )
		{
			lf.setLineWeight( lw );
			lf.setLineStyle( 0 );    // solid
		}
		else
		{
			lf.setLineStyle( 5 );    // none
		}
		if( lclr != -1 )
		{
			lf.setLineColor( lclr );
		}
		AreaFormat af = (AreaFormat) Chart.findRec( chartArr, AreaFormat.class );
		if( af == null )
		{
			af = (AreaFormat) AreaFormat.getPrototype( 1 );
			addChartRecord( af );
		}
		if( bgcolor != -1 )
		{
			af.seticvBack( bgcolor );
		}
	}

	/**
	 * returns true if this Frame is surrounded by a box
	 *
	 * @return
	 */
	public boolean hasBox()
	{
		LineFormat l = (LineFormat) Chart.findRec( chartArr, LineFormat.class );
		return (l.getLineStyle() != LineFormat.NONE);
	}

	/**
	 * return the line color as a hex String
	 *
	 * @return
	 */
	public String getLineColor()
	{
		LineFormat l = (LineFormat) Chart.findRec( chartArr, LineFormat.class );
		if( l != null )
		{
			return l.getLineColor();
		}
		return null;
	}

	/**
	 * return the svg representation of this Frame element
	 *
	 * @param coords
	 * @return
	 */
	public StringBuffer getSVG( float[] coords )
	{
		StringBuffer svg = new StringBuffer();
		String lineSVG = "";
		String bgclr = getBgColor();
//		String bgclr= "white";
		if( bgclr == null )
		{
			bgclr = "white";
		}

		LineFormat lf = (LineFormat) Chart.findRec( chartArr, LineFormat.class );
		if( lf != null )
		{
			lineSVG = lf.getSVG();
		}

		float x = coords[0] - (coords[2] / 2);    // apparently coords are center-point; adjust
		float y = coords[1] - (coords[3] / 2);
		svg.append( "<rect x='" + x + "' y='" + y + "' width='" + coords[2] + "' height='" + coords[3] +
				            "' fill='" + bgclr + "' fill-opacity='1' " + lineSVG + "/>\r\n" );

		return svg;
	}
}
