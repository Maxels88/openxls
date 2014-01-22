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
import com.extentech.formats.XLS.MSODrawingConstants;
import com.extentech.formats.escher.MsofbtOPT;

/**
 * <b>GelFrame: Fill Data(0x1066)</b>
 * The GelFrame record specifies the properties of a fill pattern for parts of a chart.
 */
public class GelFrame extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 581278144607124129L;
	private java.awt.Color fillColor = null;
	private int fillType = 0;    // default= solid

	public void init()
	{
		super.init();
		// try to interpret
		MsofbtOPT optrec = new MsofbtOPT( MSODrawingConstants.MSOFBTOPT,
		                                  0,
		                                  3 );    //version is always 3, inst is current count of properties.
		optrec.setData( this.getData() );    // sets and parses msoFbtOpt data, including imagename, shapename and imageindex
		fillColor = optrec.getFillColor();
	}

	public String toString()
	{
		return "GelFrame: fillType=" + fillType + " fillColor:" + fillColor.toString();
	}

	/**
	 * return the fill color for this frame
	 *
	 * @return Color Hex String
	 */
	public String getFillColor()
	{
		if( fillColor == null )
		{
			return null;
		}
		return FormatHandle.colorToHexString( fillColor );
	}
}
