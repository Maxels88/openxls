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

import org.openxls.formats.XLS.XLSRecord;
import org.openxls.toolkit.ByteTools;
import org.openxls.formats.XLS.XLSConstants;

/**
 * <b>Dropbar: Defines Drop Bars (0x103d)</b>
 * Controls up or down bars on a line (or stock, for 2007 v) chart with multiple series
 * the first dropBar record controls upBars
 * the second record controls downBars
 * Also, if these records exist, SeriesList cSer > 1
 * <p/>
 * <p/>
 * <p/>
 * pcGap (2 bytes): A signed integer that specifies the width of the gap between the up bars or the down bars. MUST be a value between 0 and 500.
 * The width of the gap in SPRCs can be calculated by the following formula:
 * Width of the gap in SPRCs = 1 + pcGap
 * <p/>
 * <br>
 * The DropBar record occurs in the ChartFormat subrecord stream after the Legend record,
 * and contains subrecords LineFormat, AreaFormat, [GelFrame], [ShapeProps]
 */
public class Dropbar extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6826327230442065566L;
	short pcGap;

	@Override
	public void init()
	{
		super.init();
		pcGap = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
	}

	public static XLSRecord getPrototype()
	{
		Dropbar db = new Dropbar();
		db.setOpcode( XLSConstants.DROPBAR );
		db.setData( new byte[]{ -106, 0 } );    // 150 is default gap width
		db.init();
		return db;
	}

	/**
	 * sets the width of the gap between the up bars or the down bars. MUST be a value between 0 and 500.
	 * The width of the gap in SPRCs can be calculated by the following formula:
	 * Width of the gap in SPRCs = 1 + pcGap
	 *
	 * @param gap
	 */
	public void setGapWidth( int gap )
	{
		pcGap = (short) gap;
		byte[] b = ByteTools.shortToLEBytes( pcGap );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	/**
	 * returns the width of the gap between the up bars or the down bars.
	 * <br>The gap width is a value between 0 and 500.
	 * <br>The width of the gap in SPRCs can be calculated by the following formula:
	 * Width of the gap in SPRCs = 1 + pcGap
	 *
	 * @return
	 */
	public short getGapWidth()
	{
		return pcGap;
	}

	/**
	 * return the OOXML to define this ChartLine
	 *
	 * @return
	 */
	public StringBuffer getOOXML( boolean upBars )
	{
		StringBuffer cooxml = new StringBuffer();
		String tag;
		if( upBars )
		{
			tag = "c:upBars>";
		}
		else
		{
			tag = "c:downBars>";
		}
		cooxml.append( "<" + tag );
		LineFormat lf = (LineFormat) chartArr.get( 0 );
		cooxml.append( lf.getOOXML() );
		// TODO: finish this logic
		if( !parentChart.getWorkBook().getIsExcel2007() )
		{
			AreaFormat af = (AreaFormat) chartArr.get( 1 );
			cooxml.append( af.getOOXML() );
		}
		cooxml.append( "</" + tag );
		return cooxml;
	}
}
