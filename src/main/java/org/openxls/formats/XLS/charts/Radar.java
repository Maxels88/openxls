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
 * <b>Radar: Chart Group Is a Radar Chart Group(0x103e)</b>
 * <p/>
 * 4		grbit		2
 * <p/>
 * 0		0x1		fRdrAxLab			1= chart contains radar axis labels
 * 1		0x2		fHasShadow			1= this radar series has a shadow
 */
public class Radar extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 3443368503725347910L;
	private short grbit = 0;
	private boolean fRdrAxLab = true;
	private boolean fHasShadow = false;

	@Override
	public void init()
	{
		super.init();
		chartType = ChartConstants.RADARCHART;
		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		fRdrAxLab = (grbit & 0x1) == 0x1;
		fHasShadow = (grbit & 0x2) == 0x2;
	}

	// 20070703 KSC: taken from Bar.java
	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	public static XLSRecord getPrototype()
	{
		Radar b = new Radar();
		b.setOpcode( XLSConstants.RADAR );
		b.setData( b.PROTOTYPE_BYTES );
		b.init();
		return b;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 1, 0, 18, 0 };

	/**
	 * @return String XML representation of this chart-type's options
	 */
	@Override
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		if( fRdrAxLab )
		{
			sb.append( " AxisLabels=\"true\"" );
		}
		if( fHasShadow )
		{
			sb.append( " Shadow=\"true\"" );
		}
		return sb.toString();
	}

	/**
	 * Handle setting options from XML in a generic manner
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "AxisLabels" ) )
		{
			fRdrAxLab = val.equals( "true" );
			grbit = ByteTools.updateGrBit( grbit, fRdrAxLab, 0 );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Shadow" ) )
		{
			fHasShadow = val.equals( "true" );
			grbit = ByteTools.updateGrBit( grbit, fHasShadow, 1 );
			bHandled = true;
		}
		if( bHandled )
		{
			updateRecord();
		}
		return bHandled;
	}

}
