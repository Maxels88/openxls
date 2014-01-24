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
 * <b>Serfmt: Series Format(0x105d)</b>
 * <p/>
 * Specifies series formatting information
 * <p/>
 * 0	grbit		2
 * <p/>
 * bits
 * 0		0x1	fSmoothedLine		1= the line series has a smoothed line (Line, Scatter or Radar)
 * 1		0x2	f3DBubbles			1= draw bubbles with 3-D effects
 * 2		0x4 fArShadow			1= specifies whether the data markers are displayed with a shadow on bubble,
 * scatter, radar, stock, and line chart groups.
 * rest are reserved
 */
public class Serfmt extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -8307035373276421283L;
	// 20070810 KSC: parse options ...
	private short grbit = 0;
	private boolean fSmoothedLine = false;
	private boolean f3dBubbles = false;
	private boolean fArShadow = false;

	@Override
	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		fSmoothedLine = (grbit & 0x1) == 0x1;
		f3dBubbles = (grbit & 0x2) == 0x2;
		fArShadow = (grbit & 0x4) == 0x4;
	}

	private void updateRecord()
	{
		grbit = ByteTools.updateGrBit( grbit, fSmoothedLine, 0 );
		grbit = ByteTools.updateGrBit( grbit, f3dBubbles, 1 );
		grbit = ByteTools.updateGrBit( grbit, fArShadow, 2 );
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0 };

	/**
	 * Handle setting options from XML in a generic manner
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "SmoothedLine" ) )
		{
			fSmoothedLine = val.equals( "true" );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ThreeDBubbles" ) )
		{
			f3dBubbles = val.equals( "true" );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ArShadow" ) )
		{
			fArShadow = val.equals( "true" );
			bHandled = true;
		}
		if( bHandled )
		{
			updateRecord();
		}
		return bHandled;
	}

	/**
	 * @return String XML representation of this chart-type's options
	 */
	@Override
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		if( fSmoothedLine )
		{
			sb.append( " SmoothedLine=\"true\"" );
		}
		if( f3dBubbles )
		{
			sb.append( " ThreeDBubbles=\"true\"" );
		}
		if( fArShadow )
		{
			sb.append( " ArShadow=\"true\"" );
		}
		return sb.toString();
	}

	public static XLSRecord getPrototype()
	{
		Serfmt s = new Serfmt();
		s.setOpcode( SERFMT );
		s.setData( s.PROTOTYPE_BYTES );
		s.init();
		return s;
	}

	public void setHas3dBubbles( boolean has3dBubbles )
	{
		f3dBubbles = has3dBubbles;
		updateRecord();
	}

	public boolean get3DBubbles()
	{
		return f3dBubbles;
	}

	/**
	 * sets whether the parent chart or series has smoothed lines
	 *
	 * @param b
	 */
	public void setSmoothedLine( boolean b )
	{
		fSmoothedLine = b;
		updateRecord();
	}

	public boolean getSmoothLine()
	{
		return fSmoothedLine;
	}

	/**
	 * data markers are displayed with a shadow on bubble,
	 * scatter, radar, stock, and line chart groups.
	 */
	public boolean getShadow()
	{
		return fArShadow;
	}

	/**
	 * data markers are displayed with a shadow on bubble,
	 * scatter, radar, stock, and line chart groups.
	 *
	 * @param b
	 */
	public void setHasShadow( boolean b )
	{
		fArShadow = b;
		updateRecord();
	}

}
