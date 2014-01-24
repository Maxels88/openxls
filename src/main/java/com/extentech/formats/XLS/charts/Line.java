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
 * <b>Line: Chart Group Is a Line Chart Group (0x1018)</b>
 * <p/>
 * 4	grbit		2		flags
 * <p/>
 * 0		0x1		fStacked		Stack the displayed values
 * 1		0x2		f100			Each category is broken down as a percentage
 * 2		0x4		fHasShadow		1= this line has a shadow
 */
public class Line extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6526476252082906554L;
	private short grbit = 0;
	private boolean fStacked = false;
	private boolean f100 = false;
	private boolean fHasShadow = false;

	@Override
	public void init()
	{
		super.init();
		chartType = ChartConstants.LINECHART;    // 20070703 KSC
		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		fStacked = ((grbit & 0x1) == 0x1);
		f100 = ((grbit & 0x2) == 0x2);
		fHasShadow = ((grbit & 0x4) == 0x4);

	}

	// 20070725 KSC:
	private void updateRecord()
	{
		grbit = ByteTools.updateGrBit( grbit, fStacked, 0 );
		grbit = ByteTools.updateGrBit( grbit, f100, 1 );
		grbit = ByteTools.updateGrBit( grbit, fHasShadow, 2 );
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	public static XLSRecord getPrototype()
	{
		Line b = new Line();
		b.setOpcode( LINE );
		b.setData( b.PROTOTYPE_BYTES );
		b.init();
		return b;
	}

	public void setAsStockChart()
	{
		chartType = ChartConstants.STOCKCHART;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0 };

	/**
	 * Handle setting options from XML in a generic manner
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "Stacked" ) )
		{
			fStacked = val.equals( "true" );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "PercentageDisplay" ) )
		{
			f100 = val.equals( "true" );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Shadow" ) )
		{
			fHasShadow = val.equals( "true" );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Smooth" ) )
		{    // smooth lines
/*            if (b instanceof Serfmt) {            	
            	Serfmt s= (Serfmt) b;
            	if (s.getSmoothLine()) {*/
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
		if( fStacked )
		{
			sb.append( " Stacked=\"true\"" );
		}
		if( f100 )
		{
			sb.append( " PercentageDisplay=\"true\"" );
		}
		if( fHasShadow )
		{
			sb.append( " Shadow=\"true\"" );
		}
		return sb.toString();
	}

	@Override
	public void setIsStacked( boolean bIsStacked )
	{
		fStacked = bIsStacked;
		grbit = ByteTools.updateGrBit( grbit, fStacked, 0 );
		updateRecord();
	}

	@Override
	public void setIs100Percent( boolean bOn )
	{
		f100 = bOn;
		grbit = ByteTools.updateGrBit( grbit, f100, 1 );
		updateRecord();
	}

	/**
	 * @return truth of "Chart is Stacked"
	 */
	@Override
	public boolean isStacked()
	{
		return fStacked;
	}
}
