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
 * <b>Area: Chart Group is an Area Chart Group (0x101a)</b>
 * 4	grbit		2	formatflags
 * <p/>
 * grbit:
 * 0	0	01h	fStacked	Series in this group are stacked
 * 1	02h	f100		Each cat is broken down as a percentge
 * 2	04h	fHasShadow	1= this are has a shadow
 * 1	7-0 FFh reserved	0
 */
public class Area extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -4600344312324775780L;
	private short grbit = 0;
	protected boolean fStacked = false, f100 = false, fHasShadow = false;

	@Override
	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		fStacked = ((grbit & 0x1) == 0x1);
		f100 = ((grbit & 0x2) == 0x2);
		fHasShadow = ((grbit & 0x4) == 0x4);
		chartType = ChartConstants.AREACHART;    // 20070703 KSC
	}

	protected void updateRecord()
	{
		grbit = ByteTools.updateGrBit( grbit, fStacked, 0 );
		grbit = ByteTools.updateGrBit( grbit, f100, 1 );
		grbit = ByteTools.updateGrBit( grbit, fHasShadow, 2 );
		byte[] b = ByteTools.shortToLEBytes( grbit );
		this.getData()[0] = b[0];
		this.getData()[1] = b[1];
	}

	public static XLSRecord getPrototype()
	{
		Area b = new Area();
		b.setOpcode( AREA );
		b.setData( b.PROTOTYPE_BYTES );
		b.init();
		return b;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0 };

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

	@Override
	public boolean is100Percent()
	{
		return f100;
	}

	/**
	 * Handle setting options
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		boolean bHandled = false;
		if( op.equalsIgnoreCase( "Stacked" ) )
		{
			this.fStacked = val.equals( "true" );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "PercentageDisplay" ) )
		{
			this.f100 = val.equals( "true" );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Shadow" ) )
		{
			this.fHasShadow = val.equals( "true" );
			bHandled = true;
		}
		if( bHandled )
		{
			this.updateRecord();
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
		if( this.fStacked )
		{
			sb.append( " Stacked=\"true\"" );
		}
		if( this.f100 )
		{
			sb.append( " PercentageDisplay=\"true\"" );
		}
		if( this.fHasShadow )
		{
			sb.append( " Shadow=\"true\"" );
		}
		return sb.toString();
	}

}
