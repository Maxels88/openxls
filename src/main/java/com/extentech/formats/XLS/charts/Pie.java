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
 * <b>Pie: Chart Group Is a pie Chart Group(0x1019)</b>
 * <p/>
 * 4	anStart		2		Angle of the first pie slice expressed in degrees.  Must be <= 360
 * 6	pcDonut		2		0= true pie chart, non-zero= size of center hole in a donut chart, as a percentage
 * 8	grbit		2		Option Flags
 * <p/>
 * <p/>
 * grbit:
 * 0	0x1		fHasShadow		1= has shadow
 * 1	0x2		fShowLdrLines	1= show leader lines to data labels
 */
public class Pie extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -7851320124576950635L;
	private short grbit = 0;
	private boolean fHasShadow = false;
	private boolean fShowLdrLines = false;
	protected short pcDonut = 0;
	protected short anStart = 0;

	@Override
	public void init()
	{
		super.init();
		byte[] data = getData();
		anStart = ByteTools.readShort( data[0], data[1] );
		pcDonut = data[2];
		if( pcDonut == 0 )
		{
			chartType = ChartConstants.PIECHART; // 20070703 KSC
		}
		else
		{
			chartType = ChartConstants.DOUGHNUTCHART;
		}
		grbit = ByteTools.readShort( data[4], data[5] );
		fHasShadow = ((grbit & 0x1) == 0x1);
		fShowLdrLines = ((grbit & 0x2) == 0x2);
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 0, 0, 0, 0, 2, 0 };

	public static XLSRecord getPrototype()
	{
		Pie b = new Pie();
		b.setOpcode( PIE );
		b.setData( b.PROTOTYPE_BYTES );
		b.init();
		return b;
	}

	// 20070703 KSC and below ...
	private void updateRecord()
	{
		getData()[2] = (byte) pcDonut;
		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[4] = b[0];
		getData()[5] = b[1];
	}

	public void setAsPieChart()
	{
		pcDonut = 0;
		chartType = ChartConstants.PIECHART;
		updateRecord();
	}

	public void setAsDoughnutChart()
	{
		pcDonut = 50;    // default %
		chartType = ChartConstants.DOUGHNUTCHART;
		updateRecord();
	}

	/**
	 * return the Doughnut hole size (if > 0)
	 *
	 * @return
	 */
	public int getDoughnutSize()
	{
		return pcDonut;
	}

	/**
	 * size of center hole in a donut chart, as a percentage.  0 for pie
	 *
	 * @return
	 */
	public void setDoughnutSize( int s )
	{
		pcDonut = (short) s;
		byte[] b = ByteTools.shortToLEBytes( pcDonut );
		getData()[2] = b[0];
		getData()[3] = b[1];
	}

	/**
	 * return the Angle of the first pie slice expressed in degrees.  Must be <= 360
	 *
	 * @return
	 */
	public int getAnStart()
	{
		return anStart;
	}

	/**
	 * sets the Angle of the first pie slice expressed in degrees.  Must be <= 360
	 *
	 * @return
	 */
	public void setAnStart( int a )
	{
		anStart = (short) a;
		byte[] b = ByteTools.shortToLEBytes( anStart );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	/**
	 * @return String XML representation of this chart-type's options
	 */
	@Override
	public String getOptionsXML()
	{
		StringBuffer sb = new StringBuffer();
		if( fHasShadow )
		{
			sb.append( " Shadow=\"true\"" );
		}
		if( fShowLdrLines )
		{
			sb.append( " ShowLdrLines=\"true\"" );
		}
		if( pcDonut > 0 )
		{
			sb.append( " Donut=\"" + pcDonut + "\"" );
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
		if( op.equalsIgnoreCase( "Shadow" ) )
		{
			setHasShadow( true );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "ShowLdrLines" ) )
		{
			setShowLdrLines( true );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "Donut" ) )
		{
			setDonutPercentage( val );
			bHandled = true;
		}
		else if( op.equalsIgnoreCase( "donutSize" ) )
		{
			setDonutPercentage( val );
			bHandled = true;
		}
		return bHandled;
	}

	@Override
	public String getChartOption( String op )
	{
		if( op.equals( "ShowLdrLines" ) )
		{ // Pie
			return String.valueOf( showLdrLines() );
		}
		if( op.equals( "donutSize" ) )
		{ // Pie
			return String.valueOf( getDonutPercentage() );
		}
		return super.getChartOption( op );
	}

	@Override
	public boolean hasShadow()
	{
		return fHasShadow;
	}

	public boolean showLdrLines()
	{
		return fShowLdrLines;
	}

	public void setHasShadow( boolean bHasShadow )
	{
		fHasShadow = bHasShadow;
		grbit = ByteTools.updateGrBit( grbit, fHasShadow, 0 );
		updateRecord();
	}

	public void setShowLdrLines( boolean bShowLdrLines )
	{
		fShowLdrLines = bShowLdrLines;
		grbit = ByteTools.updateGrBit( grbit, fShowLdrLines, 1 );
		updateRecord();
	}

	public void setDonutPercentage( String val )
	{
		try
		{
			pcDonut = Short.valueOf( val );
		}
		catch( Exception e )
		{
		}
		updateRecord();
	}

	public short getDonutPercentage()
	{
		return pcDonut;
	}

}
