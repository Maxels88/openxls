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
 * <b>Boppop: Bar of Pie/Pie of Pie chart options(0x1061)</b>
 * <p/>
 * <p/>
 * <p/>
 * <p/>
 * <p/>
 * pst (1 byte): An unsigned integer that specifies whether this chart group is a bar of pie chart group or a pie of pie chart group. MUST be a value from the following table:
 * 1= pie of pie 2= bar of pie
 * fAutoSplit (1 byte): A Boolean that specifies whether the split point of the chart group is determined automatically.
 * If the value is 1, when a bar of pie chart group or pie of pie chart group is initially created the data points from the primary pie are selected
 * and inserted into the secondary bar/pie automatically.
 * split (2 bytes): An unsigned integer that specifies what determines the split between the primary pie and the secondary bar/pie.
 * MUST be ignored if fAutoSplit is set to 1. MUST be a value from the following table:
 * <p/>
 * Value	 	Type of split		   	 Meaning
 * 0x0000	    Position			    The data is split based on the position of the data point in the series as specified by iSplitPos.
 * 0x0001	    Value				    The data is split based on a threshold value as specified by numSplitValue.
 * 0x0002	    Percent				    The data is split based on a percentage threshold and the data point values represented as a percentage as specified by pcSplitPercent.
 * 0x0003	    Custom				    The data is split as arranged by the user. Custom split is specified in a following BopPopCustom record.
 * <p/>
 * iSplitPos (2 bytes): A signed integer that specifies how many data points are contained in the secondary bar/pie.
 * Data points are contained in the secondary bar/pie starting from the end of the series.
 * For example, if the value is 2, the last 2 data points in the series are contained in the secondary bar/pie.
 * MUST be a value greater than or equal to 0 and less than or equal to 32000.
 * If the value is more than the number of data points in the series, the entire series will be in the secondary bar/pie, except for the first data point.
 * If split is not set to 0x0000 or fAutoSplit is set to 1, this value MUST be ignored.
 * <p/>
 * pcSplitPercent (2 bytes): A signed integer that specifies the percentage below which each data point is contained in the secondary bar/pie as opposed to the primary pie.
 * The percentage value of a data point is calculated using the following formula:
 * (value of the data point x 100) / sum of all data points in the series
 * If split is not set to 0x0002 or if fAutoSplit is set to 1, this value MUST be ignored
 * <p/>
 * pcPie2Size (2 bytes): A signed integer that specifies the size of the secondary bar/pie as a percentage of the size of the primary pie.
 * MUST be a value greater than or equal to 5 and less than or equal to 200.
 * <p/>
 * pcGap (2 bytes): A signed integer that specifies the distance between the primary pie and the secondary bar/pie.
 * The distance is specified as a percentage of the average width of the primary pie and secondary bar/pie.
 * MUST be a value greater than or equal to 0 and less than or equal to 500,
 * where 0 is 0% of the average width of the primary pie and the secondary bar/pie, and 500 is 250% of the average width of the primary pie and the secondary bar/pie.
 * <p/>
 * numSplitValue (8 bytes): An Xnum value that specifies the split when the split field is set to 0x0001
 * . The value of this field specifies the threshold that selects which data points of the primary pie move to the secondary bar/pie.
 * The secondary bar/pie contains any data points with a value less than the value of this field. If split is not set to 0x0001 or if fAutoSplit is set to 1, this value MUST be ignored.
 * <p/>
 * A - fHasShadow (1 bit): A bit that specifies whether one or more data points in the chart group have shadows.
 * <p/>
 * reserved (15 bits): MUST be zero, and MUST be ignored.
 */
public class Boppop extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8071801452993935943L;
	boolean fAutoSplit, fHasShadow;
	short split, iSplitPos, pcSplitPercent, pcPieSize, pcGap;
	float numSplitValue;
	byte pst;

	@Override
	public void init()
	{
		super.init();
		byte[] data = this.getData();
		pst = data[0];
		fAutoSplit = (data[1] == 1);
		split = ByteTools.readShort( data[2], data[3] );
		iSplitPos = ByteTools.readShort( data[4], data[5] );
		pcSplitPercent = ByteTools.readShort( data[6], data[7] );
		pcPieSize = ByteTools.readShort( data[8], data[9] );
		pcGap = ByteTools.readShort( data[10], data[11] );
		numSplitValue = (float) ByteTools.eightBytetoLEDouble( this.getBytesAt( 12, 8 ) );
		fHasShadow = (data[20] & 1) == 1;
		chartType = OFPIECHART;
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 1, 1, 0, 0, 0, 0, 0, 0, 75, 0, 100, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };

	public static XLSRecord getPrototype()
	{
		Boppop b = new Boppop();
		b.setOpcode( BOPPOP );
		b.setData( b.PROTOTYPE_BYTES );
		b.init();
		return b;
	}

	/**
	 * returns true of pie of pie, false if bar of pie
	 *
	 * @return
	 */
	public boolean isPieOfPie()
	{
		return (pst == 1);
	}

	public void setIsPieOfPie( boolean b )
	{
		if( b )
		{
			pst = 1;
		}
		else
		{
			pst = 2;
		}
		this.getData()[0] = pst;
	}

	/**
	 * specifies the distance between the primary pie and the secondary bar/pie.
	 * The distance is specified as a percentage of the average width of the primary pie and secondary bar/pie.
	 *
	 * @param g
	 */
	public void setpcGap( int g )
	{
		pcGap = (short) g;
		byte[] b = ByteTools.shortToLEBytes( pcGap );
		this.getData()[10] = b[0];
		this.getData()[11] = b[1];
	}

	/**
	 * returns the distance between the primary pie and the secondary bar/pie.
	 * The distance is specified as a percentage of the average width of the primary pie and secondary bar/pie.
	 */
	public int getpcGap()
	{
		return pcGap;
	}

	/**
	 * specifies the size of the secondary bar/pie as a percentage of the size of the primary pie.
	 *
	 * @param s
	 */
	public void setSecondPieSize( int s )
	{
		pcPieSize = (short) s;
		byte[] b = ByteTools.shortToLEBytes( pcPieSize );
		this.getData()[8] = b[0];
		this.getData()[9] = b[1];
	}

	/**
	 * returns the size of the secondary bar/pie as a percentage of the size of the primary pie.
	 */
	public int getSecondPieSize()
	{
		return pcPieSize;
	}

	/**
	 * specifies what determines the split between the primary pie and the secondary bar/pie.
	 * 0x0000	    Position			    The data is split based on the position of the data point in the series as specified by iSplitPos.
	 * 0x0001	    Value				    The data is split based on a threshold value as specified by numSplitValue.
	 * 0x0002	    Percent				    The data is split based on a percentage threshold and the data point values represented as a percentage as specified by pcSplitPercent.
	 * 0x0003	    Custom				    The data is split as arranged by the user. Custom split is specified in a following BopPopCustom record.
	 *
	 * @param t
	 */
	public void setSplitType( int t )
	{
		fAutoSplit = false;
		this.getData()[1] = 0;
		split = (short) t;
		byte[] b = ByteTools.shortToLEBytes( split );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];
	}

	/**
	 * specifies what determines the split between the primary pie and the secondary bar/pie via OOXML string value:
	 * auto	-- split point of the chart group is determined automatically
	 * pos --  The data is split based on the position of the data point in the series as specified by iSplitPos.
	 * val --  The data is split based on a threshold value as specified by numSplitValue.
	 * percent --The data is split based on a percentage threshold and the data point values represented as a percentage as specified by pcSplitPercent.
	 * cust -- The data is split as arranged by the user.
	 *
	 * @param s
	 */
	public void setSplitType( String s )
	{
		if( s == null )
		{
			fAutoSplit = true;
			this.getData()[1] = 0;
			return;
		}
		fAutoSplit = false;
		this.getData()[1] = 0;
		if( s.equals( "pos" ) )
		{
			split = 0;
		}
		else if( s.equals( "val" ) )
		{
			split = 1;
		}
		else if( s.equals( "percent" ) )
		{
			split = 2;
		}
		else if( s.equals( "cust" ) )
		{
			split = 3;
		}
		byte[] b = ByteTools.shortToLEBytes( split );
		this.getData()[2] = b[0];
		this.getData()[3] = b[1];
	}

	/**
	 * returns split type, or -1 if autosplit
	 * 0x0000	    Position			    The data is split based on the position of the data point in the series as specified by iSplitPos.
	 * 0x0001	    Value				    The data is split based on a threshold value as specified by numSplitValue.
	 * 0x0002	    Percent				    The data is split based on a percentage threshold and the data point values represented as a percentage as specified by pcSplitPercent.
	 * 0x0003	    Custom				    The data is split as arranged by the user. Custom split is specified in a following BopPopCustom record.
	 *
	 * @return
	 */
	public int getSplitType()
	{
		if( fAutoSplit )
		{
			return -1;
		}
		return split;
	}

	public String getSplitTypeOOXML()
	{
		if( fAutoSplit )
		{
			return "auto";
		}
		switch( split )
		{
			case 0:
				return "pos";
			case 1:
				return "val";
			case 2:
				return "percent";
			case 3:
				return "cust";
		}
		return "auto";    // default
	}

	/**
	 * specifies how many data points are contained in the secondary bar/pie.
	 * Data points are contained in the secondary bar/pie starting from the end of the series.
	 * For example, if the value is 2, the last 2 data points in the series are contained in the secondary bar/pie.
	 *
	 * @param sp
	 */
	public void setSplitPos( int sp )
	{
		iSplitPos = (short) sp;
		byte[] b = ByteTools.shortToLEBytes( iSplitPos );
		this.getData()[4] = b[0];
		this.getData()[5] = b[1];
	}

	/**
	 * returns how many data points are contained in the secondary bar/pie.
	 * Data points are contained in the secondary bar/pie starting from the end of the series.
	 * For example, if the value is 2, the last 2 data points in the series are contained in the secondary bar/pie.
	 */
	public int getSplitPos()
	{
		return iSplitPos;
	}

	/**
	 * Set specific options
	 */
	@Override
	public boolean setChartOption( String op, String val )
	{
		if( op.equalsIgnoreCase( "Gap" ) )
		{
			setpcGap( Integer.parseInt( val ) );
		}
		return true;
	}
}
