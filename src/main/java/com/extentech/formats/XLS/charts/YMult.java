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
 * <b>YMULT: Y Multiplier (857h)</b>
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for Charts.
 * This record describes the axis multiplier feature which scales the axis values
 * displayed by the axis tick labels. For instance, an axis multiplier value of
 * "millions" would cause the axis tick labels to show the axis value divided by one
 * million (e.g., the tick label for an axis value of 20,000,000 would show "20".)
 * This record is a "parent" record and is immediately followed by a set of records
 * surrounded by rtStartObject and rtEndObject which describes the axis multiplier label.
 * <p/>
 * Record Data
 * Offset	Field Name	Size	Contents
 * 4		rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =0857h
 * 6		grbitFrt	2		FRT flags; must be zero
 * 8		axmid		2		Axis multiplier ID, one of the following values:
 * -1 = multiplier value is stored in numLabelMultiplier
 * 0 = no multiplier (same as 1.0)
 * 1 = Hundreds, 10 2nd
 * 2 = Thousands, 10 3rd
 * 3 = Ten Thousands, 10 4th
 * 4 = Hundred Thousands, 10 5th
 * 5 = Millions, 10 6th
 * 6 = Ten Millions, 10 7th
 * 7 = Hundred Millions, 10 8th
 * 8 = billion
 * 9 = trillion
 * 16		numLabelMultiplier	4	Numeric value
 * 18		grbit		2		Option flags for y axis multiplier (see description below)*
 * <p/>
 * The grbit field contains the following category axis label option flags:
 * Bits	Mask	Flag Name	Contents
 * 0		0001h	fEnabled	=1 if the multiplier is enabled =0 otherwise
 * 1		0002h	fAutoShowMultiplier	=1 if the multiplier label is shown =0 otherwise
 * 15-2	FFFCh	(unused)	Reserved; must be zero
 */
public class YMult extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */

	private static final long serialVersionUID = -6166267220292885486L;
	short axmid;
	short grbit;
	double numLabelMultiplier;

	@Override
	public void init()
	{
		super.init();
		axmid = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		numLabelMultiplier = ByteTools.eightBytetoLEDouble( getBytesAt( 6, 8 ) );
		grbit = ByteTools.readShort( getByteAt( 14 ), getByteAt( 15 ) );
	}

	// TODO: Prototype Bytes
	private byte[] PROTOTYPE_BYTES = new byte[]{ };

	public static XLSRecord getPrototype()
	{
		YMult ym = new YMult();
		ym.setOpcode( YMULT );
		ym.setData( ym.PROTOTYPE_BYTES );
		ym.init();
		return ym;
	}

	/**
	 * returns the Axis multiplier ID, one of the following values:
	 * <li>-1 = multiplier value is stored in numLabelMultiplier
	 * <li>0 = no multiplier (same as 1.0)
	 * <li>1 = Hundreds, 10 2nd
	 * <li>2 = Thousands, 10 3rd
	 * <li>3 = Ten Thousands, 10 4th
	 * <li>4 = Hundred Thousands, 10 5th
	 * <li>5 = Millions, 10 6th
	 * <li>6 = Ten Millions, 10 7th
	 * <li>7 = Hundred Millions, 10 8th
	 * <li>8 = Thousand Millions, 10 9th
	 * <li>9 = Billions, 10 12th
	 *
	 * @return
	 */
	public short getAxMultiplierId()
	{
		return axmid;
	}

	public String getAxMultiplierIdAsString()
	{
		switch( axmid )
		{
			case -1:
				return null;
			case 0:
				return null;    //?
			case 1:
				return "hundreds";
			case 2:
				return "thousands";
			case 3:
				return "tenThousands";
			case 4:
				return "hundredThousands";
			case 5:
				return "millions";
			case 6:
				return "tenMillions";
			case 7:
				return "hundredMillions";
			case 8:
				return "billions";
			case 9:
				return "trillions";
		}
		return null;
	}

	/**
	 * Sets Axis multiplier ID, one of the following values:
	 * <li>-1 = multiplier value is stored in numLabelMultiplier
	 * <li>0 = no multiplier (same as 1.0)
	 * <li>1 = Hundreds, 10 2nd
	 * <li>2 = Thousands, 10 3rd
	 * <li>3 = Ten Thousands, 10 4th
	 * <li>4 = Hundred Thousands, 10 5th
	 * <li>5 = Millions, 10 6th
	 * <li>6 = Ten Millions, 10 7th
	 * <li>7 = Hundred Millions, 10 8th
	 * <li>8 = Thousand Millions, 10 9th
	 * <li>9 = Billions, 10 12th
	 *
	 * @param m
	 */
	public void setAxMultiplierId( int m )
	{
		if( !((m > -2) && (m < 10)) )
		{
			return;    // report error?
		}
		axmid = (short) m;
		byte[] b = ByteTools.shortToLEBytes( axmid );
		getData()[4] = b[0];
		getData()[5] = b[1];
	}

	/**
	 * Sets Axis multiplier ID via OOXML String value:
	 * <li>hundreds			Hundreds
	 * <li>thousands			Thousands
	 * <li>tenThousands		Ten Thousands
	 * <li>hundredThousands	Hundred Thousands
	 * <li>millions			Millions
	 * <li>tenMillions			Ten Millions
	 * <li>hundredMillions		Hundred Millions
	 * <li>billions			Billions
	 * <li>trillions			Trillions
	 */
	public void setAxMultiplierId( String m )
	{
		if( m.equalsIgnoreCase( "hundreds" ) )
		{
			axmid = 1;
		}
		else if( m.equalsIgnoreCase( "thousands" ) )
		{
			axmid = 2;
		}
		else if( m.equalsIgnoreCase( "tenThousands" ) )
		{
			axmid = 3;
		}
		else if( m.equalsIgnoreCase( "hundredThousands" ) )
		{
			axmid = 4;
		}
		else if( m.equalsIgnoreCase( "millions" ) )
		{
			axmid = 5;
		}
		else if( m.equalsIgnoreCase( "tenMillions" ) )
		{
			axmid = 6;
		}
		else if( m.equalsIgnoreCase( "hundredMillions" ) )
		{
			axmid = 7;
		}
		else if( m.equalsIgnoreCase( "billions" ) )
		{
			axmid = 8;
		}
		else if( m.equalsIgnoreCase( "trillions" ) )
		{
			axmid = 9;
		}
		else    // default
		{
			axmid = 0;
		}
		byte[] b = ByteTools.shortToLEBytes( axmid );
		getData()[4] = b[0];
		getData()[5] = b[1];
	}

	// dispUnits -> builtInUnit (val:  billions,
	public double getCustomMultiplier()
	{
		return numLabelMultiplier;
	}

	public void setCustomMultiplier( double m )
	{
		numLabelMultiplier = m;
		axmid = -1;    // custom
		byte[] b = ByteTools.shortToLEBytes( axmid );
		getData()[4] = b[0];
		getData()[5] = b[1];
		b = ByteTools.doubleToLEByteArray( numLabelMultiplier );
		System.arraycopy( b, 0, getData(), 6, 8 );
	}
}
