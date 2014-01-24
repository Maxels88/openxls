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
package com.extentech.formats.XLS;

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <b>NUMBER: BiffRec Value, Floating-Point Number (203h)</b><br>
 * This record stores an internal numeric type.  Stores data in one of four
 * RK 'types' which determine whether it is an integer or an IEEE floating point
 * equivalent.
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       rw          2       Row Number
 * 6       col         2       Column Number of the RK record
 * 8       ixfe        2       Index to XF cell format record
 * 10      num         8       Floating point number
 * </p></pre>
 *
 * @see RK
 * @see MULRK
 */

public final class NumberRec extends XLSCellRecord
{
	private static final Logger log = LoggerFactory.getLogger( NumberRec.class );
	private static final long serialVersionUID = 7489308348300854345L;
	//int t;
	double fpnum;

	/**
	 * Constructor which takes an Integer value
	 */
	public NumberRec( int val )
	{
		super();
		setOpcode( NUMBER );
		setLength( (short) 14 );
		//	setLabel("NUMBER");
		setData( new byte[14] );
		originalsize = 14;
		setNumberVal( val );
		isIntNumber = true;
		isFPNumber = false;
	}

	/**
	 * Constructor which takes a number value
	 */
	public NumberRec( long val )
	{
		this( (double) val );
	}

	/**
	 * Constructor which takes a number value
	 */
	public NumberRec( double val )
	{
		super();
		setOpcode( NUMBER );
		setLength( (short) 14 );
		//  setLabel("NUMBER");
		setData( new byte[14] );
		originalsize = 14;
		setNumberVal( val );
	}

	@Override
	public void init()
	{
		super.init();
		int l = 0;
		int m = 0;
		// get the row information
		super.initRowCol();
		short s = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		ixfe = s;
		// get the long
//      get the long
		fpnum = ByteTools.eightBytetoLEDouble( getBytesAt( 6, 8 ) );
		setIsValueForCell( true );
			log.trace( "NumberRec: " + getCellAddress() + ":" + getStringVal() );

		String d = String.valueOf( fpnum );
		if( d.length() > 12 )
		{
			isDoubleNumber = true;
			isFPNumber = true;
			isIntNumber = false;
		}
		else if( d.substring( d.length() - 2, d.length() ).equals( ".0" ) && (fpnum < Integer.MAX_VALUE) )
		{
			// this is for extenXLS output files, as we put int's into number records!
			isIntNumber = true;
			isFPNumber = false;
			isDoubleNumber = false;
		}
		else
		{
			if( (fpnum < Float.MAX_VALUE) || ((fpnum * -1) < Float.MAX_VALUE) )
			{
				isFPNumber = true;
				isIntNumber = false;
			}
			else
			{
				// isFPNumber=true;
				isDoubleNumber = true;
				isIntNumber = false;
			}
		}
	}

	/**
	 * Get the value of the record as a Float.
	 * Value must be parseable as an Float or it
	 * will throw a NumberFormatException.
	 */
	@Override
	public float getFloatVal()
	{
		return (float) fpnum;
	}

	@Override
	public int getIntVal()
	{
		if( fpnum > Integer.MAX_VALUE )
		{
			throw new NumberFormatException( "Cell value is larger than the maximum java signed int size" );
		}
		if( fpnum < Integer.MIN_VALUE )
		{
			throw new NumberFormatException( "Cell value is smaller than the minimum java signed int size" );
		}
		return (int) fpnum;
	}

	@Override
	public double getDblVal()
	{
		return fpnum;
	}

	@Override
	public String getStringVal()
	{
		if( isIntNumber )
		{
			return String.valueOf( (int) fpnum );
		}
		return ExcelTools.getNumberAsString( fpnum );
	}

	@Override
	public void setDoubleVal( double v )
	{
		setNumberVal( v );
	}

	@Override
	public void setFloatVal( float d )
	{
//    	setNumberVal(d);	// original
		// 20090708 KSC: handle casting issues by converting float to string first
		setNumberVal( new Double( (new Float( d )).toString() ) );
	}

	public NumberRec()
	{
		super();
	}

	@Override
	public void setIntVal( int i )
	{
		double d = i;
		setNumberVal( d );
	}

	void setNumberVal( long d )
	{
		byte[] b;
		byte[] rkdata = getData();
		b = ByteTools.toBEByteArray( d );
		System.arraycopy( b, 0, rkdata, 6, 8 );
		setData( rkdata );
		init();
	}

	void setNumberVal( double d )
	{
		byte[] b;
		byte[] rkdata = getData();
		b = ByteTools.toBEByteArray( d );
		System.arraycopy( b, 0, rkdata, 6, 8 );
		setData( rkdata );
		init();
	}
}