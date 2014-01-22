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

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

/**
 * <b>Scl: Sheet Zoom (A0h)</b><br>
 * <p/>
 * Scl stores the zoom magnification for the sheet
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       num             2       = Numerator of the view magnification fraction (num)
 * 6		denum			2		= Denumerator of the view magnification fraction (den)
 * </p></pre>
 */

public final class Scl extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 *
	 */
	private static final long serialVersionUID = -4595833226859365049L;
	//	int num = 100; 20081231 KSC: default val is 1, making the calc (num/denum)*100
	int num = 1;
	int denum = 1;

	/**
	 * default constructor
	 */
	Scl()
	{
		super();
		byte[] bs = new byte[4];
		bs[0] = 1;
		bs[1] = 0;
		bs[2] = 1;
		bs[3] = 0;
		setOpcode( SCL );
		setLength( (short) 4 );
		if( DEBUGLEVEL > DEBUG_LOW )
		{
			Logger.logInfo( "Scl.init()" + String.valueOf( this.offset ) );
		}
		this.setData( bs );
		this.originalsize = 4;
	}

	/**
	 * sets the zoom as a percentage for this sheet
	 *
	 * @param b
	 */
	public void setZoom( float b )
	{
		byte[] data = this.getData();

/* 20081231 KSC:  appears that zooming is such that 1/1=100%         
        // set our scale to 1000
        denum = 1000;
        byte[] denmbd= ByteTools.shortToLEBytes((short) denum);
        System.arraycopy(denmbd, 0, data, 2, 2);

        // take something like .2345 and come up with 24 & 100
        float nx = b * denum; // get denum
        // get the num
        num = (int)nx;

        if((denum % b)>0){
        	if(b>999) // only 2 precision places for zoom... out a warn
        		Logger.logWarn("Cannot set zoom to : " +b + " rounding to nearest valid zoom setting.");
        }
*/
		// 20081231 KSC: Convert double to fraction and set num/denum to results
		int[] n = gcd( (int) (b * 100), 100 );
		num = n[0];
		denum = n[1];
		byte[] nmbd = ByteTools.shortToLEBytes( (short) num );
		System.arraycopy( nmbd, 0, data, 0, 2 );
		nmbd = ByteTools.shortToLEBytes( (short) denum );
		System.arraycopy( nmbd, 0, data, 2, 2 );

		this.setData( data );
	}

	/**
	 * gets the zoom as a percentage for this sheet
	 *
	 * @return
	 */
	public float getZoom()
	{
		return ((float) num / (float) denum);
	}

	@Override
	public void init()
	{
		super.init();
		num = (int) ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		denum = (int) ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
		if( (DEBUGLEVEL > DEBUG_LOW) )
		{
			Logger.logInfo( "Scl.init() sheet zoom:" + getZoom() );
		}
	}

	private int[] gcd( int numerator, int denominator )
	{
		int highest;
		int n = 1;
		int d = 1;

		if( denominator > numerator )
		{
			highest = denominator;
		}
		else
		{
			highest = numerator;
		}

		for( int x = highest; x > 0; x-- )
		{
			if( denominator % x == 0 && numerator % x == 0 )
			{
				n = numerator / x;
				d = denominator / x;
				break;
			}
		}
		return new int[]{ n, d };
	}
}