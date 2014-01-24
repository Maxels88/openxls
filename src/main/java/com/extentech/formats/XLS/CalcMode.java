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

/**
 * <b>CALCMODE: (OxD)</b><br>
 * <p/>
 * It specifies whether to calculate formulas manually,
 * automatically or automatically except for multiple table operations.
 * <p><pre>
 * Offset Size Contents
 * 0 		2 	FFFFH = automatically except for multiple table operations
 * 0000H = manually
 * 0001H = automatically (default)
 * </p></pre>
 */

public final class CalcMode extends com.extentech.formats.XLS.XLSRecord
{

	private static final long serialVersionUID = -4544323710670598072L;
	short calcmode;

	@Override
	public void init()
	{
		super.init();
		/*  FFFFH = automatically except for multiple table operations
			0000H = manually
			0001H = automatically (default)
         */
		calcmode = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
	}

	/**
	 * returns the recalculation mode:
	 * 0= Manual, 1= Automatic, 2= Automatic except for Multiple Table Operations
	 *
	 * @return int recalculation mode
	 */
	public int getRecalcuationMode()
	{
		if( calcmode < 0 )
		{
			return 2;
		}
		return calcmode;
	}

	/**
	 * Sets the recalculation mode for the Workbook:
	 * <br>0= Manual
	 * <br>1= Automatic
	 * <br>2= Automatic except for multiple table operations
	 */
	public void setRecalculationMode( int mode )
	{
		if( (mode >= 0) && (mode <= 2) )
		{
			if( mode == 2 )
			{
				mode = -1;
			}
			calcmode = (short) mode;
			byte[] b = ByteTools.shortToLEBytes( calcmode );
			getData()[0] = b[0];
			getData()[1] = b[1];
		}
	}

}