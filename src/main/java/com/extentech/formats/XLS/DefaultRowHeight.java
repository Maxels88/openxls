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
 * <b>DefaultRowHeight: default row height for WorkBook (225h)</b><br>
 * <p/>
 * offset  name        size    contents
 * ---
 * 4       options     2       comments
 * 6       miyRw       2       Default height for unused rows, in twips = 1/20 of a point
 * <p/>
 * ﻿	0 2 	2 2 	Option flags:
 * Bit	Mask Contents
 * 0 0001H 1 = Row height and default font height do not match 1
 * 0002H 1 = Row is hidden 2
 * 0004H 1 = Additional space above the row 3
 * 0008H 1 = Additional space below the row
 * <p/>
 * <p/>
 * <p/>
 * </p></pre>
 *
 * @see DefColWidth
 */
public final class DefaultRowHeight extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -930064032441287284L;
	short rwh;

	/**
	 * init: save the default row height
	 */
	@Override
	public void init()
	{
		super.init();
		rwh = ByteTools.readShort( this.getData()[2], this.getData()[3] );
	}

	@Override
	public void setWorkBook( WorkBook b )
	{
		super.setWorkBook( b );
		b.setDefaultRowHeightRec( this );
	}

	/**
	 * set the default row height in twips (=1/20 of a point)
	 * <br>Twips are 20*Excel units
	 * <br>e.g. default row height in Excel units=12.75
	 * <br>20*12.75= 256 (approx) twips
	 *
	 * @param t - desired default row height in twips
	 */
	public void setDefaultRowHeight( int t )
	{
		this.rwh = (short) t;
		byte[] mydata = this.getData();
		byte[] heightbytes = ByteTools.shortToLEBytes( this.rwh );
		mydata[2] = heightbytes[0];
		mydata[3] = heightbytes[1];
	}

	/**
	 * set the sheet's default row height in Excel units or twips
	 */
	@Override
	public void setSheet( Sheet bs )
	{
		this.worksheet = bs;
		((Boundsheet) bs).setDefaultRowHeight( this.rwh / 20.0 );
	}

}