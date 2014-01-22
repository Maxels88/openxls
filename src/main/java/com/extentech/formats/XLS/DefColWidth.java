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
 * <b>DefColWidth: default column width for WorkBook (55h)</b><br>
 * <p/>
 * offset  name        size    contents
 * ---
 * 6       miyRw       2       Default Column Width
 * <p/>
 * <p/>
 * </p></pre>
 *
 * @see DefaultRowHeight
 */
public final class DefColWidth extends com.extentech.formats.XLS.XLSRecord
{
	short defWid;
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8726286841723548636L;

	public void init()
	{
		byte[] mydata = this.getData();
		defWid = ByteTools.readShort( mydata[0], mydata[1] );
	}

	public void setDefaultColWidth( int t )
	{
		this.defWid = (short) t;
		byte[] mydata = this.getData();
		byte[] heightbytes = ByteTools.shortToLEBytes( (short) t );
		mydata[0] = heightbytes[0];
		mydata[1] = heightbytes[1];
	}

	public short getDefaultWidth()
	{
		return defWid;
	}
}