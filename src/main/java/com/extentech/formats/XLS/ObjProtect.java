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
 * <b>OBJPROTECT: Protection Flag (63h)</b><br>
 * <p/>
 * OBJPROTECT stores the protection state for a sheet or workbook
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       fLock           2       = 1 if the objects are protected
 * </p></pre>
 */
public final class ObjProtect extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -5775187385375827918L;
	int fLock = -1;

	/**
	 * returns whether this sheet or workbook is protected
	 */
	boolean getIsLocked()
	{
		if( fLock > 0 )
		{
			return true;
		}
		return false;
	}

	/**
	 * default constructor
	 */
	ObjProtect()
	{
		super();
		byte[] bs = new byte[2];
		bs[0] = 1;
		bs[1] = 0;
		setOpcode( OBJPROTECT );
		setLength( (short) 2 );
		//   setLabel("OBJPROTECT" + String.valueOf(this.offset));
		this.setData( bs );
		this.originalsize = 2;
	}

	void setLocked( boolean b )
	{
		byte[] data = this.getData();
		if( b )
		{
			data[0] = 1;
		}
		else
		{
			data[0] = 0;
		}
	}

	@Override
	public void init()
	{
		super.init();
		fLock = (int) ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		if( this.getIsLocked() && (DEBUGLEVEL > DEBUG_LOW) )
		{
			Logger.logInfo( "Object Protection Enabled." );
		}
	}

}