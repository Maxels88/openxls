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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <b>SCENPROTECT: Protection Flag (DDh)</b><br>
 * <p/>
 * SCENPROTECT stores the protection state for a sheet or workbook
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       fLock           2       = 1 if the scenarios are protected
 * </p></pre>
 */

public final class ScenProtect extends com.extentech.formats.XLS.XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( ScenProtect.class );
	private static final long serialVersionUID = -3722344748446193860L;
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
	ScenProtect()
	{
		super();
		byte[] bs = new byte[2];
		bs[0] = 1;
		bs[1] = 0;
		setOpcode( SCENPROTECT );
		setLength( (short) 2 );
		//  setLabel("SCENPROTECT" + String.valueOf(this.offset));
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
		fLock = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		if( this.getIsLocked()  )
		{
			log.debug( "Scenario Protection Enabled." );
		}
	}
}