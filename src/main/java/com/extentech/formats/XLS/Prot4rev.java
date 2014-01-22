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
 * <b>PROT4REV: Protection Flag (12h)</b><br>
 * <p/>
 * PROT4REV stores the protection state for a shared workbook
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       fRevLock        2       = 1 if the Sharing with
 * track changes option is on
 * </p></pre>
 */
public final class Prot4rev extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 681662633243537043L;
	int fRevLock = -1;

	/**
	 * returns whether this sheet or workbook is protected
	 */
	boolean getIsLocked()
	{
		if( fRevLock > 0 )
		{
			return true;
		}
		return false;
	}

	public void init()
	{
		super.init();
		fRevLock = (int) ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		if( this.getIsLocked() )
		{
			Logger.logInfo( "Shared Workbook Protection Enabled." );
			// throw new InvalidFileException("Shared Workbook Protection Enabled.  Unsupported file format.");
		}
	}

}