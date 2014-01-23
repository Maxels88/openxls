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
 * <b>USERSVIEWEND: End Custom View Settings (1ABh)</b><br>
 * <p/>
 * USERSVIEWEND marks the end of a custom view for the sheet
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       fValid          2       = 1 if the settings saved are valid
 * <p/>
 * </p></pre>
 */
public final class Usersviewend extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -3120417369791123931L;
	int fValid = -1;

	/**
	 * return whether the settings
	 * for this user view are valid
	 */
	public boolean isValid()
	{
		if( fValid == 1 )
		{
			return true;
		}
		return false;
	}

	@Override
	public void init()
	{
		super.init();
		fValid = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
	}

}