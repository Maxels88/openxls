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
 * <pre><b>1904: Flag which determines whether the 1904 date system is used 0x22</b><br>
 * <p/>
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       f1904      		2       0x1 = 1904 in the house
 *
 * </pre>
 */

public final class NineteenOhFour extends com.extentech.formats.XLS.XLSRecord
{
	// if i exist, then you are in 1904.

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -4258740446375241600L;
	boolean is1904 = false;

	@Override
	public void init()
	{
		super.init();
		is1904 = (ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) ) == 1);
	}

}