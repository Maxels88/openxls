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
 * Dsf Double Stream File (161h)
 * <p/>
 * simple flag indicating whether this is a double-stream file
 */
public class Dsf extends com.extentech.formats.XLS.XLSRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4818410956902736572L;
	int fDSF = -1;

	/**
	 * all it does is indicate DSF-ness
	 */
	@Override
	public void init()
	{
		super.init();
		fDSF = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
	}
}