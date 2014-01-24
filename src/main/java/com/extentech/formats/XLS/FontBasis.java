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
 * FBI: Font Basis (1060h)
 * <p/>
 * The FBI record stores font metrics.
 * Chart use only.
 * <p/>
 * Offset | Name | Size | Contents
 * <p/>
 * 4 | dmixBasis | 2 | Width of basis when font was applied
 * <p/>
 * 6 | dmiyBasis | 2 | Height of basis when font was applied
 * <p/>
 * 8 | twpHeightBasis | 2 | Font height applied
 * <p/>
 * 10 | scab | 2 | Scale basis
 * <p/>
 * 12 | ifnt | 2 | Index number into the font table
 */
public class FontBasis extends XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6984935185426785077L;

	@Override
	public void init()
	{
		super.init();
	}

	public int getFontIndex()
	{
		return ByteTools.readShort( getData()[8], getData()[9] );
	}

	public void setFontIndex( int id )
	{
		byte[] b = ByteTools.shortToLEBytes( (short) id );
		getData()[8] = b[0];
		getData()[9] = b[1];
	}

}
