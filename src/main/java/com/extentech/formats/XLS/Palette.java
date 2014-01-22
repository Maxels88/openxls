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
import com.extentech.toolkit.CompatibleVector;

/**
 * <b>Palette: Defined Colors (92h)</b><br>
 * <p/>
 * Describe Colors in the file.
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 4       ccv             2       count of  colors
 * var     rgch            var     4 byte color data
 * <p/>
 * </p></pre>
 */
public final class Palette extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 157670739931392705L;
	CompatibleVector colorvect = new CompatibleVector();
	int ccv = -1;

	@Override
	public void init()
	{
		super.init();
		ccv = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		this.getData();
		int pos = 2;
		for( int d = 0; d < ccv; d++ )
		{
			this.getColorTable()[d + 8] = new java.awt.Color( (data[pos] < 0 ? 255 + data[pos] : data[pos]),
			                                                  (data[pos + 1] < 0 ? 255 + data[pos + 1] : data[pos + 1]),
			                                                  (data[pos + 2] < 0 ? 255 + data[pos + 2] : data[pos + 2]) );
			pos += 4;
		}
	}

}