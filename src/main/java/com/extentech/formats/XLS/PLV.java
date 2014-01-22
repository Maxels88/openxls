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
 * PLV 0x88B
 * The PLV record specifies the settings of a Page Layout view for a sheet.
 * <p/>
 * frtHeader (12 bytes): An FrtHeader structure. The frtHeader.rt field MUST be 0x088B.
 * wScalePLV (2 bytes): An unsigned integer that specifies zoom scale as a percentage for the Page Layout view of the current sheet. For example, if the value is 107, then the zoom scale is 107%.
 * The value 0 means that the zoom scale is not set.
 * If the value is nonzero, it MUST be greater than or equal to 10 and less than or equal to 400.
 * A - fPageLayoutView (1 bit): A bit that specifies whether the sheet is in the Page Layout view. If the fSLV in Window2 record is 1 for this sheet, it MUST be 0.
 * B - fRulerVisible (1 bit): A bit that specifies whether the application displays the ruler.
 * C - fWhitespaceHidden (1 bit): A bit that specifies whether the margins between pages are hidden in the Page Layout view.
 * <p/>
 * unused (13 bits): Undefined, and MUST be ignored.
 */
public class PLV extends com.extentech.formats.XLS.XLSRecord
{
	short wScalePLV;

	public void init()
	{
		super.init();
		wScalePLV = ByteTools.readShort( this.getByteAt( 12 ), this.getByteAt( 13 ) );
	}
}
