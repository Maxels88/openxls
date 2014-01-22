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
/**
 *
 */
package com.extentech.formats.XLS;

import com.extentech.toolkit.ByteTools;

/**
 * <b>XCT  CRN Count (0059h)</b><br>
 * <p/>
 * <p><pre>
 * 	This record stores the number of immediately following Crn records.
 * 	These records are used to store the cell contents of external references.
 * <p/>
 * offset  size 	contents
 * ---
 * 0 		2 		Number of following CRN records
 * 2 		2 		Index into sheet table of the involved SUPBOOK record
 * </p></pre>
 */

public class Xct extends XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -5112701341711255711L;
	private int nCRNs;    // number of External Cell References CRN record, similar to EXTERNNAME
	private int supBookIndex;

	public void init()
	{
		super.init();
		nCRNs = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		supBookIndex = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
	}

	public String toString()
	{
		return "XTC: nCRNS=" + nCRNs + " SupBook Index: " + supBookIndex;
	}
}
