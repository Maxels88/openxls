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

/**
 * CF 12 is a FRT (Future record type) record that handles extended conditional formatting information that
 * was not available before excel 2007.  It is currently unsupported in ExtenXLS, but the record looks much the same
 * to a standard CF rule
 * <p/>
 * I'm guessing when implemented we will either want a member cf record or even just extend cf with this.
 */

public class Cf12 extends XLSRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4974997074621442348L;

}
