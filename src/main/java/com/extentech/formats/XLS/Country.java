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
 * <b>COUNTRY: 8CH</b><br>
 * <p/>
 * This record stores two Windows country identifiers. The first represents the user interface language of the Excel version
 * that has saved the file, and the second represents the system regional settings at the time the file was saved.
 * Record COUNTRY, BIFF3-BIFF8:
 * <p><pre>
 * Offset Size Contents
 * 0 		2 	Windows country identifier of the user interface language of Excel
 * 2 		2 	Windows country identifier of the system regional settings
 * </p></pre>
 */

public final class Country extends com.extentech.formats.XLS.XLSRecord
{

	private static final long serialVersionUID = -4544323710670598072L;

	@Override
	public void init()
	{
		super.init();
		getData();
	}

	public int getDefaultLanguage()
	{
		return ByteTools.readShort( getData()[0], getData()[1] );
	}
}