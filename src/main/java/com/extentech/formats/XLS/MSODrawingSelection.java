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
 * <b>Msodrawingselection: MS Office Drawing Selection (EDh)</b><br>
 * <p/>
 * These records contain only data.<p><pre>
 * <p/>
 * offset  name        size    contents
 * ---
 * 4       rgMSODrSelr    var    MSO Drawing Group selectionData
 * <p/>
 * </p></pre>
 *
 * @see MSODrawing
 */
public class MSODrawingSelection extends com.extentech.formats.XLS.XLSRecord
{

	/**
	 *
	 */
	private static final long serialVersionUID = 2799490308252319737L;
	public byte[] PROTOTYPE_BYTES = { 0, 0, 25, -15, 16, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 4, 0, 0, 1, 4, 0, 0 };

	/**
	 * not a lot going on here...
	 */
	public void init()
	{
		super.init();
	}

	/**
	 * bypass continue handling for msodrawingselection, until we start
	 * modifing the record
	 */
	public byte[] getData()
	{
		return super.getData();
	}

}
