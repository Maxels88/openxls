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
package com.extentech.formats.XLS.charts;

import com.extentech.formats.XLS.XLSRecord;

/**
 * <b>STARTOBJECT: Chart Future Record Type Start Object (854h)</b>
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for Charts.
 * The STARTOBJECT/ENDOBJECT records are used for Excel 9+ elements with child
 * records instead of the BEGIN/END records.
 * <p/>
 * Record Data
 * Offset		Field Name		Size	Contents
 * 4			rt				2		Record type; this matches the BIFF rt in the first two bytes of the record; =0854h
 * 6			grbitFrt		2		FRT flags; must be zero
 * 8			iObjectKind		2		Kinds of object, e.g., AI, CRT, SS, etc.
 * 10			iObjectContext	2		Where the object lives, object-specific
 * 12			iObjectInstance1 2		Which one from a collection, object-specific
 * 14			iObjectInstance2 2		Which one from a collection, object-specific.
 */
public class StartObject extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -8933906882043608941L;

	@Override
	public void init()
	{
		super.init();
	}    // iObjectKind, 18 must be Chart, or maybe Chart label ...

	private byte[] PROTOTYPE_BYTES = new byte[]{ 84, 8, 0, 0, 18, 0, 0, 0, 0, 0, 0, 0 };

	public static XLSRecord getPrototype()
	{
		StartObject sb = new StartObject();
		sb.setOpcode( STARTOBJECT );
		sb.setData( sb.PROTOTYPE_BYTES );
		sb.init();
		return sb;
	}

	private void updateRecord()
	{
	}

}
