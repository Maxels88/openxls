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
 * <b>ENDOBJECT: Chart Future Record Type End Object (855h)</b>
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record
 * for Charts that indicates the end of an object's scope for Excel 9 and later objects.
 * <p/>
 * Record Data
 * Offset		Field Name	Size	Contents
 * 4			rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =0855h
 * 6			grbitFrt	2		FRT flags; must be zero
 * 8			iObjectKind	2		Sanity check for object scope being ended
 * 10			(unused)	6		Reserved; must be zero
 */
public class EndObject extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 8367476159843855287L;

	public void init()
	{
		super.init();
	}        // iObjectKind of 18 appears to be Chart object

	private byte[] PROTOTYPE_BYTES = new byte[]{ 85, 8, 0, 0, 18, 0, 0, 0, 0, 0, 0, 0 };

	public static XLSRecord getPrototype()
	{
		EndObject eo = new EndObject();
		eo.setOpcode( ENDOBJECT );
		eo.setData( eo.PROTOTYPE_BYTES );
		eo.init();
		return eo;
	}

	private void updateRecord()
	{
	}

}
