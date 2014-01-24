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
import com.extentech.toolkit.ByteTools;

/**
 * <b>ENDBLOCK: Chart Future Record Type End Block (853h)</b>
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for
 * Charts that indicates end of an object's scope for Pre-Excel 9 objects.
 * Paired with STARTBLOCK.
 * <p/>
 * Record Data
 * Offset		Field Name		Size		Contents
 * 4			rt				2			Record type; this matches the BIFF rt in the first two bytes of the record; =0853h
 * 6			grbitFrt		2			FRT flags; must be zero
 * 8			iObjectKind		2			Sanity check for object scope being ended
 * 10			(unused)		6			Reserved; must be zero
 */
public class EndBlock extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 1132544432743236942L;
	short iObjectKind = 0;

	@Override
	public void init()
	{
		super.init();
		iObjectKind = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
	}    // iObjectKind= 5 or 0 or 13

	private byte[] PROTOTYPE_BYTES = new byte[]{ 83, 8, 0, 0, 5, 0, 0, 0, 0, 0, 0, 0 };

	public static XLSRecord getPrototype()
	{
		EndBlock eb = new EndBlock();
		eb.setOpcode( ENDBLOCK );
		eb.setData( eb.PROTOTYPE_BYTES );
		eb.init();
		return eb;
	}

	public void setObjectKind( int i )
	{
		iObjectKind = (short) i;
		updateRecord();
	}

	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( iObjectKind );
		getData()[4] = b[0];
		getData()[5] = b[1];
	}

}
