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
 * <b>STARTBLOCK: Chart Future Record Type Start Block (852h)</b>
 * Introduced in Excel 9 (2000) this BIFF record is an FRT record for Charts.
 * Indicates the start of an object's scope for Pre-Excel 9 objects. These
 * records are used to push a chart element scope onto the parent element stack.
 * This stack is used to determine the containing element for records that are
 * used by more than one type of element. The FRAME record, for instance, is used
 * by at least four different elements.
 * <p/>
 * The STARTBLOCK/ENDBLOCK records are used for pre-Excel 9 elements with child records
 * (i.e., a record for the element followed by a BEGIN/END block for the child records.)
 * STARTBLOCK/ENDBLOCK are only written to enclose one or more child CFRT records and can
 * be placed outside the original BEGIN/END block. They may be omitted otherwise.
 * These records allow Excel 9 or later to determine the proper parent element even after
 * Excel 97 moves CFRTs to the end of the stream. Since these records are CFRTs,
 * they will stay with and keep contained any child CFRTs.
 * <p/>
 * Record Data
 * Offset		Field Name		Size	Contents
 * 4			rt				2		Record type; this matches the BIFF rt in the first two bytes of the record; =0852h
 * 6			grbitFrt		2		FRT flags; must be zero
 * 8			iObjectKind		2		See table below
 * 10			iObjectContext	2		See table below
 * 12			iObjectInstance1 2		See table below
 * 14			iObjectInstance2 2		See table below
 * <p/>
 * The following table describes the meaning of each set of possible values for iObjectKind,
 * iObjectContext, iObjectInstance1, iObjectInstance2. In some cases, these fields are indexed,
 * and the indexes are described in the documentation for the parent rt. The table also lists
 * whether the STARTBLOCK/ENDBLOCK or STARTOBJECT/ENDOBJECT rts are used, and the parent rt.
 * <p/>
 * iObjectKind	iObjectContext	iObjectInstance1	iObjectInstance2	Class	rt			Description
 * 0			0				0					0					BLOCK	AXIS PARENT	Primary axis group
 * 0			0				1					0					BLOCK	AXIS PARENT	Secondary axis group
 * 2			0				0					0					BLOCK	TEXT		Chart title
 * 2			1				xi					yi					BLOCK	TEXT		Data label for point in hidden series
 * 2			2				0					0					BLOCK	TEXT		Default data label for other cases
 * 2			2				1					0					BLOCK	TEXT		Default data label for showing values only
 * 2			4				0					0					BLOCK	TEXT		Category axis title
 * 2			4				1					0					BLOCK	TEXT		Value axis title
 * 2			4				2					0					BLOCK	TEXT		Series axis title
 * 2			5				xi					yi					BLOCK	TEXT		Data label for point in visible series, iobjInstance1 and iobjInstance2 is the DATAFORMAT xi and yi
 * 2			6				0					0					OBJECT	TEXT		Display unit label
 * 4			0				0					0					BLOCK	AXIS		Category axis
 * 4			0				1					0					BLOCK	AXIS		Value axis
 * 4			0				2					0					BLOCK	AXIS		Series axis
 * 4			0				3					0					BLOCK	AXIS		X-axis on scatter chart
 * 5			0				index				0					BLOCK	CHART FORMAT Chart group, iobjInstance1 is the index in the file
 * 6			0				0					0					BLOCK	DAT			Data table
 * 7			0				0					0					BLOCK	FRAME		Frame
 * 7			1				0					0					BLOCK	FRAME		Frame for an axis
 * 7			2				0					0					BLOCK	FRAME		Chart area frame
 * 8			0				0					0					BLOCK	GELFRAME	Frame fill
 * 8			1				0					0					BLOCK	GELFRAME	Series fill
 * 8			2				0					0					BLOCK	GELFRAME	Up/down bars fill
 * 8			3				0					0					BLOCK	GELFRAME	Floor fill
 * 8			3				1					0					BLOCK	GELFRAME	Walls fill
 * 9			0				0					0					BLOCK	LEGEND		Data table
 * 9			1				0					0					BLOCK	LEGEND		Legend
 * 10			0				iss					0					BLOCK	LEGENDXN	Legend entry
 * 11			0				0					0					BLOCK	PICF		Picture fill
 * 11			1				0					0					BLOCK	PICF		Data point picture fill
 * 12			0				index				0					BLOCK	SERIES		Series, iobjInstance1 is the index in the file
 * 13			0				0					0					BLOCK	CHART		Chart
 * 14			-1				0					0					BLOCK	DATA FORMAT Series formatting
 * 14			yi				xi					0					BLOCK	DATA FORMAT	Data point formatting
 * 15			0				0					0					BLOCK	DROPBAR		Up bars
 * 15			0				1					0					BLOCK	DROPBAR		Down bars
 * 15			0				2					0					BLOCK	AXISLINE FORMAT Floor
 * 15			0				3					0					BLOCK	AXISLINE FORMAT Walls
 * 16			0				0					0					OBJECT	YMULT		Axis multiplier
 * 17			0				verChart			0					OBJECT	FRTFONT LIST Fonts
 */
public class StartBlock extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 1895593244077899106L;
	short iObjectKind = 0;
	public static final int CHART = 13;
	public static final int AXIS = 0;
	public static final int CHARTFORMAT = 5;

	@Override
	public void init()
	{
		super.init();
		iObjectKind = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
	} // iObjectKind= 13 or 0 or 5

	private byte[] PROTOTYPE_BYTES = new byte[]{ 82, 8, 0, 0, 13, 0, 0, 0, 0, 0, 0, 0 };

	public static XLSRecord getPrototype()
	{
		StartBlock sb = new StartBlock();
		sb.setOpcode( STARTBLOCK );
		sb.setData( sb.PROTOTYPE_BYTES );
		sb.init();
		return sb;
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
