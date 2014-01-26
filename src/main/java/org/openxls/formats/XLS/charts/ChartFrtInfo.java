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
package org.openxls.formats.XLS.charts;

import org.openxls.formats.XLS.XLSRecord;

/**
 * <b>CHARTFRTINFO: Chart Future Record Type Info (850h)</b>
 * <p/>
 * Introduced in Excel 9 (2000) this BIFF record is an FRT record for Charts.
 * This record contains information describing the versions of Excel that originally
 * created and last saved the file, and the FRT IDs that are used in the file.
 * <br>In a file written by Excel 2000 or later, this record appears before the end of CHART
 * record block and before any other FRT in the Chart record stream.
 * <br>In a file written by Excel 97, this record may be missing or will appear after
 * the CHART record block. If this record appears after the END record of CHART record block,
 * the verWriter field is assumed to be 8 for Excel 97 regardless of the actual value
 * in the record.
 * <p/>
 * Record Data
 * Offset		Field Name		Size		Contents
 * 4			rt				2			Record type; this matches the BIFF rt in the first two bytes of the record; =0850h
 * 6			grbitFrt		2			FRT flags; must be zero
 * 8			verOriginator	1			Excel version that originally created the file
 * 9			verWriter		1			Excel version that last saved the file
 * 10			cCFRTID			2			Count of FRT ID value ranges in list
 * 12			rgCFRTID		var			List of FRT ID values used for charts
 * <p/>
 * CFRTID Structure
 * Offset		Field Name		Size		Contents
 * 0			rtFirst			2			First FRT in range
 * 2			rtLast			2			Last FRT in range
 */
public class ChartFrtInfo extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 7722730813154117198L;

	@Override
	public void init()
	{
		super.init();
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 80, 8, 0, 0, 10, 10, 3, 0, 80, 8, 90, 8, 97, 8, 97, 8, 106, 8, 107, 8 };

	public static XLSRecord getPrototype()
	{
		ChartFrtInfo cri = new ChartFrtInfo();
		cri.setOpcode( CHARTFRTINFO );
		cri.setData( cri.PROTOTYPE_BYTES );
		cri.init();
		return cri;
	}

	private void updateRecord()
	{
	}

}
