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
 * <b>FRTWRAPPER: Chart Future Record Type Wrapper (851h)</b>
 * Introduced in Excel 9 (2000) this BIFF record is an FRT record for Charts.
 * This record is used to disguise a normal, non-FRT record as a FRT record.
 * This is necessary whenever a new Excel element must save a pre-Excel 9 record
 * as a child record. As an FRT record, Excel 97 will keep the record together with
 * its associated STARTOBJECT/ENDOBJECT when round-tripping FRTs. The size of this
 * record varies depending on the record that was wrapped.
 * <p/>
 * <p/>
 * Record Data
 * Offset		Field Name		Size		Contents
 * 4			rt				2			Record type; this matches the BIFF rt in the first two bytes of the record; =0851h
 * 6			grbitFrt		2			FRT flags; must be zero
 * 8			rt				2			RT of wrapped record
 * 10			cb				2			Size of wrapped RT's data in bytes
 * 12			rgb				var			RT's data
 */
public class FrtWrapper extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -4527467252753642328L;
	private int type;
	public static final int DEFAULTTEXT = 0, TEXTDISPWITHDATALABELS = 1, BEGIN = 2, POS = 3, AI = 4, OBJECTLINK = 5, END = 6, TEXTDISP = 7;

	@Override
	public void init()
	{
		super.init();
	}

	// 20080703 KSC: these bytes sequences haven't been entirelyy figured out but are necessary for Series Data Labels ..
	private byte[] P_0 = new byte[]{ 81, 8, 0, 0, 36, 16, 2, 0, 0, 0, 0, 0 };    // rt= 1024, DEFTEXT
	// rt= 1025, TEXTDISP
	private byte[] P_1 = new byte[]{
			81,
			8,
			0,
			0,
			37,
			16,
			32,
			0,
			2,
			2,
			1,
			0,
			0,
			0,
			0,
			0,
			-33,
			-1,
			-1,
			-1,
			-74,
			-1,
			-1,
			-1,
			0,
			0,
			0,
			0,
			0,
			0,
			0,
			0,
			-77,
			0,
			77,
			0,
			16,
			61,
			0,
			0
	};
	private byte[] P_2 = new byte[]{ 81, 8, 0, 0, 51, 16, 0, 0, 0, 0, 0, 0 };    //rt= 1033, BEGIN
	private byte[] P_3 = new byte[]{
			81, 8, 0, 0, 79, 16, 20, 0, 2, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0
	};    // rt= 104F, POS
	private byte[] P_4 = new byte[]{ 81, 8, 0, 0, 81, 16, 8, 0, 0, 1, 0, 0, 0, 0, 0, 0 };    // rt= 1051, AI
	private byte[] P_5 = new byte[]{ 81, 8, 0, 0, 39, 16, 6, 0, 4, 0, 0, 0, 0, 0 };    // rt= 1027, OBJECTLINK
	private byte[] P_6 = new byte[]{ 81, 8, 0, 0, 52, 16, 0, 0, 0, 0, 0, 0 };    // rt= 1034, END
	private byte[] P_7 = new byte[]{
			81,
			8,
			0,
			0,
			37,
			16,
			32,
			0,
			2,
			2,
			1,
			0,
			0,
			0,
			0,
			0,
			-33,
			-1,
			-1,
			-1,
			-74,
			-1,
			-1,
			-1,
			0,
			0,
			0,
			0,
			0,
			0,
			0,
			0,
			-79,
			0,
			77,
			0,
			16,
			61,
			0,
			0
	};

	/**
	 * create a new FrtWrapper record
	 * each record wraps around another record
	 *
	 * @param type type of other record
	 * @return
	 */
	public static XLSRecord getPrototype( int type )
	{
		FrtWrapper frt = new FrtWrapper();
		frt.type = type;                    // save wrapped type
		frt.setOpcode( FRTWRAPPER );
		byte[] b = null;
		switch( type )
		{
			case DEFAULTTEXT:
				b = frt.P_0;    // generate a default text record with 0=show labels
				break;
			case TEXTDISPWITHDATALABELS:    // generate a text display with showKey (I believe!)
				b = frt.P_1;
				break;
			case TEXTDISP:    // generate a text display with showKey (I believe!) and NO data labels
				b = frt.P_7;
				break;
			case BEGIN:
				b = frt.P_2;
				break;
			case POS:
				b = frt.P_3;    // generate a POS record with default values
				break;
			case AI:
				b = frt.P_4;    // generate an AI record for SERIES Values
				break;
			case OBJECTLINK:
				b = frt.P_5;        // generate record with TYPE_DATASERIES
				break;
			case END:
				b = frt.P_6;
				break;
		}
		frt.setData( b );
		frt.init();
		return frt;
	}

	private void updateRecord()
	{
	}

}
