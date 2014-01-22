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
 * <b> Dat:  Data Table Options  (0x1063)</b>
 * <p/>
 * Offset	Name	Size	Contents
 * 4		grbit	2		Option flags (see following table)
 * <p/>
 * <p/>
 * The grbit field contains the following flags.
 * <p/>
 * Offset	Bits	Mask	Name			Contents
 * 0		0		01h		fHasBordHorz	1 = data table has horizontal borders
 * 1		02h		fHasBordVert	1 = data table has vertical borders
 * 2		04h		fhasBordOutline	1 = data table has a border
 * 3		08h		fShowSeriesKey	1 = data table shows series keys
 * 7-4		F0h		reserved		Reserved; must be zero
 * 1		7-0		FFh		reserved		Reserved; must be zero
 */
public class Dat extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 1138056714558134785L;
	private short grbit;
	boolean fHasBordHorz, fHasBordVert, fHasBordOutline, fShowSeriesKey;

	@Override
	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		fHasBordHorz = ((grbit & 0x1) == 0x1);
		fHasBordVert = ((grbit & 0x2) == 0x2);
		fHasBordOutline = ((grbit & 0x4) == 0x4);
		fShowSeriesKey = ((grbit & 0x8) == 0x8);

	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ 15, 0 };

	/**
	 * creates a new Dat record; if bCreateDataTable is true,
	 * will also add the associated records required to create
	 * a Data Table
	 *
	 * @param bCreateDataTable
	 * @return
	 */
	public static XLSRecord getPrototype( boolean bCreateDataTable )
	{
		Dat d = new Dat();
		d.setOpcode( DAT );
		d.setData( d.PROTOTYPE_BYTES );
		d.init();
		if( bCreateDataTable )
		{
			Legend l = (Legend) Legend.getPrototype();
			l.setIsDataTable( true );
			//l.setwType(Legend.NOT_DOCKED);
			d.chartArr.add( l );
			// add pos record
			Pos p = (Pos) Pos.getPrototype( Pos.TYPE_DATATABLE );
//            p.setWorkBook(book);
			l.addChartRecord( p );
			// TextDisp
			TextDisp td = (TextDisp) TextDisp.getPrototype();
			l.addChartRecord( td );
			// TextDisp sub-recs
			p = (Pos) Pos.getPrototype( Pos.TYPE_TEXTDISP );
			td.addChartRecord( p );
			Fontx f = (Fontx) Fontx.getPrototype();
//	        f.setWorkBook(book);
// EVENTUALLY!	        
			f.setIfnt( 6 );
			td.addChartRecord( f );
			Ai ai = (Ai) Ai.getPrototype( Ai.AI_TYPE_NULL_LEGEND );
//            ai.setWorkBook(book);
			td.addChartRecord( ai );
		}
		return d;
	}

	private void updateRecord()
	{
		byte[] b = ByteTools.shortToLEBytes( grbit );
		this.setData( b );
	}
}
