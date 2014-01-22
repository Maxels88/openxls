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
import com.extentech.toolkit.Logger;

import java.io.UnsupportedEncodingException;

/**
 * TableStyles 0x88E
 * <p/>
 * The TableStyles record specifies the default table and PivotTabletable styles and
 * specifies the beginning of a collection of TableStyle records as defined by the Globals substream.
 * The collection of TableStyle records specifies user-defined table styles.
 * <p/>
 * frtHeader (12 bytes): An FrtHeader structure. The frtHeader.rt field MUST be 0x088E.
 * <p/>
 * cts (4 bytes): An unsigned integer that specifies the total number of table styles in this document. This is the sum of the standard built-in table styles and all of the custom table styles. This value MUST be greater than or equal to 144 (the number of built-in table styles).
 * <p/>
 * cchDefTableStyle (2 bytes): An unsigned integer that specifies the count of characters in the rgchDefTableStyle field. This value MUST be less than or equal to 255.
 * <p/>
 * cchDefPivotStyle (2 bytes): An unsigned integer that specifies the count of characters in the rgchDefPivotStyle field. This value MUST be less than or equal to 255.
 * <p/>
 * rgchDefTableStyle (variable): An array of Unicode characters whose length is specified by cchDefTableStyle that specifies the name of the default table style.
 * <p/>
 * rgchDefPivotStyle (variable): An array of Unicode characters whose length is specified by cchDefPivotStyle that specifies the name of the default PivotTable style.
 */
public class TableStyles extends XLSRecord implements XLSConstants
{
	/**
	 * serialVersionUID
	 */
	short cts, cchDefTableStyle, cchDefPivotStyle;
	String rgchDefTableStyle = null, rgchDefPivotStyle = null;
	private static final long serialVersionUID = 2639291289806138985L;

	@Override
	public void init()
	{
		super.init();
		// -114, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,  // old frtHeader
		//-112, 0, 0, 0, 	== 144
		// An unsigned integer that specifies the total number of table styles in this document. This is the sum of the standard built-in table styles and all of the custom table styles. This value MUST be greater than or equal to 144 (the number of built-in table styles).
		cts = (short) ByteTools.readInt( this.getByteAt( 12 ), this.getByteAt( 13 ), this.getByteAt( 14 ), this.getByteAt( 14 ) );
		cchDefTableStyle = ByteTools.readShort( this.getByteAt( 16 ), this.getByteAt( 17 ) );
		cchDefPivotStyle = ByteTools.readShort( this.getByteAt( 18 ), this.getByteAt( 19 ) );
		int pos = 20;
		if( cchDefTableStyle > 0 )
		{
			byte[] tmp = this.getBytesAt( pos, (cchDefTableStyle) * (2) );
			try
			{
				rgchDefTableStyle = new String( tmp, UNICODEENCODING );
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logInfo( "encoding Table Style name in TableStyles: " + e );
			}
			pos += cchDefTableStyle * (2);
		}
		if( cchDefPivotStyle > 0 )
		{
			byte[] tmp = this.getBytesAt( pos, (cchDefPivotStyle) * (2) );
			try
			{
				rgchDefPivotStyle = new String( tmp, UNICODEENCODING );
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logInfo( "encoding Pivot Style name in TableStyles: " + e );
			}
		}
	}

	public static XLSRecord getPrototype()
	{
		TableStyles tx = new TableStyles();
		tx.setOpcode( TABLESTYLES );
		tx.setData( new byte[]{
				-114,
				8,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,
				0,  /* required id */
				-112,
				0,
				0,
				0,	/* cts */
				17,
				0, 	/* cch */
				17,
				0, 	/* cch */
				84,
				0,
				97,
				0,
				98,
				0,
				108,
				0,
				101,
				0,
				83,
				0,
				116,
				0,
				121,
				0,
				108,
				0,
				101,
				0,
				77,
				0,
				101,
				0,
				100,
				0,
				105,
				0,
				117,
				0,
				109,
				0,
				57,
				0,
				80,
				0,
				105,
				0,
				118,
				0,
				111,
				0,
				116,
				0,
				83,
				0,
				116,
				0,
				121,
				0,
				108,
				0,
				101,
				0,
				76,
				0,
				105,
				0,
				103,
				0,
				104,
				0,
				116,
				0,
				49,
				0,
				54,
				0
		} );
		tx.init();
		return tx;
	}
}
