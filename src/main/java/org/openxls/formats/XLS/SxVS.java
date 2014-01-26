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
package org.openxls.formats.XLS;

import org.openxls.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * SXVS  0xE3
 * The SXVS record specifies the type of source data used for a PivotCache.
 * This record is followed by a sequence of records that specify additional information about the source data.
 * <p/>
 * sxvs (2 bytes): An unsigned integer that specifies the type of source data used for the PivotCache. The types of records that follow this record are dictated by the value of this field. MUST be a value from the following table:
 * <p/>
 * Name		Value		Meaning
 * <p/>
 * SHEET		0x0001		Specifies that the source data is a range. This record MUST be followed by a DConRef record that specifies a simple range, or a DConName record that specifies a named range, or a DConBin record that specifies a built-in named range.
 * EXTERNAL	0x0002		Specifies that external source data is used. This record MUST be followed by a sequence of records beginning with a DbQuery record that specifies connection and query information that is used to retrieve external data.
 * CONSOLIDATION	0x0004	Specifies that multiple consolidation ranges are used as the source data. This record MUST be followed by a sequence of records beginning with an SXTbl record that specifies information about the multiple consolidation ranges.
 * SCENARIO	0x0010		The source data is populated from a temporary internal structure. In this case there is no additional source data information because the raw data does not exist as a permanent structure and the logic to produce it is application-dependent.
 */
public class SxVS extends XLSRecord implements XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( SxVS.class );
	private short sourceType = -1;
	public static final short TYPE_SHEET = 0x1;
	public static final short TYPE_EXTERNAL = 0x0002;
	public static final short TYPE_CONSOLIDATION = 0x0004;
	public static final short TYPE_SCENARIO = 0x0010;

	private static final long serialVersionUID = 2639291289806138985L;

	@Override
	public void init()
	{
		super.init();
		sourceType = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
			log.debug( "SXVS - sourceType:" + sourceType );
	}

	public short getSourceType()
	{
		return sourceType;
	}

	public void setSourceType( int st )
	{
		sourceType = (short) st;
		byte[] b = ByteTools.shortToLEBytes( sourceType );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	public static XLSRecord getPrototype()
	{
		SxVS sv = new SxVS();
		sv.setOpcode( SXVS );
		sv.setData( new byte[]{ 1, 0 } );
		sv.init();
		return sv;
	}
}
