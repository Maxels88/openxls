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
 * <b>FRTFONTLIST: Chart Font List (85Ah)</b>
 * Introduced in Excel 9 (2000), this BIFF record is an FRT record for
 * Charts. This record stores font information for Excel 9 or later chart
 * objects. On round-tripping through an earlier version of Excel, the fonts
 * for new chart objects are lost from the font table because earlier version
 * of Excel do not load the newer objects and thus dont preserve the new objects
 * fonts. This record contains a list of the font indices used by Excel 9 or later
 * objects and whether the font is auto-scaled. The fonts themselves are stored
 * information in a STARTOBJECT/ENDOBJECT block that immediately follows. The block
 * has objectKind = 17, objectContext = 0, objectInstance1 = 0, objectInstance2 = 0.
 * The block has cfont FONT records and FBI records (for those with fScaled =1 only).
 * <p/>
 * Record Data
 * Offset	Field Name	Size	Contents
 * 4		rt			2		Record type; this matches the BIFF rt in the first two bytes of the record; =085Ah
 * 6		grbitFrt	2		FRT flags; must be zero
 * 8		verChart	1		Version of Charting this list applies to
 * 9		cfont		2		Number of fonts in list
 * 11		rgFontInfo	var		Array of font IDs
 * <p/>
 * FontInfo Structure
 * Offset	Field Name	Size	Contents
 * 0		grbit		2		Option flags for chart fonts (see description below)
 * 2		ifnt		2		Font ID of this font entry
 * <p/>
 * The grbit field contains the following chart font flags:
 * Bits	Mask	Flag Name	Contents
 * 0		0001h	fScaled		=1 if the font is scaled =0 otherwise
 * 15-1	FFFEh	(unused)	Reserved; must be zero
 */
public class FrtFontList extends GenericChartObject implements ChartObject
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 5374035624852924277L;

	@Override
	public void init()
	{
		super.init();
	}

	// TODO: Prototype Bytes
	private byte[] PROTOTYPE_BYTES = new byte[]{ };

	public static XLSRecord getPrototype()
	{
		FrtFontList ffl = new FrtFontList();
		ffl.setOpcode( FRTFONTLIST );
		ffl.setData( ffl.PROTOTYPE_BYTES );
		ffl.init();
		return ffl;
	}

	private void updateRecord()
	{
	}

}
