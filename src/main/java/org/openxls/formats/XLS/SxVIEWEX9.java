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

/**
 * SXViewEx9  0x810
 * <p/>
 * The SXViewEx9 record specifies extensions to the PivotTable view.
 * <p/>
 * rt (2 bytes): An unsigned integer that specifies the record type identifier. The value MUST be 0x0810.
 * <p/>
 * A - reserved1 (1 bit): MUST be zero, and MUST be ignored.
 * <p/>
 * B - fFrtAlert (1 bit): A bit that specifies whether features of this PivotTable are not supported in earlier versions of the BIFF.
 * An application can alert the user of possible problems when saving as an earlier version of the BIFF.
 * <p/>
 * reserved2 (14 bits): MUST be zero, and MUST be ignored.
 * <p/>
 * reserved3 (4 bytes): MUST be zero, and MUST be ignored.
 * <p/>
 * C - reserved4 (1 bit): MUST be zero, and MUST be ignored.
 * <p/>
 * D - fPrintTitles (1 bit): A bit that specifies whether the print titles for the worksheet are set based on the PivotTable report.
 * The row print titles are set to the pivot item captions on the column axis and the column print titles are set to the pivot item captions on the row axis.
 * <p/>
 * E - fLineMode (1 bit): A bit that specifies whether any pivot field is in outline mode. See Subtotalling for more information.
 * <p/>
 * F - reserved5 (2 bits): MUST be zero, and MUST be ignored.
 * <p/>
 * G - fRepeatItemsOnEachPrintedPage (1 bit): A bit that specifies whether pivot item captions on the row axis are repeated on each printed page for pivot fields in tabular form.
 * <p/>
 * reserved6 (26 bits): MUST be zero, and MUST be ignored.
 * <p/>
 * itblAutoFmt (2 bytes): An AutoFmt8 structure that specifies the PivotTable AutoFormat. If the value of this field is not 1,
 * this field overrides the itblAutoFmt field in the previous SxView record.
 * <p/>
 * chGrand (variable): An XLUnicodeString structure that specifies a user-entered caption to display for grand totals when the PivotTable is recalculated.
 * The length MUST be less than or equal to 255 characters.
 */
public class SxVIEWEX9 extends XLSRecord implements XLSConstants
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2639291289806138985L;

	@Override
	public void init()
	{
		super.init();
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0x10, 0x8, /* rt= 0x810 */
			00, 00,		/* flags + reserved */
			00, 00, 00, 00, 0x20, 00, 00, 00, 01, 00,	 	/* itblAutoFmt*/
			00, 00, 00  	/* chGrand-- cch= 0, encoding= 0 */
	};

	public static XLSRecord getPrototype()
	{
		SxVIEWEX9 sxv = new SxVIEWEX9();
		sxv.setOpcode( SXVIEWEX9 );
		sxv.setData( sxv.PROTOTYPE_BYTES );
		sxv.init();
		return sxv;
	}
}
