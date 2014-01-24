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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * SxVDEx	0x100
 * <p/>
 * The SXVDEx record specifies extended pivot field properties.
 * <p/>
 * A - fShowAllItems (1 bit): A bit that specifies whether to show all pivot items for this pivot field,
 * including pivot items that do not currently exist in the source data. The value MUST be 0 for an OLAP PivotTable view.
 * MUST be a value from the following table:
 * Value		  	Meaning
 * 0x0			    Specifies that all pivot items are not displayed.
 * 0x1			    Specifies that all pivot items are displayed.
 * <p/>
 * B - fDragToRow (1 bit): A bit that specifies whether this pivot field can be placed on the row axis. This value MUST be ignored for an OLAP PivotTable view.
 * MUST be a value from the following table:
 * Value	    Meaning
 * 0x0		    Specifies that the user is prevented from placing this pivot field on the row axis.
 * 0x1		    Specifies that the user is not prevented from placing this pivot field on the row axis.
 * <p/>
 * C - fDragToColumn (1 bit): A bit that specifies whether this pivot field can be placed on the column axis. This value MUST be ignored for an OLAP PivotTable view. MUST be a value from the following table:
 * Value	    Meaning
 * 0x0		    Specifies that the user is prevented from placing this pivot field on the column axis.
 * 0x1		    Specifies that the user is not prevented from placing this pivot field on the column axis.
 * <p/>
 * D - fDragToPage (1 bit): A bit that specifies whether this pivot field can be placed on the page axis. This value MUST be ignored for an OLAP PivotTable view. MUST be a value from the following table:
 * Value	    Meaning
 * 0x0		    Specifies that the user is prevented from placing this pivot field on the page axis.
 * 0x1		    Specifies that the user is not prevented from placing this pivot field on the page axis.
 * <p/>
 * E - fDragToHide (1 bit): A bit that specifies whether this pivot field can be removed from the PivotTable view. This value MUST be ignored for an OLAP PivotTable view. MUST be a value from the following table:
 * Value	    Meaning
 * 0x0		    Specifies that the user is prevented from removing this pivot field from the PivotTable view.
 * 0x1		    Specifies that the user is not prevented from removing this pivot field from the PivotTable view.
 * <p/>
 * F - fNotDragToData (1 bit): A bit that specifies whether this pivot field can be placed on the data axis. This value MUST be ignored for an OLAP PivotTable view. MUST be a value from the following table:
 * Value	    Meaning
 * 0x0		    Specifies that the user is not prevented from placing this pivot field on the data axis.
 * 0x1		    Specifies that the user is prevented from placing this pivot field on the data axis.
 * <p/>
 * G - reserved1 (1 bit): MUST be zero, and MUST be ignored.
 * <p/>
 * H - fServerBased (1 bit): A bit that specifies whether this pivot field is server-based when on the page axis. For more information, see Source Data.
 * A value of 1 specifies that this pivot field is a server-based pivot field.
 * <p/>
 * MUST be 1 if and only if the value of the fServerBased field of the SXFDB record of the associated cache field of this pivot field is 1.
 * <p/>
 * I - reserved2 (1 bit): MUST be zero, and MUST be ignored.
 * <p/>
 * J - fAutoSort (1 bit): A bit that specifies whether AutoSort will be applied to this pivot field. For more information, see Pivot Field Sorting.
 * <p/>
 * K - fAscendSort (1 bit): A bit that specifies whether any AutoSort applied to this pivot field will sort in ascending order. MUST be a value from the following table:
 * Value		    Meaning
 * 0x0			    Sort in descending order.
 * 0x1			    Sort in ascending order.
 * <p/>
 * L - fAutoShow (1 bit): A bit that specifies whether an AutoShowfilter is applied to this pivot field. For more information, see Simple Filters.
 * <p/>
 * M - fTopAutoShow (1 bit): A bit that specifies whether any AutoShow filter applied to this pivot field shows the top-ranked or bottom-ranked values. For more information, see Simple Filters. MUST be a value from the following table:
 * Value		    Meaning
 * 0x0			    Any AutoShow filter applied to this pivot field shows the bottom-ranked values.
 * 0x1			    Any AutoShow filter applied to this pivot field shows the top-ranked values.
 * <p/>
 * N - fCalculatedField (1 bit): A bit that specifies whether this pivot field is a calculated field. A value of 1 specifies that this pivot field is a calculated field.
 * <p/>
 * MUST be 1 if and only if the value of the fCalculatedField field of the SXFDB record of the cache field associated with this pivot field is 1.
 * <p/>
 * O - fPageBreaksBetweenItems (1 bit): A bit that specifies whether a page break (2) is inserted after each pivot item when the PivotTable is printed.
 * <p/>
 * P - fHideNewItems (1 bit): A bit that specifies whether new pivot items that appear after a refresh are hidden by default. This value MUST be equal to 0 for a non-OLAP PivotTable view.
 * Value	    Meaning
 * 0x0		    New pivot items are shown by default.
 * 0x1		    New pivot items are hidden by default.
 * <p/>
 * reserved3 (5 bits): MUST be zero, and MUST be ignored.
 * <p/>
 * Q - fOutline (1 bit): A bit that specifies whether this pivot field is in outline form. For more information, see PivotTable layout.
 * <p/>
 * R - fInsertBlankRow (1 bit): A bit that specifies whether to insert a blank row after each pivot item.
 * <p/>
 * S - fSubtotalAtTop (1 bit): A bit that specifies whether subtotals are displayed at the top of the group when the fOutline field is equal to 1. For more information, see PivotTable layout.
 * <p/>
 * citmAutoShow (8 bits): An unsigned integer that specifies the number of pivot items to show when the fAutoShow field is equal to 1.
 * The value MUST be greater than or equal to 1 and less than or equal to 255.
 * <p/>
 * isxdiAutoSort (2 bytes): A signed integer that specifies the data item that AutoSort uses when the fAutoSort field is equal to 1. If the value of the fAutoSort field is one,
 * the value MUST be greater than or equal to zero and less than the count of SXDI records. MUST be a value from the following table:
 * Value    	    Meaning
 * -1			    Specifies that the values of the pivot items themselves are used.
 * Greater than or equal to zero	Specifies a data item index, as specified in Data Items, of the data item that is used.
 * <p/>
 * isxdiAutoShow (2 bytes): A signed integer that specifies the data item that AutoShow ranks by when the fAutoShow field is equal to 1.
 * For more information, see Simple Filters. If the value of the fAutoShow field is 1, this value MUST be greater than or equal to zero and less than the count of SXDI records.
 * MUST be a value from the following table:
 * Value		    Meaning
 * -1			    AutoShow is not enabled for this pivot field.
 * Greater than or equal to zero	    Specifies a data item index, as specified in Data Items, of the data item that is used.
 * <p/>
 * ifmt (2 bytes): An IFmt structure that specifies the number format of this pivot field.
 * <p/>
 * subName (variable): An optional SXVDEx_Opt structure that specifies the name of the aggregate function used to calculate this pivot field's subtotals. SHOULD<124> be present.
 */
public class SxVdEX extends XLSRecord implements XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( SxVdEX.class );
	private static final long serialVersionUID = 2639291289806138985L;
	private short citmAutoShow;
	private short isxdiAutoSort;
	private short isxdiAutoShow;
	private short ifmt;

	@Override
	public void init()
	{
		super.init();
		// TODO: flags
		citmAutoShow = getByteAt( 4 );
		isxdiAutoSort = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );
		isxdiAutoShow = ByteTools.readShort( getByteAt( 6 ), getByteAt( 7 ) );
		ifmt = ByteTools.readShort( getByteAt( 8 ), getByteAt( 9 ) );
		// TODO: subName (variable): An optional SXVDEx_Opt structure that specifies the name of the aggregate function used to calculate this pivot field's subtotals. SHOULD<124> be present.

			log.debug( "SXVDEX - citmAutoShow:" + citmAutoShow + " isxdiAutoSort:" + isxdiAutoSort + " isxdoAutoShow:" + isxdiAutoShow + " ifmt:" + ifmt );
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ // default configuration
	                                             30, 20, 0, 10, -1, -1, -1, -1, 0, 0, -1, -1, 0, 0, 0, 0, 0, 0, 0, 0
	};

	public static XLSRecord getPrototype()
	{
		SxVdEX sv = new SxVdEX();
		sv.setOpcode( SXVDEX );
		sv.setData( sv.PROTOTYPE_BYTES );
		sv.init();
		return sv;
	}
}
