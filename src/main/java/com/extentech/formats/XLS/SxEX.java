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

import com.extentech.toolkit.Logger;

/**
 * SxEX  0xF1
 * <p/>
 * The SXEx record specifies additional properties of a PivotTable view and specifies the beginning of a collection of records as defined by the Worksheet substreamABNF.
 * The collection of records specifies selection and formatting properties for the PivotTable view.
 * <p/>
 * csxformat (2 bytes): An unsigned integer that specifies the number of SxFormat records that follow this record. MUST be less than or equal to 0xFFFF.
 * <p/>
 * cchErrorString (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the stError field.
 * If the value is 0xFFFF, then stError does not exist. MUST be 0xFFFF or MUST be greater than zero and less than or equal to 0x00FF.
 * <p/>
 * cchNullString (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the stDisplayNull field.
 * If the value is 0xFFFF, then stDisplayNull does not exist. MUST be 0xFFFF or MUST be greater than zero and less than or equal to 0x00FF.
 * <p/>
 * cchTag (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the stTag field.
 * If the value is 0xFFFF, then stTag does not exist. MUST be 0xFFFF or MUST be greater than zero and less than or equal to 0x00FF.
 * <p/>
 * csxselect (2 bytes): An unsigned integer that specifies the number of SxSelect records that follow this record. MUST be less than or equal to 0xFFFF.
 * <p/>
 * crwPage (2 bytes): A DRw structure that specifies the number of rows in the page area (see Location and Body) of the PivotTable view.
 * <p/>
 * ccolPage (2 bytes): A DCol structure that specifies the number of columns in the page area (see Location and Body) of the PivotTable view.
 * <p/>
 * A - fAcrossPageLay (1 bit): A bit that specifies how pivot fields are laid out in the page area (see Location and Body) when there are multiple pivot fields on the page axis. MUST be a value from the following table:
 * Value	    Meaning
 * 0x0		    Pivot fields are displayed in the page area from the top to the bottom first, as fields are added, before moving to another column.
 * 0x1		    Pivot fields are displayed in the page area from left to right first, as fields are added, before moving to another row.
 * <p/>
 * cWrapPage (8 bits): An unsigned integer that specifies the number of pivot fields in the page area (see Location and Body) to
 * display before moving to another row or column, as specified by fAcrossPageLay.
 * MUST be less than or equal to 0xFF. A value of 0 means that no wrap is allowed.
 * <p/>
 * B - unused (1 bit): Undefined and MUST be ignored.
 * <p/>
 * C - reserved1 (1 bit): MUST be zero and MUST be ignored.
 * <p/>
 * reserved2 (5 bits): MUST be zero and MUST be ignored.
 * <p/>
 * D - fEnableWizard (1 bit): A bit that specifies whether a wizard user interface is displayed to work with the PivotTable view.
 * <p/>
 * E - fEnableDrilldown (1 bit): A bit that specifies whether details can be shown for cells in the data area, as specified by PivotTable Layout.
 * <p/>
 * F - fEnableFieldDialog (1 bit): A bit that specifies whether a user interface for setting properties of a pivot field can be displayed.
 * <p/>
 * G - fPreserveFormatting (1 bit): A bit that specifies whether formatting is preserved when the PivotTable view is recalculated.
 * If the value is 1, csxformat MUST be 0 and there MUST be no SxFormat records following this record.
 * <p/>
 * H - fMergeLabels (1 bit): A bit that specifies whether empty cells adjacent to the cells displaying pivot item captions of
 * pivot fields on the row axis and column axis of the PivotTable view are merged into a single cell with center-aligned text.
 * <p/>
 * I - fDisplayErrorString (1 bit): A bit that specifies whether the PivotTable view displays the custom error string stError in cells that contain errors.
 * <p/>
 * J - fDisplayNullString (1 bit): A bit that specifies whether the PivotTable view displays the custom string stDisplayNull in cells that contain NULL values.
 * <p/>
 * K - fSubtotalHiddenPageItems (1 bit): A bit that specifies whether hidden pivot items, as specified by SXVI records with the fHidden field equal to 1, of a
 * pivot field on the page axis with the isxvi field of the corresponding SXPI_Item structure equal to 0x7FFD are filtered out when calculating the PivotTable view.
 * MUST be 0 for non-OLAPdata sources (1) if the PivotCache functionality level is 3.
 * <p/>
 * reserved3 (8 bits): MUST be zero and MUST be ignored.
 * <p/>
 * cchPageFieldStyle (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the stPageFieldStyle field.
 * If the value is 0xFFFF, then stPageFieldStyle does not exist.
 * MUST be 0xFFFF or MUST be greater than zero and less than or equal to 0x00FF.
 * <p/>
 * cchTableStyle (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the stTableStyle field.
 * If the value is 0xFFFF, then stTableStyle does not exist.
 * MUST be 0xFFFF or MUST be greater than zero and less than or equal to 0x00FF.
 * <p/>
 * cchVacateStyle (2 bytes): An unsigned integer that specifies the length, in characters, of the XLUnicodeStringNoCch structure in the stVacateStyle field.
 * If the value is 0xFFFF, then stVacateStyle does not exist.
 * MUST be 0xFFFF or MUST be greater than zero and less than or equal to 0x00FF.
 * <p/>
 * stError (variable): An XLUnicodeStringNoCch structure that specifies a custom string displayed in cells that contain
 * errors when the value of fDisplayErrorString is 1. The length is specified in cchErrorString.
 * This field is optional and MUST NOT exist if cchErrorString is 0xFFFF.
 * <p/>
 * stDisplayNull (variable): An XLUnicodeStringNoCch structure that specifies a custom string displayed in cells that contain
 * NULL values when fDisplayNullString is 1. The length is specified in cchNullString.
 * This field is optional and MUST NOT exist if cchNullString is 0xFFFF.
 * <p/>
 * stTag (variable): An XLUnicodeStringNoCch structure that specifies a custom string saved with the PivotTable view. The length is specified in cchTag.
 * This field is optional and MUST NOT exist if cchTag is 0xFFFF.
 * <p/>
 * stPageFieldStyle (variable): An XLUnicodeStringNoCch structure that specifies the style used in the page area (see Location and Body) of the PivotTable view.
 * The style is specified by the StyleExt record with its stName field equal to this field's value.
 * If cchPageFieldStyle is 0xFFFF or less than 1, no style is applied. The length is specified in cchPageFieldStyle.
 * This field is optional and MUST NOT exist if cchPageFieldStyle is 0xFFFF.
 * <p/>
 * stTableStyle (variable): An XLUnicodeStringNoCch structure that specifies the style used in the body of the PivotTable view.
 * The style is specified by the StyleExt record with its stName field equal to this field's value.
 * If cchTableStyle is 0xFFFF or less than 1, no style is applied. The length is specified in cchTableStyle.
 * This field is optional and MUST NOT exist if cchTableStyle is 0xFFFF.
 * <p/>
 * stVacateStyle (variable): An XLUnicodeStringNoCch structure that specifies the style applied to cells that
 * become empty when the PivotTable view is recalculated. The style is specified by the StyleExt record
 * with its stName field equal to this field's value. If cchVacateStyle is 0xFFFF or less than 1, no style is applied.
 * The length is specified in cchVacateStyle.
 * This field is optional and MUST NOT exist if cchVacateStyle is 0xFFFF.
 */
public class SxEX extends XLSRecord implements XLSConstants
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2639291289806138985L;

	@Override
	public void init()
	{
		super.init();
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "SXEX - " );
		}
	}

	private byte[] PROTOTYPE_BYTES = new byte[]{ // default configuration
	                                             0, 0, -1, -1, -1, -1, -1, -1, 0, 0, 0, 0, 0, 0, 0, 2, 79, 0, -1, -1, -1, -1, -1, -1
	};

	public static XLSRecord getPrototype()
	{
		SxEX sex = new SxEX();
		sex.setOpcode( SXEX );
		sex.setData( sex.PROTOTYPE_BYTES );
		sex.init();
		return sex;
	}
}
