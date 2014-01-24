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

/**
 * FEATHEADR
 * Shared Feature Header (867h)
 * Introduced in Excel 10 (2002) the FEATHEADR record describes the common information (header) for shared features such as Protection and SmartTag.
 * For example, if you have a worksheet that contains Protection, a Shared Feature Header (FEATHEADER) record is created for all Protections which may include
 * Sheet Protection or Book Protection. Though Sheet Protection and Book Protection may have specific data that are different and are saved in the Feature Data
 * (see FEAT record for detail) portion, their common settings are stored in this header block record.
 * A worksheet may contain one or more different types of Shared Feature and each type of Shared Feature has its own Shared Feature Header
 * (FEATHEADER) record to store common information across all Shared Feature of the same type.
 * This FEATHEADER record will have different data structure layout according to the Shared Feature type (the isf field flags differentiate Shared Feature types).
 * For example, if a Workbook has both Protection and SmartTag, there is one Shared Feature Header (FEATHEADER) created for Protection, and another Shared Feature Header
 * (FEATHEADER) created for SmartTag. Therefore, the data block of the Shared Feature Header (FEATHEADER) may have a different data structure depending on which
 * Shared Feature Type the record is for.
 * <p/>
 * Though Excel currently has many different Shared Features such as Formula Error Checking, Protection, SmartTag etc.,
 * only 2 types of Shared Feature are persisted in Excel 2002: Protection and SmartTag.
 * <p/>
 * Special note: In Excel 2003, this FEATHEADER is not used for new shared features, such as Tables, since a new record of FEATFEADER11 was introduced for better feature
 * data round-tripping stories through earlier versions of Excel.
 * <p/>
 * <p/>
 * Contents	size
 * 4		rt			2	Record type; this matches the BIFF rt in the first two bytes of the record; =0867h
 * 6		grbitFrt	2	FRT cell reference flag =0 currently
 * 8		(Reserved)	8	Currently not used and set to 0.
 * 16		isf			2	Shared feature type index =2 for Enhanced Protection =4 for SmatTag
 * 18		fHdr		1	=1 since this is a feat header
 * 19		cbHdrData	4	Size of rgbHdrSData =4 for simple feature headers =0 there is no rgbHdrData =-1 for complex feature headers, the size of rgbHdrData depends on the isf type. (prior to Excel 2003, all features saved using FEATHEAER use complex features.)
 * 23		rgbHdrData	var	Byte array of extra info, including from future versions of Excel
 * <p/>
 * The rgbHdrData block for Enhanced Protection
 * Contents	size
 * 0		grbit		4	Bit flag for protection rules setting (see table below for detail bit settings)
 * <p/>
 * The bit settings for protection rules settings in the grbit:
 * Bits	Mask		Bit Name				Description
 * 0		00000001h	iprotObject				Edit object
 * 1		00000002h	iprotScenario			Edit scenario
 * 2		00000004h	iprotFormatCells		Format cells
 * 3		00000008h	iprotFormatColumns		Format columns
 * 4		00000010h	iprotFormatrows			Format rows
 * 5		00000020h	iprotInsertColumns		Insert columns
 * 6		00000040h	iprotInsertRows			Insert rows
 * 7		00000080h	iprotInsertHyperlinks	Insert hyperlinks
 * 8		00000100h	iprotDeleteColumns		Delete columns
 * 9		00000200h	iprotDeleteRows			Delete rows
 * 10		00000400h	iprotSelLockedCells		Select locked cells
 * 11		00000800h	iprotSort				Sort
 * 12		00001000h	iprotAutoFilter			Use Autofilter
 * 13		00002000h	iprotPivotTables		Use PivotTable reports
 * 14		00004000h	iprotSelUnlockedCells	Select unlocked cells
 * <p/>
 * The rgbHdrData block for SmartTag
 * Offset	Field Name	Size	Contents
 * 0		cSmartTag	4		Count of SmartTags
 * 4		rgbSmartTag	var		Array of SmartTag header data (see table below).
 * var		cbPropBagData	2	Count of bytes in Property Bag store plus unknown data, i.e., rest of the data size including this count bit
 * var		sVersion	2		Version number
 * var		(Reserved)	4		Currently not used.
 * var		Cste		4		String table entry
 * var		cbUnknown	2		Count of bytes of unknown data
 * var		pvUnknoan	var		The Unknown data
 * <p/>
 * Where the rgbSmartTag, the array of SmartTag header data block, has the fields:
 * Offset	Field Name	Size	Contents
 * 0		cbSmartTag	4		Count of Byte of the SmartTag data
 * 4		id			4		Id of this SmartTag
 * 8		cbUri		2		Count of bytes of rgbUri
 * 10		rgbUri		var		Character string of URI (Universal Resource Identifier) which is the name space portion of the SmartTag, of length cbUri e.g., urn:schemas-microsoft-com:office:smarttags
 * var		cbTag		2		Count of bytes of the SmartTag tag name rgbTag
 * var		rgbTag		var		Character string of the SmartTag tag name, of length cbTag e.g., Stockticker
 * var		cbDownLoadURL	2	Count of bytes of Download URL address string: rgbDownLoadURL
 * var		rgbDownLoadURL	var	Character string of downloading URL
 * Var		pvUnknown	var		Additional data in pvUnknown, of length cbSmartTag – cbUri – cbTag – cbDownLoadURL – 10
 * In Excel 2003, only book level SmartTags‘ headers are saved. Sheet level SmatTag headers are not persisted.
 */
public class FeatHeadr extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -5775187385375827918L;
	short isf = 0;
	int grbit = 0;
	boolean iprotObject;
	boolean iprotScenario;
	boolean iprotFormatCells;
	boolean iprotFormatColumns;
	boolean iprotFormatRows;
	boolean iprotInsertColumns;
	boolean iprotInsertRows;
	boolean iprotInsertHyperlinks;
	boolean iprotDeleteColumns;
	boolean iprotDeleteRows;
	boolean iprotSelLockedCells;
	boolean iprotSort;
	boolean iprotAutoFilter;
	boolean iprotPivotTable;
	boolean iprotSelUnlockedCells;
	public final static short ALLOWOBJECTS = 0x1;
	public final static short ALLOWSCENARIOS = 0x2;
	public final static short ALLOWFORMATCELLS = 0x4;
	public final static short ALLOWFORMATCOLUMNS = 0x8;
	public final static short ALLOWFORMATROWS = 0x10;
	public final static short ALLOWINSERTCOLUMNS = 0x20;
	public final static short ALLOWINSERTROWS = 0x40;
	public final static short ALLOWINSERTHYPERLINKS = 0x80;
	public final static short ALLOWDELETECOLUMNS = 0x100;
	public final static short ALLOWDELETEROWS = 0x200;
	public final static short ALLOWSELLOCKEDCELLS = 0x400;
	public final static short ALLOWSORT = 0x800;
	public final static short ALLOWAUTOFILTER = 0x1000;
	public final static short ALLOWPIVOTTABLES = 0x2000;
	public final static short ALLOWSELUNLOCKEDCELLS = 0x4000;

	// TODO: handle SmartTag options

	/**
	 * default constructor
	 */
	private byte[] PROTOTYPE_BYTES = {
			0x67, 0x08, // rt or record type
			0, 0,        // frt cell ref flag
			0, 0, 0, 0, 0, 0, 0, 0,    // reserved 0
			2, 0,        // isf: shared feature flag 2=enhanced protection, 4=smart tag
			1,            // always 1
			-1, -1, -1, -1,    // size of rgbhdrSData, -1=complex size depends upon isf:
			0, 0, 0, 0
	};        // for enhanced protection, 4 bytes = protection rules

	/**
	 * Create a FeatHeadr record, default defines enhanced protection features
	 *
	 * @return
	 */
	protected static XLSRecord getPrototype()
	{
		FeatHeadr f = new FeatHeadr();
		f.setOpcode( FEATHEADR );
		f.setData( f.PROTOTYPE_BYTES );
		f.init();
		return f;
	}

	/**
	 * init the record - as of now, only enhanced protection is supported
	 */
	@Override
	public void init()
	{
		super.init();
		isf = ByteTools.readShort( getData()[12], getData()[13] );
		if( isf == 2 )
		{// Enhanced Protection - only option supported for now - read last 4 bytes for enhanced protection settings
			grbit = ByteTools.readInt( getData()[19], getData()[20], getData()[21], getData()[22] );
			parseProtectionGrbit();
		}
	}

	/**
	 * parse grbit, get boolean protection values
	 */
	private void parseProtectionGrbit()
	{
		// parse grbit
		iprotObject = ((grbit & ALLOWOBJECTS) == ALLOWOBJECTS);
		iprotScenario = ((grbit & ALLOWSCENARIOS) == ALLOWSCENARIOS);
		iprotFormatCells = ((grbit & ALLOWFORMATCELLS) == ALLOWFORMATCELLS);
		iprotFormatColumns = ((grbit & ALLOWFORMATCOLUMNS) == ALLOWFORMATCOLUMNS);
		iprotFormatRows = ((grbit & ALLOWFORMATROWS) == ALLOWFORMATROWS);
		iprotInsertColumns = ((grbit & ALLOWINSERTCOLUMNS) == ALLOWINSERTCOLUMNS);
		iprotInsertRows = ((grbit & ALLOWINSERTROWS) == ALLOWINSERTROWS);
		iprotInsertHyperlinks = ((grbit & ALLOWINSERTHYPERLINKS) == ALLOWINSERTHYPERLINKS);
		iprotDeleteColumns = ((grbit & ALLOWDELETECOLUMNS) == ALLOWDELETECOLUMNS);
		iprotDeleteRows = ((grbit & ALLOWDELETEROWS) == ALLOWDELETEROWS);
		iprotSelLockedCells = ((grbit & ALLOWSELLOCKEDCELLS) == ALLOWSELLOCKEDCELLS);
		iprotSort = ((grbit & ALLOWSORT) == ALLOWSORT);
		iprotAutoFilter = ((grbit & ALLOWAUTOFILTER) == ALLOWAUTOFILTER);
		iprotPivotTable = ((grbit & ALLOWPIVOTTABLES) == ALLOWPIVOTTABLES);
		iprotSelUnlockedCells = ((grbit & ALLOWSELUNLOCKEDCELLS) == ALLOWSELUNLOCKEDCELLS);
		;
	}

	/**
	 * update protection-style grbit variables + update record
	 */
	private void updateProtectionGrbit()
	{
		byte[] b = ByteTools.cLongToLEBytes( grbit );
		getData();
		data[19] = b[0];
		data[20] = b[1];
		data[21] = b[2];
		data[22] = b[3];
		parseProtectionGrbit();
	}

	/**
	 * sets or clears the set of enhanced protection option
	 *
	 * @param protectionOption @see WorkSheetHandle.iprotXXX constants
	 * @param set
	 */
	public void setProtectionOption( int protectionOption, boolean set )
	{
		if( set )
		{
			grbit = (grbit | protectionOption);
		}
		else
		{
			grbit = (grbit & ~protectionOption);
		}
		updateProtectionGrbit();
	}

	/**
	 * returns the desired set of enhanced protection setting
	 *
	 * @param protectionOption @see @WorkSheetHandle.iprotXXX constants
	 * @return
	 */
	public boolean getProtectionOption( int protectionOption )
	{
		return ((grbit & protectionOption) == protectionOption);
	}
}