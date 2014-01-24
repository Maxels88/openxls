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

/**
 * The SXFDB record specifies properties for a cache field within a PivotCache.
 *
 *
 * grbit:
 *  A - fAllAtoms (1 bit): A bit that specifies whether this cache field has a collection of cache items. If fSomeUnhashed is equal to 1, this value MUST be equal to 0.
 B - fSomeUnhashed (1 bit): Undefined, and MUST be ignored. If the fAllAtoms field is equal to 1, MUST be equal to 0.
 C - fUsed (1 bit): Undefined, and MUST be ignored.
 D - fHasParent (1 bit): A bit that specifies whether ifdbParent specifies a reference to a parent grouping cache field. For more information, see Grouping. If the fCalculatedField field is equal to 1, then this field MUST be equal to 0.
 E - fRangeGroup (1 bit): A bit that specifies whether this cache field is grouped by using numeric grouping or date grouping, as specified by Grouping. If this field is equal to 1, then this record MUST be followed by a sequence of SXString records, as specified by the GRPSXOPER rule. The quantity of SXString records is specified by csxoper. If this field is equal to 1, then this record MUST be followed by a sequence of records that conforms to the SXRANGE rule that specifies the grouping properties for the ranges of values.
 F - fNumField (1 bit): A bit that specifies whether the cache items in this cache field contain at least one numeric cache item, as specified by SXNum. If fDateInField is equal to 1, this field MUST be equal to 0.
 G - unused1 (1 bit): Undefined and MUST be ignored.
 H - fTextEtcField (1 bit): A bit that specifies whether the cache items contain text data. If fNumField is 1, this field MUST be ignored.
 I - fnumMinMaxValid (1 bit): A bit that specifies whether a valid minimum or maximum value can be computed for the cache field. MUST be equal to 1 if fDateInField or fNumField is equal to 1.
 J - fShortIitms (1 bit): A bit that specifies whether there are more than 255 cache items in this cache field. If catm is greater than 255, this value MUST be equal to 1; otherwise it MUST be 0.
 K - fNonDates (1 bit): A bit that specifies whether the cache items in this cache field contain values that are not time or date values. If this cache field is a grouping cache field, as specified by Grouping, then this field MUST be ignored. Otherwise, if fDateInField is equal to 1, then this field MUST be 0.
 L - fDateInField (1 bit): A bit that specifies whether the cache items in this cache field contain at least one time or date cache item, as specified by SXDtr. If fNonDates is equal to 1, then this field MUST be equal to 0.
 M - unused2 (1 bit): Undefined and MUST be ignored.
 N - fServerBased (1 bit): A bit that specifies whether this cache field is a server-based page field when the corresponding pivot field is on the page axis of the PivotTable view, as specified in source data.
 This value applies only to an ODBC PivotCache. MUST NOT be equal to 1 if fCantGetUniqueItems is equal to 1. If fCantGetUniqueItems is equal to 1, then the ODBC connection cannot provide a list of unique items for the cache field.
 MUST be 0 for a cache field in a non-ODBC PivotCache.
 O - fCantGetUniqueItems (1 bit): A bit that specifies whether a list of unique values for the cache field was not available while refreshing the source data. This field applies only to a PivotCache that uses ODBC source data and is intended to be used in conjunction with optimization features. For example, the application can optimize memory usage when populating PivotCache records if it has a list of unique values for a cache field before all the records are retrieved from the ODBC connection. Or, the application can determine the appropriate setting of fServerBased based on this value.
 MUST be 0 for fields in a non-ODBC PivotCache.
 P - fCalculatedField (1 bit): A bit that specifies whether this field is a calculated field. The formula (section 2.2.2) of the calculated field is stored in a directly following SXFormula record. If fHasParent is equal to 1, this field MUST be equal to 0.

 ifdbParent (2 bytes): An unsigned integer that specifies the cache field index,
 as specified by Cache Fields, of the grouping cache field for this cache field. MUST be greater than or equal to 0x0000 and less than or equal to the cfdbTot field of the SXDB record of this PivotCache. If fHasParent is equal to 0, then this field MUST be ignored. If fHasParent is equal to 1, and fRangeGroup is equal to 1, and the iByType field of the SXRng record of this cache field is greater than 0, then the fRangeGroup of the SXFDB record of the cache field specified by ifdbParent MUST be 1 and the iByType field of the SXRng record of the cache field
 specified by ifdbParent MUST be greater than the iByType field of the SXRng record of this cache field.

 ifdbBase (2 bytes): An unsigned integer that specifies the cache field index, as specified by Cache Fields, of the base cache field,
 as specified by Grouping, for the cache field specified by this record.
 MUST be greater than or equal to 0x0000 and less than the value of the cfdbdb field of the SXDB record of this PivotCache.
 If the cache field specified by this record is not a grouping cache field, then this field MUST be ignored.

 citmUnq (2 bytes): Undefined and MUST be ignored.

 csxoper (2 bytes): An unsigned integer that specifies the number of cache items in this cache field when this cache field is a grouping cache field,
 as specified by Grouping. There MUST be an equivalent number of sequences of records that conform to the GRPSXOPER rule
 following this record that specify the cache items. If the fRangeGroup field and the fCalculatedField field are equal to 0
 and this cache field corresponds to a source data entity, this field MUST be equal to 0.
 If the fRangeGroup field is equal to 1, this value MUST be greater than or equal to 1.

 cisxoper (2 bytes): An unsigned integer that specifies the number of cache items in the base cache field that are grouped by this cache field.
 There MUST be an equivalent number of SxIsxoper records following this record that specify which cache item in this cache
 field groups each of the cache items in the base cache field. For more information, see Grouping.

 catm (2 bytes): An unsigned integer that specifies the number of cache items in the collection sequences of records that conform to the
 SRCSXOPER rule in this cache field. If fAllAtoms is 0, then this field MUST be equal to 0x0000.
 If this cache field corresponds to source data entities then there MUST be an equal number of SRCSXOPER rules in this cache field.

 stFieldName (variable): An XLUnicodeString structure that specifies the name of the cache field. MUST be less than or equal to 255 characters long.


 GROUPING:  There are three different types of grouping: numeric grouping, date grouping, and discrete grouping.
 Numeric grouping combines numeric cache items into ranges of values. Date grouping combines date cache items into date ranges.
 Discrete grouping combines specifically selected cache items into groups.
 The cache field that contains the cache items that are to be grouped is called the base cache field.
 The resultant cache field that contains the groups of cache items is called the parent grouping cache field.
 Each group of cache items in the base cache field is associated with a single cache item in the parent grouping cache field.
 Often cache items in parent grouping cache fields can be further grouped, creating a hierarchy of parent grouping cache fields.
 The base cache field is at the lowest level of the hierarchy.
 *
 */

import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.UnsupportedEncodingException;

public class SxFDB extends XLSRecord implements XLSConstants, PivotCacheRecord
{
	private static final Logger log = LoggerFactory.getLogger( SxFDB.class );
	private static final long serialVersionUID = 9027599480633995587L;
	private short ifdbParent;
	private short ifdbBase;
	private short csxoper;
	private short cisxoper;
	private short catm;
	private short grbit;
	private String stFieldName;
	// significant bit fields 
	boolean fAllAtoms;
	boolean fRangeGroup;
	boolean fNumField;
	boolean fTextEtcField;
	boolean fnumMinMaxValid;
	boolean fNonDates;
	boolean fDateInField;
	boolean fCalculatedField;
	boolean fShortItms;

	// TODO: handle ranges/grouping and all the complications that it entails
	@Override
	public void init()
	{
		super.init();
		grbit = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
		// grouping-related
		ifdbParent = ByteTools.readShort( getByteAt( 2 ), getByteAt( 3 ) );    // specifies the cache field index, as specified by Cache Fields, of the grouping cache field for this cache field
		ifdbBase = ByteTools.readShort( getByteAt( 4 ), getByteAt( 5 ) );        // specifies the cache field index, as specified by Cache Fields, of the base cache field
		//ignored:  citmUnq= ByteTools.readShort(this.getByteAt(6),this.getByteAt(7)); // "Undefined and MUST be ignored."
		csxoper = ByteTools.readShort( getByteAt( 8 ), getByteAt( 9 ) );        // specifies the number of cache items in this cache field when this cache field is a grouping cache field,
		cisxoper = ByteTools.readShort( getByteAt( 10 ), getByteAt( 11 ) );    // specifies the number of cache items in the base cache field that are grouped by this cache field
		// + fHasParent, fRangeGroup ...
		catm = ByteTools.readShort( getByteAt( 12 ), getByteAt( 13 ) );        // number of cache items
		int cch = ByteTools.readShort( getByteAt( 14 ), getByteAt( 15 ) );
		if( cch > 0 )
		{        // 0xFFFF if none
			byte encoding = getByteAt( 16 );
			byte[] tmp = getBytesAt( 17, (cch) * (encoding + 1) );
			try
			{
				if( encoding == 0 )
				{
					stFieldName = new String( tmp, DEFAULTENCODING );
				}
				else
				{
					stFieldName = new String( tmp, UNICODEENCODING );
				}
			}
			catch( UnsupportedEncodingException e )
			{
				log.warn( "SXString.init: " + e, e );
			}
		}
		fAllAtoms = (grbit & 0x1) == 0x1;
		// KSC: TESTING
//if (!fAllAtoms)
//	Logger.logWarn("SXFDB: not all cache items are used");        
		fRangeGroup = (grbit & 0x10) == 0x10;            // A bit that specifies whether this cache field is grouped by using numeric grouping or date grouping, as specified by Grouping.
		fNumField = (grbit & 0x20) == 0x20;            // at least one numeric field (SxNum records follow)
		fTextEtcField = (grbit & 0x80) == 0x80;        // text field (SxString records follow)... if fNumField must be false
		fnumMinMaxValid = (grbit & 0x100) == 0x100;    // true date or num field
		fShortItms = (grbit & 0x200) == 0x200;            // true if > 255 cache items
		fNonDates = (grbit & 0x400) == 0x400;
		fDateInField = (grbit & 0x4000) == 0x4000;        // at least one date field (SxDtr follows)
		fCalculatedField = (grbit & 0x8000) == 0x8000;    // Sxformulas follow "A calculated field is a cache field (section 2.2.5.3.5) and does not correspond to a column in the source data (section 2.2.5.3.2). The values for a calculated field are calculated based on the formula specified for the calculated field"

			log.debug( "{}", toString() );
	}

	public String toString()
	{
		return String.format(
				"SXFDB -p: %d d: %d csxoper: %d cisxoper: %d catm: %d stFieldName: %s fAllAtoms? %b text? %b num? %b calculated? %b",
				ifdbParent,
				ifdbBase,
				csxoper,
				cisxoper,
				catm,
				stFieldName,
				fAllAtoms,
				fTextEtcField,
				fNumField,
				fCalculatedField );
	}

	/**
	 * return the bytes describing this record, including the header
	 *
	 * @return
	 */
	@Override
	public byte[] getRecord()
	{
		byte[] b = new byte[4];
		System.arraycopy( ByteTools.shortToLEBytes( getOpcode() ), 0, b, 0, 2 );
		System.arraycopy( ByteTools.shortToLEBytes( (short) getData().length ), 0, b, 2, 2 );
		return ByteTools.append( getData(), b );

	}

	private byte[] PROTOTYPE_BYTES = new byte[]{
//				1, 20,    /* grbit 5121 -- fAllAtoms no dates or nums...*/  
			-128, 4,    /* grbit 1152 -- the most basic setting i.e. no fAllAtoms ...*/
			0, 0, 0, 0, 1, 0,	    /* unique count -- ignored?? */
			0, 0, 0, 0, 	    /* cisxoper */
			0, 0,		/* catm */
			-1, -1
	};    // cch

	public static XLSRecord getPrototype()
	{
		SxFDB sxfdb = new SxFDB();
		sxfdb.setOpcode( SXFDB );
		sxfdb.setData( sxfdb.PROTOTYPE_BYTES );
		sxfdb.init();
		return sxfdb;
	}

	/**
	 * sets the number of cache items (column cells in range)
	 * TODO: only set fAllAtoms if cache item is source data and not calculated ...
	 *
	 * @param n the number of cache items is the number of cache items actually on axes
	 */
	public void setNCacheItems( int n )
	{
		catm = (short) n;                // If this cache field corresponds to source data entities then there MUST be an equal number of SRCSXOPER rules in this cache field.
		byte[] b = ByteTools.shortToLEBytes( catm );
		getData()[12] = b[0];
		getData()[13] = b[1];

		if( n > 0 )
		{
			// set that this has atoms (placed on pivot table axes)
			fAllAtoms = true;
			grbit |= 0x1;
		}
		else
		{
			fAllAtoms = false;
			grbit &= ~0x1;    // turn off fAllAtoms ...
		}
		b = ByteTools.shortToLEBytes( grbit );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}

	/**
	 * returns the number of cache items (column cells in range)
	 *
	 * @return
	 */
	public int getNCacheItems()
	{
		return catm;
	}

	/**
	 * sets the cache field name == cache row cell label
	 *
	 * @param s
	 */
	public void setCacheField( String s )
	{
		stFieldName = s;
		byte[] strbytes = new byte[0];
		if( stFieldName != null )
		{
			try
			{
				strbytes = stFieldName.getBytes( DEFAULTENCODING );
			}
			catch( UnsupportedEncodingException e )
			{
				log.warn( "SxFDB: " + e, e );
			}
		}

		short cch = (short) strbytes.length;
		byte[] nm = ByteTools.shortToLEBytes( cch );
		byte[] data = new byte[14];
		System.arraycopy( getData(), 0, data, 0, 14 );
		byte[] newdata = new byte[cch + 3];    // account for encoding bytes + cch
		System.arraycopy( nm, 0, newdata, 0, 2 );
		System.arraycopy( strbytes, 0, newdata, 3, cch );

		data = ByteTools.append( newdata, data );
		setData( data );
	}

	/**
	 * returns the cache field name == cache row cell label
	 */
	public String getCachefield()
	{
		return stFieldName;
	}

	/**
	 * sets the type of the cache items within this cache field (== the column cells below the row header field in the pivot cache range)
	 *
	 * @param type
	 */
	public void setCacheItemsType( int type )
	{
		switch( type )
		{
			case XLSConstants.TYPE_STRING:
				fTextEtcField = true;
				fNumField = false;
				fNonDates = true;
				fDateInField = false;
				fnumMinMaxValid = false;
				fCalculatedField = false;
				break;
			case XLSConstants.TYPE_FP:
			case XLSConstants.TYPE_INT:
			case XLSConstants.TYPE_DOUBLE:
				fTextEtcField = false;
				fNumField = true;
				fNonDates = true;
				fDateInField = false;
				fnumMinMaxValid = true;
				fCalculatedField = false;
				break;
			case TYPE_FORMULA:
				fCalculatedField = true;
				break;
			//TYPE_BOOLEAN = 4, 
			case 6:// date
				fTextEtcField = false;
				fNumField = false;
				fNonDates = false;
				fDateInField = true;
				fnumMinMaxValid = true;
				fCalculatedField = false;
		}
		if( fNumField )
		{
			grbit |= 0x20;
		}
		else
		{
			grbit &= ~0x20;
		}
		if( fTextEtcField )
		{
			grbit |= 0x80;
		}
		else
		{
			grbit &= ~0x80;
		}
		if( fnumMinMaxValid )
		{
			grbit |= 0x100;
		}
		else
		{
			grbit &= ~0x100;
		}
		if( fNonDates )
		{
			grbit |= 0x400;
		}
		else
		{
			grbit &= ~0x400;
		}
		if( fDateInField )
		{
			grbit |= 0x4000;
		}
		else
		{
			grbit &= ~0x4000;
		}
		if( fCalculatedField )
		{
			grbit |= 0x8000;
		}
		else
		{
			grbit &= ~0x8000;
		}

		byte[] b = ByteTools.shortToLEBytes( grbit );
		getData()[0] = b[0];
		getData()[1] = b[1];
	}
}
