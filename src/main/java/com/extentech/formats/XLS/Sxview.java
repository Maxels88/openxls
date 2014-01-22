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

import com.extentech.ExtenXLS.ExcelTools;
import com.extentech.formats.XLS.SxAddl.SxcView;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;

/**
 * <b>SXVIEW B0h: This record contains top-level pivot table information.</b><br>
 * <p><pre>
 * ref (8 bytes): A Ref8U structure that specifies the PivotTable report body. For more information, see Location and Body.
 * rwFirstHead (2 bytes): An RwU structure that specifies the first row of the row area.
 * MUST be 1 if none of the axes are assigned in this PivotTable view.
 * Otherwise, the value MUST be greater than or equal to ref.rwFirst.
 * <p/>
 * rwFirstData (2 bytes): An RwU structure that specifies the first row of the data area.
 * MUST be 1 if none of the axes are assigned in this PivotTable view.
 * Otherwise, it MUST be equal to the value as specified by the following formula:
 * <p/>
 * rwFirstData = rwFirstHead + cDimCol
 * <p/>
 * colFirstData (2 bytes): A ColU structure that specifies the first column of the data area.
 * It MUST be 1 if none of the axes are assigned in this PivotTable view.
 * Otherwise, the value MUST be greater than or equal to ref.colFirst, and if the value of cDimCol or cDimData is not zero,
 * it MUST be less than or equal to ref.colLast.
 * <p/>
 * iCache (2 bytes): A signed integer that specifies the zero-based index of an SXStreamID record in the Globals substream.
 * MUST be greater than or equal to zero and less than the number of SXStreamID records in the Globals substream.
 * <p/>
 * reserved (2 bytes): MUST be zero, and MUST be ignored.
 * <p/>
 * sxaxis4Data (2 bytes): An SXAxis structure that specifies the default axis for the data field.
 * Either the sxaxis4Data.sxaxisRw field MUST be 1 or the sxaxis4Data.sxaxisCol field MUST be 1.
 * The sxaxis4Data.sxaxisPage field MUST be 0 and the sxaxis4Data.sxaxisData field MUST be 0.
 * <p/>
 * ipos4Data (2 bytes): A signed integer that specifies the row or column position for the data field in the PivotTable view.
 * The sxaxis4Data field specifies whether this is a row or column position.
 * MUST be greater than or equal to -1 and less than or equal to 0x7FFF. A value of -1 specifies the default position.
 * <p/>
 * cDim (2 bytes): A signed integer that specifies the number of pivot fields in the PivotTable view.
 * MUST equal the number of Sxvd records following this record.
 * MUST equal the number of fields in the associated PivotCache specified by iCache.
 * <p/>
 * cDimRw (2 bytes): An unsigned integer that specifies the number of fields on the row axis of the PivotTable view.
 * MUST be less than or equal to 0x7FFF. MUST equal the number of array elements in the SxIvd record in this PivotTable view that contain row items.
 * <p/>
 * cDimCol (2 bytes): An unsigned integer that specifies the number of fields on the column axis of the PivotTable view.
 * MUST be less than or equal to 0x7FFF.
 * MUST equal the number of array elements in the SxIvd record in this PivotTable view that contain column items.
 * <p/>
 * cDimPg (2 bytes): An unsigned integer that specifies the number of page fields in the PivotTable view.
 * MUST be less than or equal to 0x7FFF.
 * MUST equal the number of array elements in the SXPI record in this PivotTable view.
 * <p/>
 * cDimData (2 bytes): A signed integer that specifies the number of data fields in the PivotTable view.
 * MUST be greater than or equal to zero and less than or equal to 0x7FFF.
 * MUST equal the number of SXDI records in this PivotTable view.
 * <p/>
 * cRw (2 bytes): An unsigned integer that specifies the number of pivot lines in the row area of the PivotTable view.
 * MUST be less than or equal to 0x7FFF.
 * MUST equal the number of array elements in the first SXLI record in this PivotTable view.
 * <p/>
 * cCol (2 bytes): An unsigned integer that specifies the number of pivot lines in the column area of the PivotTable view.
 * MUST equal the number of array elements in the second SXLI record in this PivotTable view.
 * <p/>
 * A - fRwGrand (1 bit): A bit that specifies whether the PivotTable contains grand totals for rows.
 * MUST be 0 if none of the axes have been assigned in this PivotTable view.
 * <p/>
 * B - fColGrand (1 bit): A bit that specifies whether the PivotTable contains grand totals for columns.
 * MUST be 1 if none of the axes are assigned in this PivotTable view.
 * <p/>
 * C - unused1 (1 bit): Undefined and MUST be ignored.
 * <p/>
 * D - fAutoFormat (1 bit): A bit that specifies whether the PivotTable has AutoFormat applied.
 * <p/>
 * E - fAtrNum (1 bit): A bit that specifies whether the PivotTable has number AutoFormat applied.
 * <p/>
 * F - fAtrFnt (1 bit): A bit that specifies whether the PivotTable has font AutoFormat applied.
 * <p/>
 * G - fAtrAlc (1 bit): A bit that specifies whether the PivotTable has alignment AutoFormat applied.
 * <p/>
 * H - fAtrBdr (1 bit): A bit that specifies whether the PivotTable has border AutoFormat applied.
 * <p/>
 * I - fAtrPat (1 bit): A bit that specifies whether the PivotTable has pattern AutoFormat applied.
 * <p/>
 * J - fAtrProc (1 bit): A bit that specifies whether the PivotTable has width/height AutoFormat applied.
 * <p/>
 * unused2 (6 bits): Undefined and MUST be ignored.
 * <p/>
 * itblAutoFmt (2 bytes): An AutoFmt8 structure that specifies the PivotTable AutoFormat.
 * If the value of itblAutoFmt in the associated SXViewEx9 record is not 1, this field is overridden by the value of itblAutoFmt in the associated SXViewEx9.
 * <p/>
 * cchTableName (2 bytes): An unsigned integer that specifies the length, in characters, of stTable.
 * MUST be greater than or equal to zero and less than or equal to 0x00FF.
 * <p/>
 * cchDataName (2 bytes): An unsigned integer that specifies the length, in characters of stData.
 * MUST be greater than zero and less than or equal to 0x00FE.
 * <p/>
 * stTable (variable): An XLUnicodeStringNoCch structure that specifies the name of the PivotTable.
 * The length of this field is specified by cchTableName.
 * <p/>
 * stData (variable): An XLUnicodeStringNoCch structure that specifies the name of the data field.
 * The length of this field is specified by cchDataName.
 * <p/>
 * </pre>
 */
public class Sxview extends XLSRecord implements XLSConstants
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 2639291289806138985L;
	short rwFirst = 0x0; // First Row of Pivot Table
	short rwLast = 0x0; // Last Row of Pivot Table
	short colFirst = 0x0; // First Column of Pivot Table
	short colLast = 0x0; // Last Column of Pivot Table
	short rwFirstHead = 0x0; // First Row containing Pivot Table headings
	short rwFirstData = 0x0; // First Row containing Pivot Table data
	short colFirstData = 0x0; // First Column containing Pivot Table data
	short iCache = 0x0; // Index to the Cache
	//  offset 20, length 2 is reserved, must be 0 (zero)
	short sxaxis4Data = 0x0; // Default axis for a data field
	short ipos4Data = 0x0; // Default position for a data field
	short cDim = 0x0; // Number of fields
	short cDimRw = 0x0; // Number of row fields;
	short cDimCol = 0x0; // Number of column fields
	short cDimPg = 0x0; // Number of page fields;
	short cDimData = 0x0; // Number of Data fields
	short cRw = 0x0; // Number of data rows
	short cCol = 0x0; // Number of data columns
	byte grbit1 = 0x0; // if you dont know what this is....
	byte grbit2 = 0x0; // if you dont know what this is....
	short itblAutoFmt = 0x0; // Index to the Pivot Table autoformat
	short cchName = 0x0; // Length the Pivot Table name
	short cchData = 0x0; // Length of the data field name
	byte[] rgch;// PivotTableName followed by the name of a data field

	/* The following member variables are populated from parsing the
	   grbit field
	*/ boolean fRwGrand = false;  // = 1 if the Pivot Table contains grand totals for rows
	boolean fColGrand = false; // = 1 if the Pivot Table contains grand totals for Columns
	// bit 2, mask 0004h is reserved, must be 0 (zero)
	boolean fAutoFormat = false; // = 1 if the Pivot table has an autoformat applied
	boolean fWH = false; // = 1 if the width/height autoformat is applied
	boolean fFont = false; // = 1 if the font autoformat is applied
	boolean fAlign = false; // = 1 if the alignment autoformat is applied
	boolean fBorder = false; // = 1 if the border autoformat is applied
	boolean fPattern = false; // = 1 if the pattern autoformat is applied
	boolean fNumber = false; // = 1 if the number autoformat is applied

	/*
		The following member variables are derived from the rgch field
	*/ String PivotTableName = null;
	String DataFieldName = null;

	private ArrayList<BiffRec> subRecs = new ArrayList();

	private byte[] PROTOTYPE_BYTES = new byte[]{
			0, 0, 0, 0, 0, 0, 0, 0, /* ref */
			0, 0,	/* rwfirstHead */
			0, 0,   /* rwfirstData */
			0, 0,	/* colfirstHead */
			0, 0,	/* iCache */
			0, 0,	/* reserved */
			1, 0,	/* saxis4Data */
			-1, -1,	/* ipos4Data */
			0, 0,	/* cDim */
			0, 0, 	/* cDimRw */
			0, 0,	/* cDimCol */
			0, 0, 	/* cDimPg */
			0, 0,	/* cDimData */
			0, 0,	/* cRw */
			0, 0,	/* cCol */
			11, 2,	/* grbit + reserved */
			1, 0,   /* itblAutoFmt */
			0, 0,   /* cchTableName */
			0, 0,   /* cchDataName */

	};

	public static XLSRecord getPrototype()
	{
		Sxview sx = new Sxview();
		sx.setOpcode( SXVIEW );
		sx.setData( sx.PROTOTYPE_BYTES );
		sx.init();
		return sx;
	}

	@Override
	public void init()
	{
		super.init();
		if( this.getLength() <= 0 )
		{  // Is this record populated?
			if( DEBUGLEVEL > -1 )
			{
				Logger.logInfo( "no data in SXVIEW" );
			}
		}
		else
		{ // parse out all the fields
			rwFirst = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
			rwLast = ByteTools.readShort( this.getByteAt( 2 ), this.getByteAt( 3 ) );
			colFirst = ByteTools.readShort( this.getByteAt( 4 ), this.getByteAt( 5 ) );
			colLast = ByteTools.readShort( this.getByteAt( 6 ), this.getByteAt( 7 ) );
			rwFirstHead = ByteTools.readShort( this.getByteAt( 8 ), this.getByteAt( 9 ) );
			rwFirstData = ByteTools.readShort( this.getByteAt( 10 ), this.getByteAt( 11 ) );
			colFirstData = ByteTools.readShort( this.getByteAt( 12 ), this.getByteAt( 13 ) );
			iCache = ByteTools.readShort( this.getByteAt( 14 ), this.getByteAt( 15 ) );

			// 16 & 17 - reserved must be zero

			sxaxis4Data = ByteTools.readShort( this.getByteAt( 18 ), this.getByteAt( 19 ) );
			ipos4Data = ByteTools.readShort( this.getByteAt( 20 ), this.getByteAt( 21 ) );
			cDim = ByteTools.readShort( this.getByteAt( 22 ), this.getByteAt( 23 ) );
			cDimRw = ByteTools.readShort( this.getByteAt( 24 ), this.getByteAt( 25 ) );
			cDimCol = ByteTools.readShort( this.getByteAt( 26 ), this.getByteAt( 27 ) );
			cDimPg = ByteTools.readShort( this.getByteAt( 28 ), this.getByteAt( 29 ) );
			cDimData = ByteTools.readShort( this.getByteAt( 30 ), this.getByteAt( 31 ) );
			cRw = ByteTools.readShort( this.getByteAt( 32 ), this.getByteAt( 33 ) );
			cCol = ByteTools.readShort( this.getByteAt( 34 ), this.getByteAt( 35 ) );
			grbit1 = this.getByteAt( 37 );
			grbit2 = this.getByteAt( 36 );

			this.initGrbit(); // note the manual hibyting
			itblAutoFmt = ByteTools.readShort( this.getByteAt( 38 ), this.getByteAt( 39 ) );
			cchName = ByteTools.readShort( this.getByteAt( 40 ), this.getByteAt( 41 ) );
			cchData = ByteTools.readShort( this.getByteAt( 42 ), this.getByteAt( 43 ) );
			int fullnamelen = (int) ((int) cchName) + ((int) cchData);
			rgch = new byte[fullnamelen];
			int pos = 44;
			if( cchName > 0 )
			{
				//A - fHighByte (1 bit): A bit that specifies whether the characters in rgb are double-byte characters.
				// 0x0  All the characters in the string have a high byte of 0x00 and only the low bytes are in rgb.
				// 0x1  All the characters in the string are saved as double-byte characters in rgb.
				// reserved (7 bits): MUST be zero, and MUST be ignored.

				byte encoding = this.getByteAt( pos++ );

				byte[] tmp = this.getBytesAt( pos, (cchName) * (encoding + 1) );
				try
				{
					if( encoding == 0 )
					{
						PivotTableName = new String( tmp, DEFAULTENCODING );
					}
					else
					{
						PivotTableName = new String( tmp, UNICODEENCODING );
					}
				}
				catch( UnsupportedEncodingException e )
				{
					Logger.logInfo( "encoding PivotTable name in Sxview: " + e );
				}
				pos += cchName * (encoding + 1);
			}
			if( cchData > 0 )
			{
				byte encoding = this.getByteAt( pos++ );
				byte[] tmp = this.getBytesAt( pos, (cchData) * (encoding + 1) );
				try
				{
					if( encoding == 0 )
					{
						DataFieldName = new String( tmp, DEFAULTENCODING );
					}
					else
					{
						DataFieldName = new String( tmp, UNICODEENCODING );
					}
				}
				catch( UnsupportedEncodingException e )
				{
					Logger.logInfo( "encoding PivotTable name in Sxview: " + e );
				}
			}
		}
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "SXVIEW: name:" + this.getTableName() + " iCache:" + iCache + " cDim:" + cDim + " cDimRw:" + cDimRw + " cDimCol:" + cDimCol + " cDimPg:" + cDimPg + " cDimData:" + cDimData + " cRw:" + cRw + " cCol:" + cCol + " datafieldname:" + DataFieldName );
		}
	}

	/**
	 * specifies the default axis for the data field:
	 * <li>1- row
	 * <li>2= col
	 * <li>4= page
	 * <li>8= data
	 *
	 * @param s
	 */
	protected void setAxis4Data( short s )
	{
		sxaxis4Data = s;
		byte[] b = ByteTools.shortToLEBytes( sxaxis4Data );
		System.arraycopy( b, 0, data, 18, 2 );
	}

	/**
	 * retieves the default axis for the data field:
	 * <li>1- row
	 * <li>2= col
	 * <li>4= page
	 * <li>8= data
	 */
	public short getSxaxis4Data()
	{
		return sxaxis4Data;
	}

	/**
	 * A signed integer that specifies the row or column position for the data field in the PivotTable view.
	 * <br>see getSxaxis4Data to determine whether this is a row or column position
	 *
	 * @param s
	 */
	protected void setIpos4Data( short s )
	{
		ipos4Data = s;
		byte[] b = ByteTools.shortToLEBytes( ipos4Data );
		System.arraycopy( b, 0, data, 20, 2 );
	}

	/**
	 * A signed integer that specifies the row or column position for the data field in the PivotTable view.
	 * <br>see getSxaxis4Data to determine whether this is a row or column position
	 */
	public short getIpos4Data()
	{
		return ipos4Data;
	}

	/**
	 * sets the number of pivot fields (i.e. columns in the source data) for this pivot table
	 * <br>NOTE: this method wipes any existing data of the pivot table
	 *
	 * @param s
	 */
	public void setNPivotFields( short s )
	{
		cDim = s;    // number of pivot fields MUST==# of SxVd records
		byte[] b = ByteTools.shortToLEBytes( cDim );
		System.arraycopy( b, 0, data, 22, 2 );
		// remove ALL field-related records
		for( int i = 0; i < subRecs.size(); i++ )
		{
			BiffRec br = ((BiffRec) subRecs.get( i ));
			if( br.getOpcode() != SXEX )
			{
				this.getSheet().removeRecFromVec( br );
				subRecs.remove( i );
				i--;
			}
			else
			{
				break;
			}
		}
		// reset variables
		cDimRw = 0;
		cDimCol = 0;
		cDimPg = 0;
		cDimData = 0;
		cRw = 0;
		cCol = 0;
		// now add the required records for each field
		int zz = this.getRecordIndex() + 1;
		for( int i = 0; i < cDim; i++ )
		{
			Sxvd svd = (Sxvd) Sxvd.getPrototype();    // for each pivot field (which goes on an axis)
			svd.setSheet( this.getSheet() );
			this.getSheet().getSheetRecs().add( zz++, svd );
			this.subRecs.add( i * 2, svd );
			SxVdEX svdex = (SxVdEX) SxVdEX.getPrototype();
			svdex.setSheet( this.getSheet() );
			this.getSheet().getSheetRecs().add( zz++, svdex );
			this.subRecs.add( (i * 2) + 1, svdex );
		}
	}

	/**
	 * adds the pivot field corresponding to cache field at index fieldIndex to the desired axis
	 * <br>A pivot field is a cache field that has been added to the pivot table
	 * <br>Pivot fields are defined by SXVD and associated records
	 * <br>the SXVD record stores which axis the field is on
	 * <br>there are cDim pivot fields on the pivot table view
	 * <li>1= row
	 * <li>2= col
	 * <li>4= page
	 * <li>8= data
	 *
	 * @param axis
	 * @param fieldIndex 0-based pivot field index
	 * @see SxVd.AXIS_ constants
	 */
	public Sxvd addPivotFieldToAxis( int axis, int fieldIndex )
	{
		int zz = this.getRecordIndex() + 1;
		SxVdEX sxvdex = (SxVdEX) getSubRec( SXVDEX, fieldIndex );    // end of last pivot field set (PIVOTVD rule)
		if( sxvdex != null )
		{
			zz = sxvdex.getRecordIndex() + 1;
		}

		Sxvd sxvd = (Sxvd) Sxvd.getPrototype();    // for each pivot field (which goes on an axis)
		sxvd.setSheet( this.getSheet() );
		this.getSheet().getSheetRecs().add( zz++, sxvd );
		this.subRecs.add( cDim * 2, sxvd );
		SxVdEX svdex = (SxVdEX) SxVdEX.getPrototype();
		svdex.setSheet( this.getSheet() );
		this.getSheet().getSheetRecs().add( zz++, svdex );
		this.subRecs.add( (cDim * 2) + 1, svdex );
		cDim++;
		sxvd.setAxis( axis );
		return sxvd;

/*    	
    	Sxvd sxvd= (Sxvd) getSubRec(SXVD, fieldIndex);
    	if (sxvd!=null) {
    		// got it
    		sxvd.setAxis(axis);
        	return sxvd;        	
        }
        return null;*/
	}

	/**
	 * returns the number of pivot fields (==columns in the pivot table data range)
	 *
	 * @return
	 */
	public short getNPivotFields()
	{
		return cDim;
	}

	/**
	 * adds a pivot field  (0-based index) to the ROW axis
	 *
	 * @param fieldNumber
	 */
	public void addRowField( int fieldNumber )
	{
		cDimRw++;    // MUST==# row elements in SxIVD
		byte[] b = ByteTools.shortToLEBytes( cDimRw );
		System.arraycopy( b, 0, data, 24, 2 );
		Sxivd sxivd = (Sxivd) getSubRec( SXIVD, (cDimCol > 0) ? 1 : 0 );
		if( sxivd == null )
		{
			int zz = getPivotRecordInsertionIndexes( SXIVD, 1, -1 );
			if( zz > 0 )
			{ // should!!!
				sxivd = (Sxivd) Sxivd.getPrototype();
				sxivd.setSheet( this.getSheet() );
				this.getSheet().getSheetRecs().add( subRecs.get( zz ).getRecordIndex(), sxivd );
				this.subRecs.add( zz, sxivd );
			}
		}
		sxivd.addField( fieldNumber );
	}

	/**
	 * adds a pivot field (0-based index) to the COL axis
	 */
	public void addColField( int fieldNumber )
	{
		cDimCol++;
		byte[] b = ByteTools.shortToLEBytes( cDimCol );
		System.arraycopy( b, 0, data, 26, 2 );
		Sxivd sxivd = (Sxivd) getSubRec( SXIVD, 0 );
		if( sxivd == null )
		{
			int zz = getPivotRecordInsertionIndexes( SXIVD, 0, -1 );
			if( zz > 0 )
			{ // should!!
				sxivd = (Sxivd) Sxivd.getPrototype();
				sxivd.setSheet( this.getSheet() );
				this.getSheet().getSheetRecs().add( subRecs.get( zz ).getRecordIndex(), sxivd );
				this.subRecs.add( zz, sxivd );
			}
		}
		sxivd.addField( fieldNumber );
	}

	/**
	 * returns the number of pivot fields on the ROW axis
	 *
	 * @return
	 */
	public short getCDimRw()
	{
		return cDimRw;
	}

	/**
	 * returns the number of pivot fields on the COL axis
	 *
	 * @return
	 */
	public short getCDimCol()
	{
		return cDimCol;
	}

	/**
	 * adds a pivot field to the page axis
	 *
	 * @param strref
	 */
	public void addPageField( int fieldIndex, int itemIndex )
	{
		cDimPg++;
		byte[] b = ByteTools.shortToLEBytes( cDimPg );
		System.arraycopy( b, 0, data, 28, 2 );
		SxPI sxpi = (SxPI) getSubRec( SXPI, -1 );
		if( sxpi == null )
		{
			int zz = getPivotRecordInsertionIndexes( SXPI, -1, -1 );
			if( zz > 0 )
			{ // should!!!
				sxpi = (SxPI) SxPI.getPrototype();
				sxpi.setSheet( this.getSheet() );
				this.getSheet().getSheetRecs().add( subRecs.get( zz ).getRecordIndex(), sxpi );
				this.subRecs.add( zz, sxpi );
			}
		}
		sxpi.addPageField( fieldIndex, itemIndex );
	}

	/**
	 * returns the number of fields on the page axis
	 *
	 * @return
	 */
	public short getCDimPg()
	{
		return cDimPg;
	}

	/**
	 * adds a pivot field to the DATA axis
	 *
	 * @param fieldIndex
	 * @param aggregateFunction
	 */
	public void addDataField( int fieldIndex, String aggregateFunction, String name )
	{
		cDimData++;
		byte[] b = ByteTools.shortToLEBytes( cDimData );
		System.arraycopy( b, 0, data, 30, 2 );
		SxDI sxdi = null;
/*        SxDI sxdi= (SxDI) getSubRec(SXDI, -1);
        if (sxdi==null) {           */
		int zz = getPivotRecordInsertionIndexes( SXDI, -1, -1 );
		if( zz > 0 )
		{ // should!!!
			sxdi = (SxDI) SxDI.getPrototype();
			sxdi.setSheet( this.getSheet() );
			this.getSheet().getSheetRecs().add( subRecs.get( zz ).getRecordIndex(), sxdi );
			this.subRecs.add( zz, sxdi );
		}
/*	    }*/
		sxdi.addDataField( fieldIndex, aggregateFunction, name );
	}

	/**
	 * returns the number of pivot fields on the data axis
	 *
	 * @return
	 */
	public short getCDimData()
	{
		return cDimData;
	}

	/**
	 * adds a pivot item to the end of the list of items on this axis (ROW, COL, DATA or PAGE)
	 *
	 * @param axis      Axis int: ROW, COL, DATA or PAGE
	 * @param itemType  one of:
	 *                  <li>0x0000		itmtypeData	    A data value
	 *                  <li>0x0001	    itmtypeDEFAULT	Default subtotal for the pivot field
	 *                  <li>0x0002	    itmtypeSUM	    Sum of values in the pivot field
	 *                  <li>0x0003		itmtypeCOUNTA	Count of values in the pivot field
	 *                  <li>0x0004	    itmtypeAVERAGE	Average of values in the pivot field
	 *                  <li>0x0005	    itmtypeMAX	    Max of values in the pivot field
	 *                  <li>0x0006	    itmtypeMIN	    Min of values in the pivot field
	 *                  <li>0x0007	    itmtypePRODUCT  Product of values in the pivot field
	 *                  <li>0x0008	    itmtypeCOUNT	Count of numbers in the pivot field
	 *                  <li>0x0009	    itmtypeSTDEV	Statistical standard deviation (estimate) of the pivot field
	 *                  <li>0x000A	    itmtypeSTDEVP	Statistical standard deviation (entire population) of the pivot field
	 *                  <li>0x000B		itmtypeVAR	    Statistical variance (estimate) of the pivot field
	 *                  <li>0x000C	    itmtypeVARP	    Statistical variance (entire population) of the pivot field
	 * @param cacheItem A cache item index in the cache field associated with the pivot field, as specified by Cache Items.
	 */
	public void addPivotItem( Sxvd axis, int itemType, int cacheItem )
	{
		int n = axis.getNumItems();        // Axis record = Sxvd.  Identifies pivot field on axis. Pivot items records follow until SXVDEX.
		axis.setNumItems( n + 1 );
		int axisIndex = getSubRecIndex( axis );
		int zz = getPivotRecordInsertionIndexes( SXVI, n, axisIndex );    // get LAST pivot field on axis
		Sxvi sxvi = (Sxvi) Sxvi.getPrototype();                // pivot item record
		sxvi.setItemType( itemType );
		sxvi.setCacheItem( cacheItem );
		sxvi.setSheet( this.getSheet() );
		this.getSheet().getSheetRecs().add( subRecs.get( zz ).getRecordIndex(), sxvi );
		subRecs.add( zz, sxvi );
	 	/*if (axis.getAxisType()==Sxvd.AXIS_ROW) {
     	} else if (axis.getAxisType()==Sxvd.AXIS_COL) {
     	} else if (axis.getAxisType()==Sxvd.AXIS_PAGE) {
     		
     	} else if (axis.getAxisType()==Sxvd.AXIS_DATA) {
     		
     	}*/
		if( cacheItem != -1 )
		{
			PivotCache pc = this.getWorkBook().getPivotCache();    // TODO should this be here or in Sxstream?
			pc.addCacheItem( iCache, cacheItem );    // adds the cache item to the "used" list, as it were
		}
	}

	/**
	 * sets the number of pivot items or lines on the ROW axis
	 *
	 * @param strref
	 */
	public void addPivotLineToROWAxis( int repeat, int nLines, int type, short[] indexes )
	{
		cRw++;
		byte[] b = ByteTools.shortToLEBytes( cRw );
		System.arraycopy( b, 0, data, 32, 2 );
		/**
		 If the value of either of the cRw or cCol fields of the associated SxView is greater than zero,
		 then two records of this type MUST exist in the file for the associated SxView.
		 The first record contains row area pivot lines and the second record contains column area pivot lines.		*/
		Sxli sxli = (Sxli) getSubRec( SXLI, 0 );
		if( sxli == null )
		{
			int zz = getPivotRecordInsertionIndexes( SXLI, 0, -1 );
			if( zz > 0 )
			{ // should!!!
				sxli = (Sxli) Sxli.getPrototype( this.getWorkBook() );
				sxli.setSheet( this.getSheet() );
				this.getSheet().getSheetRecs().add( subRecs.get( zz ).getRecordIndex(), sxli );
				this.subRecs.add( zz, sxli );
				sxli.addField( repeat, nLines, type, indexes );
				sxli = (Sxli) Sxli.getPrototype( this.getWorkBook() );
				sxli.setSheet( this.getSheet() );
				this.getSheet().getSheetRecs().add( subRecs.get( zz + 1 ).getRecordIndex(), sxli );
				this.subRecs.add( zz + 1, sxli );
			}
		}
		else
		{
			sxli.addField( repeat, nLines, type, indexes );
		}
	}

	/**
	 * returns the number of pivot items or lines on the ROW axis
	 *
	 * @return
	 */
	public short getCRw()
	{
		return cRw;
	}

	protected boolean hasRowPivotItemsRecord()
	{
		return (getSubRec( SXLI, 0 ) != null);
	}

	/**
	 * set the number of pivot items or lines on the COL axis
	 *
	 * @param strref
	 */
	public void addPivotLineToCOLAxis( int repeat, int nLines, int type, short[] indexes )
	{
		cCol++;
		byte[] b = ByteTools.shortToLEBytes( cCol );
		System.arraycopy( b, 0, data, 34, 2 );
		Sxli sxli = (Sxli) getSubRec( SXLI, 1 );
		if( sxli == null )
		{
			int zz = getPivotRecordInsertionIndexes( SXLI, 0, -1 );
			if( zz == -1 )
			{ // shouldn't!
				sxli = (Sxli) Sxli.getPrototype();
				sxli.setSheet( this.getSheet() );
				this.getSheet().getSheetRecs().add( subRecs.get( zz ).getRecordIndex(), sxli );
				this.subRecs.add( zz, sxli );
			}
		}
		sxli.addField( repeat, nLines, type, indexes );
	}

	/**
	 * returns the number of pivot items or lines on the COL axis
	 *
	 * @return
	 */
	public short getCCol()
	{
		return cCol;
	}

	/**
	 * iCache links this pivot table to a pivot data cache
	 *
	 * @param s
	 */
	public void setICache( short s )
	{
		iCache = s;
		byte[] b = ByteTools.shortToLEBytes( iCache );
		System.arraycopy( b, 0, data, 14, 2 );
	}

	/**
	 * iCache links this pivot table to a pivot data cache
	 */
	public short getICache()
	{
		return iCache;
	}

	/**
	 * Init the grbit, populate the member variables of the SXVIEW object.
	 * This can be called both at initial init, and later to reup the boolean fields
	 * after changes to the grbit.
	 */
	private void initGrbit()
	{
		if( (grbit1 & 0x1) == 0x1 )
		{
			fRwGrand = true;
		}
		else
		{
			fRwGrand = false;
		}
		if( (grbit1 & 0x2) == 0x2 )
		{
			fColGrand = true;
		}
		else
		{
			fColGrand = false;
		}
		if( (grbit1 & 0x8) == 0x8 )
		{
			fAutoFormat = true;
		}
		else
		{
			fAutoFormat = false;
		}
		if( (grbit1 & 0x10) == 0x10 )
		{
			fWH = true;
		}
		else
		{
			fWH = false;
		}
		if( (grbit1 & 0x20) == 0x20 )
		{
			fFont = true;
		}
		else
		{
			fFont = false;
		}
		if( (grbit1 & 0x40) == 0x40 )
		{
			fAlign = true;
		}
		else
		{
			fAlign = false;
		}
		if( (grbit1 & 0x80) == 0x80 )
		{
			fBorder = true;
		}
		else
		{
			fBorder = false;
		}
		if( (grbit2 & 0x1) == 0x1 )
		{
			fPattern = true;
		}
		else
		{
			fPattern = false;
		}
		if( (grbit2 & 0x2) == 0x2 )
		{
			fNumber = true;
		}
		else
		{
			fNumber = false;
		}
		this.getData()[36] = grbit2;
		this.getData()[37] = grbit1;
	}

	/**
	 * store associated records for ease of lookup
	 *
	 * @param r
	 */
	public void addSubrecord( BiffRec r )
	{
		subRecs.add( r );
	}
    /*
        The following methods all change or get grbit fields
    */

	/**
	 * specifies whether the PivotTable contains grand totals for rows.
	 * MUST be 0 if none of the axes have been assigned in this PivotTable view.
	 */
	public void setFRwGrand( boolean b )
	{
		if( b != fRwGrand )
		{
			if( b == false )
			{
				grbit1 = (byte) (grbit1 & 0xFE);
			}
			else
			{
				grbit1 = (byte) (grbit1 | 0x1);
			}
			this.initGrbit();
		}
	}

	/**
	 * specifies whether the PivotTable contains grand totals for rows.
	 * MUST be 0 if none of the axes have been assigned in this PivotTable view.
	 */
	public boolean getFRwGrand()
	{
		return fRwGrand;
	}

	/**
	 * specifies whether the PivotTable contains grand totals for columns.
	 * MUST be 1 if none of the axes are assigned in this PivotTable view.
	 *
	 * @param b
	 */
	public void setColGrand( boolean b )
	{
		if( b != fColGrand )
		{
			if( b == false )
			{
				grbit1 = (byte) (grbit1 & 0xFD);
			}
			else
			{
				grbit1 = (byte) (grbit1 | 0x2);
			}
			this.initGrbit();
		}
	}

	/**
	 * specifies whether the PivotTable contains grand totals for columns. MUST be 1 if none of the axes are assigned in this PivotTable view.
	 */
	public boolean getFColGrand()
	{
		return fColGrand;
	}

	/**
	 * specifies whether the PivotTable has AutoFormat applied.
	 *
	 * @param b
	 */
	public void setFAutoFormat( boolean b )
	{
		if( b != fAutoFormat )
		{
			if( b == false )
			{
				grbit1 = (byte) (grbit1 & 0xFB);
			}
			else
			{
				grbit1 = (byte) (grbit1 | 0x4);
			}
			this.initGrbit();
		}
	}

	/**
	 * specifies whether the PivotTable has AutoFormat applied.
	 *
	 * @return
	 */
	public boolean getFAutoFormat()
	{
		return fAutoFormat;
	}

	public void setFWH( boolean b )
	{
		if( b != fWH )
		{
			if( b == false )
			{
				grbit1 = (byte) (grbit1 & 0xEF);
			}
			else
			{
				grbit1 = (byte) (grbit1 | 0x10);
			}
			this.initGrbit();
		}
	}

	public boolean getFWH()
	{
		return fWH;
	}

	/**
	 * specifies whether the PivotTable has font AutoFormat applied.
	 *
	 * @param b
	 */
	public void setFFont( boolean b )
	{
		if( b != fFont )
		{
			if( b == false )
			{
				grbit1 = (byte) (grbit1 & 0xDF);
			}
			else
			{
				grbit1 = (byte) (grbit1 | 0x20);
			}
			this.initGrbit();
		}
	}

	/**
	 * specifies whether the PivotTable has font AutoFormat applied.
	 *
	 * @return
	 */
	public boolean getFFont()
	{
		return fFont;
	}

	/**
	 * specifies whether the PivotTable has alignment AutoFormat applied.
	 *
	 * @param b
	 */
	public void setFAlign( boolean b )
	{
		if( b != fAlign )
		{
			if( b == false )
			{
				grbit1 = (byte) (grbit1 & 0xBF);
			}
			else
			{
				grbit1 = (byte) (grbit1 | 0x40);
			}
			this.initGrbit();
		}
	}

	/**
	 * specifies whether the PivotTable has alignment AutoFormat applied.
	 *
	 * @return
	 */
	public boolean getFAlign()
	{
		return fAlign;
	}

	/**
	 * specifies whether the PivotTable has border AutoFormat applied.
	 *
	 * @param b
	 */
	public void setFBorder( boolean b )
	{
		if( b != fBorder )
		{
			if( b == false )
			{
				grbit1 = (byte) (grbit1 & 0x7F);
			}
			else
			{
				grbit1 = (byte) (grbit1 | 0x80);
			}
			this.initGrbit();
		}
	}

	/**
	 * specifies whether the PivotTable has border AutoFormat applied.
	 *
	 * @return
	 */
	public boolean getFBorder()
	{
		return fBorder;
	}

	/**
	 * specifies whether the PivotTable has pattern AutoFormat applied.
	 *
	 * @param b
	 */
	public void setFPattern( boolean b )
	{
		if( b != fPattern )
		{
			if( b == false )
			{
				grbit2 = (byte) (grbit2 & 0xFE);
			}
			else
			{
				grbit2 = (byte) (grbit2 | 0x1);
			}
			this.initGrbit();
		}
	}

	/**
	 * specifies whether the PivotTable has pattern AutoFormat applied.
	 *
	 * @return
	 */
	public boolean getFPattern()
	{
		return fPattern;
	}

	/**
	 * @param b
	 */
	public void setFNumber( boolean b )
	{
		if( b != fNumber )
		{
			if( b == false )
			{
				grbit2 = (byte) (grbit2 & 0xFE);
			}
			else
			{
				grbit2 = (byte) (grbit2 | 0x1);
			}
			this.initGrbit();
		}
	}

	public boolean getFNumber()
	{
		return fNumber;
	}

	/**
	 * sets the location and size of this pivot table view
	 *
	 * @param range
	 */
	public void setLocation( String range )
	{
		int[] rc = ExcelTools.getRangeRowCol( range );
		setRwFirst( (short) (rc[0]) );
		setColFirst( (short) rc[1] );
		setRwLast( (short) (rc[2]) );
		setColLast( (short) rc[3] );
	}
    /*
        Set and get methods,  these all basically do the same thing, update the
        member variable, then update the body data with the new information.  Simple, but large XLSRecord
    */

	/**
	 * sets the first row of the pivot table
	 */
	public void setRwFirst( short s )
	{
		rwFirst = s;
		byte[] b = ByteTools.shortToLEBytes( rwFirst );
		System.arraycopy( b, 0, data, 0, 2 );
	}

	/**
	 * retrieves the first row of the pivot table
	 */
	public short getRwFirst()
	{
		return rwFirst;
	}

	/**
	 * sets the last row of the pivot table
	 */
	public void setRwLast( short s )
	{
		rwLast = s;
		byte[] b = ByteTools.shortToLEBytes( rwLast );
		System.arraycopy( b, 0, data, 2, 2 );
	}

	/**
	 * retieves the first row of the pivot table
	 */
	public short getRwLast()
	{
		return rwLast;
	}

	/**
	 * sets the first column of the pivot table
	 */
	public void setColFirst( short s )
	{
		colFirst = s;
		byte[] b = ByteTools.shortToLEBytes( colFirst );
		System.arraycopy( b, 0, data, 4, 2 );
	}

	/**
	 * retrieves the first column of the pivot table
	 */
	public short getColFirst()
	{
		return colFirst;
	}

	/**
	 * sets the last column of the pivot table
	 */
	public void setColLast( short s )
	{
		colLast = s;
		byte[] b = ByteTools.shortToLEBytes( colLast );
		System.arraycopy( b, 0, data, 6, 2 );
	}

	/**
	 * sets the last column of the pivot table
	 */
	public short getColLast()
	{
		return colLast;
	}

	public void setRwFirstHead( short s )
	{
		rwFirstHead = s;
		byte[] b = ByteTools.shortToLEBytes( rwFirstHead );
		System.arraycopy( b, 0, data, 8, 2 );
	}

	public short getRwFirstHead()
	{
		return rwFirstHead;
	}

	public void setRwFirstData( short s )
	{
		rwFirstData = s;
		byte[] b = ByteTools.shortToLEBytes( rwFirstData );
		System.arraycopy( b, 0, data, 10, 2 );
	}

	public short getRwFirstData()
	{
		return rwFirstData;
	}

	public void setColFirstData( short s )
	{
		colFirstData = s;
		byte[] b = ByteTools.shortToLEBytes( colFirstData );
		System.arraycopy( b, 0, data, 12, 2 );
	}

	public short getColFirstData()
	{
		return colFirstData;
	}

	/*
		Probably want to use the individual grbit options instead of this, but
		hey,  I'm being thorough.
	*/
	public void setGrbit( short s )
	{
		byte[] b = ByteTools.shortToLEBytes( s );
		grbit2 = b[0];
		grbit1 = b[1];
		System.arraycopy( b, 0, data, 36, 2 );
		this.initGrbit();
	}

	public void setItblAutoFmt( short s )
	{
		itblAutoFmt = s;
		byte[] b = ByteTools.shortToLEBytes( itblAutoFmt );
		System.arraycopy( b, 0, data, 38, 2 );
	}

	public short getItblAutoFmt()
	{
		return itblAutoFmt;
	}

	/**
	 * Sets the name of the Pivot Table.
	 */
	public void setTableName( String s )
	{
		PivotTableName = s;
		this.buildRgch();
		// also set associated qsitag pivot view name
		QsiSXTag qsi = (QsiSXTag) getSubRec( QSISXTAG, -1 );
		if( qsi != null )
		{
			qsi.setName( s );
		}
		// find SXADDL_SxView_SxDID record and set PivotTableView name - must match this view tablename
		for( int i = subRecs.size() - 1; i > 0; i-- )
		{
			BiffRec br = ((BiffRec) subRecs.get( i ));
			if( br.getOpcode() == SXADDL )
			{
				if( ((SxAddl) br).getRecordId() == SxAddl.SxcView.sxdId )
				{
					((SxAddl) br).setViewName( s );
					break;
				}
			}
		}

	}

	/**
	 * return the name of the Pivot Table.
	 */
	public String getTableName()
	{
		return PivotTableName;
	}

	/**
	 * Sets the name of the Data field
	 */
	public void setDataName( String s )
	{
		DataFieldName = s;
		this.buildRgch();
	}

	public String getDataName()
	{
		return DataFieldName;
	}

	/*
		Builds a new rgch field, changes the length of the entire record....
	*/
	private void buildRgch()
	{
		byte[] data = new byte[44];
		System.arraycopy( this.getData(), 0, data, 0, 44 );
		byte[] strbytes = new byte[0];
		byte[] databytes = new byte[0];
		try
		{
			if( PivotTableName != null )
			{
				strbytes = PivotTableName.getBytes( DEFAULTENCODING );
			}
			if( DataFieldName != null )
			{
				databytes = DataFieldName.getBytes( DEFAULTENCODING );
			}
		}
		catch( UnsupportedEncodingException e )
		{
			Logger.logInfo( "encoding pivot table name in SXVIEW: " + e );
		}

		//update the lengths:
		cchName = (short) strbytes.length;
		byte[] nm = ByteTools.shortToLEBytes( cchName );
		data[40] = nm[0];
		data[41] = nm[1];

		cchData = (short) databytes.length;
		nm = ByteTools.shortToLEBytes( cchData );
		data[42] = nm[0];
		data[43] = nm[1];

		// now append variable-length string data
		byte[] newrgch = new byte[cchName + cchData + 2];    // account for encoding bytes
		System.arraycopy( strbytes, 0, newrgch, 1, cchName );
		System.arraycopy( databytes, 0, newrgch, cchName + 2, cchData );

		data = ByteTools.append( newrgch, data );
		this.setData( data );
	}

	/**
	 * utility method to retrieve the nth pivot table subrecord (SxVd, SxPI, SxDI ...)
	 *
	 * @param opcode opcode to look up
	 * @param index  if > -1, the index or occurrence of the record to return
	 * @return
	 */
	private BiffRec getSubRec( int opcode, int index )
	{
		int j = 0;
		for( int i = 0; i < subRecs.size(); i++ )
		{
			BiffRec br = subRecs.get( i );
			if( br.getOpcode() == opcode )
			{
				if( (index == -1) || (j++ == index) )
				{
					return br;
				}
			}
		}
		return null;
	}

	/**
	 * utility method to retrieve the subrecord index of the desired pivot table record
	 *
	 * @param br pivot table record to look up
	 * @return
	 */
	private int getSubRecIndex( BiffRec br )
	{
		int i = 0;
		for(; i < subRecs.size(); i++ )
		{
			if( br == subRecs.get( i ) )
			{
				break;
			}
		}
		return i;
	}

	/**
	 * utility method to retrieve a pivot table subrecord (SxVd, SxPI, SxDI ...) index
	 *
	 * @param opcode opcode to look up
	 * @param index  if > -1, the index or occurrence of the record to return
	 * @return index in subrecords of desired record
	 */
	private int getSubRecIndex( int opcode, int index )
	{
		int j = 0;
		for( int i = 0; i < subRecs.size(); i++ )
		{
			BiffRec br = subRecs.get( i );
			if( br.getOpcode() == opcode )
			{
				if( (index == -1) || (j++ == index) )
				{
					return i;
				}
			}
		}
		return -1;
	}

	/**
	 * i
	 * finds the proper sheet rec insertion index for the desired opcode
	 * as well as the proper insertion index into the subrecord array
	 * <br>this method contains alot of intelligence regarding the order of pivot table records
	 *
	 * @param opcode one of SXVD, SXIVD, SXPI, SXDI, SXLI, or SXVDEX
	 * @param index  either 0 (for ROW records) or 1 (for COL records)
	 * @return int[]    {sheet record index, subrecord index} or -1's
	 */
	private int getPivotRecordInsertionIndexes( int opcode, int index, int pivotFieldIndex )
	{
		int i, j = 0;
		if( pivotFieldIndex < 0 )
		{
			for( i = subRecs.size() - 1; i >= 0; i-- )
			{
				BiffRec br = subRecs.get( i );
				int bropcode = br.getOpcode();
				if( bropcode == opcode )
				{ //
					if( j < index )
					{
						j++;
						continue;        // haven't found correct record jet
					}
					break;
				}
				else if( (opcode == SXLI) && ((bropcode == SXDI) || (bropcode == SXPI) || (bropcode == SXIVD) || (bropcode == SXVDEX)) )
				{
					break;
				}
				else if( (opcode == SXDI) && ((bropcode == SXPI) || (bropcode == SXIVD) || (bropcode == SXVDEX)) )
				{
					break;
				}
				else if( (opcode == SXPI) && ((bropcode == SXIVD) || (bropcode == SXVDEX)) )
				{
					break;
				}
				else if( (opcode == SXIVD) && (bropcode == SXVDEX) )
				{
					break;
				}
				else if( (opcode == SXVD) && (bropcode == SXEX) )
				{
					break;
				}
			}
			i++;    // counter reverse order
		}
		else
		{    // only lookup within the current pivot field
			for( i = pivotFieldIndex + 1; i < subRecs.size(); i++ )
			{
				BiffRec br = subRecs.get( i );
				int bropcode = br.getOpcode();
				if( bropcode == opcode )
				{ //
					if( j < index )
					{
						j++;
						continue;        // haven't found correct record jet
					}
					break;
				}
				else if( bropcode == SXVDEX )
				{
					break;
				}
			}
		}
		return i;
/*	    if (i >= 0) {
	    	return new int[] {subRecs.get(i+j).getRecordIndex(), i};
	    }
	    return new int[] {-1, i};*/
	}

	/**
	 * creates the basic, default records necessary to define a pivot table
	 *
	 * @param sheet string sheetname where pivot table is located
	 * @return arraylist of records
	 */
	public ArrayList addInitialRecords( Boundsheet sheet )
	{
    	/*
    	 * basic blank pivot table:
    	 * SXVIEW
    	 *  	after SXVIEW pivot field and item records will go: 
    	 *  		for each pivot field:  SxVd, SxVi* n items, SxVDEx
    	 *  	SXIVD if cDimRw > 0
    	 *  	SXIVD if cDimCol > 0
    	 *  	SXPI if cDimPg > 0
    	 *  	SXDI if cDimData > 0
    	 *  	SXLI if cRw > 0
    	 *  	SXLI if cCol > 0
    	 * SXEX
    	 * QXISTAG
    	 * SXADDL records
    	 * 
    	 */
		ArrayList initialrecs = new ArrayList();
		this.setSheet( sheet );
		this.setWorkBook( sheet.getWorkBook() );
		// SXEX
		SxEX sxex = (SxEX) SxEX.getPrototype();
		addInit( initialrecs, sxex, sheet );
		// QSISXTAG
		QsiSXTag qsi = (QsiSXTag) QsiSXTag.getPrototype();
		addInit( initialrecs, qsi, sheet );
		SxVIEWEX9 sxv = (SxVIEWEX9) SxVIEWEX9.getPrototype();
		addInit( initialrecs, sxv, sheet );
		// SXADDLs -required SxAddl records: stores additional PivotTableView info of a variety of types
//		byte[] b= ByteTools.cLongToLEBytes(sid);      	b= ByteTosols.append(new byte[] {0, 0}, b); // add 2 reserved bytes
		SxAddl sa = SxAddl.getDefaultAddlRecord( SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdId.sxd(), null );
		addInit( initialrecs, sa, sheet );
		//SXADDL_sxcView: record=sxdVer10Info data:[1, 65, 0, 0, 0, 0]
		sa = SxAddl.getDefaultAddlRecord( SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdVer10Info.sxd(), null );
		addInit( initialrecs, sa, sheet );
		sa = SxAddl.getDefaultAddlRecord( SxAddl.ADDL_CLASSES.sxcView,
		                                  SxcView.sxdTableStyleClient.sxd(),
		                                  new byte[]{ 0, 0, 0, 0, 0, 0, 51, 0, 0, 0 } );
		addInit( initialrecs, sa, sheet );
		sa = SxAddl.getDefaultAddlRecord( SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdVerUpdInv.sxd(), new byte[]{ 2, 0, 0, 0, 0, 0 } );
		addInit( initialrecs, sa, sheet );
		sa = SxAddl.getDefaultAddlRecord( SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdVerUpdInv.sxd(), new byte[]{ -1, 0, 0, 0, 0, 0 } );
		addInit( initialrecs, sa, sheet );
		sa = SxAddl.getDefaultAddlRecord( SxAddl.ADDL_CLASSES.sxcView, SxcView.sxdEnd.sxd(), null );
		addInit( initialrecs, sa, sheet );
		return initialrecs;
	}

	/**
	 * utility function to properly add a Pivot Table View subrec
	 *
	 * @param initialrecs
	 * @param rec
	 * @param addToInitRecords
	 * @param sheet
	 */
	private void addInit( ArrayList initialrecs, XLSRecord rec, Boundsheet sheet )
	{
		initialrecs.add( rec );
		rec.setSheet( sheet );
		rec.setWorkBook( this.getWorkBook() );
		this.addSubrecord( rec );
	}

	public String toString()
	{
		return "SXVIEW: name:" + this.getTableName() + " iCache:" + iCache + " cDim:" + cDim + " cDimRw:" + cDimRw + " cDimCol:" + cDimCol +
				" cDimPg:" + cDimPg + " cDimData:" + cDimData +
				" cRw:" + cRw + " cCol:" + cCol +
				" datafieldname:" + DataFieldName;
	}
}

