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

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.LEO.BlockByteReader;
import com.extentech.formats.LEO.Storage;
import com.extentech.formats.LEO.StorageNotFoundException;
import com.extentech.formats.LEO.StorageTable;
import com.extentech.toolkit.ByteTools;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.HashMap;

/**
 * represents the required storage _SX_CUR_DB for Pivot Tables
 *
 */

/**
 *
 */
public class PivotCache implements XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( PivotCache.class );
	WorkBook book;
	HashMap<Integer, Storage> caches = new HashMap();
	HashMap<Integer, ArrayList<BiffRec>> pivotCacheRecs = new HashMap();
//	HashMap<String, ArrayList> caches= new HashMap();

	public void init( StorageTable directories, WorkBookHandle wbh ) throws StorageNotFoundException
	{
		Storage child = directories.getChild( "_SX_DB_CUR" );
		if( wbh != null )
		{
			caches = new HashMap();
			book = wbh.getWorkBook();
		}

		while( child != null )
		{
			if( wbh != null )
			{
				caches.put( Integer.valueOf( child.getName() ), child );
			}
			ArrayList<BiffRec> curRecs = new ArrayList();
			BlockByteReader bytes = child.getBlockReader();
			int len = bytes.getLength();
			for( int i = 0; i <= (len - 4); )
			{
				byte[] headerbytes = bytes.getHeaderBytes( i );
				short opcode = ByteTools.readShort( headerbytes[0], headerbytes[1] );
				int reclen = ByteTools.readShort( headerbytes[2], headerbytes[3] );
				BiffRec rec = XLSRecordFactory.getBiffRecord( opcode );

				// init the mighty rec
				rec.setWorkBook( book );
				rec.setByteReader( bytes );
				rec.setLength( (short) reclen );
				rec.setOffset( i );
				rec.init();
/*				
// KSC: TESTING				
		try {
System.out.println(rec.getClass().getName().substring(rec.getClass().getName().lastIndexOf(".")+1) + ": " + Arrays.toString(((PivotCacheRecord)rec).getRecord()));				
		} catch (ClassCastException e) {
System.out.println(rec.getClass().getName().substring(rec.getClass().getName().lastIndexOf(".")+1) + ": " + Arrays.toString(ByteTools.shortToLEBytes(rec.getOpcode())) + Arrays.toString(ByteTools.shortToLEBytes((short)rec.getData().length)) + Arrays.toString(rec.getData()));				
		}
*/
				if( wbh != null )
				{
					curRecs.add( rec );
				}

				i += reclen + 4;
			}
			if( wbh != null )
			{
				pivotCacheRecs.put( Integer.valueOf( child.getName() ), curRecs );
			}
			child = directories.getNext( child.getName() );
		}
// KSC: TESTING		
//		Logger.logInfo("PivotCache.end init");
	}

	/**
	 * Creates the pivot table cache == defines pivot table data
	 * <br>A pivot table cache requires 2 directory storages
	 * <br>_SX_DB_CUR = parent pivot cache
	 * <br>0001, 0002 ... = child streams that define the pivot cache records
	 *
	 * @param directories
	 * @param wb
	 * @param ref         Cell Range which identifies pivot table data range
	 */
	public void createPivotCache( StorageTable directories, WorkBookHandle wbh, String ref, int sId ) throws InvalidRecordException
	{
		try
		{
			// KSC: TESTING
				log.trace( String.format( "creatpivotCache: ref: %s sid %d", ref, sId ) );
			/**
			 * the Pivot Cache Storage specifies zero or more streams, each of which specify a PivotCache
			 * The name of each stream (1) MUST be unique within the storage, and the name MUST be a four digit hexadecimal number stored as text.
			 * The number of FDB rules that occur MUST be equal to the value of cfdbTot in the SXDBrecord (section 2.4.275).
			 */
			/*
			 */
			// KSC: unsure if it's absolutely necessary to also have CompObj storage 
			try
			{
				directories.getDirectoryByName( "\1CompObj" );
			}
			catch( StorageNotFoundException e )
			{
				Storage compObj = directories.createStorage( "\1CompObj", 2, directories.getDirectoryStreamID(
						"\005DocumentSummaryInformation" ) + 1 );
				compObj.setBytesWithOverage( new byte[]{
						1,
						0,
						-2,
						-1,
						3,
						10,
						0,
						0,
						-1,
						-1,
						-1,
						-1,
						32,
						8,
						2,
						0,
						0,
						0,
						0,
						0,
						-64,
						0,
						0,
						0,
						0,
						0,
						0,
						70,
						38,
						0,
						0,
						0,
						77,
						105,
						99,
						114,
						111,
						115,
						111,
						102,
						116,
						32,
						79,
						102,
						102,
						105,
						99,
						101,
						32,
						69,
						120,
						99,
						101,
						108,
						32,
						50,
						48,
						48,
						51,
						32,
						87,
						111,
						114,
						107,
						115,
						104,
						101,
						101,
						116,
						0,
						6,
						0,
						0,
						0,
						66,
						105,
						102,
						102,
						56,
						0,
						14,
						0,
						0,
						0,
						69,
						120,
						99,
						101,
						108,
						46,
						83,
						104,
						101,
						101,
						116,
						46,
						56,
						0,
						-12,
						57,
						-78,
						113,
						0,
						0,
						0,
						0,
						0,
						0,
						0,
						0,
						0,
						0,
						0,
						0
				} );
				int compObjid = directories.getDirectoryStreamID( "\1CompObj" );
				Storage wb = directories.getDirectoryByName( "Workbook" );
				wb.setPrevStorageID( compObjid );
			}

			 /* create _SX_DB_CUR + child (actual pivot cache) directory and put in proper order in directory array */
			Storage sx_db_cur = directories.createStorage( "_SX_DB_CUR",
			                                               1,
			                                               directories.getDirectoryStreamID( "\005SummaryInformation" ) );    // Pivot Cache Storage:  (id= 1) insert just before SummaryInfo -- ALWAYS ??????
			int sxdbcurid = directories.getDirectoryStreamID( "_SX_DB_CUR" );    //
			Storage pcache1 = directories.createStorage( "0001", 2, sxdbcurid + 1 );            // TODO: handle multiple Caches ... 0002 ...
			directories.getDirectoryByName( "Root Entry" ).setChildStorageID( sxdbcurid );
			sx_db_cur.setPrevStorageID( directories.getDirectoryStreamID( "Workbook" ) );
			sx_db_cur.setChildStorageID( directories.getDirectoryStreamID( "0001" ) );            // child= 0001 Cache Stream (id= 2)
			sx_db_cur.setNextStorageID( directories.getDirectoryStreamID( "\005SummaryInformation" ) );
			Storage si = directories.getDirectoryByName( "\005SummaryInformation" );
			si.setPrevStorageID( -1 );    // Necessary?????????
			si.setNextStorageID( directories.getDirectoryStreamID( "\005DocumentSummaryInformation" ) );
			directories.getDirectoryByName( "Root Entry" ).setChildStorageID( sxdbcurid );    // ??? ALWAYS ????

			// create pivot cache records which are source of actual pivot cache data 
			byte[] newbytes = createPivotCacheRecords( ref, wbh, sId );
			pcache1.setBytesWithOverage( newbytes );
			this.init( directories, wbh );
		}
		catch( StorageNotFoundException e )
		{ // shouldn't!
		}
	}

	/**
	 * adds a specific instance of a cache field
	 * <br>A cache item is contained in a cache field. A cache field can have zero cache items if the cache field is not in use in the PivotTable view.
	 * TODO: handle unique cache items ...
	 * <p/>
	 * // TODO: need cache field index ***
	 *
	 * @param cacheItem
	 */
	public void addCacheItem( int cacheId, int cacheItem )
	{
		int insertIndex = 0;
		for( BiffRec br : pivotCacheRecs.get( cacheId + 1 ) )
		{
			if( br.getOpcode() == SXFDB )
			{
				((SxFDB) br).setNCacheItems( ((SxFDB) br).getNCacheItems() + 1 );
			}
			else if( br.getOpcode() == EOF )
			{
				insertIndex = pivotCacheRecs.get( cacheId + 1 ).indexOf( br );
				/**/
			}
		}
		// add required SXDBB for non-summary-cache items
		if( cacheItem > -1 )
		{
			/* SXDBB records only exist when put cache fields on a pivot table axis == cache item*/
			SxDBB sxdbb = (SxDBB) SxDBB.getPrototype();
			sxdbb.setCacheItemIndexes( new byte[]{
					Integer.valueOf( cacheItem ).byteValue()
			} );    //ByteTools.shortToLEBytes((short)cacheItem));
			pivotCacheRecs.get( cacheId + 1 ).add( insertIndex, sxdbb );
		}
		updateCacheRecords( cacheId );
	}

	/**
	 * take pivotCacheRecs and update the actual cache bytes
	 */
	private void updateCacheRecords( int cacheId )
	{
		byte[] newbytes = new byte[0];
		for( BiffRec br : pivotCacheRecs.get( cacheId + 1 ) )
		{
			try
			{
				newbytes = ByteTools.append( ((PivotCacheRecord) br).getRecord(), newbytes );
//System.out.println(br.getClass().getName().substring(br.getClass().getName().lastIndexOf(".")+1) + ": " + Arrays.toString(((PivotCacheRecord)br).getRecord()));				
			}
			catch( ClassCastException e )
			{
				newbytes = ByteTools.append( ByteTools.shortToLEBytes( br.getOpcode() ), newbytes );
				newbytes = ByteTools.append( ByteTools.shortToLEBytes( (short) br.getData().length ), newbytes );
				newbytes = ByteTools.append( br.getData(), newbytes );
//System.out.println(br.getClass().getName().substring(br.getClass().getName().lastIndexOf(".")+1) + ": " + Arrays.toString(ByteTools.shortToLEBytes(br.getOpcode())) + Arrays.toString(ByteTools.shortToLEBytes((short)br.getData().length)) + Arrays.toString(br.getData()));				
			}
		}
		Storage pcache1 = caches.get( cacheId + 1 );
		pcache1.setBytesWithOverage( newbytes );
// KSC: TESTING		
/*		try {
		this.init(book.factory.myLEO.getDirectoryArray(), null);
		} catch (StorageNotFoundException e) {			
		}*/
	}

	/**
	 * parse range and create required cache records, returning bytes defining said records.
	 * <br>For normal ranges, the PivotCache has one cache field for each column of the range,
	 * using the values in the first row of the range for cache field names,
	 * and all other rows are used as source data values, specified by cache records.
	 *
	 * @param ref Range Reference String, including sheet
	 * @param wbh workbookhandle
	 * @param sId Stream or cachid Id -- links back to SxStream set of records
	 */
	byte[] createPivotCacheRecords( String ref, WorkBookHandle wbh, int sId )
	{
		byte[] newbytes = new byte[0];
		try
		{
			CellRange cr = new CellRange( ref, wbh, false, true );
			CellHandle[] ch = cr.getCells();    // cells are in row-order
			int[] rows = cr.getRowInts();    // first row= header, ensuing rows are cacherecords
			int[] cols = cr.getColInts();
			int[] types = new int[cols.length];
			byte[][] cachefieldindexes = new byte[cols.length][rows.length - 1];
			SxDB sxdb = (SxDB) SxDB.getPrototype();
			sxdb.setNCacheRecords( rows.length - 1 );
			sxdb.setNCacheFields( cols.length );
			sxdb.setStreamID( sId );
			newbytes = ByteTools.append( sxdb.getRecord(), newbytes );
//System.out.println("SXDB: " + Arrays.toString(sxdb.getRecord()));				
			SXDBEx sxdbex = (SXDBEx) SXDBEx.getPrototype();    //TODO: nFormulas
			newbytes = ByteTools.append( sxdbex.getRecord(), newbytes );
//System.out.println("SXDBEX: " + Arrays.toString(sxdbex.getRecord()));				
			// TODO: cells after row header cell ***should be*** the same type -- true in ALL cases??????
			if( ch.length > cols.length )
			{ // have multiple rows
				for( int i = 0; i < cols.length; i++ )
				{
					CellHandle c = ch[i + (cols.length)];
					int type = -1;
					if( c.isDate() )
					{
						type = 6;
					}
					else
					{
						type = c.getCellType();
					}
					types[i] = type;
				}
			}
//TODO: ranges/grouping and formulas !!!!
//TODO: boolean vals?		
			for( int z = 0; z < rows.length; z++ )
			{
				for( int i = 0; i < cols.length; i++ )
				{
					if( z == 0 )
					{ // # SxFDB records==# COLUMNS==# Cache Fields
						SxFDB sxfdb = (SxFDB) SxFDB.getPrototype();
						sxfdb.setCacheItemsType( types[i] );
						sxfdb.setCacheField( ch[i].getStringVal() );    // row header values
						sxfdb.setNCacheItems( 0 );        // only set ACTUAL cache items when put cache field(s)on the pivot table (on row, page, column or data axis)
						newbytes = ByteTools.append( sxfdb.getRecord(), newbytes );
//System.out.println("SXFDB: " + Arrays.toString(sxfdb.getRecord()));				
						SXFDBType sxfdbtype = (SXFDBType) SXFDBType.getPrototype();
						newbytes = ByteTools.append( sxfdbtype.getRecord(), newbytes );
//System.out.println("SXDFBTYPE: " + Arrays.toString(sxfdbtype.getRecord()));				
						continue;
					}
					cachefieldindexes[i][z - 1] = (byte) i;
					// data cells== CACHE ITEMS 
					CellHandle c = ch[((z * (cols.length)) + i)];
// TODO: handle SxNil, SxErr, SxDtr
// TODO: handle SxFmla, SXName, SxPair, SxFormula					
					switch( types[i] )
					{
						case XLSConstants.TYPE_STRING:
							SXString sxstring = (SXString) SXString.getPrototype();
							sxstring.setCacheItem( c.getStringVal() );
							newbytes = ByteTools.append( sxstring.getRecord(), newbytes );
//System.out.println("SXSTRING: " + Arrays.toString(sxstring.getRecord()));				
							break;
						case XLSConstants.TYPE_FP:
						case XLSConstants.TYPE_INT:
						case XLSConstants.TYPE_DOUBLE:
							SXNum sxnum = (SXNum) SXNum.getPrototype();
							sxnum.setNum( c.getDoubleVal() );
							newbytes = ByteTools.append( sxnum.getRecord(), newbytes );
//System.out.println("SXNUM: " + Arrays.toString(sxnum.getRecord()));				
							break;
						case XLSConstants.TYPE_BOOLEAN:
							SXBool sxbool = (SXBool) SXBool.getPrototype();
							sxbool.setBool( c.getBooleanVal() );
							newbytes = ByteTools.append( sxbool.getRecord(), newbytes );
//System.out.println("SXBOOL: " + Arrays.toString(sxbool.getRecord()));				
							//TYPE_FORMULA = 3,		SxFmla *(SxName *SXPair)
						case 6:
							// SXDtr
					}
				}
			}
		}
		catch( Exception e )
		{
			throw new InvalidRecordException( "PivotCache.createPivotCache: invalid source range: " + ref );
		}
		// EOF -- header:
		byte[] b = new byte[4];
		System.arraycopy( ByteTools.shortToLEBytes( XLSConstants.EOF ), 0, b, 0, 2 );
		newbytes = ByteTools.append( b, newbytes );
//System.out.println("EOF: " + Arrays.toString((b)));				
		return newbytes;
	}
}
