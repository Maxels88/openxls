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
import com.extentech.toolkit.CompatibleVector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Serializable;

/**
 * <b>Index: Index Record 0x20B</b><br>
 * <p/>
 * Index records are written after the Bof record for each Boundsheet.
 * <p><pre>
 * <p/>
 * offset  name        size    contents
 * ---
 * 4       Reserved    4       Must be zero
 * 8       rwMic       4       First row that exists on the sheet
 * 12      rwMac       4       Last row that exists on the sheet, plus 1
 * 16      Reserved    4       Pointer to DIMENSIONS record for Boundsheet
 * 20      rgibRw      var     Array of file offsets to the Dbcell records
 * for each block of ROW records.  A block
 * contains up to 32 ROW records.
 * <p/>
 * <p/>
 * When a record value changes in size, it fires a CellChangeEvent
 * which cascades through the other associated objects.
 * <p/>
 * The record size change has the following effect on INDEX record fields:
 * <p/>
 * 1. ALL Dbcell records in the file starting with the one
 * containing the changed record move.  This requires updating
 * the dbcellpointers (rgibRw) in ALL INDEX records which
 * are located after the changed Dbcell within the file stream.
 * <p/>
 * </p></pre>
 *
 * @see WorkBook
 * @see Boundsheet
 * @see Dbcell
 * @see Row
 * @see Cell
 * @see XLSRecord
 */

public final class Index extends com.extentech.formats.XLS.XLSRecord implements XLSConstants
{
	int offsetStart = 0;
	private static final Logger log = LoggerFactory.getLogger( Index.class );
	private static final long serialVersionUID = -753407655976707961L;
	private static int defaultsize = 16;
	private int rwMic;
	private int rwMac;
	private int dbnum = 0;
	//    private dbCellPointer[] dbcellarray; not used
	private CompatibleVector dbcells = new CompatibleVector();
	private int indexnum;
	private Dimensions dims;

	/**
	 * create a new INDEX rec
	 */
	public static XLSRecord getPrototype()
	{
		Index idx = new Index();
		byte[] dt = new byte[defaultsize]; // default val
		idx.originalsize = defaultsize;
		idx.setData( dt );
		idx.setOpcode( INDEX );
		idx.setLength( (short) defaultsize );
		idx.init();
		return idx;
	}

	/**
	 * get the index number for
	 * addressing.
	 */
	public int getIndexNum()
	{
		return indexnum;
	}

	/**
	 * set the index number for
	 * addressing.
	 */
	public void setIndexNum( int n )
	{
		indexnum = n;
	}

	@Override
	public void init()
	{
		super.init();
		// 1st 4 are reseverd-0
		rwMic = ByteTools.readInt( this.getBytesAt( 4, 4 ) );
		rwMac = ByteTools.readInt( this.getBytesAt( 8, 4 ) );
		// next 4 are position of defColWidth record - skip
/* no need to read in dbcell offsets as we don't do anything with it
 * 		int pos= 16;
		int recsize= 4; // KSC added
		int numdbcells = (this.getLength()-pos)/recsize;
//		dbcellarray = new dbCellPointer[numdbcells];
		// rest of data is position of dbCell records: rgibRw (variable): An array of FilePointer. Each FilePointer specifies the file position of each referenced DBCell record
		for(int i = 0;i< numdbcells;i++){
		    if(DEBUGLEVEL > 6)Logger.logInfo("Index -> initializing dbcell pointer: " + i);
		    byte[] bite = this.getBytesAt(pos,4);
		    dbCellPointer pointer = new dbCellPointer(bite);
		    pos += 4;
//		    dbcellarray[i] = pointer;
		}
//		*/
	}

	/**
	 * Prestream for Index is going to create the correct size record, and populate the correct number of dbcells.
	 * The actual values will not yet be populated, but the record sizes will be correct in order to get offsets
	 * working correctly.
	 * <p/>
	 * Once offsets are correctly calculated in bytestreamer.stream, we can come back and
	 * populate without the getIndex call overhead.
	 */
	@Override
	public void preStream()
	{
		// rebuild the record with the correct length body data to fit the new dbcells
		this.getData();
		int arrsize = 16 + (dbcells.size() * 4);
		byte[] newBytes = new byte[arrsize];
		Dbcell dbc = null;
		// KSC: Changed from copying 12 bytes to copying 16 bytes to keep DIMENSIONS reference
		System.arraycopy( this.getData(), 0, newBytes, 0, 16 );
		this.setData( newBytes );
	}

	/**
	 * update the min/max cols and rows
	 * 8       rwMic       4       First row that exists on the sheet
	 * 12      rwMac       4       Last row that exists on the sheet, plus 1
	 */
	public void updateRowDimensions( int lowRow, int hiRow )
	{
		byte[] rw = ByteTools.cLongToLEBytes( lowRow );
		System.arraycopy( rw, 0, this.getData(), 4, 4 );
		rw = ByteTools.cLongToLEBytes( hiRow + 1 );
		System.arraycopy( rw, 0, this.getData(), 8, 4 );
	}

	/**
	 * fire the cell change event
	 * <p/>
	 * public void fireCellChangeEvent(CellChangeEvent c){
	 * // do its thing
	 * // this.doCellSizeChangeAction(c);
	 * // then pass it along...
	 * //  this.getSheet().fireCellChangeEvent(c);
	 * }
	 */

	void setDimensions( Dimensions d )
	{
		this.dims = d;
	}

	// Not used??
	void setDimensionsOffset( int offset )
	{
		byte[] recData = this.getData();
		byte[] newoff = ByteTools.cLongToLEBytes( offset );
		System.arraycopy( newoff, 0, recData, 12, 4 );
		this.setData( recData );
	}

	/**
	 * add an associated Dbcell object
	 * to this INDEX.
	 */
	void addDBCell( Dbcell dbc )
	{
		boolean bAdd = true;
/* KSC: TESTING testLostBorders bug   for (int i= 0; i < dbcells.size() && bAdd; i++) {
    	if (Arrays.equals(((Dbcell) dbcells.get(i)).data,  dbc.data))
			bAdd= false;
    }*/
		if( bAdd )
		{
//        if(!dbcells.contains(dbc)){
			dbc.setDBCNum( dbnum++ );
			dbcells.add( dbc );
//        }
		}
	}

	void resetDBCells()
	{
		dbnum = 0;
		dbcells = new CompatibleVector();
	}

	/** compute the location of Dbcell records using
	 the INDEX dbcellpointers and the firstBof record position
	 in the workbook.

	 NOT USED
	 *
	 int getDbcellPosition(int pointernum){
	 int firstBofloc = wkbook.getFirstBof().offset;
	 byte[] b = dbcellarray[pointernum].cdb;
	 int pointerloc = ByteTools.readInt(b[0],b[1],b[2],b[3]);
	 return pointerloc + firstBofloc;
	 }*/

	/**
	 * Update the Dbcell Pointers
	 * <p/>
	 * At this point the index should have all of it's dbcells in place, as well as
	 * thier offsets being populated, so all that should need to be done is to create
	 * the index correctly out of its values
	 */

	void updateDbcellPointers()
	{
		streamer = getSheet().streamer;
		// first, get the collection of Rows from sheet
		Boundsheet bs = this.getSheet();
		Row[] rowz = bs.getRows();
		if( rowz.length != 0 )
		{
			this.updateRowDimensions( rowz[0].getRowNumber(), rowz[(rowz.length) - 1].getRowNumber() );
		}
		// create the new Dbcells if any rows exist within the sheet

		// rebuild the record with the correct length body data to fit the new dbcells
		int arrsize = 16 + (dbcells.size() * 4);
		byte[] newBytes = new byte[arrsize];
		Dbcell dbc = null;
		System.arraycopy( this.getData(), 0, newBytes, 0, 16 );
		int offset = 16;
		for( int i = 0; i < dbcells.size(); i++ )
		{
			Dbcell db = (Dbcell) dbcells.elementAt( i );
			int dbOff = db.getOffset();
			byte[] b = ByteTools.cLongToLEBytes( dbOff );
			newBytes[offset++] = b[0];
			newBytes[offset++] = b[1];
			newBytes[offset++] = b[2];
			newBytes[offset++] = b[3];
		}
		this.setData( newBytes );
	}

	/**
	 * update the dimensions info based on Dimensions rec
	 */
	void updateDimensions()
	{
		byte[] rkdata = this.getData();
		byte[] newb = ByteTools.cLongToLEBytes( dims.offset );
		rkdata[12] = newb[0];
		rkdata[13] = newb[1];
		rkdata[14] = newb[2];
		rkdata[15] = newb[3];

        /* these should match rwMic/rwMac on dimensions rec
        dims.rwMic;
        dims.rwMac;
        */
	}

	/**
	 * return the associated Dbcell objects
	 * for this INDEX.
	 */
	Dbcell[] getDBCells()
	{
		Object[] obj = dbcells.toArray();
		Dbcell[] dbcs = new Dbcell[obj.length];
		System.arraycopy( obj, 0, dbcs, 0, obj.length );
		return dbcs;
	}

	/**
	 * Called from streamer, this updates individual dbcell offset values.
	 * <p/>
	 * Will only run correctly if called sequentially, ie dboffset [0], [1], [2]
	 *
	 * @param DbcellNumber - which dbcell to update
	 * @param DbOffset     - the pure offset from beginning of file
	 */
	void setDbcellPosition( int DbcellNumber, int DbOffset )
	{
		if( offsetStart == 0 )
		{
			offsetStart = this.getSheet().getMyBof().getOffset();
		}
		int insertOffset = DbOffset - offsetStart;
		log.trace( "Setting DBBiffRec Position, offsetStart:" + offsetStart + " & InsertOffset = " + insertOffset );
		offsetStart += insertOffset;
		int insertloc = 16 + (DbcellNumber * 4);
		byte[] off = ByteTools.cLongToLEBytes( insertOffset );
		System.arraycopy( off, 0, data, insertloc, 4 );

	}

	/**
	 * file offset to the Dbcell record
	 */
	class dbCellPointer implements Serializable
	{

		int cellloc = 0;
		int datasiz = 0;
		short s2, s3;
		byte[] cdb = new byte[4];
		/**
		 * serialVersionUID
		 */
		private static final long serialVersionUID = -5132922970171084839L;

		dbCellPointer( byte[] b )
		{
			cdb = b;
			cellloc = ByteTools.readShort( b[0], b[1] );
			datasiz = ByteTools.readShort( b[2], b[3] );
		}

		/**
		 * Updates location of Dbcell pointer and data size(?)
		 */
		void adjustPosition( int i )
		{
			cellloc += i;
			datasiz += i;
		}

		byte[] getBytes()
		{
			byte[] bite = new byte[4];
			System.arraycopy( ByteTools.shortToLEBytes( (short) cellloc ), 0, bite, 0, 2 );
			System.arraycopy( ByteTools.shortToLEBytes( (short) datasiz ), 0, bite, 2, 2 );
			return bite;
		}
	}
}

