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

import java.util.Iterator;
import java.util.List;

/**
 * <b>Dbcell: Stream Offsets 0xD7</b><br>
 * <p/>
 * Offsets for value records.  There is one DBCELL for
 * each 32-row block of Row records and associated cell records.
 * <p/>
 * <br>
 * <p/>
 * <pre>
 * offset  name        size    contents
 * ---
 * 4       dbRtrw      4       Offset from the start of the DBCELL
 * to the start of the first Row in the block
 * 8       rgdb        var     Array of stream offsets, 2 bytes each.
 *
 * The internal format layout is basically:
 *
 *
 * When a record value changes in size, it fires a CellChangeEvent
 * which cascades through the other associated objects.
 *
 * The record size change has the following effects on DBCELL record fields:
 *
 * 1. Row records for the data block move relative to the
 * DBCELL for the block.  The dbRtrw field tracks the position
 * of the first Row record, so this needs to be updated in the
 * DBCELL only for the changed record block.
 *
 * 2. The *record*  offsets for the Row stored in the rgdb array
 * change starting with the changed value.  These only change
 * for the record block containing the changed value -- subsequent
 * blocks maintain their relative position to their row and row
 * records.
 * </p></pre>
 *
 * @see WorkBook
 * @see Boundsheet
 * @see Index
 * @see Dbcell
 * @see Row
 * @see Cell
 * @see XLSRecord
 */
public final class Dbcell extends com.extentech.formats.XLS.XLSRecord implements XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( Dbcell.class );
	private static final long serialVersionUID = -3169134298616400374L;
	private Bof mybof;

	void setBof( Bof b )
	{
		mybof = b;
	}

	private byte[] rkdata = new byte[0];
	private int numrecs = -1;
	private int dbRtrw;
	private int dbcnum;
	private short[] rgdb;
	private Index myidx;

	Row[] myrows = new Row[32];

	/**
	 * Init the dbcell with it's new values
	 * <p/>
	 * TODO: review these calls to look for performance gain possibilities.  This method can
	 * get called a lot on output.
	 *
	 * @param offsetToFirstRow - offset from the dbcell to the top row
	 * @param valrecOffsets    - array of offsets from row to row in valrecs
	 */
	public void initDbCell( int offsetToFirstRow, List valrecOffsets )
	{
		byte[] newData = new byte[4 + ((valrecOffsets.size() - 1) * 2)];
		byte[] firstOffset = ByteTools.cLongToLEBytes( offsetToFirstRow );
		System.arraycopy( firstOffset, 0, newData, 0, 4 );
		Iterator it = valrecOffsets.iterator();
		int pointer = 4;
		while( it.hasNext() )
		{
			Short s = (Short) it.next();
			// we don't want the last length, not needed, end of chain!
			if( it.hasNext() )
			{
				short ns = s;
				byte[] b = ByteTools.shortToLEBytes( ns );
				newData[pointer++] = b[0];
				newData[pointer++] = b[1];
			}
		}
		setData( newData );
	}

	/**
	 * set the Index record for this RowBlock
	 * <p/>
	 * there are many RowBlocks per Index record
	 * Index records need to update their DBCELL
	 * pointers when we add or move a DBCELL.
	 */
	@Override
	public void setIndex( Index idx )
	{
		myidx = idx;
	}

	int rwct = 0;

	/**
	 * this method adds a Row to the Dbcell, as well
	 * as updating the Dbcell position based on existing
	 * row data.  For use in Index updateDbcellPOinters.
	 */
	boolean addRow( Row r )
	{
		//   if(!this.isFull()){
		myrows[rwct++] = r;
		return true;
		//    }else return false;
	}

	/**
	 * gets the value and updates the pointer.  this is the position
	 * of the first row in the row block.
	 */
	public int getdbRtrw()
	{
		if( rwct > 0 )
		{
			Row rw1 = myrows[0];
			int i = rw1.getOffset();
			int y = getOffset();
			return y - i;
		}
		return -1;
	}

	/** iterates through XLSRecords Contained in Rows
	 and gets offset pointers

	 Deprecated?  THinkso

	 void updateIndexes(){
	 Iterator e = myrows.iterator();
	 int startingSize = this.getLength();
	 int[] idxes = new int[myrows.size()];
	 int i = 0, pos = 4;
	 byte[] newrgdb = new byte[(myrows.size()*2)+4];
	 int offst = -1;
	 while(e.hasNext()){
	 Row c = (Row) e.next();
	 idxes[i]=c.getFirstCellOffset();
	 if(i==0){  // read page 443 of Excel Format book for info on this...
	 offst = idxes[i];
	 idxes[i] += 6;
	 Row r1 = null;
	 if(myrows.size() > 1) r1 = (Row) this.myrows.get(1);
	 else r1 = (Row) this.myrows.get(0);
	 int dd = r1.offset + r1.getLength();
	 if(false)Logger.logInfo("Initializing of Dbcell: " + this.getdbRtrw() + " - " + dd);
	 idxes[i] = idxes[i] - dd;
	 if(myrows.size()>1)idxes[i]+=r1.getLength();
	 }else{
	 idxes[i] -= offst;
	 //  offst+=idxes[i];
	 }
	 if(false)Logger.logInfo(" dbcellpointer: " + String.valueOf( idxes[i] ));
	 byte[] barr = ByteTools.shortToLEBytes((short)idxes[i++]);
	 System.arraycopy(barr, 0, newrgdb, pos, 2);
	 pos+=2;
	 }
	 this.setData(newrgdb);
	 this.setDbRtrw(this.getdbRtrw());
	 this.init();
	 }
	 */

	/** returns whether this RowBlock can
	 hold any more Row records

	 boolean isFull(){
	 if(myrows.size() >= 32)return true;
	 return false;
	 }*/

	//Row[] getRows(){return myrows;}

	/**
	 * returns the byte array containing the DBCELL location
	 * as an offset from the Bof for the BOUNDSHEET.
	 * <p/>
	 * this is used by the Index recordto locate RowBlocks.
	 */
	byte[] getDBCELLPointerPos()
	{
		int bofpos = mybof.offset;
		int thispos = offset;
		int diff = thispos - bofpos;
		return ByteTools.cLongToLEBytes( diff );
	}

	/**
	 * get a new, empty DBCELL
	 */
	public static XLSRecord getPrototype()
	{
		Dbcell dbc = new Dbcell();
		byte[] dt = new byte[4]; // default val
		dbc.originalsize = 4;
		dbc.setData( dt );
		dbc.setOpcode( DBCELL );
		dbc.setLength( (short) 4 );
		dbc.init();
		return dbc;
	}

	/** handle a cell change event

	 public void fireCellChangeEvent(CellChangeEvent c){
	 // do its thing,
	 this.doCellSizeChangeAction(c);
	 // then pass it along...
	 this.getIDX().fireCellChangeEvent(c);
	 }*/

	/**
	 * set the DBCELL number for use by the Index
	 */
	void setDBCNum( int l )
	{
		dbcnum = l;
	}

	/**
	 * get the DBCELL number for use by the Index
	 */
	int getDBCNum()
	{
		return dbcnum;
	}

	/** set the rows for this dbcell

	 public void setRows(AbstractList r){
	 this.myrows = r;
	 for(int i = 0;i<myrows.size();i++){
	 Row rt = (Row) myrows.get(i);
	 rt.setDBCell(this);
	 }
	 }*/

	/**
	 * returns the number of rows
	 * contained in this DBCELL.  There should
	 * never be more than 32.
	 */
	public int getNumRows()
	{
		return rwct;
		// if(ret > 32)if(DEBUGLEVEL > -1)Logger.logInfo("DBCELL has too many rows: "+ String.valueOf(ret));
		//return ret;
	}

	/**
	 * set the associated dbcell index
	 */
	void setIDX( Index indx )
	{
		myidx = indx;
	}

	/**
	 * get the associated dbcell index
	 */
	public Index getIDX()
	{
		return myidx;
	}

	/**
	 * Initialize the Dbcell
	 */
	@Override
	public void init()
	{
		super.init();
		dbRtrw = ByteTools.readInt( getByteAt( 0 ), getByteAt( 1 ), getByteAt( 2 ), getByteAt( 3 ) );
		numrecs = (getLength() - 8) / 2;
		int pos = 4;
		rgdb = new short[numrecs];
		for( int i = 0; i < numrecs; i++ )
		{
			rgdb[i] = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) );
		}
	}

	/**
	 * set the dbRtrw pointer location.
	 */
	void setDbRtrw( int val )
	{
		dbRtrw = val;
		byte[] b = ByteTools.cLongToLEBytes( val );
		System.arraycopy( b, 0, getData(), 0, 4 );
	}

	/**
	 */
	@Override
	public void preStream()
	{
		//this.updateIndexes();
	}

	@Override
	public void close()
	{
		super.close();
		rgdb = null;
		myidx = null;
		myrows = null;
	}

}