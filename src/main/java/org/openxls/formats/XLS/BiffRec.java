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

import org.openxls.ExtenXLS.CellRange;
import org.openxls.formats.LEO.BlockByteReader;

import java.util.List;

public interface BiffRec
{

	// integrating Cell Methods
	public Object getInternalVal();

	public void setHyperlink( Hlink h );

	public Hlink getHyperlink();

	public CellRange getMergeRange();

	public void setMergeRange( CellRange r );

	public Row getRow();

	public String getFormatPattern();

	public void copyFormat( BiffRec source );

	public Font getFont();

	public Xf getXfRec();

	public boolean remove( boolean t );

	public Formula getFormulaRec();

	public void setRow( Row r );

	public abstract void setRowNumber( int i );

	public void setIxfe( int xf );

	public int getIntVal();

	public double getDblVal();

	public boolean getBooleanVal();

	public boolean isReadOnly();

	/**
	 * return the record index of this object
	 */
	public abstract int getRecordIndex();

	/**
	 * adds a CONTINUE record to the array of CONTINUE records for this record,
	 * containing all data beyond the 8224 byte record size limit.
	 */
	public abstract void addContinue( Continue c );

	/**
	 * remove all Continue records
	 */
	public abstract void removeContinues();

	/**
	 * returns the array of CONTINUE records for this record, containing all
	 * data beyond the 8224 byte record size limit.
	 */

	public abstract List getContinueVect();

	/**
	 * returns whether this record has a CONTINUE record containing data beyond
	 * the 8228 byte record size limit.
	 * <p/>
	 * XLSRecords can have 0 or more CONTINUE records.
	 */
	public abstract boolean hasContinues();

	/**
	 * set whether this record contains the value of the Cell.
	 */
	public abstract void setIsValueForCell( boolean b );

	/**
	 * associate this record with its Index record
	 */
	public abstract void setIndex( Index id );

	/**
	 * Associate this record with a worksheet. First checks to see if there is
	 * already a cell with this address.
	 */
	public abstract void setSheet( Sheet b );

	/**
	 * get the WorkSheet for this record.
	 */
	public abstract Boundsheet getSheet();

	public abstract void setWorkBook( WorkBook wk );

	public abstract WorkBook getWorkBook();

	/**
	 * set the column
	 */
	public abstract void setCol( short i );

	public abstract void setRowCol( int[] x );

	public abstract short getColNumber();

	public abstract int getRowNumber();

	/**
	 * get a string address for the cell based on row and col ie: "H22"
	 */
	public abstract String getCellAddress();
	public abstract String getCellAddressWithSheet();

	/**
	 * perform record initialization
	 */
	public abstract void init();

	/**
	 * get a default "empty" data value for this record
	 */
	public abstract Object getDefaultVal();

	/**
	 * get the data type name for this record
	 */
	public abstract String getDataType();

	/**
	 * Get the value of the record as a Float. Value must be parseable as an
	 * Float or it will throw a NumberFormatException.
	 */
	public abstract float getFloatVal();

	/**
	 * Get the value of the record as a String.
	 */
	public abstract String getStringVal();

	/**
	 * Get the value of the record as a String.
	 */
	public abstract String getStringVal( String encoding );

	public abstract void setStringVal( String v );

	public abstract void setBooleanVal( boolean b );

	public abstract void setIntVal( int v );

	public abstract void setDoubleVal( double v );

	/**
	 * do any pre-streaming processing such as expensive index updates or other
	 * deferrable processing.
	 */
	public abstract void preStream();

	/**
	 * do any post-streaming cleanup such as expensive index updates or other
	 * deferrable processing.
	 */
	public abstract void postStream();

	/**
	 * set the XF (format) record for this rec
	 */
	public abstract void setXFRecord();

	/**
	 * set the XF (format) record for this rec
	 */
	public abstract void setXFRecord( int i );

	/**
	 * get the ixfe
	 */
	public abstract int getIxfe();

	public abstract short getOpcode();

	public abstract void setOpcode( short opcode );

	/**
	 * Returns the length of this record, including the 4 header bytes
	 */
	public abstract int getLength();

	public abstract void setLength( int len );

	public abstract void setByteReader( BlockByteReader db );

	public abstract BlockByteReader getByteReader();

	public abstract void setEncryptedByteReader( BlockByteReader db );

	public abstract BlockByteReader getEncryptedByteReader();

	//** Thread Safing ExtenXLS **//
	public abstract void setData( byte[] b );

	/**
	 * gets the full record bytes for this record including header bytes
	 *
	 * @return byte[] of all rec bytes
	 */
	public byte[] getBytes();

	/**
	 * gets the record data merging any Continue record data.
	 */
	public abstract byte[] getData();

	/**
	 * Gets the byte from the specified position in the record byte array.
	 *
	 * @param off
	 * @return
	 */
	public abstract byte getByteAt( int off );

	/**
	 * Gets the byte from the specified position in the record byte array.
	 *
	 * @param off
	 * @return
	 */
	public byte[] getBytesAt( int off, int len );

	/**
	 * Set the relative position within the data underlying the block vector
	 * represented by the BlockByteReader.
	 * <p/>
	 * In other words, this is the relative position used by the BlockByteReader
	 * to offset the Consumer's read position within the collection of Data
	 * Blocks.
	 * <p/>
	 * This may be an offset relative to the data in a file, or within a Storage
	 * contained in a file.
	 * <p/>
	 * The Workbook Storage for example will contain a non-contiguous collection
	 * of Blocks containing data from any number of positions in a file.
	 * <p/>
	 * This collection forms a contiguous span of bytes comprising an XLS
	 * Workbook. The XLSRecords within this span of bytes will set their
	 * relative position within this 'virtual' array. Thus the XLSRecord
	 * positions are relative to the order of bytes contained in the Block
	 * collection. The BOF record then is at offset 0 within the data of the
	 * first Block, even though the underlying data of this first Block may be
	 * anywhere on disk.
	 *
	 * @param pos
	 */
	public abstract void setOffset( int pos );

	/**
	 * Get the relative position within the data underlying the block vector
	 * represented by the BlockByteReader.
	 *
	 * @return relative position
	 */
	public abstract int getOffset();

	/**
	 * @return
	 */
	public abstract ByteStreamer getStreamer();

	/**
	 * @param streamer
	 */
	public abstract void setStreamer( ByteStreamer streamer );

	/**
	 * @return
	 */
	public abstract boolean isValueForCell();

	/**
	 * Dumps this record as a human-readable string.
	 * This method's output is more verbose than that of {@link #toString()}.
	 * It generally includes a hex dump of the record's contents.
	 */
	public String toHexDump();

	/**
	 * collect cell change listeners and fire cell change event upon, well ...
	 * @param t
	 * 20080204 KSC

	public void addCellChangeListener(CellChangeListener t);
	public void removeCellChangeListener(CellChangeListener t);
	// public void fireCellChangeEvent();
	public void setCachedValue(Object newValue);
	public Object getCachedValue();
	 */
}