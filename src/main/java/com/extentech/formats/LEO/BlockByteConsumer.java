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
package com.extentech.formats.LEO;

/**
 * The BlockByteConsumer interface describes an Object
 * which reads its data from a scattered collection of Blocks.
 * <p/>
 * By tracking the blocks containing the record's data,
 */
public interface BlockByteConsumer
{

	/**
	 * Set the relative position within the data
	 * underlying the block vector represented by
	 * the BlockByteReader.
	 * <p/>
	 * In other words, this is the relative position
	 * used by the BlockByteReader to offset the Consumer's
	 * read position within the collection of Data Blocks.
	 * <p/>
	 * This may be an offset relative to the data in a file,
	 * or within a Storage contained in a file.
	 * <p/>
	 * The Workbook Storage for example will contain a non-contiguous
	 * collection of Blocks containing data from any number of
	 * positions in a file.
	 * <p/>
	 * This collection forms a contiguous span of bytes comprising
	 * an XLS Workbook.  The XLSRecords within this span of bytes will
	 * set their relative position within this 'virtual' array.  Thus
	 * the XLSRecord positions are relative to the order of bytes contained
	 * in the Block collection.  The BOF record then is at offset 0 within the
	 * data of the first Block, even though the underlying data of this
	 * first Block may be anywhere on disk.
	 *
	 * @param pos
	 */
	void setOffset( int pos );

	/**
	 * Get the relative position within the data
	 * underlying the block vector represented by
	 * the BlockByteReader.
	 *
	 * @return relative position
	 */
	int getOffset();

	/** Get the blocks containing this Consumer's data
	 *
	 * @return
	 */
// KSC: NOT USED	Block[] getBlocks();

	/** Set the blocks containing this Consumer's data
	 *
	 * @param myblocks
	 */
// KSC: NOT USED void setBlocks(Block[] myblocks);

	/**
	 * Sets the index of the first block
	 *
	 * @return
	 */
	void setFirstBlock( int i );

	/**
	 * Sets the index of the last block
	 *
	 * @return
	 */
	void setLastBlock( int i );

	/**
	 * Returns the index of the first block
	 *
	 * @return
	 */
	int getFirstBlock();

	/**
	 * Returns the index of the last block
	 *
	 * @return
	 */
	int getLastBlock();

	/**
	 * Returns the length of the record.
	 *
	 * @return
	 */
	int getLength();

	/**
	 * Set the BlockByteReader for this Consumer
	 *
	 * @param db
	 */
	void setByteReader( BlockByteReader db );

	/**
	 * Get the BlockByteReader for this Consumer
	 *
	 * @param db
	 */
	BlockByteReader getByteReader();

}
