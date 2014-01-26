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
package org.openxls.formats.LEO;

import java.io.OutputStream;
import java.nio.ByteBuffer;
import java.util.List;

/**
 * the basic unit of data in a LEO file.  Can either be BIG or SMALL.
 */
public interface Block extends java.util.Iterator
{

	public final static int SMALL = 0;
	public final static int BIG = 1;

	/**
	 * @return
	 */
	public boolean isXBAT();

	/**
	 * @param b
	 */
	public void setIsExtraSector( boolean b );

	/**
	 * link to the vector of blocks for the storage
	 */
	public void setBlockVector( List v );

	/**
	 * get the size of the Block data in bytes
	 *
	 * @return block data size
	 */
	public int getBlockSize();

	/**
	 * get the index of this Block in the storage
	 * Vector
	 */
	public int getBlockIndex();

	/**
	 * link the next Block in the chain
	 */
	void setNextBlock( Block b );

	/**
	 * set the storage for this Block
	 */
	void setStorage( Storage s );

	/**
	 * set the storage for this Block
	 */
	Storage getStorage();

	/**
	 * returns whether this block has been
	 * added to the output stream
	 */
	boolean getStreamed();

	/**
	 * sets whether this block has been
	 * added to the output stream
	 */
	void setStreamed( boolean b );

	/**
	 * returns the int representing the block type
	 */
	int getBlockType();

	/**
	 * returns true if this is a Block Depot block
	 * that needs to be ignored when reading byte storages
	 */
	boolean getIsSpecialBlock();

	/**
	 * returns true if this is a Block Depot block
	 * that needs to be ignored when reading byte storages
	 */
	boolean getIsDepotBlock();

	/**
	 * set to true if this is a Block Depot block
	 */
	void setIsDepotBlock( boolean b );

	/**
	 * init the Block Data
	 */
	void init( ByteBuffer d, int origidx, int origp );

	/**
	 * set whether this Block has been read yet...
	 */
	void setInitialized( boolean b );

	/**
	 * returns whether this Block has been read yet...
	 */
	boolean getInitialized();

	/**
	 * set the data bytes  on this Block
	 */
	void setBytes( ByteBuffer b );

	/**
	 * get the data bytes  on this Block
	 */
	ByteBuffer getByteBuffer();

	/**
	 * return the byte Array for this BLOCK
	 */
	public byte[] getBytes( int start, int end );

	/**
	 * get the data bytes  on this Block
	 */
	byte[] getBytes();

	/**
	 * write the data bytes on this Block to out
	 */
	void writeBytes( OutputStream out );

	/**
	 * return the original BB position in the file
	 */
	int getOriginalPos();

	/**
	 * return the original BB position in the file
	 */
	int getOriginalIdx();

	/**
	 * set the original BB position in the file
	 */
	void setOriginalIdx( int x );
}