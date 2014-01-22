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

import com.extentech.toolkit.CompatibleVectorHints;
import com.extentech.toolkit.Logger;

import java.io.OutputStream;
import java.io.Serializable;
import java.nio.ByteBuffer;
import java.util.List;

/**
 * LEO File Block Information Record
 * <p/>
 * These blocks of data contain information related to the
 * LEO file format data blocks.
 */
public abstract class BlockImpl implements com.extentech.formats.LEO.Block, CompatibleVectorHints, Serializable
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4833713921208834278L;
	/*allows the block to be populated with a byte array
	 rather than just a bytebuffer, easing debugging
    **/ boolean DEBUG = false;

	/**
	 * methods from CompatibleVectorHints
	 */
	transient int recordIdx = -1, lastidx = -1;
	transient ByteBuffer data = null; // new byte[SIZE];
	private List blockvec = null;

	// implement iterator for use in
	// chaining blocks
	public Block nextblock = null;
	private boolean isXBAT = false;

	int originalidx, originalpos;
	boolean isBBDepotBlock = false;
	boolean isSBDepotBlock = false;
	boolean isSpecialBlock = false;
	private boolean initialized = false;
	private boolean streamed = false;
	public Storage mystore = null;

	public void close()
	{
		if( blockvec != null )
		{
			blockvec.clear();
			blockvec = null;
		}
		if( nextblock != null && nextblock != this )
		{
			nextblock = null;
		}
		mystore = null;
		if( data != null )
		{
			data.clear();
			data = null;
		}
	}

	/**
	 * return the ByteBuffer for this BLOCK
	 */
	public ByteBuffer getByteBuffer()
	{
		return data;
	}

	/**
	 * Write the entire bytes directly to out
	 *
	 * @see com.extentech.formats.LEO.Block#writeBytes(java.io.OutputStream)
	 */
	public void writeBytes( OutputStream out )
	{
		try
		{
			out.write( getBytes() );
		}
		catch( Exception exp )
		{
			Logger.logErr( "BlockImpl.writeBytes failed.", exp );
		}
	}

	//byte[] delbytes = null;

	/**
	 * return the byte Array for this BLOCK
	 */
	public byte[] getBytes( int start, int end )
	{
		//if (delbytes == null) delbytes = getBytes();
		int SIZE = end - start;
		if( (end) > this.getBlockSize() )
		{
			throw new RuntimeException( "WARNING: BlockImpl.getBytes(): read position > block size:" + SIZE + start );
		}

		byte[] ret = new byte[SIZE];
		int capcheck = data.capacity();
		// TODO: track why this is occurring
		if( capcheck <= SIZE )
		{
			originalpos = 0;
		}
		// CAN HAPPEN ON OUT-OF-SPEC FILES whom have last block size < 512
		if( capcheck < SIZE + originalpos )
		{
			SIZE = capcheck - originalpos;
		}
		try
		{
			start += originalpos;
			data.position( start );
			data.get( ret, 0, SIZE );
		}
		catch( Exception e )
		{
			Logger.logWarn( "BlockImpl.getBytes() start: " + start + " size: " + SIZE + ": " + e );
		}
		return ret;
	}

	public int getBlockSize()
	{
		if( this.getBlockType() == BIG )
		{
			return BIGBLOCK.SIZE;
		}
		else if( this.getBlockType() == SMALL )
		{
			return SMALLBLOCK.SIZE;
		}
		else
		{
			return 0;
		}
	}

	/**
	 * return the byte Array for this BLOCK
	 */
	public byte[] getBytes()
	{
		int SIZE = 0;
		if( this.getBlockType() == BIG )
		{
			SIZE = BIGBLOCK.SIZE;
		}
		else if( this.getBlockType() == SMALL )
		{
			SIZE = SMALLBLOCK.SIZE;
		}
		byte[] ret = new byte[SIZE];

		int capcheck = data.capacity();
		if( capcheck <= originalpos )    // why is this hitting????????
		{
			originalpos = 0;
		}
		if( capcheck < SIZE + originalpos )
		{
			SIZE = capcheck - originalpos;    // CAN HAPPEN ON OUT-OF-SPEC FILES whom have last block size < 512
		}
		try
		{
			data.position( originalpos );
			data.get( ret, 0, SIZE );
		}
		catch( Exception e )
		{
			Logger.logWarn( "BlockImpl.getBytes(0," + SIZE + "): " + e );
		}
		return ret;
	}

	/**
	 * set the data bytes  on this Block
	 */
	public void setBytes( ByteBuffer b )
	{
		data = b;
		//if(DEBUG) {
		// Logger.logInfo("Debugging turned on in BlockImpl.setBytes");
		//delbytes = getBytes();
		//}
	}

	/**
	 * provide a hint to the CompatibleVector
	 * about this objects likely position.
	 */
	public int getRecordIndexHint()
	{
		return recordIdx;
	}

	/**
	 * set index information about this
	 * objects likely position.
	 */
	public void setRecordIndexHint( int i )
	{
		recordIdx = i;
	}

	/**
	 * link to the vector of blocks for the storage
	 */
	public void setBlockVector( List v )
	{
		blockvec = v;
	}

	/**
	 * get the index of this Block in the storage
	 * Vector
	 */
	public int getBlockIndex()
	{
		if( blockvec == null )
		{
			return this.recordIdx;
		}
		return blockvec.indexOf( this );
	}

	/**
	 * set the original BB position in the file
	 */
	public void setOriginalIdx( int x )
	{
		originalidx = x;
	}

	public void setNextBlock( Block b )
	{
		nextblock = b;
/*		if (LEOFile.DEBUG)
			Logger.logInfo(
				"INFO: BlockImpl setNextBlock(): " + b.toString());*/
	}

	public boolean hasNext()
	{
		if( nextblock != null )
		{
			return true;
		}
		return false;
	}

	public Object next()
	{
		return nextblock;
	}

	public void remove()
	{
		this.mystore.removeBlock( this );
		nextblock = null;
	}

	/**
	 * set the storage for this Block
	 */
	public void setStorage( Storage s )
	{
		mystore = s;
	}

	/**
	 * get the storage for this Block
	 */
	public Storage getStorage()
	{
		return mystore;
	}

	/**
	 * return the original position of this BIGBLOCK
	 */
	public int getOriginalIdx()
	{
		return originalidx;
	}

	/** returns the original BB pos

	 public int getOriginalPos(){return  this.originalpos;}
	 */
	/**
	 * returns whether this block has been
	 * added to the output stream
	 */
	public boolean getStreamed()
	{
		return streamed;
	}

	/**
	 * sets whether this block has been
	 * added to the output stream
	 */
	public void setStreamed( boolean b )
	{
		streamed = b;
	}

	/**
	 * set whether this Block has been read yet...
	 */
	public void setInitialized( boolean b )
	{
		initialized = b;
	}

	/**
	 * returns whether this Block has been read yet...
	 */
	public boolean getInitialized()
	{
		return initialized;
	}

	/**
	 * returns true if this is a Block Depot block
	 * that needs to be ignored when reading byte storages
	 */
	public boolean getIsSpecialBlock()
	{
		return isSpecialBlock;
	}

	/**
	 * returns true if this is a Block Depot block
	 * that needs to be ignored when reading byte storages
	 */
	public boolean getIsDepotBlock()
	{
		return isBBDepotBlock;
	}

	/**
	 * set to true if this is a Block Depot block
	 */
	public void setIsDepotBlock( boolean b )
	{
		isSpecialBlock = b;
		isBBDepotBlock = b;
	}

	/**
	 * init the BIGBLOCK Data
	 */
	public void init( ByteBuffer d, int origidx, int origp )
	{
		originalidx = origidx;
		originalpos = origp;
		this.setBytes( d );
	}

	/**
	 * return the original position of this
	 * BIGBLOCK record in the array of BIGBLOCKS
	 * that make up the file.
	 */
	public int getOriginalPos()
	{
		return originalpos;
	}

	/**
	 * @return
	 */
	public boolean isXBAT()
	{
		return isXBAT;
	}

	/**
	 * true if ths block is represents an extra DIFAT sector
	 *
	 * @param b
	 */
	public void setIsExtraSector( boolean b )
	{
		isXBAT = b;
	}

}