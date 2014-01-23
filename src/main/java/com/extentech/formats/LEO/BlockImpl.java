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
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
	private static final Logger log = LoggerFactory.getLogger( BlockImpl.class );
	private static final long serialVersionUID = 4833713921208834278L;
	/*allows the block to be populated with a byte array
	 rather than just a bytebuffer, easing debugging
    **/

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
		if( (nextblock != null) && (nextblock != this) )
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
	@Override
	public ByteBuffer getByteBuffer()
	{
		return data;
	}

	/**
	 * Write the entire bytes directly to out
	 *
	 * @see com.extentech.formats.LEO.Block#writeBytes(java.io.OutputStream)
	 */
	@Override
	public void writeBytes( OutputStream out )
	{
		try
		{
			out.write( getBytes() );
		}
		catch( Exception exp )
		{
			log.error( "BlockImpl.writeBytes failed.", exp );
		}
	}

	//byte[] delbytes = null;

	/**
	 * return the byte Array for this BLOCK
	 */
	@Override
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
		if( capcheck < (SIZE + originalpos) )
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
			log.error( "BlockImpl.getBytes() start: " + start + " size: " + SIZE + ": " + e, e );
		}
		return ret;
	}

	@Override
	public int getBlockSize()
	{
		if( this.getBlockType() == BIG )
		{
			return BIGBLOCK.SIZE;
		}
		if( this.getBlockType() == SMALL )
		{
			return SMALLBLOCK.SIZE;
		}
		return 0;
	}

	/**
	 * return the byte Array for this BLOCK
	 */
	@Override
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
		if( capcheck < (SIZE + originalpos) )
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
			log.error( "BlockImpl.getBytes(0," + SIZE + "): " + e, e );
		}
		return ret;
	}

	/**
	 * set the data bytes  on this Block
	 */
	@Override
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
	@Override
	public int getRecordIndexHint()
	{
		return recordIdx;
	}

	/**
	 * set index information about this
	 * objects likely position.
	 */
	@Override
	public void setRecordIndexHint( int i )
	{
		recordIdx = i;
	}

	/**
	 * link to the vector of blocks for the storage
	 */
	@Override
	public void setBlockVector( List v )
	{
		blockvec = v;
	}

	/**
	 * get the index of this Block in the storage
	 * Vector
	 */
	@Override
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
	@Override
	public void setOriginalIdx( int x )
	{
		originalidx = x;
	}

	@Override
	public void setNextBlock( Block b )
	{
		nextblock = b;
/*		if (LEOFile.DEBUG)
			Logger.logInfo(
				"INFO: BlockImpl setNextBlock(): " + b.toString());*/
	}

	@Override
	public boolean hasNext()
	{
		if( nextblock != null )
		{
			return true;
		}
		return false;
	}

	@Override
	public Object next()
	{
		return nextblock;
	}

	@Override
	public void remove()
	{
		this.mystore.removeBlock( this );
		nextblock = null;
	}

	/**
	 * set the storage for this Block
	 */
	@Override
	public void setStorage( Storage s )
	{
		mystore = s;
	}

	/**
	 * get the storage for this Block
	 */
	@Override
	public Storage getStorage()
	{
		return mystore;
	}

	/**
	 * return the original position of this BIGBLOCK
	 */
	@Override
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
	@Override
	public boolean getStreamed()
	{
		return streamed;
	}

	/**
	 * sets whether this block has been
	 * added to the output stream
	 */
	@Override
	public void setStreamed( boolean b )
	{
		streamed = b;
	}

	/**
	 * set whether this Block has been read yet...
	 */
	@Override
	public void setInitialized( boolean b )
	{
		initialized = b;
	}

	/**
	 * returns whether this Block has been read yet...
	 */
	@Override
	public boolean getInitialized()
	{
		return initialized;
	}

	/**
	 * returns true if this is a Block Depot block
	 * that needs to be ignored when reading byte storages
	 */
	@Override
	public boolean getIsSpecialBlock()
	{
		return isSpecialBlock;
	}

	/**
	 * returns true if this is a Block Depot block
	 * that needs to be ignored when reading byte storages
	 */
	@Override
	public boolean getIsDepotBlock()
	{
		return isBBDepotBlock;
	}

	/**
	 * set to true if this is a Block Depot block
	 */
	@Override
	public void setIsDepotBlock( boolean b )
	{
		isSpecialBlock = b;
		isBBDepotBlock = b;
	}

	/**
	 * init the BIGBLOCK Data
	 */
	@Override
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
	@Override
	public int getOriginalPos()
	{
		return originalpos;
	}

	/**
	 * @return
	 */
	@Override
	public boolean isXBAT()
	{
		return isXBAT;
	}

	/**
	 * true if ths block is represents an extra DIFAT sector
	 *
	 * @param b
	 */
	@Override
	public void setIsExtraSector( boolean b )
	{
		isXBAT = b;
	}

}