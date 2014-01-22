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

import com.extentech.formats.XLS.WorkBookException;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.io.ByteArrayOutputStream;
import java.io.Serializable;
import java.nio.Buffer;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.nio.CharBuffer;
import java.nio.DoubleBuffer;
import java.nio.FloatBuffer;
import java.nio.IntBuffer;
import java.nio.LongBuffer;
import java.nio.ShortBuffer;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.List;

/**
 * Provide a translation layer between the block vector and a byte array.
 * <p/>
 * <p/>
 * A record (Storage, XLSRecord) can retrieve bytes from Scattered blocks
 * transparently.
 * <p/>
 * The BlockByteReader Allocates a ByteBuffer containing only byte references
 * contained within Blocs assigned to the implementation class.
 * <p/>
 * A Class using this reader will either subclass the reader and manage its
 * blocks, or interact with a shared reader.
 * <p/>
 * In the case of a shared reader,
 */
public class BlockByteReader implements Serializable
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 4845306509411520019L;
	private boolean applyRelativePosition = true;

	protected BlockByteReader()
	{
		// empty constructor...
	}

	private List blockmap = new ArrayList();
	private boolean ro = false;
	private int length = -1;
	transient ByteBuffer backingByteBuffer = ByteBuffer.allocate( 0 );

	public BlockByteReader( List blokz, int len )
	{
		this.blockmap = blokz;
		this.length = len;
	}

	public boolean isReadOnly()
	{
		return ro;
	}

	/**
	 * Allows for getting of header bytes without setting blocks on a rec
	 *
	 * @param startpos
	 * @return
	 */
	public byte[] getHeaderBytes( int startpos )
	{
		try
		{
			int SIZE = BIGBLOCK.SIZE; // normal case
			if( this.length < StorageTable.BIGSTORAGE_SIZE )
			{
				SIZE = SMALLBLOCK.SIZE;
			}
			int block = startpos / SIZE;
			int check = startpos % SIZE;
			// handle EOF that falls right on boundary
			if( ((check + 4) > SIZE) && ((blockmap.size() - 1) == block) )
			{
				// Last EOF falls within 4 bytes of 512 boundary... junkrec
				byte[] junk = { 0x0, 0x0, 0x0, 0x0 };
				return junk;
			}
			else if( (check + 4) > SIZE )
			{ // SPANNER!
				Block bx = (Block) this.blockmap.get( block );
				int l1 = ((SIZE * (block + 1)) - startpos);
				int s2 = startpos % SIZE;
				byte[] b1 = bx.getBytes( s2, s2 + l1 );
				bx = (Block) this.blockmap.get( block + 1 );
				l1 = 4 - l1;
				byte[] b2 = bx.getBytes( 0, l1 );
				return ByteTools.append( b2, b1 );
			}

			Block bx = (Block) this.blockmap.get( block );
			startpos -= (block * SIZE);
			return bx.getBytes( startpos, startpos + 4 );
		}
		catch( RuntimeException e )
		{
			throw new WorkBookException(
					"Smallblock based workbooks are unsupported in ExtenXLS: see http://extentech.com/uimodules/docs/docs_detail.jsp?showall=true&meme_id=195",
					WorkBookException.SMALLBLOCK_FILE );
		}
	}

	/* Return the byte from the blocks at the proper locations...
	 * 
	 * @see java.nio.ByteBuffer#get()
	 */
	public byte get( BlockByteConsumer rec, int startpos )
	{
		byte ret = this.get( rec, startpos, 1 )[0];
		return ret;
	}

	/* Return the bytes from the blocks at the proper locations...
	* 
	* as opposed to when we are traversing the entire collection of bytes as in WorkBookFactory.parse()
	* 
	* @ see java.nio.ByteBuffer # get()
	*/
	public byte[] get( BlockByteConsumer rec, int startpos, int len )
	{
		rec.setByteReader( this );
		//	we only want to add the offset when
		//  we are fetching data from 'within' a record, ie: rkdata.get(i)
		int recoffy = rec.getOffset();
		if( this.getApplyRelativePosition() )
		{

			startpos += 4; // add the offset
		}
		startpos += recoffy;
		// reality checks
		if( false ) // ((startpos + len) > getLength())
		{
			Logger.logWarn( "WARNING: BlockByteReader.get(rec," + startpos + "," + rec.getLength() + ") error.  Attempt to read past end of Block buffer." );
		}

		// return the bytes from the rec
		return getRecBytes( rec, startpos, len );
	}

	/* Handles the spanning of Record Bytes over Block boundaries
	 * then returns requested bytes.
	 * 
	 */
	private byte[] getRecBytes( BlockByteConsumer rec, int startpos, int len )
	{
		if( (startpos < 0) || (startpos > (startpos + len)) )
		{
			throw new RuntimeException( "ERROR: BBR.getRecBytes(" + rec.getClass()
			                                                           .getName() + "," + startpos + "," + (startpos + len) + ") failed - OUT OF BOUNDS." );
		}
		// get the block byte boundaries
		int[] pos = this.getReadPositions( startpos, len );
		int numblocks = pos.length / 3;
		int blkdef = 0;
		//	backingByteBuffer = blokx.getByteBuffer();
		//	Temporarily use BAOS...
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		for( int t = 0; t < numblocks; t++ )
		{
			try
			{ // inlining byte read
				Block b1 = (Block) this.blockmap.get( pos[blkdef++] );
				out.write( b1.getBytes( pos[blkdef++], pos[blkdef++] ) );
				if( false )
				{
					Logger.logInfo( "INFO: BBR.getRecBytes() " + rec.getClass()
					                                                .getName() + " ACCESSING DATA for block:" + b1.getBlockIndex() + ":" + pos[0] + "-" + pos[1] );
				}
			}
			catch( Exception a )
			{
				Logger.logWarn( "ERROR: BBR.getRecBytes streaming " + rec.toString() + " bytes for block failed: " + a );
			}
		}
		return out.toByteArray();
	}

	/**
	 * returns the lengths of the two byte
	 * <pre>
	 * ie: start 10 len 8
	 * blk0 = 10-18
	 * 0,10,18
	 *
	 * ie: start 514 len 1
	 * blk1 = 2-3
	 * 1,2,3
	 *
	 * ie: start 480 len 1124
	 *
	 *  blk0 = 480-512
	 *  blk1 = 512-1024
	 *  blk2 = 1024-1536
	 *  blk3 = 1536-1604
	 *
	 * [0, 480, 512,1, 0, 512, 2, 0, 512, 3, 0, 68]
	 *
	 * 8139, 693
	 * </pre>
	 *
	 * @param startpos
	 * @return
	 */
	public static int[] getReadPositions( int startpos, int len, boolean BIGBLOCKSTORAGE )
	{
		//	Logger.logInfo("BBR.getReadPositions()"+startpos+":"+endpos);

		// 20100323 KSC: handle small blocks
		int SIZE = BIGBLOCK.SIZE; // normal case
		if( !BIGBLOCKSTORAGE )
		{
			SIZE = SMALLBLOCK.SIZE;
		}

		int firstblock = startpos / SIZE;
		int lastblock = (startpos + len) / SIZE;
		int numblocks = lastblock - firstblock;
		numblocks++;
		int origlen = len;

		int ct = startpos / SIZE;

		int pos1 = startpos;
		int[] ret = new int[numblocks * 3];
		int t = 0;
		// for each block, create 2 byte positions
		while( len > 0 )
		{
			if( t >= ret.length )
			{
				Logger.logWarn( "BlockByteReader.getReadPositions() wrong guess on NumBlocks." );
				numblocks++;
				int[] retz = new int[numblocks * 3];
				System.arraycopy( ret, 0, retz, 0, ret.length );
				ret = retz;
			}
			ret[t++] = firstblock++;
			int check = pos1 % SIZE; //leftover
			check += len;
			if( check > SIZE )
			{ // SPANNER!
				pos1 = startpos - ((SIZE) * ct);
				if( pos1 < 0 )
				{
					pos1 = 0;
				}
				// int s1 = pos1- ((SIZE)*(ct));
				int s2 = pos1 % SIZE;
				if( s2 < 0 )
				{
					s2 = 0;
					pos1 = 0;
				}
				ret[t++] = s2;
				ret[t++] = SIZE;
			}
			else
			{
				pos1 = startpos - ((SIZE) * ct);
				int strt = startpos - ((SIZE) * ct);
				if( strt < 0 )
				{
					ret[t++] = 0;
					ret[t++] = len;
				}
				else
				{
					ret[t++] = strt;
					ret[t++] = (startpos + origlen) - ((SIZE) * ct);
				}

			}
			ct++;
			int ctdn = ret[t - 1] - pos1;
			len -= (ctdn);
			pos1 = 0;//startpos;
		}
		return ret;
	}

	/**
	 * Gets the list of blocks needed to read the given sequence.
	 */
	public int[] getReadPositions( int startpos, int len )
	{
		return getReadPositions( startpos, len, (this.length >= StorageTable.BIGSTORAGE_SIZE) );
	}

	/**
	 * Gets the mapping from stream offsets to file offsets over the given
	 * range.
	 * <p/>
	 * This returns an array of integers arranged in pairs. The first value of
	 * each pair is an offset from <code>start</code> and the second is the
	 * corresponding offset in the source file.
	 */
	public int[] getFileOffsets( int start, int size )
	{
		int[] smap = this.getReadPositions( start, size );
		int[] fmap = new int[(smap.length / 3) * 2];

		int offset = 0;
		int fidx = 0;
		Block block = null;
		Block prev;
		for( int sidx = 0; sidx < smap.length; sidx += 3 )
		{
			prev = block;
			block = (Block) this.blockmap.get( smap[sidx] );

			if( (prev == null) || ((block.getOriginalPos() + smap[sidx + 1]) != (prev.getOriginalPos() + smap[sidx - 1])) )
			{
				fmap[fidx++] = offset;
				fmap[fidx++] = block.getOriginalPos() + smap[sidx + 1];
			}

			offset += smap[sidx + 2] - smap[sidx + 1];
		}

		int[] ret = new int[fidx];
		System.arraycopy( fmap, 0, ret, 0, fidx );
		return ret;
	}

	public String getFileOffsetString( int start, int size )
	{
		String ret = "";
		int[] map = this.getFileOffsets( start, size );

		for( int idx = 0; idx < map.length; idx += 2 )
		{
			ret += ((idx != 0) ? " " : "") + Integer.toHexString( map[idx + 0] ).toUpperCase() + ":" + Integer.toHexString( map[idx + 1] )
			                                                                                                .toUpperCase();
		}

		return ret;
	}

	/**Assign blocks to the recs
	 *
	 * @param rec
	 * @param startpos
	 * @return
	 */
/* Not used
	public int setBlocks(BlockByteConsumer rec, int startpos) {
		if (rec.getBlocks() != null)
			return rec.getFirstBlock();

		int reclen = rec.getLength();

		int numblocks = -1, firstblock = -1, lastblock = -1;

		// Handle assignment of blocks
		//	span blocks on either side of the block boundary
		if (this.getApplyRelativePosition()) {
			firstblock = rec.getOffset();
			firstblock /= BIGBLOCK.SIZE;
			lastblock = reclen + rec.getOffset();
			lastblock /= BIGBLOCK.SIZE;
		} else {
			firstblock = startpos;
			firstblock /= BIGBLOCK.SIZE;
			lastblock = (reclen + startpos);
			lastblock /= BIGBLOCK.SIZE;
		}

 		numblocks = lastblock - (firstblock - 1);

		if (numblocks < 1)
			numblocks = 1;

		Block[] blks = new Block[numblocks];
		for (int t = 0; t < numblocks; t++) {
			try { // inlining byte read
				int getblk = firstblock + t;
				//Logger.logInfo("GETTING BLOCK: "+ getblk);
				blks[t] = (Block) this.getBlockmap().get(getblk);
			} catch (Exception a) {
				Logger.logWarn(
					"ERROR: Bytes for block:"
						+ blks[t]
						+ " failed: "
						+ a
						+ " Output Corrupted.");
			}
		}
		rec.setFirstBlock(firstblock);
		rec.setLastBlock(lastblock);
		rec.setBlocks(blks);
		return firstblock;
	}
*/

	/**
	 * For whatever reason, get the blockmap
	 *
	 * @return the map of blocks
	 */
	public List getBlockmap()
	{
		return blockmap;
	}

	/**
	 * Set the map of Blocks contained in this reader
	 *
	 * @param list the map of blocks
	 */
	public void setBlockmap( AbstractList list )
	{
		blockmap = list;
	}

	/**
	 * @return
	 */
	public int getLength()
	{
		return length;
	}

	/**
	 * @param i
	 */
	public void setLength( int i )
	{
		length = i;
	}

	/**
	 * @param arg0
	 * @return
	 */
	public static ByteBuffer allocate( int arg0 )
	{
		return ByteBuffer.allocate( arg0 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public static ByteBuffer allocateDirect( int arg0 )
	{
		return ByteBuffer.allocateDirect( arg0 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public final static ByteBuffer wrap( byte[] arg0 )
	{
		return ByteBuffer.wrap( arg0 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @param arg2
	 * @return
	 */
	public final static ByteBuffer wrap( byte[] arg0, int arg1, int arg2 )
	{
		return ByteBuffer.wrap( arg0, arg1, arg2 );
	}

	/**
	 * @return
	 */
	public byte[] array()
	{
		return backingByteBuffer.array();
	}

	/**
	 * @return
	 */
	public int arrayOffset()
	{
		return backingByteBuffer.arrayOffset();
	}

	/**
	 * @return
	 */
	public CharBuffer asCharBuffer()
	{
		return backingByteBuffer.asCharBuffer();
	}

	/**
	 * @return
	 */
	public DoubleBuffer asDoubleBuffer()
	{
		return backingByteBuffer.asDoubleBuffer();
	}

	/**
	 * @return
	 */
	public FloatBuffer asFloatBuffer()
	{
		return backingByteBuffer.asFloatBuffer();
	}

	/**
	 * @return
	 */
	public IntBuffer asIntBuffer()
	{
		return backingByteBuffer.asIntBuffer();
	}

	/**
	 * @return
	 */
	public LongBuffer asLongBuffer()
	{
		return backingByteBuffer.asLongBuffer();
	}

	/**
	 * @return
	 */
	public ByteBuffer asReadOnlyBuffer()
	{
		return backingByteBuffer.asReadOnlyBuffer();
	}

	/**
	 * @return
	 */
	public ShortBuffer asShortBuffer()
	{
		return backingByteBuffer.asShortBuffer();
	}

	/**
	 * @return
	 */
	public int capacity()
	{
		return backingByteBuffer.capacity();
	}

	/**
	 * @return
	 */
	public Buffer clear()
	{
		return backingByteBuffer.clear();
	}

	/**
	 * @return
	 */
	public ByteBuffer compact()
	{
		return backingByteBuffer.compact();
	}

	/**
	 * @param arg0
	 * @return public int compareTo(Object arg0) {
	return backingByteBuffer.compareTo(arg0);
	}*/

	/**
	 * @return
	 */
	public ByteBuffer duplicate()
	{
		return backingByteBuffer.duplicate();
	}

	/* (non-Javadoc)
	 * @see java.lang.Object#equals(java.lang.Object)
	 */
	public boolean equals( Object arg0 )
	{
		return backingByteBuffer.equals( arg0 );
	}

	/**
	 * @return
	 */
	public Buffer flip()
	{
		return backingByteBuffer.flip();
	}

	/**
	 * @return
	 */
	public byte get()
	{
		return backingByteBuffer.get();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer get( byte[] arg0 )
	{
		return backingByteBuffer.get( arg0 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @param arg2
	 * @return
	 */
	public ByteBuffer get( byte[] arg0, int arg1, int arg2 )
	{
		return backingByteBuffer.get( arg0, arg1, arg2 );
	}

	/**
	 * @return
	 */
	public char getChar()
	{
		return backingByteBuffer.getChar();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public char getChar( int arg0 )
	{
		return backingByteBuffer.getChar( arg0 );
	}

	/**
	 * @return
	 */
	public double getDouble()
	{
		return backingByteBuffer.getDouble();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public double getDouble( int arg0 )
	{
		return backingByteBuffer.getDouble( arg0 );
	}

	/**
	 * @return
	 */
	public float getFloat()
	{
		return backingByteBuffer.getFloat();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public float getFloat( int arg0 )
	{
		return backingByteBuffer.getFloat( arg0 );
	}

	/**
	 * @return
	 */
	public int getInt()
	{
		return backingByteBuffer.getInt();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public int getInt( int arg0 )
	{
		return backingByteBuffer.getInt( arg0 );
	}

	/**
	 * @return
	 */
	public long getLong()
	{
		return backingByteBuffer.getLong();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public long getLong( int arg0 )
	{
		return backingByteBuffer.getLong( arg0 );
	}

	/**
	 * @return
	 */
	public short getShort()
	{
		return backingByteBuffer.getShort();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public short getShort( int arg0 )
	{
		return backingByteBuffer.getShort( arg0 );
	}

	/**
	 * @return
	 */
	public boolean hasArray()
	{
		return backingByteBuffer.hasArray();
	}

	/* (non-Javadoc)
	 * @see java.lang.Object#hashCode()
	 */
	public int hashCode()
	{
		return backingByteBuffer.hashCode();
	}

	/**
	 * @return
	 */
	public boolean hasRemaining()
	{
		return backingByteBuffer.hasRemaining();
	}

	/**
	 * @return
	 */
	public boolean isDirect()
	{
		return backingByteBuffer.isDirect();
	}

	/**
	 * @return
	 */
	public int limit()
	{
		return backingByteBuffer.limit();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public Buffer limit( int arg0 )
	{
		return backingByteBuffer.limit( arg0 );
	}

	/**
	 * @return
	 */
	public Buffer mark()
	{
		return backingByteBuffer.mark();
	}

	/**
	 * @return
	 */
	public ByteOrder order()
	{
		return backingByteBuffer.order();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer order( ByteOrder arg0 )
	{
		return backingByteBuffer.order( arg0 );
	}

	/**
	 * @return
	 */
	public int position()
	{
		return backingByteBuffer.position();
	}

	/**
	 * @param arg0
	 * @return
	 */
	public Buffer position( int arg0 )
	{
		return backingByteBuffer.position( arg0 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer put( byte arg0 )
	{
		return backingByteBuffer.put( arg0 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer put( byte[] arg0 )
	{
		return backingByteBuffer.put( arg0 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @param arg2
	 * @return
	 */
	public ByteBuffer put( byte[] arg0, int arg1, int arg2 )
	{
		return backingByteBuffer.put( arg0, arg1, arg2 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @return
	 */
	public ByteBuffer put( int arg0, byte arg1 )
	{
		return backingByteBuffer.put( arg0, arg1 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer put( ByteBuffer arg0 )
	{
		return backingByteBuffer.put( arg0 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer putChar( char arg0 )
	{
		return backingByteBuffer.putChar( arg0 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @return
	 */
	public ByteBuffer putChar( int arg0, char arg1 )
	{
		return backingByteBuffer.putChar( arg0, arg1 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer putDouble( double arg0 )
	{
		return backingByteBuffer.putDouble( arg0 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @return
	 */
	public ByteBuffer putDouble( int arg0, double arg1 )
	{
		return backingByteBuffer.putDouble( arg0, arg1 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer putFloat( float arg0 )
	{
		return backingByteBuffer.putFloat( arg0 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @return
	 */
	public ByteBuffer putFloat( int arg0, float arg1 )
	{
		return backingByteBuffer.putFloat( arg0, arg1 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer putInt( int arg0 )
	{
		return backingByteBuffer.putInt( arg0 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @return
	 */
	public ByteBuffer putInt( int arg0, int arg1 )
	{
		return backingByteBuffer.putInt( arg0, arg1 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @return
	 */
	public ByteBuffer putLong( int arg0, long arg1 )
	{
		return backingByteBuffer.putLong( arg0, arg1 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer putLong( long arg0 )
	{
		return backingByteBuffer.putLong( arg0 );
	}

	/**
	 * @param arg0
	 * @param arg1
	 * @return
	 */
	public ByteBuffer putShort( int arg0, short arg1 )
	{
		return backingByteBuffer.putShort( arg0, arg1 );
	}

	/**
	 * @param arg0
	 * @return
	 */
	public ByteBuffer putShort( short arg0 )
	{
		return backingByteBuffer.putShort( arg0 );
	}

	/**
	 * @return
	 */
	public int remaining()
	{
		return backingByteBuffer.remaining();
	}

	/**
	 * @return
	 */
	public Buffer reset()
	{
		return backingByteBuffer.reset();
	}

	/**
	 * @return
	 */
	public Buffer rewind()
	{
		return backingByteBuffer.rewind();
	}

	/**
	 * @return
	 */
	public ByteBuffer slice()
	{
		return backingByteBuffer.slice();
	}

	/* (non-Javadoc)
	 * @see java.lang.Object#toString()
	 */
	public String toString()
	{
		return backingByteBuffer.toString();
	}

	/**
	 * @return
	 */
	public ByteBuffer getBackingByteBuffer()
	{
		return backingByteBuffer;
	}

	/**
	 * @param buffer
	 */
	public void setBackingByteBuffer( ByteBuffer buffer )
	{
		backingByteBuffer = buffer;
	}

	/**
	 * @return
	 */
	public boolean getApplyRelativePosition()
	{
		return applyRelativePosition;
	}

	/**
	 * Only add the offset when we are fetching
	 * data from 'within' a record, ie: rkdata.get(i)
	 * as opposed to when we are traversing the entire collection
	 * of bytes as in WorkBookFactory.parse()...
	 * <p/>
	 * * @param b
	 */
	public void setApplyRelativePosition( boolean b )
	{
		applyRelativePosition = b;
	}

}
