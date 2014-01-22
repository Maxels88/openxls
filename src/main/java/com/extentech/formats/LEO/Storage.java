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

import com.extentech.formats.XLS.XLSConstants;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.Logger;

import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.nio.ByteBuffer;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

/**
 * Defines a 'file' in the LEO filesystem.  contains pointers
 * to Block storages as well as the storage type etc.
 * <p/>
 * <p/>
 * Header (128 bytes):
 * Directory Entry Name (64 bytes): This field MUST contain a Unicode string for the storage or stream name encoded in UTF-16. The name MUST be terminated with a UTF-16 terminating null character. Thus storage and stream names are limited to 32 UTF-16 code points, including the terminating null character. When locating an object in the compound file except for the root storage, the directory entry name is compared using a special case-insensitive upper-case mapping, described in Red-Black Tree. The following characters are illegal and MUST NOT be part of the name: '/', '\', ':', '!'.
 * Directory Entry Name Length (2 bytes): This field MUST match the length of the Directory Entry Name Unicode string in bytes. The length MUST be a multiple of 2, and include the terminating null character in the count. This length MUST NOT exceed 64, the maximum size of the Directory Entry Name field.
 * Object Type (1 byte): This field MUST be 0x00, 0x01, 0x02, or 0x05, depending on the actual type of object. All other values are not valid.
 * 0= Unknown or unallocated
 * 1= Storage Object
 * 2- Stream Object
 * 5= Root Storage Object
 * Color Flag (1 byte): This field MUST be 0x00 (red) or 0x01 (black). All other values are not valid.
 * Left Sibling ID (4 bytes): This field contains the Stream ID of the left sibling. If there is no left sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).
 * Right Sibling ID (4 bytes): This field contains the Stream ID of the right sibling. If there is no right sibling, the field MUST be set to NOSTREAM (0xFFFFFFFF).
 * Child ID (4 bytes): This field contains the Stream ID of a child object. If there is no child object, then the field MUST be set to NOSTREAM (0xFFFFFFFF).
 * CLSID (16 bytes): This field contains an object classGUID, if this entry is a storage or root storage. If there is no object class GUID set on this object, then the field MUST be set to all zeroes. In a stream object, this field MUST be set to all zeroes. If not NULL, the object class GUID can be used as a parameter to launch applications.
 * State Bits (4 bytes): This field contains the user-defined flags if this entry is a storage object or root storage object. If there are no state bits set on the object, then this field MUST be set to all zeroes.
 * Creation Time (8 bytes): This field contains the creation time for a storage object. The Windows FILETIME structure is used to represent this field in UTC. If there is no creation time set on the object, this field MUST be all zeroes. For a root storage object, this field MUST be all zeroes, and the creation time is retrieved or set on the compound file itself.
 * Modified Time (8 bytes): This field contains the modification time for a storage object. The Windows FILETIME structure is used to represent this field in UTC. If there is no modified time set on the object, this field MUST be all zeroes. For a root storage object, this field MUST be all zeroes, and the modified time is retrieved or set on the compound file itself.
 * Starting Sector Location (4 bytes): This field contains the first sector location if this is a stream object. For a root storage object, this field MUST contain the first sector of the mini st	ream, if the mini stream exists.
 * Stream Size (8 bytes): This 64-bit integer field contains the size of the user-defined data, if this is a stream object. For a root storage object, this field contains the size of the mini stream.
 * <p/>
 * ??
 * offset  type value           const?  function
 * 00:     stream $pps_rawname       !  name of the pps
 * 40:     word $pps_sizeofname      !  size of $pps_rawname
 * 42:     byte $pps_type		      !  type of pps (1=storage|2=stream|5=root)
 * 43:	    byte $pps_uk0		      !  ?
 * 44:	    long $pps_prev	          !  previous pps
 * 48:     long $pps_next            !  next pps
 * 4c:     long $pps_dir             !  directory pps
 * 50:     stream 00 09 02 00        .  ?
 * 54:     long 0                    .  ?
 * 58:     long c0                   .  ?
 * 5c:     stream 00 00 00 46        .  ?
 * 60:     long 0                    .  ?
 * 64:     long $pps_ts1s            !  timestamp 1 : "seconds"		creation time
 * 68:     long $pps_ts1d            !  timestamp 1 : "days"
 * 6c:     long $pps_ts2s            !  timestamp 2 : "seconds"		modified time
 * 70:     long $pps_ts2d            !  timestamp 2 : "days"
 * 74:     long $pps_sb              !  starting block of property
 * 78:     long $pps_size            !  size of property
 * 7c:     long                      .  ?
 */
public class Storage extends BlockByteReader
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2065921767253066667L;
	// make an enum in > 1.4
	static byte TYPE_INVALID = 0;
	static byte TYPE_DIRECTORY = 1;
	static byte TYPE_STREAM = 2;
	static byte TYPE_LOCKBYTES = 3;
	static byte TYPE_PROPERTY = 4;
	static byte TYPE_ROOT = 5;

	byte DIR_COLOR_RED = 0;
	byte DIR_COLOR_BLACK = 1;

	private transient ByteBuffer headerData = ByteBuffer.allocate( 128 );

	// properties of this storage file.
	String name = "";
	int nameSize = -1;

	public byte storageType = -1;
	public byte directoryColor = -1;
	public int prevStorageID = -1;
	public int nextStorageID = -1;
	public int childStorageID = -1;

	public int sz = -1;
	public boolean miniStreamStorage = false;
	private boolean isSpecial = false;

	protected List myblocks;
	private int startBlock = 0;
	private int SIZE = -1;
	private int blockType = -1;
	private boolean initialized = false;
	AbstractList idxs = new CompatibleVector();

	Block lastblock = null;

	public void setBlocks( Block[] blks )
	{
		myblocks = new ArrayList();
		for( int t = 0; t < blks.length; t++ )
		{
			this.addBlock( blks[t] );
		}
	}

	/**
	 * returns a new BlockByteReader
	 *
	 * @return
	 */
	public BlockByteReader getBlockReader()
	{
		BlockByteReader ret = new BlockByteReader( myblocks, this.getActualFileSize() );
		return ret;
	}

	public List getBlockVect()
	{
		return myblocks;
	}

	public void setIsSpecial( boolean b )
	{
		this.isSpecial = b;
	}

	public boolean getIsSpecial()
	{
		return this.isSpecial;
	}

	public boolean getInitialized()
	{
		return this.initialized;
	}

	public void setInitialized( boolean b )
	{
		this.initialized = b;
	}

	/**
	 * remove a block from this Storage's headerData
	 */
	void removeBlock( Block b )
	{
		myblocks.remove( b );
	}

	/**
	 * sets whether this Storage's headerData blocks are contained
	 * in the Small or Big Block arrays.
	 */
	public void setBlockType( int type )
	{
		this.blockType = type;
		switch( type )
		{
			case Block.BIG:
				SIZE = BIGBLOCK.SIZE;
				break;
			case Block.SMALL:
				SIZE = SMALLBLOCK.SIZE;
				break;
		}
	}

	/**
	 * returns whether this Storage's headerData blocks are contained
	 * in the Small or Big Block arrays.
	 */
	public int getBlockType()
	{
		return this.blockType;
	}

	public int getStorageType()
	{
		return storageType;
	}

	public int getDirectoryColor()
	{
		return directoryColor;
	}

	/**
	 * set the value of the prevProp variable
	 */
	public void setDirectoryColor( int o )
	{
		int pos = 0x43;
		headerData.position( pos );
		headerData.put( (byte) o );
		this.directoryColor = (byte) o;
	}

	public String getName()
	{
		return name;
	}

	public String toString()
	{
		return getName() + " n:" + nextStorageID + " p:" + prevStorageID + " c:" + childStorageID + " sz:" + sz;
	}

	/**
	 * sets the storage name
	 *
	 * @param nm
	 */
	public void setName( String nm )
	{
		try
		{
			byte[] b = nm.getBytes( XLSConstants.UNICODEENCODING );
			int pos = 0;                // unicode name bytes
			headerData.position( pos );
			headerData.put( b );
			pos = 0x40;                    // short name size
			headerData.position( pos );
			headerData.putShort( (short) (b.length + 2) );
		}
		catch( UnsupportedEncodingException e )
		{
		}
		name = nm;
	}

	/**/
	protected Storage()
	{
		// empty constructor...
	}

	/**
	 * create a new Storage record with a byte
	 * array containing its record headerData.
	 */
	protected Storage( ByteBuffer buff )
	{
		headerData = buff;
		int pos = 0x40;
		headerData.position( pos );
		nameSize = headerData.getShort();
		if( nameSize > 0 )
		{
			// get the name
			pos = 0;
			byte[] namebuf = new byte[nameSize];
			try
			{
				for( int i = 0; i < nameSize; i++ )
				{
					namebuf[i] = headerData.get( i );
				}
			}
			catch( Exception e )
			{
				;
			}
			try
			{
				name = new String( namebuf, XLSConstants.UNICODEENCODING );
			}
			catch( UnsupportedEncodingException e )
			{
				Logger.logWarn( "Storage error decoding storage name " + e );
			}
			name = name.substring( 0, name.length() - 1 );
			// if this line fails, your header BBD index is wrong...

		}
		else
		{
			// empty storage   
		}
		pos = 0x42;
		headerData.position( pos );
		storageType = headerData.get();
		directoryColor = headerData.get();
		prevStorageID = headerData.getInt();
		nextStorageID = headerData.getInt();
		childStorageID = headerData.getInt();

		sz = this.getActualFileSize();
		if( sz > 0 && sz < BIGBLOCK.SIZE )
		{
			miniStreamStorage = true;
		}
		if( LEOFile.DEBUG )
		{
			Logger.logInfo( "Storage: " + name + " storageType: " + storageType + " directoryColor:" + directoryColor +
					                " prevSID:" + prevStorageID + " nextSID:" + nextStorageID + " childSID:" + childStorageID + " sz:" + sz );
		}
	}

	/**
	 * get the value of the prevProp variable
	 */
	public int getPrevStorageID()
	{
		return this.prevStorageID;
	}

	/**
	 * set the value of the prevProp variable
	 */
	public void setPrevStorageID( int o )
	{
		int pos = 0x44;
		headerData.position( pos );
		headerData.putInt( o );
		this.prevStorageID = o;
	}

	/**
	 * get the value of the nextProp variable
	 */
	public int getNextStorageID()
	{
		return this.nextStorageID;
	}

	/**
	 * set the value of the nextProp variable
	 */
	public void setNextStorageID( int o )
	{
		int pos = 0x48;
		headerData.position( pos );
		headerData.putInt( o );
		this.nextStorageID = o;
	}

	/**
	 * get the value of the child storage id
	 */
	public int getChildStorageID()
	{
		return this.childStorageID;
	}

	/**
	 * set the value of the child storage id
	 */
	public void setChildStorageID( int o )
	{
		int pos = 0x4C;
		headerData.position( pos );
		headerData.putInt( o );
		this.childStorageID = o;
	}

	/**
	 * get the position of this Storage in the file bytes
	 */
	public int getFilePos()
	{
		int ret = (this.getStartBlock() + 1) * SIZE;
		return ret;
	}

	/**
	 * return the existing header headerData for this storage
	 */
	public ByteBuffer getHeaderData()
	{
		return headerData;
	}

	/**
	 * return the underlying byte array for this
	 * Storage
	 */
	public byte[] getBytes()
	{
		return LEOFile.getBytes( myblocks );
	}

	/**
	 * return the underlying byte array for this
	 * Storage
	 */
	public OutputStream getByteStream()
	{
		return LEOFile.getByteStream( myblocks );
	}

	/**
	 * return the underlying byte array for this
	 * Storage
	 */
	public void writeBytes( OutputStream out )
	{
		Iterator itx = myblocks.iterator();
		while( itx.hasNext() )
		{
			((Block) itx.next()).writeBytes( out );
		}
	}

	/**
	 * set the underlying byte array for this
	 * Storage
	 */
	public void writeBytes( OutputStream out, int blen )
	{
		Block[] bs = BlockFactory.getBlocksFromOutputStream( out, blen, Block.BIG );
		if( myblocks != null )
		{
			myblocks.clear();
			lastblock = null;
		}
		else
		{
			myblocks = new ArrayList( bs.length );
		}
		for( int d = 0; d < bs.length; d++ )
		{
			this.addBlock( bs[d] );
		}
	}

	/**
	 * set the underlying byte array for this
	 * Storage
	 */
	public void setOutputBytes( OutputStream b, int blen )
	{
		Block[] bs = BlockFactory.getBlocksFromOutputStream( b, blen, Block.BIG );
		if( myblocks != null )
		{
			myblocks.clear();
			lastblock = null;
		}
		else
		{
			myblocks = new ArrayList( bs.length );
		}
		for( int d = 0; d < bs.length; d++ )
		{
			this.addBlock( bs[d] );
		}
	}

	/**
	 * set the underlying byte array for this
	 * Storage
	 */
	public void setBytes( byte[] b )
	{
		if( this.miniStreamStorage )
		{
			this.setMiniFATSectorBytes( b );
			return;
		}
		Block[] bs = BlockFactory.getBlocksFromByteArray( b, Block.BIG );
		if( myblocks != null )
		{
			myblocks.clear();
			lastblock = null;
		}
		else
		{
			myblocks = new ArrayList( bs.length );
		}
		for( int d = 0; d < bs.length; d++ )
		{
			this.addBlock( bs[d] );
		}
	}

	/**
	 * Sets bytes on a miniFAT storage
	 *
	 * @param b
	 */
	private void setMiniFATSectorBytes( byte[] b )
	{
		Block[] bs = BlockFactory.getBlocksFromByteArray( b, Block.SMALL );
		if( myblocks != null )
		{
			myblocks.clear();
			lastblock = null;
		}
		else
		{
			myblocks = new ArrayList( bs.length );
		}
		for( int d = 0; d < bs.length; d++ )
		{
			this.addBlock( bs[d] );
		}
	}

	/**
	 * sets bytes for this storage; length of newbytes determines
	 * whether the storage is miniFAT or regular
	 * <br>Incldes padding of bytes to ensure blocks are a factor of required block size
	 *
	 * @param newbytes
	 */
	public void setBytesWithOverage( byte[] newbytes )
	{
		int actuallen = newbytes.length;
		myblocks = new ArrayList();    // clear out
		if( newbytes.length < StorageTable.BIGSTORAGE_SIZE )
		{    // usual case
			int overage = newbytes.length % 128;
			if( overage > 0 )
			{
				byte[] b = new byte[128 - overage];
				newbytes = ByteTools.append( b, newbytes );
			}
			Block[] smallblocks = BlockFactory.getBlocksFromByteArray( newbytes, Block.SMALL );
			for( int i = 0; i < smallblocks.length; i++ )
			{
				this.addBlock( smallblocks[i] );
			}
			this.setBlockType( Block.SMALL );
		}
		else
		{
			int overage = newbytes.length % BIGBLOCK.SIZE;
			if( overage > 0 )
			{
				byte[] b = new byte[BIGBLOCK.SIZE - overage];
				newbytes = ByteTools.append( b, newbytes );
			}
			Block[] blocks = BlockFactory.getBlocksFromByteArray( newbytes, Block.BIG );
			for( int i = 0; i < blocks.length; i++ )
			{
				this.addBlock( blocks[i] );
			}
			this.setBlockType( Block.BIG );
		}
		this.setActualFileSize( actuallen );
	}

	/**
	 * Associate this storage with it's data
	 * (obtained by walking the miniFAT sector index to reference blocks in the miniStream
	 */
	public void initFromMiniStream( List miniStream, int[] miniFAT ) throws LEOIndexingException
	{
		if( this.getStartBlock() < 0 )
		{
			return;
		}

		if( miniFAT == null )
		{ // error: trying to access smallblocks but no smallblock container found
			if( LEOFile.DEBUG )
			{
				Logger.logWarn( "initMiniFAT: no miniFAT container found" );
			}
			return;
		}
		myblocks = new ArrayList();
		boolean endloop = false;
		Block thisBlock = null;

		int idx = this.getStartBlock();
		while( idx >= 0 )
		{
			switch( idx )
			{
				case -1:    // unused sector, shouldn't get to here
				case -2:    // end of sector marker, exit
					endloop = true;
					break;

				default:
					if( idx >= miniStream.size() )
					{
						Logger.logWarn( "MiniStream Error initting Storage: " + this.getName() );
					}
					else
					{
						thisBlock = (Block) miniStream.get( idx );    // miniFAT is 0-based (no header sector at position 0 as in regular FAT)
						this.addBlock( thisBlock );
					}
			}
			if( endloop )
			{
				break;
			}
			idx = miniFAT[idx];    // otherwise, walk the sector id chain
		}
		if( LEOFile.DEBUG )
		{
			if( (int) Math.ceil( this.getActualFileSize() / 64.0 ) != myblocks.size() )
			{
				Logger.logErr( "Number of miniStream Sectors does not equal storage size.  Expected: " + (int) Math.ceil( this.getActualFileSize() / 64.0 ) + ". Is: " + myblocks
						.size() );
			}
		}
		this.setInitialized( true );
	}

	/**
	 * associate this storage with its blocks of data
	 * (obtained by walking down the FAT sector index chain to obtain blocks from the dta block store)
	 * <p/>
	 * for some bizarre reason, at idx 30208 the extraDIFAT jumps 5420 to 35628
	 */
	public void init( List dta, int[] FAT, boolean keepStartBlock )
	{
		myblocks = new ArrayList();
		boolean endloop = false;
		if( getStartBlock() < 0 )
		{
			return;
		}
		Block thisbb = null;
		int nextIdx = 0; //, lastIdx = 0, specialOffset = 1;

		// ksc: for root block and miniFAT cont., we add start block to block list 
		if( keepStartBlock )
		{
			// for root storages, add rootstart block
			thisbb = (Block) dta.get( startBlock + 1 );
			this.addBlock( thisbb ); //;
		}
		for( int i = startBlock; i < FAT.length; )
		{
			nextIdx = FAT[i];

			switch( nextIdx )
			{

				case -4: // extraDIFAT sector
					Logger.logInfo( "INFO: Storage.init() encountered extra DIFAT sector." );
					break;

				case -3: // special block	= DIFAT - defines the FAT
					if( this.getActualFileSize() > 0 )
					{
						if( LEOFile.DEBUG )
						{
							Logger.logWarn( "WARNING: Storage.init() Special block containing headerData." );
						}
						this.setIsSpecial( true );

						thisbb = (Block) dta.get( i++ );
						if( !thisbb.getIsSpecialBlock() )
						{
							this.addBlock( thisbb ); //;
						}
						nextIdx = i;
					}
					else
					{
						endloop = true;
					}
					break;

				case -1: // unused
					endloop = true;
					break; // ksc

				case -2: // end of Storage - keep end block
					if( i + 1 < dta.size() )
					{
						// get the "padding" block for later retrieval
						thisbb = (Block) dta.get( i + 1 );
						if( thisbb == null )
						{
							break;
						}
						this.addBlock( thisbb ); //
						//}
					}
					endloop = true;
					break;

				default: // normal block
					if( dta.size() > nextIdx )
					{
						thisbb = (Block) dta.get( nextIdx );
					}
					if( thisbb == null )
					{
						break;
					}
					if( nextIdx != i + 1 )
					{
						//the next is a jumper, pickup the orphan
						if( LEOFile.DEBUG )
						{
							Logger.logInfo( "INFO: Storage init: jumper skipping: " + String.valueOf( i ) );
						}
						Block skipbb = (Block) dta.get( i + 1 );

						this.addBlock( skipbb ); //
					}
					else if( !thisbb.getIsSpecialBlock() )
					{ // just skip as probably a bbdix in the midst of the secid chain
						this.addBlock( thisbb ); //

					}
			}
			i = nextIdx;
			if( endloop )
			{
				break;
			}
		}

		if( LEOFile.DEBUG )
		{
			int sz = this.getActualFileSize();
			if( sz != 0 )
			{
				if( Math.ceil( sz / 512.0 ) != myblocks.size() )
				{
					Logger.logWarn( "Storage.init:  Number of blocks do not equal storage size" );
				}
			}
		}
		this.setInitialized( true );
	}

	/**
	 * adds a block of data to this storage
	 *
	 * @param b
	 */
	public void addBlock( Block b )
	{
		if( lastblock != null )
		{
			lastblock.setNextBlock( b );
		}

		if( b.getInitialized() )
		{
			if( LEOFile.DEBUG )
			{
				Logger.logWarn( "ERROR: " + this.toString() + " - Block is already initialized." );
			}
			return;
		}
		b.setStorage( this );
		b.setInitialized( true );
		if( myblocks == null )
		{
			myblocks = new ArrayList();
		}
		myblocks.add( b );
		lastblock = b;
	}

	/**
	 * set the storage type
	 */
	public void setStorageType( int i )
	{
		int pos = 0x42;
		headerData.position( pos );
		headerData.put( (byte) i );
		storageType = (byte) i;
	}

	/**
	 * set the size of the Storage headerData
	 */
	public void setActualFileSize( int i )
	{
		int pos = 0x78;
		this.sz = i;
		headerData.position( pos );
		headerData.putInt( i );
	}

	/**
	 * get the size of the Storage headerData
	 */
	public int getActualFileSize()
	{
		int pos = 0x78;
		headerData.position( pos );
		this.sz = headerData.getInt();
		return this.sz;
	}

	/**
	 * set the starting block in the Block table
	 * array for this Storage's headerData.
	 */
	public void setStartBlock( int i )
	{
		this.startBlock = i;
		int pos = 0x74;
		this.headerData.position( pos );
		this.headerData.putInt( i );
	}

	/**
	 * get the starting block for the Storage headerData
	 */
	public int getStartBlock()
	{
		int pos = 0x74;
		headerData.position( pos );
		startBlock = headerData.getInt();
		return startBlock;
	}

	/**
	 * return this Storage's existing headerData Blocks.
	 */
	Block[] getBlocks()
	{
		if( myblocks.size() < 1 )
		{
			return this.initBigBlocks();
		}
		Block[] blox = new Block[this.myblocks.size()];
		blox = (Block[]) myblocks.toArray( blox );
		return blox;
	}

	/**
	 * return this Storage's headerData in new Blocks
	 */
	private Block[] initBigBlocks()
	{
		// byte[] bb = this.getBytes();
		if( this.getLength() > 0 )
		{
			Block[] blks = BlockFactory.getBlocksFromByteArray( this.getBytes(), Block.BIG );
			int t = 0;
			myblocks.clear();
			for(; t < blks.length; t++ )
			{
				if( t + 1 < blks.length )
				{
					blks[t].setNextBlock( blks[t + 1] );
				}
				this.addBlock( blks[t] );
			}
			return blks;
		}
		return null;
	}

	/**
	 * get the number of Block records
	 * that this Storage needs to store its
	 * byte array
	 */
	public int getSizeInBlocks()
	{
		return myblocks.size();
	}

	/**
	 * Track BB Index info
	 */
	public void addIdx( int x )
	{
		idxs.add( Integer.valueOf( x ) );
	}

	public boolean equals( Object other )
	{
		if( other.toString().equals( this.toString() ) )
		{
			return true;
		}
		return false;
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		if( myblocks != null )
		{
			myblocks.clear();
		}
		if( idxs != null )
		{
			idxs.clear();
		}
		lastblock = null;
	}
}