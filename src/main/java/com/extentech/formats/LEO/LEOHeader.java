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

// NIO based API requires JDK1.3+

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Serializable;
import java.nio.ByteBuffer;

/**
 * Header record containing information on the Storage records in the LEOFile.
 * <p/>
 * Header (block 1) -- 512 (0x200) bytes
 * Field Description Offset Length Default value or const
 * FILETYPE Magic number identifying this as a LEO filesystem. 0x0000 Long 0xE11AB1A1E011CFD0
 * UK1 Unknown constant 0x0008 Integer 0
 * UK2 Unknown Constant 0x000C Integer 0
 * UK3 Unknown Constant 0x0014 Integer 0
 * UK4 Unknown Constant (revision?) 0x0018 Short 0x003B
 * UK5 Unknown Constant (version?) 0x001A Short 0x0003
 * UK6 Unknown Constant 0x001C Short -2
 * LOG_2_BIG_BLOCK_SIZE Log, base 2, of the big block size 0x001E Short 9 (2 ^ 9 = 512 bytes)
 * LOG_2_SMALL_BLOCK_SIZE Log, base 2, of the miniFAT size 0x0020 Integer 6 (2 ^ 6 = 64 bytes)
 * UK7 Unknown Constant 0x0024 Integer 0
 * UK8 Unknown Constant 0x0028 Integer 0
 * FAT_COUNT Number of elements in the FAT array 0x002C Integer required
 * PROPERTIES_START Block index of the first block of the property table 0x0030 Integer required
 * UK9 Unknown Constant 0x0034 Integer 0
 * UK10 Unknown Constant 0x0038 Integer 0x00001000
 * MINIFAT_START Block index of first big block containing the mini FATallocation table (MINIFAT) 0x003C Integer -2
 * UK11 Unknown Constant 0x0040 Integer 1
 * extraDIFAT Block index of the first block in the Extended Block Allocation Table 0x0044 Integer -2
 * extraDIFATCount Number of elements in the Extended File Allocation Table (to be added to the FAT) 0x0048 Integer 0
 * FAT_ARRAY Array of block indicies constituting the File Allocation Table (FAT) 0x004C, 0x0050, 0x0054 ... 0x01FC Integer[ ] -1 for unused elements, at least first element must be filled.
 * N/A Header block data not otherwise described in this table N/A
 */
public class LEOHeader implements Serializable
{
	private static final Logger log = LoggerFactory.getLogger( LEOHeader.class );

	private static final long serialVersionUID = -422489164065975273L;
	public final static int HEADER_SIZE = 0x200;
	private transient ByteBuffer data;
	private static int DIFATPOSITION = 76;    // DIFAT start position WITHIN HEADER (DIFAT=where to find the FAT i.e the index to the FAT)
	private int numFATSectors = -1;            // FAT= array of Sector numbers
	private int rootstart = -1;                // Root Directory sector start
	private int miniFATStart = -2;            // miniFAT stores sectors for storages < 4096
	private int extraDIFATStart = -1;        // if more than 109 FAT sectors are needed this stores start position of remaining DIFAT
	private int numExtraDIFATSectors = -1;    // usually 109 FAT sectors is enough; if > -1 more DIFAT is stored in other sectors
	private int numMiniFATSectors = 0;
	private int minStreamSize = 4096;    // = 8 blocks minimum for a stream
	public final static byte[] majick = {
			(byte) 0xd0,
			(byte) 0xcf,
			(byte) 0x11,
			(byte) 0xe0,
			(byte) 0xa1,
			(byte) 0xb1,
			(byte) 0x1a,
			(byte) 0xe1,
			(byte) 0x00,
			(byte) 0x00,
			(byte) 0x00,
			(byte) 0x00,
			(byte) 0x00,
			(byte) 0x00,
			(byte) 0x00,
			(byte) 0x00
	};

	/**
	 * initialize from existing data
	 */
	protected boolean init()
	{
		return this.init( this.data );
	}

	/**
	 * structure of typical header:
	 * // 24, 25= 62, 0  (rev #) = 0x3E
	 * // 26, 27= 3, 0   (vers #) = 0x03
	 * // 28, 29= -2, -1 (little endian)
	 * // 30, 31= 9, 0   (size of sector (9= 512, 7= 128))
	 * // 32, 33= 6, 0   (size of short sector (6=64 bytes))
	 * // 34->43= 0      (unused)
	 * // 44->47= Total # sectors in SAT ***
	 * // 48->51= Sector id of 1st sector
	 * // 52->55= 0      (unused)
	 * // 56->59= Minimum size of a standard stream (usually 4096)
	 * // 60->63= SecID of first sector in short-SAT
	 * // 64->67= Total # sectors in Short-SAT ***
	 * // 68->71= SecID of 1st sector of MSAT or -2 if no additional sectors are needed
	 * // 72->75= Total # sectors used for the MSAT
	 * // 76->436= MSAT (1st 109 sectors)
	 */
	private static void displayHeader( ByteBuffer data )
	{
		for( int i = 0; i < data.capacity(); i++ )
		{
			System.out.println( "[" + i + "] " + data.get( i ) );
		}
	}

	/**
	 * initialize the leo doc header
	 */
	protected boolean init( ByteBuffer dta )
	{
		this.data = dta;
		int pos;

		if( dta.limit() < 1 )
		{
			throw new InvalidFileException( "Empty input file." );
		}

		try
		{
			// sanity check -- is it a valid WorkBook?
			for( int t = 0; t < 16; t++ )
			{
				if( dta.get( t ) != majick[t] )
				{
					throw new InvalidFileException( "File is not valid OLE format: Bad Majick." );
				}
			}
		}
		catch( Exception e )
		{
			throw new InvalidFileException( "File is not valid OLE format:" + dta.limit() );
		}

		// get the rootstart 0x30
		pos = 0x30;
		data.position( pos );    // secId of 1st storage pos (==rootstart)[48]
		rootstart = data.getInt();
		rootstart++;
		rootstart *= BIGBLOCK.SIZE;

		// has DIFAT Extra sectors? -- extended data blocks for large XLS files
		pos = 0x44;    // [68]
		data.position( pos );
		extraDIFATStart = data.getInt();    //  sector id of 1st sector of extra DIFAT sectors or -2 of no extra sectors used
		numExtraDIFATSectors = data.getInt();    // total # sectors used for the DIFAT -- if # sectors or blocks > 109, more sectors are used

		// get minimum size of standard stream
		pos = 56;
		data.position( pos );
		minStreamSize = data.getInt();

		// starting at 0x2c (44) THE numFATSectors
		pos = 0x2C;
		data.position( pos );
		numFATSectors = data.getInt();    // total # blocks used in the FAT chain (each block can hold 128 sector ids)

		// 0x3c (60)  THE miniFAT start
		pos = 60;
		data.position( pos );
		miniFATStart = data.getInt();    // secId of 1st sector of the miniFAT sector chain or -2 if non-existant

		// pos 0x40= # miniFAT sectors [64]
		pos = 0x40;
		data.position( pos );
		numMiniFATSectors = data.getInt();
		return true;
	}

	/**
	 * get the number of the Extended DIFAT sectors
	 */
	public int getNumExtraDIFATSectors()
	{
		return numExtraDIFATSectors;
	}

	/**
	 * get the position of the extra DIFAT sector start
	 */
	public int getExtraDIFATStart()
	{
		return this.extraDIFATStart;
	}

	/**
	 * get the position of the Root Storage/1st Directory
	 */
	public int getRootStartPos()
	{
		return rootstart;
	}

	/**
	 * get the position of the miniFAT Start Block
	 */
	public int getMiniFATStart()
	{
		return miniFATStart;
	}

	/**
	 * get the FAT Sectors (sectors which hold the FAT or the chain that references all sectors in the file)
	 */
	public int[] getDIFAT()
	{
		int numblks = Math.min( this.numFATSectors, 109 ); // more than 109 goes into DIFAT
		int[] FAT = new int[numblks];
		int pos = DIFATPOSITION;    // START OF DIFAT (indexes the FAT)
		data.position( pos );    // start of the 1st 109 secIds (4 bytes each==436 bytes)
		for( int i = 0; i < numblks; i++ )
		{
			FAT[i] = data.getInt();
			FAT[i]++;
		}
		return FAT;
	}

	/**
	 * get the Header bytes
	 */
	public ByteBuffer getBytes()
	{
		return this.data;
	}

	/**
	 * get a BIGBLOCK containing this Header's data
	 * ie: create a new header byte block.
	 */
	static final LEOHeader getPrototype( int[] FAT )
	{

		LEOHeader retval = new LEOHeader();

		// get an empty BIGBLOCK
		BIGBLOCK retblock = (BIGBLOCK) BlockFactory.getPrototypeBlock( Block.BIG );
		retval.setData( retblock.data );
		retval.initMajickBytes();
		retval.initConstants();

		// recreate the FAT and indexes to ... 
		retval.setFAT( FAT );
		retval.setMiniFATStart( -2 );    // we don't rebuild miniFATs at this time ...
		retval.setRootStorageStart( 1 );
		retval.setExtraDIFATStart( -2 );
		retval.setNumExtraDIFATSectors( 0 );
		return retval;
	}

	void initConstants()
	{
		int pos = 0x08;
		data.position( pos );
		data.putInt( 0 );

		pos = 0xc;
		data.position( pos );
		data.putInt( 0 );

		pos = 0x10;
		data.position( pos );
		data.putInt( 0 );

		pos = 0x14;
		data.position( pos );
		data.putInt( 0 );

		pos = 0x18;
		data.position( pos );
		data.putShort( (short) 0x3e );

		pos = 0x1a;
		data.position( pos );
		data.putShort( (short) 0x3 );

		pos = 0x1c;
		data.position( pos );
		data.putShort( (short) -2 );

		pos = 0x1e;
		data.position( pos );
		data.putShort( (short) 9 );

		pos = 0x20;
		data.position( pos );
		data.putInt( 6 );

		pos = 0x24;
		data.position( pos );
		data.putInt( 0 );

		pos = 0x28;
		data.position( pos );
		data.putInt( 0 );

		pos = 0x34;
		data.position( pos );
		data.putInt( 0 );

		pos = 0x38;
		data.position( pos );
		data.putInt( 0x00001000 );

	}

	void initMajickBytes()
	{
		// create the majick number
		this.data.position( 0 );
		data.put( majick );
	}

	void setData( ByteBuffer dta )
	{
		this.data = dta;
	}

	/**
	 * Set the FAT or sector chain index
	 *
	 * @param FAT
	 */
	void setFAT( int[] FAT )
	{
		int pos = DIFATPOSITION;
		data.position( pos );
		for( int aFAT : FAT )
		{
			if( (FAT.length * 4) >= (this.data.limit() - 4) )
			{
				// TODO: Create extra DIFAT sectors FAT.length too big...
				log.error( "WARNING: LEOHeader.setFAT() creating Extra FAT Sectors Not Implemented.  Output file too large." );
			}
			else
			{
				data.putInt( (aFAT - 1) );    // todo: why decrement here?  necessitates an increment in LEOFile.writeBytes
			}
		}
		// fill rest with -1's === Empty Sector denotation		
		for( int i = data.position(); i < 512; i += 4 )
		{
			data.putInt( -1 );
		}

		setNumFATSectors( FAT.length );
	}

	void setRootStorageStart( int i )
	{
		int pos = 0x30;
		data.position( pos );
		data.putInt( i );
	}

	void setMiniFATStart( int i )
	{
		int pos = 0x3c;
		this.miniFATStart = i;
		data.position( pos );
		data.putInt( i );
		miniFATStart = i;
	}

	void setExtraDIFATStart( int i )
	{
		int pos = 0x44;
		data.position( pos );
		data.putInt( i );
		extraDIFATStart = i;
	}

	void setNumExtraDIFATSectors( int i )
	{
		int pos = 0x48;
		data.position( pos );
		data.putInt( i );
		numExtraDIFATSectors = i;
	}

	void setNumFATSectors( int i )
	{
		int pos = 0x2c;
		data.position( pos );
		data.putInt( i );
		numFATSectors = i;
	}

	int getNumFATSectors()
	{
		return numFATSectors;
	}

	/**
	 * set the number of miniFAT sectors (each sector=64 bytes)
	 */
	void setNumMiniFATSectors( int i )
	{
		int pos = 0x40;
		data.position( pos );
		data.putInt( i );
		numMiniFATSectors = i;
	}

	/**
	 * return the number of miniFAT sectors
	 *
	 * @return
	 */
	int getNumMiniFATSectors()
	{
		return numMiniFATSectors;
	}

	/**
	 * return the minimum stream size
	 * usually 4096 or 8 blocks
	 *
	 * @return
	 */
	int getMinStreamSize()
	{
		return minStreamSize;
	}
}