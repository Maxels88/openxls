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

import com.extentech.formats.XLS.BiffRec;
import com.extentech.formats.XLS.WorkBookException;
import com.extentech.formats.XLS.XLSRecord;
import com.extentech.formats.XLS.XLSRecordFactory;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Serializable;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.Hashtable;
import java.util.Iterator;
import java.util.List;

/**
 * The directory system for an LEO file
 * <p/>
 * <a href = "http://www.extentech.com">Extentech Inc.</a>
 */
public class StorageTable implements Serializable
{
	private static final Logger log = LoggerFactory.getLogger( StorageTable.class );
	private static final long serialVersionUID = 3399830613453524580L;
	public final static int TABLE_SIZE = 0x200;
	public final static int DIRECTORY_SIZE = 0x80;
	public final static int BIGSTORAGE_SIZE = 4096;        // default, should read from LEOHeader

	private LEOHeader myheader;

	CompatibleVector directoryVector = new CompatibleVector();    // all directories in the LEO file

	// Directory collection
	Hashtable directoryHashtable = new Hashtable( 100, 0.9f );

	private int dupct = 0;

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		myheader = null;

		Iterator ii = directoryHashtable.keySet().iterator();
		while( ii.hasNext() )
		{
			Storage s = (Storage) directoryHashtable.get( ii.next() );
			s.close();
		}
		directoryHashtable = new Hashtable( 100, 0.9f );

		for( Object aDirectoryVector : directoryVector )
		{
			Storage s = (Storage) aDirectoryVector;
			s.close();
		}
		directoryVector.clear();
	}

	/**
	 * initialize the directory entry array
	 */
	public void init( ByteBuffer dta, LEOHeader h, List blockvect, int[] FAT )
	{
		myheader = h;
		byte[] data = LEOFile.getBytes( initDirectoryStream( blockvect, FAT ) );
		int psbsize = data.length;
		int numRecs = psbsize / DIRECTORY_SIZE;
			log.debug( "Number of Directories: " + numRecs );
			log.debug( "Directories:  {}", Arrays.toString( data ) );

		int pos = 0;
		for( int i = 0; i < numRecs; i++ )
		{
			ByteBuffer b = ByteBuffer.allocate( DIRECTORY_SIZE );
			b.order( ByteOrder.LITTLE_ENDIAN );
			b.put( data, pos, DIRECTORY_SIZE );
			pos += DIRECTORY_SIZE;
			Storage rec = null;
			try
			{
				rec = new Storage( b );
			}
			catch( Exception ex )
			{
				throw new WorkBookException( "StorageTable.init failed:" + ex.toString(), WorkBookException.UNSPECIFIED_INIT_ERROR );
			}
			if( i == 0 )
			{
				rec = new RootStorage( b );
				if( !rec.getName().equals( "Root Entry" ) )
				{
					rec.setName( "Root Entry" );    // can happen upon a mac-sourced file
				}
			}
			addStorage( rec, -1 );
		}
	}

	/**
	 * return the number of existing directories
	 *
	 * @return
	 */
	public int getNumDirectories()
	{
		return directoryVector.size();
	}

	/**
	 * Create a new Storage and add to the directory array
	 * <br>NOTE: add any associated data separately
	 *
	 * @param name      Storage name - must be unique
	 * @param type      Storage Type - 1= storage 2= stream 5= root 0=unknown or unallocated
	 * @param insertIdx id where to insert, or -1 to insert at end
	 * @return
	 */
	public Storage createStorage( String name, int type, int insertIdx )
	{
		Storage s = null;
		try
		{
			ByteBuffer b = ByteBuffer.allocate( DIRECTORY_SIZE );
			b.order( ByteOrder.LITTLE_ENDIAN );
			b.put( new byte[DIRECTORY_SIZE] );
			s = new Storage( b );
			s.setName( name );
			s.setStorageType( type );
			s.setBlocks( new Block[0] ); // init empty storages with 0 blocksW
			s.setPrevStorageID( -1 );
			s.setNextStorageID( -1 );
			s.setChildStorageID( -1 );
			addStorage( s, insertIdx );
		}
		catch( Exception ex )
		{
			throw new WorkBookException( "Storage.createStorage failed:" + ex.toString(), WorkBookException.UNSPECIFIED_INIT_ERROR );
		}
		return s;
	}

	/**
	 * init the directories and gather their associated data (if present)
	 * including those directories whose data is gathered from the miniStream
	 */
	int[] miniFAT = null;    // index into the miniStream

	public void initDirectories( List blockvect, int[] FAT )
	{
		Enumeration rcs = directoryVector.elements();
		List miniStream = null;
		int totrecsize = 0;

		while( rcs.hasMoreElements() )
		{
			Storage rec = (Storage) rcs.nextElement();
			String name = rec.getName();
			int recsize = rec.getActualFileSize();
//			if (LEOFile.DEBUG)
//				Logger.logInfo("Initializing Directory: "	+ name + ".  Start=" + rec.getStartBlock() + ". Size=" + recsize);
			totrecsize += recsize;
			if( name.equals( "Root Entry" ) )
			{
				// also sets miniFAT ... ugly, I know ...
				miniStream = initMiniStream( blockvect,
				                                  FAT );    // grab the mini stream (short sector container) (if any), indexed by miniFAT
				if( (miniFAT != null) )
				{
					log.debug( "miniFAT: {} ", Arrays.toString( miniFAT ) );
				}
			}
			// this storage has it's data in the miniStream
			else if( (recsize > 0) && (recsize < BIGSTORAGE_SIZE) )
			{
				rec.setBlockType( Block.SMALL );
				initStorage( rec, miniStream, miniFAT, Block.SMALL );    // the miniStream is indexed by the miniFAT
				// Regular Sector file storage
			}
			else if( (recsize >= BIGSTORAGE_SIZE) )
			{
				initStorage( rec, blockvect, FAT, Block.BIG );

				// a storage-less directory
			}
			else if( recsize == 0 )
			{
				rec.setBlocks( new Block[0] ); // init empty storages with 0 blocksW
			}
			else
			{
					log.warn( "Storage has no Block Type." );
			}
		}
			log.debug( "Total Size used by Directories : " + String.valueOf( totrecsize ) );
	}

	/**
	 * // whenever a stream is shorter than a specific length
	 * // it is stored as a short-stream or ministream.  ministreamentries do not directly
	 * // use sectors to store their data, but are all embedded in a specific
	 * // internal control stream.
	 * // first used sector is obtained from the root store ==> miniFAT chain
	 * // it's secID chain is contained in the miniFAT
	 * // The data used by all of the short-sectors container stream are concatenated
	 * // in order of the secID chain.  There is no header, so the first
	 * // mini sector (secId= 0) is always located at position 0 in the mini Stream
	 * // The miniFAT is the same as the FAT except the secID chains refer
	 * // to miniSectors (64 bytes) rather than regular sectors or blocks (512 bytes)
	 */
	public List initMiniStream( List blockvect, int[] FAT )
	{
		int pos = myheader.getMiniFATStart();
		if( pos == -2 )
		{
			return null; // no miniStream sectors
		}

		List miniFATSectors = getMiniFAT( pos, blockvect, FAT );
		if( miniFATSectors.size() > 0 )
		{
			try
			{
				Iterator sbz = miniFATSectors.iterator();
				while( sbz.hasNext() )
				{
					BIGBLOCK sbb = (BIGBLOCK) sbz.next();
					sbb.setIsDepotBlock( true );
				}

				miniFAT = LEOFile.readFAT( miniFATSectors );

				RootStorage rootStore;
				try
				{
					rootStore = (RootStorage) getDirectoryByName( "Root Entry" );
				}
				catch( StorageNotFoundException e )
				{
					throw new InvalidFileException( "Error parsing OLE File. OLE FileSystem Out of Spec: No Root Entry." );
				}
				// capture the short-stream container stream
				int miniStreamStart = rootStore.getStartBlock();
				int miniStreamSize = rootStore.getActualFileSize();

				Storage miniStream = new Storage();
				miniStream.setStartBlock( miniStreamStart );
				miniStream.setActualFileSize( miniStreamSize );
				miniStream.setName( "miniStream" );
				// obtain the miniStream from the regular bigblock store
				initStorage( miniStream, blockvect, FAT, Block.BIG );
				// now that we have the entire miniStream , break it up into mini Sector-sized blocks
				// NOTE: only miniStreamSize bytes are usable - ignore rest
				byte[] b = new byte[miniStreamSize];
				System.arraycopy( miniStream.getBytes(), 0, b, 0, miniStreamSize );
				Block[] miniStreamBlocks = BlockFactory.getBlocksFromByteArray( b, Block.SMALL );
				ArrayList miniStreamBlockList = new ArrayList();
				for( Block miniStreamBlock : miniStreamBlocks )
				{    // should equal sbbsize
					miniStreamBlockList.add( miniStreamBlock );
				}
				return miniStreamBlockList;
			}
			catch( LEOIndexingException e )
			{
				log.error( "initSBStorages: Error obtaining sbdIdx" );
			}
		}
		return null;
	}

	/**
	 * extract the miniFAT sector index from the blockvect
	 *
	 * @param pos
	 * @param blockvect
	 * @param FAT
	 * @return
	 */
	private static List getMiniFAT( int pos, List blockvect, int[] FAT )
	{
		Storage miniFATContainer = null;
		miniFATContainer = new Storage();
		miniFATContainer.setBlockType( Block.BIG );
		miniFATContainer.setStartBlock( pos );
		miniFATContainer.setStorageType( 5 ); // set as root type to distinguish from regular storages
			log.debug( "StorageTable.getMiniFAT() Initializing miniFAT Container." );
		miniFATContainer.init( blockvect, FAT, true );
		//	miniFAT.setName("SBidx");
		// miniFAT index
		return miniFATContainer.getBlockVect();
	}

	/**
	 * add a Storage to the directory array
	 *
	 * @param Storage   storage to insert
	 * @param insertIdx -1 if add at end, otherwise insert at spot
	 */
	void addStorage( Storage rec, int insertIdx )
	{
		String nm = rec.getName();
		if( (directoryHashtable.get( nm ) != null) && !nm.equals( "" ) )
		{
/*	KSC: with 2012 code changes, this breaks output:
 * 		if (LEOFile.DEBUG)
				Logger.logInfo(
					"INFO: StorageTable.addStorage() Dupe Storage Name: " + nm);
			nm = nm + "|^" + dupct++;
			rec.setName(nm);
 Does not appear necessary			
*/
		}
		directoryHashtable.put( rec.getName(), rec );
		if( insertIdx == -1 )
		{
			directoryVector.add( rec );
		}
		else
		{
			directoryVector.add( insertIdx, rec );
			for( Object aDirectoryVector : directoryVector )
			{    // adjust prev, next, child ids if necessary
				Storage s = (Storage) aDirectoryVector;
				if( s.getChildStorageID() >= insertIdx )
				{
					s.setChildStorageID( s.getChildStorageID() + 1 );
				}
				if( s.getPrevStorageID() >= insertIdx )
				{
					s.setPrevStorageID( s.getPrevStorageID() + 1 );
				}
				if( s.getNextStorageID() >= insertIdx )
				{
					s.setNextStorageID( s.getNextStorageID() + 1 );
				}

			}
		}
	}

	/* remove a Storage
	*/
	void removeStorage( Storage st )
	{
		directoryHashtable.remove( st );
		directoryVector.remove( st );
	}

	/**
	 * init a Storage
	 */
	static void initStorage( Storage rec, List sourceblocks, int[] idx, int blocktype )
	{
		String name = rec.getName();
		int recsize = rec.getActualFileSize();
			log.debug( "Initializing Storage: " + name + " Retrieving Data." +
					                " Size: " + String.valueOf( recsize ) + " type: " + String.valueOf( blocktype ) + " startidx: " + rec.getStartBlock() + ((rec
					.getBlockType() == Block.SMALL) ? " MiniFAT" : "") );
		rec.setBlockType( blocktype );
		if( ("Root Entry").equals( name ) ) // ksc: shouldn't!
		{
			return;
		}

		if( rec.getBlockType() == Block.BIG )
		{
			rec.init( sourceblocks, idx, false );
		}
		else if( rec.getBlockType() == Block.SMALL )
		{
			rec.initFromMiniStream( sourceblocks, idx );
		}

		rec.idxs = null;
	}

	/**
	 * get the directory BLOCKS or Sectors
	 * contains the header info for all of the directories
	 */
	List initDirectoryStream( List blockvect, int[] FAT )
	{
		//getDirectoryBlocks()
		Storage directories = null;
		directories = new Storage();
		int pstart = (myheader.getRootStartPos() / BIGBLOCK.SIZE) - 1;
		directories.setStartBlock( pstart );
		directories.setBlockType( Block.BIG );
		directories.setStorageType( 5 ); // set to root directory
		directories.init( blockvect, FAT, true ); // get additional directory stores, if any
//		directories.setName("StorageTable");
//		if (LEOFile.DEBUG)
//			StorageTable.writeitout(directories.getBlockVect(),
//									"directoryStorage.dat");
		return directories.getBlockVect();

	}

	/**
	 * generate new RootStorage bytes from
	 * all of the Storage directories
	 */
	byte[] rebuildRootStore()
	{
		/*
		 * Free (unused) directory entries are marked with Object Type 0x0 (unknown or unallocated). 
		 * The entire directory entry should consist of all zeroes except for the child, right sibling, 
		 * and left sibling pointers, which should be initialized to NOSTREAM (0xFFFFFFFF).
		 */
		while( (directoryVector.size() % 4) != 0 )
		{ // add "null" storages to ensure multiples of 4 (128*4=512==minimum size)
			createStorage( "", 0, -1 );
		}
		Enumeration e = directoryVector.elements();
		byte[] bytebuff = new byte[directoryVector.size() * DIRECTORY_SIZE];
		int pos = 0;
		while( e.hasMoreElements() )
		{
			Storage s = (Storage) e.nextElement();
			ByteBuffer buff = getDirectoryHeaderBytes( s );
			System.arraycopy( buff.array(), 0, bytebuff, pos, DIRECTORY_SIZE );
			pos += DIRECTORY_SIZE;
		}
		return bytebuff;
	}

	/**
	 * get the Hashtable of Storages within the LEO file
	 */
	Hashtable getStorageHash()
	{
		return directoryHashtable;
	}

	/**
	 * get the CompatibleVector of all Directories within the LEO file
	 */
	CompatibleVector getAllDirectories()
	{
		return directoryVector;
	}

	/**
	 * get the Root Storage Record
	 */
	RootStorage getRootStorage() throws StorageNotFoundException
	{
		return (RootStorage) getDirectoryByName( "Root Entry" );
	}

	/**
	 * get the directory by name.  throws StorageNotFoundException if not found.
	 */
	public Storage getDirectoryByName( String name ) throws StorageNotFoundException
	{
		if( directoryHashtable.get( name ) != null )
		{
			return (Storage) directoryHashtable.get( name );
		}
		throw new StorageNotFoundException( "Storage: " + name + " not located" );
	}

	/**
	 * returns the Stream ID of the named directory, or -1 if not found
	 *
	 * @param name
	 * @return
	 * @throws StorageNotFoundException
	 */
	public int getDirectoryStreamID( String name ) throws StorageNotFoundException
	{
		for( int i = 0; i < directoryVector.size(); i++ )
		{
			Storage s = (Storage) directoryVector.get( i );
			if( s.getName().equals( name ) )
			{
				return i;
			}
		}
		return -1;
	}

	/**
	 * get the child directory (storage or stream) of the named storage
	 *
	 * @param name
	 * @return Storage or null
	 * @throws StorageNotFoundException
	 */
	public Storage getChild( String name ) throws StorageNotFoundException
	{
		Storage s = getDirectoryByName( name );
		if( s != null )
		{
			int child = s.getChildStorageID();
			if( child > -1 )
			{
				return (Storage) directoryVector.get( child );
			}
		}
		return null;
	}

	/**
	 * get the next directory (storage or stream) of the named storage
	 *
	 * @param name
	 * @return Storage or null
	 * @throws StorageNotFoundException
	 */
	public Storage getNext( String name ) throws StorageNotFoundException
	{
		Storage s = getDirectoryByName( name );
		if( s != null )
		{
			int next = s.getNextStorageID();
			if( next > -1 )
			{
				return (Storage) directoryVector.get( next );
			}
		}
		return null;
	}

	/**
	 * get the previous directory (storage or stream) of the named storage
	 *
	 * @param name
	 * @return Storage or null
	 * @throws StorageNotFoundException
	 */
	public Storage getPrevious( String name ) throws StorageNotFoundException
	{
		Storage s = getDirectoryByName( name );
		if( s != null )
		{
			int prev = s.getNextStorageID();
			if( prev > -1 )
			{
				return (Storage) directoryVector.get( prev );
			}
		}
		return null;
	}

	/**
	 * get the header record for a Directory
	 */
	static ByteBuffer getDirectoryHeaderBytes( Storage thisStorage )
	{
		// create a new byte buffer
		ByteBuffer buff = null;
		buff = thisStorage.getHeaderData();
		// 20100304 KSC: don't reset anything, just return
		return buff;
	}

	/**
	 * write out the directory bytes for debugging
	 */
	public final static void writeitout( List blocks, String name )
	{
		try
		{
			FileOutputStream fos = new FileOutputStream( new File( System.getProperty( "user.dir" ) + "\\storages\\" + name ) );
			fos.write( LEOFile.getBytes( blocks ) );
		}
		catch( IOException a )
		{
			;
		}
	}

	public void DEBUG()
	{
		log.info( "DIRECTORY CONTENTS:" );
		for( Object aDirectoryVector : directoryVector )
		{
			Storage s = (Storage) aDirectoryVector;
			String n = s.getName();
			log.info( "Storage: " + n + " storageType: " + s.getStorageType() + " directoryColor:" + s.getDirectoryColor() +
					                " prevSID:" + s.getPrevStorageID() + " nextSID:" + s.getNextStorageID() + " childSID:" + s.getChildStorageID() + " sz:" + s
					.getActualFileSize() );
			// special storages
			if( n.equals( "Root Entry" ) )
			{
				log.info( "Root Header: " + Arrays.toString( s.getHeaderData().array() ) );

				/***********************************
				 // KSC: TESTING for XLS-97:
				 !!!  Set creation time and modified time on root storage to 0
				 int p = 100;
				 s.getHeaderData().position(p);
				 Long tsCreated= s.getHeaderData().getLong();
				 Long tsModified= s.getHeaderData().getLong();
				 */
				if( s.myblocks != null )
				{
					int zz = 0;
					if( (s.myblocks.get( zz ) instanceof BIGBLOCK) )
					{
						log.debug( "BLOCK 1:\t" + zz + "-" + Arrays.toString( ((BIGBLOCK) s.myblocks.get( zz )).getBytes() ) );
					}
					else
					{
						log.debug( "BLOCK 1:\t" + zz + "-" + Arrays.toString( ((SMALLBLOCK) s.myblocks.get( zz )).getBytes() ) );
					}
				}
			}
			else if( n.equals( "Workbook" ) )
			{
				; //skip
			}
			else if( n.equals( "\1CompObj" ) )
			{
				BlockByteReader bytes = s.getBlockReader();
				int len = bytes.getLength();
				BiffRec rec = new XLSRecord();        // 4 bytes are header ...
				rec.setByteReader( bytes );
				rec.setLength( len );
				int slen = ByteTools.readInt( rec.getBytesAt( 24, 4 ) );    // actually position 28
				if( slen >= 0 )
				{
					String ss = new String( rec.getBytesAt( 28,
					                                        slen ) );        // AnsiUserType= a display name of the linked object or embedded object.
					log.debug( "\tOLE Object:" + ss );
				}
				// AnsiClipboardFormat (variable)
//				System.out.println("\t" + Arrays.toString(rec.getData()));				
			}
			else if( n.startsWith( "000" ) )
			{        // pivot cache
				if( s.myblocks != null )
				{
					for( int zz = 0; zz < s.myblocks.size(); zz++ )
					{
						if( (s.myblocks.get( zz ) instanceof BIGBLOCK) )
						{
							log.debug( "\t" + zz + "-" + Arrays.toString( ((BIGBLOCK) s.myblocks.get( zz )).getBytes() ) );
						}
						else
						{
							log.debug( "\t" + zz + "-" + Arrays.toString( ((SMALLBLOCK) s.myblocks.get( zz )).getBytes() ) );
						}
					}
				}
				BlockByteReader bytes = s.getBlockReader();
				int len = bytes.getLength();
				for( int z = 0; z <= (len - 4); )
				{
					byte[] headerbytes = bytes.getHeaderBytes( z );
					short opcode = ByteTools.readShort( headerbytes[0], headerbytes[1] );
					int reclen = ByteTools.readShort( headerbytes[2], headerbytes[3] );
					BiffRec rec = XLSRecordFactory.getBiffRecord( opcode );
					rec.setByteReader( bytes );
					rec.setOffset( z );
					rec.setLength( (short) reclen );
					rec.init();
					log.debug( "\t\t" + rec.toString() );
					z += reclen + 4;
				}
			}
			else
			{
				if( s.myblocks != null )
				{
					for( int zz = 0; zz < s.myblocks.size(); zz++ )
					{
						if( (s.myblocks.get( zz ) instanceof BIGBLOCK) )
						{
							log.debug( "\t" + zz + "-" + Arrays.toString( ((BIGBLOCK) s.myblocks.get( zz )).getBytes() ) );
						}
						else
						{
							log.debug( "\t" + zz + "-" + Arrays.toString( ((SMALLBLOCK) s.myblocks.get( zz )).getBytes() ) );
						}
					}
				}
			}
		}
	}
}