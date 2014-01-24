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
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.JFileWriter;
import com.extentech.toolkit.ResourceLoader;
import com.extentech.toolkit.TempFileManager;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Serializable;
import java.nio.ByteBuffer;
import java.nio.ByteOrder;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;

/**
 * LEOFile is an archive format compatible with other popular archive formats such as OLE.
 * <p/>
 * It contains multiple files or "Storages" which can
 * subsequently contain data in popular desktop application formats.
 */
public class LEOFile implements Serializable
{
	private static final Logger log = LoggerFactory.getLogger( LEOFile.class );
	private static final long serialVersionUID = 2760792940329331096L;
	public final static int MAXDIFATLEN = 109;    // maximum DIFATLEN = 109; if more sectors are needed goes into extraDIFAT
	public final static int IDXBLOCKSIZE = 128;    // number of indexes that can be stored in 1 block
	public static int actualOutput = 0;
	private List bigBlocks;
	private boolean readok = false;
	private LEOHeader header = null;
	private StorageTable directories;
	private FileBuffer fb = null;
	String fileName = "New Spreadsheet";
	byte[] encryptionStorageOverage = null;
	boolean encryptedXLSX = false;

	// TODO: fix temp file issues and implement shutdown cleanup -- currently no way to delete -jm

	public LEOHeader getHeader()
	{
		return header;
	}

	/**
	 * Close the underlying IO streams for this LEOFile.
	 * <p/>
	 * Also deletes the temp file.
	 *
	 * @throws Exception
	 */
	public void close() throws IOException
	{
		if( fb != null )
		{
			fb.close();
		}
		// KSC: close out other object refs
		fb = null;
		//header.getBytes().clear();
		header = null;//new LEOHeader();
		if( directories != null )
		{
			directories.close();
			directories = null;
		}
		if( bigBlocks != null )
		{
			for( Object bigBlock : bigBlocks )
			{
				BlockImpl b = (BlockImpl) bigBlock;
				b.close();
			}
			bigBlocks.clear();
		}
		bigBlocks = null;
//	    FAT= null;
	}

	/**
	 * just closes the filebuffer withut clearing out buffers and storage tables
	 */
	public void closefb() throws IOException
	{
		if( fb != null )
		{
			fb.close();
		}
		fb = null;
	}

	public void shutdown()
	{
		try
		{
			close();
		}
		catch( Exception e )
		{
				log.warn( "could not close workbook cleanly.", e );
		}
	}

	/**
	 * a new LEO file containing LEO archive entries
	 *
	 * @param String a file path containing a valid LEOfile
	 */
	public LEOFile( String fname )
	{
		if( fname.indexOf( ".ser" ) > -1 )
		{
			initFromPrototype( fname );
			return;
		}
		fileName = fname;
		fb = LEOFile.readFile( fname );
		initWrapper( fb.getBuffer() );
	}

	/**
	 * Create a leo file from a prototype string/path
	 * <p/>
	 * PROTOTYPE_LEO_ENCRYPTED
	 *
	 * @param fname
	 */
	private void initFromPrototype( String fname )
	{
		try
		{
			byte[] b = ResourceLoader.getBytesFromJar( fname );
			if( b == null )
			{
				throw new com.extentech.formats.XLS.WorkBookException(
						"Required Class files not on the CLASSPATH.  Check location of .jar file and/or jarloc System property.",
						WorkBookException.LICENSING_FAILED );
			}
			ByteBuffer bbf = ByteBuffer.wrap( b );
			bbf.order( ByteOrder.LITTLE_ENDIAN );
			initWrapper( bbf );
		}
		catch( Exception e )
		{
			throw new InvalidFileException( "WorkBook could not be instantiated: " + e.toString() );
		}
	}

	/**
	 * Instantiate a Leo file from an encrypted document.  We currently have a hack
	 * in place for encrypted documents due to their having truncated bigblocks that our
	 * interface does not correctly handle.
	 *
	 * @param a       file containing a valid LEOfile (XLS BIFF8)
	 * @param whether to use a temp file
	 * @param if      the file is encrypted xlsx format
	 */
	public LEOFile( File fpath, boolean usetempfile, boolean encryptedXLSX )
	{
		this.encryptedXLSX = encryptedXLSX;
		fileName = fpath.getAbsolutePath();
		fb = LEOFile.readFile( fpath, usetempfile );
		initWrapper( fb.getBuffer() );
	}

	/**
	 * Checks whether the given byte array starts with the LEO magic number.
	 */
	public static boolean checkIsLEO( byte[] data, int count )
	{
		if( count < LEOHeader.majick.length )
		{
			return false;
		}
		for( int idx = 0; idx < LEOHeader.majick.length; idx++ )
		{
			if( data[idx] != LEOHeader.majick[idx] )
			{
				return false;
			}
		}
		return true;
	}

	/**
	 * a new LEO file containing LEO archive entries
	 *
	 * @param a       file containing a valid LEOfile (XLS BIFF8)
	 * @param whether to use a temp file
	 */
	public LEOFile( File fpath, boolean usetempfile )
	{
		fileName = fpath.getAbsolutePath();
		fb = LEOFile.readFile( fpath, usetempfile );
		initWrapper( fb.getBuffer() );
	}

	/**
	 * a new LEO file containing LEO archive entries
	 *
	 */
	public LEOFile( File fpath )
	{
		fileName = fpath.getAbsolutePath();
		fb = LEOFile.readFile( fpath );
		initWrapper( fb.getBuffer() );
	}

	/**
	 * a new LEO file containing LEO archive entries
	 *
	 * @param byte[] a byte array containing a valid LEOfile
	 */
	public LEOFile( ByteBuffer bytebuff )
	{
		initWrapper( bytebuff );
	}

	/**
	 * This is just removing some duplicate code from our constructors
	 * <p/>
	 * We should add some exception handling in here!
	 */
	public void initWrapper( ByteBuffer bytebuff )
	{
		int[] FAT = init( bytebuff );
		if( FAT != null )
		{
			directories.initDirectories( bigBlocks, FAT );
			FAT = null;
			readok = true;
		}
		else
		{
			readok = false;
		}
	}

	public void clearAfterInit()
	{
		bigBlocks.clear();
	}

	/**
	 * Reads in an encrypted LEO stream.  May be able to remove this and use the standard
	 * constructor, we shall see.
	 *
	 * @param encryptedFile
	 */
	public void readEncryptedFile( File encryptedFile )
	{
		fb = readFile( encryptedFile );
		initWrapper( fb.getBuffer() );
	}

	public String getFileName()
	{
		return fileName;
	}

	/**
	 * return whether the LEOFile contains a valid workbook
	 */
	public boolean hasWorkBook()
	{
		Storage book;
		try
		{
			book = directories.getDirectoryByName( "Workbook" );
		}
		catch( StorageNotFoundException e )
		{
			try
			{
				book = directories.getDirectoryByName( "Book" );
			}
			catch( StorageNotFoundException e1 )
			{
				return false;
			}
		}
		return true;
	}

	/**
	 * return whether the LEOFile contains a valid doc
	 */
	public boolean hasDoc()
	{
		if( readok )
		{
			try
			{
				Storage doc = directories.getDirectoryByName( "WordDocument" );
			}
			catch( StorageNotFoundException e )
			{
				return false;
			}
			return true;
		}
		return false;
	}

	/**
	 * return whether the LEOFile contains the _SX_DB_CUR Pivot Table Cache storage (required for Pivot Table)
	 *
	 * @return
	 */
	public boolean hasPivotCache()
	{
		if( readok )
		{
			try
			{
				Storage doc = directories.getDirectoryByName( "_SX_DB_CUR" );
			}
			catch( StorageNotFoundException e )
			{
				return false;
			}
			return true;
		}
		return false;
	}

	/**
	 *
	 */
	public LEOFile()
	{
		super();
	}

	/**
	 * Create a LEOFile from an input stream.  Unfortunately because of our byte backer
	 * we need to write to a temporary file in order to do this
	 *
	 * @param stream
	 */
	public LEOFile( InputStream stream ) throws IOException
	{
		File target = TempFileManager.createTempFile( "ExtenXLS_temp", ".leo" );
		JFileWriter.writeToFile( stream, target );

		fileName = target.getAbsolutePath();
		fb = LEOFile.readFile( target );
		initWrapper( fb.getBuffer() );
		target.deleteOnExit();
		target.delete();
	}

	/**
	 * calculate number of FAT blocks necessary to describe compound file
	 */
	final static int getNumFATSectors( int storageTotal )
	{
		int nFAT = storageTotal * 4;
		nFAT /= BIGBLOCK.SIZE;
		float realnum = ((float) (storageTotal * 4) / BIGBLOCK.SIZE);
		if( ((realnum - nFAT) > 0) || (nFAT > MAXDIFATLEN) )
		{
			nFAT++;
		}
		return nFAT;
	}

	/**
	 * Creates all the necessary headers and associated storages/structures that,
	 * along with the actual workbook data, makes up the document
	 * <p/>
	 * Basic Structure
	 * 1st sector or block= Header + 1st 109 sector chains (==FAT===array of sector numbers)
	 * Location of FAT==DIFAT
	 * Header denotes sector of "Root Directory" which in turn denotes sectors of all other directories
	 * (Workbook is but one of several directories)
	 * Each directory denotes the start block of any associated data into the FAT (sector chain index)
	 * (if directory is >0 an <4096 in size, it's associated data is stored in miniStorage
	 * <p/>
	 * <p/>
	 * <p/>
	 * <p/>
	 * <p/>
	 * <p/>
	 * <p/>
	 * <p/>
	 * Ok, so this is a little odd, it looks as if it collects all of the storages, and handles
	 * all of the preprocessing necessary to create FAT etc.
	 * returns a list of all the storages, presumably to be handled by ByteStreamer.
	 * <p/>
	 * <p/>
	 * A little more information:
	 * A list of storages is generated.  These storages do not include
	 * a) the workbook storage(!) which is meant to be the first record written after receiving the byte array
	 * b) individual miniFAT storages, these are all combined into one container storage that is not represented in the header
	 * c) the LeoHeader, which is written to the output stream already
	 * <p/>
	 * <p/>
	 * It seems as if this method needs several things
	 * 1) better header comments on its functions
	 * 2) break out logic to private methods to keep overall functions understandable.
	 * 3) Handle file formats other than XLS (ie doc, encrypted xls, ppt, etc)
	 * 4) be clear about its usage to outputStream.  Why only the 1 storage written, maybe none would be better?
	 * 5) probably throw an exception
	 * <p/>
	 * 6) WHAT IS WITH THE STORAGES THAT ARE NOT STORAGES? -- these are non-directory storages used to write out the bytes contained within in the correct order
	 * for the document.  These non-directory storages include the FAT and the minFAT blocks and miniStream
	 *
	 * @param outputstream to write out
	 * @param book         to write
	 * @param book         bytes size
	 * @return storages to write out in streamer
	 */
	public synchronized List writeBytes( OutputStream out, int workbook_byte_size )
	{
		/***** storages to be written - rebuild from saved storages and recreated workbook storage*/
		AbstractList storages = new Vector();
		// header variables
		int numFATSectors = 0;     //num FAT indexes
		int rootstart = 0;   // block position where root sector is located
		int numExtraDIFATSectors = 0;    // num extra sectors/blocks
		int extraDIFATStart = -2;   // extra sector start (-2 for none)
		int numMiniFATSectors = 0;      //num short sectors/miniFAT indexes
		int miniFATStart = -2; // start block position of short sector container blocks (-2 for none)
		int sbidxpos = -2;   // start position of short sector idx chain (-2 for none)
		int sbsz = 0;        // size of short sector container
		int[] DIFAT;            // holds the sector id chains for the 1st 109 sector index blocks (DIFAT)
		int[] extraDIFAT;       // holds the sector id chains for any sector index blocks (DIFAT) > 109
		int numblocks = 1; // header is 1st

		// workbook directory
		Storage book = null;
		boolean isEncrypted = false;
		try
		{
			book = directories.getDirectoryByName( "Workbook" );
		}
		catch( StorageNotFoundException e3 )
		{
			try
			{
				book = directories.getDirectoryByName( "EncryptedPackage" );
				isEncrypted = !true;
			}
			catch( StorageNotFoundException e1 )
			{
				//this is an error state, ideally we would be throwing a (non-runtime) exception?
				throw new InvalidFileException( "Input LEO file not valid for output" );
			}
		}

		// root directory
		RootStorage rootStore = null;
		try
		{
			rootStore = directories.getRootStorage();
		}
		catch( StorageNotFoundException e2 )
		{
			//this is an error state, ideally we would be throwing a (non-runtime) exception?
			throw new InvalidFileException( "Input LEO file not valid for output" );
		}

		// get existing directories + store all except workbook + root (which are rebuilt)
		Enumeration e = directories.getAllDirectories().elements();
		while( e.hasMoreElements() )
		{
			Storage thisStore = (Storage) e.nextElement();
			if( (!thisStore.equals( book )) && (!thisStore.equals( rootStore )) )
			{ // && thisStore!=encryptionInfo?
				storages.add( thisStore );
				if( thisStore.getBlockType() == Block.SMALL )
				{    // count number of miniStream blocks
					thisStore.miniStreamStorage = true;    // TODO: investigate why this setting gets lost!
					numMiniFATSectors += thisStore.getBlockVect().size();
				}
			}
		}

		// if have miniStream sectors, rebuild miniFAT index 
		// + convert mini blocks to one or more big blocks (=miniFAT container) 
		if( numMiniFATSectors > 0 )
		{
			Storage[] sbs = buildMiniFAT( storages,
			                              numMiniFATSectors );    // returns two non-directory storages:  miniStream Storage + miniFAT index
			storages.add( sbs[0] );
			numMiniFATSectors = sbs[1].getBlockVect().size();            // trap number of miniStream block indexes
			storages.add( sbs[1] );
		}

		// we need to count up the total block array
		// count number of blocks:
		// add blocks necessary for root storage
		numblocks += Math.ceil( directories.directoryVector.size() / 4 );    // each 4 directories==1 block (512 byte sector)
		// count Storage Blocks (except root and workbook)
		for( int t = 0; t < storages.size(); t++ )
		{
			Storage nstr = (Storage) storages.get( t );    // saved stores containt thier original blocks + padding block signifying end of sector/block
			if( !nstr.miniStreamStorage )
			{
				numblocks += nstr.getSizeInBlocks();
			}
		}

		// take current workbook byte size and calculate # bigblocks
		book.miniStreamStorage = false;    // if had any initially, now it's big blocks

		int workbook_idx_block_size = LEOFile.getSizeInBlocks( workbook_byte_size, BIGBLOCK.SIZE );
		workbook_idx_block_size = Math.max( workbook_idx_block_size, getMinBlocks() - 1 );// ensure minimum # blocks
		numblocks += workbook_idx_block_size + 1;// to account for end of block sector
		if( isEncrypted )
		{  // Encrypted workbooks are already set to correct size
			numblocks--;
			workbook_idx_block_size--;
		}
		// Given amount of blocks to write, calculate id chain necessary to describe
		numFATSectors = LEOFile.getNumFATSectors( numblocks );    // get # of secids necessary to describe blocks

		// now include the FAT itself
		// total number of Sectors(blocks)= number of Sectors + FAT chain block(s) 
		numFATSectors = getNumFATSectors( numFATSectors + numblocks );
		int[] FAT = new int[IDXBLOCKSIZE * numFATSectors];
		for( int i = 0; i < FAT.length; i++ )
		{
			FAT[i] = -1;  // create byte array of -1's 128 bytes - total number of sector ids
		}

		// allocate FAT array index 
		// important positions that will be recorded in the header or FAT/miniFAT
		DIFAT = new int[Math.min( numFATSectors, MAXDIFATLEN )];    // 1 up to the 1st 109 sector index blocks used to describe this file
		extraDIFAT = new int[Math.max( numFATSectors - MAXDIFATLEN, 0 )]; // if necessary, sector index blocks in excess of 109

		/************************ now start to lay out block positions of all directories' associated data ************************/
		/************************ note: order of directories in storages is order that they will be written out 
		 * 						  this order is indexed in the FAT via the block indexes used for each sector 
		 * 						  upon return, storages contains all the directories, their associated data, the FAT ...
		 * 						  in the calling method, the workbook records (==associated data) is written first because
		 * 						  the workbook storage is first in the FAT.  
		 Eventually, want to change this so all other storages and the FAT are written first, then
		 the calling method only needs to write the workbook records.
		 **************************************************************************************************************************/
		// first start block index with workbook blocks (they will be written out in ByteStreamer before all other blocks except the header)
		int blockpos = 0;    // initial sector position
		book.setStartBlock( blockpos );
		for( int t = 0; t < workbook_idx_block_size; t++ )
		{
			FAT[blockpos++] = blockpos;
		}
		FAT[blockpos++] = -2;    // end of sector flag
		if( !isEncrypted )
		{
			// this is a questionable line.  It is setting the storage length to that of the padded bigblock.
			// causes errors in encrypted files
			book.setActualFileSize( (workbook_idx_block_size + 1) * BIGBLOCK.SIZE );
				log.debug( "Workbook actual bytes: " + book.getActualFileSize() );
		}

		// now rest of "static" stores (summary, document summary, comp obj ...)
		for( int t = 0; t < storages.size(); t++ )
		{
			Storage nstr = (Storage) storages.get( t );
			if( !nstr.miniStreamStorage )
			{    // miniStream blocks are handled separately
				if( nstr.getName().equals( "miniStream" ) )
				{    // miniStream (short sector container)
					miniFATStart = blockpos;    // this goes in rootstore startblock to denote start of short sector container
					sbsz = nstr.getActualFileSize();        // actual size of short sector (miniFAT) container - stored in root storage
				}
				else if( nstr.getName().equals( "miniFAT" ) )
				{ // miniFAT (short sector) index
					sbidxpos = blockpos;
				}
				nstr.setStartBlock( -2 );        // default, not linked to any blocks
// 20100325 KSC: since original blocks are kept for non-workbook storage, 
//				and since size is not under our control, keep original size - otherwise can error 				
//				nstr.setActualFileSize(0);	// default, 0 bytes
				// record index of block position for each block of the storage
				Block[] blks = nstr.getBlocks();
				if( blks != null )
				{
					nstr.setStartBlock( blockpos );    // start of block chain
					for( int i = 0; i < (blks.length - 1); i++ )
					{    // -1 to trap last block with special end-of-block flag (-2)
						FAT[blockpos++] = blockpos;    // start of this store's blocks as indexed in block index
					}
					FAT[blockpos++] = -2;    // end of sector/block flag
//					nstr.setActualFileSize((blks.length)*512);	// set file size -- see above why we can't do this at this time
				}
			}
		}

		// after other storages, add root store:
		// must rebuild after info from above as root storage is built from from all other storages 
		rootStore.setStartBlock( miniFATStart );        // root store start points to miniFAT container, or -2 if none
		rootStore.setActualFileSize( sbsz );            // 0 or miniFAT container size
		// handle rootstorage blocks = data of all other directory storages
		rootStore.setBytes( directories.rebuildRootStore() );    // now rebuild rootstore from all other storages
		rootstart = blockpos;
		for( int i = 0; i < (rootStore.getBlockVect().size() - 1); i++ )
		{
			FAT[blockpos++] = blockpos;
		}
		FAT[blockpos++] = -2;    // end of sector/block flag
		// add rootstore to list of blocks writing out 
		storages.add( rootStore );

		// mark position of FAT itself within blocks (==DIFAT)
		// the 1st MAXDIFATLEN DIFAT indexes goes into header
		for( int i = 0; i < DIFAT.length; i++ )
		{
			DIFAT[i] = blockpos + 1;    // each pos is decremented in method below, so increment here
			FAT[blockpos++] = -3;    // mark loc of FAT with special value -3
		}

		// if any FAT sectors > MAX, goes into extra blocks
		if( numFATSectors > MAXDIFATLEN )
		{    // handle extra blocks - need extra blocks to store the indexes
			for( int n = MAXDIFATLEN; n < numFATSectors; n++ )
			{
				extraDIFAT[n - MAXDIFATLEN] = blockpos + 1;
				FAT[blockpos++] = -3;
			}
			for( int i = 0; i < (int) Math.ceil( (numFATSectors - MAXDIFATLEN) / (IDXBLOCKSIZE * 1.0) ); i++ )
			{
				FAT[blockpos++] = -4;    // flag for MSAT or XBB
			}
		}

		// now that all the blocks referenced by storages such as workbook and document summary, 
		// plus blocks referenced by the root storage, 
		// transform the blockindex to blocks & input a phantom or special store so can write out
		// have been indexed, store the index blocks themselves
		Storage idxstore = new Storage();
		idxstore.setName( "IDXStorage" );
		Block[] FATSectors = getIDXBlocks( FAT );    // convert each int to 4-bytes in sector/block format --> FAT Sectors
		idxstore.setBlocks( FATSectors );    //input the FAT Sectors to a non-directory store for later writing
		storages.add( idxstore );    // output FAT Sectors after all other blocks

		if( numFATSectors > MAXDIFATLEN )
		{
			// now build and add the extra sectors necessary to store the FAT 
			extraDIFATStart = blockpos - (int) Math.ceil( (numFATSectors - MAXDIFATLEN) / (IDXBLOCKSIZE * 1.0) );
			Storage xbbstore = buildExtraDIFAT( extraDIFAT, extraDIFATStart );    // create block(s) from the extraDIFAT
			numExtraDIFATSectors = xbbstore.getBlockVect().size();
			storages.add( xbbstore );
		}

		// now that blocks have been accounted for and indexed and storages and all their blocks
		// (except for workbook blocks) are stored, we can build and output the header block
		// build header block:

		header = LEOHeader.getPrototype( DIFAT ); // input secIdChain to header loc 76
		// start sector of short sector/miniFAT index chain or -2 if n/a
		header.setMiniFATStart( sbidxpos );
		// the block # where the storages begin
		header.setRootStorageStart( rootstart );
		// number of block indexes needed to describe 1st 109 sectors/blocks (= DIFAT) 
		header.setNumFATSectors( numFATSectors );
		// number of block indexes needed to describe short sectors/miniFAT sectors if any 		
		header.setNumMiniFATSectors( numMiniFATSectors );
		// number of block indexes necessary to describe extra sectors if any 
		header.setNumExtraDIFATSectors( numExtraDIFATSectors );
		// start of extra sector block index or -2 if n/a
		header.setExtraDIFATStart( extraDIFATStart );

		// setup the header info...
		if( !header.init() )
		{
			throw new RuntimeException( "LEO File Header Not Initialized" ); // invalid WorkBook File
		}

//		 return the storages
		return storages;
	}

	/**
	 * build Extra Storage Sector index (extra DIFAT)
	 *
	 * @param difatidx - indexes of blocks > 13952 (= 109 blocks described by DIFAT * 128 indexes in each block)
	 * @return
	 */
	private Storage buildExtraDIFAT( int[] difatidx, int xbbpos )
	{
//			handle Extra blocks beyond the 109*128 blocks described by the regular DIFAT (109 limit)
		if( difatidx != null )
		{
			ArrayList outblocks = new ArrayList();
			byte[] xbytes = new byte[difatidx.length * 4];
			// convert the int difatidx to bytes    
			for( int i = 0; i < difatidx.length; i++ )
			{
				byte[] idx = ByteTools.cLongToLEBytes( difatidx[i] - 1 );
				System.arraycopy( idx, 0, xbytes, i * 4, 4 );
			}
			int counter = 0;
			for( int i = 0; i < xbytes.length; )
			{
				ByteBuffer bl = BlockFactory.getPrototypeBlock( Block.BIG ).getByteBuffer();
				int len = xbytes.length - i;

				// handle eob
				if( len > (BIGBLOCK.SIZE - 4) )
				{
					len = BIGBLOCK.SIZE - 4;
				}
				bl.position( 0 );
				bl.put( xbytes, i, len );

				// handle eob
				counter++;
				if( len == (BIGBLOCK.SIZE - 4) )
				{
					bl.putInt( xbbpos + counter );
				}

				BIGBLOCK xbbBlock = new BIGBLOCK();
				xbbBlock.init( bl, 0, 0 );
				outblocks.add( xbbBlock );

				i += BIGBLOCK.SIZE - 4;
			}
			// create new storage so can be added to end
			Storage xbbstore = new Storage();
			xbbstore.setName( "XBBStore" );
			Block[] xblx = new Block[outblocks.size()];
			outblocks.toArray( xblx );
			xbbstore.setBlocks( xblx );
			return xbbstore;
		}
		return null;
	}

	/**
	 * Gather up all existing smallblocks from the existing storages
	 * and create contiguous bigblocks from it; since smallblocks are
	 * only 128 bytes long, 4 smallblocks fit into 1 bigblock
	 * while building the miniFAT Sectors,also build the miniFAT index
	 * which corresponds to it
	 *
	 * @param storages existing Storages (except workbook and root, which are handled separately)
	 * @param numsbs   total number of miniFAT blocks
	 * @return Storage[] - two storage containers: 1 contains smallblocks, 2 contains smallblock index
	 */
	private Storage[] buildMiniFAT( List storages, int numsbs )
	{
		// once more, now gathering the sbidx's (miniFAT index)		
		int[] sbidx = new int[(int) Math.ceil( numsbs / (IDXBLOCKSIZE * 1.0) ) * IDXBLOCKSIZE];
		byte[] smallblocks = new byte[0];    // sum up all short sectors, will be subsequently converted to big blocks
		for( int i = 0; i < sbidx.length; i++ )
		{
			sbidx[i] = -1;  // init sbidx index
		}
		int z = 0;
		for( Object storage : storages )
		{
			Storage thisStore = (Storage) storage;
			if( thisStore.getBlockType() == Block.SMALL )
			{
				//since these are static blocks we can copy original data
				thisStore.setStartBlock( z );    // set start position of miniFATs as may have changed
				// keep length as that hasn't changed
				for( int j = 0; j < (thisStore.getBlockVect().size() - 1); j++ )
				{
					sbidx[z++] = z;
				}
				sbidx[z++] = -2;    // end of sector
				smallblocks = ByteTools.append( thisStore.getBytes(), smallblocks );
// 20100407 KSC: this screws up subsequent usages of book after writing ... see TestCreateNewName NPE 
				//thisStore.setBlocks(new Block[0]);
			}
		}

		// concatenate all miniFAT sectors to build the miniFAT or short sector container
		// i.e. miniFAT blocks are concatenated and stored in big block(s):
		Block[] smallBlocksToBig = BlockFactory.getBlocksFromByteArray( smallblocks, Block.BIG );
		Storage miniFATContainer = new Storage();
		miniFATContainer.setName( "miniStream" );
		miniFATContainer.setBlocks( smallBlocksToBig );
		miniFATContainer.setActualFileSize( smallblocks.length );
		smallblocks = null; // no need anymore
		Storage miniFAT = new Storage();
		miniFAT.setName( "miniFAT" );
		Block[] sbIDX = getIDXBlocks( sbidx );    // convert each int to 4-bytes in  sector/block format --> BBDIX
		miniFAT.setBlocks( sbIDX );
		return new Storage[]{ miniFATContainer, miniFAT };
	}

	/**
	 * create the idx from the actual locations
	 * of the Blocks in the outblock array
	 */
	final static void initSmallBlockIndex( int[] newidx, Block b )
	{
		while( b.hasNext() && !(b.equals( b.next() )) )
		{
			int origps = b.getBlockIndex();
			if( origps < 0 )
			{
				log.warn( "WARNING: LEOFile Block Not In MINIFAT vector: " + String.valueOf( b.getOriginalIdx() ) );
			}
			else
			{
				int newp = ((Block) b.next()).getBlockIndex();
				newidx[origps] = newp;
			}
			b = (Block) b.next();
		}
		if( b.getBlockIndex() >= 0 )
		{
			newidx[b.getBlockIndex()] = -2; // end of storage
		}
	}

	/** create the idx from the actual locations
	 of the Blocks in the outblock array
	 *
	 final static int initWBBlockIndex(int[] newidx, int sz) {
	 int start = 0;
	 for(int r=newidx.length-1;r>=0;r--) {
	 if(newidx[r]!=-1) {
	 start = r+1;
	 break;
	 }
	 }
	 for(int i=0;i<sz-1;i++) {
	 newidx[(start+i)] = start+(i+1);
	 }
	 if(start>0)
	 newidx[(start+sz-1)] = -2; // end of storage
	 else
	 newidx[(start+sz-1)] = -2; // end of storage
	 return start;
	 }
	 */

	/**
	 * returns an empty new FAT which
	 * accounts for the size of its own blocks
	 */
	final static int[] getEmptyDIFAT( int totblocks, int numFATSectors )
	{

		// allocate space in idx for FAT recs, StorageTable and for final -2
		int[] bbdi = new int[numFATSectors + totblocks + 1];
		for( int x = 0; x < bbdi.length; x++ )
		{
			bbdi[x] = (byte) -1;
		}
		return bbdi;
	}

	/**
	 * Create the array of BB Idx blocks
	 */
	final static Block[] getIDXBlocks( int[] bbdidx )
	{
		// step through and create index recs for all Blocks
		byte[] b = new byte[bbdidx.length * 4];
		int bv = 0;

		// can we wrap t in an nio buffered and read?

		for( int t = 0; t < b.length/* 20100304 KSC: why? - 4*/; )
		{
			byte[] bs = ByteTools.cLongToLEBytes( bbdidx[bv++] );
			// for(int z=0;z<4;z++)
			b[t++] = bs[0];
			b[t++] = bs[1];
			b[t++] = bs[2];
			b[t++] = bs[3];
		}
		return BlockFactory.getBlocksFromByteArray( b, Block.BIG );
	}

	/**
	 * get a handle to a Storage within this LEO file
	 */
	public Storage getStorageByName( String s ) throws StorageNotFoundException
	{
		return directories.getDirectoryByName( s );
	}

	/**
	 * get all Directories in this LEO file
	 */
	public Storage[] getAllDirectories()
	{
		CompatibleVector v = directories.getAllDirectories();
		Storage[] s = new Storage[v.size()];
		s = (Storage[]) v.toArray( s );
		return s;
	}

	/**
	 * return the Directory Array for this LEOFile
	 * <br>holds all directories (storages and streams) in correct order
	 * <br>Directories in turn can reference streams of data
	 *
	 * @return
	 */
	public StorageTable getDirectoryArray()
	{
		return directories;
	}

	/**  Simply setting the byte size on the book storage
	 *
	 private void checkIfSmallBlockOutput(int blen) throws StorageNotFoundException {
	 Storage book;
	 try {
	 book = storageTable.getStorageByName("Workbook");
	 }catch(StorageNotFoundException e) {
	 book = storageTable.getStorageByName("EncryptedPackage");
	 }

	 if (blen >= StorageTable.BIGSTORAGE_SIZE) {
	 if (book.getBlockType() == Block.SMALL) {
	 Logger.logWarn(
	 "WARNING: Modifying SmallBlock Workbook: "+ this.getFileName() +
	 ". Modifying SmallBlock-type Workbooks " +
	 " can cause corrupt output files. See: " +
	 "http://extentech.com/uimodules/docs/docs_detail.jsp?meme_id=195");
	 throw new WorkBookException("SmallBlock File Detected reading: "+this.getFileName(), WorkBookException.SMALLBLOCK_FILE);
	 // this.convertSBtoBB(book);
	 }
	 book.setActualFileSize(blen);
	 } else
	 book.setActualFileSize(StorageTable.BIGSTORAGE_SIZE); // set at the minimum
	 }
	 */

	/**
	 * get the bytes of the DOC file.
	 */
	public BlockByteReader getDocBlockBytes()
	{
		// get the bytes for the WordDocument substream only
		Storage doc;
		try
		{
			doc = directories.getDirectoryByName( "WordDocument" );
		}
		catch( StorageNotFoundException e )
		{
			throw new InvalidFileException( "InvalidFileException: Not Word '97 or later version.  Unsupported file format." );
		}
		return doc.getBlockReader();
	}

	/**
	 * get the bytes of the XLS file.
	 */
	public BlockByteReader getXLSBlockBytes()
	{
		// get the bytes for the Workbook substream only
		Storage book;
		try
		{
			book = directories.getDirectoryByName( "Workbook" );
		}
		catch( StorageNotFoundException e )
		{
			try
			{
				book = directories.getDirectoryByName( "Book" );
			}
			catch( StorageNotFoundException e1 )
			{
				log.warn( "Not Excel '97 (BIFF8) or later version.  Unsupported file format." );
				throw new InvalidFileException( "InvalidFileException: Not Excel '97 (BIFF8) or later version.  Unsupported file format." );
			}
		}
		return book.getBlockReader();
	}

	/**
	 * read LEO file information from header.
	 */
	public synchronized int[] init( ByteBuffer bbuf )
	{
		int pos = 0;
		CompatibleVector FATSectors = new CompatibleVector();    // one or more sectors which hold the FAT (File Allocation Table or indexes into the sectors)

		bigBlocks = new ArrayList();
		int len = bbuf.limit() / BIGBLOCK.SIZE;
		// get ALL BIGBLOCKS (512 byte chunks of file)
			log.debug( "INIT: Total Number of bigblocks:  " + len );
		for( int i = 0; i < len; i++ )
		{
			BIGBLOCK bbd = new BIGBLOCK();
			bbd.init( bbuf, i, pos );
			pos += BIGBLOCK.SIZE;
			bigBlocks.add( bbd );
		}

		// Encrypted workbooks can have random overages.
		// not ideal, but store this value in LEO and get from the storage if its named 'EncryptedPackage'
		int encryptionStorageOverageLen = (bbuf.limit() % BIGBLOCK.SIZE);
		if( encryptionStorageOverageLen > 0 )
		{
			int filepos = len * BIGBLOCK.SIZE;
			if( encryptedXLSX )
			{
				bbuf.position( filepos );
				encryptionStorageOverage = new byte[encryptionStorageOverageLen];
				bbuf.get( encryptionStorageOverage, 0, encryptionStorageOverage.length );
			}
			else
			{
				BIGBLOCK bbd = new BIGBLOCK();
				bbd.init( bbuf, len, pos );    // filepos);
				pos += encryptionStorageOverageLen;
				bigBlocks.add( bbd );
			}
		}

		/***** Read in the file header */
		// header holds directory start sector and FAT/miniFAT/extraFAT or Sector Chain index info 
		header = new LEOHeader(); // read the LEO file header rec

		// is a valid WorkBook File?
		if( !header.init( bbuf ) )
		{
			throw new InvalidFileException( getFileName() + " is not a valid OLE File." );
		}

			log.debug( "Header: " );
			log.debug( "numbFATSectors: " + header.getNumFATSectors() );
			log.debug( "numMiniFATSectors: " + header.getNumMiniFATSectors() );
			log.debug( "numbExtraDIFATSectors: " + header.getNumExtraDIFATSectors() );
			log.debug( "rootstart: " + header.getRootStartPos() );
			log.debug( "miniFATStart: " + header.getMiniFATStart() );
		BIGBLOCK headerblock = (BIGBLOCK) bigBlocks.get( 0 );
		headerblock.setInitialized( true );

		/***** Read in the FAT sectors - the sectors or blocks used for the FAT (==Sector Chain Index or File Allocation Table)  */
		FATSectors = getFATSectors();

		/***** turn the FAT Sectors into the FAT or Sector Index Chain */
		byte[] blx = LEOFile.getBytes( FATSectors );
		int[] FAT = null;
		FAT = LEOFile.readFAT( blx );
		blx = null; // done

			log.debug( "FAT: {}", Arrays.toString( FAT ) );

		/*****  Read the Directory blocks */
		directories = new StorageTable();
		directories.init( bbuf, header, bigBlocks, FAT );
		return FAT;
	}

	/**
	 * create a FileBuffer from a file, use system property to determine
	 * whether to use a temp file.
	 *
	 * @param file  containing XLS bytes
	 * @param fpath
	 * @return
	 */
	public final static FileBuffer readFile( File fpath )
	{
		boolean usetempfile = false;
		String tmpfu = (String) System.getProperties().get( "com.extentech.formats.LEO.usetempfile" );

		if( tmpfu != null )
		{
			usetempfile = tmpfu.equalsIgnoreCase( "true" );
		}

		return readFile( fpath, usetempfile );
	}

	/**
	 * create a FileBuffer from a file, use boolean parameter to determine
	 * whether to use a temp file.
	 *
	 * @param file    containing XLS bytes
	 * @param whether to use a tempfile
	 * @return
	 */
	public final static FileBuffer readFile( File fpath, boolean usetempfile )
	{
		if( usetempfile )
		{
			return FileBuffer.readFileUsingTemp( fpath );
		}
		return FileBuffer.readFile( fpath );
	}

	/**
	 * read in a WorkBook ByteBuffer from a file path.
	 * <p/>
	 * from here on out, we're reading pointers to the bytes on disk
	 * <p/>
	 * access data directly on disk through the ByteBuffer.
	 * <p/>
	 * <br> By default, ExtenXLS will lock open WorkBook files, to close the file after parsing and work with a
	 * temporary file instead, use the following setting:
	 * <br><br>
	 * System.getProperties().put("com.extentech.formats.LEO.usetempfile", "true");
	 * <br><br>
	 * IMPORTANT NOTE: You will need to clean up temp files occasionally in your user directory (temp filenames will begin
	 * with "ExtenXLS_".)
	 * <br><br>
	 */
	public final static FileBuffer readFile( String fpath )
	{
		return LEOFile.readFile( new File( fpath ) );
	}

	/**
	 * returns the table of int locations
	 * in the file for the BIGBLOCK linked list.
	 */
	public final static int[] readFAT( List vect )
	{
		byte[] data = getBytes( vect );
		return readFAT( data );
	}

	/**
	 * returns the table of int locations
	 * in the file for the BIGBLOCK linked list.
	 */
	private final static int[] readFAT( byte[] data )
	{
		int[] bbs = new int[data.length / 4];
		int pos = 0;
		for( int i = 0; i < data.length; )
		{
			bbs[pos++] = ByteTools.readInt( data[i++], data[i++], data[i++], data[i++] );
		}
		data = null;
		return bbs;
	}

	public final static byte[] getBytes( Block[] outblocks )
	{
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		getBytes( outblocks, out );
		return out.toByteArray();
	}

	public final static OutputStream getByteStream( Block[] outblocks )
	{
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		getBytes( outblocks, out );
		return out;
	}

	/**
	 * Call this method for each storage, root storage first...
	 *
	 * @param outblocks
	 * @param out
	 */
	public final static void getBytes( Block[] outblocks, OutputStream out )
	{
		List blxt = new ArrayList();
		for( Block outblock : outblocks )
		{
			blxt.add( outblock );
		}
		getBytes( blxt, out );
	}

	/**
	 * Call this method for each storage, root storage first...
	 *
	 * @param outblocks
	 * @param out
	 */
	public final static void getBytes( List outblocks, OutputStream out )
	{
		Iterator it = outblocks.iterator();
		int i = 0;
		while( it.hasNext() )
		{
			Block boc = (Block) it.next();
			try
			{
				if( boc != null )
				{
//					if (LEOFile.DEBUG)Logger.logInfo("INFO: LEOFile.getBytes() getting bytes from Block: "+ i++	+ ":" + boc.getBlockIndex());
					out.write( boc.getBytes() );
					i++;    // count how many blocks written
				}
			}
			catch( IOException a )
			{
				log.warn( "ERROR: gettting bytes from blocks failed: " + a );
			}
		}
	}

	/**
	 * return the bytes for the blocks in a vector
	 */
	public final static byte[] getBytes( List bbvect )
	{
		if( bbvect == null )
		{
			return new byte[]{ };
		}
		Block[] b = new Block[bbvect.size()];
		b = (Block[]) bbvect.toArray( b );
		return LEOFile.getBytes( b );
	}

	/**
	 * return the bytes for the blocks in a vector
	 */
	public final static OutputStream getByteStream( List bbvect )
	{
		if( bbvect == null )
		{
			return new ByteArrayOutputStream();
		}
		Block[] b = new Block[bbvect.size()];
		b = (Block[]) bbvect.toArray( b );
		return LEOFile.getByteStream( b );
	}

	/**
	 * get the number of Block records
	 * that this Storage needs to store its byte array
	 */
	public final static int getSizeInBlocks( int sz, int blocksize )
	{
		int size = sz / blocksize;
		float realnum = ((float) (sz) / blocksize);
		if( (realnum - size) > 0 )
		{
			size++;
		}
		return size;
	}

	/**
	 * Get the minimum blocks that this file can encompass
	 *
	 * @return
	 */
	public int getMinBlocks()
	{
		return (header.getMinStreamSize() / BIGBLOCK.SIZE);
	}

	/**
	 * @return Returns the encryptionStorageOverage.
	 */
	public byte[] getEncryptionStorageOverage()
	{
		if( encryptionStorageOverage == null )
		{
			encryptionStorageOverage = new byte[0];
		}
		return encryptionStorageOverage;
	}

	/**
	 * reads in the FAT sectors from the compound file using info from the LEOFile header
	 *
	 * @return FATSectors == one or more sectors used to define the Sector Chain Index (==File Allocation Table == FAT)
	 */
	private CompatibleVector getFATSectors()
	{
		CompatibleVector FATSectors = new CompatibleVector();
		int[] DIFAT = header.getDIFAT(); // the Index for the FAT(which is the File Allocation Table or Sector Chain Array), id's of 1st 109 sectors or blocks 
			log.debug( "FAT Blocks: {}" , Arrays.toString( DIFAT ) );
		//**** read in the FAT index ****//
		//**** First get the inital blocks before the extra DIFAT sectors ****//
		int FATidx = 0;
		int blockidx = 0;
		int bsz = bigBlocks.size() - 1;
		for(; blockidx < Math.min( DIFAT.length, MAXDIFATLEN ); blockidx++ )
		{
			FATidx = DIFAT[blockidx];
			if( FATidx <= bsz )
			{
				BIGBLOCK bbd = (BIGBLOCK) bigBlocks.get( FATidx );
				// each FAT entry is a 1-based index
				//if (DEBUG)Logger.logInfo("INFO: LEOFile Got A FAT index at: "+ String.valueOf(FATidx));
				bbd.setIsDepotBlock( true );
				FATSectors.add( bbd );
			}
			else
			{
				// Usually caused by CHECK ByteStreamer.writeOut() dlen/filler handling
				log.error( "LEOFile.init failed. FAT Index Attempting to fetch Block past end of blocks." );
				throw new InvalidFileException(
						"Input file truncated. LEOFile.init failed. FAT Index Attempting to fetch Block past end of blocks." );
			}
		}

		//******* read in any Exrra DIFAT sectors  *******//
		int numExtraDIFATSectors = header.getNumExtraDIFATSectors();
		int extraDIFATStart = header.getExtraDIFATStart() + 1;
		int chainIndex = extraDIFATStart;
		for( int x = 0; x < numExtraDIFATSectors; x++ )
		{
			BIGBLOCK xind = (BIGBLOCK) bigBlocks.get( chainIndex );
			ByteBuffer extraSectorBytes = ByteBuffer.wrap( xind.getBytes() );
			extraSectorBytes.position( 0 );
			extraSectorBytes.order( ByteOrder.LITTLE_ENDIAN );
			int numExtraSectorElements = Math.min( (header.getNumFATSectors() - blockidx), IDXBLOCKSIZE );
			for( int i = 0; i < numExtraSectorElements; i++ )
			{
				if( i < (IDXBLOCKSIZE - 1) )
				{  // -1 because must always have room for end of sector mark
					int bbloc = extraSectorBytes.getInt() + 1;
					if( bbloc <= bsz )
					{
						BIGBLOCK bbd = (BIGBLOCK) bigBlocks.get( bbloc ); //fails at 41k recs jpm test
						bbd.setIsDepotBlock( true );
						bbd.setIsExtraSector( true );
						FATSectors.add( bbd );
						blockidx++;
					}
					else
					{
						log.error( "LEOFile.init failed. Attempting to fetch Invalid Extra Sector Block." );
						throw new InvalidFileException( "LEOFile.init failed. Attempting to fetch Invalid Extra Sector Block." );
					}
				}
				else
				{
					// the index to the next extra DIFAT sector is the last int in the chain
					chainIndex = extraSectorBytes.getInt() + 1;
				}
			}
		}
		return FATSectors;
	}

}