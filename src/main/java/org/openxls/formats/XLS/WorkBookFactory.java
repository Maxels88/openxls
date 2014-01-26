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

import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.formats.LEO.BlockByteConsumer;
import org.openxls.formats.LEO.BlockByteReader;
import org.openxls.formats.LEO.LEOFile;
import org.openxls.toolkit.ByteTools;
import org.openxls.toolkit.ProgressListener;
import org.openxls.toolkit.ProgressNotifier;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Serializable;
import java.util.Iterator;
import java.util.LinkedHashMap;

/**
 * Factory for creating WorkBook objects from Byte streams.
 *
 * @see WorkBook
 * @see XLSRecord
 */
public class WorkBookFactory implements ProgressNotifier, XLSConstants, Serializable
{
	public static final long serialVersionUID = 1233423412323l;
	private static final Logger log = LoggerFactory.getLogger( WorkBookFactory.class );
	protected LEOFile myLEO;
	private String fname;

	// Methods from ProgressNotifier
	private ProgressListener progresslistener;
	private int progress = 0;
	private boolean done = false;
	private String progresstext = "";

	@Override
	public void register( ProgressListener j )
	{
		progresslistener = j;
		j.addTarget( this );
	}

	@Override
	public void fireProgressChanged()
	{
		// if (progresslistener != null) {
		// progresslistener.updateProgress();
		// }
	}

	@Override
	public int getProgress()
	{
		return progress;
	}

	@Override
	public String getProgressText()
	{
		return progresstext;
	}

	@Override
	public void setProgress( int progress )
	{
		this.progress = progress;
	}

	@Override
	public void setProgressText( String s )
	{
		progresstext = s;
	}

	@Override
	public boolean iscompleted()
	{
		return done;
	}

	// end ProgressNotifier methods

	/**
	 * get the file name for the WorkBook
	 */
	public String getFileName()
	{
		if( fname == null )
		{
			fname = myLEO.getFileName();
		}
		return fname;
	}

	/**
	 * sets the workbook filename associated with this wbfactory
	 *
	 * @param f
	 */
	public void setFileName( String f )
	{
		fname = f;
	}

	/**
	 * return the next opcode/length in the Stream from the given record.
	 */
	public static short lookAhead( BiffRec rec )
	{
		int i = rec.getOffset() + rec.getLength();
		BlockByteReader parsedata = rec.getByteReader();
		if( parsedata == null )
		{
			return rec.getOpcode();
		}

		byte[] b1 = parsedata.getHeaderBytes( i );

		short opcode = ByteTools.readShort( b1[0], b1[1] );
		return opcode;
	}

	LEOFile getLEOFile()
	{
		return myLEO;
	}

	/**
	 * read in a WorkBook from a byte array.
	 */
	public Book getWorkBook( BlockByteReader parsedata, LEOFile leo ) throws InvalidRecordException
	{
		Book book = new WorkBook();
		return initWorkBook( book, parsedata, leo );
	}

	/**
	 * Initialize the workbook
	 *
	 * @param book
	 * @param parsedata
	 * @param leo
	 * @return
	 * @throws InvalidRecordException
	 */
	public Book initWorkBook( Book book, BlockByteReader parsedata, LEOFile leo ) throws InvalidRecordException
	{

		BlockByteReader blockByteReader = parsedata;
		blockByteReader.setApplyRelativePosition( true );

		/** KSC: record-level validation */
		boolean bPerformRecordLevelValidation = false;    // perform record-level validation if set
		if( System.getProperty( VALIDATEWORKBOOK ) != null )
		{
			if( System.getProperty( VALIDATEWORKBOOK ).equals( "true" ) )
			{
				bPerformRecordLevelValidation = true;
			}
		}
		java.util.LinkedHashMap<Short, R> curSubstream = null;
		java.util.LinkedHashMap<Short, R> sheetSubstream = null;
		if( bPerformRecordLevelValidation )
		{
			java.util.LinkedHashMap<Short, R> globalSubstream = new java.util.LinkedHashMap();
			fillGlobalSubstream( globalSubstream );
			sheetSubstream = new java.util.LinkedHashMap();
			fillWorksSheetSubstream( sheetSubstream );
			curSubstream = globalSubstream;
		}

		myLEO = leo;

		book.setFactory( this );
		boolean infile = false;
		boolean isWBBOF = true;
		short opcode = 0x00;
		short reclen = 0x00;
		short lastOpcode = 0x00;
		int BofCount = 0; // track the number of 'Bof' records

		BiffRec rec = null;
		int blen = parsedata.getLength();

		// init the progress listener
		progresstext = "Initializing Workbook...";
		progress = 0;
		if( progresslistener != null )
		{
			progresslistener.setMaxProgress( blen );
		}
		fireProgressChanged();
			log.info( "XLS File Size: " + String.valueOf( blen ) );

		for( int i = 0; i <= (blen - 4); )
		{

			fireProgressChanged(); // ""
			byte[] headerbytes = parsedata.getHeaderBytes( i );
			opcode = ByteTools.readShort( headerbytes[0], headerbytes[1] );
			reclen = ByteTools.readShort( headerbytes[2], headerbytes[3] );

			if( ((lastOpcode == EOF) && (opcode == 0)) || (opcode == 0xffffffff) )
			{
				int startpos = i - 3;
				int junkreclen = 0;
				int offset = 0;

				if( offset != 0 )
				{
					junkreclen = (offset - startpos);
				}
				else
				{
					junkreclen = blen - i;
				}
				i += junkreclen - 1;
				i += junkreclen - 1;
			}
			else
			{ // REAL REC
				// sanity checks
				if( reclen < 0 )
				{
					throw new InvalidRecordException( "WorkBookFactory.getWorkBook() Negative Reclen encountered pos:" + i + " opcode:0x" + Integer
							.toHexString( opcode ) );
				}
				if( (reclen + 1) > blen )
				{
					throw new InvalidRecordException( "WorkBookFactory.getWorkBook() Reclen longer than data pos:" + i + " opcode:0x" + Integer
							.toHexString( opcode ) );
				}

				if( (opcode == BOF) || infile )
				{ // if the first Bof has been
					// reached, start
					infile = true;

					// Init Record'
					rec = parse( book, opcode, i, reclen, blockByteReader );

					if( progresslistener != null )
					{
						progresslistener.setValue( i );
					}

					/**** KSC: record-level validation ****/
					if( bPerformRecordLevelValidation && (curSubstream != null) )
					{
						markRecord( curSubstream, rec, opcode );
					}

					// write to the dump file if necessary
					if( WorkBookHandle.dump_input != null )
					{
						try
						{
							WorkBookHandle.dump_input.write( "-------------------------------------" + "-------------------------\n" + ((XLSRecord) rec)
									.getRecDesc() + ByteTools.getByteDump( blockByteReader.get( (XLSRecord) rec, 0, reclen ), 0 ) + "\n" );
							WorkBookHandle.dump_input.flush();
						}
						catch( Exception e )
						{
							log.error( "error writing to dump file, ceasing dump output: ", e );
							WorkBookHandle.dump_input = null;
						}
					}

					if( rec == null )
					{ // Effectively an EOF
							log.debug( "done parsing WorkBook storage." );
						done = true;
						progresstext = "Done Reading WorkBook.";
						fireProgressChanged();
						return book;
					}
					// not used anymore ((XLSRecord)rec).resetCacheBytes();
					// int reco = rec.getOffset() ;
					// int recl = rec.getLength();
					int thisrecpos = i + reclen + 4;

					if( opcode == BOF )
					{
						if( isWBBOF )
						{    // do first Bof initialization
							book.setFirstBof( (Bof) rec );
							isWBBOF = false;
						}
						else if( bPerformRecordLevelValidation && (BofCount == 0) && (lastOpcode != EOF) && (curSubstream != null) )
						{
							/***** KSC: record-level validation ****/
							// invalid record structure-  no EOF before BOF
							validateRecords( curSubstream, book, rec.getSheet() );
						}

						BofCount++;
						/***** KSC: record-level validation ****/
						if( bPerformRecordLevelValidation && (curSubstream == null) )
						{
							// after global substream is processed, switch to sheet substream
							reInitSubstream( sheetSubstream );
							curSubstream = sheetSubstream;
							curSubstream.get( BOF ).isPresent = true;
							curSubstream.get( BOF ).recordPos = 0;
						}
					}
					else if( opcode == EOF )
					{
						BofCount--;

						/***** KSC: record-level validation ****/
						if( bPerformRecordLevelValidation && (BofCount == 0) && (curSubstream != null) )
						{
							validateRecords( curSubstream, book, rec.getSheet() );
							curSubstream = null;
						}
					}
					// end of Workbook
					if( BofCount == -1 )
					{
							log.debug( "Last Bof" );
						i += reclen;
						thisrecpos = blen;
					}
					if( thisrecpos > 0 )
					{
						i = thisrecpos;
					}
					lastOpcode = opcode;
				}
				else
				{
					throw new InvalidRecordException( "No valid record found." );
				}

			}
		}
			log.info( "done" );
		progress = blen;
		progresstext = "Done Reading WorkBook.";
		fireProgressChanged();

		done = true;
		// flag the book so we know it's ready for shared access
		// book.setReady(true); ENTERPRISE ONLY
		// recordata.setApplyRelativePosition(false);
		return book;
	}

	/**
	 * create the individual records based on type
	 */
	protected synchronized BiffRec parse( Book book, short opcode, int offset, int datalen, BlockByteReader bytebuf ) throws
	                                                                                                                  InvalidRecordException
	{
//Logger.logInfo( "Opcode/Offset/DataLen: " + opcode + ", " + offset + ", " + datalen );
		// sanity checks
		if( (datalen < 0) || (datalen > XLSConstants.MAXRECLEN) )
		{
			throw new InvalidRecordException( "InvalidRecordException BAD RECORD LENGTH: " + " off: " + offset + " op: " + Integer.toHexString(
					opcode ) + " len: " + datalen );
		}
		if( (offset + datalen) > bytebuf.getLength() )
		{
			throw new InvalidRecordException( "InvalidRecordException RECORD LENGTH LONGER THAN FILE: " + " off: " + offset + " op: " + Integer
					.toHexString( opcode ) + " len: " + datalen + " buflen:" + bytebuf.getLength() );
		}

		// Create a new Record
		BiffRec rec = XLSRecordFactory.getBiffRecord( opcode );

		// init the mighty rec
		rec.setWorkBook( (WorkBook) book );
		rec.setByteReader( bytebuf );
		rec.setLength( (short) datalen );
		rec.setOffset( offset );
		rec.setStreamer( book.getStreamer() );

		// send it to the CONTINUE handler
		book.getContinueHandler().addRec( rec, (short) datalen );
		// add it to the record stream
		return book.addRecord( rec, true );
	}

	/**
	 * Find the location of the next particular opcode
	 */
	protected static int getNextOpcodeOffset( short op, BlockByteConsumer rec, BlockByteReader parsedata )
	{
		boolean found = false;
		int x = rec.getOffset();
		short opcode = 0x0;
		while( !found && (x < (parsedata.getLength() - 2)) )
		{
			opcode = ByteTools.readShort( parsedata.get( rec, x ), parsedata.get( rec, ++x ) );
			if( opcode == op )
			{
				found = true;
				break;
			}
		}
		if( !found )
		{
			return 0;
		}
		return x - 3;
	}

	/**
	 * record-level validation:
	 * <li>have ordered list of records for each substream (wb/global, worksheet)
	 * <li>upon each record that is processed, look up in list and mark present + record pos in correseponding list (streamer or sheet)
	 * <li>upon EOF for stream, traverse thru list, if required and not present, add
	 * <br>Limitations:
	 * This methodology does NOT validate Chart records, Chart-only sheets, Macro Sheets, Dialog Sheets or Custom Views
	 */
	/**
	 * fill map with EVERY possible global (workbook)-level record, IN ORDER, plus flag if they are required or not
	 * Used in record-level validation
	 *
	 * @param map
	 */
	private static void fillGlobalSubstream( LinkedHashMap<Short, R> map )
	{
		// ordered list of all records in the global (workbook) substream, along
		// with if they are required or not
		map.put( BOF, new R( true ) );
		map.put( (short) 134, new R( false ) ); // WriteProtect
		map.put( FILEPASS, new R( false ) );
		map.put( (short) 96, new R( false ) ); // Template
		map.put( INTERFACE_HDR, new R( true ) );
		map.put( (short) 193, new R( true ) ); // Mms
		map.put( (short) 226, new R( true ) ); // InterfaceEnd
		map.put( (short) 92, new R( true ) ); // WriteAccess
		map.put( (short) 91, new R( false ) ); // FileSharing
		map.put( (short) 66, new R( true ) ); // CodePage
		map.put( (short) 441, new R( false ) ); // LEL - SHOULD BE is required??
		map.put( (short) 353, new R( true ) ); // DSF - is required??
		map.put( (short) 448, new R( false ) ); // Excel9File
		map.put( TABID, new R( true ) );
		map.put( (short) 211, new R( false ) ); // ObjProj
		map.put( (short) 445, new R( false ) ); // *[ObNoMacros]
		map.put( CODENAME, new R( false ) );
		map.put( (short) 156, new R( false ) ); // BuiltInFnGroupCount
		map.put( (short) 154, new R( false ) ); // FnGroupName
		map.put( (short) 2200, new R( false ) ); // FnGrp12
		map.put( (short) 222, new R( false ) ); // OleObjectSize
		map.put( WINDOW_PROTECT, new R( true ) );
		map.put( PROTECT, new R( true ) );
		map.put( PASSWORD, new R( true ) );
		map.put( PROT4REV, new R( true ) );
		map.put( (short) 444, new R( true ) ); // Prot4RevPass
		map.put( WINDOW1, new R( true ) );
		map.put( BACKUP, new R( true ) );
		map.put( (short) 141, new R( true ) ); // HideObj
		map.put( DATE1904, new R( true ) );
		map.put( (short) 14, new R( true ) ); // CalcPrecision
		map.put( (short) 439, new R( true ) ); // RefreshAll
		map.put( BOOKBOOL, new R( true ) );
		map.put( FONT, new R( true ) ); // Is required???
		map.put( FORMAT, new R( true ) ); // Is required??
		map.put( XF, new R( true ) ); // Is required???
		map.put( (short) 2172, new R( false ) ); // XfCRC
		map.put( (short) 2173, new R( false ) ); // XfExt
		map.put( (short) 2189, new R( false ) ); // DXF- Is required???
		map.put( STYLE, new R( true ) ); // Is required???
		map.put( (short) 2194, new R( false ) ); // StyleExt
		map.put( TABLESTYLES, new R( false ) );
		map.put( (short) 2191, new R( false ) ); // TableStyle
		map.put( (short) 2192, new R( false ) ); // TABLESTYLEELEMENT
		map.put( PALETTE, new R( false ) );
		map.put( (short) 4188, new R( false ) ); // CLRTCLIENT
		map.put( SXSTREAMID, new R( false ) );
		map.put( SXVS, new R( false ) );
		map.put( DCONNAME, new R( false ) );
		map.put( DCONBIN, new R( false ) );
		map.put( DCONREF, new R( false ) );
		map.put( (short) 208, new R( false ) ); // SXTbl
		map.put( (short) 210, new R( false ) ); // SxTbpg
		map.put( (short) 209, new R( false ) ); // SXTBRGIITM
		map.put( SXSTRING, new R( false ) );
		map.put( (short) 220, new R( false ) ); // DbOrParamQry
		map.put( SXADDL, new R( false ) );

		map.put( (short) 184, new R( false ) ); // DocRoute
		map.put( (short) 185, new R( false ) ); // RECIPNAME
		map.put( USERBVIEW, new R( false ) ); // SHOULD BE Is but isn't present
		// *************************
		map.put( (short) 352, new R( true ) ); // UsesELFs
		map.put( BOUNDSHEET, new R( true ) );
		map.put( (short) 2180, new R( false ) ); // MDTInfo
		// *ContinueFrt12
		map.put( (short) 2181, new R( false ) ); // MDXStr
		// *ContinueFrt12
		map.put( (short) 2182, new R( false ) ); // MDXTuple
		// *ContinueFrt12
		map.put( (short) 2183, new R( false ) ); // MDXSet
		map.put( (short) 2184, new R( false ) ); // MDXProp
		map.put( (short) 2185, new R( false ) ); // MDXKPI
		map.put( (short) 2186, new R( false ) ); // MDB
		map.put( (short) 2202, new R( false ) ); // MTRSettings
		map.put( (short) 2211, new R( false ) ); // ForceFullCalculation
		map.put( COUNTRY, new R( true ) );
		map.put( SUPBOOK, new R( false ) ); // SHOULD BE IsRequired but isn't present ******************
		map.put( EXTERNNAME, new R( false ) );
		map.put( XCT, new R( false ) );
		map.put( CRN, new R( false ) );
		map.put( EXTERNSHEET, new R( false ) );
		map.put( NAME, new R( false ) ); // Name is Required???
		map.put( (short) 2196, new R( false ) ); // NameCmt
		map.put( (short) 2201, new R( false ) ); // NameFnGrp12
		map.put( (short) 2195, new R( false ) ); // NamePublish
		map.put( (short) 2067, new R( false ) ); // RealTimeData A required record
		// is not present: 430
		// *ContinueFrt
		map.put( (short) 449, new R( false ) ); // RecalcId
		map.put( (short) 2150, new R( false ) ); // HFPicture should be required????
		map.put( MSODRAWINGGROUP, new R( false ) ); // should be required??
		// Continue
		map.put( SST, new R( true ) );        // should be required??
		// Continue
		map.put( EXTSST, new R( true ) );
		map.put( (short) 2049, new R( false ) ); // WebPub should be required???
		map.put( (short) 2059, new R( false ) ); // WOpt
		map.put( (short) 2149, new R( false ) ); // CrErr
		map.put( (short) 2147, new R( false ) ); // BookExt
		map.put( (short) 2151, new R( false ) ); // FeatHdr should be required???
		map.put( (short) 2166, new R( false ) ); // DConn should be required???
		map.put( (short) 2198, new R( false ) ); // Theme
		// *ContinueFrt12
		map.put( (short) 2203, new R( false ) ); // CompressPictures
		map.put( (short) 2188, new R( false ) ); // Compat12
		map.put( (short) 2199, new R( false ) ); // GUIDTypeLib
		map.put( EOF, new R( true ) );
	}

	/**
	 * fill map with EVERY possible worksheet-level record, IN ORDER, plus flag if they are required or not.
	 * Used in record-level validation
	 *
	 * @param map
	 */
	private static void fillWorksSheetSubstream( LinkedHashMap<Short, R> map )
	{
		map.put( BOF, new R( true ) );
		map.put( (short) 94, new R( false ) );    //	Uncalced
		map.put( INDEX, new R( true ) );
		map.put( CALCMODE, new R( true ) );
		map.put( CALCCOUNT, new R( true ) );
		map.put( (short) 15, new R( true ) );    //CalcRefMode
		map.put( (short) 17, new R( true ) );    // CalcIter
		map.put( (short) 16, new R( true ) );    // CalcDelta
		map.put( (short) 95, new R( true ) );    // CalcSaveRecalc
		map.put( PRINTROWCOL, new R( true ) );
		map.put( PRINTGRID, new R( true ) );
		map.put( (short) 130, new R( true ) );    // GridSet
		map.put( GUTS, new R( true ) );
		map.put( DEFAULTROWHEIGHT, new R( true ) );
		map.put( WSBOOL, new R( true ) );
		map.put( (short) 151, new R( false ) ); //	[Sync]
		map.put( (short) 152, new R( false ) ); //	[LPr]
		map.put( HORIZONTAL_PAGE_BREAKS, new R( false ) ); //[HorizontalPageBreaks]
		map.put( VERTICAL_PAGE_BREAKS, new R( false ) );    //[VerticalPageBreaks]
		map.put( HEADERREC, new R( true ) );
		map.put( FOOTERREC, new R( true ) );
		map.put( HCENTER, new R( true ) );
		map.put( VCENTER, new R( true ) );
		map.put( LEFT_MARGIN, new R( false ) );
		map.put( RIGHT_MARGIN, new R( false ) );
		map.put( TOP_MARGIN, new R( false ) );
		map.put( BOTTOM_MARGIN, new R( false ) );
		map.put( PLS, new R( false ) );
//	Continue 
		map.put( SETUP, new R( true ) );
		map.put( (short) 0x89c, new R( false ) );    //[HeaderFooter]
		map.put( (short) 233, new R( false ) );    //[ BkHim ]
		map.put( (short) 1048, new R( false ) );    // BigName
//       *ContinueBigName 
		map.put( PROTECT, new R( false ) );
		map.put( SCENPROTECT, new R( false ) );
		map.put( OBJPROTECT, new R( false ) );
		map.put( PASSWORD, new R( false ) );
		map.put( DEFCOLWIDTH, new R( true ) );
		map.put( COLINFO, new R( false ) );
		map.put( (short) 174, new R( false ) );    // [ScenMan *(
		map.put( (short) 175, new R( false ) );    // SCENARIO
		// Continue
		map.put( (short) 144, new R( false ) );    // Sort
		map.put( (short) 2197, new R( false ) );    // SortData
		// ContinueFrt12
		map.put( (short) 155, new R( false ) );    // Filtermode
		map.put( (short) 2164, new R( false ) );    // DropDownObjIds
		map.put( (short) 157, new R( false ) );    // AutoFilterInfo
		map.put( AUTOFILTER, new R( false ) );
		map.put( (short) 2174, new R( false ) );    // AutoFilter12
		map.put( (short) 2197, new R( false ) );    // SortData
//*ContinueFrt12]
		map.put( DIMENSIONS, new R( true ) );
//	[CELLTABLE]
		map.put( ROW, new R( false ) );
		// celltable entries
		map.put( BLANK, new R( false ) );
		map.put( MULBLANK, new R( false ) );
		map.put( RK, new R( false ) );
		map.put( BOOLERR, new R( false ) );
		map.put( NUMBER, new R( false ) );
		map.put( LABELSST, new R( false ) );
		map.put( (short) 94, new R( false ) );    // uncalced
		map.put( FORMULA, new R( false ) );
		map.put( MULRK, new R( false ) );
		map.put( STRINGREC, new R( false ) );
		;
		map.put( SHRFMLA, new R( false ) );
		map.put( ARRAY, new R( false ) );
		map.put( TABLE, new R( false ) );
		map.put( (short) 450, new R( false ) );    // EntExU2
//	[OBJECTS]
		map.put( MSODRAWING, new R( false ) );
		map.put( TXO, new R( false ) );
		map.put( OBJ, new R( false ) );
		map.get( OBJ ).altPrecedor = new short[]{ MSODRAWING };        // Obj can follow MSODRAWING or TXO
		//Charts -- most record opcodes are > 4000 - ignore !!
		map.put( CHARTFRTINFO, new R( false ) );
		// do not include the majority of chart records but do include these:
		map.put( STARTBLOCK, new R( false ) );
		map.put( CRTLAYOUT12, new R( false ) );
		map.put( CRTLAYOUT12A, new R( false ) );
		map.put( CATLAB, new R( false ) );
		map.put( DATALABEXT, new R( false ) );
		map.put( DATALABEXTCONTENTS, new R( false ) );
		map.put( (short) 2138, new R( false ) );    // FrtFontList
		map.put( (short) 2213, new R( false ) );    // TextPropsStream
		map.put( (short) 2214, new R( false ) );    // RichTextStream
		map.put( (short) 2212, new R( false ) );    // ShapePropsStream
		map.put( (short) 2206, new R( false ) );    // CtrlMlFrt
		map.put( ENDBLOCK, new R( false ) );
		map.put( STARTOBJECT, new R( false ) );
		map.put( FRTWRAPPER, new R( false ) );
		map.put( ENDOBJECT, new R( false ) );
//	CHARTSHEETCONTENT = BOF [WriteProtect] [SheetExt] [WebPub] *HFPicture PAGESETUP PrintSize [HeaderFooter] [BACKGROUND] *Fbi *Fbi2 [ClrtClient] [PROTECTION] [Palette] [SXViewLink] [PivotChartBits] [SBaseRef] [MsoDrawingGroup] OBJECTS Units CHARTFOMATS SERIESDATA *WINDOW *CUSTOMVIEW [CodeName] [CRTMLFRT] EOF
//Chart Begin *2FONTLIST SclPlotGrowth [FRAME] *SERIESFORMAT *SS ShtProps *2DFTTEXT AxesUsed 1*2AXISPARENT [CrtLayout12A] [DAT] *ATTACHEDLABEL [CRTMLFRT] *([DataLabExt StartObject] ATTACHEDLABEL [EndObject]) [TEXTPROPS] *2CRTMLFRT End
		map.put( MSODRAWINGSELECTION, new R( false ) );
		map.put( (short) 2150, new R( false ) );    //*HFPicture
		map.put( NOTE, new R( false ) );
		//*PIVOTVIEW
		map.put( SXVIEW, new R( false ) );
		map.put( SXVD, new R( false ) );
		map.put( SXVI, new R( false ) );
		map.put( SXVDEX, new R( false ) );
		map.put( SXIVD, new R( false ) );
		map.put( SXPI, new R( false ) );
		map.put( SXDI, new R( false ) );
		map.put( SXLI, new R( false ) );
		map.put( SXEX, new R( false ) );
		map.put( (short) 247, new R( false ) );    //SXSelect
		map.put( (short) 240, new R( false ) );    // SXRULE
		map.put( SXFORMAT, new R( false ) );
//	PIVOTRULE
		map.put( (short) 244, new R( false ) );    // SXDXF
		map.put( QSISXTAG, new R( false ) );
		//DBQUERYEXT
		map.put( (short) 2051, new R( false ) );    // DBQUERYEX
		map.put( (short) 2052, new R( false ) );    // EXTSTRING
		map.put( SXSTRING, new R( false ) );
		map.put( (short) 2060, new R( false ) );    // SXVIEWEX
		map.put( (short) 2061, new R( false ) );    // SXTH
		map.put( (short) 2062, new R( false ) );    // SXPIEx
		map.put( (short) 256, new R( false ) );    // SXVDTEx
		map.put( SXVIEWEX9, new R( false ) );
		map.put( SXADDL, new R( false ) );

		map.put( DCON, new R( false ) );
		map.put( DCONNAME, new R( false ) );
		map.put( DCONBIN, new R( false ) );
		map.put( DCONREF, new R( false ) );
		map.put( WINDOW2, new R( true ) );
		map.put( PLV, new R( false ) );
		map.put( SCL, new R( false ) );
		map.put( PANE, new R( false ) );
		map.put( SELECTION, new R( false ) );
		map.put( (short) 426, new R( false ) );    // UserSViewBegin
/*		map.put(SELECTION, new R(false));
		map.put(HORIZONTAL_PAGE_BREAKS, new R(false)); 
		map.put(VERTICAL_PAGE_BREAKS, new R(false));	 
		map.put(HEADERREC, new R(false));
		map.put(FOOTERREC, new R(false));
		map.put(HCENTER,  new R(false));
		map.put(VCENTER, new R(false));
		map.put(LEFT_MARGIN, new R(false));
		map.put(RIGHT_MARGIN,  new R(false));
		map.put(TOP_MARGIN,  new R(false));
		map.put(BOTTOM_MARGIN,  new R(false));
		map.put(PLS, new R(false));
		map.put(SETUP, new R(false));
		map.put((short) 51, new R(false));	//([PrintSize] 
		map.put((short) 2204, new R(false));	//[HeaderFooter]*/
		// here ??? [AUTOFILTER] 
		map.put( (short) 427, new R( false ) );    //UserSViewEnd

		map.put( (short) 319, new R( false ) );    // RRSort
		map.put( (short) 153, new R( false ) );    //[DxGCol]
		map.put( MERGEDCELLS, new R( false ) );    //*MergeCells
		map.put( (short) 351, new R( false ) );    // [LRng]
//	*QUERYTABLE 
		map.put( PHONETIC, new R( false ) );
		map.put( CONDFMT, new R( false ) );
		map.put( CF, new R( false ) );
		map.put( CONDFMT12, new R( false ) );
		map.put( CF12, new R( false ) );
		map.put( (short) 2171, new R( false ) );    // CFEx
		map.put( HLINK, new R( false ) );
		map.put( (short) 2048, new R( false ) );    //HLinkTooltip
		map.put( DVAL, new R( false ) );
		map.put( DV, new R( false ) );
		map.put( CODENAME, new R( false ) );    // [CodeName]
		map.put( (short) 2049, new R( false ) );    // *WebPub
		map.put( (short) 2156, new R( false ) );    // *CellWatch
		map.put( (short) 2146, new R( false ) );    // SheetExt]
		map.put( FEATHEADR, new R( false ) );
		map.put( (short) 2152, new R( false ) );    // FEAT
//		*FEAT11 
//		*RECORD12
		map.put( (short) 2248, new R( false ) );    // UNKNOWN RECORD!!!
		map.put( EOF, new R( true ) );
	}

	/**
	 * add all missing REQUIRED records from the appropriate substream for record-level validation
	 *
	 * @param map  substream list storing if present, required ...
	 * @param book
	 */
	private void validateRecords( LinkedHashMap<Short, R> map, Book book, Boundsheet bs )
	{
		// use array instead of iterator as may have to traverse more than once
		Short[] opcodes = new Short[0];
		java.util.Set<Short> ss = map.keySet();
		opcodes = ss.toArray( opcodes );
		R lastR = new R( false );
		lastR.recordPos = 0;
		int lastOp = -1;
		if( (bs != null) && bs.isChartOnlySheet() )
		{
			return;    // don't validate chart-only sheets
		}

		// traverse thru stream list, ensuring required records are present; create if not
		for( Short op : opcodes )
		{
			R r = map.get( op );
			if( !r.isPresent && r.isRequired )
			{
				// System.out.println("A required record is not present: " +  op);
				// Create a new Record
				BiffRec rec = createMissingRequiredRecord( op, book, bs );
				int recPos = lastR.recordPos + 1;
				try
				{
					while( true )
					{    // now get to LAST record if there are multiple records
						BiffRec lastRec = (bs == null) ? book.getStreamer().getRecordAt( recPos ) : (BiffRec) bs.getSheetRecs()
						                                                                                        .get( recPos );
						if( lastRec.getOpcode() != lastOp )
						{
							break;
						}
						recPos++;
					}
				}
				catch( IndexOutOfBoundsException e )
				{
				}
				book.getStreamer().addRecordAt( rec, recPos );
				book.addRecord( rec, false );    // false= already added to streamer or sheet
				if( bs != null )
				{
					rec.setSheet( bs );
				}

				r.recordPos = rec.getRecordIndex();

				// now must adjust ensuing record positions to account for inserted record
				for( Short opcode : opcodes )
				{
					R nextR = map.get( opcode );
					if( nextR.isPresent && !r.equals( nextR ) && (nextR.recordPos >= r.recordPos) )
					{
						nextR.recordPos++;
					}
				}
				r.isPresent = true;
			}
			if( r.isPresent )
			{
				lastR = r;
				lastOp = op;
			}
		}
		// go thru 1 more time to ensure record order is correct
		validateRecordOrder( map, ((bs == null) ? book.getStreamer().records : bs.getSheetRecs()), opcodes );
	}

	/**
	 * given substream map, actual list of records and list of ordered opcodes, validate record order in substream
	 * <br>if necessary, reorgainze records to follow order in map
	 *
	 * @param map     ordered map of substream records indexed by opcodes
	 * @param list    list of actual records
	 * @param opcodes ordered list of opcodes
	 */
	private static void validateRecordOrder( LinkedHashMap<Short, R> map, java.util.List list, Short[] opcodes )
	{
	/* debugging:
	System.out.println("BeFORE order:");	
	for (int zz= 0; zz < list.size(); zz++) {
		    System.out.println(zz + "-" + list.get(zz));
	}*/
		R lastR = map.get( BOF );
		short lastOp = BOF;
		for( int i = 1; i < opcodes.length; i++ )
		{
			short op = opcodes[i];
			R r = map.get( op );
			if( r.isPresent )
			{
				// compare ordered list of records (R) to actual order (denoted via recordPos)
				if( (r.recordPos < lastR.recordPos) && (r.recordPos >= 0) )
				{ // Out Of Order (NOTE: CellTable entries will have a record pos = -1)
					if( r.altPrecedor != null )
					{    // record can have more than 1 valid predecessor
						// TODO: this is ugly - do a different way ...
						for( int zz = 0; zz < r.altPrecedor.length; zz++ )
						{
							R newR = map.get( r.altPrecedor[zz] );
							if( r.recordPos > newR.recordPos )
							{
								lastR = newR;
								break;
							}
						}
						if( r.recordPos > lastR.recordPos )
						{
							continue;
						}
					}

					int origRecPos = r.recordPos;
//System.out.println("Record out of order:  r:" + op + "/" + r.recordPos + " lastR:" + lastOp + "/" + lastR.recordPos);
					// find correct insertion point by looking at ordered map
					for( int zz = i - 1; zz > 0; zz-- )
					{
						R prevr = map.get( opcodes[zz] );
						if( prevr.isPresent && (r.recordPos < prevr.recordPos) )
						{
//System.out.println("\tInsert at " + prevr.recordPos + " before op= " + opcodes[zz]);
							int recsMovedCount = 0;
							BiffRec recToMove = (BiffRec) list.get( origRecPos );
							do
							{
								list.remove( origRecPos );
								list.add( prevr.recordPos /*+ recsMovedCount*/, recToMove );
								if( recsMovedCount == 0 )
								{
									r.recordPos = recToMove.getRecordIndex();
								}
//System.out.println("\tMoved To " + recToMove.getRecordIndex());				    
								recsMovedCount++;
								recToMove = (BiffRec) list.get( origRecPos );
							} while( recToMove.getOpcode() == op );

							// after moved all the records necessary, adjust record positions
							for( Short opcode : opcodes )
							{
								R nextR = map.get( opcode );
								if( nextR.isPresent && (nextR.recordPos >= origRecPos) && (nextR.recordPos <= r.recordPos) && (opcode != op) )
								{
									nextR.recordPos -= recsMovedCount;
								}
							}
							break;
						}
					}
				}
				lastR = r;
				lastOp = op;
			}
		}
	/*System.out.println("AFTER order:");	
	for (int zz= 0; zz < list.size(); zz++) {
		    System.out.println(zz + "-" + list.get(zz));
	}*/
	}

	// TODO: handle fonts,formats,xf, style -- what's minimum necessary??
	// TODO: handle TabId
	// TODO: is Index OK?

	/**
	 * create missing required records for record-level validation
	 *
	 * @param opcode opcode of missing record
	 * @param book
	 * @param bs     boundsheet- if null it's a wb level record
	 * @return new BiffRec
	 */
	private static BiffRec createMissingRequiredRecord( short opcode, Book book, Boundsheet bs )
	{
		BiffRec record = XLSRecordFactory.getBiffRecord( opcode );
		if( bs != null )
		{
			record.setSheet( bs );
		}
		byte[] data = null;
		try
		{
			switch( opcode )
			{
				case INTERFACE_HDR:
					data = new byte[]{
							(byte) 0xB0, 0x04
					};    //codePage (2 bytes): An unsigned integer that specifies the code page. 1200==Unicode
					break;
				case (short) 193:    // Mms
					data = new byte[]{ 0, 0 };
					break;
				case (short) 226:    // InterfaceEnd
					data = new byte[]{ };
					break;
				case (short) 92:        // WriteAccess	-- user name - can be all blank
					data = new byte[]{
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20,
							0x20
					};
					break;
				case (short) 66:        // CodePage	same as interfacehdr 1200==Unicode
					data = new byte[]{ (byte) 0xB0, 0x04 };
					break;
				case (short) 353:    // DSF- truly is required?
					data = new byte[]{ 0, 0 };
					break;
				case TABID:            //
					// count how many sheets in wb
					int nSheets = 0;
					for( int z = book.getStreamer().records.size() - 1; z > 0; z-- )
					{
						BiffRec b = (BiffRec) book.getStreamer().records.get( z );
						if( b.getOpcode() == BOUNDSHEET )
						{
							while( (z > 0) && (b.getOpcode() == BOUNDSHEET) )
							{
								nSheets++;
								b = (BiffRec) book.getStreamer().records.get( --z );
							}
							break;
						}
					}
					data = new byte[]{ };
					for( int i = 0; i < nSheets; i++ )
					{
						data = ByteTools.append( ByteTools.shortToLEBytes( (short) i ), data );
					}
					break;
				case WINDOW_PROTECT:
					data = new byte[]{ 0, 0 };
					break;
				case PROTECT:
					data = new byte[]{ 0, 0 };
					break;
				case PASSWORD:
					data = new byte[]{ 0, 0 };
					break;
				case PROT4REV:
					data = new byte[]{ 0, 0 };
					break;
				case (short) 444:    // Prot4RevPass
					data = new byte[]{ 0, 0 };
					break;
				case WINDOW1:            // very general window settings
					data = new byte[]{ (byte) 0xE0, 1, 0x69, 0, 0x13, 0x38, 0x1F, 0x1D, 0x38, 0, 0, 0, 0, 0, 1, 0, 0x58, 0x02 };
					break;
				case BACKUP:
					data = new byte[]{ 0, 0 };
					break;
				case (short) 141:    // HideObj
					data = new byte[]{ 0, 0 };
					break;
				case DATE1904:
					data = new byte[]{ 0, 0 };
					break;
				case (short) 14:    // CalcPrecision
					data = new byte[]{ 1, 0 };
					break;
				case (short) 439:    // RefreshAll
					data = new byte[]{ 0, 0 };
					break;
				case BOOKBOOL:
					data = new byte[]{ 0, 0 };
					break;
				case FONT:    // truly required??
					data = new byte[]{
							(byte) 0xC8,
							0,
							0,
							0,
							(byte) 0xFF,
							0x7F,
							(byte) 0x90,
							1,
							0,
							0,
							0,
							0,
							0,
							0,
							5,
							1,
							0x41,
							0,
							0x72,
							0,
							0x69,
							0,
							0x61,
							0,
							0x6C,
							0
					};
					break;
//		case FORMAT:	// truly required??
//		    break;
				case XF:    // truly is required??
					data = new byte[]{ 0, 0, 0, 0, (byte) 0xF5, (byte) 0xFF, 0x20, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, (byte) 0xC0, 0x20 };
					break;
//		case STYLE:	// truly is required?
//		    break;
				case COUNTRY:
					data = new byte[]{ 1, 0, 1, 0 };
					break;
				case SST:
					data = new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0 };
					break;
				case EXTSST:
					data = new byte[]{ 0, 0 };    // should be rebuilt upon output ...
					break;
				case (short) 352:    // UsesELFs
					data = new byte[]{ 0, 0 };
					break;
				case INDEX:
					data = new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0x4E, 6, 0, 0 };
					break;
				case CALCMODE:
					data = new byte[]{ 1, 0 };
					break;
				case CALCCOUNT:
					data = new byte[]{ 0x64, 0 };
					break;
				case 15: //CalcRefMode
					data = new byte[]{ 1, 0 };
					break;
				case 17:    // CalcIter
					data = new byte[]{ 0, 0 };
					break;
				case 16:    // CalcDelta	An Xnum value that specifies the amount of change in value for a given cell from the previously calculated value for that cell that MUST exist for the iteration to continue. The value MUST be greater than or equal to 0.
					data = new byte[]{ (byte) 0xFC, (byte) 0xA9, (byte) 0xF1, (byte) 0xD2, 0x4D, 0x62, 0x50, 0x3F };
				case 95:    // CalcSaveRecalc
					data = new byte[]{ 1, 0 };
					break;
				case PRINTROWCOL:
					data = new byte[]{ 0, 0 };
					break;
				case PRINTGRID:
					data = new byte[]{ 0, 0 };
					break;
				case 130:    // GridSet
					data = new byte[]{ 0, 0 };
					break;
				case GUTS:
					data = new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0 };
					break;
				case DEFAULTROWHEIGHT:
					data = new byte[]{ 0, 0, (byte) 0xFF, 0 };
					break;
				case WSBOOL:
					data = new byte[]{ (byte) 0xC1, 4 };
					break;
				case HEADERREC:
				case FOOTERREC:
					data = new byte[]{ };
					break;
				case HCENTER:
					data = new byte[]{ 0, 0 };
					break;
				case VCENTER:
					data = new byte[]{ 0, 0 };
					break;
				case SETUP:
					data = new byte[]{
							0,
							0,
							(byte) 0xFF,
							0,
							1,
							0,
							1,
							0,
							1,
							0,
							4,
							0,
							0,
							0,
							0,
							0,
							0,
							0,
							0,
							0,
							0,
							0,
							(byte) 0xE0,
							0x3F,
							0,
							0,
							0,
							0,
							0,
							0,
							(byte) 0xE0,
							0x3F,
							0,
							0
					};
					break;
				case DEFCOLWIDTH:
					data = new byte[]{ 8, 0 };
					break;
				case DIMENSIONS:
					data = new byte[]{ 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
					record.setData( data );
					if( bs != null )
					{ // see if rows have already been added
						bs.setDimensions( (Dimensions) record );
						if( bs.getRowMap().size() > 0 )
						{
							int z = bs.getRows()[bs.getRowMap().size() - 1].getRowNumber();
							((Dimensions) record).setRowLast( z );
						}

//			z= ((Colinfo)bs.getColinfos().get(bs.getColinfos().size())).getColLast();
//			((Dimensions) record).setColLast(z);			
					}
					break;
				case WINDOW2:
					data = new byte[]{ (byte) 0xB6, 6, 0, 0, 0, 0, 0x40, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
					break;
				case EOF:
					break;
				default:
					System.out.println( "Must create required rec: " + opcode );
			}

			record.setData( data );
			try
			{
				record.init();
			}
			catch( Exception e )
			{
				// TODO: ExtSst needs sst ...
			}

		}
		catch( Exception e )
		{
			log.error( "Record Validation: Error creating missing record: " + opcode, e );
		}
		return record;

	}

	/**
	 * reset Record members in stream so can be used again for record-level validation
	 *
	 * @param map
	 */
	private static void reInitSubstream( LinkedHashMap<Short, R> map )
	{
		Iterator ii = map.keySet().iterator();
		while( ii.hasNext() )
		{
			short op = (Short) ii.next();
			R r = map.get( op );
			r.isPresent = false;
			r.recordPos = -1;
		}
	}

	/**
	 * mark record present and record pertinent information for record-level validation
	 */
	private static void markRecord( LinkedHashMap<Short, R> map, BiffRec rec, short opcode )
	{
		try
		{
			R r = map.get( opcode );
			if( !r.isPresent )
			{  // THIS STATEMENT FOR SOME REASON IS VERY VERY VERY TIME CONSUMING -- see testPerformance3
				r.isPresent = true;
				r.recordPos = rec.getRecordIndex();
			}
		}
		catch( NullPointerException ne )
		{
/*	    if (opcode != CONTINUE && opcode!=DBCELL 
		    && opcode < 4000 /* chart records * /
		    && !rec.isValueForCell()) // ignore CELLTABLE records
//		System.out.println("COULDN'T FIND Opcode: " + opcode);*/

		}
	}

	/**
	 * debug utility
	 */
	private static void displayRecsInStream( LinkedHashMap<Short, R> map )
	{
		Iterator<Short> ii = map.keySet().iterator();
		log.info( "Present Records" );
		while( ii.hasNext() )
		{
			short op = ii.next();
			R r = map.get( op );
			if( r.isPresent )
			{
				log.info( op + " at " + r.recordPos );
			}
		}
	}
}

/**
 * represents pertinent info for a BiffRec in an easy-to-access class
 */
class R
{
	public boolean isRequired;
	public boolean isPresent;
	public int recordPos = -1;
	public short[] altPrecedor = null;

	public R( boolean req )
	{
		isRequired = req;
	}
}

/*
 * missing records:
 * notes: 2262 record is ????? see TestInsertRows.testInsertRow0FormulaMovement
 * workingdir + "InsertRowBug1.xls"); TestColumns.testInsertColMoveReferences
 * (this.workingdir + "InsertColumnBug1.xls");
 * 
 * TestReadWrite.testNPEOnOpen: workingdir + "equilar/proxycomp3_pwc.xls");
 * COULDN'T FIND Opcode: 194
 * 
 * Sheet-level recs:
 * COULDN'T FIND Opcode: 171 ???  sheet substream just before 153 then EOF -- see testRecalc
   COULDN'T FIND Opcode: 148 -->TestNPEOnOpen
 * 
 * 
 * 
 * 
 * testCorruption.testOutOfSpec -- missing required records: workingdir +
 * "Caribou_North_And_South.xls" A required record is not present: 225 A
 * required record is not present: 193 A required record is not present: 226 A
 * required record is not present: 92 A required record is not present: 25 A
 * required record is not present: 19 A required record is not present: 431 A
 * required record is not present: 444
 * 
 * TestFormulaCalculator.testRecalc: workingdir + "Sakonnet/smile.xls" A
 * required record is not present: 225 A required record is not present: 193 A
 * required record is not present: 226 A required record is not present: 353 A
 * required record is not present: 317 A required record is not present: 18 A
 * required record is not present: 431 A required record is not present: 444 A
 * required record is not present: 64 A required record is not present: 141 A
 * required record is not present: 439 A required record is not present: 218 A
 * required record is not present: 352 A required record is not present: 255
 */

/*
 * record order issues: 659= Style 352= UsesELFs 140= Country 146= Palette
 * 
 * 2173= XfExt 2189= DXF 2190= TableStyles 2211= ForceFullCalculation
 * 
 * 
 * testValidationHandle.testEquilarFile workingdir +
 * "equilar/proxycomp3_formatted_new.xls" Record out of order: r:659/682
 * lastR:2189/903 Record out of order: r:352/732 lastR:2190/956 Record out of
 * order: r:140/753 lastR:2211/958
 * 
 * TestScenarios.borderbrokenxx, formularesultxx, numericerrorxx
 * TestFormulaParse.testRajeshOOM: workingdir + "OOXML/proxycomp3_cap_new.xls"
 * Record out of order: r:659/1304 lastR:2189/1557 Record out of order:
 * r:146/1375 lastR:2190/1610 Record out of order: r:140/1403 lastR:2211/1612
 * 
 * TestAutoFilter.XXX: workingdir + "testAutoFilter.xls" Record out of order:
 * r:659/128 lastR:2173/204 Record out of order: r:352/175 lastR:2190/292 Record
 * out of order: r:140/178 lastR:2211/294
 * 
 * also see TestAddAF w.r.t. insertion position -- supbook is wrong? ...
 * 
 * TestMemoryUsage.testMemUsageWithManyReferences: workingdir +
 * "equilar/proxycomp3_pmp_Model_T_new.xls" Record out of order: r:659/1202
 * lastR:2189/1448 Record out of order: r:146/1273 lastR:2190/1501 Record out of
 * order: r:140/1297 lastR:2211/1503 Record out of order: r:89/1300
 * lastR:35/1308
 * 
 * TestInsertRows.testInsertRowsWithFormulas.this.workingsheetsterdir +
 * "aviasphere/AirframeMaster_F20Import.xls" Record out of order: r:659/274
 * lastR:2173/352 Record out of order: r:352/323 lastR:2190/537 Record out of
 * order: r:140/330 lastR:2211/539
 * 
 * TestXMLSS.testClasicXMLReadWriteError: workingdir + "xmlsavecorruption.xls"
 * Record out of order: r:659/144 lastR:2173/208 Record out of order: r:352/194
 * lastR:2190/364 Record out of order: r:140/196 lastR:2211/366
 * 
 * TestRowCols.testRowInsertionFormatLoss.workingdir + "claritas/input/400.xls"
 * Record out of order: r:659/230 lastR:2189/426 Record out of order: r:352/280
 * lastR:2190/481 Record out of order: r:140/288 lastR:2211/483
 * 
 * TEstRowCols.testRowFormats:testRowFormats.xls Record out of order: r:659/278
 * lastR:2173/356 Record out of order: r:352/327 lastR:2190/541 Record out of
 * order: r:140/334 lastR:2211/543
 * 
 * TestRowCols.testAutoAjdustRowHeight.C:\eclipse\workspace\Testfiles\ExtenXLS\input
 * \equilar/tcr_formatted_2003.xls Record out of order: r:659/709 lastR:2189/973
 * Record out of order: r:352/774 lastR:2190/1038 Record out of order: r:140/800
 * lastR:2211/1040
 */