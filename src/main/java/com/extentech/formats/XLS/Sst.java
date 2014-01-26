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
package com.extentech.formats.XLS;

import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.formats.OOXML.OOXMLConstants;
import com.extentech.formats.OOXML.Ss_rPr;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xmlpull.v1.XmlPullParser;
import org.xmlpull.v1.XmlPullParserFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * <b>Sst: Shared String Table 0xFCh</b><br>
 * <p/>
 * Sst records contain a table of Strings possibly spanning multiple Continue
 * Records
 * <p/>
 * <p>
 * <p/>
 * <pre>
 *     offset  name        size    contents
 *     ---
 *     4       cstTotal    4       Total number of strings in this and the
 *                                 EXTSST record.
 *     8       cstUnique   4       Number of unique strings in this table.
 * 12      rgb         var     Array of unique strings
 *
 * </p>
 * </pre>
 *
 * @see Sst
 * @see Labelsst
 * @see Extsst
 */
public final class Sst extends com.extentech.formats.XLS.XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Sst.class );	private static final long serialVersionUID = 6966063306230877101L;
	private int cstTotal = -1;
	private int cstUnique = -1;
	private int boundincrement = 0;

	// continue handling
	private int numconts = -1;
	private int[] boundaries = null;
	private byte[] grbits = null;
	private List stringvector = new SstArrayList();
	private HashSet dupeSstEntries = new HashSet();
	private HashSet existingSstEntries = new HashSet();
	private Extsst myextsst = null;
	int origsstlen = 0;

	int getOrigSstLen()
	{
		return origsstlen;
	}

	Continue thiscont = null;
	int datalen = -1;
	int currbound = 0;

	byte[] deldata = null;

	public int getRealOriginalSize()
	{
		return originalsize;
	}

	@Override
	public void setData( byte[] b )
	{
		if( data == null )
		{
			originalsize = b.length;
		}
		super.setData( b );
	}

	/**
	 * set the Extsst rec for this Sst
	 */
	void setExtsst( Extsst e )
	{
		myextsst = e;
	}

	Extsst getExtsst()
	{
		return myextsst;
	}

	/**
	 * removes all existing Continues from the Sst
	 */
	@Override
	public void removeContinues()
	{
		super.removeContinues();
		continues = null;
		thiscont = null;
	}

	/**
	 * initialize the Continue Handling counters
	 */
	private void initContinues()
	{
		// create the array of record boundary offsets
		// this allows us to detect spanning UNICODE strings in CONTINUE
		// records...
		datalen = getLength();
		numconts = getContinueVect().size();
		boundaries = new int[numconts + 1];
		int thisbound = 0;
		int ir = 0;
		grbits = new byte[numconts];
		// continues = this.getContinueVect();
		Iterator it = continues.iterator();
		while( it.hasNext() )
		{
			Continue ci = (Continue) it.next();
			// byte[] b = ci.getData(); // REMOVE
			grbits[ir++] = ci.getGrbit();
			getStreamer().removeRecord( ci ); // remove existing continues
			// from stream
		}
		thisbound = getRealOriginalSize();
		boundaries[0] = thisbound;
		int lastcontlen = 0;
		for( int i = 1; i < boundaries.length; i++ )
		{
			Continue cxi = (Continue) continues.get( i - 1 );
			int contlen = cxi.getLength();
			if( cxi.getHasGrbit() )
			{
				contlen--;
			}
			thisbound += contlen;
			// if(DEBUGLEVEL > 5)Logger.logInfo( contlen + ",");
			lastcontlen += contlen;
			datalen += contlen;
			boundaries[i] = thisbound - (4 * i);
			cxi.setContinueOffset( boundaries[i - 1] );
		}
		// if(DEBUGLEVEL > 5) Logger.logInfo("");
		// for(int i = 1;i<boundaries.length;i++){
		// int contlen =continues[i-1].getLength();
		// if(DEBUGLEVEL > 5) Logger.logInfo("0x" + continues[i-1].getGrbit() +
		// ",");
		// }
		// if(DEBUGLEVEL > 5) Logger.logInfo("");
		thisbound = 0;
	}

	/**
	 * because the SST comes between the BOUNDSHEET records and all BOUNDSHEET
	 * BOFs, the lbPlyPos needs to change for all of them when record size
	 * changes.
	 */
	public static boolean getUpdatesAllBOFPositions()
	{
		return true;
	}

	@Override
	public void init()
	{
		if( originalsize == 0 )
		{
			originalsize = reclen;
		}
		Sst.init( this );
	}

	/**
	 * Initializes the sst as well as initializing the UnicodeStrings contained
	 * within
	 */
	public static void init( Sst sst )
	{
		sst.origsstlen = sst.getLength();
		sst.currbound = 0;
		sst.stringvector.clear();

		// init the string cache for fast access of init vals

		// get the row, col and ixfe information
		sst.cstTotal = ByteTools.readInt( sst.getByteAt( 0 ), sst.getByteAt( 1 ), sst.getByteAt( 2 ), sst.getByteAt( 3 ) );
		sst.cstUnique = ByteTools.readInt( sst.getByteAt( 4 ), sst.getByteAt( 5 ), sst.getByteAt( 6 ), sst.getByteAt( 7 ) );
		int strlen = 0;
		int strpos = 8;

		log.debug( "INFO: initializing Sst: " + sst.cstTotal + " total Strings, " + sst.cstUnique + " unique Strings." );
		// Initialize continues records
		sst.initContinues();

		// initialize the Unicodestrings from the byte array
		for( int d = 0; d < sst.cstUnique; d++ )
		{
			// Unicodestring values
			int numruns = 0;
			int runlen = 0;
			// the number of formatting runs each one adds 4 bytes
			int basereclen = 3; // the base length of the ustring being created
			int cchExtRst = 0; // the length of any Extended string data
			boolean doubleByte = false; // whether this is a double-byte string
			byte grbit = 0x0;
			// the grbit tells us what kind of Unicodestring this is
				log.trace( "Initializing String: " + String.valueOf( d ) + "/" + sst.cstTotal );
			// figure out the boundary offsets
			int offr = sst.boundaries.length;
			if( offr < 1 )
			{
				offr = 0;
			}
			else
			{
				offr = 1;
			}
			if( strpos >= sst.boundaries[sst.boundaries.length - offr] )
			{
				break;
			}
			short[] recdef = sst.getNextStringDefData( strpos );
			// get the length of the Unicodestring
			strlen = recdef[0];
			grbit = (byte) recdef[1];

			// we only want the bottom 4 bytes of the grbit, & not bit 2.. other
			// stuff is junk
			// grbit = (byte)(0xD & grbit);

			// init the string cache for fast access of init vals
			// st.initCacheBytes(strpos, 10); // commented out as it causes
			// array errors when short strings at end of continue boundary

			XLSRecord currec = sst;
				log.trace( "StrLen:" + strlen + " Strpos:" + strpos + " bound:" + sst.boundaries[sst.currbound] );

			if( strpos >= sst.boundaries[0] )
			{
				currec = sst.thiscont;
			}
			switch( grbit )
			{

				case 0x1: // non-rich, double-byte string
					doubleByte = true;
					break;

				case 0x4: // non-rich, single byte string
					cchExtRst = ByteTools.readInt( currec.getByteAt( strpos + 3 ),
					                               currec.getByteAt( strpos + 4 ),
					                               currec.getByteAt( strpos + 5 ),
					                               currec.getByteAt( strpos + 6 ) );
					basereclen = 7;
					doubleByte = false;
					break;

				case 0x5: // extended, non-rich, double-byte string
					cchExtRst = ByteTools.readInt( currec.getByteAt( strpos + 3 ),
					                               currec.getByteAt( strpos + 4 ),
					                               currec.getByteAt( strpos + 5 ),
					                               currec.getByteAt( strpos + 6 ) );
					basereclen = 7;
					doubleByte = true;
					break;

				case 0x8: // rich single-byte UNICODE string
					numruns = ByteTools.readShort( currec.getByteAt( strpos + 3 ), currec.getByteAt( strpos + 4 ) );
					runlen = numruns * 4;
					basereclen = 5;
					doubleByte = false;
					break;

				case 0x9: // rich double-byte UNICODE string
					numruns = ByteTools.readShort( currec.getByteAt( strpos + 3 ), currec.getByteAt( strpos + 4 ) );
					runlen = numruns * 4;
					basereclen = 5;
					doubleByte = true;
					break;

				case 0xc: // rich single-byte eastern string
					numruns = ByteTools.readShort( currec.getByteAt( strpos + 3 ), currec.getByteAt( strpos + 4 ) );
					cchExtRst = ByteTools.readInt( currec.getByteAt( strpos + 5 ),
					                               currec.getByteAt( strpos + 6 ),
					                               currec.getByteAt( strpos + 7 ),
					                               currec.getByteAt( strpos + 8 ) );
					runlen = numruns * 4;
					basereclen = 9;
					doubleByte = false;
					break;

				case 0xd: // rich double-byte eastern string
					numruns = ByteTools.readShort( currec.getByteAt( strpos + 3 ), currec.getByteAt( strpos + 4 ) );
					cchExtRst = ByteTools.readInt( currec.getByteAt( strpos + 5 ),
					                               currec.getByteAt( strpos + 6 ),
					                               currec.getByteAt( strpos + 7 ),
					                               currec.getByteAt( strpos + 8 ) );
					runlen = numruns * 4;
					basereclen = 9;
					doubleByte = true;
					break;

				default:
					doubleByte = false;
					cchExtRst = 0;
					basereclen = 3;
					if( grbit != 0x0 )
					{
						// if(st.DEBUGLEVEL > 10)
						log.warn( "ERROR: Invalid Unicodestring grbit:" + String.valueOf( grbit ) );
					}
			}
			// create the String
			if( strlen == 0 )
			{
					log.warn( "WARNING: Attempt to initialize Zero-length String." );
			}
			if( doubleByte )
			{
				strlen *= 2;
			}
			// it's a double-byte string so total size is *2
			try
			{
				strpos = sst.initUnicodeString( strlen, strpos, basereclen, cchExtRst, runlen, doubleByte );

				// Logger.logInfo("SST Currbound: " + sst.currbound +" strpos: "
				// + strpos);
				// if(st.DEBUGLEVEL > 5)Logger.logInfo("numruns: "
				// +String.valueOf(numruns)+" @"+String.valueOf(strpos)
				// +" len: " + String.valueOf(strlen) + " gr: "
				/// +String.valueOf(grbit) +
				// " base: "+String.valueOf(basereclen)+
				// " cchExtRst: "+String.valueOf(cchExtRst));
			}
			catch( Exception e )
			{
				log.warn( "ERROR: Error Reading String @ " + strpos + e.toString() + " Skipping..." );
				strpos += strlen + basereclen + runlen;
			}
		}
			log.debug( "Done reading SST." );
	}

	/**
	 * retrieves the sst string at the location pos and returns the next
	 * position
	 *
	 * @param ustrLen    actual unicode string length, (not including formatting runs,
	 *                   phonetic data or double byte multiplication)
	 * @param pos        position in the source data buffer
	 * @param ustrStart  start of unicode string within single sst record
	 * @param cchExtRst  phonetic data length
	 * @param runlen     fomratting run length
	 * @param doublebyte true if unicode string data is double byte (and then the size
	 *                   of the unicode string data array is ustrLen * 2)
	 * @return
	 */
	int initUnicodeString( int ustrLen, int pos, int ustrStart, int cchExtRst, int runlen, boolean doublebyte )
	{
		int bufferBoundary = boundaries[currbound]; // get the current boundary
		int totalStrLen = ustrStart + ustrLen + cchExtRst + runlen;        // calculate the total byte length of the unicode string
		int posEnd = pos + totalStrLen; // end position -- if > current record length must span and access next record/continues
		AtomicInteger uLen = new AtomicInteger( ustrLen ); // same as ustrLen but mutable in order to allow changing value in getData method

		// begin checking string against current record buffer boundary
		if( posEnd < bufferBoundary )
		{// string does not cross current boundary - easy! retrieve totalStringLen bytes and create unicode string
			byte[] newStringBytes = getData( uLen, pos, ustrStart, cchExtRst, runlen, doublebyte, false );
			initString( newStringBytes, pos, false );
			return posEnd;
		}
		if( posEnd == bufferBoundary )
		{// string is on the boundary - easy!
			if( (numconts == 0) || (numconts == contcounter) )
			{
					log.debug( "Last String in SST encountered." );
			}
			byte[] newStringBytes = getData( uLen, pos, ustrStart, cchExtRst, runlen, doublebyte, false );
			initString( newStringBytes, pos, false );

			/* "If fHighByte is 0x1 and rgb is extended with a Continue record the break
			   MUST occur at the double-byte character boundary."
			*/  // because we ended on a string, there is no grbit on the next continue
			if( continues.size() > currbound )
			{
				thiscont = (Continue) continues.get( currbound );
				currbound++;
				if( thiscont.getHasGrbit() )
				{
					thiscont.setHasGrbit( false );
					shiftBoundaries( 1 );
				}
			}
			return posEnd;
		}

		// spans or crosses the continue boundary
		byte[] newStringBytes = getData( uLen,
		                                 pos,
		                                 ustrStart,
		                                 cchExtRst,
		                                 runlen,
		                                 doublebyte,
		                                 true );    // retrieve the bytes, accounting for spanning (true)
		initString( newStringBytes, pos, false );
		return pos + (uLen.intValue() + ustrStart + cchExtRst + runlen);        // in most cases should be same as pos + totalStrLen but it's possible for uLen to be changed in getData
	}

	/**
	 * gets the sst string at strpos length allstrlen and returns the next
	 * position
	 *
	 * @param allstrlen
	 * @param strpos
	 * @param strend
	 * @param STATE
	 * @return
	 *
	 *
	 *         KSC: replacing int getString(int allstrlen, int strpos, int
	 *         strend, int[] STATE){ byte[] newStringBytes = getData(allstrlen,
	 *         strpos, STATE); int nextpos = strpos + allstrlen;
	 *         if(STATE[SI_SPANSTATE]==Sst.STATE_EXRSTSPAN){
	 *         this.initString(newStringBytes, strpos,true); }else{
	 *         this.initString(newStringBytes, strpos,false); }
	 *
	 *         if(strend != nextpos) Logger.logWarn(
	 *         "Sanity Check in Sst initUnicodeString(): strend != nextpos.");
	 *
	 *         return nextpos; }
	 */

	/**
	 * Adjust the boundary pointers based on whether we need to compensate for
	 * grbit anomalies
	 *
	 *
	 * NOT USED
	 *
	 void shiftBoundariesX(int x) {
	 int ct = 0;
	 Iterator it = continues.iterator();
	 while (it.hasNext()) {
	 Continue nextcont = (Continue) it.next();
	 if (ct++ >= this.currbound) {
	 nextcont.setContinueOffset(nextcont.getContinueOffset() + x);
	 boundaries[ct] = nextcont.getContinueOffset();
	 if (DEBUGLEVEL > -5)
	 Logger.logInfo("Sst.shiftBoundaries() Updated " + nextcont
	 + " : " + nextcont.getContinueOffset());
	 }
	 }
	 if (boundaries.length == (this.continues.size() + 1)) {
	 boundaries[this.continues.size()] += x;
	 }
	 }*/

	/**
	 * Adjust the boundary pointers based on whether we need to compensate for
	 * grbit anomalies
	 */
	void shiftBoundaries( int x )
	{
		// int ret = 0;
		for( int t = currbound; t < continues.size(); t++ )
		{
			Continue nextcont = (Continue) continues.get( t );
			nextcont.setContinueOffset( nextcont.getContinueOffset() + x );
			boundaries[t] = nextcont.getContinueOffset();
				log.debug( "Sst.shiftBoundaries() Updated " + nextcont + " : " + nextcont.getContinueOffset() );
		}
		if( boundaries.length == (continues.size() + 1) )
		{
			boundaries[continues.size()] += x;
		}
	}

	/**
	 * Refactoring Continue data access
	 *
	 * @param i
	 * @return
	 */
	short[] getNextStringDefData( int start )
	{
		short[] ret = { (short) 0x0, (short) 0x0 };
		try
		{
			// int thiscont = -1;
			int end = start + 3;
			if( end <= boundaries[0] )
			{ // it's in the main Sst data
				ret[0] = ByteTools.readShort( getByteAt( start++ ), getByteAt( start++ ) );
				ret[1] = (short) getByteAt( start );
				return ret;
			}

			// KSC: no need as Sst.getData increments correctly ... 
			// this.thiscont = this.getContinue(end);
			byte b0 = thiscont.getByteAt( start++ );
			byte b1 = thiscont.getByteAt( start++ );
			byte b2 = thiscont.getByteAt( start++ );
			ret[0] = ByteTools.readShort( b0, b1 );
			ret[1] = (short) b2;
		}
		catch( Exception e )
		{
				log.warn( "possible problem parsing String table getting next string def data: " + e, e );
		}
		return ret;
	}

	/**
	 * return the continue that contains up to t length
	 *
	 * @param t
	 * @return
	 */
	Continue getContinue( int t )
	{
		if( (t - 1) == datalen )
		{
			return (Continue) continues.get( continues.size() - 1 );
		}
		for( int x = boundaries.length - 1; x >= 0; x-- )
		{
			if( t > boundaries[x] )
			{
				return (Continue) continues.get( x );
			}
		}
		return null;
	}

	/**
	 * get the string data from the proper place (either this sst record, or one or more continues)
	 * <p/>
	 * <br>
	 * NOTES: if the string spans a continue, the length in bytes of each part
	 * is contained in the first two ints.
	 * <p/>
	 * if the record spans a Continue, we need to see if the border falls within
	 * text data or extra data
	 * <p/>
	 * 10 len, really 15 bytes uncomp ||gr comp 2,0,1,0,2,0,3,0,3,0,||0
	 * ,4,5,5,7,8
	 *
	 * @param ustrLen    unicode string length
	 * @param pos        position in buffer
	 * @param ustrStart  start of unicode string part (after initial length(s))
	 * @param cchExtRst  phonetic data length or 0 if none
	 * @param runlen     formatting runs length or 0 if none
	 * @param doublebyte true of doublebyte
	 * @param bSpans     true if spans records
	 * @return byte[] defining unicode string
	 */
	byte[] getData( AtomicInteger ustrLen, int pos, int ustrStart, int cchExtRst, int runlen, boolean doublebyte, boolean bSpans )
	{
		int totalStrLen = ustrStart + ustrLen.intValue() + cchExtRst + runlen;
		int posEnd = pos + totalStrLen; // buffer end position

		if( posEnd <= boundaries[0] )
		{ // it's in the main Sst data just grab string and return
			return getBytesAt( pos, totalStrLen );
		}

		// if it's in the current continues without spanning, just get the bytes and return
		if( !bSpans )
		{ // Simple -- no Span, return data
			pos += thiscont.grbitoff;
			int thisoff = pos - thiscont.getContinueOffset();
			return thiscont.getBytesAt( thisoff, totalStrLen );
		}

		// if string spans two or more records, must deal with boundaries and grbits and lots of complications ...
		int bufferBoundary = boundaries[currbound]; // get the current boundary
			log.debug( "Crossing Boundary: " + bufferBoundary + ".  Double-Bytes: " + doublebyte );

		// get ensuing record (previous==thiscont.predecessor)
		if( (currbound) < continues.size() )
		{
			thiscont = (Continue) continues.get( currbound++ );
		}

		// find out where break is
		int currpos = pos + totalStrLen;
		boolean bfoundBreak = false;
		boolean bUnCompress = false; // true if string on previous boundary must be uncompressed
		boolean bUnCompress1 = false; // true if string1 must be uncompressed ** this one is confusing but works **

		// check if break is in ExtRst data (==phonetic data)
		if( cchExtRst > 0 )
		{
			currpos -= cchExtRst;
			if( currpos <= bufferBoundary )
			{
					log.debug( "Continue Boundary in ExtRst data." );
				if( thiscont.getHasGrbit() )
				{
					thiscont.setHasGrbit( false );
					shiftBoundaries( 1 );
				}
				bfoundBreak = true;
			}
		}

		// check if break is in formatting run data
		if( runlen > 0 )
		{
			currpos -= runlen;
			if( !bfoundBreak && (currpos <= bufferBoundary) )
			{ // check against japanese!
					log.debug( "Continue Boundary in Formatting Run data." );
				if( thiscont.getHasGrbit() )
				{
					shiftBoundaries( 1 );
					thiscont.setHasGrbit( false );
				}
				bfoundBreak = true;
			}
		}

		// otherwise the break is in unicode stringdata part
		currpos = pos + ustrStart;
		if( !bfoundBreak && (currpos < bufferBoundary) )
		{
			if( ustrLen.intValue() == 0 )
			{ // a ONE BYTE String on the boundary! Add the grbit back to the Continue
					log.debug( "1 byte length String on the Continue Boundary." );
				boundaries[boundaries.length - 1]++; // increment the last boundary...
			}
		}

		// check if break is within the actual ustring data
		if( ((currpos <= bufferBoundary) && ((currpos + ustrLen.intValue()) > bufferBoundary)) )
		{ // is break within String portion// ?
				log.debug( "Continue Boundary in String data." );
			if( !thiscont.getHasGrbit() )
			{ // when does this happen???
				thiscont.setHasGrbit( true );
				shiftBoundaries( -1 );
			}

			byte b = thiscont.getGrbit();
			// If it changes double --> single or single --> double then adjust accordingly (plus set bUnCompress or bUnCompress1 flags which govern how bytes are accessed)
			if( doublebyte && (b == 0x0) )
			{ // it is in doublebytes but it really should be compressed
				int preBoundaryBytes = bufferBoundary - pos - ustrStart;
				int postBoundaryBytes = (pos + ustrStart + ustrLen.intValue()) - bufferBoundary;
				postBoundaryBytes = postBoundaryBytes / 2;
				ustrLen.set( preBoundaryBytes + postBoundaryBytes );
				bUnCompress1 = true; // dunno what this means but it works ...
			}
			else if( !doublebyte && (b == 0x1) )
			{ // string portion on prevouos boundary should be uncompressed/converted to doublebyte
				int preBoundaryBytes = bufferBoundary - pos - ustrStart;
				int postBoundaryBytes = (pos + ustrStart + ustrLen.intValue()) - bufferBoundary;
				postBoundaryBytes = postBoundaryBytes * 2;
				ustrLen.set( preBoundaryBytes + postBoundaryBytes );
				bUnCompress = true;
			}
			// ustrLen may have changed above - reset vars
			totalStrLen = ustrStart + ustrLen.intValue() + cchExtRst + runlen;
			posEnd = pos + totalStrLen;
		}

		// calculate length on current record (=thiscont.predecessor) and length on ensuing continues
		int string1ByteLength = pos - thiscont.getContinueOffset(); // bytes on 1st record or continue
		if( string1ByteLength < 0 )
		{
			string1ByteLength *= -1;
		}
		int string2ByteLength = totalStrLen - string1ByteLength; // bytes on 2nd// or// ensuing// continues// ==// spanned// bytes
		int extraData = (cchExtRst + runlen); // non-string-data (phonetic info// and/or formatting runs)

		// remove ExtRst and runlen info from 2nd String length
		string2ByteLength -= extraData;
		if( string2ByteLength <= 0 )
		{
			// if it spans we want extr just to be the bytes on the second continue
			extraData = string2ByteLength + extraData;
			string2ByteLength = 0;// all String data is contained in prior // Continue
		}

		byte[] string1bytes = null;
		byte[] string2bytes = null;

		// we need to expand the first section bytes to fit the last one (????)
		if( thiscont.predecessor instanceof Continue )
		{
			pos = ((thiscont.predecessor.getLength()) - (string1ByteLength));
			pos -= 4;
		}
		if( thiscont.getHasGrbit() )
		{
			thiscont.grbitoff = 1;
		}
		else
		{
			thiscont.grbitoff = 0;
		}

		// *********************************************************************************************************
		// handle the part of unicode string on previous continues
		if( !bUnCompress )
		{
			string1bytes = thiscont.predecessor.getBytesAt( pos, string1ByteLength );
		}
		else
		{ // portion of string on previous boundary is singlebyte; ensuing portion is doublebyte; must convert previous to doublebyte
			string1bytes = convertCompressedBytesToDoubleBytes( pos, string1ByteLength, ustrStart );
		}

		// *********************************************************************************************************
		// handle part on current (and ensuing, if necessary) continues
		if( string2ByteLength < MAXRECLEN )
		{ // 99.9% usual case
			if( !bUnCompress1 )
			{
				string2bytes = thiscont.getBytesAt( 0 + thiscont.grbitoff, string2ByteLength );
			}
			else
			{ // Expand the second string bytes
				string2ByteLength *= 2;
				string2bytes = new byte[string2ByteLength];
				for( int t = 0; t < (string2ByteLength / 2); t++ )
				{
					string2bytes[(t * 2)] = thiscont.getByteAt( t + thiscont.getContinueOffset() );
				}
			}
			// since we've accessed the last bytes of the prior Continue, blow
			// it out!
			if( thiscont.predecessor instanceof Continue )
			{
				thiscont.predecessor.setData( null );
			}
		}
		else
		{ // string2ByteLength spans continue(s) ************************************************* see infoteria/cannotread824315.xls
			int blen = string2ByteLength;
			int idx = 0;
			int start = 0;
			string2bytes = new byte[blen];
			while( blen > 0 )
			{                // loop thru ensuing continues until correct length is read in
				int curlen = Math.min( (start + thiscont.getLength()) - thiscont.grbitoff, blen );
				if( !bUnCompress1 )
				{
					byte[] tmp = thiscont.getBytesAt( start + thiscont.grbitoff, curlen );
					System.arraycopy( tmp, 0, string2bytes, idx, curlen );
				}
				else
				{ // Expand the second string bytes - NOTE: This has not been hit so hasn't been tested ...
					curlen *= 2;
					byte[] tmp = new byte[curlen];
					for( int t = 0; t < (curlen / 2); t++ )
					{
						tmp[(t * 2)] = thiscont.getByteAt( t + thiscont.getContinueOffset() );
					}
					System.arraycopy( tmp, 0, string2bytes, idx, curlen );
				}
				// since we've accessed the last bytes of the prior Continue, blow it out!
				if( thiscont.predecessor instanceof Continue )
				{
					thiscont.predecessor.setData( null );
				}
				if( curlen >= (thiscont.getLength() - thiscont.grbitoff) )
				{ // finished this one, get next continue
					if( (currbound) < continues.size() )
					{
						thiscont = (Continue) continues.get( currbound++ );
					}
					if( thiscont.getHasGrbit() )
					{
						thiscont.grbitoff = 1;
					}
					else
					{
						thiscont.grbitoff = 0;
					}
					start = 4;        // don't understand this but, hey, it works ...
				}
				else
				{    // we are done
					// get current length in current continues only (==start postion for extra data, if any)
					string2ByteLength = curlen + start;
					break;
				}
				idx += curlen;
				blen -= curlen;
			}
		}

		// ***********************************************************************************************************
		// now put together the string bytes - string1bytes and string2bytes are the ustring only, excluding extraData (formatting or phoentic data)
		byte[] returnstringbytes = new byte[string1bytes.length + string2bytes.length + extraData];
		System.arraycopy( string1bytes, 0, returnstringbytes, 0, string1bytes.length );
		System.arraycopy( string2bytes, 0, returnstringbytes, string1bytes.length, string2bytes.length );

		// does it have ExtRst or Formatting Runs?
		if( extraData > 0 )
		{
			if( posEnd <= boundaries[currbound] )
			{ // usual case!!
				int startpos = string2ByteLength;
				if( bUnCompress1 )
				{
					startpos /= 2;
				}
				if( string2ByteLength != 0 )
				{
					startpos += thiscont.grbitoff;        // dunno why but it works ...
				}
				byte[] rx2 = thiscont.getBytesAt( startpos, extraData );
				System.arraycopy( rx2, 0, returnstringbytes, (string1bytes.length + string2bytes.length), extraData );
			}
			else
			{    // extraData spans continues ...
				// have to get portion on prev continue and rest on next continue ... sigh ....
				int startpos = string2ByteLength;
				if( bUnCompress1 )
				{
					startpos /= 2;
				}
				byte[] rx2 = new byte[extraData];
				string1ByteLength = thiscont.getLength() - startpos; // bytes on 1st record or continue
				string2ByteLength = extraData - string1ByteLength; // bytes on 2nd// or// ensuing// continues// ==// spanned// bytes
				if( (currbound) < continues.size() )
				{
					thiscont = (Continue) continues.get( currbound++ );
					if( thiscont.getHasGrbit() )
					{
						thiscont.grbitoff = 1;
					}
					else
					{
						thiscont.grbitoff = 0;
					}
				}
				pos = ((thiscont.predecessor.getLength()) - (string1ByteLength));
				pos -= 4;
				int start = 4;    // why?????
				System.arraycopy( thiscont.predecessor.getBytesAt( pos, string1ByteLength ), 0, rx2, 0, string1ByteLength );
				System.arraycopy( thiscont.getBytesAt( start + thiscont.grbitoff, string2ByteLength ),
				                  0,
				                  rx2,
				                  string1ByteLength,
				                  string2ByteLength );
				System.arraycopy( rx2, 0, returnstringbytes, (string1bytes.length + string2bytes.length), extraData );
				ustrLen.set( ustrLen.get() - 1 );    // ???? DO NOT UNDERSTAND THIS BUT IT APPEARS TO WORK - hits on infoteria/cannotread824315.xls
			}
		}

			log.trace( "Total Length from Continue: " + returnstringbytes.length );
		return returnstringbytes;
	}

	/**
	 * for rare occurrences where string portion on previous boundary is flagged singlebyte/compressed,
	 * and the ensuing continue is flagged doublebyte; in these cases the unicode-string-portion on the
	 * previous boundary must be converted to doublebyte
	 *
	 * @param pos       positon on previous boundary
	 * @param totallen  total string length on previous boudary
	 * @param uStrStart start of unicode string portion
	 * @return
	 */
	byte[] convertCompressedBytesToDoubleBytes( int pos, int totallen, int uStrStart )
	{
		int uLenOnPrevious = (totallen - uStrStart);    // unicode string portion on previous boundary
		byte[] converted = new byte[uStrStart + (uLenOnPrevious * 2)];
		System.arraycopy( thiscont.predecessor.getBytesAt( pos, uStrStart ), 0, converted, 0, uStrStart );
		byte[] ustr = thiscont.predecessor.getBytesAt( pos + uStrStart, uLenOnPrevious );
		converted[2] = (byte) (converted[2] | 0x1);    // flag as doublebyte/uncompressed for unicode string processing
		for( int i = 0; i < uLenOnPrevious; i++ )
		{    // copy rest of unicode string portion on prev boundary as doublebyte
			converted[uStrStart + (i * 2)] = ustr[i];
		}
		return converted;

	}

	/**
	 * given unicode bytes, create a Unicodestring and add it to the string vector
	 */
	Unicodestring initString( byte[] newStringBytes, int strpos, boolean extrstbrk )
	{
		// create a new Unicodestring, set its data
		Unicodestring newString = new Unicodestring();
		newString.setSSTPos( strpos );
		newString.init( newStringBytes, extrstbrk );

		// add the new String to the String table and return the new pointer
			log.trace( " val: " + newString.toString() );
		if( newString.getLen() == 0 )
		{
			log.trace( "Adding zero-length string!" );
		}
		else
		{
			putString( newString );
		}
		return newString;
	}

	int retpos = -1;

	private int putString( Unicodestring newString )
	{
		++retpos;
		((SstArrayList) stringvector).put( newString, retpos );
		return retpos;
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	@Override
	public void close()
	{
		cbounds.removeAllElements();
		sstgrbits.removeAllElements();
		stringvector.clear();
		stringvector = new SstArrayList();
		dupeSstEntries.clear();
		dupeSstEntries = new HashSet();
		existingSstEntries.clear();
		existingSstEntries = new HashSet();
	}

	/**
	 * call this method after changing the value of an SST Unicode string to
	 * update the underlying SST byte array.
	 */
	void updateUnicodestrings()
	{
		// TODO: OPTIMIZE: check that the Sst has changed
		// reset defaults
		cbounds = new CompatibleVector();
		sstgrbits = new CompatibleVector();
		lastwasbreakable = true;
		stringisonbound = false;
		laststringwasonbound = false;
		islast = false;
		thisbounds = WorkBookFactory.MAXRECLEN;
		lastbounds = 0;
		contcounter = 0;
		lastlen = 0;
		grbitct = 0;
		dl = 0;
		leftoverlen = 0;
		gr = 0x0;

		// loop through the strings and copy their
		// bytes to the SST byte array.
		// byte[] tmp = new byte[0];
		byte[] cstot = ByteTools.cLongToLEBytes( cstTotal );
		byte[] cstun = ByteTools.cLongToLEBytes( cstUnique );

		// TODO: OPTIMIZE!!
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		try
		{
			out.write( cstot );
			out.write( cstun );
		}
		catch( IOException e )
		{
			log.warn( "Exception getting String bytes: " + e , e);
		}

		if( stringvector.size() > 0 )
		{
			// now get the continue boundaries
			int thispos = 8;
			int lastpos = 0;
			cbounds.removeAllElements();
			sstgrbits.removeAllElements();
			byte[] strb = null;
			Iterator it = stringvector.iterator();
			while( it.hasNext() )
			{
				Object ob = it.next();
				Unicodestring str = (Unicodestring) ob;

				// from updateUnicodeStrings()
				str.setSSTPos( thispos );

				strb = str.read();
				try
				{
					out.write( strb );
				}
				catch( IOException e )
				{
					log.warn( "Exception getting String bytes: " + e , e);
				}
				lastpos = thispos;
				thispos = lastpos + strb.length;
				checkOnBoundary( str, lastpos, thispos + 4, strb ); // add 4
				// because 4
				// added to
				// boundaries
			}

			if( leftoverlen > 0 )
			{// there was leftover data!
				cbounds.add( leftoverlen );
				sstgrbits.add( gr );
			}
			if( cbounds.size() > 0 )
			{
				numconts = cbounds.size() - 1;
			}
			else
			{
				numconts = 0;
			}

			byte[] bb = out.toByteArray();
			if( sanityCheck( bb.length ) )
			{
				setData( bb );
			}
			else
			{
				datalen = bb.length + 4;
				updateUnicodestrings();
			}

		}
	}

	/**
	 * Checks that the continues make sense. In some cases the continue lengths
	 * will be wrong due to the datalen being off. Datalen is created via
	 * offsets, not absolutes, so this can occur. If so, we will reset the
	 * datalen to the correct number.
	 *
	 * @return
	 */
	private boolean sanityCheck( int realLen )
	{
		long contLens = 0;
		for( Object cbound : cbounds )
		{
			Integer intgr = (Integer) cbound;
			contLens += intgr;
		}
		if( ((datalen - 4) - contLens) > 8223 )
		{
				log.warn( "SST continue lengths not correct, regenerating" );
			return false;
		}
		return true;
	}

	CompatibleVector cbounds = new CompatibleVector();
	CompatibleVector sstgrbits = new CompatibleVector();
	boolean lastwasbreakable = true;
	boolean stringisonbound = false;
	boolean laststringwasonbound = false;
	boolean islast = false;
	int thisbounds = WorkBookFactory.MAXRECLEN;
	int lastbounds = 0;
	int contcounter = 0;
	int lastlen = 0;
	int grbitct = 0;
	int dl = 0;
	int leftoverlen = 0;
	byte gr = 0x0;

	/*
	 * handle the checking of Continue boundary strings
	 */ int stringnumber = 0;
	int continuenumber = -1;
	int lastgrbit = 0;

	void checkOnBoundary( Unicodestring str, int lastpos, int thispos, byte[] strb )
	{
		while( thispos >= (thisbounds) )
		{
			continuenumber++;
				log.info( String.valueOf( thisbounds ) );

			// check whether the string can safely be split
			boolean breaksok = str.isBreakable( thisbounds );
			int contlen = 0;

			// get the Continue grbit
			gr = strb[0]; // default is a non-grbit -- if it doesn't break, we
			// don't want it

			if( (breaksok) )
			{
				gr = getContinueGrbitFromString( str );
			}
				log.debug( " String @: " + thispos + " is breakable: " + breaksok );

			// deal with string break subtleties
			contlen = WorkBookFactory.MAXRECLEN; // the default
			if( islast )
			{
				contlen = leftoverlen;
				leftoverlen = 0;
				contlen++;
				if( !lastwasbreakable )
				{
					contlen++;
				}
			}
			else if( !breaksok )
			{
				stringisonbound = true; // we are
				contlen = lastpos - lastbounds;
				lastbounds = lastpos;
			}
			else
			{
				// check if it's double byte, if so, make sure that the break is
				// not in the middle of a character.
				if( breaksok && (gr == 1) )
				{
					if( str.charBreakOnBounds( thisbounds + lastgrbit ) )
					{
						contlen--;
					}
				}
			}

			// set continue length
			if( (!laststringwasonbound) && lastwasbreakable && (contcounter > 0) )
			{ // normal w/grbit
				if( !breaksok )
				{
					cbounds.add( contlen );
				}
				else
				{
					if( !islast )
					{
						thisbounds--;
					}
					cbounds.add( contlen - 1 );
				}
			}
			else
			{
				cbounds.add( contlen );
			}

			// set grbit add null if the Continue should not have a grbit
			if( str.cch < 2 )
			{
				sstgrbits.add( null );
				lastgrbit = 0;
			}
			else if( !breaksok && ((gr < 0x2) && (gr >= 0x0)) )
			{
				sstgrbits.add( null );
				lastgrbit = 0;
			}
			else
			{
				sstgrbits.add( gr );
				lastgrbit = 1;
			}

			contcounter++;

			// reset stuff
			lastwasbreakable = breaksok;
			laststringwasonbound = stringisonbound;
			stringisonbound = false;

			lastlen = contlen;
			if( breaksok )
			{
				lastbounds = thisbounds;
			}

			// datalen will be smaller than reclen
			// if continues were not created
			if( reclen > datalen )
			{
				dl = reclen;
			}
			else
			{
				dl = datalen;
			}
			// 20060518 KSC: handle segments that fall between the extra 4 added
			// to the boundary ...
			if( (thisbounds + contlen + 4) < dl )
			{ // not the last one
				thisbounds += contlen;
				lastpos += contlen; // 20090407 KSC: If !breaksok but still
				// loops, must increment lastpos or infinite
				// loops [BUGTRACKER 2355 Infoteria OOM]
			}
			else if( !islast )
			{
				leftoverlen = dl + 4 - lastlen;
				if( !lastwasbreakable && (leftoverlen > 0) )
				{
					leftoverlen++;
				}
				thisbounds = dl;
				islast = true;
			}
			else
			{
				thisbounds += contlen;
			}

		}
	}

	/**
	 * This returns the Continue record grbit which is either 0 or 1 -- NOT the
	 * string's grbit which determines much more...
	 */
	static byte getContinueGrbitFromString( Unicodestring str )
	{
		byte grb = 0x0;
		switch( str.getGrbit() )
		{
			case 0x1:
				grb = 0x1;
				break;
			case 0x5:
				grb = 0x1;
				break;
			case 0x9:
				grb = 0x1;
				break;
			case 0xd:
				grb = 0x1;
				break;
			default:
		}
		return grb;
	}

	Object[] continueDef;

	/**
	 * return the sizes of Continue records for an Sst caches the read if
	 * neccesary
	 */
	public static Object[] getContinueDef( Sst rec, boolean cached )
	{
		if( cached )
		{
			return rec.continueDef;
		}
		Integer[] cbs = new Integer[rec.cbounds.size()];
		Byte[] sstgrs = new Byte[rec.sstgrbits.size()];
		for( int t = 0; t < cbs.length; t++ )
		{
			cbs[t] = ((Integer) rec.cbounds.get( t ));
		}

		for( int t = 0; t < sstgrs.length; t++ )
		{
			sstgrs[t] = ((Byte) rec.sstgrbits.get( t ));
		}

		rec.continueDef = new Object[2];
		rec.continueDef[0] = cbs;
		rec.continueDef[1] = sstgrs;
		return rec.continueDef;

	}

	/**
	 * Called from LabelSst on initialization from a new workbook, this
	 * pre-populates the list of strings that are currently shared.
	 */
	void initSharingOnStrings( int isst )
	{
		Integer iSst = isst;
		if( existingSstEntries.contains( iSst ) )
		{
			if( !dupeSstEntries.contains( iSst ) )
			{ // really is just a switch -doesn't track # times string is shared ...
				dupeSstEntries.add( iSst );
			}
		}
		else
		{
			existingSstEntries.add( iSst );
		}
	}

	// Optimization -- don't check UStr on add
	int STRING_ENCODING_MODE = Sst.STRING_ENCODING_UNICODE;

	public void setStringEncodingMode( int mode )
	{
		STRING_ENCODING_MODE = mode;
	}

	/**
	 * remove a Unicodestring from the table
	 *
	 * @param idx
	 */
	void removeUnicodestring( Unicodestring str )
	{
		stringvector.remove( idx );
		retpos--;
		reclen -= str.getLen();
	}

	/**
	 * used when modifying existing sst entry update data + rec lens
	 *
	 * @param delta amt of adjustment
	 */
	void adjustSstLength( int delta )
	{
		reclen += delta;
		datalen += delta;
	}

	/**
	 * insert a new Unicodestring into the array of strings composing this
	 * String Table
	 */
	int insertUnicodestring( Unicodestring us )
	{
		int retpos = -1;
		cstTotal++;
		boolean isuni = false;
		// get the existing position of this string
		// but only if we're not ignoring dupes
		if( getWorkBook().isSharedupes() )
		{
			retpos = ((SstArrayList) stringvector).find( us ); // indexOf will not
			// match entire
			// unicode
			// string
			// (including
			// formatting)
		}
		if( retpos == -1 )
		{ // unicode string isn't in yet
			cstUnique++;
			int strlen = us.getLen();
			reclen += strlen + (us.isRichString() ? 5 : 3);
			datalen += strlen + (us.isRichString() ? 5 : 3);
			if( isuni )
			{
				reclen += strlen; // utf double encoding.
				datalen += strlen;
			}
			retpos = putString( us );
		}
		else
		{
			// this is a duplicate string, track it!
			dupeSstEntries.add( retpos );
		}

		return retpos;
	}

	/**
	 * create a new unicode string from string and formatting information and
	 * add it to the Sst string array formatting runs, if present, contain list
	 * of short[] {char index, font index} where char index is start index in
	 * the string to apply font at font index
	 *
	 * @param s
	 * @param formattingRuns
	 * @return
	 */
	int addUnicodestring( String s, ArrayList formattingRuns )
	{
		cstTotal++;
		cstUnique++;
		Unicodestring str = createUnicodeString( s, formattingRuns, STRING_ENCODING_MODE );

		reclen += str.getLen();
		datalen += str.getLen();
		retpos = putString( str );
		return retpos;
	}

	/**
	 * Create a unicode string
	 *
	 * @param s
	 * @param formattingRuns
	 * @param ENCODINGMODE
	 * @return
	 */
	public static Unicodestring createUnicodeString( String s, ArrayList formattingRuns, int ENCODINGMODE )
	{
		try
		{
			boolean isuni = false;
			if( ENCODINGMODE == WorkBook.STRING_ENCODING_AUTO )
			{
				isuni = ByteTools.isUnicode( s );
			}
			else if( ENCODINGMODE == WorkBook.STRING_ENCODING_COMPRESSED )
			{
				isuni = false;
			}
			else if( ENCODINGMODE == WorkBook.STRING_ENCODING_UNICODE )
			{
				isuni = true;
			}
			if( formattingRuns != null )
			{
				isuni = true;
			}

			byte[] charbytes = s.getBytes( XLSConstants.DEFAULTENCODING );
			int strlen = charbytes.length; // .length();
			byte[] strbytes = null;

			// handle string sizes
			if( (strlen * 2) > Short.MAX_VALUE )
			{
				isuni = false; // can't fit larger than Short String length
			}
			if( strlen > (Short.MAX_VALUE - 3) )
			{ // if strlen is greater than
				// the maximum value for
				// excel cells, truncate
				strlen = Short.MAX_VALUE - 3; // maximum value
				charbytes = new byte[strlen];
				System.arraycopy( s.getBytes( XLSConstants.DEFAULTENCODING ), 0, charbytes, 0, strlen );
			}

			if( formattingRuns != null )
			{
				isuni = true;
			}

			if( isuni )
			{ // encode string bytes
				try
				{// if you use a string here for the encoding rather than a
					// reference to a static String, performance in JDK 4.2
					// will suffer. Why? Dunno, but it's bad!
					charbytes = s.getBytes( WorkBook.UNICODEENCODING );
				}
				catch( UnsupportedEncodingException e )
				{
					log.warn( "error encoding string: " + e + " with default encoding 'UnicodeLittleUnmarked'", e );
				}
				if( formattingRuns == null )
				{
					strbytes = new byte[charbytes.length + 3];
				}
				else
				{
					strbytes = new byte[charbytes.length + 5]; // need 2 extra
				}
				// bytes to
				// store
				// formatting
				// run info
			}
			else
			{
				strbytes = new byte[charbytes.length + 3];
			}

			// given info, create strbytes for Unicode init
			int pos = 0;
			int encodedlen = charbytes.length;
			byte[] lenbytes = ByteTools.shortToLEBytes( (short) strlen );
			strbytes[pos++] = lenbytes[0]; // cch bytes 0 & 1
			strbytes[pos++] = lenbytes[1];
			if( !isuni )
			{
				strbytes[pos++] = 0x0; // grbit byte 2
			}
			else
			{
				strbytes[pos++] = 0x1; // grbit byte 2
				if( formattingRuns != null )
				{ //
					strbytes[pos - 1] |= 0x8; // set Rich Text attribute
					byte[] fr = ByteTools.shortToLEBytes( (short) formattingRuns.size() );
					strbytes[pos++] = fr[0]; // # formatting runs bytes 3 & 4
					strbytes[pos++] = fr[1];
				}
			}
			System.arraycopy( charbytes, 0, strbytes, pos, encodedlen );

			if( formattingRuns != null )
			{
				// formatting runs (charindex, fontindex)*n after string data
				byte[] frs = new byte[(formattingRuns.size() * 4)];
				for( int i = 0; i < formattingRuns.size(); i++ )
				{
					short[] o = (short[]) formattingRuns.get( i );
					byte[] charIndex = ByteTools.shortToLEBytes( o[0] );
					byte[] fontIndex = ByteTools.shortToLEBytes( o[1] );
					System.arraycopy( charIndex, 0, frs, (i * 4), 2 );
					System.arraycopy( fontIndex, 0, frs, (i * 4) + 2, 2 );
				}
				// Append frs to end of strbytes
				byte[] newdata = new byte[strbytes.length + frs.length];
				System.arraycopy( strbytes, 0, newdata, 0, strbytes.length );
				System.arraycopy( frs, 0, newdata, strbytes.length, frs.length );
				strbytes = newdata;
			}
			// create a new one, set its data
			Unicodestring str = new Unicodestring();
			str.init( strbytes, false );
			return str;
		}
		catch( UnsupportedEncodingException e )
		{
			log.warn( "error encoding string: " + e.toString() , e);
		}
		return null;
	}

	/**
	 * insert a new Unicodestring into the array of strings composing this
	 * String Table
	 */
	int insertUnicodestring( String s )
	{
		int retpos = -1;
		// get the existing position of this string
		// but only if we're not ignoring dupes
		if( getWorkBook().isSharedupes() )
		{
			retpos = stringvector.indexOf( s );
			if( retpos > -1 )
			{
				Unicodestring str = (Unicodestring) stringvector.get( retpos );
				if( str.hasFormattingRuns() )
				{
					retpos = -1; // do not match if there are formatting runs
				}
				// embedded
			}
		}

		if( retpos == -1 )
		{ // it's a new string
			retpos = addUnicodestring( s, null ); // add with no formatting
			// information
		}
		else
		{
			cstTotal++;
			// this is a duplicate string, track it!
			dupeSstEntries.add( retpos );
		}

		return retpos;
	}

	/**
	 * Determine if the isst passed in is for a duplicate string or not.
	 */
	boolean isSharedString( int sstLoc )
	{
		if( dupeSstEntries.contains( Integer.valueOf( sstLoc ) ) )
		{
			return true;
		}
		return false;
	}

	/**
	 * Return the Unicodestring at the corresponding index
	 */
	Unicodestring getUStringAt( int i )
	{
		return (Unicodestring) stringvector.get( i );
	}

	/**
	 * find this unicode string (including formatting) in stringarray
	 *
	 * @param us
	 * @return
	 */
	int find( Unicodestring us )
	{
		return ((SstArrayList) stringvector).find( us );
	}

	/**
	 * Returns the String vector
	 */
	public List getStringVector()
	{
		return stringvector;
	}

	/**
	 * return the total number of strings in the SST
	 *
	 * @return
	 */
	public int getNumTotal()
	{
		return cstTotal;
	}

	/**
	 * return the number of unique strings in the SST
	 *
	 * @return
	 */
	public int getNumUnique()
	{
		return cstUnique;
	}

	/**
	 * return # continues
	 *
	 * @return
	 */
	public int getNumContinues()
	{
		return numconts;
	}

	/**
	 * we need to override stream to update changes to the byte array
	 */
	@Override
	public void preStream()
	{
		updateUnicodestrings();
	}

	// For debugging purposes
	public String toString()
	{
		StringBuffer sb = new StringBuffer();
		sb.append( "cstTotal:" + cstTotal + " cstUnique:" + cstUnique + " numConts:" + numconts );
		for( Object aStringvector : stringvector )
		{
			sb.append( "\n " + aStringvector );
		}
		return sb.toString();
	}

	/**
	 * Override ArrayList to allow matching based on .toString. Required because
	 * we call ArrayList.indexOf(String) when ArrayList contains UnicodeStrings.
	 */
	private class SstArrayList extends ArrayList
	{
		/**
		 * serialVersionUID
		 */
		private static final long serialVersionUID = 7904551471519095640L;
		private HashMap container = new HashMap();

		public boolean put( Object o, Integer isst )
		{
			container.put( ((Unicodestring) o).toCachingString(), isst );
			return super.add( o );
		}

		@Override
		public int indexOf( Object o )
		{
			Object oo = container.get( o.toString() );
			if( oo == null )
			{
				return -1;
			}
			return (Integer) oo;
		}

		@Override
		public boolean remove( Object o )
		{
			log.warn( "String being removed from SST array, Indexing may be off" );
			container.remove( ((Unicodestring) o).toCachingString() );
			return super.remove( o );
		}

		/**
		 * find this particular unicode string, including formatting
		 *
		 * @param us
		 * @return
		 */
		public int find( Unicodestring us )
		{
			return (super.indexOf( us ));
		}
	}

	/**
	 * generate the OOXML necessary to describe this string table, also fill
	 * sststrings list with unique sststrings
	 *
	 * @param sststrings
	 * @return sstooxml
	 * @throws IOException
	 */
	public void writeOOXML( Writer zip ) throws IOException
	{
		StringBuffer sstooxml = new StringBuffer();

		zip.write( OOXMLConstants.xmlHeader );
		zip.write( "\r\n" );
		zip.write( ("<sst xmlns=\"" + OOXMLConstants.xmlns + "\" count=\"" + cstTotal + "\" uniqueCount=\"" + cstUnique + "\">") );
		zip.write( "\r\n" );
		for( int i = 0; i < getStringVector().size(); i++ )
		{
			Unicodestring us = ((Unicodestring) getStringVector().get( i ));
			ArrayList frs = us.getFormattingRuns();
			String s = us.getStringVal();
			s = OOXMLAdapter.stripNonAscii( s ).toString();
			// sststrings.add(OOXMLAdapter.stripNonAscii(s));// zip.write(s); //
			// used as an index for cell values in parsing sheet ooxml

			// TODO: below should be in Unicodestring as .getOOXML?
			zip.write( "<si>" );
			zip.write( "\r\n" );

			if( frs == null )
			{ // no intra-string formattingz
				if( (s.indexOf( " " ) == 0) || (s.lastIndexOf( " " ) == (s.length() - 1)) )
				{
					zip.write( ("<t xml:space=\"preserve\">" + s + "</t>") );
				}
				else
				{
					zip.write( ("<t>" + s + "</t>") );
				}
				zip.write( "\r\n" );
			}
			else
			{ // have formatting runs which split up string into areas
				// with separate formats applied
				int begIdx = 0;
				for( int j = 0; j < frs.size(); j++ )
				{
					short[] idxs = (short[]) frs.get( j );
					if( idxs[0] > begIdx )
					{ // +1!!
						if( j == 0 )
						{
							zip.write( "<r>" ); // new rich text run
							zip.write( ("<t xml:space=\"preserve\">" + OOXMLAdapter.stripNonAscii( s.substring( begIdx,
							                                                                                    idxs[0] ) ) + "</t>") );
							zip.write( "</r>" );
							zip.write( "\r\n" );
						}
						else
						{
							zip.write( ("<t xml:space=\"preserve\">" + OOXMLAdapter.stripNonAscii( s.substring( begIdx,
							                                                                                    idxs[0] ) ) + "</t>") );
							zip.write( "</r>" );
							zip.write( "\r\n" );
						}
						begIdx = idxs[0];
					}
					zip.write( "<r>" ); // new rich text run
					Ss_rPr rp = Ss_rPr.createFromFont( getWorkBook().getFont( idxs[1] ) );
					zip.write( rp.getOOXML() );
				}
				if( begIdx < s.length() ) // output remaining string
				{
					s = s.substring( begIdx );
				}
				else
				{
					s = "";
				}
				zip.write( ("<t xml:space=\"preserve\">" + OOXMLAdapter.stripNonAscii( s ) + "</t>") );
				zip.write( "\r\n" );
				zip.write( "</r>" );
			}
			zip.write( "</si>" );
			zip.write( "\r\n" );
		}
		zip.write( "</sst>" );
		// return sstooxml.toString();
	}

	/**
	 * given SharedStrings.xml OOXML inputstream, read in string and formatting
	 * data, if any and parse into ArrayList for later use in parseSheetOOXML
	 *
	 * @param bk WorkBookHandle
	 * @param ii InputStream
	 * @return String ArrayList return list of shared strings
	 * @see parseSheetOOXML
	 */
	public static ArrayList parseOOXML( WorkBookHandle bk, InputStream ii )
	{
		// NOTE:
		// apparently can have dup entries in sharedstring.xml
		// index of string links to cell value so must keep dups here
		// reset after parsing
		boolean shareDups = false;
		if( bk.getWorkBook().isSharedupes() )
		{
			bk.getWorkBook().setSharedupes( false );
			shareDups = true;
		}
		try
		{
			XmlPullParserFactory factory = XmlPullParserFactory.newInstance();
			factory.setNamespaceAware( true );
			XmlPullParser xpp = factory.newPullParser();

			xpp.setInput( ii, "UTF-8" ); // using XML 1.0 specification
			int eventType = xpp.getEventType();
			while( eventType != XmlPullParser.END_DOCUMENT )
			{
				if( eventType == XmlPullParser.START_TAG )
				{
					String tnm = xpp.getName();
					if( tnm.equals( "si" ) )
					{ // parse si single string table
						// entry
						String s = "";
						ArrayList formattingRuns = null;
						while( eventType != XmlPullParser.END_DOCUMENT )
						{
							if( eventType == XmlPullParser.START_TAG )
							{
								if( xpp.getName().equals( "rPr" ) )
								{ // intra-string
									// formatting
									// properties
									int idx = s.length(); // index into
									// character string
									// to apply
									// formatting to
									Ss_rPr rp = (Ss_rPr) Ss_rPr.parseOOXML( xpp, bk ).cloneElement();
									Font f = rp.generateFont( bk ); // NOW CONVERT
									// ss_rPr to
									// a font!!
									int fIndex = bk.getWorkBook().getFontIdx( f ); // index
									// for
									// specific
									// font
									// formatting
									if( fIndex == -1 ) // must insert new font
									{
										fIndex = bk.getWorkBook().insertFont( f ) + 1;
									}
									if( formattingRuns == null )
									{
										formattingRuns = new ArrayList();
									}
									formattingRuns.add( new short[]{
											Integer.valueOf( idx ).shortValue(), Integer.valueOf( fIndex ).shortValue()
									} );
								}
								else if( xpp.getName().equals( "t" ) )
								{
									/*
									 * boolean bPreserve= false; if
									 * (xpp.getAttributeCount()>0) { if
									 * (xpp.getAttributeName(0).equals("space")
									 * &&
									 * xpp.getAttributeValue(0).equals("preserve"
									 * )) bPreserve= true; }
									 */
									eventType = xpp.next();
									while( (eventType != XmlPullParser.END_DOCUMENT) && (eventType != XmlPullParser.END_TAG) && (eventType != XmlPullParser.TEXT) )
									{
										eventType = xpp.next();
									}
									if( eventType == XmlPullParser.TEXT )
									{
										s += xpp.getText();
									}
								}
							}
							else if( (eventType == XmlPullParser.END_TAG) && xpp.getName().equals( "si" ) )
							{
								bk.getWorkBook().getSharedStringTable().addUnicodestring( s, formattingRuns ); // create
								// a
								// new
								// unicode
								// string
								// with
								// formatting
								// runs
								break;
							}
							eventType = xpp.next();
						}
					}
				}
				else if( eventType == XmlPullParser.END_TAG )
				{
				}
				eventType = xpp.next();
			}
		}
		catch( Exception e )
		{
			log.error( "SST.parseXML: " + e.toString() );
		}
		if( shareDups )
		{
			bk.getWorkBook().setSharedupes( true );
		}

		return (ArrayList) bk.getWorkBook().getSharedStringTable().getStringVector();
	}

	/**
	 * Returns all strings that are in the SharedStringTable for this workbook.
	 * The SST contains all standard string records in cells, but may not
	 * include such things as strings that are contained within formulas. This
	 * is useful for such things as full text indexing of workbooks
	 *
	 * @return Strings in the workbook.
	 */
	public ArrayList getAllStrings()
	{
		ArrayList al = new ArrayList( stringvector.size() );
		for( Object aStringvector : stringvector )
		{
			al.add( aStringvector.toString() );
		}
		return al;
	}

	/**
	 * Returns the length of this record, including the 4 header bytes
	 */
	@Override
	public int getLength()
	{
		int len = super.getLength();
		// if "hasGrbit" must account for additional size taken up by it
		// see ContinueHandler.createSstContinues
		for( int i = 0; i < (sstgrbits.size() - 1); i++ )
		{
			Byte b = (Byte) sstgrbits.get( i );
			if( b != null )
			{
				byte grbyte = b;
				if( (grbyte < 0x2) && (grbyte >= 0x0) ) // Sst grbit is either 0h
				// or 1h, otherwise it's
				// String data
				{
					len++;
				}
			}
		}
		return len;
	}
}
