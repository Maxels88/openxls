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

import com.extentech.toolkit.Logger;

import java.io.OutputStream;
import java.io.Serializable;
import java.util.Iterator;
import java.util.List;

/**
 * this class takes care of the tasks related to
 * working with Continue records.
 * <p/>
 * when a Continue record is created, it needs to be
 * associated with its related XLSRecord -- the record whose
 * data it contains.
 */
public class ContinueHandler implements Serializable, XLSConstants
{
	/**
	 *
	 */
	private static final long serialVersionUID = 164009339243774537L;
	private static int DEBUGLEVEL = 0;
	private static boolean processContinues = true; // debug setting
	private BiffRec continued;
	private boolean handleTxo = false;
	private boolean handleObj = false;
	private Txo lastTxo;
	private Obj lastObj;
	private Continue lastCont;
	private WorkBook book;

	private BiffRec splitPrevRec;
	private BiffRec splitContRec;

	public ContinueHandler( WorkBook b )
	{
		this.book = b;
	}

	/**
	 * add an XLSRecord to this handler
	 * check if it needs a Continue, if
	 * so, put in our continued_recs.
	 */
	public void addRec( BiffRec rec, int datalen )
	{
		// check if this record has Continue records
		// if so, delay initialization until all Continues are read.
		// In the body.getData() method we then need to check
		// if the record has Continues and if so, then read/write init
		// data from them.

		short opcode = rec.getOpcode();
		short nextOpcode = 0x0;

		if( opcode != EOF )
		{
			nextOpcode = book.getFactory().lookAhead( rec );
		}

		if( nextOpcode == CONTINUE )
		{
			if( DEBUGLEVEL > 11 )
			{
				Logger.logInfo( "Next OPCODE IS CONTINUE: " + Integer.toHexString( nextOpcode ) );
			}
		}
		if( (nextOpcode == CONTINUE) && (opcode != CONTINUE) )
		{ // the continued rec
			if( continued != null )
			{
				continued.init();
				continued = null;
			}
			this.continued = rec;
			this.splitPrevRec = this.continued;
			// if the rec is a Txo, we need to process special
			if( continued instanceof Txo )
			{
				this.handleTxo = true;
				lastTxo = (Txo) continued;
				lastObj = null;
				this.handleObj = false;    // ""
			}
			else if( continued instanceof Obj )
			{
				this.handleObj = true;
				lastObj = (Obj) continued;
				// obj records need to be init'd before setsheet
				lastObj.init();
				lastTxo = null;
				this.handleTxo = false;    // ""
			}
			else
			{
				this.handleTxo = false;
				this.handleObj = false;
			}
			lastCont = null;
		}
		else if( opcode == CONTINUE )
		{ // add to the continued rec
			splitContRec = rec;
			rec.init();
			if( (continued == null) && (lastCont == null) )
			{
				//This is a use case where a chart is in the middle of an Obj, PLS, or Txo record.  A continue
				// record appears at the end of the last chart EOF, and needs to remain inorder to not cause corruption
				if( splitPrevRec != null )
				{
					((Continue) rec).setPredecessor( splitContRec );
				}
				if( DEBUGLEVEL > 0 )
				{
					Logger.logWarn( "Warning:  Out of spec split txo continue record found, reconstructing." + splitPrevRec.toString() );
				}

			}
			else
			{
				if( lastCont != null )
				{
					((Continue) rec).setPredecessor( lastCont );
				}
				else
				{
					((Continue) rec).setPredecessor( continued );
				}
				if( (continued.getOpcode() == SST) || (continued.getOpcode() == STRINGREC) )
				{
					if( DEBUGLEVEL > 2 )
					{
						Logger.logInfo( "Sst Continue.  grbit:" + ((Continue) rec).getGrbit() );
					}
				}
				else
				{
					((Continue) rec).setHasGrbit( false ); // if it can't have one, don't
				}
			}

			// last data rec was a Txo -- add next 2 Continues
			if( handleTxo )
			{
				// txo's have either 2 continues following (text, formatting runs), no continues (an "empty" txo); may also contain continues masking mso's
				if( lastTxo.text == null )
				{
					if( !isMaskedMSODrawingRec( rec.getData() ) )  // in almost all cases, the next continue is a Text continue
					{
						lastTxo.text = (Continue) rec;
					}
					else
					{ // it is possible that an empty TXO (one that does NOT contain Text) is followed by a Continue rec which is masking an MSODrawing
						continued = createMSODrawingFromContinue( rec );        // create a new MSODrawing rec from the Continue rec's data
						((Continue) rec).maskedMso = (MSODrawing) continued;    // set maskedMso in Continue to identify
					}
				}
				else if( lastTxo.formattingruns == null )
				{
					lastTxo.formattingruns = (Continue) rec;
				}
				else
				{  // a third continues: will be a masked mso or possibly a "big rec" continues
					try
					{
						if( isMaskedMSODrawingRec( rec.getData() ) )
						{
							continued = createMSODrawingFromContinue( rec );        // create a new MSODrawing rec from the Continue rec's data
							((Continue) rec).maskedMso = (MSODrawing) continued;    // set maskedMso in Continue to identify
						}
						else if( continued != null )
						{    // then it's a "big rec" continue
							continued.addContinue( (Continue) rec );
							((XLSRecord) continued).mergeContinues();    // must merge continues "by hand" because data member var is set already
							continued.removeContinues();
						}
					}
					catch( Exception e )
					{
						if( DEBUGLEVEL > 0 )
						{
							Logger.logErr( "ContinueHandler.txo parsing- encountered unknown Continue record" );
						}
					}
					lastCont = (Continue) rec;
				}
			}
			else if( handleObj )
			{
				try
				{    // When a Continue record follows an Obj record, it is either masking a Msodrawing record - or it's a "big rec" continues
					if( isMaskedMSODrawingRec( rec.getData() ) )
					{
						continued = createMSODrawingFromContinue( rec );        // create a new MSODrawing rec from the Continue rec's data
						((Continue) rec).maskedMso = (MSODrawing) continued;    // set maskedMso in Continue to identify
					}
					else if( continued != null )
					{    // then it's a "big rec" continue
						continued.addContinue( (Continue) rec );
						((XLSRecord) continued).mergeContinues();    // must merge continues "by hand" because data member var is set already
						continued.removeContinues();
					}
					lastCont = (Continue) rec;
				}
				catch( Exception e )
				{
					if( DEBUGLEVEL > 0 )
					{
						Logger.logErr( "ContinueHandler.Obj parsing- encountered unknown Continue record" );
					}
				}
			}
			else
			{
				// null continued?  bizarre case testfiles/assess.xls -jm 10/01/2004
				if( continued != null )
				{
					continued.addContinue( (Continue) rec );
				}

				lastCont = (Continue) rec;
			}
		}
		else
		{
			if( continued != null )
			{
				continued.init();
		         /* If Formula Attached String was a continue, Formula cachedValue WAS NOT init-ed. Do here.*/
				if( continued.getOpcode() == STRINGREC )
				{
					this.book.lastFormula.setCachedValue( ((StringRec) continued).getStringVal() );
				}
				continued = null;
			}

			// associate boundsheet for init
			if( book.getLastbound() != null )
			{
				if(/*opcode == NAME ||*/
						opcode == FORMULA )
				{
					rec.setSheet( book.getLastbound() );
				}
			}
			// init names without init'ing expression -- must init expression after loading sheet recs ...
			if( opcode != XLSConstants.NAME )
			{
				rec.init();
			}
			else
			{
				((Name) rec).init( false );    // don't init expression here; do after loading sheet recs ...
			}
			lastCont = null;
		}
	}

	/**
	 * returns true if this id is one of an MSODrawing
	 * <br>occurs when a Continue record is masking an MSO
	 * (i.e. contains the record structure of an MSO in it's data,
	 * with an opcode of Continue)
	 *
	 * @param data - record data
	 * @return true if data is in form of an MSODrawing record
	 */
	private boolean isMaskedMSODrawingRec( byte[] data )
	{
		if( data.length > 3 )
		{
			int id = (((0xFF & (byte) data[3]) << 8) | (0xFF & data[2]));
			return (((id == MSODrawingConstants.MSOFBTSPCONTAINER) ||
					(id == MSODrawingConstants.MSOFBTSOLVERCONTAINER) ||
					(id == MSODrawingConstants.MSOFBTSPGRCONTAINER) ||
					(id == MSODrawingConstants.MSOFBTCLIENTTEXTBOX)));
		}
		return false;
	}

	/**
	 * create an MSODrwawing Record from a Continue record
	 * which is masking an MSODrawing (i.e. contains the record
	 * structure of an MSO in it's data, with an opcode of Continue)
	 *
	 * @param rec
	 * @return
	 */
	private MSODrawing createMSODrawingFromContinue( BiffRec rec )
	{
		// create an mso and add to drawing recs ...
		MSODrawing mso = new MSODrawing();
		mso.setOpcode( MSODRAWING );
		mso.setWorkBook( rec.getWorkBook() );
		mso.setData( rec.getData() );
		mso.setLength( rec.getData().length );
		mso.setDebugLevel( this.DEBUGLEVEL );
		mso.setStreamer( book.getStreamer() );
		return mso;
	}

	/**
	 * check if the record needs to have its data
	 * spanned across Continue records.
	 */
	public static boolean createContinues( BiffRec rec, OutputStream out, ByteStreamer streamer )
	{
		int datalen = rec.getLength();
		int opc = rec.getOpcode();
		// Logger.logInfo("ContinueHandler creating output continues for: " + rec.toString() + " datalen: " + datalen);
		// if greater than 8023, we need Continues
		if( opc == CONTINUE )
		{
			if( ((Continue) rec).isBigRecContinue() ) // skip ensuing Continues as should be written by main record
			{
				return true;
			}
			// handle masked mso's which have continues separately
			if( (((Continue) rec).maskedMso != null) && ((((Continue) rec).maskedMso.getLength() - 4) > MAXRECLEN) )
			{
				((Continue) rec).maskedMso.setOpcode( CONTINUE );    // so can add the correct record to output
				createBigRecContinues( ((Continue) rec).maskedMso, out, streamer );
				((Continue) rec).maskedMso.setOpcode( MSODRAWING );    // reset so can continue working with this record set
				return true;    // processed, return true
			}
		}
		else if( (opc == SST) )
		{
			createSstContinues( (Sst) rec, out, streamer );
			return true;
		}
		else if( opc == TXO )
		{
			createTxoContinues( (Txo) rec, out, streamer );
			return true;
		}
		else if( opc == MSODRAWINGGROUP )
		{
			createMSODGContinues( rec, out, streamer );
			return true;
		}
		else if( ((datalen - 4) > MAXRECLEN) )
		{
			createBigRecContinues( rec, out, streamer );
			return true;
		}
		return false;
	}

	/**
	 * check if the record needs to have its data
	 * spanned across Continue records.
	 */
	static int createContinues( BiffRec rec, int insertLoc )
	{
		int datalen = rec.getLength();

		//Logger.logInfo("ContinueHandler creating output continues for: " + rec.toString() + " datalen: " + datalen);
		// if greater than 8023, we need Continues
		if( (rec instanceof Obj) || (rec instanceof MSODrawing) || (rec instanceof MSODrawingGroup) )
		{
			return createObjContinues( rec );
		}
		if( (rec instanceof Sst) )
		{
			return createSstContinues( (Sst) rec, insertLoc );
		}
		if( (datalen > MAXRECLEN) && (!(rec instanceof Continue)) )
		{
			return createBigContinues( rec, insertLoc );
		}
		if( rec instanceof Txo )
		{
			return createTxoContinues( (Txo) rec );
		}
		return 0;
	}

	/**
	 * generate Continue records for Sst records
	 */
	public static int createSstContinues( Sst rec, int insertLoc )
	{
		byte[] dta = rec.getData();
		int datalen = dta.length;
		if( datalen < MAXRECLEN )
		{
			return 0;
		}
		// get the grbits and continue sizes
		Object[] continuedef = rec.getContinueDef( rec, false );
		Integer[] continuesizes = (Integer[]) continuedef[0];
		Byte[] sstgrbits = (Byte[]) continuedef[1];
		// int sstoffset =  continuesizes[0].intValue() - rec.getOrigSstLen();
		int numconts = continuesizes.length - 1;

		// blow out old record Continues
		removeContinues( rec );

		// account for the offset caused by the Sstgrbits
		int sizer = 0;
		int dtapos = 0;

		// create Continues, skip the first which is Sst recordbody
		for( int i = 1; i <= numconts; i++ )
		{ // start after the 1st continue length which is the Sst data body
			if( continuesizes[i] == 0 )
			{
				break;
			}
			// numconts is one less than continuesizes, match continuesizes[1] with numconts[0]...
			Byte thisgr = sstgrbits[i - 1];
			dtapos += continuesizes[i - 1];

			// check for a grbit -- null grbit means Continue Breaks on a one-char UString
			boolean hasGrbit = false;
			if( thisgr != null )
			{
				hasGrbit = ((thisgr < 0x2) && (thisgr >= 0x0)); // Sst grbit is either 0h or 1h, otherwise it's String data
			}
			if( continuesizes[i] == MAXRECLEN )
			{
				hasGrbit = false; // this is a non-standard Sst Continue
			}

			sizer = continuesizes[i];
			if( i == numconts )
			{
				sizer = (datalen - dtapos);
			}
			if( hasGrbit )
			{
				sizer++;
			}
			byte[] continuedata = new byte[sizer];

			if( hasGrbit )
			{
				if( DEBUGLEVEL > 1 )
				{
					Logger.logInfo( "New Continue. HAS grbit." );
					Logger.logInfo( "Continue GRBIT: " + String.valueOf( thisgr ) );
				}
				continuedata[0] = thisgr; // set a grbit on the new Continue
				System.arraycopy( dta, dtapos, continuedata, 1, continuedata.length - 1 );
			}
			else
			{
				if( DEBUGLEVEL > 1 )
				{
					Logger.logInfo( "New Continue. NO grbit." );
					Logger.logInfo( "Continue GRBIT: " + String.valueOf( dta[dtapos] & 0x1 ) );
				}
				System.arraycopy( dta, dtapos, continuedata, 0, continuedata.length );
			}

			Continue thiscont = addContinue( rec, continuedata, insertLoc, rec.wkbook );
			if( hasGrbit )
			{
				thiscont.setHasGrbit( true );
			}
			insertLoc++;
		}
		int sstsize = continuesizes[0];
		trimRecSize( rec, sstsize );
		return numconts;
	}

	/**
	 * Write a record out to the output stream.  I have zero idea why this is in here and
	 * not in bytestreamer or elsewhere.  Anyway, we are passing around the streamer in here
	 * for when it's needed...
	 *
	 * @param rec
	 * @param out
	 */
	private static void writeRec( BiffRec rec, OutputStream out, ByteStreamer streamer )
	{
		if( rec.getOpcode() == CONTINUE )
		{
			rec.preStream();
		}

		try
		{ // output the rec bytes
			streamer.writeRecord( out, rec );
		}
		catch( Exception a )
		{
			Logger.logErr( "Streaming WorkBook Bytes for record:" + rec.toString() + " failed: " + a + " Output Corrupted." );
		}
	}

	/**
	 * generate Continue records for Sst records
	 */
	public static void createSstContinues( Sst rec, OutputStream out, ByteStreamer streamer )
	{
		byte[] dta = rec.getData();
		int datalen = dta.length;
		// get the grbits and continue sizes
		Object[] continuedef = rec.getContinueDef( rec, false );
		Integer[] continuesizes = (Integer[]) continuedef[0];

		int sstsize = 0;
		if( continuesizes.length > 0 )
		{
			Integer sstz = continuesizes[0];
			if( sstz != null )
			{
				sstsize = sstz;
			}
			trimRecSize( rec, sstsize );
		}
		// output the original rec
		writeRec( rec, out, streamer );

		Byte[] sstgrbits = (Byte[]) continuedef[1];
		int numconts = continuesizes.length - 1;

		// blow out old record Continues
		removeContinues( rec ); // are they even there? Should not be!

		// account for the offset caused by the Sstgrbits
		int sizer = 0;
		int dtapos = 0;

		// create Continues, skip the first which is Sst recordbody
		for( int i = 1; i <= numconts; i++ )
		{ // start after the 1st continue length which is the Sst data body
			if( continuesizes[i] == 0 )
			{
				break;
			}
			// numconts is one less than continuesizes, match continuesizes[1] with numconts[0]...
			Byte thisgr = sstgrbits[i - 1];
			dtapos += continuesizes[i - 1];

			// check for a grbit -- null grbit means Continue Breaks on a one-char UString
			boolean hasGrbit = false;
			if( thisgr != null )
			{
				hasGrbit = ((thisgr < 0x2) && (thisgr >= 0x0)); // Sst grbit is either 0h or 1h, otherwise it's String data
			}
			if( continuesizes[i] == MAXRECLEN )
			{
				hasGrbit = false; // this is a non-standard Sst Continue
			}

			sizer = continuesizes[i];
			if( i == numconts )
			{
				sizer = (datalen - dtapos);
			}
			if( hasGrbit )
			{
				sizer++;
			}
			byte[] continuedata = new byte[sizer];

			if( hasGrbit )
			{
				if( DEBUGLEVEL > 1 )
				{
					Logger.logInfo( "New Continue. HAS grbit." );
					Logger.logInfo( "Continue GRBIT: " + String.valueOf( thisgr ) );
				}
				continuedata[0] = thisgr; // set a grbit on the new Continue
				System.arraycopy( dta, dtapos, continuedata, 1, continuedata.length - 1 );
			}
			else
			{
				if( DEBUGLEVEL > 1 )
				{
					Logger.logInfo( "New Continue. NO grbit." );
					Logger.logInfo( "Continue GRBIT: " + String.valueOf( dta[dtapos] & 0x1 ) );
				}
				System.arraycopy( dta, dtapos, continuedata, 0, continuedata.length );
			}

			Continue thiscont = createContinue( continuedata, rec.wkbook );
			if( hasGrbit )
			{
				thiscont.setHasGrbit( true );
			}
			// output the original rec
			writeRec( thiscont, out, streamer );

		}
	}

	/**
	 * remove Continues from a record
	 * <p/>
	 * TODO:  Can this be removed now?  We shouldn't really have continues in memory, just on stream, NO?
	 */
	public static void removeContinues( BiffRec rec )
	{
		// remove existing Continues (if any)
		List oldconts = rec.getContinueVect();
		if( oldconts == null )
		{
			return;
		}
		if( oldconts.size() > 0 )
		{
			Iterator it = oldconts.iterator();
			while( it.hasNext() )
			{
				Object ob = it.next();
				rec.getStreamer().removeRecord( (BiffRec) ob );
				ob = null; // faster!!  will it work?
			}
			rec.removeContinues();
		}
	}

	/**
	 * generate Continue records for records with lots of data
	 */
	public static void createBigRecContinues( BiffRec rec, OutputStream out, ByteStreamer streamer )
	{
		byte[] dta = rec.getData();
		int datalen = dta.length;
		int numconts = datalen / MAXRECLEN;

		if( datalen > MAXRECLEN )
		{
			trimRecSize( rec, MAXRECLEN );
			writeRec( rec, out, streamer );
		}
		else
		{
			writeRec( rec, out, streamer );
			return;
		}

		// create Continues
		int[] boundaries = ContinueHandler.getBoundaries( numconts );
		int sizer = MAXRECLEN;
		for( int i = 0; i < numconts; i++ )
		{
			// if this is the last Continue rec it is probably shorter than CONTINUESIZE
			if( (datalen - boundaries[i]) < MAXRECLEN )
			{
				sizer = (datalen - boundaries[i]);
			}
			byte[] continuedata = new byte[sizer];
			System.arraycopy( dta, boundaries[i], continuedata, 0, continuedata.length );
			Continue cr = createContinue( continuedata, rec.getWorkBook() );
			writeRec( cr, out, streamer );
		}
	}

	/**
	 * generate Continue records for records with lots of data
	 */
	// 20070921 KSC: Is this used?
	public static int createBigContinues( BiffRec rec, int insertLoc )
	{
		byte[] dta = rec.getData();
		int datalen = dta.length;
		int numconts = datalen / MAXRECLEN;
		// create Continues
		Continue[] conts = new Continue[numconts];
		int[] boundaries = ContinueHandler.getBoundaries( numconts );
		// remove existing Continues (if any)
		removeContinues( rec );
		int sizer = MAXRECLEN;
		for( int i = 0; i < numconts; i++ )
		{
			// if this is the last Continue rec it is probably shorter than CONTINUESIZE
			if( (datalen - boundaries[i]) < MAXRECLEN )
			{
				sizer = (datalen - boundaries[i]);
			}
			byte[] continuedata = new byte[sizer];
			System.arraycopy( dta, boundaries[i], continuedata, 0, continuedata.length );
			conts[i] = addContinue( rec, continuedata, insertLoc, rec.getWorkBook() );
			insertLoc++;
		}
		trimRecSize( rec, MAXRECLEN );
		return numconts;
	}

	/**
	 * Create and initialize a new Continue record
	 *
	 * @param rec       the XLSRecord owner of the Continue
	 * @param data      the pre-sized Continue body data
	 * @param streampos the position to insert the new Continue into the data stream
	 */
	public static Continue createContinue( byte[] data, Book book )
	{
		Continue cont = new Continue();
		cont.setWorkBook( (WorkBook) book );
		cont.setData( data );
		cont.setStreamer( book.getStreamer() );        // 20070921 KSC: Addded
		int len = data.length;
		cont.setOpcode( CONTINUE );
		cont.setLength( (short) len );
		return cont;
	}

	/**
	 * Create and initialize a new Continue record
	 *
	 * @param rec       the XLSRecord owner of the Continue
	 * @param data      the pre-sized Continue body data
	 * @param streampos the position to insert the new Continue into the data stream
	 */
	public static Continue addContinue( BiffRec rec, byte[] data, int streampos, Book book )
	{
		Continue cont = new Continue();
		rec.addContinue( cont );
		cont.setWorkBook( (WorkBook) book );
		cont.setData( data );
		int len = data.length;
		cont.setOpcode( CONTINUE );
		cont.setLength( (short) len );
		return cont;
	}

	/**
	 * generate the mandatory Continue records for the Txo rec type
	 * <p/>
	 * Txo must have at least 2 Continue recs
	 * first one contains text data
	 * second (last) one contains formatting runs
	 */
	public static int createTxoContinues( Txo rec )
	{
		List txoConts = rec.getContinueVect();
		if( txoConts == null )
		{
			return 0;
		}
		// iterate through existing Continues and update their data.
		//   for(int i = 0;i< txoConts.length;i++){
		// step into first Continue and find its
		// boundary -- everything after goes into
		// the 'formatting' Continue rec.  For now
		// we assume that this can't change

		// now create data Continue(s) for text

		// we can ignore the formatting Continue
		// but text may be formatted wierd -- SORRY! : )

		// insert NEW Continues into byte stream

		//  }
		if( rec.getLength() > (MAXRECLEN + 4) )
		{
			ContinueHandler.trimRecSize( rec, MAXRECLEN );
		}
		return 0;//txoConts.length;

	}

	/**
	 * generate the mandatory Continue records for the Txo rec type
	 * <p/>
	 * Txo must have at least 2 Continue recs
	 * first one contains text data
	 * second (last) one contains formatting runs
	 */
	public static void createTxoContinues( Txo rec, OutputStream out, ByteStreamer streamer )
	{
		byte[] dta = rec.getBytes();
		WorkBook book = rec.getWorkBook();
		if( dta.length > (MAXRECLEN + 4) )
		{
			ContinueHandler.trimRecSize( rec, MAXRECLEN );
			// stream the rec to out...
			ContinueHandler.writeRec( rec, out, streamer );
			// and now the restof the recs
			Continue[] myconts = getContinues( dta, MAXRECLEN, book );
			for( int x = 0; x < myconts.length; x++ )
			{
				ContinueHandler.writeRec( myconts[x], out, streamer );
			}
		}
		else
		{
			// stream the rec to out...
			ContinueHandler.writeRec( rec, out, streamer );
		}

	}

	/**
	 * generate the mandatory Continue records for the Obj rec type
	 */
	public static void createMSODGContinues( BiffRec rec, OutputStream out, ByteStreamer streamer )
	{
		byte[] dta = rec.getData();
		int datalen = dta.length;
		int numconts = datalen / MAXRECLEN;

		if( datalen > MAXRECLEN )
		{
			trimRecSize( rec, MAXRECLEN );
			writeRec( rec, out, streamer );
		}
		else
		{
			writeRec( rec, out, streamer );
			return;
		}
		// create Continues
		int[] boundaries = ContinueHandler.getBoundaries( numconts );
		int sizer = MAXRECLEN;
		for( int i = 0; i < numconts; i++ )
		{
			// if this is the last Continue rec it is probably shorter than CONTINUESIZE
			if( (datalen - boundaries[i]) < MAXRECLEN )
			{
				sizer = (datalen - boundaries[i]);
			}
			if( sizer == 0 )// reclen hits boundary exactly; ignore last continue; see ByteStreamer.writeOut for boundary issues
			{
				break;
			}
			byte[] continuedata = new byte[sizer];
			System.arraycopy( dta, boundaries[i], continuedata, 0, continuedata.length );

			BiffRec cr = null;
//          now add a second MSODG -- acts like a continue, exists to confuse and dismay

			if( i == 0 )
			{
				cr = MSODrawingGroup.getPrototype();
				cr.setData( continuedata );
			}
			else
			{
				cr = createContinue( continuedata, rec.getWorkBook() );
			}
			writeRec( cr, out, streamer );
		}
		// reset rec
		rec.setData( dta );
	}

	private static Continue[] getContinues( byte[] dta, int start, WorkBook book )
	{

		int clen = dta.length - start;
		int len = 0, pos = 0;
		int numconts = clen / MAXRECLEN;
		numconts++;
		Continue[] retconts = new Continue[numconts];

		Logger.logInfo( "Creating continues: " + numconts );

		byte[] dtx = null;
		for( int x = 0; x < numconts; x++ )
		{
			;
			if( clen > MAXRECLEN )
			{
				len = MAXRECLEN;
			}
			else
			{
				len = clen;
			}

			if( len > 0 )
			{
				dtx = new byte[len];
			}
			else
			{
				dtx = new byte[clen];
				len = clen;
			}

			// populate the bytes
			System.arraycopy( dta, start, dtx, 0, len );
			retconts[x] = createContinue( dtx, book );
			clen -= len;
		}
		return retconts;
	}

	/**
	 * trims the current rec to MAXRECLEN size
	 */
	public static int createObjContinues( BiffRec rec )
	{
		if( rec.getLength() > (MAXRECLEN + 4) )
		{
			ContinueHandler.trimRecSize( rec, MAXRECLEN );
		}
		return 0;//objConts.length;

	}

	/**
	 * get record size boundaries which determine the
	 * break-point of Continue record data
	 *
	 * @param x the number of boundaries needed
	 */
	static int[] getBoundaries( int x )
	{
		int[] boundaries = new int[x];
		int thisbound = 0;
		for( int i = 0; i < x; i++ )
		{
			thisbound += MAXRECLEN;
			boundaries[i] = thisbound;
		}
		return boundaries;
	}

	/**
	 * trim original rec to max rec size
	 * (for records with size-related Continues, not Txos)
	 */
	public static void trimRecSize( BiffRec rec, int CONTINUESIZE )
	{
		byte[] dta = rec.getData();
		byte[] newdata = new byte[CONTINUESIZE];
		System.arraycopy( dta, 0, newdata, 0, CONTINUESIZE );
		rec.setData( newdata );
	}

	/**
	 * clear out object references in prep for closing workbook
	 */
	public void close()
	{
		continued = null;
		lastTxo = null;
		lastObj = null;
		lastCont = null;
		book = null;

	}
}