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

import com.extentech.formats.LEO.BIGBLOCK;
import com.extentech.formats.LEO.Block;
import com.extentech.formats.LEO.LEOFile;
import com.extentech.formats.LEO.Storage;
import com.extentech.formats.XLS.charts.ChartObject;
import com.extentech.formats.XLS.charts.SeriesText;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.FastAddVector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.AbstractList;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;
import java.util.Vector;

/**
 * ByteStreamer Handles the low-level byte array streaming
 * of the WorkBook.  Basically a collection of XLS and methods for getting
 * at their data.
 *
 * @see WorkBook
 * @see Boundsheet
 */
public class ByteStreamer implements Serializable, XLSConstants
{
	private static final Logger log = LoggerFactory.getLogger( ByteStreamer.class );
	private static final long serialVersionUID = -8188652784510579406L;
	AbstractList records = new FastAddVector();
	private byte[] bytes;
	private String myname = "";
	protected WorkBook workbook = null;
	private OutputStream out = null;
	int dlen = -1;

	public int getRecordIndex( int opcode )
	{

		for( int i = 0; i < records.size(); i++ )
		{
			XLSRecord rec = (XLSRecord) records.get( i );
			if( rec.getOpcode() == opcode )
			{
				return i;
			}
		}
		return -1;

	}

	public void initTestRecVect()
	{
		// Useful for looking at the output vector & debugging ... THANKS, NICK! -jm
		Vector testVect;
		if( true )
		{
			log.info( "TestVector on in bytestreamer.stream" );
			testVect = new Vector();
			testVect.addAll( records );
		}
		log.info( "TESTING Recvec done." );
	}

	public ByteStreamer( WorkBook bk )
	{
		this.workbook = bk;
	}

	/**
	 * Returns an array of all the BiffRecs in the streamer.  For debug purposes,
	 * as there is no way to directy view the FastAddVector
	 *
	 * @return
	 */
	public Object[] getBiffRecords()
	{
		return records.toArray();
	}

	// 20060601 KSC: setBiffRecords
	public void setBiffRecords( Collection recs )
	{
		if( Arrays.equals( records.toArray(), recs.toArray() ) )
		{
			// already set!
			return;
		}
		records.clear();
		records.addAll( recs );
	}

	// public ByteStreamer(WorkBook b){this.book = b;}

	public void setBytes( byte[] b )
	{
		this.bytes = b;
	}

	public byte[] getBytes()
	{
		return this.bytes;
	}

	public int getRecVecSize()
	{
		return records.size();
	}

	public String getName()
	{
		return myname;
	}

	public String getSubstreamTypeName()
	{
		return "ByteStreamer";
	}

	public short getSubstreamType()
	{
		return WK_WORKSHEET;
	}

	/**
	 * get a record from the underlying array
	 */
	public BiffRec getRecordAt( int t, BiffRec rec )
	{
		if( (rec instanceof Boundsheet) || (rec instanceof Dbcell) )
		{
			return this.getRecordAt( t );
		}
		if( rec.getSheet() != null )
		{
			Sheet bs = rec.getSheet();
			return (BiffRec) bs.getSheetRecs().get( t );
		}
		throw new InvalidRecordException( "ERROR: ByteStreamer.getRecord() could not retrieve record from" );
	}

	/**
	 * get a record from the underlying array
	 */
	public BiffRec getRecordAt( int t )
	{
		return (BiffRec) records.get( t );
	}

	/**
	 * remove a record from the underlying array
	 */
	public boolean removeRecord( BiffRec rec )
	{
/*        if(rec instanceof Mulled) {
            if(((Mulled)rec).getMyMul()!=null) return true;
        }*/
		if( rec.getSheet() != null )
		{
			Sheet bs = rec.getSheet();
			if( bs.getSheetRecs().contains( rec ) )
			{
				return bs.getSheetRecs().remove( rec );
			}
		}
		return records.remove( rec );
	}

	/** get the recvec index of this record

	 public List getRecordSubList(int start, int end){
	 return  records.subList(start,end);
	 }*/

	/**
	 * get the recvec index of this record
	 */
	public int getRecordIndex( BiffRec rec )
	{
		if( rec.getSheet() != null )
		{
			Sheet bs = rec.getSheet();
			//	Logger.logInfo("ByteStreamer getting index of sheetRec: " + bs.toString());
			int ret = bs.getSheetRecs().indexOf( rec ); // bs.getRidx(); //
			if( ret > -1 )
			{
				return ret;
			}
			return records.indexOf( rec );
		}
		return records.indexOf( rec );
	}

	/**
	 * get the real (non-boundsheet based) recvec index of this record
	 */
	public int getRealRecordIndex( BiffRec rec )
	{
		return records.indexOf( rec );
	}

	int ridx = 0;

	/**
	 * add an BiffRec to this streamer.
	 */
	public void addRecord( BiffRec rec )
	{
		rec.setStreamer( this );
		Sheet sht = rec.getSheet();
		if( sht != null )
		{
			sht.getSheetRecs().add( rec );
		}
		else
		{
			records.add( rec );
		}
	}

	/**
	 * Bypass the sheet vecs...
	 *
	 * @param rec
	 * @param idx
	 */
	public void addRecordToBookStreamerAt( BiffRec rec, int idx )
	{
		try
		{
			records.add( idx, rec );
		}
		catch( ArrayIndexOutOfBoundsException e )
		{
			records.add( records.size(), rec );
		}
	}

	/**
	 * add an BiffRec to this streamer at the specified index.
	 */
	public void addRecordToSheetStreamerAt( BiffRec rec, int idx, Boundsheet sht )
	{
		rec.setStreamer( this );
		List sr = sht.getSheetRecs();
		sr.add( idx, rec );
	}

	/**
	 * add an BiffRec to this streamer at the specified index.
	 */
	public void addRecordAt( BiffRec rec, int idx )
	{
		rec.setStreamer( this );
		Sheet sht = rec.getSheet();
		if( sht != null )
		{
			List sr = sht.getSheetRecs();
			//  if(!sr.contains(rec))
			sr.add( idx, rec );
			// Logger.logInfo("ByteStreamer adding recAT: " + rec.toString() + " to Sheet: " + sht.toString());
		}
		else
		{
			// Logger.logInfo("ByteStreamer adding non-Sheet recAT: " + rec.toString());
			try
			{
				records.add( idx, rec );
				//     rec.setRecordIndexHint(idx);
			}
			catch( Exception e )
			{
				records.add( records.size(), rec );
				//   rec.setRecordIndexHint(records.size());
			}
		}
	}

	/**
	 * stream the bytes to an outputstream
	 */
	public int streamOut( OutputStream _out )
	{
		writeOut( _out );
		return dlen - 4;
	}

	public void writeRecord( OutputStream out, BiffRec rec ) throws IOException
	{
		byte[] op = ByteTools.shortToLEBytes( rec.getOpcode() );
		byte[] dt = rec.getData();
		byte[] ln = ByteTools.shortToLEBytes( (short) dt.length );
		out.write( op );
		out.write( ln );
		if( dt.length > 0 )
		{
			out.write( dt );
		}
			LEOFile.actualOutput += (op.length + ln.length + dt.length); // debugging
		rec.postStream();
	}

	/**
	 * Write all of the records to the output stream, including
	 * creating lbplypos records, assembling continues, etc.
	 */
	public StringBuffer writeOut( OutputStream out )
	{
		// create a byte level lockdown file in same directory as output
		StringBuffer lockdown = new StringBuffer();
		boolean lockit = false;
		if( System.getProperties().get( "com.extentech.ExtenXLS.autocreatelockdown" ) != null )
		{
			lockit = System.getProperties().get( "com.extentech.ExtenXLS.autocreatelockdown" ).equals( "true" );
		}

		// update tracker cells, packs formats ...
		this.workbook.prestream();
		byte[] dt;
		int recpos = 0, recctr = 0, dlen = 0;

		// get a private list of the records
		AbstractList rex = new FastAddVector( records.size() );
		rex.addAll( records );

		// first pass -- prepare SST
		BiffRec rec = null;
		Iterator e = rex.iterator();
		while( e.hasNext() )
		{
			rec = (BiffRec) e.next();
			++recctr;
			if( rec != null )
			{
				if( rec.getByteReader() != null )
				{
					rec.getByteReader().setApplyRelativePosition( true );
				}
				// Logger.logInfo("ByteStreamer.stream() PREStreaming: "+ rec);
				if( rec.getOpcode() == BOUNDSHEET )
				{
//  				add sheet recs to output vector
					List lst = ((Boundsheet) rec).assembleSheetRecs();
					rex.addAll( rex.size(), lst );
				}
				else if( rec.getOpcode() == SST )
				{
					// add extra bytes necessary for adding continue recx
					rec.preStream();
					recpos = (((Sst) rec).getNumContinues() * 4);
					if( recpos < 0 )// deal with empty SST
					{
						recpos = 0;
					}
					dlen += recpos;
				}
				else
				{
					rec.preStream(); // perform expensive processes
				}
			}
			else
			{
				log.warn( "Body Rec missing while preStreaming(): " + rec.toString() );
			}
		}
		e = rex.iterator();
		Index lastindex = null;
		int ctr = 0;
		while( e.hasNext() )
		{
			rec = (BiffRec) e.next();
			if( ctr == 0 )
			{ // handle the first BOF offset
				rec.setOffset( 0 );
				ctr++;
			}
			else
			{
				rec.setOffset( recpos );
			}

			if( rec.getOpcode() == INDEX )
			{ // need to get all it's component dbcell offsets set before we can process!
				if( lastindex != null )
				{
					lastindex.updateDbcellPointers();
				}
				lastindex = (Index) rec;
			}
//            offset dlen by number of new continue headers
			if( (rec.getOpcode() == CONTINUE) && (((Continue) rec).maskedMso != null) )
			{
				rec.setData( ((Continue) rec).maskedMso.getData() );    // ensure any mso changes are propogated up
			}
			int rln = rec.getLength();  //length of total rec data including continue
			int numcx = rln / MAXRECLEN;    // num continues?
			if( ((rln % MAXRECLEN) <= 4) && (numcx > 0) )    // hits boundary; since rlen==datalen+4, numcx is 1 more than actual continues
			{
				numcx--;
			}
			if( rec.getOpcode() == CONTINUE )
			{
				Continue thiscont = (Continue) rec;
				if( thiscont.isBigRecContinue() )
				{// could cause bugs... if related rec is trimmed
					rln = 0; // do not count data byte size for Continues...
				}
			}
			if( ((rln > (MAXRECLEN + 4)) && (numcx > 0) && (rec.getOpcode() != SST)) )
			{
				dlen += (numcx * 4);
				recpos += (numcx * 4);
			}
			dlen += rln;
			recpos += rln;
		}
		if( lastindex != null )
		{
			lastindex.updateDbcellPointers();
		}
		e = rex.iterator();

		/**
		 *  Get the updated Storages from LEO... output the RootStorage
		 *  and return the other storages and block index for output after
		 *  the workbook recs.
		 *
		 */
		LEOFile leo = this.workbook.factory.myLEO;
		List storages = null;
		storages = leo.writeBytes( out, (dlen) );
		BIGBLOCK hdrBlock = new BIGBLOCK();
		hdrBlock.init( leo.getHeader().getBytes(), 0, 0 );

		// ********************** WRITE OUT FILE IN CORRECT ORDER **************************************/
		/** Header = 1st sector/block */
		hdrBlock.writeBytes( out );

		// now output the workbook biff records
			LEOFile.actualOutput = 0;    // debugging
		while( e.hasNext() )
		{
			rec = (BiffRec) e.next();

			try
			{ // output the rec bytes
				// deal with CONTINUE record changes before streaming
				if( ContinueHandler.createContinues( rec, out, this ) )
				{
					// Logger.logInfo("Created continues for: " + rec.toString());
				}
				else
				{// Not a continued rec!
					this.writeRecord( out, rec );

					if( lockit )
					{
						byte[] op = ByteTools.shortToLEBytes( rec.getOpcode() );
						dt = rec.getData();
						byte[] ln = ByteTools.shortToLEBytes( (short) dt.length );

						// Logger.logInfo("=== WRITING RECORD DATA ===");
						//lockdown.append("rec:" + rec.toString());
						//lockdown.append("\r\n");
						lockdown.append( "opc:0x" + Integer.toHexString( rec.getOpcode() ) + " [" + ByteTools.getByteString( op,
						                                                                                                     false ) + "]" );
						lockdown.append( "\r\n" );
						lockdown.append( "len:0x" + Integer.toHexString( rec.getLength() ) + " [" + ByteTools.getByteString( ln,
						                                                                                                     false ) + "]" );
						lockdown.append( "\r\n" );
						//lockdown.append("off:0x" + Integer.toHexString(off+0x200));
						//lockdown.append("\r\n");
						lockdown.append( ByteTools.getByteDump( dt, 1 ) );
						lockdown.append( "\r\n" );
					}
				}
			}
			catch( Exception a )
			{
				throw new WorkBookException( "Streaming WorkBook Bytes failed for record: " + rec.toString() + ": " + a + " Output Corrupted.",
				                             WorkBookException.WRITING_ERROR,
				                             a );
			}
		}

		// pad to fit FAT size
			if( LEOFile.actualOutput != dlen )
			{
				log.debug( "Expected:" + dlen + " Actual: " + LEOFile.actualOutput + " Diff: " + (LEOFile.actualOutput - dlen) );
			}
		int leftover;
		int nBlocks = Math.max( leo.getMinBlocks(), (int) Math.ceil( dlen / (BIGBLOCK.SIZE * 1.0) ) + 1 );
		leftover = (nBlocks * BIGBLOCK.SIZE) - dlen;    // padding

		byte[] filler = new byte[leftover];
		for( int t = 0; t < filler.length; t++ )
		{
			filler[t] = 0;
		}

		// finally output the rest of the storages
		try
		{
			out.write( filler ); // workbook data filler
			Iterator its = storages.iterator();
			while( its.hasNext() )
			{
				Storage ob = (Storage) its.next();
				if( ob != null )
				{
					if( !ob.getName().equals( "Root Entry" ) || true )
					{  // was newLeoProcessing, so should always run?
						if( ob.getBlockType() != Block.SMALL )
						{// already written in miniFAT Container - see buildSAT
							ob.writeBytes( out );
							//Logger.logInfo("Streamed Storage:" + ob.getName() + ":"+ob.getActualFileSize());
						}
					}
				}
			}
		}
		catch( Exception a )
		{
			log.error( "Streaming WorkBook Storage Bytes failed.", a );
			throw new WorkBookException( "ByteStreamer.stream(): Body Rec missing while preStreaming(): " + rec.toString() + " " + a.toString() + " Output Corrupted.",
			                             WorkBookException.WRITING_ERROR );

		}
		if( lockit )
		{ // write the lockdown file
			return lockdown;
		}
		return null;
	}

	public void WriteAllRecs( String fName )
	{
		WriteAllRecs( fName, false );
	}

	/**
	 * Debug utility to write out ALL BIFF records
	 *
	 * @param fName          output filename
	 * @param bWriteSheeRecs truth of "include sheet records in output"
	 */
	public void WriteAllRecs( String fName, boolean bWriteSheetRecs )
	{
		try
		{
			java.io.File f = new java.io.File( fName );
			BufferedWriter writer = new BufferedWriter( new FileWriter( f ) );
			ArrayList recs = new ArrayList( java.util.Arrays.asList( this.getBiffRecords() ) );
			int ctr = 0;
			ctr = ByteStreamer.writeRecs( recs, writer, ctr, 0 );
			List sheets = this.workbook.getSheetVect();
			for( Object sheet : sheets )
			{
				if( bWriteSheetRecs )
				{
					ArrayList lst = (ArrayList) ((Boundsheet) sheet).assembleSheetRecs();
					ctr = ByteStreamer.writeRecs( lst, writer, ctr, 0 );
				}
				else
				{
					ctr = ByteStreamer.writeRecs( (ArrayList) ((Boundsheet) sheet).getSheetRecs(), writer, ctr, 0 );
				}
			}
			writer.flush();
			writer.close();
			writer = null;
		}
		catch( Exception e )
		{
		}
	}

	/**
	 * Debug utility to write out BIFF records in friendly fashion
	 *
	 * @param recArr array of records
	 * @param writer BufferedWriter
	 * @param level  For ChartRecs that have sub-records, recurse Level
	 */
	public static int writeRecs( ArrayList recArr, BufferedWriter writer, int ctr, int level )
	{
		String tabs = "\t\t\t\t\t\t\t\t\t\t";
		for( Object aRecArr : recArr )
		{
			try
			{
				BiffRec b = (BiffRec) aRecArr;
				if( b == null )
				{
					break;
				}
				writer.write( tabs.substring( 0, level ) + b.getClass().toString().substring( b.getClass()
				                                                                               .toString()
				                                                                               .lastIndexOf( '.' ) + 1 ) );
				if( b instanceof SeriesText )
				{
					writer.write( "\t[" + b.toString() + "]" );
				}
				else if( b instanceof Continue )
				{
					if( ((Continue) b).maskedMso != null )
					{
						writer.write( "\t[MASKED MSO ******************" );
						writer.write( "\t[" + ((Continue) b).maskedMso.toString() + "]" );
						writer.write( ((Continue) b).maskedMso.debugOutput() );
						writer.write( "\t[" + ByteTools.getByteDump( b.getData(), 0 ).substring( 11 ) + "]" );
					}
					else
					{
						writer.write( "\t[" + ByteTools.getByteDump( b.getData(), 0 ).substring( 11 ) + "]" );
					}
				}
				else if( b instanceof MSODrawing )
				{
					writer.write( "\t[" + b.toString() + "]" );
//					writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
					writer.write( ((MSODrawing) b).debugOutput() );
					writer.write( "\t[" + ByteTools.getByteDump( b.getData(), 0 ).substring( 11 ) + "]" );
				}
				else if( b instanceof Obj )
				{
					writer.write( ((Obj) b).debugOutput() );
				}
				else if( b instanceof MSODrawingGroup )
				{
					writer.write( "\t[" + b.toString() + "]" );
//					writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
				}
				else if( b instanceof Label )
				{
					writer.write( "\t[" + b.getStringVal() + "]" );
				}
				else if( b instanceof Mulblank )
				{
					writer.write( "\t[" + b.getCellAddress() + "]" );
				}
				else if( b instanceof Name )
				{
					try
					{
						writer.write( "\t[" + ((Name) b).getName() + "-" + ((Name) b).getLocation() + "]" );
					}
					catch( Exception ce )
					{
						writer.write( "\t[" + ((Name) b).getName() + "-ERROR IN LOCATION]" );
					}
					//writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
				}
				else if( b instanceof Sst )
				{
					writer.write( "\t[" + b.toString() + "]" );
				}
				else if( b instanceof Pls )
				{
					writer.write( "\t[" + b.toString() + "]" );
				}
				else if( b instanceof Supbook )
				{
					writer.write( "\t[" + b.toString() + "]" );
//					writer.write("\t[" + ByteTools.getByteDump(b.getData(), 0).substring(11)+ "]");
				}
				else if( b instanceof Crn )
				{
					writer.write( "\t[" + b.toString() + "]" );
				}
				else if( (b instanceof Formula) || (b instanceof Rk) || (b instanceof NumberRec) || (b instanceof Blank) || (b instanceof Labelsst) )
				{
					writer.write( " " + ((XLSRecord) b).getCellAddressWithSheet() + "\t[" + ByteTools.getByteDump( b.getData(), 0 )
					                                                                                 .substring( 11 ) + "]" );
				}
				else // all else, write bytes
				{
					writer.write( "\t[" + ByteTools.getByteDump( ByteTools.shortToLEBytes( b.getOpcode() ),
					                                             0 ) + "][" + ByteTools.getByteDump( b.getData(), 0 )
					                                                                   .substring( 11 ) + "]" );
				}
				writer.newLine();
				if( b instanceof ChartObject )
				{
					writeRecs( ((ChartObject) b).getChartRecords(), writer, ctr, level + 1 );
				}
			}
			catch( Exception e )
			{
			}
		}
		return ctr;
	}

	/**
	 * clear out object references
	 */
	public void close()
	{
		for( int i = 0; i < records.size(); i++ )
		{
			XLSRecord r = (XLSRecord) records.get( i );
			r.close();
		}
		records.clear();
		workbook = null;
	}
}