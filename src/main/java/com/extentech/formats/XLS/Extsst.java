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

import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.Logger;

import java.util.List;

/**
 * <b>Extsst: Extended Shared String Table (FFh)</b><br>
 * hashtable that optimizes external copy operations.
 * <p><pre>
 * offset  name        size    contents
 * ---
 * 4       Dsst        2       Number of strings in each bucket
 * 6       Rgisstinf   var     Array of ISSTINF Structures
 * </p></pre>
 * <p/>
 * <p/>
 * The first WORD in each ISSTINF entry in the EXTSST record is an absolute
 * offset into the Workbook stream.  It's value isn't record dependent.  The
 * second WORD in each ISSTINF entry in the EXTSST record is the offset into
 * the record where the string is found.  In the case where the string is
 * found in a CONTINUE record, the value should be the offset into the
 * CONTINUE record.
 * <p/>
 * For example:
 * In an Excel workbook with a lot of strings, the SST record might start at
 * offset 0x07F2 into the Workbook stream.  The next record, a CONTINUE
 * record, starts at offset 0x2810.  The first 52 ISSTINF entries in the
 * EXTSST record work just as we would expect, with offsets into the SST
 * record.  The 53 ISSTINF entry contains an offset into the CONTINUE record,
 * and absolute addressing continues normally.
 * <p/>
 * ISSTINF entries 52 and 53 look like this:
 * <p/>
 * ISSTINF 52 absolute offset: 0x27E5
 * ISSTINF 52 relative offset: 0x1FF3
 * <p/>
 * ISSTINF 53 absolute offset: 0x28A4
 * ISSTINF53 relative offset: 0x0094
 * <p/>
 * As you can see, while the absolute offset continues to increment, the
 * relative offset resets to begin addressing from the beginning of the
 * CONTINUE record.
 *
 * @see SST
 * @see LABELSST
 * @see EXTSST
 */

public final class Extsst extends com.extentech.formats.XLS.XLSRecord
{
	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = -2409458213287279987L;
	Sst mysst = null;
	short Dsst;
	int numstructs;
	Isstinf[] mystinfs = null;
	boolean debug;

	void setSst( Sst s )
	{
		this.mysst = s;
		s.setExtsst( this );
	}

	/**
	 * because the EXTSST comes between the BOUNDSHEET
	 * records and all BOUNDSHEET BOFs, the lbPlyPos needs
	 * to change for all Worksheets when this record's size changes.
	 */
	public boolean getUpdatesAllBOFPositions()
	{
		return true;
	}

	@Override
	public void init()
	{
		debug = true;

		super.init();
		Dsst = ByteTools.readShort( this.getByteAt( 0 ), this.getByteAt( 1 ) );
		int bpos = 2;
		if( false )
		{
			byte[] b = this.getData();
			while( bpos < b.length )
			{
				int sststartpos = ByteTools.readInt( b[bpos++], b[bpos++], b[bpos++], b[bpos++] );
				int continueStrPos = ByteTools.readInt( b[bpos++], b[bpos++], b[bpos++], b[bpos++] );
				Logger.logInfo( "ExtSST IB = " + (sststartpos - 16) );
				Logger.logInfo( "ExtSST CB = " + continueStrPos );
				Logger.logInfo( "" );
			}
		}

	}

	/**
	 * update the Dsst number
	 */
	void setDsst( int newdst )
	{
		byte[] newdt = ByteTools.shortToLEBytes( (short) newdst );
		this.getData()[0] = newdt[0];
		this.getData()[1] = newdt[1];
		this.Dsst = (short) newdst;
	}

	/**
	 * create a new Isstinf array from Sst strings
	 * <p/>
	 * If you don't update these you will have problems with particular files.  Sorry!!!
	 */
	public void updateIsstinfs()
	{
		List strs = mysst.getStringVector();
		int totstrs = strs.size();

		// if the total of unique UStrings is < 1024,
		// then Dsst = 8, total Isstinfs is cstun/8
		int newdsst = 8;

		// otherwise, divide cstun by 128 and round up
		// to get Dsst.  total Isstinfs  is cstun/Dsst.
		if( totstrs > 1024 )
		{
			newdsst = totstrs / 128;
			if( (totstrs % newdsst) > 0 )
			{
				newdsst++;
			}
		}
		int totissts = totstrs / newdsst;
		if( (totstrs % newdsst) > 0 )
		{
			totissts++;
		}
		int ctr = 0;
		// this is the location of the actual first Ustring data in the Workbook stream
		// add the opcodeLen, length, and sst fields to this value to get the correct
		// pointer directly into the sst
		int sststartpos = mysst.getOffset() + 12;

		Isstinf[] mynewstinfs = new Isstinf[totissts];
		int continueStrPos = 0, lastContinueStrPos = 0, strlen = 0, off3 = 0, bytepos = 2, sstpos = 0, contoff = 0;
		byte[] bd = new byte[((totissts) * 8) + 2];
		byte[] numStrsPer = ByteTools.shortToLEBytes( (short) newdsst );
		bd[0] = numStrsPer[0];
		bd[1] = numStrsPer[1];
		int RecLen = WorkBookFactory.MAXRECLEN;
		int whichContinue = -1;
		boolean updateRecLen = false;

		int lastsstpos = 0;

		// iterate through the issts
		int sstOffset = mysst.getOffset();
		for( int t = 0; t < totissts; t++ )
		{
			Unicodestring str;
			if( t == totissts )
			{
				str = (Unicodestring) mysst.getUStringAt( totstrs - 1 );
			}
			else
			{
				// get the data
				str = (Unicodestring) mysst.getUStringAt( ctr );
			}
			ctr += newdsst;

			lastsstpos = sstpos;
			lastContinueStrPos = continueStrPos;
			sstpos = str.getSSTPos() + 4; // always 4 off

			boolean newbucket = false;
			// set the offset
			//int stringlength=0;
			continueStrPos = (sstpos - contoff);
			if( (continueStrPos > RecLen) && ((whichContinue + 1) < mysst.continues.size()) )
			{ // new bucket
				updateRecLen = true;
				continueStrPos = ((continueStrPos - RecLen) + 1);

				if( updateRecLen )
				{
					List v = mysst.continues;
					BiffRec br = (BiffRec) v.get( ++whichContinue );
					RecLen = br.getLength();
					updateRecLen = false;
				}
				Object[] grbits = mysst.getContinueDef( mysst, false );
				Byte[] grbytes = (Byte[]) grbits[1];
				if( (grbytes[whichContinue] != null) && ((grbytes[whichContinue] == 0x0) || (grbytes[whichContinue] == 0x1)) )
				{
					contoff += (RecLen);
					sstOffset += 5;
				}
				else
				{
					contoff += (RecLen);
					sstOffset += 4;
				}

			}

			sststartpos = sstpos + sstOffset;
			// get the data into bytes

			if( !true )
			{
				Logger.logInfo( "ExtSST IB = " + sststartpos );
				Logger.logInfo( "ExtSST CB = " + (continueStrPos - 4) );
				Logger.logInfo( "String number " + (t * newdsst) );
				Logger.logInfo( "Continue number " + whichContinue );
				Logger.logInfo( "" );
			}

			byte[] ib1 = ByteTools.cLongToLEBytes( sststartpos );
			byte[] cb1 = ByteTools.shortToLEBytes( (short) continueStrPos );//((short)0);

			bd[bytepos++] = ib1[0];
			bd[bytepos++] = ib1[1];
			bd[bytepos++] = ib1[2];
			bd[bytepos++] = ib1[3];
			bd[bytepos++] = cb1[0];
			bd[bytepos++] = cb1[1];
			bd[bytepos++] = 0x0;
			bd[bytepos++] = 0x0;

			lastContinueStrPos = continueStrPos;
		}
		this.setData( bd );
		this.setDsst( newdsst );
		//int orsz = this.getOriginalDataSize();
		int newsz = this.getLength();
		if( DEBUGLEVEL > 3 )
		{
			Logger.logInfo( "Done creating Isstinfs" );
		}
	}

	@Override
	public void preStream()
	{
		try
		{
			this.updateIsstinfs();
		}
		catch( Exception e )
		{
			Logger.logErr( "Extsst.preStream() failed updating Isstinfs.", e );
		}
	}

}
