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

import com.extentech.ExtenXLS.CellRange;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.toolkit.ByteTools;
import com.extentech.toolkit.CompatibleVector;
import com.extentech.toolkit.Logger;

import java.util.Iterator;
import java.util.List;

/**
 * <b>Mergedcells: Merged Cells for Sheet (E5h)</b><br>
 * <p/>
 * Merged Cells record
 * <p/>
 * <p><pre>
 * offset  name            size    contents
 * ---
 * 0	 					var		BiffRec range addresses
 * <p/>
 * </p></pre>
 */
public final class Mergedcells extends com.extentech.formats.XLS.XLSRecord
{

	/**
	 * serialVersionUID
	 */
	private static final long serialVersionUID = 6638569392267433468L;
	public static int MAXRANGES = 1024;
	private int nummerges = 0;
	private CompatibleVector ranges = new CompatibleVector();

	public void init()
	{
		super.init();
		if( DEBUGLEVEL > 5 )
		{
			Logger.logInfo( "Mergedcells record." );
		}
	}

	public static XLSRecord getPrototype()
	{
		Mergedcells newrec = new Mergedcells();
		// for larger records, we need to read in from a file...
		newrec.setOpcode( MERGEDCELLS );
		newrec.setData( new byte[]{ 0x0, 0x0, 0x0, 0x0 } );
		return newrec;
	}

	/**
	 * Un-merge a CellRange
	 *
	 * @param rng
	 */
	public void removeCellRange( CellRange rng )
	{
		this.ranges.remove( rng );
		this.update();
	}

	/**
	 * merge a CellRange
	 *
	 * @param rng
	 */
	public void addCellRange( CellRange rng )
	{
		//  rng.setIsmerge(true);
		this.ranges.add( rng );
	}

	/**
	 * returns the Merged BiffRec Ranges for this WorkSheet
	 *
	 * @return an array of CellRanges each containing Merged Cells.
	 */
	public CellRange[] getMergedRanges()
	{
		if( ranges.size() < 1 )
		{
			return null;
		}
		CellRange[] ret = new CellRange[ranges.size()];
		ranges.toArray( ret );
		return ret;
	}

	/**
	 * update the underlying bytes with any new CellRange info
	 */
	public void update()
	{
		if( ranges.size() > MAXRANGES )
		{
			this.handleMultiRec();
		}
		this.nummerges = ranges.size();
		int datasz = nummerges * 8;
		datasz += 2;
		data = new byte[datasz];
		// get the number of CellRanges
		byte[] szbt = ByteTools.shortToLEBytes( (short) ranges.size() );
		data[0] = szbt[0];
		data[1] = szbt[1];
		int pos = 2;
		if( DEBUGLEVEL > DEBUG_LOW )
		{
			Logger.logInfo( "updating Mergedcell with " + nummerges + " merges." );
		}
		for( int t = 0; t < ranges.size(); t++ )
		{
			CellRange thisrng = (CellRange) ranges.get( t );
// why call this now??? much overhead ...			if(!thisrng.update())
// note: range is updated upon setRange ...			
//				return;
			int[] rints = thisrng.getRowInts();
			int[] cints = thisrng.getColInts();

			//Logger.logInfo(rints[0]+",");
			//for(int x=0;x<cints.length;x++)Logger.logInfo(cints[x]+",");
			//Logger.logInfo();

			byte[] rowmin = ByteTools.shortToLEBytes( (short) (rints[0] - 1) );
			data[pos++] = rowmin[0];
			data[pos++] = rowmin[1];

			byte[] rowmax = ByteTools.shortToLEBytes( (short) (rints[rints.length - 1] - 1) );
			data[pos++] = rowmax[0];
			data[pos++] = rowmax[1];

			byte[] colmin = ByteTools.shortToLEBytes( (short) cints[0] );
			data[pos++] = colmin[0];
			data[pos++] = colmin[1];

			byte[] colmax = ByteTools.shortToLEBytes( (short) cints[cints.length - 1] );
			data[pos++] = colmax[0];
			data[pos++] = colmax[1];

		}
		this.setData( data );
	}

	/**
	 * Mergedranges do not use Continues but instead
	 * there are multiple Mergedranges.
	 *
	 * @return
	 */
	void handleMultiRec()
	{
		//if(true)return;
		if( ranges.size() < MAXRANGES )
		{
			return;
		}
		this.nummerges = MAXRANGES;
		Mergedcells mcfresh = (Mergedcells) Mergedcells.getPrototype();
		List substa = ranges.subList( MAXRANGES, ranges.size() );
		Iterator ita = substa.iterator();
		while( ita.hasNext() )
		{
			mcfresh.addCellRange( (CellRange) ita.next() );
		}
		Iterator removes = mcfresh.ranges.iterator();
		while( removes.hasNext() )
		{
			ranges.remove( removes.next() );
		}
		this.getWorkBook().addRecord( mcfresh, false );

		int idx = this.getRecordIndex() + 1;
		mcfresh.setSheet( this.getSheet() );
		this.getSheet().addMergedCellsRec( mcfresh );
		//Logger.logInfo("ADDING Mergedcells at idx: " + idx);
		this.getStreamer().addRecordAt( mcfresh, idx );
		mcfresh.init();
		mcfresh.update();
	}

	/**
	 * Initialize the CellRanges containing the Merged Cells.
	 *
	 * @param the workbook containing the cells
	 */
	public void initCells( WorkBookHandle wbook )
	{
		nummerges = (int) ByteTools.readShort( this.getByteAt( 0 ), getByteAt( 1 ) );
		ranges = new CompatibleVector();
		int pos = 2; // pointer to the indexes
		for( int x = 0; x < nummerges; x++ )
		{
			int[] cellcoords = new int[4];
			cellcoords[0] = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) ); // first col onebased
			cellcoords[2] = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) ); // last col onebased
			cellcoords[1] = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) ); // first row
			cellcoords[3] = ByteTools.readShort( getByteAt( pos++ ), getByteAt( pos++ ) ); // last row
			try
			{
				// for(int xd=0;xd<cellcoords.length;xd++)
				//    Logger.logInfo("" + cellcoords[xd]+",");
				WorkSheetHandle shtr = wbook.getWorkSheet( this.getSheet().getSheetName() );

				// TODO: testing -- this saves about 30MB in parsing the Reflexis 700+ sheet problem
				CellRange cr = new CellRange( shtr, cellcoords, false );

				//	Logger.logInfo(" init: " + cr);;
				//	Logger.logInfo(x);
				cr.setWorkBook( wbook );
				ranges.add( cr );
				BiffRec[] ch = cr.getCellRecs();
				Mulblank aMul = null;
				for( int t = 0; t < ch.length; t++ )
				{
					// set the range of merged cells
					if( ch[t] != null )
					{
						if( ch[t].getOpcode() == MULBLANK )
						{
							if( aMul == (Mulblank) ch[t] )
							{
								continue;    // skip- already handled
							}
							else
							{
								aMul = (Mulblank) ch[t];
							}
						}
						ch[t].setMergeRange( cr );
					}
				}
			}
			catch( Throwable e )
			{
				//	Logger.logWarn("initializing Merged Cells failed: " + e);
			}
		}
	}

}