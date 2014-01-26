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

import org.openxls.ExtenXLS.CellRange;
import org.openxls.ExtenXLS.WorkBookHandle;
import org.openxls.ExtenXLS.WorkSheetHandle;
import org.openxls.toolkit.ByteTools;
import org.openxls.toolkit.CompatibleVector;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
public final class Mergedcells extends XLSRecord
{
	private static final Logger log = LoggerFactory.getLogger( Mergedcells.class );
	private static final long serialVersionUID = 6638569392267433468L;
	public static int MAXRANGES = 1024;
	private int nummerges = 0;
	private CompatibleVector ranges = new CompatibleVector();

	@Override
	public void init()
	{
		super.init();
			log.trace( "Mergedcells record." );
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
		ranges.remove( rng );
		update();
	}

	/**
	 * merge a CellRange
	 *
	 * @param rng
	 */
	public void addCellRange( CellRange rng )
	{
		//  rng.setIsmerge(true);
		ranges.add( rng );
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
			handleMultiRec();
		}
		nummerges = ranges.size();
		int datasz = nummerges * 8;
		datasz += 2;
		data = new byte[datasz];
		// get the number of CellRanges
		byte[] szbt = ByteTools.shortToLEBytes( (short) ranges.size() );
		data[0] = szbt[0];
		data[1] = szbt[1];
		int pos = 2;
			log.debug( "updating Mergedcell with " + nummerges + " merges." );
		for( Object range : ranges )
		{
			CellRange thisrng = (CellRange) range;
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
		setData( data );
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
		nummerges = MAXRANGES;
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
		getWorkBook().addRecord( mcfresh, false );

		int idx = getRecordIndex() + 1;
		mcfresh.setSheet( getSheet() );
		getSheet().addMergedCellsRec( mcfresh );
		//Logger.logInfo("ADDING Mergedcells at idx: " + idx);
		getStreamer().addRecordAt( mcfresh, idx );
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
		nummerges = ByteTools.readShort( getByteAt( 0 ), getByteAt( 1 ) );
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
				WorkSheetHandle shtr = wbook.getWorkSheet( getSheet().getSheetName() );

				// TODO: testing -- this saves about 30MB in parsing the Reflexis 700+ sheet problem
				CellRange cr = new CellRange( shtr, cellcoords, false );

				cr.setWorkBook( wbook );
				ranges.add( cr );
				BiffRec[] ch = cr.getCellRecs();
				Mulblank aMul = null;
				for( BiffRec aCh : ch )
				{
					// set the range of merged cells
					if( aCh != null )
					{
						if( aCh.getOpcode() == MULBLANK )
						{
							if( aMul.equals( aCh ) )
							{
								continue;    // skip- already handled
							}
							aMul = (Mulblank) aCh;
						}
						aCh.setMergeRange( cr );
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