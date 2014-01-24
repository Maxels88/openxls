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

import com.extentech.formats.XLS.charts.Chart;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Random;

/**
 * WorkBookAssembler handles the details of how a biff8 file
 * should be constructed, what records go in what order, etc
 */
public class WorkBookAssembler implements XLSConstants
{

	/**
	 * assembleSheetRecs assembles the array of records, then ouputs
	 * the ordered list to the bytestreamer, which should be the only
	 * thing calling this.
	 * <p/>
	 * The SheetRecs should contain all records excepting valrecs, rows, index, and dbcells at this
	 * point, where they should be created.
	 * <p/>
	 * The dbcell creation and population also occurs here.  The first pointer is the offset between
	 * the beginning of the second row and the first valrec.  After that it is just incrementing valrecs
	 * <p/>
	 * TODO:  John,  please review use of collections & performance.  Check the Dbcell.initDbCell method as well.
	 * TODO:  Add handling for creation of mulblank and mulrk records on the fly.  We no longer have them!
	 */
	public static List assembleSheetRecs( Boundsheet thissheet )
	{
		List addVec = new ArrayList();
		WorkBookAssembler.preProcessSheet( thissheet );
		addVec.addAll( thissheet.getSheetRecs() );
		if( !thissheet.isChartOnlySheet() )
		{
			addVec = WorkBookAssembler.assembleRows( thissheet, addVec );
		}
		if( thissheet.getWorkBook().getCharts().length > 0 )
		{
			addVec = WorkBookAssembler.assembleChartRecs( thissheet, addVec );
		}
		return addVec;
	}

	/**
	 * Handles functionality that needs to occur before the boundsheet can be
	 * reliably created
	 */
	private static void preProcessSheet( Boundsheet thissheet )
	{
		if( thissheet.hasMergedCells() )
		{
			// update mergedcells first, as they may grow in record size and are not handled by continues.
			Iterator itx = thissheet.getMergedCellsRecs().iterator();
			while( itx.hasNext() )
			{
				Mergedcells mrg = (Mergedcells) itx.next();
				mrg.update();
			}
		}
	}

	/**
	 * Add the rows of data into the worksheet level stream
	 *
	 * @param thissheet
	 * @param addVec
	 * @return
	 */
	private static List assembleRows( Boundsheet thissheet, List addVec )
	{
		Dimensions dim = thissheet.getDimensions();
		Random gen = new Random();
		int randomNumber = gen.nextInt();
		int insertRowidx = 0;
		if( dim != null )
		{
			insertRowidx = dim.getRecordIndex() + 1;
		}
		int insertValidx = insertRowidx;
		int rowCount = 0; // use this to break every 32 rows for a new dbcell
		int dbOffset = 0; // the offset between the first row and the dbcell
		int maxRow = 0;
		int maxCol = 0;
		List dbOffsets = new ArrayList(); // dbcell offsets
		int valrecOffset = 0;
		// KSC: clear out index dbcells
		if( thissheet.getSheetIDX() != null )
		{
			thissheet.getSheetIDX().resetDBCells();
		}
		// NOTE: if below sorting is time-consuming, should input sort inserting and deleting rows and not here ...
		Iterator outRows = thissheet.getSortedRows().keySet().iterator();
		while( outRows.hasNext() )
		{
			if( rowCount == 32 )
			{
				// Add a new dbcell for every 32 rows
				Dbcell d = (Dbcell) Dbcell.getPrototype();
				dbOffsets.add( 0, (short) ((rowCount - 1) * 20) );
				d.initDbCell( dbOffset, dbOffsets );
				addVec.add( insertValidx++, d );
				if( thissheet.getSheetIDX() != null )
				{
					thissheet.getSheetIDX().addDBCell( d );
				}
				insertRowidx = insertValidx;
				dbOffsets = new ArrayList();
				rowCount = 0;
				dbOffset = 0;
			}
			Row r = thissheet.getRowMap().get( outRows.next() );
			rowCount++;
			maxRow = Math.max( r.getRowNumber(), maxRow );
			dbOffset += r.getLength();
			addVec.add( insertRowidx++, r );
			insertValidx++;

			// insert the valrec and child recs from the row
			Mulblank skipMull = null;
			List outRecs = r.getValRecs( randomNumber );
			Iterator it = outRecs.iterator();
			while( it.hasNext() )
			{
				BiffRec or = (BiffRec) it.next();
				if( skipMull != null )
				{
					if( or.equals( skipMull ) )
					{
						continue;
					}
					skipMull = null;
				}
				addVec.add( insertValidx++, or );
				valrecOffset += or.getLength();
				dbOffset += or.getLength();
				short orc = or.getColNumber();
				if( (!it.hasNext()) && (orc > maxCol) )
				{
					maxCol = orc;
				}
				if( or.getOpcode() == MULBLANK )
				{
					skipMull = (Mulblank) or;
				}
			}
			dbOffsets.add( (short) valrecOffset );
			valrecOffset = 0;
		}
		// add the final dbcell.  Chart only sheets will not have an index, so ignore if so.
		if( dbOffsets.size() > 0 )
		{
			Dbcell d = (Dbcell) Dbcell.getPrototype();
			dbOffsets.add( 0, (short) ((rowCount - 1) * 20) );
			d.initDbCell( dbOffset, dbOffsets );
			d.setOffset( dbOffset );    // KSC: Added to set dbCell offset
			addVec.add( insertValidx++, d );
			if( thissheet.getSheetIDX() != null )
			{
				thissheet.getSheetIDX().addDBCell( d );
			}
		}
		thissheet.updateDimensions( maxRow, maxCol/* 20100225 KSC: incrementing does not match Excel results: take out +1*/ );
		return addVec;
	}

	/**
	 * Get the index to the last db cell in the boundsheet
	 * this is where objects can start being inserted
	 *
	 * @return
	 */
	private static int getLastDBCellLocation( List addVec )
	{
		for( int i = addVec.size() - 1; i >= 0; i-- )
		{
			BiffRec b = (BiffRec) addVec.get( i );
			if( b.getOpcode() == DBCELL )
			{
				return i;
			}
		}
		return 0;
	}

	/**
	 * Add the chart records to the array.  There are two types of charts to be added,
	 * those that have a pointer into the worksheet level stream (ie they existed when the file was parsed)
	 * and those that have been added post parse and/or are converted from ooxml.
	 * <p/>
	 * We hold the location of the parsed ones to deal with odd inconsistencies that exist
	 * with multiple drawing objects and txos
	 *
	 * @param thissheet
	 * @param addVec
	 * @return
	 */
	private static List assembleChartRecs( Boundsheet thissheet, List addVec )
	{
		// insert charts that have bound obj records in the stream
		int insertValidx = WorkBookAssembler.getLastDBCellLocation( addVec );
		ArrayList insertedCharts = new ArrayList();
		int chartInsert = insertValidx;
		while( chartInsert < addVec.size() )
		{
			XLSRecord x = (XLSRecord) addVec.get( chartInsert );
			if( x.getOpcode() == OBJ )
			{
				Obj o = (Obj) x;
				if( o.getChart() != null )
				{
					insertedCharts.add( o.getChart().getTitle() );
					List l = o.getChart().assembleChartRecords();
					addVec.addAll( chartInsert + 1, l );
					chartInsert += l.size();
					insertValidx += l.size();
				}
			}
			chartInsert++;
		}

		// insert charts that are new, either from transfers, insertions, etc
		Chart[] chts = thissheet.getWorkBook().getCharts();
		chartInsert = insertValidx;
		for( Chart cht : chts )
		{
			if( !insertedCharts.contains( cht.getTitle() ) && cht.getSheet().equals( thissheet ) )
			{
				// if it's a chart only sheet, insertValidx will be '0' here, put it at the end of the current recordset
				if( insertValidx == 0 )
				{
					insertValidx = addVec.size();
				}
				List l = cht.assembleChartRecords();
				if( cht.getObj() != null )
				{
					l.add( 0, cht.getObj() );
				}
				if( cht.getMsodrawobj() != null )
				{
					l.add( 0, cht.getMsodrawobj() );
				}
				int spid = 0;
				boolean isHeader = false;
				if( l.get( 0 ) instanceof MSODrawing )
				{
					isHeader = ((MSODrawing) l.get( 0 )).isHeader();
					spid = ((MSODrawing) l.get( 0 )).getSPID();
				}
				while( chartInsert < addVec.size() )
				{
					XLSRecord x = (XLSRecord) addVec.get( chartInsert );
					if( x.getOpcode() == MSODRAWING )
					{
						MSODrawing mm = (MSODrawing) x;
						if( ((mm.getSPID() > spid) && (spid != 0) && !mm.isHeader()) || isHeader )
						{
							insertValidx = chartInsert;
							chartInsert = addVec.size();
						}
					}
					if( (x.getOpcode() == WINDOW2) || (x.getOpcode() == MSODRAWINGSELECTION) || (x.getOpcode() == NOTE) )
					{
						insertValidx = chartInsert;
						chartInsert = addVec.size();
					}
					chartInsert++;
				}
				addVec.addAll( insertValidx, l );
				insertValidx += l.size();
			}
		}
		return addVec;
	}
}
