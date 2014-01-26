package com.extentech.formats.XLS;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.SortedSet;
import java.util.TreeSet;

/**
 * User: npratt
 * Date: 1/24/14
 * Time: 13:03
 */
public class CellsByColImpl implements CellsByCol
{
	private static final Logger log = LoggerFactory.getLogger( CellsByColImpl.class );
	private final Map<Integer, SortedSet<CellRec>> cells;

	public CellsByColImpl()
	{
		cells = new HashMap<>();
	}

	@Override
	public List<CellRec> get( int col )
	{
		SortedSet<CellRec> colCells = cells.get( col );
		if( colCells != null )
		{
			return new ArrayList<>( colCells );
		}

		return new ArrayList<>();
	}

	@Override
	public void add( CellRec cell )
	{
		int first = cell.getColFirst();
		int last = cell.getColLast();

		validate( first, last );

		log.debug( "Adding R{}C{}-{}", cell.getRowNumber(), first, last);
		//
		// While we do double loop on this, I want the data structure usage safety over any potential performance issues right now.
		// If add() is being invoked for a Cell that already exists, that implies an incorrect usage and it should be fixed.
		//

		for( int col = first; col <= last; col++ )
		{
			SortedSet<CellRec> colCells = cells.get( col );
			if( (colCells != null) && colCells.contains( cell ) )
			{
				String msg = String.format( "Attempt to add cell '%s' that is already in Column %d.  Existing cell: {}", cell, col );
				log.error( msg );
				// FIXME: Code upstream invokes this multiple times and needs to be fixed before re-enabling this exception (which I intend to do)
//				throw new IllegalArgumentException( msg );
			}
		}

		for( int col = first; col <= last; col++ )
		{
			SortedSet<CellRec> colCells = cells.get( col );
			if( colCells == null )
			{
				colCells = new TreeSet<>( new RowComparator() );
				cells.put( col, colCells );
			}

			boolean added = colCells.add( cell );
			if( !added )
			{
				log.error( "Unable to add cell to Collection - it seems to already exist! Cell: {}, Col: {}", cell, col );
			}
		}

		log.debug( "Added Cell '{}' to Cols: {}-{}", cell, first, last );
	}

	@Override
	public CellRec remove( CellRec cell )
	{
		int first = cell.getColFirst();
		int last = cell.getColLast();

		validate( first, last );

		//
		// While we double loop on this, I want the data structure usage safety over any potential performance issues right now.
		// If remove() is being invoked for a Cell that doesnt exist in the data structure, that implies an incorrect usage and it should
		// be fixed.
		//

		for( int col = first; col <= last; col++ )
		{
			SortedSet<CellRec> colCells = cells.get( col );
			if( (colCells == null) || !colCells.contains( cell ) )
			{
				// Perhaps the cell was modified before being removed from here.
				String msg = String.format( "Attempt to remove cell '%s' that doesn't exist in Column %d", cell, col );
				log.warn( msg );
				throw new IllegalArgumentException( msg );
			}
		}

		for( int col = first; col <= last; col++ )
		{
			SortedSet<CellRec> colCells = cells.get( col );
			colCells.remove( cell );
		}

		log.debug( "Removed cell '{}' from Cols {}-{}", cell, first, last );
		return cell;
	}

	private static void validate( int first, int last )
	{
		if( first > last )
		{
			log.error( "What sort of messed up scenario is this? first({}) > last({})", first, last );
		}
	}

	private class RowComparator implements Comparator<CellRec>
	{
		@Override
		public int compare( CellRec c1, CellRec c2 )
		{
			return Integer.compare( c1.getRowNumber(), c2.getRowNumber() );
		}
	}
}
