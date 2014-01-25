package com.extentech.formats.XLS;

import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.Map;
import java.util.SortedMap;
import java.util.TreeMap;

import static org.junit.Assert.assertEquals;

/**
 * User: npratt
 * Date: 1/23/14
 * Time: 17:34
 */
public class ColumnMajorComparatorTest
{
	private static final Logger log = LoggerFactory.getLogger( ColumnMajorComparatorTest.class );
	private CellAddressible.ColumnMajorComparator comparator;
	private TreeMap<CellAddressible, Integer> treemap;

	@Before
	public void setUp() throws Exception
	{
		comparator = new CellAddressible.ColumnMajorComparator();
		treemap = new TreeMap<>( comparator );

		int val = 0;
		treemap.put( new RowCols( 0, 0 ), val++ );
		treemap.put( new RowCols( 0, 1 ), val++ );
		treemap.put( new RowCols( 0, 2 ), val++ );
		treemap.put( new RowCols( 0, 3 ), val++ );
		treemap.put( new RowCols( 0, 4 ), val++ );
		treemap.put( new RowCols( 1, 0 ), val++ );
		treemap.put( new RowCols( 1, 1 ), val++ );
		treemap.put( new RowCols( 1, 2 ), val++ );
		treemap.put( new RowCols( 1, 3 ), val++ );
		treemap.put( new RowCols( 1, 4 ), val++ );
		treemap.put( new RowCols( 2, 0 ), val++ );
		treemap.put( new RowCols( 2, 1 ), val++ );
		treemap.put( new RowCols( 2, 2 ), val++ );
		treemap.put( new RowCols( 2, 3 ), val++ );
		treemap.put( new RowCols( 2, 4 ), val++ );
	}

	@Test
	public void testCompareWhenLessThan() throws Exception
	{
		assertEquals( -1, comparator.compare( new RowCols( 0, 0 ), new RowCols( 0, 1 ) ) );
	}

	@Test
	public void testCompareWhenLHSGreaterThanRHS() throws Exception
	{
		assertEquals( 1, comparator.compare( new RowCols( 0, 1 ), new RowCols( 2, 0 ) ) );
	}

	@Test
	public void testCompareWhenLHSLessThanRHS() throws Exception
	{
		assertEquals( -1, comparator.compare( new RowCols( 2, 0 ), new RowCols( 1, 5 ) ) );
	}

	@Test
	public void testCompareWhenGreaterThanRowSameCol() throws Exception
	{
		assertEquals( 1, comparator.compare( new RowCols( 1, 1), new RowCols( 0, 1 ) ) );
	}

	@Test
	public void testCompareWhenGreaterThanColSameRow() throws Exception
	{
		assertEquals( 1, comparator.compare( new RowCols( 0, 2), new RowCols( 0, 1 ) ) );
	}

	@Test
	public void testSubMapForAll() throws Exception
	{
		// Note: We have to go beyond the end of the key range to get everything - check the javadocs for treemap.subMap()
		SortedMap<CellAddressible,Integer> submap = treemap.subMap( new RowCols( 0, 0 ), new RowCols( 2, 5 ) );

		assertEquals( 15, submap.size() );
	}

	@Test
	public void testSubMapForSingleColumn() throws Exception
	{
		SortedMap<CellAddressible,Integer> submap = treemap.subMap( new RowCols( 0, 0 ), new RowCols( 0, 1 ) );
		assertEquals( 3, submap.size() );
	}

	@Test
	public void testSubMapForAnotherSingleColumn() throws Exception
	{
		SortedMap<CellAddressible,Integer> submap = treemap.subMap( new RowCols( 0, 2 ), new RowCols( 0, 3 ) );
		assertEquals( 3, submap.size() );
	}

	//
	//
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//
	//

	private void dumpMap( Map<CellAddressible, Integer> map )
	{
		for( Map.Entry<CellAddressible, Integer> entry : map.entrySet() )
		{
			log.info( "Entry {} = {}", entry.getKey(), entry.getValue() );
		}
	}

	private static class RowCols implements CellAddressible
	{
		private final int row;
		private final int colFirst;
		private final int colLast;

		public RowCols( int row, int col )
		{
			this.row = row;
			colFirst = col;
			colLast = col;
		}

		@Override
		public int getRowNumber()
		{
			return row;
		}

		@Override
		public int getColFirst()
		{
			return colFirst;
		}

		@Override
		public int getColLast()
		{
			return colLast;
		}

		@Override
		public boolean isSingleCol()
		{
			return colFirst == colLast;
		}

		@Override
		public String toString()
		{
			return "RowCols{" +
					"row=" + row +
					", cols=" + colFirst +
					"-" + colLast +
					'}';
		}
	}
}
