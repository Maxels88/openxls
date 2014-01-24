package com.extentech.formats.XLS;

import org.junit.Before;
import org.junit.Test;

import java.util.TreeMap;

import static org.junit.Assert.assertEquals;

/**
 * User: npratt
 * Date: 1/23/14
 * Time: 16:28
 */
public class ComparatorTest
{
	private ColumnRange.Comparator comparator;

	@Before
	public void setUp() throws Exception
	{
		comparator = new ColumnRange.Comparator();
	}

	@Test
	public void testCompareSame() throws Exception
	{
		ColumnRange cr1 = new CRStub( 0, 0 );
		ColumnRange cr2 = new CRStub( 0, 0 );
		assertEquals( 0, comparator.compare( cr1, cr2 ) );
	}

	@Test
	public void testCompareExclusiveAscending() throws Exception
	{
		ColumnRange cr1 = new CRStub( 0, 1 );
		ColumnRange cr2 = new CRStub( 2, 3 );
		assertEquals( -1, comparator.compare( cr1, cr2 ) );
	}

	@Test
	public void testCompareExclusiveDescending() throws Exception
	{
		ColumnRange cr1 = new CRStub( 2, 3 );
		ColumnRange cr2 = new CRStub( 0, 1 );
		assertEquals( 1, comparator.compare( cr1, cr2 ) );
	}

	@Test
	public void testLHSInsideRHS() throws Exception
	{
		ColumnRange cr1 = new CRStub( 2, 3 );
		ColumnRange cr2 = new CRStub( 0, 4 );
		assertEquals( 1, comparator.compare( cr1, cr2 ) );
	}

	@Test
	public void testRHSInsideLHS() throws Exception
	{
		ColumnRange cr1 = new CRStub( 0, 4 );
		ColumnRange cr2 = new CRStub( 2, 3 );
		assertEquals( -1, comparator.compare( cr1, cr2 ) );
	}

	@Test
	public void testOverlapOnLeftSide() throws Exception
	{
		ColumnRange cr1 = new CRStub( 0, 3 );
		ColumnRange cr2 = new CRStub( 2, 4 );
		assertEquals( -1, comparator.compare( cr1, cr2 ) );
	}

	@Test
	public void testOverlapOnRightSide() throws Exception
	{
		ColumnRange cr1 = new CRStub( 3, 7 );
		ColumnRange cr2 = new CRStub( 2, 4 );
		assertEquals( 1, comparator.compare( cr1, cr2 ) );
	}

	static class CRStub implements ColumnRange
	{
		private final int colFirst;
		private final int colLast;

		private CRStub( int colFirst, int colLast )
		{
			this.colFirst = colFirst;
			this.colLast = colLast;
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
	}
}
